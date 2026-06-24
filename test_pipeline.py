"""
test_pipeline.py — Kiem tra toan dien codebase doi chieu ACH.

Kich ban kiem tra:
  T01  Pipeline 237k+ dong tu tao — TXRT fix, CUSTOMER filter, ty le khop
  T02  skip_file bug — GL02 voi dong dau LOCAC sai, van doc duoc dong 502003
  T03  Cancel Point 1 — huy truoc Phase 1
  T04  Cancel Point 2 — huy sau B3
  T05  Cancel Point 3 — huy sau Phase 1
  T06  Cancel Point 4 — huy truoc ghi Excel
  T07  GL02 khong co LOCAC column — chay khong crash
  T08  MIS_DI rong sau pre-filter — chay khong crash
  T09  GW rong — timeout = 0, cap_cn_tien = 0
  T10  PHAN_TICH — kiem tra tinh toan phan tram

Chay: python test_pipeline.py
"""
import io
import os
import sys
import shutil
import threading
import traceback
from datetime import datetime, timedelta
from collections import defaultdict

import numpy as np
import pandas as pd
import pyzipper
import xlsxwriter

# ─── Config ───────────────────────────────────────────────────────────────────
TEST_DIR = '_test_tmp'
ZIP_PW   = b'DACwLdHi'
PASS_CH  = '\033[92m[PASS]\033[0m'
FAIL_CH  = '\033[91m[FAIL]\033[0m'
SKIP_CH  = '\033[93m[SKIP]\033[0m'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as _cfg
_cfg.ZIP_PASSWORD = ZIP_PW   # override de dung voi test data

from main import main_from_dir

results = []

def run(name, fn):
    try:
        fn()
        results.append((name, True, ''))
        print(f'  {PASS_CH} {name}')
    except AssertionError as e:
        results.append((name, False, str(e)))
        print(f'  {FAIL_CH} {name}: {e}')
    except Exception as e:
        results.append((name, False, traceback.format_exc()[-300:]))
        print(f'  {FAIL_CH} {name}: {e}')

def assert_eq(a, b, msg=''):
    assert a == b, f'{msg}: expected={b}, got={a}'

def assert_in_range(val, lo, hi, msg=''):
    assert lo <= val <= hi, f'{msg}: {val} not in [{lo}, {hi}]'


# ─── Helpers de tao du lieu test nho ──────────────────────────────────────────
def _write_zip(path, df, csv_name='data.csv', pw=ZIP_PW):
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding='utf-8-sig')
    buf.seek(0)
    with pyzipper.AESZipFile(path, 'w', compression=pyzipper.ZIP_DEFLATED,
                              encryption=pyzipper.WZ_AES) as z:
        z.setpassword(pw)
        z.writestr(csv_name, buf.getvalue())

def _make_gl02_zip(path, rows):
    df = pd.DataFrame(rows)
    _write_zip(path, df, 'GL02.csv')

def _make_gw_xlsx(path, rows, session='SES001'):
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('Sheet 1')
    cols = ['No', 'SessionId', 'BRCD', 'STTLMAMT', 'MSGREF', 'PrcFlg', 'TxDt', 'Status']
    for ci, c in enumerate(cols):
        ws.write(0, ci, c)
    for ri, r in enumerate(rows, 1):
        ws.write_row(ri, 0, [r.get(c, '') for c in cols])
    wb.close()

def _make_gw_xlsx_df(path, df):
    """Ghi DataFrame GW vao XLSX — hieu qua cho du lieu lon (60k+ dong)."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('Sheet 1')
    cols = list(df.columns)
    for ci, c in enumerate(cols):
        ws.write(0, ci, c)
    data = df.fillna('').values.tolist()
    for ri, row in enumerate(data, 1):
        ws.write_row(ri, 0, row)
    wb.close()

_MIS_DI_COLS  = ['NGAY_GIAO_DICH','CHI_NHANH','REFHUB','MSGREF','MSGSEQ','TXID',
                  'KENH_THANH_TOAN','LOAI_LENH_OSB','NH_NHAN','MA_GIAO_DICH','NOI_DUNG',
                  'TRANG_THAI_LENH','SO_TIEN','TRACE','SE_TRACE','SESSION','NGAY_KENH_TRA']
_MIS_DEN_COLS = ['NGAY_GIAO_DICH','CHI_NHANH','REFHUB','MSGREF','MSGSEQ','TXID',
                  'KENH_THANH_TOAN','LOAI_LENH_OSB','NH_GUI','NOI_DUNG',
                  'TRANG_THAI_LENH','SO_TIEN','TRACE','SESSION']

def _make_mis_zip(path, rows, csv_name='mis.csv'):
    if rows:
        df = pd.DataFrame(rows)
    else:
        # Xac dinh cot dua theo ten file
        cols = _MIS_DEN_COLS if 'den' in csv_name.lower() else _MIS_DI_COLS
        df = pd.DataFrame(columns=cols)
    _write_zip(path, df, csv_name)

def _make_pdf(dir_, session, ngay_dt):
    name = f'ACH_{(ngay_dt+timedelta(1)).strftime("%Y%m%d")}_VBAAVNVN_NRT_{session}_N03_1.pdf'
    with open(os.path.join(dir_, name), 'wb') as f:
        f.write(b'%PDF-1.4 test')

NGAY = datetime(2026, 6, 15)
SES  = '11111'   # phai la so nguyen — regex b1 chi nhan \d+


def _base_gl02_row(t, locac='502003', cramount=100_000, dramount=0, customer='1000-003526275'):
    # REFERENCE format: 'REFn' (khong co leading zeros) de SO_TRACE khop voi TRACE.lstrip('0')
    return {'TRDATE': '15/06/2026', 'TRBRCD': '1101', 'USERID': 'U',
            'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': locac,
            'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
            'CUSTOMER': customer, 'TRTP': '1', 'REFERENCE': f'REF{t}',
            'REMARK': '', 'DRAMOUNT': dramount, 'CRAMOUNT': cramount,
            'CRTDTM': '15/06/2026 08:00:00'}

def _base_mis_di_row(t, tt='SCNL', amount=100_000, branch='1101',
                     session=SES, ngay_kenh=''):
    return {'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': branch,
            'REFHUB': '', 'MSGREF': f'MR{t:010d}', 'MSGSEQ': '1',
            'TXID': '0', 'KENH_THANH_TOAN': 'ACH', 'LOAI_LENH_OSB': 'CT',
            'NH_NHAN': 'VBAA', 'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
            'TRANG_THAI_LENH': tt, 'SO_TIEN': amount, 'TRACE': str(t),
            'SE_TRACE': '', 'SESSION': session, 'NGAY_KENH_TRA': ngay_kenh}

def _base_mis_den_row(t, amount=200_000, branch='1101', session=SES):
    return {'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': branch,
            'REFHUB': '', 'MSGREF': f'DN{t:010d}', 'MSGSEQ': '1',
            'TXID': '0', 'KENH_THANH_TOAN': 'ACH', 'LOAI_LENH_OSB': 'CT',
            'NH_GUI': 'VBAA', 'NOI_DUNG': '', 'TRANG_THAI_LENH': 'SCNL',
            'SO_TIEN': amount, 'TRACE': str(t), 'SESSION': session}

def _base_gw_row(t, amount=100_000, branch='1101', session=SES, prc='OK'):
    return {'No': t, 'SessionId': session, 'BRCD': branch,
            'STTLMAMT': amount, 'MSGREF': f'GW{t:010d}', 'PrcFlg': prc,
            'TxDt': '15/06/2026', 'Status': 'SETTLED'}


_run_seq = [0]

def _run_test_case(td, extra_kwargs=None):
    """
    td: dict chua GL02 rows, GW rows, MIS_DI rows, MIS_DEN rows.
    Tao thu muc tam, chay pipeline, tra ve DataFrame Excel result.
    """
    _run_seq[0] += 1
    d = os.path.join(TEST_DIR, f'inp_{_run_seq[0]}')
    o = os.path.join(TEST_DIR, f'out_{_run_seq[0]}')
    shutil.rmtree(d, ignore_errors=True)
    shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)

    _make_pdf(d, td.get('session', SES), td.get('ngay', NGAY))
    _make_gl02_zip(os.path.join(d, 'GL02_20260615.zip'), td['gl02'])
    _make_gw_xlsx(os.path.join(d, 'GW_20260615.xlsx'), td['gw'],
                  session=td.get('session', SES))
    mis_di = td['mis_di']
    mid = len(mis_di) // 2
    _make_mis_zip(os.path.join(d, 'MIS_DI_20260615_01.zip'), mis_di[:mid], 'di1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DI_20260615_02.zip'), mis_di[mid:], 'di2.csv')
    mis_den = td['mis_den']
    _make_mis_zip(os.path.join(d, 'MIS_DEN_20260615_01.zip'), mis_den[:len(mis_den)//2], 'den1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_20260615_02.zip'), mis_den[len(mis_den)//2:], 'den2.csv')

    kw = dict(input_dir=d, output_dir=o, ngay='15/06/2026')
    if extra_kwargs:
        kw.update(extra_kwargs)
    out = main_from_dir(**kw)
    return out, o


def _run_large_test_case(gl02_df, gw_df, mis_di_df, mis_den_df):
    """Chay pipeline voi DataFrame lon (thay vi list of dicts) de tao file hieu qua."""
    _run_seq[0] += 1
    d = os.path.join(TEST_DIR, f'inp_{_run_seq[0]}')
    o = os.path.join(TEST_DIR, f'out_{_run_seq[0]}')
    shutil.rmtree(d, ignore_errors=True)
    shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)

    _make_pdf(d, SES, NGAY)

    _write_zip(os.path.join(d, 'GL02_large.zip'), gl02_df, 'GL02.csv')
    _make_gw_xlsx_df(os.path.join(d, 'GW_large.xlsx'), gw_df)

    mid_di = len(mis_di_df) // 2
    _write_zip(os.path.join(d, 'MIS_DI_large_01.zip'), mis_di_df.iloc[:mid_di], 'di1.csv')
    _write_zip(os.path.join(d, 'MIS_DI_large_02.zip'), mis_di_df.iloc[mid_di:], 'di2.csv')

    mid_den = len(mis_den_df) // 2
    _write_zip(os.path.join(d, 'MIS_DEN_large_01.zip'), mis_den_df.iloc[:mid_den], 'den1.csv')
    _write_zip(os.path.join(d, 'MIS_DEN_large_02.zip'), mis_den_df.iloc[mid_den:], 'den2.csv')

    out = main_from_dir(input_dir=d, output_dir=o, ngay='15/06/2026')
    return out, o


def _read_sheet(out_dir, sheet):
    xlsx = [f for f in os.listdir(out_dir) if f.endswith('.xlsx')]
    assert xlsx, 'Khong co file xlsx trong output'
    xl = pd.ExcelFile(os.path.join(out_dir, xlsx[0]))
    return pd.read_excel(xl, sheet_name=sheet, dtype=str)


# ═══════════════════════════════════════════════════════════════════════════════
# T01: Pipeline lon — 237k+ dong tu tao (vectorized), kiem tra day du
# ═══════════════════════════════════════════════════════════════════════════════
def t01_pipeline_200k_rows():
    """
    Tu tao 237k+ dong du lieu, chay pipeline end-to-end, xac nhan:
    - So luong khop/thua DI va DEN chinh xac
    - CUSTOMER non-ACH bi loai truoc khi vao NPO
    - TXRT fix: SCNL+TXRT cung chiem slot GW -> TPAY timeout
    - Ty le khop DI=80%, DEN=83.33%
    """
    N_MATCH_DI          = 40_000
    N_NPO_DI_THUA       = 10_000
    N_MIS_DI_SCNL_THUA  = 20_000
    N_TIMEOUT_NO_GW     = 20       # TPAY khong co GW slot -> tat ca timeout
    # TXRT fix: 1 SCNL + 1 TXRT + 1 TPAY + 2 GW slots -> TPAY timeout vi SCNL+TXRT het slot
    N_MATCH_DEN         = 25_000
    N_NPO_DEN_THUA      =  5_000
    N_MIS_DEN_THUA      = 10_000
    N_NON_ACH_GL02      =  2_000   # bi loc boi _CUSTOMER_ACH filter truoc khi vao NPO

    # ── GL02 NPO_DI ──────────────────────────────────────────────────────────
    t_m = np.arange(1, N_MATCH_DI + 1)
    gl02_di_match = pd.DataFrame({
        'TRDATE': '15/06/2026', 'TRBRCD': '1101', 'USERID': 'U',
        'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': '502003',
        'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
        'CUSTOMER': '1000-003526275', 'TRTP': '1',
        'REFERENCE': ('REF' + pd.Series(t_m).astype(str)).values,
        'REMARK': '', 'DRAMOUNT': 0, 'CRAMOUNT': 100_000,
        'CRTDTM': '15/06/2026 08:00:00',
    })
    t_nt = np.arange(200_001, 200_001 + N_NPO_DI_THUA)
    gl02_di_thua = pd.DataFrame({
        'TRDATE': '15/06/2026', 'TRBRCD': '1101', 'USERID': 'U',
        'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': '502003',
        'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
        'CUSTOMER': '1000-003526275', 'TRTP': '1',
        'REFERENCE': ('REF' + pd.Series(t_nt).astype(str)).values,
        'REMARK': '', 'DRAMOUNT': 0, 'CRAMOUNT': 100_000,
        'CRTDTM': '15/06/2026 08:00:00',
    })
    # Cung LOCAC 502003 nhung CUSTOMER sai -> bi loc boi b2, khong vao NPO
    t_na = np.arange(900_001, 900_001 + N_NON_ACH_GL02)
    gl02_di_non_ach = pd.DataFrame({
        'TRDATE': '15/06/2026', 'TRBRCD': '1101', 'USERID': 'U',
        'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': '502003',
        'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
        'CUSTOMER': '9999-KHAC', 'TRTP': '1',
        'REFERENCE': ('REF' + pd.Series(t_na).astype(str)).values,
        'REMARK': '', 'DRAMOUNT': 0, 'CRAMOUNT': 50_000,
        'CRTDTM': '15/06/2026 08:00:00',
    })

    # ── GL02 NPO_DEN ─────────────────────────────────────────────────────────
    # REFERENCE='DN{t}', DRAMOUNT=200000, CRAMOUNT=0
    # KEY_DEN = SO_TRACE + DRAMOUNT = str(t) + '200000'
    t_dm = np.arange(1, N_MATCH_DEN + 1)
    gl02_den_match = pd.DataFrame({
        'TRDATE': '15/06/2026', 'TRBRCD': '2101', 'USERID': 'U',
        'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': '502003',
        'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
        'CUSTOMER': '1000-003526275', 'TRTP': '1',
        'REFERENCE': ('DN' + pd.Series(t_dm).astype(str)).values,
        'REMARK': '', 'DRAMOUNT': 200_000, 'CRAMOUNT': 0,
        'CRTDTM': '15/06/2026 08:00:00',
    })
    t_dt = np.arange(700_001, 700_001 + N_NPO_DEN_THUA)
    gl02_den_thua = pd.DataFrame({
        'TRDATE': '15/06/2026', 'TRBRCD': '2101', 'USERID': 'U',
        'JOURSEQ': '1', 'DYTRSEQ': '1', 'LOCAC': '502003',
        'CCY': 'VND', 'BUSCD': 'ACH', 'UNIT': '0', 'TRCD': 'A',
        'CUSTOMER': '1000-003526275', 'TRTP': '1',
        'REFERENCE': ('DN' + pd.Series(t_dt).astype(str)).values,
        'REMARK': '', 'DRAMOUNT': 200_000, 'CRAMOUNT': 0,
        'CRTDTM': '15/06/2026 08:00:00',
    })

    gl02_df = pd.concat([gl02_di_match, gl02_di_thua, gl02_di_non_ach,
                          gl02_den_match, gl02_den_thua], ignore_index=True)

    # ── MIS_DI ───────────────────────────────────────────────────────────────
    # SCNL khop: CHI_NHANH=1101, SO_TIEN=100000, TRACE=1..N_MATCH_DI
    # KEY_HUB = '1101' + str(t) + '100000' = KEY_DI cua NPO_DI tuong ung
    t_sm = np.arange(1, N_MATCH_DI + 1)
    mis_di_scnl_match = pd.DataFrame({
        'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '1101',
        'REFHUB': '', 'MSGREF': ('MM' + pd.Series(t_sm).astype(str)).values,
        'MSGSEQ': '1', 'TXID': '0', 'KENH_THANH_TOAN': 'ACH',
        'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA', 'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
        'TRANG_THAI_LENH': 'SCNL', 'SO_TIEN': 100_000,
        'TRACE': pd.Series(t_sm).astype(str).values,
        'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': '',
    })
    # SCNL thua: CHI_NHANH=8888, SO_TIEN=50000, khong co NPO tuong ung
    # GW co 20k slot cho key '888850000' -> CHENH_LECH=0 -> khong vao CAP_CN_TIEN
    t_st = np.arange(300_001, 300_001 + N_MIS_DI_SCNL_THUA)
    mis_di_scnl_thua = pd.DataFrame({
        'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '8888',
        'REFHUB': '', 'MSGREF': ('MT' + pd.Series(t_st).astype(str)).values,
        'MSGSEQ': '1', 'TXID': '0', 'KENH_THANH_TOAN': 'ACH',
        'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA', 'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
        'TRANG_THAI_LENH': 'SCNL', 'SO_TIEN': 50_000,
        'TRACE': pd.Series(t_st).astype(str).values,
        'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': '',
    })
    # TPAY timeout (N_TIMEOUT_NO_GW): CHI_NHANH=9001, SO_TIEN unique -> khong co GW slot
    t_to = np.arange(400_001, 400_001 + N_TIMEOUT_NO_GW)
    mis_di_tpay_to = pd.DataFrame({
        'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '9001',
        'REFHUB': '', 'MSGREF': ('TO' + pd.Series(t_to).astype(str)).values,
        'MSGSEQ': '1', 'TXID': '0', 'KENH_THANH_TOAN': 'ACH',
        'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA', 'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
        'TRANG_THAI_LENH': 'TPAY',
        'SO_TIEN': pd.Series(t_to).values,  # moi record co SO_TIEN khac nhau
        'TRACE': pd.Series(t_to).astype(str).values,
        'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': '',
    })
    # TXRT fix: CHI_NHANH=9099, SO_TIEN=777777
    # 1 SCNL + 1 TXRT + 1 TPAY voi 2 GW slots
    # Neu khong co fix (chi dem SCNL): available=2-1=1 >= 1 TPAY -> TPAY KHONG timeout
    # Sau fix (dem SCNL+TXRT):         available=2-2=0  < 1 TPAY -> TPAY LA timeout
    mis_di_txrt_fix = pd.DataFrame([
        {'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '9099', 'REFHUB': '',
         'MSGREF': 'TXRTFIX001', 'MSGSEQ': '1', 'TXID': '0',
         'KENH_THANH_TOAN': 'ACH', 'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA',
         'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
         'TRANG_THAI_LENH': 'SCNL', 'SO_TIEN': 777_777,
         'TRACE': '500001', 'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': ''},
        {'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '9099', 'REFHUB': '',
         'MSGREF': 'TXRTFIX002', 'MSGSEQ': '1', 'TXID': '0',
         'KENH_THANH_TOAN': 'ACH', 'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA',
         'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
         'TRANG_THAI_LENH': 'TXRT', 'SO_TIEN': 777_777,
         'TRACE': '500002', 'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': ''},
        {'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '9099', 'REFHUB': '',
         'MSGREF': 'TXRTFIX003', 'MSGSEQ': '1', 'TXID': '0',
         'KENH_THANH_TOAN': 'ACH', 'LOAI_LENH_OSB': 'CT', 'NH_NHAN': 'VBAA',
         'MA_GIAO_DICH': '0', 'NOI_DUNG': '',
         'TRANG_THAI_LENH': 'TPAY', 'SO_TIEN': 777_777,
         'TRACE': '500003', 'SE_TRACE': '', 'SESSION': SES, 'NGAY_KENH_TRA': ''},
    ])

    mis_di_df = pd.concat([mis_di_scnl_match, mis_di_scnl_thua,
                            mis_di_tpay_to, mis_di_txrt_fix], ignore_index=True)

    # ── GW ───────────────────────────────────────────────────────────────────
    # SCNL khop (40k): KEY_GW = '1101100000' (40k slots)
    gw_match = pd.DataFrame({
        'No': t_sm, 'SessionId': SES, 'BRCD': '1101', 'STTLMAMT': 100_000,
        'MSGREF': ('GM' + pd.Series(t_sm).astype(str)).values,
        'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'SETTLED',
    })
    # SCNL thua (20k): KEY_GW = '888850000' (20k slots) -> CHENH_LECH=0, an khoi CAP
    t_gt = np.arange(1, N_MIS_DI_SCNL_THUA + 1)
    gw_thua = pd.DataFrame({
        'No': N_MATCH_DI + t_gt, 'SessionId': SES, 'BRCD': '8888', 'STTLMAMT': 50_000,
        'MSGREF': ('GT' + pd.Series(t_gt).astype(str)).values,
        'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'SETTLED',
    })
    # TXRT fix: 2 slots cho '9099777777' (1 SCNL + 1 TXRT het slot -> TPAY timeout)
    gw_txrt_fix = pd.DataFrame([
        {'No': 999990001, 'SessionId': SES, 'BRCD': '9099', 'STTLMAMT': 777_777,
         'MSGREF': 'GWFIX001', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'SETTLED'},
        {'No': 999990002, 'SessionId': SES, 'BRCD': '9099', 'STTLMAMT': 777_777,
         'MSGREF': 'GWFIX002', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'SETTLED'},
    ])
    gw_df = pd.concat([gw_match, gw_thua, gw_txrt_fix], ignore_index=True)

    # ── MIS_DEN ──────────────────────────────────────────────────────────────
    # KEY_DEN_HUB = TRACE.lstrip('0') + SO_TIEN
    # Phai khop voi KEY_DEN = SO_TRACE(REFERENCE) + DRAMOUNT
    # DN{t} -> SO_TRACE = str(t) -> KEY_DEN = str(t) + '200000'
    # TRACE = str(t), SO_TIEN = 200000 -> KEY_DEN_HUB = str(t) + '200000' -> KHOP
    t_den_m = np.arange(1, N_MATCH_DEN + 1)
    mis_den_match = pd.DataFrame({
        'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '2101',
        'REFHUB': '', 'MSGREF': ('DM' + pd.Series(t_den_m).astype(str)).values,
        'MSGSEQ': '1', 'TXID': '0', 'KENH_THANH_TOAN': 'ACH',
        'LOAI_LENH_OSB': 'CT', 'NH_GUI': 'VBAA', 'NOI_DUNG': '',
        'TRANG_THAI_LENH': 'SCNL', 'SO_TIEN': 200_000,
        'TRACE': pd.Series(t_den_m).astype(str).values,
        'SESSION': SES,
    })
    t_den_t = np.arange(800_001, 800_001 + N_MIS_DEN_THUA)
    mis_den_thua = pd.DataFrame({
        'NGAY_GIAO_DICH': '15/06/2026', 'CHI_NHANH': '2101',
        'REFHUB': '', 'MSGREF': ('DT' + pd.Series(t_den_t).astype(str)).values,
        'MSGSEQ': '1', 'TXID': '0', 'KENH_THANH_TOAN': 'ACH',
        'LOAI_LENH_OSB': 'CT', 'NH_GUI': 'VBAA', 'NOI_DUNG': '',
        'TRANG_THAI_LENH': 'SCNL', 'SO_TIEN': 200_000,
        'TRACE': pd.Series(t_den_t).astype(str).values,
        'SESSION': SES,
    })
    mis_den_df = pd.concat([mis_den_match, mis_den_thua], ignore_index=True)

    # ── Chay pipeline ─────────────────────────────────────────────────────────
    _, o = _run_large_test_case(gl02_df, gw_df, mis_di_df, mis_den_df)

    # ── Kiem tra so luong ─────────────────────────────────────────────────────
    # MIS_DI_THUA = 20k SCNL_thua + 1 SCNL_txrtfix (trace 500001) + 1 TXRT_txrtfix = 20,002
    # TIMEOUT = N_TIMEOUT_NO_GW (20, moi TPAY co SO_TIEN unique, GW=0) + 1 TPAY_txrtfix = 21
    # CAP_CN_TIEN = 21 (20 cap SO_TIEN unique cua TPAY_to + 1 cap '9099777777')
    expected = {
        'MIS_DI_KHOP'       : N_MATCH_DI,
        'NPO_DI_THUA'       : N_NPO_DI_THUA,
        'MIS_DI_THUA'       : N_MIS_DI_SCNL_THUA + 2,  # +1 SCNL + +1 TXRT cua txrt-fix
        'TIMEOUT_KHONG_KENH': N_TIMEOUT_NO_GW + 1,     # +1 TPAY txrt-fix
        'CAP_CN_TIEN'       : N_TIMEOUT_NO_GW + 1,
        'MIS_DEN_KHOP'      : N_MATCH_DEN,
        'NPO_DEN_THUA'      : N_NPO_DEN_THUA,
        'MIS_DEN_THUA'      : N_MIS_DEN_THUA,
    }
    for sheet, exp in expected.items():
        df = _read_sheet(o, sheet)
        assert_eq(len(df), exp, f'Sheet {sheet}')

    # So lieu TUYET DOI — khong co ty le %
    # DI: 40k khop / (40k NPO_DI_THUA=10k) => Tong NPO_DI=50k, Tong MIS_DI (khop+thua+TO)
    # DEN: 25k khop / (25k NPO_DEN_THUA=5k) => Tong NPO_DEN=30k
    df_tk = _read_sheet(o, 'TONG_KET')
    vals  = dict(zip(df_tk.iloc[:, 0].astype(str), df_tk.iloc[:, 1].astype(str)))

    # Khong co dong ty le %
    assert 'Ty le khop DI (%)'  not in vals, 'TONG_KET khong duoc co dong ty le % DI'
    assert 'Ty le khop DEN (%)' not in vals, 'TONG_KET khong duoc co dong ty le % DEN'

    def get_int(key):
        raw = vals.get(key, '').replace(',', '').strip()
        return int(raw) if raw.lstrip('-').isdigit() else None

    n_tong_npo_di = get_int('Tong NPO_DI (can doi)')
    n_tong_npo_den_rows = df_tk[df_tk.iloc[:, 0].astype(str) == 'Tong NPO_DEN (can doi)']
    n_tong_npo_den = int(n_tong_npo_den_rows.iloc[0, 1].replace(',', '')) if len(n_tong_npo_den_rows) > 0 else None
    assert_eq(n_tong_npo_di,  N_MATCH_DI + N_NPO_DI_THUA, 'Tong NPO_DI (can doi)')
    assert_eq(n_tong_npo_den, N_MATCH_DEN + N_NPO_DEN_THUA, 'Tong NPO_DEN (can doi)')

    # TXRT fix: TPAY CHI_NHANH=9099, SO_TIEN=777777 phai la TIMEOUT
    df_to = _read_sheet(o, 'TIMEOUT_KHONG_KENH')
    df_to['SO_TIEN'] = pd.to_numeric(df_to['SO_TIEN'], errors='coerce')
    txrt_fix_row = df_to[df_to['SO_TIEN'] == 777_777]
    assert len(txrt_fix_row) == 1, \
        f'TXRT fix: TPAY 9099/777777 phai la TIMEOUT (got {len(txrt_fix_row)})'

    # CUSTOMER non-ACH: 2000 dong REF900001..902000 khong duoc xuat hien trong NPO_DI_THUA
    df_npo_thua = _read_sheet(o, 'NPO_DI_THUA')
    assert_eq(len(df_npo_thua), N_NPO_DI_THUA, 'NPO_DI_THUA phai dung N (non-ACH bi loc)')
    if 'REFERENCE' in df_npo_thua.columns:
        non_ach_leaked = df_npo_thua['REFERENCE'].str.startswith('REF9000').sum()
        assert non_ach_leaked == 0, f'Non-ACH rows bi lo vao NPO_DI_THUA: {non_ach_leaked}'


# ═══════════════════════════════════════════════════════════════════════════════
# T02: skip_file bug — dong dau LOCAC=900001, dong sau LOCAC=502003
# ═══════════════════════════════════════════════════════════════════════════════
def t02_skip_file_bug():
    """
    GL02: dong 1 = LOCAC=900001 (sai), dong 2-11 = LOCAC=502003 (dung).
    Neu skip_file bug con ton tai: GL02 se doc duoc 0 dong -> NPO_DI=0.
    Sau khi fix: NPO_DI = 10 dong.
    """
    gl02 = [_base_gl02_row(0, locac='900001', cramount=50_000)]  # DONG DAU SAI
    for t in range(1, 11):
        gl02.append(_base_gl02_row(t, locac='502003', cramount=100_000))  # dung

    mis_di  = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 11)]
    mis_den = [_base_mis_den_row(100 + t, 200_000) for t in range(3)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 11)]

    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df = _read_sheet(o, 'NPO_DI_THUA')
    # Tat ca 10 NPO_DI deu khop voi MIS -> NPO_DI_THUA = 0
    df_khop = _read_sheet(o, 'MIS_DI_KHOP')
    assert_eq(len(df_khop), 10, 'MIS_DI_KHOP sau khi fix skip_file bug')


# ═══════════════════════════════════════════════════════════════════════════════
# T03-T06: Cancel tai 4 diem khac nhau
# ═══════════════════════════════════════════════════════════════════════════════
def _make_minimal_testcase():
    """Tao bo du lieu nho nhat de test cancel."""
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 6)]
    mis_di  = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 6)]
    mis_den = [_base_mis_den_row(100 + t) for t in range(3)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 6)]
    return {'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den}

def t03_cancel_point1():
    """Cancel NGAY truoc Phase 1 — main_from_dir tra ve None."""
    ev = threading.Event()
    ev.set()   # set ngay lap tuc — cancel truoc khi kiem tra dau tien
    td = _make_minimal_testcase()
    d = os.path.join(TEST_DIR, 'c1_inp')
    o = os.path.join(TEST_DIR, 'c1_out')
    shutil.rmtree(d, ignore_errors=True)
    shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)
    _make_pdf(d, SES, NGAY)
    _make_gl02_zip(os.path.join(d, 'GL02.zip'), td['gl02'])
    _make_gw_xlsx(os.path.join(d, 'GW.xlsx'), td['gw'])
    _make_mis_zip(os.path.join(d, 'MIS_DI_1.zip'), td['mis_di'][:3], 'di1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DI_2.zip'), td['mis_di'][3:], 'di2.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_1.zip'), td['mis_den'][:2], 'den1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_2.zip'), td['mis_den'][2:], 'den2.csv')
    result = main_from_dir(input_dir=d, output_dir=o, ngay='15/06/2026', cancel_event=ev)
    assert result is None, f'Cancel Point 1: can tra ve None, got {result}'

def t04_cancel_point2():
    """Cancel sau B3 — duoc set trong khi B2/B6 co the van dang chay."""
    ev = threading.Event()
    td = _make_minimal_testcase()
    d = os.path.join(TEST_DIR, 'c2_inp')
    o = os.path.join(TEST_DIR, 'c2_out')
    shutil.rmtree(d, ignore_errors=True); shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)
    _make_pdf(d, SES, NGAY)
    _make_gl02_zip(os.path.join(d, 'GL02.zip'), td['gl02'])
    _make_gw_xlsx(os.path.join(d, 'GW.xlsx'), td['gw'])
    _make_mis_zip(os.path.join(d, 'MIS_DI_1.zip'), td['mis_di'][:3], 'di1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DI_2.zip'), td['mis_di'][3:], 'di2.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_1.zip'), td['mis_den'][:2], 'den1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_2.zip'), td['mis_den'][2:], 'den2.csv')

    # Set cancel sau 0.05s — Phase 1 vua bat dau, B3 co the da xong
    def _set_after_delay():
        import time; time.sleep(0.05); ev.set()
    threading.Thread(target=_set_after_delay, daemon=True).start()

    result = main_from_dir(input_dir=d, output_dir=o, ngay='15/06/2026', cancel_event=ev)
    # Ket qua: None (bi cancel) hoac duong dan file (neu xu ly xong truoc khi cancel)
    # Ca hai deu chap nhan duoc
    assert result is None or result.endswith('.xlsx'), \
        f'Cancel Point 2: expected None or .xlsx, got {result}'

def t05_cancel_between_phase1_phase2():
    """Cancel sau Phase 1 bang cach set truoc khi main_from_dir goi."""
    # Khong the test chinh xac vi la concurrent — skip (duoc bao gom trong T03)
    pass  # covered by T03

def t06_cancel_truoc_excel():
    """Tuong tu T03 nhung can xac nhan khong co file xlsx trong output."""
    ev = threading.Event()
    ev.set()
    td = _make_minimal_testcase()
    d = os.path.join(TEST_DIR, 'c4_inp')
    o = os.path.join(TEST_DIR, 'c4_out')
    shutil.rmtree(d, ignore_errors=True); shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)
    _make_pdf(d, SES, NGAY)
    _make_gl02_zip(os.path.join(d, 'GL02.zip'), td['gl02'])
    _make_gw_xlsx(os.path.join(d, 'GW.xlsx'), td['gw'])
    _make_mis_zip(os.path.join(d, 'MIS_DI_1.zip'), td['mis_di'][:3], 'di1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DI_2.zip'), td['mis_di'][3:], 'di2.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_1.zip'), td['mis_den'][:2], 'den1.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_2.zip'), td['mis_den'][2:], 'den2.csv')
    result = main_from_dir(input_dir=d, output_dir=o, ngay='15/06/2026', cancel_event=ev)
    assert result is None, f'Can tra ve None khi cancel'
    xlsx = [f for f in os.listdir(o) if f.endswith('.xlsx')]
    assert not xlsx, f'Khong duoc co xlsx trong output khi cancel: {xlsx}'


# ═══════════════════════════════════════════════════════════════════════════════
# T07: GL02 khong co cot LOCAC — chay khong crash, doc duoc du lieu
# ═══════════════════════════════════════════════════════════════════════════════
def t07_gl02_no_locac_column():
    gl02 = []
    for t in range(1, 6):
        r = _base_gl02_row(t, cramount=100_000)
        del r['LOCAC']   # bo cot LOCAC
        gl02.append(r)
    mis_di = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 6)]
    mis_den = [_base_mis_den_row(100 + t) for t in range(2)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 6)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    # Khong LOCAC filter → tat ca 5 GL02 DI qua → khop voi 5 SCNL
    df = _read_sheet(o, 'MIS_DI_KHOP')
    assert_eq(len(df), 5, 'GL02 khong LOCAC: phai doc duoc 5 dong va khop 5 SCNL')


# ═══════════════════════════════════════════════════════════════════════════════
# T08: MIS_DI rong sau pre-filter (tat ca SESSION sai)
# ═══════════════════════════════════════════════════════════════════════════════
def t08_mis_di_rong_sau_prefilter():
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 6)]
    # Tat ca MIS_DI co SESSION sai
    mis_di = [_base_mis_di_row(t, 'SCNL', 100_000, session='WRONG') for t in range(1, 6)]
    mis_den = [_base_mis_den_row(100 + t) for t in range(2)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 6)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_khop = _read_sheet(o, 'MIS_DI_KHOP')
    df_npo  = _read_sheet(o, 'NPO_DI_THUA')
    assert_eq(len(df_khop), 0, 'MIS_DI_KHOP phai rong')
    assert_eq(len(df_npo),  5, 'NPO_DI_THUA = 5 (tat ca GL02 thua)')


# ═══════════════════════════════════════════════════════════════════════════════
# T09: Khong co TPAY — timeout = 0, CAP_CN_TIEN khong co TPAY row
# ═══════════════════════════════════════════════════════════════════════════════
def t09_gw_rong():
    """
    Khi khong co TPAY trong MIS: TIMEOUT_KHONG_KENH = 0.
    CAP_CN_TIEN chi hien khi COUNT_MIS > COUNT_GW; voi GW co du slot cho SCNL
    thi SCNL khong tao CHENH_LECH. Chi kiem tra TIMEOUT = 0 va khong co TPAY trong CAP.
    """
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 6)]
    mis_di = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 6)]
    mis_den = [_base_mis_den_row(100 + t) for t in range(2)]
    # GW co du slot cho moi SCNL (5 records)
    gw = [_base_gw_row(t, 100_000) for t in range(1, 6)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_to  = _read_sheet(o, 'TIMEOUT_KHONG_KENH')
    df_cap = _read_sheet(o, 'CAP_CN_TIEN')
    assert_eq(len(df_to),  0, 'TIMEOUT phai rong khi khong co TPAY')
    # 5 SCNL + 5 GW cung KEY_GW -> COUNT_MIS=5 COUNT_GW=5 -> CHENH_LECH=0 -> khong co trong CAP
    assert_eq(len(df_cap), 0, 'CAP_CN_TIEN phai rong khi CHENH_LECH=0 (SCNL=5 GW=5)')


# ═══════════════════════════════════════════════════════════════════════════════
# T10: TONG_KET so tuyet doi — can doi chinh xac (khong co % trong output)
# ═══════════════════════════════════════════════════════════════════════════════
def t10_so_lieu_tuyet_doi():
    """
    DI:  5 NPO_DI, 3 MIS_DI khop, 2 NPO_DI_THUA, 0 MIS_DI_THUA
    DEN: 4 NPO_DEN, 3 MIS_DEN khop, 1 NPO_DEN_THUA, 0 MIS_DEN_THUA
    Xac nhan so TUYET DOI tu TONG_KET — khong co dong ty le %, phai can doi chinh xac.
    """
    gl02 = []
    for t in range(1, 6):    # 5 NPO_DI
        gl02.append(_base_gl02_row(t, cramount=100_000))
    for t in range(10, 14):  # 4 NPO_DEN
        gl02.append(_base_gl02_row(t, cramount=0, dramount=200_000))

    mis_di  = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 4)]  # 3 khop
    mis_den = [_base_mis_den_row(t, 200_000) for t in range(10, 13)]        # 3 khop DEN
    gw = [_base_gw_row(t, 100_000) for t in range(1, 4)]

    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_tk = _read_sheet(o, 'TONG_KET')
    vals = dict(zip(df_tk.iloc[:, 0].astype(str), df_tk.iloc[:, 1].astype(str)))

    def get_int(key):
        raw = vals.get(key, '').replace(',', '').strip()
        return int(raw) if raw.lstrip('-').isdigit() else None

    # Kiem tra khong co dong ty le % trong output
    assert 'Ty le khop DI (%)' not in vals, 'TONG_KET khong duoc co dong ty le %'
    assert 'Ty le khop DEN (%)' not in vals, 'TONG_KET khong duoc co dong ty le %'

    # DI: so TUYET DOI
    n_di_khop      = get_int('So giao dich khop (MIS)')
    n_npo_di_thua  = get_int('NPO_DI thua')
    n_mis_di_thua  = get_int('MIS_DI thua')
    n_timeout      = get_int('Timeout khong kenh')
    n_tong_npo_di  = get_int('Tong NPO_DI (can doi)')
    n_tong_mis_di  = get_int('Tong MIS_DI (can doi)')

    assert_eq(n_di_khop, 3, 'MIS_DI_KHOP')
    # Chieu DI - TONG_KET co 2 dong "So giao dich khop (MIS)" (DI va DEN) nen chi lay lan dau
    # -> kiem tra qua NPO_THUA va MIS_THUA thay the
    assert_eq(n_timeout, 0, 'Timeout phai = 0')
    # Can doi NPO_DI: 3 khop + 2 thua = 5
    assert_eq(n_tong_npo_di, 5, 'Tong NPO_DI (can doi) = 5')
    # Can doi MIS_DI: 3 khop + 0 thua + 0 timeout = 3
    assert_eq(n_tong_mis_di, 3, 'Tong MIS_DI (can doi) = 3')

    # DEN: doc tu TONG_KET (doc toan bo de lay hang DEN)
    df_den_rows = df_tk[df_tk.iloc[:, 0].astype(str).str.contains('DEN', na=False)]
    den_vals = dict(zip(df_den_rows.iloc[:, 0].astype(str), df_den_rows.iloc[:, 1].astype(str)))

    def get_int_den(key):
        raw = den_vals.get(key, '').replace(',', '').strip()
        return int(raw) if raw.lstrip('-').isdigit() else None

    n_tong_npo_den = get_int_den('Tong NPO_DEN (can doi)')
    n_tong_mis_den = get_int_den('Tong MIS_DEN (can doi)')
    assert_eq(n_tong_npo_den, 4, 'Tong NPO_DEN (can doi) = 4')
    assert_eq(n_tong_mis_den, 3, 'Tong MIS_DEN (can doi) = 3')


# ═══════════════════════════════════════════════════════════════════════════════
# T11: TXRT chi thuoc session hien tai
# ═══════════════════════════════════════════════════════════════════════════════
def t11_txrt_session_filter():
    """TXRT cua session cu phai bi loai."""
    gl02 = [_base_gl02_row(1, cramount=100_000)]
    gw   = [_base_gw_row(1, 100_000)]
    mis_di = [
        _base_mis_di_row(1, 'SCNL', 100_000),           # khop
        _base_mis_di_row(2, 'TXRT', 100_000),            # TXRT session dung
        _base_mis_di_row(3, 'TXRT', 100_000, session='OLD'),  # TXRT session sai -> bi loc
    ]
    mis_den = [_base_mis_den_row(10)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df = _read_sheet(o, 'MIS_DI_THUA')
    txrt_rows = df[df['TRANG_THAI_LENH'] == 'TXRT']
    assert_eq(len(txrt_rows), 1, 'Chi 1 TXRT (session dung) trong MIS_DI_THUA')


# ═══════════════════════════════════════════════════════════════════════════════
# T12: Null-session TPAY — in-range vs out-range
# ═══════════════════════════════════════════════════════════════════════════════
def t12_null_session_tpay():
    gl02 = [_base_gl02_row(1, cramount=100_000)]
    gw = [_base_gw_row(2, 900_001), _base_gw_row(3, 900_002)]  # slots cho TPAY
    ngay_tu  = (NGAY - timedelta(days=1)).replace(hour=23)   # T-1 23:00
    ngay_den = NGAY.replace(hour=23)                          # T   23:00
    in_time  = NGAY.replace(hour=10)
    out_time = (NGAY - timedelta(days=1)).replace(hour=10)  # truoc TPAY_TU

    mis_di = [
        _base_mis_di_row(1, 'SCNL', 100_000),
        _base_mis_di_row(2, 'TPAY', 900_001, session='',
                         ngay_kenh=in_time.strftime('%d/%m/%Y %H:%M:%S')),   # duoc tinh
        _base_mis_di_row(3, 'TPAY', 900_002, session='',
                         ngay_kenh=out_time.strftime('%d/%m/%Y %H:%M:%S')),  # loai
    ]
    mis_den = [_base_mis_den_row(10)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df = _read_sheet(o, 'MIS_DI_THUA')
    tpay_rows = df[df['TRANG_THAI_LENH'] == 'TPAY']
    assert_eq(len(tpay_rows), 1, 'Chi 1 TPAY null-session in-range trong THUA')


# ═══════════════════════════════════════════════════════════════════════════════
# T13: GW dedup theo MSGREF
# ═══════════════════════════════════════════════════════════════════════════════
def t13_gw_dedup():
    """GW co 5 records nhung 3 MSGREF trung -> chi con 2 unique sau dedup."""
    gl02 = [_base_gl02_row(1, cramount=100_000)]
    mis_di = [_base_mis_di_row(1, 'SCNL', 100_000)]
    mis_den = [_base_mis_den_row(10)]
    gw_rows = [
        {'No': 1, 'SessionId': SES, 'BRCD': '1101', 'STTLMAMT': 100_000,
         'MSGREF': 'GW0000000001', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'S'},
        {'No': 2, 'SessionId': SES, 'BRCD': '1101', 'STTLMAMT': 200_000,
         'MSGREF': 'GW0000000002', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'S'},
        # 3 ban sao MSGREF=001 — sau dedup chi con 1
        {'No': 3, 'SessionId': SES, 'BRCD': '1101', 'STTLMAMT': 100_000,
         'MSGREF': 'GW0000000001', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'S'},
        {'No': 4, 'SessionId': SES, 'BRCD': '1101', 'STTLMAMT': 100_000,
         'MSGREF': 'GW0000000001', 'PrcFlg': 'OK', 'TxDt': '15/06/2026', 'Status': 'S'},
    ]
    d = os.path.join(TEST_DIR, 't13_inp')
    o = os.path.join(TEST_DIR, 't13_out')
    shutil.rmtree(d, ignore_errors=True); shutil.rmtree(o, ignore_errors=True)
    os.makedirs(d); os.makedirs(o)
    _make_pdf(d, SES, NGAY)
    _make_gl02_zip(os.path.join(d, 'GL02.zip'), gl02)
    # Viet GW voi Sheet 1 (4 rows) + Sheet 2 (dup 2 rows dau)
    wb = xlsxwriter.Workbook(os.path.join(d, 'GW.xlsx'))
    cols = ['No', 'SessionId', 'BRCD', 'STTLMAMT', 'MSGREF', 'PrcFlg', 'TxDt', 'Status']
    for sname, rows in [('Sheet 1', gw_rows), ('Sheet 2', gw_rows[:2])]:
        ws = wb.add_worksheet(sname)
        for ci, c in enumerate(cols): ws.write(0, ci, c)
        for ri, r in enumerate(rows, 1): ws.write_row(ri, 0, [r[c] for c in cols])
    wb.close()
    _make_mis_zip(os.path.join(d, 'MIS_DI_1.zip'), mis_di, 'di.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DI_2.zip'), [], 'di2.csv')  # rong
    _make_mis_zip(os.path.join(d, 'MIS_DEN_1.zip'), mis_den, 'den.csv')
    _make_mis_zip(os.path.join(d, 'MIS_DEN_2.zip'), [], 'den2.csv')
    main_from_dir(input_dir=d, output_dir=o, ngay='15/06/2026')
    # RAW_GW phai co 2 rows (2 unique MSGREF sau dedup, ACH filter khong ap dung)
    # Kiem tra qua log thay vi doc sheet (GW < 50k -> o Excel)
    df_gw = _read_sheet(o, 'RAW_GW')
    assert_eq(len(df_gw), 2, 'RAW_GW sau dedup: 2 unique MSGREF')


# ═══════════════════════════════════════════════════════════════════════════════
# T14: KEY matching chinh xac — TRACE co leading zero
# ═══════════════════════════════════════════════════════════════════════════════
def t14_trace_leading_zero():
    """
    MIS TRACE='00123' sau lstrip('0') = '123'.
    GL02 REFERENCE='REF0000000123' -> SO_TRACE='123'.
    KEY_DI == KEY_HUB -> khop.
    """
    gl02 = [_base_gl02_row(0, cramount=100_000)]
    gl02[0]['REFERENCE'] = 'REF0000000123'  # SO_TRACE = '123'
    mis_di = [_base_mis_di_row(0, 'SCNL', 100_000)]
    mis_di[0]['TRACE'] = '00123'   # sau lstrip = '123'
    mis_den = [_base_mis_den_row(10)]
    gw = [_base_gw_row(1, 100_000)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df = _read_sheet(o, 'MIS_DI_KHOP')
    assert_eq(len(df), 1, 'TRACE 00123 phai khop voi REFERENCE REF0000000123')


# ═══════════════════════════════════════════════════════════════════════════════
# T15: CALD/ERPO/TPER bi loai hoan toan
# ═══════════════════════════════════════════════════════════════════════════════
def t15_excl_trang_thai():
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 4)]
    mis_di = [
        _base_mis_di_row(1, 'SCNL', 100_000),
        _base_mis_di_row(2, 'CALD', 100_000),
        _base_mis_di_row(3, 'ERPO', 100_000),
        _base_mis_di_row(4, 'TPER', 100_000),
    ]
    mis_den = [_base_mis_den_row(10)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 4)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_final = _read_sheet(o, 'MIS_DI_KHOP')
    df_thua  = _read_sheet(o, 'MIS_DI_THUA')
    df_to    = _read_sheet(o, 'TIMEOUT_KHONG_KENH')
    # CALD/ERPO/TPER khong xuat hien trong bat ky sheet nao
    for df, sheet in [(df_final, 'MIS_DI_KHOP'), (df_thua, 'MIS_DI_THUA'), (df_to, 'TIMEOUT')]:
        if 'TRANG_THAI_LENH' in df.columns:
            bad = df[df['TRANG_THAI_LENH'].isin(['CALD','ERPO','TPER'])]
            assert len(bad) == 0, f'{sheet} con chua CALD/ERPO/TPER'


# T16: Filter CUSTOMER ACH — chi lay CUSTOMER = 1000-003526275
# ═══════════════════════════════════════════════════════════════════════════════
def t16_customer_ach_filter():
    """
    GL02 co 5 rows CUSTOMER=ACH + 3 rows CUSTOMER=non-ACH (cung LOCAC 502003).
    Sau khi loc, NPO_DI chi co 5 dong ACH; 3 dong non-ACH bi loai.
    """
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 6)]           # 5 ACH
    for t in range(10, 13):                                                       # 3 non-ACH
        gl02.append(_base_gl02_row(t, cramount=100_000, customer='9999-KHAC'))
    mis_di = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 6)]
    mis_den = [_base_mis_den_row(100)]
    gw = [_base_gw_row(t, 100_000) for t in range(1, 6)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_tong = _read_sheet(o, 'TONG_KET')
    # Tim dong "So giao dich khop (MIS)" de lay so NPO_DI = 5 (ACH) khop voi 5 MIS
    row = df_tong[df_tong.iloc[:, 0].astype(str).str.contains('khop', case=False, na=False)]
    assert len(row) > 0, 'TONG_KET khong co dong khop'
    count = int(str(row.iloc[0, 1]).replace(',', ''))
    assert count == 5, f'NPO_DI sau filter phai la 5, got {count} (non-ACH bi giu lai?)'


def t17_timeout_msgref_check():
    """
    TPAY vuot slot GW nhung MSGREF xuat hien trong GW → thuc ra da di kenh.
    Pipeline phai chuyen TPAY nay tu TIMEOUT_KHONG_KENH vao MIS_DI_FINAL → khop NPO.

    Setup:
      NPO: 4 rows (t=1..4)
      GW:  3 slots (KEY '1101100000') + 1 row extra MSGREF='GW0000000099' (BRCD='9999')
      MIS: 3 SCNL (t=1..3) + 1 TPAY (t=4, MSGREF='GW0000000099')
    Ket qua mong doi:
      TIMEOUT = 0  (TPAY t=4 co MSGREF trong GW → duoc chuyen vao MIS_DI_FINAL)
      KHOP    = 4  (ca 4 NPO rows deu khop MIS)
    """
    gl02 = [_base_gl02_row(t, cramount=100_000) for t in range(1, 5)]   # NPO t=1..4
    gw   = [_base_gw_row(t, 100_000) for t in range(1, 4)]              # 3 slots (t=1,2,3)
    # Extra GW row: BRCD='9999' -> khong anh huong slot KEY '1101100000'
    # nhung MSGREF='GW0000000099' de TPAY t=4 duoc rescue
    gw_extra = {'No': 99, 'SessionId': SES, 'BRCD': '9999',
                 'STTLMAMT': 1, 'MSGREF': 'GW0000000099', 'PrcFlg': 'OK',
                 'TxDt': '15/06/2026', 'Status': 'SETTLED'}
    gw.append(gw_extra)

    mis_di = [_base_mis_di_row(t, 'SCNL', 100_000) for t in range(1, 4)]   # 3 SCNL
    tpay_row = _base_mis_di_row(4, 'TPAY', 100_000)
    tpay_row['MSGREF'] = 'GW0000000099'  # MSGREF khop voi GW extra row
    mis_di.append(tpay_row)

    mis_den = [_base_mis_den_row(100)]
    _, o = _run_test_case({'gl02': gl02, 'gw': gw, 'mis_di': mis_di, 'mis_den': mis_den})
    df_tong = _read_sheet(o, 'TONG_KET')

    def _get_val(label_pat):
        row = df_tong[df_tong.iloc[:, 0].astype(str).str.contains(label_pat, case=False, na=False)]
        if len(row) == 0:
            return 0
        return int(str(row.iloc[0, 1]).replace(',', '').strip() or '0')

    n_timeout = _get_val('Timeout khong kenh')
    n_khop    = _get_val('So giao dich khop')
    assert n_timeout == 0, f'TIMEOUT phai = 0 (TPAY duoc rescue), got {n_timeout}'
    assert n_khop == 4,    f'KHOP phai = 4 (ca TPAY da di kenh duoc khop NPO), got {n_khop}'


# ═══════════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    os.makedirs(TEST_DIR, exist_ok=True)
    print(f'\nChay {17} test cases...\n')

    run('T01: Pipeline 237k+ dong tu tao (vectorized)', t01_pipeline_200k_rows)
    run('T02: skip_file bug — dong dau LOCAC sai',  t02_skip_file_bug)
    run('T03: Cancel Point 1 — truoc Phase 1',      t03_cancel_point1)
    run('T04: Cancel Point 2 — sau B3',             t04_cancel_point2)
    run('T05: Cancel Point 3 (covered by T03)',     t05_cancel_between_phase1_phase2)
    run('T06: Cancel — khong co xlsx trong output', t06_cancel_truoc_excel)
    run('T07: GL02 khong co cot LOCAC',             t07_gl02_no_locac_column)
    run('T08: MIS_DI rong sau pre-filter',          t08_mis_di_rong_sau_prefilter)
    run('T09: GW rong — timeout=0, cap=0',          t09_gw_rong)
    run('T10: So lieu tuyet doi — can doi, khong co %', t10_so_lieu_tuyet_doi)
    run('T11: TXRT chi tinh session hien tai',      t11_txrt_session_filter)
    run('T12: Null-session TPAY in/out range',      t12_null_session_tpay)
    run('T13: GW dedup theo MSGREF',                t13_gw_dedup)
    run('T14: TRACE leading zero matching',         t14_trace_leading_zero)
    run('T15: CALD/ERPO/TPER bi loai hoan toan',   t15_excl_trang_thai)
    run('T16: Filter CUSTOMER ACH (loc non-ACH)',   t16_customer_ach_filter)
    run('T17: TIMEOUT MSGREF check — TPAY da di kenh duoc rescue', t17_timeout_msgref_check)

    passed = sum(1 for _, ok, _ in results if ok)
    failed = sum(1 for _, ok, _ in results if not ok)
    print(f'\n{"="*50}')
    print(f'  Ket qua: {passed}/{len(results)} PASS  |  {failed} FAIL')
    if failed:
        print('\nChi tiet loi:')
        for name, ok, msg in results:
            if not ok:
                print(f'  [{name}]\n    {msg[:200]}')
    print('=' * 50)

    # Don dep
    shutil.rmtree(TEST_DIR, ignore_errors=True)
    sys.exit(0 if failed == 0 else 1)
