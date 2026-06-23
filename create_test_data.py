"""
create_test_data.py — Tao du lieu test tong hop ~200k dong, day du trang thai.

Thiet ke:
- SCNL va TPAY dung amount KHAC NHAU hoan toan -> khong tranh GW slot
- TPAY_NORM: amount [600001..601000] (ngoai AMOUNTS list) -> moi cap chi co 1 GW slot
- TIMEOUT_BA: amount [700001..700020] -> GW=1, MIS=2 -> 1 timeout/cap
- NULL_IN: amount [601001..601300] -> GW=1, tpay=1 -> 0 timeout

Ket qua mong doi:
  B4:  TIMEOUT=20, Final=72,120
  B5:  KHOP=30k, NPO_THUA=10k, MIS_THUA=42,120
  B7:  KHOP=20k, NPO_THUA=5k, MIS_THUA=15k
  TIMEOUT sheet: 20 dong, CAP_CN_TIEN: 20 dong (CHENH_LECH=1)

Chay: python create_test_data.py [--clean]
"""
import io
import os
import sys
import random
import shutil
from datetime import datetime, timedelta
from collections import defaultdict

import pandas as pd
import pyzipper
import xlsxwriter

# ─── Config ───────────────────────────────────────────────────────────────────
SEED        = 42
SESSION_ID  = '99001'
SESSION_BAD = '77777'
NGAY_DT     = datetime(2026, 6, 15)
NGAY_TRUOC  = NGAY_DT - timedelta(days=1)
TPAY_TU     = NGAY_TRUOC.replace(hour=23)
TPAY_DEN    = NGAY_DT.replace(hour=23)
ZIP_PW      = b'DACwLdHi'
OUT         = 'test_input'

rng = random.Random(SEED)

if '--clean' in sys.argv:
    shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT, exist_ok=True)

# 10 branches x 7 amounts = 70 cap (chi dung cho SCNL/DEN matching)
BRANCHES = ['1101', '1102', '1201', '1301', '2101',
            '2201', '3101', '4101', '5101', '9101']
AMOUNTS  = [50_000, 100_000, 200_000, 500_000,
            1_000_000, 2_000_000, 5_000_000]

def _ba(seed):
    """Branch + amount tu random seed (chi trong BRANCHES x AMOUNTS)."""
    r = random.Random(seed)
    return r.choice(BRANCHES), r.choice(AMOUNTS)

def _tpay_ba(trace):
    """Amount RIENG cho TPAY_NORM: ngoai AMOUNTS list -> khong trung SCNL."""
    return BRANCHES[trace % len(BRANCHES)], 600_001 + (trace - 90_001)

def _null_ba(trace):
    """Amount RIENG cho null-session TPAY."""
    return BRANCHES[trace % len(BRANCHES)], 601_001 + (trace - 102_001)

def ngay_str(dt): return dt.strftime('%d/%m/%Y')
def ngay_hm(dt):  return dt.strftime('%d/%m/%Y %H:%M:%S')

def write_zip(path, df, csv_name='data.csv'):
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding='utf-8-sig')
    buf.seek(0)
    with pyzipper.AESZipFile(path, 'w',
                              compression=pyzipper.ZIP_DEFLATED,
                              encryption=pyzipper.WZ_AES) as z:
        z.setpassword(ZIP_PW)
        z.writestr(csv_name, buf.getvalue())
    print(f'  [OK] {os.path.basename(path):50s}  {len(df):>8,} dong')


# ─── Trace pools ──────────────────────────────────────────────────────────────
DI_KHOP      = list(range(10_001, 40_001))    # 30k  SCNL se khop GL02
DI_NPO_ONLY  = list(range(40_001, 50_001))    # 10k  chi co trong GL02 (NPO_DI_THUA)
DI_MIS_ONLY  = list(range(50_001, 90_001))    # 40k  chi co trong MIS (MIS_DI_THUA SCNL)
TPAY_NORM    = list(range(90_001, 91_001))    #  1k  TPAY co GW slot, khong timeout
TXRT_POOL    = list(range(100_001, 100_801))  #  800 TXRT
EXCL_POOL    = list(range(101_001, 101_401))  #  400 CALD/ERPO/TPER -> loai
NULL_IN      = list(range(102_001, 102_301))  #  300 null-session TPAY in-range
NULL_OUT     = list(range(102_301, 102_401))  #  100 null-session TPAY out-range -> loai
BAD_SES      = list(range(103_001, 113_001))  # 10k  SESSION sai -> pre-filter loai

DEN_KHOP     = list(range(200_001, 220_001))  # 20k  se khop DEN
DEN_NPO_ONLY = list(range(220_001, 225_001))  #  5k  NPO_DEN_THUA
DEN_MIS_ONLY = list(range(225_001, 240_001))  # 15k  MIS_DEN_THUA
DEN_RJCT     = list(range(240_001, 240_501))  #  500 RJCT -> loai

# TIMEOUT: 20 cap, amount 700001-700020 (ngoai AMOUNTS), GW=1, MIS=2 -> 1 timeout/cap
TIMEOUT_BA = [(BRANCHES[i % len(BRANCHES)], 700_001 + i) for i in range(20)]
TIMEOUT_TRACE_BASE = 120_001

print(f'\nSinh du lieu test: {OUT}/')
print(f'Session={SESSION_ID}, Ngay={NGAY_DT.strftime("%d/%m/%Y")}')
print('=' * 60)


# ─── 1. PDF ───────────────────────────────────────────────────────────────────
pdf_name = f'ACH_{(NGAY_DT+timedelta(1)).strftime("%Y%m%d")}_VBAAVNVN_NRT_{SESSION_ID}_N03_1.pdf'
with open(os.path.join(OUT, pdf_name), 'wb') as f:
    f.write(b'%PDF-1.4 test')
print(f'  [OK] {pdf_name}')


# ─── 2. GL02 ZIP ──────────────────────────────────────────────────────────────
_GL02_BASE = dict(TRDATE=ngay_str(NGAY_DT), USERID='USR001', JOURSEQ='1',
                  DYTRSEQ='1', CCY='VND', BUSCD='ACH', UNIT='0',
                  TRCD='ACH01', CUSTOMER='0', TRTP='1', REMARK='TEST',
                  CRTDTM=ngay_hm(NGAY_DT))

gl02_rows = []
# NPO_DI: LOCAC=502003, CRAMOUNT!=0
for t in DI_KHOP + DI_NPO_ONLY:
    b, a = _ba(t)
    gl02_rows.append({**_GL02_BASE, 'TRBRCD': b, 'LOCAC': '502003',
                      'REFERENCE': f'ACHTEST{t}', 'CRAMOUNT': a, 'DRAMOUNT': 0})

# NPO_DEN: LOCAC=502003, CRAMOUNT=0, DRAMOUNT!=0
for t in DEN_KHOP + DEN_NPO_ONLY:
    _, a = _ba(t + 500_000)
    gl02_rows.append({**_GL02_BASE, 'TRBRCD': '0000', 'LOCAC': '502003',
                      'REFERENCE': f'ACHTEST{t}', 'CRAMOUNT': 0, 'DRAMOUNT': a})

# LOCAC sai -> bi loc boi B2
for t in range(150_001, 155_001):
    b, a = _ba(t)
    gl02_rows.append({**_GL02_BASE, 'TRBRCD': b, 'LOCAC': '900001',
                      'REFERENCE': f'ACHTEST{t}', 'CRAMOUNT': a, 'DRAMOUNT': 0})

write_zip(os.path.join(OUT, 'GL02_20260615.zip'), pd.DataFrame(gl02_rows), 'GL02_20260615.csv')


# ─── 3. GW Excel (2 sheet, co dup MSGREF) ─────────────────────────────────────
# Muc tieu CAP_CN_TIEN: chi hien TIMEOUT pairs (CHENH_LECH > 0)
# -> GW can du slot cho TAT CA MIS records (SCNL + TXRT + TPAY non-timeout) + buffer
# -> TIMEOUT pairs: GW = count-1 = 2-1 = 1 -> CHENH_LECH = 1

# Buoc 1: dem truoc tat ca (branch, amount) pairs se co trong mis_di_final + df_timeout
_all_mis_ba = defaultdict(int)
for t in DI_KHOP:                         # SCNL matched
    b, a = _ba(t);             _all_mis_ba[(b, a)] += 1
for t in DI_MIS_ONLY:                     # SCNL unmatched
    b, a = _ba(t + 100);       _all_mis_ba[(b, a)] += 1
for t in TXRT_POOL:                       # TXRT
    b, a = _ba(t);             _all_mis_ba[(b, a)] += 1
for t in TPAY_NORM:                       # TPAY non-timeout (unique amounts)
    b, a = _tpay_ba(t);        _all_mis_ba[(b, a)] += 1
for t in NULL_IN:                         # null-session TPAY in-range (unique amounts)
    b, a = _null_ba(t);        _all_mis_ba[(b, a)] += 1
for b, a in TIMEOUT_BA:                   # TIMEOUT: 2 TPAY / cap (du ca mis_di_final + df_timeout)
    _all_mis_ba[(b, a)] += 2

_timeout_ba_set = set(TIMEOUT_BA)

gw_rows = []
_gw_seq = [1]

def add_gw(b, a, prc='OK'):
    s = _gw_seq[0]
    gw_rows.append({'No': s, 'SessionId': SESSION_ID, 'BRCD': b,
                    'STTLMAMT': a, 'MSGREF': f'MR{s:010d}', 'PrcFlg': prc,
                    'TxDt': ngay_str(NGAY_DT), 'Status': 'SETTLED'})
    _gw_seq[0] += 1

# GW: count+5 cho non-timeout pairs -> CHENH_LECH = -5 (khong hien thi)
#      count-1 cho timeout pairs    -> CHENH_LECH = +1 (hien thi, SO_TIMEOUT=1)
for (b, a), cnt in _all_mis_ba.items():
    if (b, a) in _timeout_ba_set:
        add_gw(b, a)               # 1 slot, MIS=2 -> CHENH_LECH=1
    else:
        for _ in range(cnt + 5):   # cnt+5 slots -> CHENH_LECH=-5 (an)
            add_gw(b, a)

# GW ACH_Tu_choi -> bi loai trong B3
for i in range(200):
    add_gw(BRANCHES[i % len(BRANCHES)], 999_000, prc='ACH Từ chối')

df_gw_main = pd.DataFrame(gw_rows)
# Sheet 2: lap lai 3k MSGREF tu Sheet 1 (test dedup)
n_dup = min(3000, len(df_gw_main))
df_gw_dup = df_gw_main.head(n_dup).copy()

gw_path = os.path.join(OUT, 'GW_20260615.xlsx')
wb = xlsxwriter.Workbook(gw_path)
for sname, df_s in [('Sheet 1', df_gw_main), ('Sheet 2 dup', df_gw_dup)]:
    ws = wb.add_worksheet(sname)
    for ci, col in enumerate(df_s.columns):
        ws.write(0, ci, col)
    for ri, row in enumerate(df_s.values.tolist(), 1):
        ws.write_row(ri, 0, row)
wb.close()
print(f'  [OK] {"GW_20260615.xlsx":50s}  {len(df_gw_main)+n_dup:>8,} dong '
      f'(sheet1={len(df_gw_main)}, dup={n_dup})')


# ─── 4. MIS_DI ZIPs (2 file, xao tron 60/40) ─────────────────────────────────
_MIS_DI_BASE = dict(NGAY_GIAO_DICH=ngay_str(NGAY_DT), REFHUB='REF', MSGSEQ='1',
                    TXID='0', KENH_THANH_TOAN='ACH', LOAI_LENH_OSB='CT',
                    NH_NHAN='VBAA', MA_GIAO_DICH='0', NOI_DUNG='TEST')

def mis_di_row(trace, tt, branch, amount, session=SESSION_ID,
               ngay_kenh_tra='', se_trace=''):
    return {**_MIS_DI_BASE, 'CHI_NHANH': branch, 'TRANG_THAI_LENH': tt,
            'SO_TIEN': amount, 'TRACE': str(trace), 'SE_TRACE': se_trace,
            'SESSION': session, 'MSGREF': f'MIS{trace:012d}',
            'NGAY_KENH_TRA': ngay_kenh_tra}

mis_di = []

# SCNL matched (trace va amount phai khop voi GL02 DI)
for t in DI_KHOP:
    b, a = _ba(t)
    mis_di.append(mis_di_row(t, 'SCNL', b, a))

# SCNL unmatched -> MIS_DI_THUA SCNL
for t in DI_MIS_ONLY:
    b, a = _ba(t + 100)   # amount khac GL02, khong khop
    mis_di.append(mis_di_row(t, 'SCNL', b, a))

# TPAY non-timeout (amount 600001-601000, GW=1, SCNL=0 -> available=1 -> 0 timeout)
for t in TPAY_NORM:
    b, a = _tpay_ba(t)
    mis_di.append(mis_di_row(t, 'TPAY', b, a,
                              ngay_kenh_tra=ngay_hm(NGAY_DT.replace(hour=14))))

# TPAY timeout: 2 ban ghi / cap (cung branch+amount trong TIMEOUT_BA)
for i, (b, a) in enumerate(TIMEOUT_BA):
    t1 = TIMEOUT_TRACE_BASE + i * 2
    t2 = TIMEOUT_TRACE_BASE + i * 2 + 1
    for t in (t1, t2):
        mis_di.append(mis_di_row(t, 'TPAY', b, a,
                                  ngay_kenh_tra=ngay_hm(NGAY_DT.replace(hour=10))))

# TXRT -> MIS_DI_THUA TXRT
for t in TXRT_POOL:
    b, a = _ba(t)
    mis_di.append(mis_di_row(t, 'TXRT', b, a))

# null-session TPAY IN range -> duoc tinh, co GW slot
in_dt = NGAY_DT.replace(hour=10)
for t in NULL_IN:
    b, a = _null_ba(t)
    mis_di.append(mis_di_row(t, 'TPAY', b, a, session='',
                              ngay_kenh_tra=ngay_hm(in_dt)))

# null-session TPAY OUT of range -> bi loai (ngoai khoang [TPAY_TU, TPAY_DEN))
out_dt = NGAY_TRUOC.replace(hour=10)   # 14/06 10:00 < TPAY_TU=14/06 23:00
for t in NULL_OUT:
    b, a = _ba(t + 400_000)
    mis_di.append(mis_di_row(t, 'TPAY', b, a, session='',
                              ngay_kenh_tra=ngay_hm(out_dt)))

# CALD/ERPO/TPER -> bi loai hoan toan
for i, t in enumerate(EXCL_POOL):
    b, a = _ba(t)
    mis_di.append(mis_di_row(t, ['CALD', 'ERPO', 'TPER'][i % 3], b, a))

# SESSION sai -> bi pre-filter (SESSION != 99001 va != null)
for t in BAD_SES:
    b, a = _ba(t)
    mis_di.append(mis_di_row(t, 'SCNL', b, a, session=SESSION_BAD))

rng.shuffle(mis_di)
cut = int(len(mis_di) * 0.6)
write_zip(os.path.join(OUT, 'MIS_DI_20260615_01.zip'), pd.DataFrame(mis_di[:cut]), 'MIS_DI_01.csv')
write_zip(os.path.join(OUT, 'MIS_DI_20260615_02.zip'), pd.DataFrame(mis_di[cut:]), 'MIS_DI_02.csv')


# ─── 5. MIS_DEN ZIPs (2 file) ─────────────────────────────────────────────────
_MIS_DEN_BASE = dict(NGAY_GIAO_DICH=ngay_str(NGAY_DT), REFHUB='REF', MSGSEQ='1',
                     TXID='0', KENH_THANH_TOAN='ACH', LOAI_LENH_OSB='CT',
                     NH_GUI='VBAA', NOI_DUNG='TEST')

def mis_den_row(trace, tt, branch, amount, session=SESSION_ID):
    return {**_MIS_DEN_BASE, 'CHI_NHANH': branch, 'TRANG_THAI_LENH': tt,
            'SO_TIEN': amount, 'TRACE': str(trace),
            'SESSION': session, 'MSGREF': f'DEN{trace:012d}'}

mis_den = []

# DEN matched (cung seed voi GL02_DEN: _ba(t + 500_000))
for t in DEN_KHOP:
    b, a = _ba(t + 500_000)
    mis_den.append(mis_den_row(t, 'SCNL', b, a))

# DEN unmatched -> MIS_DEN_THUA
for t in DEN_MIS_ONLY:
    b, a = _ba(t + 600_000)
    mis_den.append(mis_den_row(t, 'SCNL', b, a))

# RJCT -> bi loai boi B6
for t in DEN_RJCT:
    b, a = _ba(t)
    mis_den.append(mis_den_row(t, 'RJCT', b, a))

rng.shuffle(mis_den)
cut_den = int(len(mis_den) * 0.6)
write_zip(os.path.join(OUT, 'MIS_DEN_20260615_01.zip'), pd.DataFrame(mis_den[:cut_den]), 'MIS_DEN_01.csv')
write_zip(os.path.join(OUT, 'MIS_DEN_20260615_02.zip'), pd.DataFrame(mis_den[cut_den:]), 'MIS_DEN_02.csv')


# ─── Tong ket ─────────────────────────────────────────────────────────────────
total = (len(gl02_rows) + len(df_gw_main) + n_dup
         + len(mis_di) + len(mis_den))

print('=' * 60)
print(f'Tong so dong tao ra: {total:,}')
print()
print('KET QUA MONG DOI sau main.py:')
print(f'  [B2] NPO_DI={len(DI_KHOP)+len(DI_NPO_ONLY):,}  NPO_DEN={len(DEN_KHOP)+len(DEN_NPO_ONLY):,}  (LOCAC sai bi loc: 5,000)')
print(f'  [B3] GW unique={len(df_gw_main)-200:,}  (dup={n_dup} bi dedup, ACH_Tu_choi=200 bi loai)')
print(f'  [B4] SCNL={len(DI_KHOP)+len(DI_MIS_ONLY):,}  TXRT={len(TXRT_POOL):,}  '
      f'TPAY={len(TPAY_NORM)+len(NULL_IN)+len(TIMEOUT_BA)*2:,}  '
      f'Timeout=20  Pre-filter loai: {len(BAD_SES):,}+{len(NULL_OUT):,}+{len(EXCL_POOL):,}')
print(f'  [B5] DI_KHOP=30,000  NPO_THUA=10,000  MIS_THUA~42,120')
print(f'       MIS_DI_THUA: SCNL={len(DI_MIS_ONLY):,} TPAY~1,320 TXRT={len(TXRT_POOL):,}')
print(f'  [B7] DEN_KHOP=20,000  NPO_THUA=5,000  MIS_THUA=15,000  (RJCT={len(DEN_RJCT):,} bi loai)')
print(f'  TIMEOUT sheet: 20 dong')
print(f'  CAP_CN_TIEN: 20 dong (CHENH_LECH=1)')
print()
print(f'  Ty le khop DI:  30000/(30000+10000) = 75.00%')
print(f'  Ty le khop DEN: 20000/(20000+5000)  = 80.00%')
print()
print('Chay: python main.py --input test_input --output test_output')
