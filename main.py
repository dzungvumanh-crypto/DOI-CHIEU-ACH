"""
Chuong trinh doi chieu ACH: GL02 (NPO) vs MIS

Cach chay:
    python main.py
    python main.py --input ".\\file du lieu" --output ".\\output"
"""
import sys
import os
import glob
import argparse

# Force UTF-8 stdout de tranh crash khi in ky tu dac biet tren Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from modules.b1_doc_session   import doc_session
from modules.b2_xu_ly_gl02    import xu_ly_gl02
from modules.b3_xu_ly_gw      import xu_ly_gw
from modules.b4_xu_ly_mis_di  import xu_ly_mis_di
from modules.b5_doi_chieu_di  import doi_chieu_di
from modules.b6_xu_ly_mis_den import xu_ly_mis_den
from modules.b7_doi_chieu_den import doi_chieu_den
from modules.b8_phan_tich    import phan_tich

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

import pandas as pd
import xlsxwriter
from tqdm import tqdm


# ─── Cot can giu cho tung loai DataFrame ──────────────────────────
# C3: lay tu config.py — nguon duy nhat, dong bo voi b2_xu_ly_gl02.py
_COLS_NPO = config.COLS_NPO

_COLS_MIS_DI = [
    'NGAY_GIAO_DICH', 'CHI_NHANH', 'CN tiền Hub', 'REFHUB', 'MSGREF',
    'MSGSEQ', 'TXID', 'KENH_THANH_TOAN', 'TRANG_THAI_LENH', 'SO_TIEN',
    'TRACE', 'SE_TRACE', 'SESSION', 'LOAI_LENH_OSB', 'NH_NHAN',
    'MA_GIAO_DICH', 'NOI_DUNG', 'NGAY_KENH_TRA',
]

_COLS_MIS_DEN = [
    'NGAY_GIAO_DICH', 'CHI_NHANH', 'REFHUB', 'MSGREF', 'MSGSEQ', 'TXID',
    'KENH_THANH_TOAN', 'TRANG_THAI_LENH', 'SO_TIEN', 'TRACE',
    'SESSION', 'LOAI_LENH_OSB', 'NH_GUI', 'NOI_DUNG',
]

# Nguong dong: sheet lon hon se xuat ra CSV thay vi ghi vao Excel
CSV_THRESHOLD = 50_000


def _clean(df: pd.DataFrame, cols: list, label: str = '') -> pd.DataFrame:
    """Chi giu cac cot co trong df va thuoc danh sach cols. Log warning neu thieu cot."""
    if df is None or len(df) == 0:
        return df
    existing = [c for c in cols if c in df.columns]
    missing  = [c for c in cols if c not in df.columns]
    if missing:
        print(f'[WARN] _clean({label}): thieu cot {missing}')
    return df[existing].copy()


def _tong_tien(df: pd.DataFrame, col: str) -> int:
    if df is None or len(df) == 0 or col not in df.columns:
        return 0
    return int(pd.to_numeric(df[col], errors='coerce').fillna(0).sum())


def _tao_cap_cn_tien(mis_di_final, df_timeout, dict_gw_count):
    """
    So sanh count CN+TIEN giua TAT CA MIS (SCNL+TXRT+TPAY) va GW.
    Logic giong manual pivot: COUNT_MIS - COUNT_GW, chi hien cap CHENH_LECH > 0.
    Quan trong: phai tinh toan TAT CA giao dich MIS (khong chi TPAY) de bat dung
    truong hop SCNL da dung het slot GW khien TPAY thanh timeout.
    Vi du: CN=3617 SO_TIEN=35000: MIS=2 (1 SCNL + 1 TPAY timeout), GW=1 -> CHENH_LECH=1.
    """
    cn_col = 'CN tiền Hub'
    frames = []
    if mis_di_final is not None and len(mis_di_final) > 0:
        frames.append(mis_di_final)
    if df_timeout is not None and len(df_timeout) > 0:
        frames.append(df_timeout)

    if not frames:
        return pd.DataFrame(columns=['CHI_NHANH', 'SO_TIEN', 'COUNT_MIS', 'COUNT_GW', 'CHENH_LECH', 'SO_TIMEOUT'])

    df_all = pd.concat(frames, ignore_index=True)

    if cn_col not in df_all.columns:
        return pd.DataFrame(columns=['CHI_NHANH', 'SO_TIEN', 'COUNT_MIS', 'COUNT_GW', 'CHENH_LECH', 'SO_TIMEOUT'])

    cnt = df_all.groupby(cn_col, sort=False).size().rename('COUNT_MIS').reset_index()
    cnt['COUNT_GW']   = cnt[cn_col].map(dict_gw_count).fillna(0).astype(int)
    cnt['CHENH_LECH'] = cnt['COUNT_MIS'] - cnt['COUNT_GW']

    if df_timeout is not None and len(df_timeout) > 0 and cn_col in df_timeout.columns:
        to_cnt = df_timeout.groupby(cn_col, sort=False).size().rename('SO_TIMEOUT')
        cnt = cnt.merge(to_cnt, on=cn_col, how='left')
        cnt['SO_TIMEOUT'] = cnt['SO_TIMEOUT'].fillna(0).astype(int)
    else:
        cnt['SO_TIMEOUT'] = 0

    ref = df_all.drop_duplicates(subset=[cn_col])
    cnt['CHI_NHANH'] = cnt[cn_col].map(ref.set_index(cn_col)['CHI_NHANH'].to_dict())
    cnt['SO_TIEN']   = cnt[cn_col].map(ref.set_index(cn_col)['SO_TIEN'].to_dict())

    result = cnt[cnt['CHENH_LECH'] > 0][
        ['CHI_NHANH', 'SO_TIEN', 'COUNT_MIS', 'COUNT_GW', 'CHENH_LECH', 'SO_TIMEOUT']
    ].copy()
    return result.sort_values('CHENH_LECH', ascending=False).reset_index(drop=True)


# ─── Mau sac ──────────────────────────────────────────────────────
_XANH_LA  = '#C6EFCE'
_DO       = '#FFC7CE'
_CAM      = '#FFEB9C'
_XANH_LAM = '#DDEBF7'
_XANH_NHAT = '#E2EFDA'


def _viet_phan_tich(workbook, worksheet, df: pd.DataFrame):
    """Ghi sheet PHAN_TICH voi mau sac theo _type: header/sub_header/canh_bao/data."""
    if df is None or len(df) == 0:
        worksheet.write(0, 0, '(Khong co du lieu)')
        return

    fmt_col_hdr  = workbook.add_format({'bold': True, 'font_size': 10, 'bg_color': '#DDEBF7', 'border': 1})
    fmt_header   = workbook.add_format({'bold': True, 'font_size': 10, 'bg_color': '#BDD7EE', 'border': 1})
    fmt_sub      = workbook.add_format({'bold': True, 'font_size': 10, 'bg_color': '#E2EFDA', 'border': 1})
    fmt_canh_bao = workbook.add_format({'bold': True, 'font_size': 10, 'bg_color': '#FFEB9C', 'border': 1})
    fmt_data     = workbook.add_format({'font_size': 10, 'border': 1})

    cols   = ['Chi tieu', 'Gia tri', 'Ghi chu']
    widths = [60, 30, 70]
    for ci, (col, w) in enumerate(zip(cols, widths)):
        worksheet.write(0, ci, col, fmt_col_hdr)
        worksheet.set_column(ci, ci, w)

    for row_idx, row in df.iterrows():
        typ = str(row.get('_type', ''))
        if typ == 'header':
            fmt = fmt_header
        elif typ == 'sub_header':
            fmt = fmt_sub
        elif typ == 'canh_bao':
            fmt = fmt_canh_bao
        else:
            fmt = fmt_data
        for ci, col in enumerate(cols):
            val = row[col] if row[col] else ''
            worksheet.write(row_idx + 1, ci, str(val), fmt)


def _viet_sheet(workbook, worksheet, df: pd.DataFrame, header_color: str):
    if df is None or len(df) == 0:
        worksheet.write(0, 0, '(Khong co du lieu)')
        return

    fmt_header = workbook.add_format({'bold': True, 'bg_color': header_color, 'border': 1, 'font_size': 10})
    fmt_cell   = workbook.add_format({'font_size': 10, 'border': 1})

    # Chuyen cot datetime -> string truoc khi ghi de tranh hien thi so serial Excel
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')

    for col_idx, col_name in enumerate(df.columns):
        # Auto column width: dua tren do dai ten cot, toi thieu 8, toi da 30
        width = min(max(len(str(col_name)), 8) + 2, 30)
        worksheet.set_column(col_idx, col_idx, width)
        worksheet.write(0, col_idx, str(col_name), fmt_header)

    rows = df.fillna('').values.tolist()
    for row_idx, row in enumerate(rows, start=1):
        worksheet.write_row(row_idx, 0, row, fmt_cell)



def _viet_tong_ket(workbook, ws, session_id, ngay_doi_chieu_str,
                   n_di_khop,      s_di_khop,
                   n_npo_di_thua,  s_npo_di_thua,
                   n_mis_di_thua,  s_mis_di_thua,
                   n_timeout,      s_timeout,
                   n_den_khop,     s_den_khop,
                   n_npo_den_thua, s_npo_den_thua,
                   n_mis_den_thua, s_mis_den_thua):
    fmt_label  = workbook.add_format({'bold': True, 'font_size': 10})
    fmt_header = workbook.add_format({'bold': True, 'font_size': 10,
                                      'bg_color': '#DDEBF7', 'border': 1})
    fmt_num    = workbook.add_format({'font_size': 10, 'num_format': '#,##0'})
    fmt_val    = workbook.add_format({'font_size': 10})

    ws.write(0, 0, 'Chi tieu',           fmt_header)
    ws.write(0, 1, 'So giao dich',       fmt_header)
    ws.write(0, 2, 'Tong so tien (VND)', fmt_header)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 22)

    n_npo_di  = n_di_khop  + n_npo_di_thua
    n_npo_den = n_den_khop + n_npo_den_thua
    ty_le_di  = round(n_di_khop  / n_npo_di  * 100, 2) if n_npo_di  > 0 else 0.0
    ty_le_den = round(n_den_khop / n_npo_den * 100, 2) if n_npo_den > 0 else 0.0

    data = [
        ('Ngay doi chieu',           ngay_doi_chieu_str,    ''),
        ('Session',                  session_id,             ''),
        ('',                         '',                     ''),
        ('=== CHIEU DI ===',         '',                     ''),
        ('So giao dich khop (MIS)',  n_di_khop,     s_di_khop),
        ('NPO_DI thua',              n_npo_di_thua, s_npo_di_thua),
        ('MIS_DI thua',              n_mis_di_thua, s_mis_di_thua),
        ('Timeout khong kenh',       n_timeout,     s_timeout),
        ('Ty le khop DI (%)',        ty_le_di,      f'{n_di_khop:,} / {n_npo_di:,}'),
        ('',                         '',             ''),
        ('=== CHIEU DEN ===',        '',             ''),
        ('So giao dich khop (MIS)',  n_den_khop,    s_den_khop),
        ('NPO_DEN thua',             n_npo_den_thua, s_npo_den_thua),
        ('MIS_DEN thua',             n_mis_den_thua, s_mis_den_thua),
        ('Ty le khop DEN (%)',       ty_le_den,     f'{n_den_khop:,} / {n_npo_den:,}'),
    ]

    for row_idx, (label, val, tien) in enumerate(data, start=1):
        ws.write_string(row_idx, 0, label, fmt_label)
        if isinstance(val, int):
            ws.write(row_idx, 1, val, fmt_num)
        else:
            ws.write(row_idx, 1, val, fmt_val)
        if isinstance(tien, int) and tien > 0:
            ws.write(row_idx, 2, tien, fmt_num)
        elif tien != '':
            ws.write(row_idx, 2, tien, fmt_val)


# ─── Tim file ─────────────────────────────────────────────────────

def _tim_ngay_tu_pdf(input_dir: str) -> str:
    """
    Doc ten file PDF de suy ra ngay doi chieu.
    ACH_20260612_VBAAVNVN_NRT_15882_... -> ngay T+1 = 20260612 -> T = 11/06/2026.
    """
    import re as _re
    from datetime import timedelta as _td
    for root, _, files in os.walk(os.path.abspath(input_dir)):
        for f in files:
            if f.endswith('.pdf'):
                m = _re.search(r'_(\d{8})_', f)
                if m:
                    d = datetime.strptime(m.group(1), '%Y%m%d') - _td(days=1)
                    return d.strftime('%d/%m/%Y')
    return None


def _tim_file(input_dir: str, pattern: str) -> list:
    """Tim file theo pattern, tim de quy qua tat ca subfolder (dung abs path)."""
    abs_dir = os.path.abspath(input_dir)
    return sorted(glob.glob(os.path.join(abs_dir, '**', pattern), recursive=True))


def _tim_gw_xlsx(input_dir: str) -> str:
    """Tim file .xlsx co chua cot BRCD va SessionId (dau hieu cua file GW)."""
    abs_dir = os.path.abspath(input_dir)
    for f in glob.glob(os.path.join(abs_dir, '**', '*.xlsx'), recursive=True):
        try:
            xl = pd.ExcelFile(f, engine='calamine')
            for sheet in xl.sheet_names:
                df_peek = pd.read_excel(xl, sheet_name=sheet, header=None,
                                        nrows=8, dtype=str, engine='calamine')
                flat = set(str(v).strip() for v in df_peek.values.flatten() if str(v) != 'nan')
                if 'BRCD' in flat and 'SessionId' in flat:
                    return f
        except Exception:
            continue
    raise FileNotFoundError('Khong tim thay file GW .xlsx trong: ' + abs_dir)


# ─── Xuat Excel / CSV ─────────────────────────────────────────────

def xuat_excel(output_path: str, session_id: str,
               df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
               df_timeout, df_mis_den_khop, df_npo_den_thua,
               df_mis_den_thua, df_gw_raw,
               df_cap_cn_tien=None, df_phan_tich=None,
               log_callback=None):

    output_dir  = os.path.dirname(os.path.abspath(output_path))
    ngay_str    = os.path.basename(output_path).replace('doi_chieu_', '').replace('.xlsx', '')
    # Chuyen YYYYMMDD -> dd/mm/yyyy de hien thi trong TONG_KET
    ngay_display = datetime.strptime(ngay_str, '%Y%m%d').strftime('%d/%m/%Y')
    df_gw_clean = df_gw_raw.drop(columns=['KEY_GW'], errors='ignore') if df_gw_raw is not None else None

    sheets = [
        ('TONG_KET',           None,                                                     '#FFFFFF'),
        ('PHAN_TICH',          df_phan_tich,                                             _XANH_NHAT),
        ('MIS_DI_KHOP',        _clean(df_mis_di_khop,  _COLS_MIS_DI, 'MIS_DI_KHOP'),   _XANH_LA),
        ('NPO_DI_THUA',        _clean(df_npo_di_thua,  _COLS_NPO,    'NPO_DI_THUA'),   _DO),
        ('MIS_DI_THUA',        _clean(df_mis_di_thua,  _COLS_MIS_DI, 'MIS_DI_THUA'),   _DO),
        ('TIMEOUT_KHONG_KENH', _clean(df_timeout,       _COLS_MIS_DI, 'TIMEOUT'),        _CAM),
        ('CAP_CN_TIEN',        df_cap_cn_tien,                                           _CAM),
        ('MIS_DEN_KHOP',       _clean(df_mis_den_khop, _COLS_MIS_DEN,'MIS_DEN_KHOP'),  _XANH_LA),
        ('NPO_DEN_THUA',       _clean(df_npo_den_thua, _COLS_NPO,    'NPO_DEN_THUA'),  _DO),
        ('MIS_DEN_THUA',       _clean(df_mis_den_thua, _COLS_MIS_DEN,'MIS_DEN_THUA'),  _DO),
        ('RAW_GW',             df_gw_clean,                                              _XANH_LAM),
    ]

    workbook = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False, 'constant_memory': True})
    csv_writes = []   # list of (sheet_name, csv_path, future)
    _log = log_callback or print
    total_sheets = len(sheets)

    with ThreadPoolExecutor(max_workers=3) as csv_pool:
        for i, (sheet_name, df, color) in enumerate(tqdm(sheets, desc='Ghi Excel', unit='sheet'), 1):
            _log(f'[EXCEL] ({i}/{total_sheets}) Ghi sheet: {sheet_name}...')
            # Sheet lon → ghi CSV song song, ghi note vao Excel ngay khong cho CSV xong
            if df is not None and len(df) > CSV_THRESHOLD:
                csv_path = os.path.join(output_dir, f'{sheet_name}_{ngay_str}.csv')
                fut = csv_pool.submit(df.to_csv, csv_path, index=False, encoding='utf-8-sig')
                csv_writes.append((sheet_name, csv_path, fut))
                ws = workbook.add_worksheet(sheet_name)
                ws.set_tab_color(color)
                ws.write(0, 0, f'[Du lieu lon - xem file: {os.path.basename(csv_path)}]')
                ws.write(1, 0, f'Tong so dong: {len(df):,}')
                ws.write(2, 0, 'LUU Y: Mo file CSV qua Excel > Data > Tu Van ban/CSV (khong double-click truc tiep).')
                ws.write(3, 0, 'Double-click co the mat so 0 dau o cot TRACE, MSGSEQ va sai dinh dang so tien.')
                _log(f'[CSV] {sheet_name}: {len(df):,} dong → dang ghi nen...')
                continue

            ws = workbook.add_worksheet(sheet_name)
            ws.set_tab_color(color)
            if sheet_name == 'TONG_KET':
                _viet_tong_ket(
                    workbook, ws, session_id, ngay_display,
                    len(df_mis_di_khop)  if df_mis_di_khop  is not None else 0,
                    _tong_tien(df_mis_di_khop,   'SO_TIEN'),
                    len(df_npo_di_thua)  if df_npo_di_thua  is not None else 0,
                    _tong_tien(df_npo_di_thua,   'CRAMOUNT'),
                    len(df_mis_di_thua)  if df_mis_di_thua  is not None else 0,
                    _tong_tien(df_mis_di_thua,   'SO_TIEN'),
                    len(df_timeout)      if df_timeout      is not None else 0,
                    _tong_tien(df_timeout,       'SO_TIEN'),
                    len(df_mis_den_khop) if df_mis_den_khop is not None else 0,
                    _tong_tien(df_mis_den_khop,  'SO_TIEN'),
                    len(df_npo_den_thua) if df_npo_den_thua is not None else 0,
                    _tong_tien(df_npo_den_thua,  'DRAMOUNT'),
                    len(df_mis_den_thua) if df_mis_den_thua is not None else 0,
                    _tong_tien(df_mis_den_thua,  'SO_TIEN'),
                )
            elif sheet_name == 'PHAN_TICH':
                _viet_phan_tich(workbook, ws, df)
            else:
                _viet_sheet(workbook, ws, df, color)

        workbook.close()
        _log(f'[DONE] Excel: {output_path}')
    # csv_pool.__exit__ da doi tat ca CSV xong moi thoat
    for name, path, fut in csv_writes:
        fut.result()  # raise neu co loi
        _log(f'       CSV  : {path}  ({name})')


# ─── main_from_dir — dung cho Web UI (thread-safe) ────────────────

def _cancelled(ev) -> bool:
    """Tra ve True neu cancel_event duoc set (nguoi dung bam Dung)."""
    return ev is not None and ev.is_set()


def main_from_dir(input_dir: str, output_dir: str,
                  ngay: str = None, log_callback=None,
                  cancel_event=None) -> str:
    """
    Phien ban cua main() dung cho Web UI.
    - input_dir: thu muc da co file upload
    - output_dir: noi luu file ket qua
    - ngay: 'dd/mm/yyyy' hoac None (lay tu config)
    - log_callback: ham(msg: str) de emit log real-time
    Tra ve: duong dan file output .xlsx
    Thread-safe: ngay doi chieu tinh local, KHONG sua config module-level.
    """
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    # Tinh ngay local thay vi mutate config global
    if ngay:
        ngay_dt      = datetime.strptime(ngay.strip(), '%d/%m/%Y')
        ngay_str_cfg = ngay.strip()
    else:
        # Tu dong phat hien tu ten file PDF (ACH_YYYYMMDD_... -> ngay T-1)
        auto_ngay = _tim_ngay_tu_pdf(input_dir)
        if auto_ngay:
            ngay_dt      = datetime.strptime(auto_ngay, '%d/%m/%Y')
            ngay_str_cfg = auto_ngay
            log(f'[AUTO] Phat hien ngay doi chieu tu PDF: {ngay_str_cfg}')
        else:
            ngay_dt      = config.NGAY_DT
            ngay_str_cfg = config.NGAY_DOI_CHIEU

    # Tinh tpay de truyen tuong minh xuong b4 (thread-safe)
    tpay_tu  = (ngay_dt - timedelta(days=1)).replace(hour=23, minute=0, second=0)
    tpay_den = ngay_dt.replace(hour=23, minute=0, second=0)

    os.makedirs(output_dir, exist_ok=True)

    # Kiem tra file output co bi khoa (Excel dang mo) TRUOC khi xu ly 7 phut
    ngay_check = ngay_dt.strftime('%Y%m%d')
    output_xlsx = os.path.join(output_dir, f'doi_chieu_{ngay_check}.xlsx')
    if os.path.exists(output_xlsx):
        try:
            with open(output_xlsx, 'a'):
                pass
        except PermissionError:
            raise PermissionError(
                f'\n[LOI] File dang mo trong Excel. Vui long DONG FILE roi chay lai:\n'
                f'       {os.path.abspath(output_xlsx)}'
            )

    log(f'Ngay doi chieu: {ngay_str_cfg}')

    session_id    = doc_session(input_dir, log_callback)
    gl02_files    = _tim_file(input_dir, 'GL02*.zip')
    gw_path       = _tim_gw_xlsx(input_dir)
    mis_di_files  = _tim_file(input_dir, '*_DI_*.zip')
    mis_den_files = _tim_file(input_dir, '*_DEN_*.zip')

    if not gl02_files:
        raise FileNotFoundError('Khong tim thay GL02*.zip')
    if len(mis_di_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DI zip, chi tim thay {len(mis_di_files)}')
    if len(mis_den_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DEN zip, chi tim thay {len(mis_den_files)}')

    log(f'Tim thay: GL02={len(gl02_files)}, DI={len(mis_di_files)}, DEN={len(mis_den_files)}')

    # [CANCEL POINT 1] Truoc khi bat dau xu ly nang — kiem tra sau khi doc file list
    if _cancelled(cancel_event):
        log('[CANCELLED] Nguoi dung da dung. Khong xu ly.')
        return None

    # Phase 1: B2 + B3 + B6 song song; B4 khoi dong ngay khi B3 xong (khong cho B2/B6)
    with ThreadPoolExecutor(max_workers=3) as ex:
        f_gl02    = ex.submit(xu_ly_gl02,    gl02_files[0], log_callback)
        f_gw      = ex.submit(xu_ly_gw,      gw_path, session_id, log_callback)
        f_mis_den = ex.submit(xu_ly_mis_den, mis_den_files, session_id, ngay_dt, log_callback)

        # B4 can dict_gw_count tu B3; bat dau ngay khi B3 xong, chay song song B2/B6 con lai
        dict_gw_count, df_gw_raw = f_gw.result()

        # [CANCEL POINT 2] Sau khi B3 xong — truoc khi nop B4 (buoc nang nhat)
        if _cancelled(cancel_event):
            log('[CANCELLED] Nguoi dung da dung sau B3. Cho B2/B6 hoan thanh...')
            f_gl02.result()
            f_mis_den.result()
            return None

        f_mis_di = ex.submit(
            xu_ly_mis_di,
            mis_di_files, dict_gw_count, session_id,
            tpay_tu, tpay_den, log_callback,
        )

        npo_di, npo_den          = f_gl02.result()
        df_mis_den               = f_mis_den.result()
        mis_di_final, df_timeout = f_mis_di.result()

    # [CANCEL POINT 3] Sau Phase 1 — truoc Phase 2 va ghi Excel
    if _cancelled(cancel_event):
        log('[CANCELLED] Nguoi dung da dung sau Phase 1.')
        return None

    # CAP_CN_TIEN: so sanh count MIS_TPAY vs GW per CN+TIEN (thay pivot thu cong)
    df_cap_cn_tien = _tao_cap_cn_tien(mis_di_final, df_timeout, dict_gw_count)

    # Phase 2: B5 + B7 song song
    with ThreadPoolExecutor(max_workers=2) as ex:
        f_di  = ex.submit(doi_chieu_di,  npo_di,  mis_di_final, log_callback)
        f_den = ex.submit(doi_chieu_den, npo_den, df_mis_den,   log_callback)

        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua    = f_di.result()
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua = f_den.result()

    # [CANCEL POINT 4] Truoc ghi Excel (buoc cuoi)
    if _cancelled(cancel_event):
        log('[CANCELLED] Nguoi dung da dung truoc khi ghi Excel.')
        return None

    # PHAN_TICH: dashboard chat luong doi chieu (thay phan tich thu cong)
    df_phan_tich = phan_tich(
        df_npo_di_thua, df_mis_di_thua,
        df_npo_den_thua, df_mis_den_thua,
        len(df_mis_di_khop), len(df_mis_den_khop), df_timeout,
    )

    output_path = os.path.join(output_dir, f'doi_chieu_{ngay_dt.strftime("%Y%m%d")}.xlsx')
    xuat_excel(
        output_path, session_id,
        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
        df_timeout,
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua,
        df_gw_raw,
        df_cap_cn_tien=df_cap_cn_tien,
        df_phan_tich=df_phan_tich,
        log_callback=log_callback,
    )
    log(f'Hoan thanh: {output_path}')
    return output_path


# ─── Flow chinh (CLI) ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Doi chieu ACH GL02 vs MIS')
    parser.add_argument('--input',  default=config.INPUT_DIR,  help='Thu muc chua file input')
    parser.add_argument('--output', default=config.OUTPUT_DIR, help='Thu muc ket qua output')
    parser.add_argument('--date',   default=None, help='Ngay doi chieu dd/mm/yyyy (mac dinh: lay tu config.py)')
    args = parser.parse_args()

    print('=' * 60)
    print(f'DOI CHIEU ACH  —  Ngay: {args.date or config.NGAY_DOI_CHIEU}')
    print('=' * 60)

    main_from_dir(
        input_dir=args.input,
        output_dir=args.output,
        ngay=args.date,
    )


if __name__ == '__main__':
    main()
