"""
Chuong trinh doi chieu ACH: GL02 (NPO) vs MIS

Cach chay:
    python main.py
    python main.py --input ".\file du lieu" --output ".\output"
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

from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import xlsxwriter


# ─── Cot can giu cho tung loai DataFrame ──────────────────────────
_COLS_NPO = [
    'TRDATE', 'TRBRCD', 'USERID', 'JOURSEQ', 'DYTRSEQ', 'LOCAC', 'CCY',
    'BUSCD', 'UNIT', 'TRCD', 'CUSTOMER', 'TRTP', 'REFERENCE',
    'REMARK', 'DRAMOUNT', 'CRAMOUNT', 'CRTDTM',
]

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


def _clean(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    """Chi giu cac cot co trong df va thuoc danh sach cols."""
    if df is None or len(df) == 0:
        return df
    existing = [c for c in cols if c in df.columns]
    return df[existing].copy()


def _tong_tien(df: pd.DataFrame, col: str) -> int:
    if df is None or len(df) == 0 or col not in df.columns:
        return 0
    return int(pd.to_numeric(df[col], errors='coerce').fillna(0).sum())


# ─── Mau sac ──────────────────────────────────────────────────────
_XANH_LA  = '#C6EFCE'
_DO       = '#FFC7CE'
_CAM      = '#FFEB9C'
_XANH_LAM = '#DDEBF7'


def _viet_sheet(workbook, worksheet, df: pd.DataFrame, header_color: str):
    if df is None or len(df) == 0:
        worksheet.write(0, 0, '(Khong co du lieu)')
        return

    fmt_header = workbook.add_format({'bold': True, 'bg_color': header_color, 'border': 1, 'font_size': 10})
    fmt_cell   = workbook.add_format({'font_size': 10, 'border': 1})

    for col_idx, col_name in enumerate(df.columns):
        worksheet.write(0, col_idx, str(col_name), fmt_header)

    rows = df.fillna('').values.tolist()
    for row_idx, row in enumerate(rows, start=1):
        worksheet.write_row(row_idx, 0, row, fmt_cell)


def _viet_tong_ket(workbook, ws, session_id,
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

    data = [
        ('Ngay doi chieu',           config.NGAY_DOI_CHIEU, ''),
        ('Session',                  session_id,             ''),
        ('',                         '',                     ''),
        ('=== CHIEU DI ===',         '',                     ''),
        ('So giao dich khop (MIS)',  n_di_khop,     s_di_khop),
        ('NPO_DI thua',              n_npo_di_thua, s_npo_di_thua),
        ('MIS_DI thua',              n_mis_di_thua, s_mis_di_thua),
        ('Timeout khong kenh',       n_timeout,     s_timeout),
        ('',                         '',             ''),
        ('=== CHIEU DEN ===',        '',             ''),
        ('So giao dich khop (MIS)',  n_den_khop,    s_den_khop),
        ('NPO_DEN thua',             n_npo_den_thua, s_npo_den_thua),
        ('MIS_DEN thua',             n_mis_den_thua, s_mis_den_thua),
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


# ─── Xuat Excel ───────────────────────────────────────────────────

def xuat_excel(output_path: str, session_id: str,
               df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
               df_timeout, df_mis_den_khop, df_npo_den_thua,
               df_mis_den_thua, df_gw_raw):

    # Lam sach cot truoc khi ghi
    df_gw_clean = df_gw_raw.drop(columns=['KEY_GW'], errors='ignore') if df_gw_raw is not None else None

    sheets = [
        ('TONG_KET',           None,                                   '#FFFFFF'),
        ('MIS_DI_KHOP',        _clean(df_mis_di_khop,  _COLS_MIS_DI), _XANH_LA),
        ('NPO_DI_THUA',        _clean(df_npo_di_thua,  _COLS_NPO),    _DO),
        ('MIS_DI_THUA',        _clean(df_mis_di_thua,  _COLS_MIS_DI), _DO),
        ('TIMEOUT_KHONG_KENH', _clean(df_timeout,       _COLS_MIS_DI), _CAM),
        ('MIS_DEN_KHOP',       _clean(df_mis_den_khop, _COLS_MIS_DEN), _XANH_LA),
        ('NPO_DEN_THUA',       _clean(df_npo_den_thua, _COLS_NPO),    _DO),
        ('MIS_DEN_THUA',       _clean(df_mis_den_thua, _COLS_MIS_DEN), _DO),
        ('RAW_GW',             df_gw_clean,                            _XANH_LAM),
    ]

    workbook = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False, 'constant_memory': True})

    for sheet_name, df, color in sheets:
        ws = workbook.add_worksheet(sheet_name)
        ws.set_tab_color(color)
        if sheet_name == 'TONG_KET':
            _viet_tong_ket(
                workbook, ws, session_id,
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
        else:
            _viet_sheet(workbook, ws, df, color)

    workbook.close()
    print(f'\n[DONE] Da xuat: {output_path}')


# ─── Flow chinh ───────────────────────────────────────────────────

def main():
    from datetime import timedelta
    parser = argparse.ArgumentParser(description='Doi chieu ACH GL02 vs MIS')
    parser.add_argument('--input',  default=config.INPUT_DIR,  help='Thu muc chua file input')
    parser.add_argument('--output', default=config.OUTPUT_DIR, help='Thu muc ket qua output')
    parser.add_argument('--date',   default=None, help='Ngay doi chieu dd/mm/yyyy (mac dinh: lay tu config.py)')
    args = parser.parse_args()

    # Neu truyen --date thi ghi de config runtime (khong sua file config.py)
    if args.date:
        from datetime import datetime
        ngay_dt = datetime.strptime(args.date.strip(), '%d/%m/%Y')
        config.NGAY_DOI_CHIEU = args.date.strip()
        config.NGAY_DT        = ngay_dt
        config.NGAY_TRUOC_DT  = ngay_dt - timedelta(days=1)
        config.TPAY_TU        = config.NGAY_TRUOC_DT.replace(hour=23, minute=0, second=0)
        config.TPAY_DEN       = ngay_dt.replace(hour=23, minute=0, second=0)

    input_dir  = args.input
    output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    print('=' * 60)
    print(f'DOI CHIEU ACH  —  Ngay: {config.NGAY_DOI_CHIEU}')
    print('=' * 60)

    # B1
    session_id = doc_session(input_dir)

    # Tim file (nhanh, tuan tu)
    gl02_files    = _tim_file(input_dir, 'GL02*.zip')
    gw_path       = _tim_gw_xlsx(input_dir)
    mis_di_files  = _tim_file(input_dir, '*_DI_*.zip')
    mis_den_files = _tim_file(input_dir, '*_DEN_*.zip')

    if not gl02_files:
        raise FileNotFoundError('Khong tim thay GL02*.zip')
    if len(mis_di_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DI zip, chi tim thay {len(mis_di_files)}: {mis_di_files}')
    if len(mis_den_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DEN zip, chi tim thay {len(mis_den_files)}: {mis_den_files}')

    # Phase 1: B2 + B3 + B6 song song (doc lap nhau)
    with ThreadPoolExecutor(max_workers=3) as ex:
        f_gl02    = ex.submit(xu_ly_gl02,    gl02_files[0])
        f_gw      = ex.submit(xu_ly_gw,      gw_path, session_id)
        f_mis_den = ex.submit(xu_ly_mis_den, mis_den_files, session_id, config.NGAY_DT)

        npo_di, npo_den          = f_gl02.result()
        dict_gw_count, df_gw_raw = f_gw.result()
        df_mis_den               = f_mis_den.result()

    # B4: can dict_gw_count tu B3 (tuan tu)
    mis_di_final, df_timeout = xu_ly_mis_di(mis_di_files, dict_gw_count, session_id)

    # Phase 2: B5 + B7 song song
    with ThreadPoolExecutor(max_workers=2) as ex:
        f_di  = ex.submit(doi_chieu_di,  npo_di,  mis_di_final)
        f_den = ex.submit(doi_chieu_den, npo_den, df_mis_den)

        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua    = f_di.result()
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua = f_den.result()

    # Xuat Excel
    ngay_str    = config.NGAY_DT.strftime('%Y%m%d')
    output_path = os.path.join(output_dir, f'doi_chieu_{ngay_str}.xlsx')
    xuat_excel(
        output_path, session_id,
        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
        df_timeout,
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua,
        df_gw_raw,
    )


if __name__ == '__main__':
    main()
