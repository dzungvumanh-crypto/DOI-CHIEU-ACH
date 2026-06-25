import io
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List
from datetime import datetime
import numpy as np
import pyzipper
import pandas as pd
from config import ZIP_PASSWORD

_TRANG_THAI_LOAI_TRU = {'CALD', 'ERPO', 'TPER'}

_COLS = [
    'NGAY_GIAO_DICH', 'CHI_NHANH', 'REFHUB', 'MSGREF', 'MSGSEQ', 'TXID',
    'KENH_THANH_TOAN', 'TRANG_THAI_LENH', 'SO_TIEN', 'TRACE',
    'SE_TRACE', 'SESSION', 'LOAI_LENH_OSB', 'NH_NHAN',
    'MA_GIAO_DICH', 'NOI_DUNG', 'NGAY_KENH_TRA',
]

_NULL_SESSION = frozenset({'', 'nan', 'None', 'NaN'})


def _detect_encoding(z: pyzipper.AESZipFile, name: str) -> str:
    """Phat hien encoding bang cach peek 512 byte dau — tranh re-read toan bo file."""
    with z.open(name) as f:
        raw = f.read(512)
    if raw[:3] == b'\xef\xbb\xbf':
        return 'utf-8-sig'
    try:
        raw.decode('utf-8')
        return 'utf-8'
    except UnicodeDecodeError:
        return 'cp1252'


def _doc_zip(zip_path: str, session_filter: str = None) -> pd.DataFrame:
    """
    Doc ZIP chua CSV MIS_DI, su dung streaming (z.open) thay vi z.read() de giam RAM.
    session_filter: neu truyen, chi giu dong co SESSION == session_filter hoac SESSION null —
    giam ~60-70% du lieu truoc khi xu ly tiep theo.
    """
    frames = []
    with pyzipper.AESZipFile(zip_path, 'r') as z:
        z.setpassword(ZIP_PASSWORD)
        for name in z.namelist():
            if not name.lower().endswith('.csv'):
                continue
            enc = _detect_encoding(z, name)
            for errors in ('strict', 'replace'):
                try:
                    with z.open(name) as raw_f:
                        wrapped = io.TextIOWrapper(raw_f, encoding=enc, errors=errors)
                        if session_filter:
                            sid = str(session_filter)
                            keep_sessions = frozenset({sid} | _NULL_SESSION)
                            chunk_list = []
                            for chunk in pd.read_csv(
                                wrapped, dtype=str,
                                usecols=lambda c: c in _COLS,
                                chunksize=200_000, low_memory=False,
                            ):
                                if 'SESSION' in chunk.columns:
                                    sess = (chunk['SESSION'].fillna('').astype(str)
                                            .str.strip().str.lstrip("'"))
                                    mask = sess.isin(keep_sessions)
                                    chunk = chunk[mask]
                                if not chunk.empty:
                                    chunk_list.append(chunk)
                            if chunk_list:
                                frames.append(pd.concat(chunk_list, ignore_index=True))
                        else:
                            df = pd.read_csv(
                                wrapped, dtype=str,
                                usecols=lambda c: c in _COLS,
                                low_memory=False,
                            )
                            frames.append(df)
                    break
                except UnicodeDecodeError:
                    if errors == 'strict':
                        print(f'[B4][WARN] Encoding detect mismatch trong {name}, thu lai errors=replace')
                        continue
                    raise
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_COLS)


def _tao_so_trace(df: pd.DataFrame) -> pd.Series:
    """SE_TRACE neu co gia tri, nguoc lai dung TRACE. Bo dau nháy don va leading zero."""
    se = df['SE_TRACE'].fillna('').astype(str).str.strip().str.lstrip("'0")
    tr = df['TRACE'].fillna('').astype(str).str.strip().str.lstrip("'0")
    return se.where(se.ne(''), tr)


def _get_timeout_indices(df_tpay: pd.DataFrame, df_non_tpay: pd.DataFrame,
                         dict_gw_count: Dict[str, int]) -> pd.Index:
    """
    Tra ve Index cua cac dong TPAY thua so voi GW (theo doc 3.2).
    df_non_tpay: SCNL + TXRT — tat ca lenh da su dung slot GW.
    surplus = count_all_mis - count_gw; n_timeout = min(surplus, count_tpay).
    Vectorized: khong dung Python for-loop, dung cumcount + map.
    """
    if df_tpay.empty:
        return pd.Index([], dtype='int64')

    key_col  = 'CN tiền Hub'
    cnt_tpay = df_tpay[key_col].value_counts()
    cnt_non  = (df_non_tpay[key_col].value_counts()
                if not df_non_tpay.empty else pd.Series(dtype='int64'))

    keys      = cnt_tpay.index
    c_gw      = pd.Series({k: dict_gw_count.get(str(k), 0) for k in keys}, dtype='int64')
    c_non     = cnt_non.reindex(keys, fill_value=0)
    available = (c_gw - c_non).clip(lower=0)
    n_thua    = (cnt_tpay - available).clip(lower=0)

    cc_rev    = df_tpay.groupby(key_col, sort=False).cumcount(ascending=False)
    threshold = df_tpay[key_col].map(n_thua.to_dict()).fillna(0)
    return df_tpay.index[cc_rev < threshold]


def _doc_mis_di_raw(zip_paths: List[str], session_id: str, log_callback=None) -> pd.DataFrame:
    """Doc 2 ZIP MIS_DI song song, tra ve DataFrame thu (chua xu ly). Dung cho P7 parallel I/O."""
    _log = log_callback or print
    _log('[B4] Doc MIS_DI tu 2 ZIP...')
    sid = str(session_id)
    with ThreadPoolExecutor(max_workers=2) as ex:
        futures = [ex.submit(_doc_zip, p, sid) for p in zip_paths]
        frames  = [f.result() for f in futures]
    return pd.concat(frames, ignore_index=True)


def _process_mis_di(df: pd.DataFrame, dict_gw_count: Dict[str, int], session_id: str,
                    df_gw: pd.DataFrame = None,
                    tpay_tu: datetime = None, tpay_den: datetime = None,
                    log_callback=None):
    """Xu ly DataFrame MIS_DI da doc truoc (tu _doc_mis_di_raw hoac xu_ly_mis_di)."""
    import config
    _tpay_tu  = tpay_tu  if tpay_tu  is not None else config.TPAY_TU
    _tpay_den = tpay_den if tpay_den is not None else config.TPAY_DEN

    _log = log_callback or print
    sid = str(session_id)

    # Bo trang thai loai tru
    df = df[~df['TRANG_THAI_LENH'].isin(_TRANG_THAI_LOAI_TRU)].copy()

    # Parse SO_TIEN
    df['SO_TIEN'] = pd.to_numeric(df['SO_TIEN'], errors='coerce').fillna(0).astype('int64')

    # Tao SO_TRACE
    df['SO_TRACE'] = _tao_so_trace(df)

    # Parse NGAY_KENH_TRA
    df['NGAY_KENH_TRA'] = pd.to_datetime(
        df['NGAY_KENH_TRA'].str.strip(), format='%d/%m/%Y %H:%M:%S', errors='coerce'
    )

    # Chuan hoa SESSION
    df['SESSION'] = df['SESSION'].fillna('').astype(object).astype(str).str.strip().str.lstrip("'")
    df['SESSION_NULL'] = df['SESSION'].isin(['', 'nan', 'None', 'NaN'])

    # SCNL: SESSION = session_id
    mask_scnl = df['TRANG_THAI_LENH'] == 'SCNL'
    df_scnl = df[mask_scnl & (df['SESSION'] == sid)].copy()

    # TXRT: chi lay trong session hien tai (tranh lay TXRT tu session cu)
    df_txrt = df[(df['TRANG_THAI_LENH'] == 'TXRT') & (df['SESSION'] == sid)].copy()

    # TPAY: SESSION = session_id HOAC (SESSION null VA NGAY_KENH_TRA trong khoang)
    mask_tpay = df['TRANG_THAI_LENH'] == 'TPAY'
    mask_session_ok = df['SESSION'] == sid
    mask_null_ok = (
        df['SESSION_NULL']
        & df['NGAY_KENH_TRA'].notna()
        & (df['NGAY_KENH_TRA'] >= _tpay_tu)
        & (df['NGAY_KENH_TRA'] < _tpay_den)
    )
    df_tpay = df[mask_tpay & (mask_session_ok | mask_null_ok)].copy()

    # Gop lai — giu index goc
    df_mis_di = pd.concat([df_scnl, df_txrt, df_tpay])

    # Tao KEY_HUB va CN tien Hub — CHI_NHANH strip 1 lan
    cn_clean = df_mis_di['CHI_NHANH'].astype(str).str.strip()
    df_mis_di['KEY_HUB'] = cn_clean + df_mis_di['SO_TRACE'] + df_mis_di['SO_TIEN'].astype(str)
    cn_tien = cn_clean + df_mis_di['SO_TIEN'].astype(str)
    loc = df_mis_di.columns.get_loc('CHI_NHANH') + 1
    df_mis_di.insert(loc, 'CN tiền Hub', cn_tien)

    # Tinh timeout — SCNL va TXRT deu chiem slot GW (theo doc 3.2: dem ca SCNL+TXRT+TPAY vs GW)
    df_non_tpay_in_mis = df_mis_di[df_mis_di['TRANG_THAI_LENH'].isin(['SCNL', 'TXRT'])]
    df_tpay_in_mis = df_mis_di[df_mis_di['TRANG_THAI_LENH'] == 'TPAY']
    timeout_idx = _get_timeout_indices(df_tpay_in_mis, df_non_tpay_in_mis, dict_gw_count)

    df_timeout_candidates = df_mis_di.loc[timeout_idx].copy()
    df_mis_di_final       = df_mis_di[~df_mis_di.index.isin(timeout_idx)].copy()

    # MSGREF check: TPAY co MSGREF trong GW → giu trong timeout va danh dau CO_TRONG_GW=True
    # (giao dich nay co the da di kenh thanh cong, can check thu cong)
    _QUOTE = "'"
    df_timeout = df_timeout_candidates.copy()
    n_in_gw = 0
    if df_gw is not None and 'MSGREF' in df_gw.columns and len(df_timeout_candidates) > 0:
        gw_msgref_set = set(
            df_gw['MSGREF'].fillna('').astype(object).astype(str)
            .str.strip().str.lstrip(_QUOTE)
        )
        tpay_msgref = (
            df_timeout['MSGREF'].fillna('').astype(object).astype(str)
            .str.strip().str.lstrip(_QUOTE)
        )
        mask_in_gw = tpay_msgref.isin(gw_msgref_set)
        df_timeout.insert(df_timeout.columns.get_loc('MSGREF') + 1, 'CO_TRONG_GW', mask_in_gw.values)
        n_in_gw = int(mask_in_gw.sum())
        if n_in_gw > 0:
            _log(f'[B4] MSGREF check: {n_in_gw} TPAY co MSGREF trong GW → danh dau CO_TRONG_GW (can check thu cong)')
    else:
        df_timeout['CO_TRONG_GW'] = False

    _log(
        f'[B4] MIS_DI → tong truoc timeout: {len(df_mis_di):,} | '
        f'SCNL: {len(df_scnl):,} | TXRT: {len(df_txrt):,} | TPAY: {len(df_tpay):,} | '
        f'Timeout khong kenh: {len(df_timeout):,} (co trong GW: {n_in_gw}) | Final: {len(df_mis_di_final):,}'
    )
    return df_mis_di_final.reset_index(drop=True), df_timeout.reset_index(drop=True)


def xu_ly_mis_di(zip_paths: List[str], dict_gw_count: Dict[str, int], session_id: str,
                 df_gw: pd.DataFrame = None,
                 tpay_tu: datetime = None, tpay_den: datetime = None, log_callback=None):
    """
    Doc 2 zip MIS_DI song song, xu ly va tra ve (df_mis_di_final, df_timeout_khong_kenh).
    Wrapper: _doc_mis_di_raw + _process_mis_di (cho CLI va Web UI).
    De song song hoa I/O voi B3, dung _doc_mis_di_raw tach biet trong main.py Phase 1.
    """
    df_raw = _doc_mis_di_raw(zip_paths, session_id, log_callback)
    return _process_mis_di(df_raw, dict_gw_count, session_id, df_gw, tpay_tu, tpay_den, log_callback)
