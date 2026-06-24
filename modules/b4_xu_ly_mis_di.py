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

_NULL_SESSION = {'', 'nan', 'None', 'NaN'}


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
            for enc in ('utf-8-sig', 'cp1252'):
                try:
                    with z.open(name) as raw_f:
                        wrapped = io.TextIOWrapper(raw_f, encoding=enc, errors='strict')
                        if session_filter:
                            sid = str(session_filter)
                            chunk_list = []
                            for chunk in pd.read_csv(
                                wrapped, dtype=str,
                                usecols=lambda c: c in _COLS,
                                chunksize=100_000, low_memory=False,
                            ):
                                if 'SESSION' in chunk.columns:
                                    sess = (chunk['SESSION'].fillna('').astype(object).astype(str)
                                            .str.strip().str.lstrip("'"))
                                    mask = sess.isin({sid} | _NULL_SESSION)
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
                    break  # encoding thanh cong
                except UnicodeDecodeError:
                    continue
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_COLS)


def _tao_so_trace(df: pd.DataFrame) -> pd.Series:
    """SE_TRACE neu co gia tri, nguoc lai dung TRACE. Bo dau nháy don va leading zero."""
    # fillna('') + astype(object) tranh pandas 3.x StringDtype dung pd.NA (isin tra False cho pd.NA)
    se = df['SE_TRACE'].fillna('').astype(object).astype(str).str.strip().str.lstrip("'").str.lstrip('0')
    tr = df['TRACE'].fillna('').astype(object).astype(str).str.strip().str.lstrip("'").str.lstrip('0')
    has_se = se.ne('')
    return se.where(has_se, tr)


def _get_timeout_indices(df_tpay: pd.DataFrame, df_non_tpay: pd.DataFrame,
                         dict_gw_count: Dict[str, int]) -> pd.Index:
    """
    Tra ve Index cua cac dong TPAY thua so voi GW (theo doc 3.2).
    df_non_tpay: SCNL + TXRT — tat ca lenh da su dung slot GW.
    surplus = count_all_mis - count_gw; n_timeout = min(surplus, count_tpay).
    """
    count_non_tpay = df_non_tpay['CN tiền Hub'].value_counts().to_dict()

    idx_list = []
    for key, group in df_tpay.groupby('CN tiền Hub', sort=False):
        count_tpay   = len(group)
        count_gw     = dict_gw_count.get(str(key), 0)
        count_others = count_non_tpay.get(str(key), 0)
        available_gw = max(0, count_gw - count_others)
        thua = count_tpay - available_gw
        if thua > 0:
            idx_list.append(group.tail(thua).index)
    if not idx_list:
        return pd.Index([], dtype='int64')
    return pd.Index(np.concatenate([i.to_numpy() for i in idx_list]))


def xu_ly_mis_di(zip_paths: List[str], dict_gw_count: Dict[str, int], session_id: str,
                 df_gw: pd.DataFrame = None,
                 tpay_tu: datetime = None, tpay_den: datetime = None, log_callback=None):
    """
    Doc 2 zip MIS_DI song song, xu ly va tra ve (df_mis_di_final, df_timeout_khong_kenh).
    df_gw: DataFrame GW da loc (tu b3) — dung de check MSGREF cua TPAY timeout.
           TPAY co MSGREF trong GW thi thuc ra da di kenh → dua ve MIS_DI_FINAL.
    tpay_tu / tpay_den: neu truyen thi dung gia tri nay (thread-safe cho Web UI);
                        neu None thi lay tu config (CLI mode).
    """
    import config
    _tpay_tu  = tpay_tu  if tpay_tu  is not None else config.TPAY_TU
    _tpay_den = tpay_den if tpay_den is not None else config.TPAY_DEN

    _log = log_callback or print
    _log('[B4] Doc MIS_DI tu 2 ZIP...')

    sid = str(session_id)

    # Doc 2 ZIP song song, loc session ngay trong qua trinh doc (giam RAM ~60-70%)
    with ThreadPoolExecutor(max_workers=2) as ex:
        futures = [ex.submit(_doc_zip, p, sid) for p in zip_paths]
        frames  = [f.result() for f in futures]
    df = pd.concat(frames, ignore_index=True)

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

    # Tao KEY_HUB
    df_mis_di['KEY_HUB'] = (
        df_mis_di['CHI_NHANH'].astype(str).str.strip()
        + df_mis_di['SO_TRACE']
        + df_mis_di['SO_TIEN'].astype(str)
    )

    # Them cot "CN tien Hub"
    cn_tien = df_mis_di['CHI_NHANH'].astype(str).str.strip() + df_mis_di['SO_TIEN'].astype(str)
    loc = df_mis_di.columns.get_loc('CHI_NHANH') + 1
    df_mis_di.insert(loc, 'CN tiền Hub', cn_tien)

    # Tinh timeout — SCNL va TXRT deu chiem slot GW (theo doc 3.2: dem ca SCNL+TXRT+TPAY vs GW)
    df_non_tpay_in_mis = df_mis_di[df_mis_di['TRANG_THAI_LENH'].isin(['SCNL', 'TXRT'])]
    df_tpay_in_mis = df_mis_di[df_mis_di['TRANG_THAI_LENH'] == 'TPAY']
    timeout_idx = _get_timeout_indices(df_tpay_in_mis, df_non_tpay_in_mis, dict_gw_count)

    df_timeout_candidates = df_mis_di.loc[timeout_idx].copy()
    df_mis_di_final       = df_mis_di[~df_mis_di.index.isin(timeout_idx)].copy()

    # MSGREF check: TPAY co MSGREF trong GW thi thuc ra da di kenh → dua ve MIS_DI_FINAL
    _QUOTE = "'"
    n_rescued = 0
    if df_gw is not None and 'MSGREF' in df_gw.columns and len(df_timeout_candidates) > 0:
        gw_msgref_set = set(
            df_gw['MSGREF'].fillna('').astype(object).astype(str)
            .str.strip().str.lstrip(_QUOTE)
        )
        tpay_msgref = (
            df_timeout_candidates['MSGREF'].fillna('').astype(object).astype(str)
            .str.strip().str.lstrip(_QUOTE)
        )
        mask_in_gw = tpay_msgref.isin(gw_msgref_set)
        df_rescued  = df_timeout_candidates[mask_in_gw]
        df_timeout  = df_timeout_candidates[~mask_in_gw]
        if len(df_rescued) > 0:
            df_mis_di_final = pd.concat([df_mis_di_final, df_rescued]).reset_index(drop=True)
            n_rescued = len(df_rescued)
            _log(f'[B4] MSGREF check: {n_rescued} TPAY co MSGREF trong GW → chuyen vao MIS_DI_FINAL (da di kenh)')
    else:
        df_timeout = df_timeout_candidates

    _log(
        f'[B4] MIS_DI → tong truoc timeout: {len(df_mis_di):,} | '
        f'SCNL: {len(df_scnl):,} | TXRT: {len(df_txrt):,} | TPAY: {len(df_tpay):,} | '
        f'Timeout khong kenh: {len(df_timeout):,} | Final: {len(df_mis_di_final):,}'
    )
    return df_mis_di_final.reset_index(drop=True), df_timeout.reset_index(drop=True)
