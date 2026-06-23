import io
from concurrent.futures import ThreadPoolExecutor
from typing import List
from datetime import datetime
import pyzipper
import pandas as pd
from config import ZIP_PASSWORD

_COLS = [
    'NGAY_GIAO_DICH', 'CHI_NHANH', 'REFHUB', 'MSGREF', 'MSGSEQ', 'TXID',
    'KENH_THANH_TOAN', 'TRANG_THAI_LENH', 'SO_TIEN', 'TRACE',
    'SESSION', 'LOAI_LENH_OSB', 'NH_GUI', 'NOI_DUNG',
]

_NULL_SESSION = {'', 'nan', 'None', 'NaN'}


def _doc_zip(zip_path: str, session_filter: str = None) -> pd.DataFrame:
    """
    Doc ZIP chua CSV MIS_DEN, su dung streaming (z.open) thay vi z.read() de giam RAM.
    session_filter: neu truyen, chi giu dong co SESSION == session_filter hoac SESSION null.
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
                                    sess = (chunk['SESSION'].astype(str)
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


def xu_ly_mis_den(zip_paths: List[str], session_id: str, ngay_doi_chieu: datetime, log_callback=None):
    """
    Doc 2 zip MIS_DEN song song, tra ve df_mis_den da xu ly.
    """
    sid = str(session_id)

    # Doc 2 ZIP song song, loc session ngay trong qua trinh doc
    with ThreadPoolExecutor(max_workers=2) as ex:
        futures = [ex.submit(_doc_zip, p, sid) for p in zip_paths]
        frames  = [f.result() for f in futures]
    df = pd.concat(frames, ignore_index=True)

    # Parse NGAY_GIAO_DICH
    df['NGAY_GIAO_DICH'] = pd.to_datetime(
        df['NGAY_GIAO_DICH'].str.strip(), format='%d/%m/%Y', errors='coerce'
    )

    # Chuan hoa SESSION
    df['SESSION'] = df['SESSION'].astype(str).str.strip().str.lstrip("'")
    df['SESSION_NULL'] = df['SESSION'].isin(['', 'nan', 'None', 'NaN'])

    ngay_ts = pd.Timestamp(ngay_doi_chieu.date())

    # Loc SESSION chinh xac (da pre-filter o tren, loc chinh xac them voi ngay)
    mask_ok = (df['SESSION'] == sid) | (
        df['SESSION_NULL'] & (df['NGAY_GIAO_DICH'] == ngay_ts)
    )
    df = df[mask_ok].copy()

    # Chuyen lai thanh string de xuat Excel dung format
    df['NGAY_GIAO_DICH'] = df['NGAY_GIAO_DICH'].dt.strftime('%d/%m/%Y').fillna('')

    # Bo TRANG_THAI_LENH = 'RJCT'
    df = df[df['TRANG_THAI_LENH'].astype(str).str.strip() != 'RJCT'].copy()

    # Parse SO_TIEN
    df['SO_TIEN'] = pd.to_numeric(df['SO_TIEN'], errors='coerce').fillna(0).astype('int64')

    # TRACE: bo dau nháy don roi bo leading zero
    df['TRACE'] = df['TRACE'].astype(str).str.strip().str.lstrip("'").str.lstrip('0')

    # KEY_DEN_HUB
    df['KEY_DEN_HUB'] = df['TRACE'] + df['SO_TIEN'].astype(str)

    _log = log_callback or print
    _log(f'[B6] MIS_DEN | {len(df):,} dong sau loc')
    return df.reset_index(drop=True)
