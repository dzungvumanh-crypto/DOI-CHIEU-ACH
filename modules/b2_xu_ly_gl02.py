import re
import io
import pyzipper
import pandas as pd
from config import ZIP_PASSWORD, COLS_NPO as _COLS_NPO

_COLS_REQUIRED = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']

_RE_TRACE = re.compile(r'[A-Za-z]+(\d+)$')


def _so_trace(ref: str):
    """Lay phan so cuoi cua REFERENCE. Tra None neu khong co."""
    m = _RE_TRACE.search(str(ref))
    return m.group(1) if m else None


_LOCAC_TARGET = '502003'


def _doc_zip(zip_path: str) -> pd.DataFrame:
    """
    Doc ZIP co nhieu CSV, loc LOCAC=502003 tung chunk (khong doc toan bo roi filter sau).
    Toi uu: skip ca file neu dong dau khong phai LOCAC_TARGET.
    """
    frames = []
    with pyzipper.AESZipFile(zip_path, 'r') as z:
        z.setpassword(ZIP_PASSWORD)
        for name in sorted(z.namelist()):
            if not name.lower().endswith('.csv'):
                continue
            for enc in ('utf-8-sig', 'cp1252'):
                try:
                    with z.open(name) as raw_f:
                        wrapped = io.TextIOWrapper(raw_f, encoding=enc, errors='strict')
                        file_frames = []
                        skip_file   = False
                        for i, chunk in enumerate(
                            pd.read_csv(
                                wrapped, dtype=str,
                                usecols=lambda c: c.strip() in _COLS_NPO,
                                chunksize=50_000, low_memory=False,
                            )
                        ):
                            chunk.columns = [c.strip() for c in chunk.columns]
                            if i == 0:
                                missing = [c for c in _COLS_REQUIRED if c not in chunk.columns]
                                if missing:
                                    raise ValueError(f'Thieu cot: {missing}')
                                # Toi uu: skip ca file neu dong dau co LOCAC khac target
                                if ('LOCAC' in chunk.columns and len(chunk) > 0
                                        and chunk.iloc[0]['LOCAC'].strip() != _LOCAC_TARGET):
                                    skip_file = True
                                    break
                            # Loc LOCAC ngay tai chunk — khong giu dong thua
                            if 'LOCAC' in chunk.columns:
                                chunk = chunk[chunk['LOCAC'].str.strip() == _LOCAC_TARGET]
                            if not chunk.empty:
                                file_frames.append(chunk)
                        if not skip_file and file_frames:
                            frames.append(pd.concat(file_frames, ignore_index=True))
                    break  # encoding thanh cong
                except UnicodeDecodeError:
                    continue
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_COLS_NPO)


def xu_ly_gl02(zip_path: str, log_callback=None):
    """
    Doc GL02 zip, tra ve (df_npo_di, df_npo_den).
    """
    df = _doc_zip(zip_path)

    # Parse so tien thanh int64
    df['CRAMOUNT'] = pd.to_numeric(df['CRAMOUNT'], errors='coerce').fillna(0).astype('int64')
    df['DRAMOUNT'] = pd.to_numeric(df['DRAMOUNT'], errors='coerce').fillna(0).astype('int64')

    # LOCAC da duoc loc trong _doc_zip — khong can filter lai
    df['LOCAC'] = df['LOCAC'].astype(str).str.strip()

    # Tao SO_TRACE
    df['SO_TRACE'] = df['REFERENCE'].map(_so_trace)
    df['_trace_str'] = df['SO_TRACE'].fillna('')

    # NPO_DI: LOCAC=502003 AND CRAMOUNT != 0 (giao dich di - credit)
    npo_di = df[df['CRAMOUNT'] != 0].copy()
    npo_di['KEY_DI'] = (
        npo_di['TRBRCD'].str.strip()
        + npo_di['_trace_str']
        + npo_di['CRAMOUNT'].astype(str)
    )

    # NPO_DEN: LOCAC=502003 AND CRAMOUNT == 0 (giao dich den - debit, DRAMOUNT != 0)
    npo_den = df[df['CRAMOUNT'] == 0].copy()
    npo_den['KEY_DEN'] = (
        npo_den['_trace_str']
        + npo_den['DRAMOUNT'].astype(str)
    )

    _log = log_callback or print
    _log(f'[B2] GL02 | NPO_DI: {len(npo_di):,} dong | NPO_DEN: {len(npo_den):,} dong')
    return npo_di.reset_index(drop=True), npo_den.reset_index(drop=True)
