import io
import pyzipper
import pandas as pd
from config import ZIP_PASSWORD, COLS_NPO as _COLS_NPO

_COLS_REQUIRED = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']


_LOCAC_TARGET   = '502003'
_CUSTOMER_ACH   = '1000-003526275'  # Ma khach hang kenh ACH — loc khi LOCAC 502003 co nhieu CUSTOMER


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
            enc = _detect_encoding(z, name)
            for errors in ('strict', 'replace'):
                try:
                    with z.open(name) as raw_f:
                        wrapped = io.TextIOWrapper(raw_f, encoding=enc, errors=errors)
                        file_frames = []
                        for i, chunk in enumerate(
                            pd.read_csv(
                                wrapped, dtype=str,
                                usecols=lambda c: c.strip() in _COLS_NPO,
                                chunksize=100_000, low_memory=False,
                            )
                        ):
                            chunk.columns = [c.strip() for c in chunk.columns]
                            if i == 0:
                                missing = [c for c in _COLS_REQUIRED if c not in chunk.columns]
                                if missing:
                                    raise ValueError(f'Thieu cot: {missing}')
                            if 'LOCAC' in chunk.columns:
                                chunk = chunk[chunk['LOCAC'].str.strip() == _LOCAC_TARGET]
                            if 'CUSTOMER' in chunk.columns:
                                chunk = chunk[chunk['CUSTOMER'].str.strip() == _CUSTOMER_ACH]
                            if not chunk.empty:
                                file_frames.append(chunk)
                        if file_frames:
                            frames.append(pd.concat(file_frames, ignore_index=True))
                    break
                except UnicodeDecodeError:
                    if errors == 'strict':
                        print(f'[B2][WARN] Encoding detect mismatch trong {name}, thu lai errors=replace')
                        continue
                    raise
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
    if 'LOCAC' in df.columns:
        df['LOCAC'] = df['LOCAC'].astype(str).str.strip()

    # Tao SO_TRACE: vectorized str.extract thay vi Python loop (.map)
    _extracted = df['REFERENCE'].str.extract(r'[A-Za-z]+(\d+)$', expand=False)
    _stripped   = _extracted.str.lstrip('0')
    # Neu lstrip het → so la '0'; neu khong match → None (nhat quan voi logic cu)
    df['SO_TRACE']   = _stripped.where(_stripped != '', other='0').where(_extracted.notna(), other=None)
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
