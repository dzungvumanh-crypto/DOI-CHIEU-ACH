import re
import io
import pyzipper
import pandas as pd
from config import ZIP_PASSWORD

_COLS = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']

_RE_TRACE = re.compile(r'[A-Za-z]+(\d+)$')


def _so_trace(ref: str):
    """Lay phan so cuoi cua REFERENCE. Tra None neu khong co."""
    m = _RE_TRACE.search(str(ref))
    return m.group(1) if m else None


def _doc_zip(zip_path: str) -> pd.DataFrame:
    frames = []
    with pyzipper.AESZipFile(zip_path, 'r') as z:
        z.setpassword(ZIP_PASSWORD)
        for name in z.namelist():
            if name.lower().endswith('.csv'):
                raw = z.read(name)
                for enc in ('utf-8-sig', 'cp1252'):
                    try:
                        df = pd.read_csv(
                            io.BytesIO(raw),
                            dtype=str,
                            usecols=_COLS,
                            encoding=enc,
                            low_memory=False,
                        )
                        frames.append(df)
                        break
                    except UnicodeDecodeError:
                        continue
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_COLS)


def xu_ly_gl02(zip_path: str):
    """
    Doc GL02 zip, tra ve (df_npo_di, df_npo_den).
    """
    df = _doc_zip(zip_path)

    # Parse so tien thanh int64
    df['CRAMOUNT'] = pd.to_numeric(df['CRAMOUNT'], errors='coerce').fillna(0).astype('int64')
    df['DRAMOUNT'] = pd.to_numeric(df['DRAMOUNT'], errors='coerce').fillna(0).astype('int64')

    # Tao SO_TRACE
    df['SO_TRACE'] = df['REFERENCE'].map(_so_trace)

    # Tao KEY (SO_TRACE co the la None → se khong khop voi MIS → tu dong ra THUA)
    df['_trace_str'] = df['SO_TRACE'].fillna('')

    # NPO_DI: CRAMOUNT != 0
    npo_di = df[df['CRAMOUNT'] != 0].copy()
    npo_di['KEY_DI'] = (
        npo_di['TRBRCD'].str.strip()
        + npo_di['_trace_str']
        + npo_di['CRAMOUNT'].astype(str)
    )

    # NPO_DEN: DRAMOUNT != 0
    # KEY_DEN khong co TRBRCD vi GL02 luon post vao branch 5227 (clearing),
    # trong khi MIS_DEN dung CHI_NHANH thuc cua khach hang
    npo_den = df[df['DRAMOUNT'] != 0].copy()
    npo_den['KEY_DEN'] = (
        npo_den['_trace_str']
        + npo_den['DRAMOUNT'].astype(str)
    )

    print(f'[B2] GL02 | NPO_DI: {len(npo_di):,} dong | NPO_DEN: {len(npo_den):,} dong')
    return npo_di.reset_index(drop=True), npo_den.reset_index(drop=True)
