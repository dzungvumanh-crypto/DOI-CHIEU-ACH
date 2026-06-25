import pandas as pd


def _doc_mot_sheet(xl: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """Doc mot sheet GW, tu dong phat hien header row. Doc toan bo 1 lan, khong doc 2 lan."""
    df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None,
                           dtype=str, engine='calamine')
    header_row = 0
    for i, row in df_raw.iterrows():
        if 'BRCD' in row.values:
            header_row = i
            break
    df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df_raw.iloc[header_row]]
    return df


def _sheet_co_session(xl: pd.ExcelFile, sheet_name: str, session_id: str) -> bool:
    """Peek 60 dong dau de kiem tra sheet co chua SessionId can tim khong.
    Tranh doc ca trieu dong khi phan lon sheets khong thuoc session hien tai."""
    try:
        df_peek = pd.read_excel(xl, sheet_name=sheet_name, header=None,
                                nrows=60, dtype=str, engine='calamine')
    except Exception:
        return False
    for i, row in df_peek.iterrows():
        if 'SessionId' in row.values:
            try:
                sid_idx = list(row).index('SessionId')
            except ValueError:
                return False
            data = df_peek.iloc[i + 1:, sid_idx].astype(str)
            return str(session_id) in data.values
    return False  # Khong co cot SessionId → bo qua sheet nay


def xu_ly_gw(xlsx_path: str, session_id: str, log_callback=None):
    """
    Doc file GW Excel (chi cac sheet co session can tim), tra ve (dict_gw_count, df_gw_raw).
    dict_gw_count: {KEY_GW: count}  —  KEY_GW = str(BRCD) + str(STTLMAMT_int)
    df_gw_raw: DataFrame day du sau khi loc, de xuat sheet RAW_GW
    Toi uu: peek 60 dong moi sheet truoc — bo qua sheet khong co session (tranh doc 2M+ dong).
    """
    xl = pd.ExcelFile(xlsx_path, engine='calamine')

    # Lazy reading: chi doc sheet co session phu hop
    matching = [s for s in xl.sheet_names if _sheet_co_session(xl, s, session_id)]
    if matching:
        frames = [_doc_mot_sheet(xl, s) for s in matching]
    else:
        # Fallback: doc het neu peek khong tim thay (an toan)
        frames = [_doc_mot_sheet(xl, s) for s in xl.sheet_names]

    df = pd.concat(frames, ignore_index=True)

    # Loai ban ghi trung theo MSGREF: xay ra khi GW file co sheet phu
    # (VD: "di GW 12.06") la ban sao loc cua sheet chinh (Sheet 1)
    if 'MSGREF' in df.columns:
        df = df.drop_duplicates(subset=['MSGREF'])

    # Loc session va bo PrcFlg = 'ACH Tu choi' trong 1 buoc (giam 1 lan allocate)
    mask = (
        (df['SessionId'].astype(str).str.strip() == str(session_id)) &
        (df['PrcFlg'].astype(str).str.strip() != 'ACH Từ chối')
    )
    df = df[mask].copy()

    # Parse STTLMAMT: bo 'VND', dau phay, khoang trang -> int64
    df['STTLMAMT'] = (
        df['STTLMAMT']
        .astype(str)
        .str.replace(r'[VND,\s]', '', regex=True)
    )
    df['STTLMAMT'] = pd.to_numeric(df['STTLMAMT'], errors='coerce').fillna(0).astype('int64')

    # KEY_GW = BRCD + STTLMAMT (string concat)
    df['KEY_GW'] = df['BRCD'].astype(str).str.strip() + df['STTLMAMT'].astype(str)

    dict_gw_count = df['KEY_GW'].value_counts().to_dict()

    _log = log_callback or print
    _log(f'[B3] GW | {len(df):,} dong (session {session_id}) | {len(dict_gw_count):,} KEY_GW unique')
    return dict_gw_count, df.reset_index(drop=True)
