import pandas as pd


def _doc_mot_sheet(xl: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """
    Doc mot sheet GW, tu dong phat hien header row.
    Tim dong dau tien chua gia tri 'BRCD' de dung lam header.
    """
    # Doc raw khong co header de tim vi tri header row
    df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10,
                           dtype=str, engine='calamine')
    header_row = 0
    for i, row in df_raw.iterrows():
        if 'BRCD' in row.values:
            header_row = i
            break

    # Doc lai voi dung header row
    df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row,
                       dtype=str, engine='calamine')
    df.columns = [str(c).strip() for c in df.columns]
    return df


def xu_ly_gw(xlsx_path: str, session_id: str):
    """
    Doc file GW Excel (tat ca sheet), tra ve (dict_gw_count, df_gw_raw).
    dict_gw_count: {KEY_GW: count}  —  KEY_GW = str(BRCD) + str(STTLMAMT_int)
    df_gw_raw: DataFrame day du sau khi loc, de xuat sheet RAW_GW
    """
    xl = pd.ExcelFile(xlsx_path, engine='calamine')

    frames = [_doc_mot_sheet(xl, s) for s in xl.sheet_names]
    df = pd.concat(frames, ignore_index=True)

    # Loc session
    df = df[df['SessionId'].astype(str).str.strip() == str(session_id)].copy()

    # Bo PrcFlg = 'ACH Tu choi'
    df = df[df['PrcFlg'].astype(str).str.strip() != 'ACH Từ chối'].copy()

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

    print(f'[B3] GW | {len(df):,} dong (session {session_id}) | {len(dict_gw_count):,} KEY_GW unique')
    return dict_gw_count, df.reset_index(drop=True)
