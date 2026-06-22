import pandas as pd


def phan_tich(df_npo_di_thua, df_mis_di_thua, df_npo_den_thua, df_mis_den_thua,
              n_di_khop, n_den_khop, df_timeout):
    """
    Phan tich chat luong doi chieu. Tra ve DataFrame 3 cot ('Chi tieu', 'Gia tri', 'Ghi chu')
    de ghi vao sheet PHAN_TICH.
    """
    def safe_len(df):
        return len(df) if df is not None else 0

    n_npo_di  = n_di_khop + safe_len(df_npo_di_thua)
    n_mis_di  = n_di_khop + safe_len(df_mis_di_thua)
    n_npo_den = n_den_khop + safe_len(df_npo_den_thua)
    n_mis_den = n_den_khop + safe_len(df_mis_den_thua)

    def pct(num, den):
        return f'{num / den * 100:.2f}' if den > 0 else 'N/A'

    timeout_tien = 0
    if df_timeout is not None and len(df_timeout) > 0 and 'SO_TIEN' in df_timeout.columns:
        timeout_tien = int(pd.to_numeric(df_timeout['SO_TIEN'], errors='coerce').fillna(0).sum())

    rows = []

    # === Section 1: KPI ===
    rows.append(('--- 1. CHI TIEU CHAT LUONG DOI CHIEU ---', '', ''))
    rows.append(('Ty le khop NPO_DI (%)', pct(n_di_khop, n_npo_di),   f'{n_di_khop:,} / {n_npo_di:,}'))
    rows.append(('Ty le khop MIS_DI (%)', pct(n_di_khop, n_mis_di),   f'{n_di_khop:,} / {n_mis_di:,}'))
    rows.append(('Ty le khop NPO_DEN (%)', pct(n_den_khop, n_npo_den), f'{n_den_khop:,} / {n_npo_den:,}'))
    rows.append(('Ty le khop MIS_DEN (%)', pct(n_den_khop, n_mis_den), f'{n_den_khop:,} / {n_mis_den:,}'))
    rows.append(('Timeout khong di kenh', f'{safe_len(df_timeout):,} lenh', f'{timeout_tien:,} VND'))
    rows.append(('', '', ''))

    # === Section 2: MIS_DI_THUA breakdown ===
    rows.append(('--- 2. PHAN TICH MIS_DI_THUA ---', '', ''))
    total_mis_di = safe_len(df_mis_di_thua)
    if total_mis_di > 0 and 'TRANG_THAI_LENH' in df_mis_di_thua.columns:
        _NOTE = {
            'SCNL': 'Da thanh toan — co the thuoc session khac, binh thuong',
            'TPAY': 'Chua duoc xu ly — can theo doi',
            'TXRT': 'Hoan tra — can kiem tra',
        }
        for tt, cnt in df_mis_di_thua['TRANG_THAI_LENH'].value_counts().items():
            rows.append((f'  {tt}', f'{cnt:,}  ({cnt / total_mis_di * 100:.1f}%)', _NOTE.get(str(tt), '')))

        if (df_npo_di_thua is not None and len(df_npo_di_thua) > 0
                and {'TRBRCD', 'CRAMOUNT'} <= set(df_npo_di_thua.columns)
                and {'CHI_NHANH', 'SO_TIEN'} <= set(df_mis_di_thua.columns)):
            npo_pairs = set(zip(
                df_npo_di_thua['TRBRCD'].astype(str).str.strip(),
                df_npo_di_thua['CRAMOUNT'].astype(str)
            ))
            mis_pairs = set(zip(
                df_mis_di_thua['CHI_NHANH'].astype(str).str.strip(),
                df_mis_di_thua['SO_TIEN'].astype(str)
            ))
            overlap = len(npo_pairs & mis_pairs)
            rows.append(('  Cap (CN+TIEN) co mat o ca 2 phia', f'{overlap:,}',
                          'Cung chi nhanh + so tien nhung TRACE khac — Nen kiem tra'))
    else:
        rows.append(('  (Khong co du lieu)', '', ''))
    rows.append(('', '', ''))

    # === Section 3: Top 10 CHI_NHANH MIS_DI_THUA ===
    rows.append(('--- 3. TOP 10 CHI_NHANH CO MIS_DI_THUA NHIEU NHAT ---', '', ''))
    rows.append(('  CHI_NHANH', 'MIS_DI_THUA', 'NPO_DI_THUA'))
    if (total_mis_di > 0 and 'CHI_NHANH' in df_mis_di_thua.columns):
        top10 = df_mis_di_thua.groupby('CHI_NHANH').size().nlargest(10)
        npo_by_cn = {}
        if df_npo_di_thua is not None and 'TRBRCD' in df_npo_di_thua.columns:
            npo_by_cn = df_npo_di_thua.groupby('TRBRCD').size().to_dict()
        for cn, cnt_mis in top10.items():
            cnt_npo = npo_by_cn.get(str(cn), 0)
            rows.append((f'  {cn}', f'{cnt_mis:,}', f'{cnt_npo:,}'))
    rows.append(('', '', ''))

    # === Section 4: MIS_DEN_THUA breakdown ===
    rows.append(('--- 4. PHAN TICH MIS_DEN_THUA ---', '', ''))
    total_mis_den = safe_len(df_mis_den_thua)
    if total_mis_den > 0 and 'TRANG_THAI_LENH' in df_mis_den_thua.columns:
        for tt, cnt in df_mis_den_thua['TRANG_THAI_LENH'].value_counts().items():
            rows.append((f'  {tt}', f'{cnt:,}  ({cnt / total_mis_den * 100:.1f}%)', ''))

        if (df_npo_den_thua is not None and len(df_npo_den_thua) > 0
                and {'TRBRCD', 'DRAMOUNT'} <= set(df_npo_den_thua.columns)
                and {'CHI_NHANH', 'SO_TIEN'} <= set(df_mis_den_thua.columns)):
            npo_den_pairs = set(zip(
                df_npo_den_thua['TRBRCD'].astype(str).str.strip(),
                df_npo_den_thua['DRAMOUNT'].astype(str)
            ))
            mis_den_pairs = set(zip(
                df_mis_den_thua['CHI_NHANH'].astype(str).str.strip(),
                df_mis_den_thua['SO_TIEN'].astype(str)
            ))
            overlap_den = len(npo_den_pairs & mis_den_pairs)
            rows.append(('  Cap (CN+TIEN) co mat o ca 2 phia', f'{overlap_den:,}',
                          'Cung chi nhanh + so tien nhung TRACE khac — Nen kiem tra'))
    else:
        rows.append(('  (Khong co du lieu)', '', ''))

    return pd.DataFrame(rows, columns=['Chi tieu', 'Gia tri', 'Ghi chu'])
