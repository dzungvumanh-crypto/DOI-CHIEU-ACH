import pandas as pd


def _doi_chieu(df_npo: pd.DataFrame, key_npo: str,
               df_mis: pd.DataFrame, key_mis: str,
               label: str, log_callback=None):
    """
    Doi chieu theo count (vectorized cumcount).
    Tra ve (df_npo_khop, df_mis_khop, df_npo_thua, df_mis_thua).
    """
    # Dem count tung KEY moi phia
    cnt_npo = df_npo.groupby(key_npo, sort=False).size()
    cnt_mis = df_mis.groupby(key_mis, sort=False).size()

    common = set(cnt_npo.index) & set(cnt_mis.index)
    dict_min = {k: min(int(cnt_npo[k]), int(cnt_mis[k])) for k in common}

    # --- NPO side ---
    cc_npo = df_npo.groupby(key_npo, sort=False).cumcount()
    npo_min = df_npo[key_npo].map(dict_min).fillna(0).astype(int)
    df_npo_khop = df_npo[cc_npo < npo_min].copy()
    df_npo_thua = df_npo[cc_npo >= npo_min].copy()

    # --- MIS side ---
    cc_mis = df_mis.groupby(key_mis, sort=False).cumcount()
    mis_min = df_mis[key_mis].map(dict_min).fillna(0).astype(int)
    df_mis_khop = df_mis[cc_mis < mis_min].copy()
    df_mis_thua = df_mis[cc_mis >= mis_min].copy()

    _log = log_callback or print
    _log(
        f'[{label}] Khop: NPO={len(df_npo_khop):,} MIS={len(df_mis_khop):,} | '
        f'NPO thua: {len(df_npo_thua):,} | MIS thua: {len(df_mis_thua):,}'
    )
    return df_npo_khop, df_mis_khop, df_npo_thua, df_mis_thua


def doi_chieu_di(df_npo_di: pd.DataFrame, df_mis_di_final: pd.DataFrame, log_callback=None):
    """
    Doi chieu chieu DI: KEY_DI (NPO) vs KEY_HUB (MIS).
    Tra ve (df_mis_di_khop, df_npo_di_thua, df_mis_di_thua).
    """
    npo_k, mis_k, npo_t, mis_t = _doi_chieu(
        df_npo_di, 'KEY_DI',
        df_mis_di_final, 'KEY_HUB',
        'B5', log_callback,
    )
    return mis_k, npo_t, mis_t
