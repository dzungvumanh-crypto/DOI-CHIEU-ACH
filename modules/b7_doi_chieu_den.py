import pandas as pd
from modules.b5_doi_chieu_di import _doi_chieu


def doi_chieu_den(df_npo_den: pd.DataFrame, df_mis_den: pd.DataFrame, log_callback=None):
    """
    Doi chieu chieu DEN: KEY_DEN (NPO) vs KEY_DEN_HUB (MIS).
    Tra ve (df_mis_den_khop, df_npo_den_thua, df_mis_den_thua).
    """
    npo_k, mis_k, npo_t, mis_t = _doi_chieu(
        df_npo_den, 'KEY_DEN',
        df_mis_den, 'KEY_DEN_HUB',
        'B7', log_callback,
    )
    return mis_k, npo_t, mis_t
