import os
from datetime import datetime, timedelta

# C5: ZIP_PASSWORD doc tu bien moi truong, fallback sang gia tri mac dinh
ZIP_PASSWORD   = os.environ.get('DOI_CHIEU_ZIP_PASSWORD', 'DACwLdHi').encode()

INPUT_DIR      = './input'
OUTPUT_DIR     = './output'

# Ngay can doi chieu - dinh dang 'dd/mm/yyyy'
NGAY_DOI_CHIEU = '11/06/2026'

# C3: Danh sach cot can giu khi doc file GL02 — nguon duy nhat de tranh sai lech
COLS_NPO = [
    'TRDATE', 'TRBRCD', 'USERID', 'JOURSEQ', 'DYTRSEQ', 'LOCAC', 'CCY',
    'BUSCD', 'UNIT', 'TRCD', 'CUSTOMER', 'TRTP', 'REFERENCE',
    'REMARK', 'DRAMOUNT', 'CRAMOUNT', 'CRTDTM',
]

# ---- Khong chinh sua ben duoi ----
_ngay_dt: datetime = datetime.strptime(NGAY_DOI_CHIEU, '%d/%m/%Y')
NGAY_DT       = _ngay_dt                          # datetime object
NGAY_TRUOC_DT = _ngay_dt - timedelta(days=1)      # T-1

# Moc thoi gian cho TPAY
TPAY_TU   = NGAY_TRUOC_DT.replace(hour=23, minute=0, second=0)   # T-1 23:00:00
TPAY_DEN  = _ngay_dt.replace(hour=23, minute=0, second=0)         # T   23:00:00
