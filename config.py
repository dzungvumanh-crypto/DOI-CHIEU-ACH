from datetime import datetime, timedelta

ZIP_PASSWORD   = b'DACwLdHi'
INPUT_DIR      = './input'
OUTPUT_DIR     = './output'

# Ngay can doi chieu - dinh dang 'dd/mm/yyyy'
NGAY_DOI_CHIEU = '11/06/2026'

# ---- Khong chinh sua ben duoi ----
_ngay_dt: datetime = datetime.strptime(NGAY_DOI_CHIEU, '%d/%m/%Y')
NGAY_DT       = _ngay_dt                          # datetime object
NGAY_TRUOC_DT = _ngay_dt - timedelta(days=1)      # T-1

# Moc thoi gian cho TPAY
TPAY_TU   = NGAY_TRUOC_DT.replace(hour=23, minute=0, second=0)   # T-1 23:00:00
TPAY_DEN  = _ngay_dt.replace(hour=23, minute=0, second=0)         # T   23:00:00
