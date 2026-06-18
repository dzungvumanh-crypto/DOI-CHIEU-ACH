# KẾ HOẠCH CODE — ĐỐI CHIẾU ACH (GL02 vs MIS)
## Dùng cho Claude Code

---

## 1. CẤU TRÚC THƯ MỤC

```
doi_chieu_ACH/
├── main.py
├── config.py
├── modules/
│   ├── __init__.py
│   ├── b1_doc_session.py
│   ├── b2_xu_ly_gl02.py
│   ├── b3_xu_ly_gw.py
│   ├── b4_xu_ly_mis_di.py
│   ├── b5_doi_chieu_di.py
│   ├── b6_xu_ly_mis_den.py
│   └── b7_doi_chieu_den.py
├── input/
│   ├── GL02_yyyymmdd_1000.zip
│   ├── doichieugd_yyyymmdd__01_DI_9999_N.zip   (ngày T-1)
│   ├── doichieugd_yyyymmdd__01_DI_9999_N.zip   (ngày T)
│   ├── doichieugd_yyyymmdd__01_DEN_9999_N.zip  (ngày T-1)
│   ├── doichieugd_yyyymmdd__01_DEN_9999_N.zip  (ngày T)
│   ├── di_GW_dd_mm.xlsx
│   └── ACH_yyyymmdd_VBAAVNVN_NRT_SSSSS_N03_1.pdf
└── output/
    └── doi_chieu_YYYYMMDD.xlsx
```

---

## 2. THƯ VIỆN CẦN CÀI

```
pip install pandas openpyxl xlsxwriter pyzipper re
```

---

## 3. config.py

```python
ZIP_PASSWORD = b'DACwLdHi'
INPUT_DIR    = './input'
OUTPUT_DIR   = './output'

# Ngày cần đối chiếu (truyền vào khi chạy hoặc tự nhận từ tên file GL02)
# Định dạng: 'dd/mm/yyyy'  ví dụ: '11/06/2026'
NGAY_DOI_CHIEU = '11/06/2026'
```

---

## 4. BƯỚC 1 — b1_doc_session.py

**Mục tiêu:** Đọc số SESSION từ tên file PDF.

**Input:** Tên file PDF — vd: `ACH_20260612_VBAAVNVN_NRT_15882_N03_1.pdf`

**Logic:**
```
- Dùng regex: r'_NRT_(\d+)_' để lấy số session
- Trả về: session_id = '15882' (string)
```

**Output:** `session_id` (string)

---

## 5. BƯỚC 2 — b2_xu_ly_gl02.py

**Mục tiêu:** Đọc GL02.zip → tạo NPO_ĐI và NPO_ĐẾN.

**Input:** GL02_yyyymmdd_1000.zip (password DACwLdHi)

**Cấu trúc file bên trong zip:**
- 2 file CSV, cùng cấu trúc cột
- Cột: TRDATE, TRBRCD, USERID, JOURSEQ, DYTRSEQ, LOCAC, CCY,
        BUSCD, UNIT, TRCD, CUSTOMER, TRTP, REFERENCE,
        REMARK, DRAMOUNT, CRAMOUNT, CRTDTM

**Logic chi tiết:**

```
BƯỚC 2.1 — Giải nén & gộp:
  - Giải nén cả 2 file CSV trong zip (password DACwLdHi)
  - Gộp thành 1 DataFrame

BƯỚC 2.2 — Tạo SO_TRACE từ cột REFERENCE:
  - REFERENCE dạng '1000API145934227' → SO_TRACE = '145934227'
    (lấy phần số sau chuỗi chữ cái: regex r'[A-Za-z]+(\d+)$')
  - REFERENCE dạng '1000NPO000000002' → SO_TRACE = '000000002'
  - REFERENCE = '1000OSB' → SO_TRACE = None (null)

BƯỚC 2.3 — Tách NPO_ĐI:
  - Điều kiện: CRAMOUNT != 0
  - Tạo cột khóa: KEY_DI = str(TRBRCD) + str(SO_TRACE) + str(CRAMOUNT)
  - Ghi chú: dòng SO_TRACE=None → KEY_DI sẽ không khớp với MIS
              → tự động ra NPO_ĐI THỪA, không cần loại trước

BƯỚC 2.4 — Tách NPO_ĐẾN:
  - Điều kiện: DRAMOUNT != 0
  - Tạo cột khóa: KEY_DEN = str(TRBRCD) + str(SO_TRACE) + str(DRAMOUNT)
```

**Output:** df_npo_di, df_npo_den

---

## 6. BƯỚC 3 — b3_xu_ly_gw.py

**Mục tiêu:** Xử lý file GW → tạo tập so sánh với TPAY.

**Input:** di_GW_dd_mm.xlsx (Sheet tên chứa 'GW', bắt đầu từ dòng 6 — bỏ 5 dòng header)

**Cấu trúc cột:**
STT, MSGREF, BRCD, CN tiền GW, (cột trống), Kênh kết nối,
Ngày giờ, NH Thụ hưởng, STTLMAMT, Ngày giờ Hạch Toán,
POSTINGAMT, STTLMDT, CORETRACE, DBTRNM, DBTRACCT,
CDTRACCT, CDTRNM, Nội dung Thanh toán, StsFlg,
PrcFlg, SessionId, Số bút toán, Ghi chú

**Logic chi tiết:**

```
BƯỚC 3.1 — Lọc Session:
  - Chỉ lấy dòng SessionId = session_id (vd: '15882')

BƯỚC 3.2 — Bỏ trạng thái:
  - Bỏ PrcFlg = 'ACH Từ chối'
  - Giữ lại: 'Lệnh Hoàn thành', 'Lệnh Timeout', 'Chờ hoàn trả'

BƯỚC 3.3 — Xử lý STTLMAMT:
  - Bỏ chữ 'VND' và dấu phẩy
  - Chuyển sang số nguyên (int)
  - Ví dụ: '459,000VND' → 459000

BƯỚC 3.4 — Tạo khóa:
  - KEY_GW = str(BRCD) + str(STTLMAMT)
  - Đếm số lần xuất hiện mỗi KEY_GW (value_counts)
    → tạo dict: {KEY_GW: count}
```

**Output:** dict_gw_count {KEY_GW: count}

---

## 7. BƯỚC 4 — b4_xu_ly_mis_di.py

**Mục tiêu:** Xử lý 2 file zip MIS_ĐI → tạo tập đối chiếu + tách TIMEOUT.

**Input:**
- doichieugd_yyyymmdd__01_DI_9999_N.zip × 2 (ngày T-1 và T)
- dict_gw_count (từ Bước 3)
- session_id, NGAY_DOI_CHIEU

**Cấu trúc cột MIS_ĐI:**
NGAY_GIAO_DICH, CHI_NHANH, REFHUB, MSGREF, MSGSEQ, TXID,
KENH_THANH_TOAN, TRANG_THAI_LENH, SO_TIEN, TRACE,
SE_TRACE, SESSION, LOAI_LENH_OSB, NH_NHAN,
MA_GIAO_DICH, NOI_DUNG, NGAY_KENH_TRA

**Logic chi tiết:**

```
BƯỚC 4.1 — Đọc & gộp 2 file:
  - Giải nén 2 zip (password DACwLdHi), đọc CSV
  - Gộp thành 1 DataFrame

BƯỚC 4.2 — Bỏ trạng thái loại trừ:
  - Bỏ TRANG_THAI_LENH trong ('CALD', 'ERPO', 'TPER')

BƯỚC 4.3 — Tạo SO_TRACE:
  - Nếu SE_TRACE không null và không rỗng → SO_TRACE = SE_TRACE
  - Ngược lại → SO_TRACE = TRACE
  - Strip dấu nháy đơn đầu chuỗi nếu có (vd: '145789877 → 145789877)

BƯỚC 4.4 — Tạo khóa MIS_ĐI:
  - KEY_HUB = str(CHI_NHANH) + str(SO_TRACE) + str(SO_TIEN)

BƯỚC 4.5 — Lọc theo trạng thái & session:

  [SCNL]:
  - Chỉ lấy SESSION = session_id
  - Bỏ SESSION khác hoặc null

  [TXRT]:
  - Lấy toàn bộ (không lọc session)

  [TPAY]:
  - Lấy SESSION = session_id
  - Lấy SESSION = null VÀ NGAY_KENH_TRA >= (NGAY_DOI_CHIEU - 1 ngày) 23:00:00
                       VÀ NGAY_KENH_TRA <   NGAY_DOI_CHIEU 23:00:00
    Ví dụ đối chiếu ngày 11/6:
    NGAY_KENH_TRA >= 10/06/2026 23:00:00
    NGAY_KENH_TRA <  11/06/2026 23:00:00
  - Bỏ SESSION không phải session_id và không phải null

BƯỚC 4.6 — Gộp tất cả lại:
  - df_mis_di = concat(df_scnl, df_txrt, df_tpay)

BƯỚC 4.7 — Xử lý TPAY vs GW (so khớp bằng count):
  - Tạo cột KEY_TIEN = str(CHI_NHANH) + str(SO_TIEN)
    (chỉ dùng CHI_NHANH+SO_TIEN, không dùng TRACE)
  - Đếm KEY_TIEN trong df_mis_di chỉ với TPAY: dict_mis_tpay_count
  - So sánh với dict_gw_count:
    - Với mỗi KEY_TIEN trong TPAY:
      thừa = count_mis - count_gw  (nếu count_gw không có → thừa = count_mis)
      → Nếu thừa > 0: lấy 'thừa' dòng cuối cùng của KEY_TIEN đó
        đẩy vào df_timeout_khong_kenh
  - df_mis_di_final = df_mis_di trừ đi df_timeout_khong_kenh
```

**Output:** df_mis_di_final, df_timeout_khong_kenh

---

## 8. BƯỚC 5 — b5_doi_chieu_di.py

**Mục tiêu:** Đối chiếu NPO_ĐI vs MIS_ĐI theo count.

**Input:** df_npo_di, df_mis_di_final

**Logic chi tiết:**

```
BƯỚC 5.1 — Đếm KEY theo count (không phải chỉ có/không có):
  - dict_npo = value_counts của KEY_DI trong df_npo_di
  - dict_mis = value_counts của KEY_HUB trong df_mis_di_final

BƯỚC 5.2 — So sánh:
  - Với mỗi KEY xuất hiện ở cả 2 phía:
    min_count = min(count_npo, count_npo)
    → Lấy min_count dòng từ mỗi phía → df_mis_di_khop
  - Dòng thừa ở NPO_ĐI → df_npo_di_thua
  - Dòng thừa ở MIS_ĐI → df_mis_di_thua
```

**Output:** df_mis_di_khop, df_npo_di_thua, df_mis_di_thua

---

## 9. BƯỚC 6 — b6_xu_ly_mis_den.py

**Mục tiêu:** Xử lý 2 file zip MIS_ĐẾN → tạo tập đối chiếu.

**Input:**
- doichieugd_yyyymmdd__01_DEN_9999_N.zip × 2 (ngày T-1 và T)
- session_id, NGAY_DOI_CHIEU

**Cấu trúc cột MIS_ĐẾN:**
NGAY_GIAO_DICH, CHI_NHANH, REFHUB, MSGREF, MSGSEQ, TXID,
KENH_THANH_TOAN, TRANG_THAI_LENH, SO_TIEN, TRACE,
SESSION, LOAI_LENH_OSB, NH_GUI, NOI_DUNG

**Logic chi tiết:**

```
BƯỚC 6.1 — Đọc & gộp 2 file:
  - Giải nén 2 zip, đọc CSV, gộp thành 1 DataFrame

BƯỚC 6.2 — Lọc SESSION:
  - Lấy SESSION = session_id (từ cả 2 file)
  - Lấy SESSION = null VÀ NGAY_GIAO_DICH = NGAY_DOI_CHIEU (11/06/2026)
  - Bỏ SESSION khác (vd: 15862, 15902)

BƯỚC 6.3 — Bỏ trạng thái loại trừ:
  - Bỏ TRANG_THAI_LENH = 'RJCT'
  - Giữ lại tất cả còn lại kể cả TRANG_THAI_LENH null/rỗng

BƯỚC 6.4 — Xử lý TRACE:
  - Strip dấu nháy đơn đầu chuỗi nếu có
  - MIS_ĐẾN không có SE_TRACE → chỉ dùng cột TRACE

BƯỚC 6.5 — Tạo khóa:
  - KEY_DEN_HUB = str(CHI_NHANH) + str(TRACE) + str(SO_TIEN)
```

**Output:** df_mis_den

---

## 10. BƯỚC 7 — b7_doi_chieu_den.py

**Mục tiêu:** Đối chiếu NPO_ĐẾN vs MIS_ĐẾN theo count.

**Input:** df_npo_den, df_mis_den

**Logic chi tiết:**

```
BƯỚC 7.1 — Đếm KEY theo count:
  - dict_npo = value_counts của KEY_DEN trong df_npo_den
  - dict_mis = value_counts của KEY_DEN_HUB trong df_mis_den

BƯỚC 7.2 — So sánh (giống Bước 5):
  - Khớp → df_mis_den_khop
  - Thừa NPO_ĐẾN → df_npo_den_thua
  - Thừa MIS_ĐẾN → df_mis_den_thua
```

**Output:** df_mis_den_khop, df_npo_den_thua, df_mis_den_thua

---

## 11. XUẤT KẾT QUẢ — trong main.py

**Output:** 1 file Excel — doi_chieu_YYYYMMDD.xlsx

**Gồm 9 sheet:**

| Sheet | Nội dung | Màu nền |
|---|---|---|
| TONG_KET | Bảng số liệu tổng hợp | Trắng |
| MIS_DI_KHOP | Giao dịch đi khớp đúng | Xanh lá |
| NPO_DI_THUA | NPO_ĐI thừa (không có ở MIS) | Đỏ |
| MIS_DI_THUA | MIS_ĐI thừa (không có ở NPO) | Đỏ |
| TIMEOUT_KHONG_KENH | TPAY thừa, không đi kênh | Cam |
| MIS_DEN_KHOP | Giao dịch đến khớp đúng | Xanh lá |
| NPO_DEN_THUA | NPO_ĐẾN thừa | Đỏ |
| MIS_DEN_THUA | MIS_ĐẾN thừa | Đỏ |

**Sheet TONG_KET gồm:**
- Ngày đối chiếu
- Session
- Chiều ĐI: số khớp / NPO thừa / MIS thừa / timeout không kênh
- Chiều ĐẾN: số khớp / NPO thừa / MIS thừa

---

## 12. main.py — FLOW CHÍNH

```python
# Cách chạy:
# python main.py --input ./input --output ./output

1. Tìm file PDF trong input/ → đọc session_id
2. Tìm file GL02*.zip → chạy b2_xu_ly_gl02 → npo_di, npo_den
3. Tìm file GW*.xlsx → chạy b3_xu_ly_gw → dict_gw_count
4. Tìm 2 file zip MIS_DI (tên chứa '_DI_') → chạy b4 → mis_di_final, timeout
5. Chạy b5 → mis_di_khop, npo_di_thua, mis_di_thua
6. Tìm 2 file zip MIS_DEN (tên chứa '_DEN_') → chạy b6 → mis_den
7. Chạy b7 → mis_den_khop, npo_den_thua, mis_den_thua
8. Xuất Excel
```

---

## 13. LƯU Ý KỸ THUẬT CHO CLAUDE CODE

```
1. Thư viện đọc ZIP có password: dùng 'pyzipper' (zipfile thường không hỗ trợ AES)
   import pyzipper
   with pyzipper.AESZipFile(zip_path) as z:
       z.setpassword(b'DACwLdHi')

2. Encoding file CSV: thử 'utf-8-sig' trước, nếu lỗi thử 'cp1252'

3. Dấu nháy đơn đầu chuỗi trong CSV (vd: '145789877):
   df['TRACE'] = df['TRACE'].astype(str).str.lstrip("'").str.strip()

4. So khớp theo COUNT (không phải merge 1-1):
   Dùng value_counts() rồi so sánh dict, không dùng pd.merge()
   Khi lấy dòng thừa: dùng groupby + tail(n_thua) để lấy đúng số dòng

5. NGAY_KENH_TRA parse datetime:
   pd.to_datetime(df['NGAY_KENH_TRA'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

6. NGAY_GIAO_DICH parse date:
   pd.to_datetime(df['NGAY_GIAO_DICH'].str.strip(), format='%d/%m/%Y', errors='coerce')

7. SO_TIEN, CRAMOUNT, DRAMOUNT: đảm bảo là số nguyên khi tạo KEY
   str(int(float(value)))

8. File GW có 5 dòng đầu là tiêu đề meta, dữ liệu bắt đầu từ dòng 6:
   pd.read_excel(path, sheet_name=..., header=5)
```

