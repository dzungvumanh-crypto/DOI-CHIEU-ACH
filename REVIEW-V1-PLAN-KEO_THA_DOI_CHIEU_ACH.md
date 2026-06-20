# REVIEW & ĐỀ XUẤT TỐI ƯU — DOI-CHIEU-ACH
> Người review: Claude Sonnet 4.6  
> Ngày: 19/06/2026  
> Phương pháp: Đọc trực tiếp codebase + chạy benchmark thực tế (xlsxwriter, pandas, ThreadPool)

---

## TÓM TẮT NHANH

| Mục | Kết luận |
|---|---|
| Kế hoạch Claude Code nhìn chung | ✅ Tốt — phân tích đúng bottleneck, loại đúng A4 |
| **A1** — ThreadPool đọc 2 ZIP | ✅ Giữ — có lợi ở I/O disk, nhưng **hạ kỳ vọng xuống 10–20%** |
| **A2** — `write_column` thay `write_row` | ❌ **BÁC BỎ** — benchmark cho thấy **speedup 1.0x** (không cải thiện) |
| **A3** — `engine='c'` tường minh | ❌ **BÁC BỎ** — pandas đã dùng C parser mặc định, không có khác biệt |
| **A5** — tqdm progress bar | ✅ Giữ (UX, không ảnh hưởng tốc độ) |
| **Giải pháp mới — A_NEW** | 🔴 **ĐỀ XUẤT THÊM**: Output sheet lớn ra CSV — tiết kiệm **4–5 phút** thực tế |
| **Web UI (UPGRADE B)** | ✅ Thiết kế đúng — có 1 lỗ hổng thread-safety cần vá |

---

## 1. PHÂN TÍCH CHI TIẾT TỪNG ĐIỂM TRONG KẾ HOẠCH

### A1 — ThreadPool đọc 2 ZIP bên trong B4 và B6

**Đánh giá: ✅ ĐÚNG HƯỚNG — nhưng cần điều chỉnh kỳ vọng**

Kế hoạch đã đúng khi ghi nhận GIL giới hạn phần AES decryption. Benchmark bổ sung:

```
Sequential đọc 2x BytesIO 1.2M dòng: 6.02s
Parallel   đọc 2x BytesIO 1.2M dòng: 6.17s  (speedup 0.98x)
```

> **Khi dữ liệu đã nằm trong RAM** (BytesIO sau khi giải nén), ThreadPool không có lợi vì bottleneck là CPU (parse CSV + GIL). Lợi ích thực tế chỉ đến từ **phần đọc đĩa (I/O) chồng lên nhau** trong khi đợi giải nén ZIP.

**Kết luận sửa kỳ vọng:**
- Máy chủ dùng SSD: ~15–20% cải thiện (I/O nhanh, overlap không nhiều)
- Máy chủ dùng HDD: ~20–30% cải thiện (I/O chậm, overlap nhiều hơn)
- **Ước lượng trong kế hoạch (~1–2 phút) là hợp lý**, không cần bác bỏ

**Code đề xuất trong kế hoạch là đúng** — giữ nguyên.

---

### A2 — `write_column` thay `write_row` + bỏ `constant_memory`

**Đánh giá: ❌ BÁC BỎ — Benchmark thực tế bác bỏ hoàn toàn**

Kết quả đo trực tiếp trên môi trường này:

```
N=200,000 rows × 18 cols:
  write_row  + border + constant_memory:    34.07s
  write_column + border + no constant_mem:  34.49s
  Speedup: 0.99x (không cải thiện)

N=500,000 rows × 18 cols:
  write_row  + border:   85.5s
  write_column + border: 85.1s
  Speedup: 1.0x (không cải thiện)
```

**Nguyên nhân kỹ thuật:** Bottleneck KHÔNG phải `write_row` vs `write_column`. Cả hai đều phải duyệt qua toàn bộ dữ liệu trong Python loop — đây là **O(N × COLS) Python object overhead** không thể tránh. Việc duyệt theo hàng hay theo cột không thay đổi số lần truy cập.

**Hậu quả nếu áp dụng A2:**
- Bỏ `constant_memory=True` mà không có lợi ích tốc độ
- Tốn thêm RAM (xlsxwriter phải buffer toàn bộ workbook trong RAM thay vì stream từng row)
- Với 2.4M dòng, có thể gây OOM (Out of Memory) trên máy chủ ít RAM

**→ Giữ `constant_memory=True` và `write_row` như hiện tại.**

---

### A3 — Thêm `engine='c'` tường minh

**Đánh giá: ❌ BÁC BỎ — Không có khác biệt**

```
read_csv default (C engine ngầm):  3.12s / 1.2M dòng
read_csv engine='c' tường minh:   3.10s / 1.2M dòng
Khác biệt: 0.02s = nhiễu đo lường
```

pandas đã mặc định dùng C engine khi không có tham số đặc biệt. Thêm `engine='c'` tường minh không thay đổi gì. Thêm `on_bad_lines='skip'` có thể che giấu lỗi dữ liệu thực sự — **không nên thêm**.

---

### A5 — tqdm progress bar

**Đánh giá: ✅ Giữ — Không ảnh hưởng tốc độ, cải thiện UX đáng kể**

Đặc biệt hữu ích khi chạy CLI và không thấy tiến trình ghi Excel (vốn tốn 5+ phút).

---

### UPGRADE B — Web UI Flask + SocketIO

**Đánh giá: ✅ Thiết kế đúng — có 1 lỗ hổng thread-safety nghiêm trọng**

#### Điểm tốt
- Dùng `job_id` riêng cho mỗi upload → tránh xung đột file
- `main_from_dir()` nhận `ngay` làm tham số → không mutate config global
- SocketIO emit log real-time là đúng hướng

#### Lỗ hổng thread-safety cần vá

**`b4_xu_ly_mis_di.py` dòng 6:**
```python
from config import ZIP_PASSWORD, TPAY_TU, TPAY_DEN   # ← IMPORT LÚC MODULE LOAD
```

`TPAY_TU` và `TPAY_DEN` được import **một lần** khi module load → tất cả jobs đều dùng giá trị từ lúc khởi động server. Nếu 2 người dùng xử lý 2 ngày khác nhau cùng lúc, **B4 sẽ lọc sai ngày cho cả 2 job**.

Kế hoạch Claude Code đã **nhận diện** vấn đề này nhưng liệt kê là "bước cải tiến bổ sung". Đây cần là **bắt buộc** trước khi deploy Web UI.

**Vá lỗi:**
```python
# b4_xu_ly_mis_di.py — Thay đổi signature:
# TRƯỚC:
from config import ZIP_PASSWORD, TPAY_TU, TPAY_DEN
def xu_ly_mis_di(zip_paths, dict_gw_count, session_id):
    # ... dùng TPAY_TU, TPAY_DEN global

# SAU:
from config import ZIP_PASSWORD
from datetime import datetime

def xu_ly_mis_di(zip_paths, dict_gw_count, session_id,
                  tpay_tu: datetime = None, tpay_den: datetime = None):
    import config
    _tpay_tu = tpay_tu if tpay_tu is not None else config.TPAY_TU
    _tpay_den = tpay_den if tpay_den is not None else config.TPAY_DEN
    # ... dùng _tpay_tu, _tpay_den thay vì global
```

Và trong `main_from_dir()`:
```python
# Tính tpay_tu/tpay_den từ ngay_dt local, truyền xuống b4:
from datetime import timedelta
tpay_tu = (ngay_dt - timedelta(days=1)).replace(hour=23, minute=0, second=0)
tpay_den = ngay_dt.replace(hour=23, minute=0, second=0)
mis_di_final, df_timeout = xu_ly_mis_di(
    mis_di_files, dict_gw_count, session_id,
    tpay_tu=tpay_tu, tpay_den=tpay_den
)
```

#### Vấn đề upload 500MB qua LAN

`MAX_CONTENT_LENGTH = 500 * 1024 * 1024` là giới hạn Flask — nhưng **timeout của eventlet** mặc định có thể ngắt connection khi upload file lớn qua LAN chậm. Cần thêm:

```python
socketio = SocketIO(app, async_mode='eventlet',
                    cors_allowed_origins='*',
                    ping_timeout=300,      # ← thêm
                    ping_interval=25)      # ← thêm
```

---

## 2. GIẢI PHÁP MỚI — A_NEW: OUTPUT SHEET LỚN RA CSV

**Đây là cải tiến tốc độ LỚN NHẤT mà kế hoạch hiện tại bỏ sót.**

### Benchmark

```
500,000 dòng × 18 cột:
  Ghi vào Excel (write_row + border):  85.5s
  Ghi vào CSV (pandas to_csv):          2.5s
  Speedup: 34x nhanh hơn
```

### Phân tích sheet nào lớn

| Sheet | Kích thước thực tế | Đề xuất |
|---|---|---|
| MIS_DI_KHOP | Có thể 1–2M dòng (phần khớp ≈90%+) | **CSV** |
| MIS_DEN_KHOP | Tương tự | **CSV** |
| NPO_DI_THUA | Nhỏ (vài trăm–vài ngàn) | Excel OK |
| MIS_DI_THUA | Nhỏ | Excel OK |
| TIMEOUT_KHONG_KENH | Nhỏ | Excel OK |
| NPO_DEN_THUA | Nhỏ | Excel OK |
| MIS_DEN_THUA | Nhỏ | Excel OK |
| RAW_GW | Vài nghìn dòng | Excel OK |
| TONG_KET | <20 dòng | Excel OK |

### Ước tính tiết kiệm thực tế

Nếu MIS_DI_KHOP + MIS_DEN_KHOP mỗi sheet ~1M dòng:
- Hiện tại: 2 sheets × 85s/500k × 2 = **~340s (5.7 phút) chỉ để ghi 2 sheet này**
- Sau A_NEW: 2 CSV × 2.5s = **5 giây**
- **Tiết kiệm: ~5.5 phút**

### Cách triển khai A_NEW

Sửa hàm `xuat_excel()` trong `main.py`:

```python
# Ngưỡng dòng: sheet lớn hơn sẽ ra CSV
CSV_THRESHOLD = 50_000   # tuỳ chỉnh theo thực tế

def xuat_excel(output_path: str, session_id: str,
               df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
               df_timeout, df_mis_den_khop, df_npo_den_thua,
               df_mis_den_thua, df_gw_raw):

    output_dir  = os.path.dirname(output_path)
    ngay_str    = os.path.basename(output_path).replace('doi_chieu_', '').replace('.xlsx', '')
    df_gw_clean = df_gw_raw.drop(columns=['KEY_GW'], errors='ignore') if df_gw_raw is not None else None

    # Định nghĩa sheets
    sheets = [
        ('TONG_KET',           None,                                   '#FFFFFF'),
        ('MIS_DI_KHOP',        _clean(df_mis_di_khop,  _COLS_MIS_DI), _XANH_LA),
        ('NPO_DI_THUA',        _clean(df_npo_di_thua,  _COLS_NPO),    _DO),
        ('MIS_DI_THUA',        _clean(df_mis_di_thua,  _COLS_MIS_DI), _DO),
        ('TIMEOUT_KHONG_KENH', _clean(df_timeout,       _COLS_MIS_DI), _CAM),
        ('MIS_DEN_KHOP',       _clean(df_mis_den_khop, _COLS_MIS_DEN), _XANH_LA),
        ('NPO_DEN_THUA',       _clean(df_npo_den_thua, _COLS_NPO),    _DO),
        ('MIS_DEN_THUA',       _clean(df_mis_den_thua, _COLS_MIS_DEN), _DO),
        ('RAW_GW',             df_gw_clean,                            _XANH_LAM),
    ]

    workbook = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False,
                                                   'constant_memory': True})
    csv_files_created = []

    for sheet_name, df, color in sheets:
        # Sheet lớn → ghi CSV riêng
        if (df is not None and len(df) > CSV_THRESHOLD
                and sheet_name in ('MIS_DI_KHOP', 'MIS_DEN_KHOP')):
            csv_path = os.path.join(output_dir, f'{sheet_name}_{ngay_str}.csv')
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            csv_files_created.append((sheet_name, csv_path))
            # Để lại note trong Excel sheet
            ws = workbook.add_worksheet(sheet_name)
            ws.set_tab_color(color)
            ws.write(0, 0, f'[Du lieu lon - xem file: {os.path.basename(csv_path)}]')
            ws.write(1, 0, f'Tong so dong: {len(df):,}')
            print(f'[CSV] {sheet_name}: {len(df):,} dong → {csv_path}')
            continue

        ws = workbook.add_worksheet(sheet_name)
        ws.set_tab_color(color)
        if sheet_name == 'TONG_KET':
            _viet_tong_ket(workbook, ws, session_id,
                           len(df_mis_di_khop)  if df_mis_di_khop  is not None else 0,
                           _tong_tien(df_mis_di_khop,   'SO_TIEN'),
                           # ... các tham số còn lại giữ nguyên
                           )
        else:
            _viet_sheet(workbook, ws, df, color)

    workbook.close()

    print(f'\n[DONE] Excel: {output_path}')
    for name, path in csv_files_created:
        print(f'       CSV  : {path}  ({name})')
```

### Tuỳ chọn: Giá trị `CSV_THRESHOLD`

Khuyến nghị đặt `CSV_THRESHOLD = 50_000`. Lý do:
- Dưới 50k dòng: Excel ghi dưới 10s, chấp nhận được
- Trên 50k dòng: thời gian ghi Excel tăng tuyến tính, CSV tiết kiệm đáng kể

---

## 3. PHÂN TÍCH "KẾT QUẢ KHÔNG MONG MUỐN"

Ngoài tốc độ, codebase có 2 điểm dễ gây kết quả sai:

### Vấn đề 1: `on_bad_lines='skip'` (đề xuất trong A3 — nên từ chối)

Nếu file ZIP bị corrupt hoặc encoding lạ tạo ra dòng lỗi, `skip` sẽ **bỏ qua lặng lẽ** mà không cảnh báo. Với dữ liệu ngân hàng, mất 1 giao dịch có thể dẫn đến kết quả đối chiếu sai.

**Khuyến nghị:** Không thêm `on_bad_lines='skip'`. Giữ hành vi hiện tại (raise exception khi gặp dòng lỗi) để phát hiện sớm vấn đề dữ liệu.

### Vấn đề 2: Thread-safety của B4 khi Web UI

Đã phân tích ở mục B ở trên. Cần vá trước khi deploy.

### Vấn đề 3: `_clean()` im lặng khi thiếu cột

```python
def _clean(df, cols):
    existing = [c for c in cols if c in df.columns]
    return df[existing].copy()   # ← Không cảnh báo cột thiếu
```

Nếu dữ liệu nguồn thay đổi schema (thiếu cột), output Excel sẽ thiếu cột mà không có warning. **Khuyến nghị thêm log:**

```python
def _clean(df, cols, label=''):
    existing = [c for c in cols if c in df.columns]
    missing  = [c for c in cols if c not in df.columns]
    if missing:
        print(f'[WARN] _clean({label}): thiếu cột {missing}')
    return df[existing].copy()
```

---

## 4. BẢNG TỔNG HỢP — ĐÁNH GIÁ TỪNG ĐIỂM KẾ HOẠCH

| Mục | Kế hoạch Claude Code | Review | Hành động |
|---|---|---|---|
| A1 — ThreadPool ZIP B4+B6 | ✅ Đúng | ✅ Xác nhận | Triển khai, kỳ vọng 10–20% |
| A2 — write_column | ⚠️ Ước tính 2–4 phút | ❌ **Benchmark: 0% cải thiện** | **Bỏ A2** |
| A3 — engine='c' | ⚠️ Ước tính 0.5–1 phút | ❌ **Benchmark: 0% cải thiện** | **Bỏ A3** |
| A4 — z.open() | ❌ Đã loại | ✅ Đồng ý | Giữ loại |
| A5 — tqdm | ✅ | ✅ | Triển khai |
| Thread-safety B4 TPAY | "Bổ sung sau" | 🔴 **Bắt buộc trước Web UI** | **Nâng ưu tiên** |
| SocketIO ping timeout | Không đề cập | ⚠️ Cần thêm | Thêm vào |
| **A_NEW — CSV cho sheet lớn** | **Không có** | 🔴 **Tiết kiệm 4–5 phút** | **Thêm mới** |
| _clean() cảnh báo thiếu cột | Không đề cập | ⚠️ Chất lượng kết quả | Thêm mới |

---

## 5. KẾ HOẠCH TRIỂN KHAI ĐỀ XUẤT (SAU REVIEW)

Sắp xếp lại theo mức độ tác động thực tế:

| Bước | Nội dung | Tác động | Ước công |
|---|---|---|---|
| **1** | **A_NEW** — CSV output cho MIS_DI_KHOP + MIS_DEN_KHOP | 🔴 **~4–5 phút tiết kiệm** | 45 phút |
| **2** | **A1** — ThreadPool ZIP bên trong B4 + B6 | 🟡 ~10–20% đọc dữ liệu | 30 phút |
| **3** | **Vá thread-safety B4** — tham số hoá TPAY_TU/DEN | 🔴 Bắt buộc Web UI | 30 phút |
| **4** | **Web-B4** — `main_from_dir()` + refactor `main()` | 🔴 Cần trước Web | 45 phút |
| **5** | **Web-B3** — `web_app.py` + thêm ping timeout | 🔴 Core Web UI | 60 phút |
| **6** | **Web-B5** — `templates/index.html` | 🔴 Core Web UI | 45 phút |
| **7** | **A5** — tqdm progress bar | 🟢 UX | 15 phút |
| **8** | **_clean() warning** — log cột thiếu | 🟡 Chất lượng | 10 phút |
| **9** | **Web-B6** — `START_WEB.bat` + `requirements.txt` | 🟢 | 10 phút |
| **10** | Test tổng thể trên dữ liệu thật (đo timer từng bước) | 🔴 Bắt buộc | 60 phút |
| | **Tổng** | | **~6 giờ** |

---

## 6. MỤC TIÊU THỜI GIAN XỬ LÝ SAU NÂNG CẤP

| Giai đoạn | Hiện tại (ước tính) | Sau nâng cấp |
|---|---|---|
| Đọc dữ liệu B2+B3+B6 (song song) | ~3–4 phút | ~2.5–3 phút (A1) |
| Đọc B4 MIS_DI | ~3–4 phút | ~2.5–3 phút (A1) |
| Đối chiếu B5+B7 | ~30 giây | ~30 giây (không đổi) |
| Ghi Excel | ~4–6 phút | **~15–30 giây** (A_NEW: 2 sheet lớn ra CSV) |
| **Tổng** | **~10.7 phút** | **~5–7 phút** |

> **Ghi chú thực tế:** Ước lượng tiết kiệm phụ thuộc vào tỷ lệ dữ liệu khớp/thừa thực tế.
> Nếu MIS_DI_KHOP < 50k dòng (do lọc session khắt khe), A_NEW không kích hoạt và thời gian ghi Excel đã nhỏ.
> Cần đo timer thực tế sau bước 1 để xác định mức tiết kiệm chính xác.

---

*Review bởi Claude Sonnet 4.6 — 19/06/2026*  
*Benchmark chạy trực tiếp trong môi trường container Python 3.x với xlsxwriter 3.2.9, pandas ≥2.0*
