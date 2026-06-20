# KẾ HOẠCH NÂNG CẤP — DOI-CHIEU-ACH
> Phiên bản: 20/06/2026 (v3 — cập nhật từ REVIEW-V2, benchmark độc lập)
> Mục tiêu: Tăng tốc xử lý + Thêm Web UI kéo thả folder qua mạng LAN

---

## MỤC LỤC

1. [Phân tích hiện trạng & bottleneck](#1-phân-tích-hiện-trạng--bottleneck)
2. [UPGRADE A — Tăng tốc xử lý](#2-upgrade-a--tăng-tốc-xử-lý)
3. [UPGRADE B — Web UI kéo thả folder qua LAN](#3-upgrade-b--web-ui-kéo-thả-folder-qua-lan)
4. [UPGRADE C — Sửa lỗi & chất lượng (từ REVIEW-V2)](#4-upgrade-c--sửa-lỗi--chất-lượng-từ-review-v2)
5. [Kiến trúc tổng thể sau nâng cấp](#5-kiến-trúc-tổng-thể-sau-nâng-cấp)
6. [Thứ tự triển khai & ước lượng công việc](#6-thứ-tự-triển-khai--ước-lượng-công-việc)
7. [Chi tiết code cần thêm / sửa](#7-chi-tiết-code-cần-thêm--sửa)
8. [Kiểm thử & đo lường kết quả](#8-kiểm-thử--đo-lường-kết-quả)

---

## 1. Phân tích hiện trạng & bottleneck

### Số liệu thực tế
| Chỉ tiêu | Giá trị |
|---|---|
| Thời gian xử lý 1 ngày | ~10.7 phút |
| Dữ liệu thực tế | ~2.4 triệu dòng |
| Mục tiêu sau nâng cấp | < 6 phút |

### Kiến trúc luồng xử lý hiện tại (main.py)

```
main() hiện tại:
  B1  doc_session              → Đọc PDF (glob KHÔNG đệ quy — lỗi tiềm ẩn)   ⚠️
  [Phase 1 — ThreadPool ×3]
    B2  xu_ly_gl02             → Đọc 1 ZIP AES + TẤT CẢ cột (thừa)            ⚠️
    B3  xu_ly_gw               → Đọc Excel calamine, lọc session                ✅
    B6  xu_ly_mis_den          → Đọc 2 ZIP song song (đã A1)                    ✅
  [Hết Phase 1]
  B4  xu_ly_mis_di             → Đọc 2 ZIP song song (đã A1)                    ✅
  [Phase 2 — ThreadPool ×2]
    B5  doi_chieu_di           → groupby + cumcount (vectorized)                 ✅
    B7  doi_chieu_den          → groupby + cumcount (vectorized)                 ✅
  [Hết Phase 2]
  xuat_excel                   → Sheet lớn → CSV (34x); sheet nhỏ → Excel       ✅
```

### Bottleneck & vấn đề (xếp theo ưu tiên — v3)

| # | Vấn đề | Mức độ | Trạng thái |
|---|---|---|---|
| 🔴 1 | **Web UI log bị treo/mất kết nối** — `eventlet` thiếu `monkey_patch()` + `ThreadPoolExecutor` CPU-bound block event loop | Nghiêm trọng | ❌ Cần sửa |
| 🔴 2 | **Ghi 2 sheet lớn MIS_DI_KHOP + MIS_DEN_KHOP** — ~85s/500k dòng khi dùng Excel | Tốc độ lớn nhất | ✅ Đã xong (A_NEW) |
| 🟡 3 | **B2 đọc thừa cột** — không có `usecols`, đọc toàn bộ cột GL02 thay vì chỉ 17 cột cần | Tốc độ vừa | ❌ Cần sửa |
| 🟡 4 | **A1 ThreadPool đọc 2 ZIP** — tỷ lệ I/O:CPU thực = 6%:94%; trên máy ít core có thể phản tác dụng (benchmark 0.66x) | Phụ thuộc phần cứng | ⚠️ Cần đo trên máy thật |
| 🟡 5 | **B1 tìm PDF không đệ quy** — nếu PDF trong subfolder thì crash, trong khi các file khác được tìm đệ quy | Lỗi tiềm ẩn | ❌ Cần sửa |
| 🟡 6 | **CSV mở bằng Excel có thể mất leading-zero** — TRACE/MSGSEQ có nguy cơ hiển thị sai | Rủi ro dữ liệu | ❌ Cần ghi chú |
| 🟢 7 | **ZIP_PASSWORD plaintext trong config.py** — được commit vào Git | Bảo mật | ❌ Nên sửa |
| 🟢 8 | **4 file .md chồng chéo ở gốc dự án** | Dọn dẹp | ❌ Nên làm |

> **Đã loại khỏi bottleneck list (benchmark bác bỏ):**
> - `write_row` vs `write_column`: benchmark 0% khác biệt → giữ `constant_memory=True` + `write_row`
> - `engine='c'`: pandas đã mặc định dùng C engine → 0.02s khác biệt = nhiễu đo lường

---

## 2. UPGRADE A — Tăng tốc xử lý

### A1 — Song song hoá đọc 2 ZIP bên trong B4 và B6

**Trạng thái: ✅ ĐÃ TRIỂN KHAI — nhưng cần xác minh hiệu quả trên máy chủ thật**

**File đã sửa:** `modules/b4_xu_ly_mis_di.py`, `modules/b6_xu_ly_mis_den.py`

```python
with ThreadPoolExecutor(max_workers=2) as ex:
    futures = [ex.submit(_doc_zip, p) for p in zip_paths]
    frames  = [f.result() for f in futures]
```

> **⚠️ Cảnh báo từ REVIEW-V2 (benchmark độc lập):**
>
> Tỷ lệ thực tế khi xử lý ZIP: **I/O ~6% : CPU ~94%** (đọc đĩa + giải mã AES rất nhỏ so với parse CSV).
>
> ```
> [Benchmark độc lập — 1 CPU core]
> Tuần tự đọc+giải nén+parse 2 ZIP:  4.47s
> Song song (ThreadPool) 2 ZIP:       4.21s  → 1.06x (gần bằng)
> Parse CSV thuần (data đã trong RAM): tuần tự 7.42s vs song song 11.18s → 0.66x (CHẬM HƠN)
> ```
>
> Kết quả phụ thuộc **số CPU core thật của máy chủ**:
> - **≤2 core**: A1 có thể không có lợi hoặc phản tác dụng (GIL + context-switch overhead). Cân nhắc revert về sequential.
> - **≥4 core**: Có khả năng có lợi, nhưng cần đo thực tế.
>
> **Bắt buộc phải đo trên máy chủ thật trước khi tin vào ước tính 10–30%.**
>
> ```python
> # Thêm vào đầu main_from_dir() để đo khi chạy lần đầu:
> import os
> print(f'[INFO] So CPU core: {os.cpu_count()}')
> ```

---

### A2 — ĐÃ LOẠI BỎ (write_column vs write_row — benchmark 0% cải thiện)

~~Tối ưu `_viet_sheet()` bỏ `constant_memory`, dùng `write_column`~~

**Benchmark độc lập (REVIEW-V2) xác nhận lại:**
```
200k rows × 18 cols:
  write_row + constant_memory:    21.8s
  write_column, no constant_mem:  22.8s  → write_row vẫn nhanh hơn (1.05x)
```

**→ Giữ nguyên `constant_memory=True` và `write_row` như hiện tại.**

---

### A3 — ĐÃ LOẠI BỎ (engine='c' — benchmark 0% + on_bad_lines nguy hiểm)

pandas đã mặc định dùng C engine. `on_bad_lines='skip'` nguy hiểm cho dữ liệu ngân hàng.

---

### A_NEW — CSV output cho sheet lớn ✅ ĐÃ TRIỂN KHAI

**Benchmark độc lập (REVIEW-V2) xác nhận:**
```
500k dòng × 18 cột:
  Ghi Excel (write_row + border + constant_memory):  60.5s
  Ghi CSV (pandas to_csv, utf-8-sig):                 2.2s
  Speedup đo được: 27.6x  (plan V2 báo 34x — chênh do khác máy, kết luận giữ nguyên)
```

`CSV_THRESHOLD = 50_000` — đã có trong `main.py`. Chỉ áp dụng cho `MIS_DI_KHOP` và `MIS_DEN_KHOP`.

> **Lưu ý bỏ border không nên áp dụng:** REVIEW-V2 benchmark thêm: bỏ border nhanh hơn 1.33x (22.0s → 16.6s) nhưng các sheet nhỏ vốn đã nhanh (<50k dòng → <3s) — không đáng đánh đổi thẩm mỹ.

---

### A_CLEAN — `_clean()` warning thiếu cột ✅ ĐÃ TRIỂN KHAI

```python
def _clean(df, cols, label=''):
    existing = [c for c in cols if c in df.columns]
    missing  = [c for c in cols if c not in df.columns]
    if missing:
        print(f'[WARN] _clean({label}): thieu cot {missing}')
    return df[existing].copy()
```

---

### A5 — tqdm progress bar ✅ ĐÃ TRIỂN KHAI

---

### A_B2 — Thêm `usecols` đúng 17 cột cho B2 + gộp định nghĩa cột về `config.py` ❌ CẦN LÀM

**File cần sửa:** `config.py`, `modules/b2_xu_ly_gl02.py`, `main.py`

**Vấn đề:** `b2_xu_ly_gl02.py` đọc TOÀN BỘ cột CSV của GL02 (không có `usecols`). Đây là hậu quả của một lần sửa bug cũ (usecols=4 cột → thiếu 13 cột trong sheet NPO_*_THUA) được sửa bằng cách bỏ hẳn usecols thay vì sửa đúng danh sách.

**Benchmark (REVIEW-V2):**
```
600k dòng, GL02 gốc có 30 cột (17 cần + 13 thừa):
  Đọc TẤT CẢ 30 cột:         1.97s
  Đọc CHỈ 17 cột (usecols):  1.32s  → 1.49x nhanh hơn
```
Mức tiết kiệm thực tế phụ thuộc số cột thật của GL02 gốc (chưa biết chính xác).

**Giải pháp — định nghĩa một lần ở `config.py` để tránh lặp lại bug cũ (2 nơi định nghĩa cùng 1 danh sách cột là nguyên nhân gốc của bug trước):**

```python
# config.py — THÊM:
COLS_NPO = [
    'TRDATE', 'TRBRCD', 'USERID', 'JOURSEQ', 'DYTRSEQ', 'LOCAC', 'CCY',
    'BUSCD', 'UNIT', 'TRCD', 'CUSTOMER', 'TRTP', 'REFERENCE',
    'REMARK', 'DRAMOUNT', 'CRAMOUNT', 'CRTDTM',
]
```

```python
# modules/b2_xu_ly_gl02.py — THÊM usecols:
from config import ZIP_PASSWORD, COLS_NPO as _COLS_NPO

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
                            usecols=lambda c: c in _COLS_NPO,   # ← THÊM
                            encoding=enc,
                            low_memory=False,
                        )
                        frames.append(df)
                        break
                    except UnicodeDecodeError:
                        continue
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_COLS_NPO)
```

```python
# main.py — THAY _COLS_NPO local bằng import từ config:
from config import ..., COLS_NPO as _COLS_NPO_CFG
# Hoặc đơn giản hơn:
import config
_COLS_NPO = config.COLS_NPO
```

> **Rủi ro thấp:** Chỉ khác với sửa bug cũ ở chỗ usecols bây giờ dùng đủ 17 cột (không phải 4 cột như trước gây bug). Cần kiểm tra sau khi sửa rằng sheet NPO_DI_THUA và NPO_DEN_THUA vẫn đủ cột.

---

### Tổng hợp lợi ích tốc độ

| # | Giải pháp | Ước lượng tiết kiệm | Trạng thái |
|---|---|---|---|
| A1 | ThreadPool ZIP B4+B6 | ~0–20% (cần đo máy thật) | ✅ Triển khai, ⚠️ cần xác minh |
| ~~A2~~ | ~~write_column~~ | ~~0~~ (bác bỏ) | ✅ Đã loại |
| ~~A3~~ | ~~engine='c'~~ | ~~0~~ (bác bỏ) | ✅ Đã loại |
| **A_NEW** | CSV cho MIS_DI_KHOP + MIS_DEN_KHOP | **~4–5 phút (27x benchmark)** | ✅ Xong |
| A_CLEAN | `_clean()` warning | 0 (chất lượng) | ✅ Xong |
| A5 | tqdm progress bar | 0 (UX) | ✅ Xong |
| **A_B2** | `usecols` đúng 17 cột cho B2 | **~1.5x đọc GL02** | ❌ Cần làm |

---

## 3. UPGRADE B — Web UI kéo thả folder qua LAN

### Mô hình triển khai

```
[Máy chủ — cài sẵn Python + ứng dụng]
        ↑
   HTTP LAN (port 8080)
        ↑
[Máy người dùng — trình duyệt Chrome/Edge]
  → Kéo thả folder input
  → Xem tiến trình real-time
  → Tải file output .xlsx (+ CSV nếu sheet lớn)
```

### Kiến trúc Web UI

**Thư viện backend:** `Flask` + `Flask-SocketIO` (real-time progress)  
**Frontend:** HTML5 + Drag-and-drop API + `webkitdirectory`  
**Tương thích trình duyệt:** Chrome/Edge hỗ trợ đầy đủ. Firefox KHÔNG hỗ trợ `webkitdirectory` trong drag-and-drop.

---

### Web-B1 — Cài đặt thư viện ✅ ĐÃ LÀM

```
pip install flask flask-socketio tqdm
```

`requirements.txt` đã cập nhật: flask>=3.0, flask-socketio>=5.3, tqdm.

> **REVIEW-V2:** Bỏ `eventlet>=0.35` khỏi requirements.txt (xem Web-B3 bên dưới).

---

### Web-B2 — Cấu trúc thư mục ✅ ĐÃ TẠO

```
DOI-CHIEU-ACH/
├── main.py          (sửa: A_NEW, A_CLEAN, main_from_dir)
├── web_app.py       ← ĐÃ TẠO: Flask server
├── config.py        (sửa: thêm COLS_NPO, ZIP_PASSWORD env)
├── modules/
│   ├── b1_doc_session.py   (sửa: recursive glob PDF)
│   ├── b2_xu_ly_gl02.py    (sửa: thêm usecols COLS_NPO)
│   ├── b4_xu_ly_mis_di.py  (sửa: A1 + thread-safety tpay)
│   └── b6_xu_ly_mis_den.py (sửa: A1)
├── templates/
│   └── index.html   ← ĐÃ TẠO
├── uploads/         ← ĐÃ TẠO
├── output/
├── docs/            ← CẦN TẠO: gom file .md
└── START_WEB.bat    ← ĐÃ TẠO
```

---

### Web-B3 — `web_app.py` ✅ ĐÃ TẠO — ⚠️ CẦN SỬA `eventlet` → `threading`

**Vấn đề nghiêm trọng từ REVIEW-V2 (ưu tiên #1):**

`web_app.py` hiện dùng `async_mode='eventlet'` nhưng **KHÔNG gọi `eventlet.monkey_patch()`**. Khi `main_from_dir()` chạy `ThreadPoolExecutor` (CPU-bound, pandas), eventlet không nhường event loop → **toàn bộ SocketIO emit bị chặn cứng** trong suốt quá trình xử lý. Hậu quả:
- Log `socketio.emit('log', ...)` bị dồn cục, phát ra hàng loạt khi xử lý xong thay vì real-time
- Client bị ngắt WebSocket dù đã đặt `ping_timeout=300` (server không trả lời ping được)
- Người dùng tưởng bị lỗi, refresh trang → mất tham chiếu job_id dù job vẫn chạy

**Sửa theo REVIEW-V2 (Hướng A — đơn giản, khuyến nghị cho công cụ LAN nội bộ):**

```python
# web_app.py — SỬA:
socketio = SocketIO(
    app,
    async_mode='threading',     # ← bỏ 'eventlet', dùng 'threading'
    cors_allowed_origins='*',
    ping_timeout=300,
    ping_interval=25,
)
```

```
# requirements.txt — XOÁ dòng:
# eventlet>=0.35
```

> `async_mode='threading'` chạy mỗi job trên 1 OS thread thật (giống cách `web_app.py` đã tự quản lý bằng `threading.Thread`). Không có monkey-patch hay greenlet — `ThreadPoolExecutor` bên trong `main_from_dir()` hoạt động bình thường.
>
> Nhược điểm duy nhất: `socketio.on('disconnect')` có độ trễ vài chục giây — không quan trọng với use-case nội bộ LAN.

---

### Web-B4 — Vá thread-safety B4 (tpay tham số hoá) ✅ ĐÃ TRIỂN KHAI

`xu_ly_mis_di()` đã nhận `tpay_tu`/`tpay_den` làm tham số — không import từ config tại module load.

---

### Web-B5 — `main_from_dir()` + refactor `main()` ✅ ĐÃ TRIỂN KHAI

---

### Web-B6 — `templates/index.html` ✅ ĐÃ TẠO

---

### Web-B7 — `START_WEB.bat` + `requirements.txt` ✅ ĐÃ TẠO

---

## 4. UPGRADE C — Sửa lỗi & chất lượng (từ REVIEW-V2)

### C1 — Sửa `eventlet` → `threading` trong `web_app.py` ❌ CẦN LÀM (ƯU TIÊN #1)

Xem chi tiết tại Web-B3 ở trên.

**File cần sửa:** `web_app.py`, `requirements.txt`

```python
# web_app.py — thay:
# async_mode='eventlet'
# bằng:
async_mode='threading'
```

```
# requirements.txt — xoá dòng:
# eventlet>=0.35
```

---

### C2 — Sửa `b1_doc_session.py` tìm PDF đệ quy ❌ CẦN LÀM

**Vấn đề:** `b1_doc_session.py` dùng `glob.glob(pattern)` không đệ quy, trong khi `main.py._tim_file()` tìm đệ quy cho tất cả file khác. Nếu PDF nằm trong subfolder (rất dễ xảy ra khi upload qua Web UI với `webkitdirectory`), B1 sẽ raise lỗi ngay bước đầu.

**File cần sửa:** `modules/b1_doc_session.py`

```python
# Thay:
pattern = os.path.join(input_dir, '*.pdf')
pdfs = glob.glob(pattern)

# Thành:
pattern = os.path.join(os.path.abspath(input_dir), '**', '*.pdf')
pdfs = sorted(glob.glob(pattern, recursive=True))
```

---

### C3 — Thêm `usecols` cho B2 + gộp `_COLS_NPO` về `config.py` ❌ CẦN LÀM

Xem chi tiết tại A_B2 ở trên.

---

### C4 — Ghi chú mở CSV đúng cách trong sheet Excel ❌ CẦN LÀM

**Vấn đề:** Khi người dùng double-click file CSV trực tiếp từ Excel (thói quen phổ biến), Excel tự suy luận kiểu dữ liệu → mất leading-zero của `TRACE`, `MSGSEQ`, `REFHUB` (vd: `000123456` hiển thị thành `123456`) hoặc số lớn hiển thị dạng khoa học (`9.99999E+16`).

Lưu ý: Dữ liệu GỐC trong file CSV vẫn đúng — chỉ là hiển thị sai khi mở trực tiếp.

**File cần sửa:** `main.py` — trong `xuat_excel()`, phần tạo note sheet cho CSV:

```python
# Sửa đoạn ghi note trong Excel khi sheet lớn → CSV:
ws.write(0, 0, f'[Du lieu lon - xem file: {os.path.basename(csv_path)}]')
ws.write(1, 0, f'Tong so dong: {len(df):,}')
# THÊM 2 dòng cảnh báo:
ws.write(2, 0, 'LUU Y: Mo file CSV qua Excel > Data > Tu Van ban/CSV')
ws.write(3, 0, 'KHONG double-click truc tiep - tranh mat so 0 dau va sai dinh dang so.')
```

---

### C5 — Chuyển `ZIP_PASSWORD` sang biến môi trường ❌ NÊN LÀM

**Vấn đề:** `ZIP_PASSWORD = b'DACwLdHi'` trong `config.py` đang được commit thẳng vào Git — mật khẩu giải mã file ZIP dữ liệu ngân hàng nằm trong lịch sử Git vĩnh viễn.

**File cần sửa:** `config.py`

```python
# config.py — thay:
# ZIP_PASSWORD = b'DACwLdHi'
# bằng:
import os
ZIP_PASSWORD = os.environ.get('DOI_CHIEU_ZIP_PASSWORD', 'DACwLdHi').encode()
```

> Không phá vỡ CLI hiện tại (có giá trị mặc định fallback). Để thay đổi mật khẩu sau này: đặt biến môi trường `DOI_CHIEU_ZIP_PASSWORD` trước khi chạy, không cần sửa code.

---

### C6 — Dọn file `.md` vào thư mục `docs/` ❌ NÊN LÀM

**Hiện tại:** 4–5 file `.md` kế hoạch/review chồng chéo ở gốc dự án.

**Đề xuất cấu trúc:**
```
docs/
├── 00-KIEN-TRUC-GOC.md              (từ KE_HOACH_CODE.md)
├── 01-fix-bug-toc-do-v1.md           (từ PLAN_DOI_CHIEU_ACH.md)
├── 02-web-ui-toc-do-v2.md            (từ PLAN-KEO_THA_DOI_CHIEU_ACH.md — file này)
├── 03-review-v1.md                   (từ REVIEW-V1-PLAN-KEO_THA_DOI_CHIEU_ACH.md)
└── 04-review-v2.md                   (từ REVIEW-V2-PLAN-KEO_THA_DOI_CHIEU_ACH.md)
```

Không ảnh hưởng tới vận hành (`.bat` không tham chiếu file `.md`).

---

## 5. Kiến trúc tổng thể sau nâng cấp

```
┌─────────────────────────────────────────────────────────┐
│                      MÁY CHỦ                           │
│                                                         │
│  START_WEB.bat                                          │
│       │                                                 │
│  web_app.py (Flask + SocketIO, port 8080)               │
│       │   async_mode='threading' ← SỬA từ eventlet      │
│       │   ping_timeout=300, ping_interval=25            │
│       ├── /upload    → lưu file vào uploads/job_id/     │
│       ├── /download  → trả file output về client        │
│       └── SocketIO   → emit log real-time               │
│                │                                        │
│         main_from_dir()  ← thread-safe, ngày local      │
│                │                                        │
│    ┌───────────┼───────────┐                            │
│   B2          B3          B6     (ThreadPool × 3)       │
│   (17 cột)    │    (2 ZIP song song bên trong)          │
│    └───────────┼───────────┘                            │
│     B4 (2 ZIP song song, tpay tham số hoá)              │
│    ┌───────────┴───────────┐                            │
│   B5                      B7     (ThreadPool × 2)       │
│    └───────────┬───────────┘                            │
│           xuat_excel()                                   │
│             ├── MIS_DI_KHOP > 50k → CSV + ghi chú      │
│             ├── MIS_DEN_KHOP > 50k → CSV + ghi chú     │
│             └── Các sheet nhỏ → Excel (constant_memory) │
│                │                                        │
│      output/doi_chieu_YYYYMMDD.xlsx                     │
│      output/MIS_DI_KHOP_YYYYMMDD.csv  (nếu > 50k)      │
│      output/MIS_DEN_KHOP_YYYYMMDD.csv (nếu > 50k)      │
└─────────────────────────────────────────────────────────┘
         ▲ HTTP LAN (port 8080)
┌──────────────────────────┐
│  MÁY NGƯỜI DÙNG          │
│  Chrome / Edge           │
│  → Kéo thả folder input  │
│  → Xem log real-time     │
│  → Click tải .xlsx + CSV │
└──────────────────────────┘
```

---

## 6. Thứ tự triển khai & ước lượng công việc

> **v3 — sắp xếp lại theo REVIEW-V2: ưu tiên "kết quả không mong muốn" trước "tốc độ"**

| Bước | Nội dung | Ưu tiên | Ước công |
|---|---|---|---|
| **1** | **C1** — Bỏ `eventlet`, chuyển `async_mode='threading'` | 🔴 Khắc phục log treo/mất kết nối | 15 phút |
| **2** | **C2** — Sửa B1 tìm PDF đệ quy | 🔴 Tránh crash khi folder lồng nhau | 5 phút |
| **3** | **C3 / A_B2** — Thêm `usecols` cho B2 + gộp `COLS_NPO` về `config.py` | 🟡 ~1.5x tốc độ đọc GL02, tránh bug cũ | 20 phút |
| **4** | **C4** — Thêm ghi chú CSV "mở đúng cách" trong sheet Excel | 🟡 Tránh hiểu nhầm dữ liệu | 10 phút |
| **5** | **Đo A1 trên máy thật** — `os.cpu_count()` + so sánh B4/B6 tuần tự vs song song | 🟡 Xác nhận A1 có lợi hay cần bỏ | 30 phút |
| **6** | **C5** — `ZIP_PASSWORD` → biến môi trường | 🟢 Bảo mật | 15 phút |
| **7** | **C6** — Dọn file `.md` vào `docs/` | 🟢 Gọn gàng | 10 phút |
| **8** | Test tổng thể trên dữ liệu thật (đo timer từng bước) | 🔴 Bắt buộc | 60 phút |
| | **Tổng** | | **~2.5 giờ** |

**Đã hoàn thành (không cần làm thêm):**
- A_NEW (CSV output 27x), A_CLEAN, A5 (tqdm), thread-safety B4 (tpay), main_from_dir(), web_app.py, index.html, START_WEB.bat, requirements.txt

---

## 7. Chi tiết code cần thêm / sửa

### Tóm tắt file cần đụng (còn lại)

| File | Loại | Nội dung |
|---|---|---|
| `web_app.py` | **Sửa** | C1: `async_mode='threading'`, bỏ eventlet |
| `requirements.txt` | **Sửa** | C1: xoá `eventlet>=0.35` |
| `modules/b1_doc_session.py` | **Sửa** | C2: glob đệ quy cho PDF |
| `config.py` | **Sửa** | C3: thêm `COLS_NPO = [...]`; C5: `ZIP_PASSWORD` từ env |
| `modules/b2_xu_ly_gl02.py` | **Sửa** | C3: `usecols=lambda c: c in _COLS_NPO` |
| `main.py` | **Sửa** | C3: dùng `config.COLS_NPO` thay vì local; C4: thêm ghi chú CSV |
| `docs/` | **Tạo mới** | C6: gom file .md |

**Không đổi:** `modules/b3_xu_ly_gw.py`, `modules/b5_doi_chieu_di.py`, `modules/b7_doi_chieu_den.py`, `templates/index.html`, `START_WEB.bat`

---

## 8. Kiểm thử & đo lường kết quả

### Checklist kiểm thử tổng

**Web UI:**
- [ ] **C1 (eventlet → threading)**: Kéo thả folder → log xuất hiện real-time từng dòng (không dồn cục cuối)
- [ ] **C1**: Xử lý 2.4M dòng qua LAN → không ngắt kết nối giữa chừng
- [ ] **C2 (B1 đệ quy)**: Upload folder có PDF trong subfolder → không crash
- [ ] **Web upload**: Kéo thả folder → thấy log → tải được .xlsx + CSV

**Tốc độ:**
- [ ] **A_NEW**: Log `[CSV]` xuất hiện với dữ liệu thật; số dòng CSV khớp kết quả cũ
- [ ] **A_B2 / C3**: NPO_DI_THUA và NPO_DEN_THUA vẫn đủ 17 cột sau khi thêm `usecols`
- [ ] **C3 edge case 1**: Test với file GL02 mẫu có ít hơn 17 cột → chương trình vẫn chạy đúng, không raise lỗi giả (usecols chỉ lọc cột tồn tại; chỉ 4 cột `_COLS_REQUIRED` mới bắt buộc — đã xác minh cả 4 đều nằm trong 17 cột)
- [ ] **C3 edge case 2**: Test với file GL02 sai encoding → `UnicodeDecodeError` vẫn được raise trước khi `usecols` xử lý → fallback `cp1252` hoạt động đúng như cũ
- [ ] **A1**: Đo `os.cpu_count()` trên máy chủ thật. Đo thời gian B4+B6 với A1 BẬT và TẮT → quyết định giữ hay bỏ

**Dữ liệu:**
- [ ] **C4 (CSV warning)**: Sheet Excel tóm tắt có ghi chú "mở qua Data→Text/CSV"
- [ ] **Thread-safety**: 2 job khác ngày cùng lúc → kết quả không bị trộn lẫn ngày
- [ ] **A_CLEAN**: Thử với DataFrame thiếu cột → warning `[WARN] _clean(...)` xuất hiện

### Đo tốc độ mục tiêu

| Giai đoạn | Hiện tại | Sau nâng cấp |
|---|---|---|
| Đọc GL02 (B2) | ? phút | ~1.5x nhanh hơn (A_B2) |
| Đọc B2+B3+B6 (song song) | ~3–4 phút | ~2.5–3 phút (A1, nếu có lợi) |
| Đọc B4 MIS_DI | ~3–4 phút | ~2.5–3 phút (A1, nếu có lợi) |
| Đối chiếu B5+B7 | ~30 giây | ~30 giây (không đổi) |
| Ghi Excel/CSV | ~4–6 phút | **~15–30 giây** (A_NEW khi kích hoạt) |
| **Tổng** | **~10.7 phút** | **~4–6 phút** |

> A_NEW là yếu tố quyết định thực sự. A1 cần đo trên máy thật để xác nhận.

---

## Lịch sử phiên bản

| Phiên bản | Ngày | Nội dung |
|---|---|---|
| v1 | 19/06/2026 | Phân tích ban đầu: A1, A2, A3, Web UI, thread-safety |
| v2 | 19/06/2026 | Tích hợp REVIEW-V1: loại A2/A3 (benchmark), thêm A_NEW (CSV 34x), nâng thread-safety lên bắt buộc |
| **v3** | **20/06/2026** | **Tích hợp REVIEW-V2: C1 (eventlet→threading), C2 (B1 đệ quy), C3 (B2 usecols+config.COLS_NPO), C4 (CSV warning), C5 (ZIP_PASSWORD env), C6 (docs/); làm rõ A1 phụ thuộc số core** |
| **v3.1** | **20/06/2026** | **Tích hợp REVIEW-V3: xác nhận C1–C6 đúng kỹ thuật; thêm 2 edge case C3 vào checklist; ghi nhận rủi ro collision file khi 2 web job cùng ngày (low priority, không sửa ngay)** |

> **Ghi chú REVIEW-V3 (collision risk):** Nếu 2 người dùng Web UI xử lý cùng ngày đối chiếu cùng lúc, file output `doi_chieu_YYYYMMDD.xlsx` và `MIS_*_KHOP_YYYYMMDD.csv` sẽ ghi đè lẫn nhau (tên file đặt theo ngày, không có job_id). Rủi ro thấp với vài người dùng nội bộ — không cần sửa ngay, chỉ biết để cân nhắc nếu mở rộng số người dùng đồng thời.

---

*Kế hoạch lập bởi Claude Sonnet 4.6 — v3 cập nhật từ benchmark độc lập REVIEW-V2 (20/06/2026)*
