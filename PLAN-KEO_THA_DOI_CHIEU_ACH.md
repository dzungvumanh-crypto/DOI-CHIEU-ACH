# KẾ HOẠCH NÂNG CẤP — DOI-CHIEU-ACH
> Phiên bản: 19/06/2026 (v2 — cập nhật từ benchmark thực tế)
> Mục tiêu: Tăng tốc xử lý + Thêm Web UI kéo thả folder qua mạng LAN

---

## MỤC LỤC

1. [Phân tích hiện trạng & bottleneck](#1-phân-tích-hiện-trạng--bottleneck)
2. [UPGRADE A — Tăng tốc xử lý](#2-upgrade-a--tăng-tốc-xử-lý)
3. [UPGRADE B — Web UI kéo thả folder qua LAN](#3-upgrade-b--web-ui-kéo-thả-folder-qua-lan)
4. [Kiến trúc tổng thể sau nâng cấp](#4-kiến-trúc-tổng-thể-sau-nâng-cấp)
5. [Thứ tự triển khai & ước lượng công việc](#5-thứ-tự-triển-khai--ước-lượng-công-việc)
6. [Chi tiết code cần thêm / sửa](#6-chi-tiết-code-cần-thêm--sửa)
7. [Kiểm thử & đo lường kết quả](#7-kiểm-thử--đo-lường-kết-quả)

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
  B1  doc_session              → Nhanh (đọc tên file PDF)                ✅
  [Phase 1 — ThreadPool ×3]
    B2  xu_ly_gl02             → Đọc 1 ZIP mã hoá AES (pyzipper)         ⚠️
    B3  xu_ly_gw               → Đọc Excel calamine, lọc session          ⚠️
    B6  xu_ly_mis_den          → Đọc 2 ZIP, mỗi ZIP đọc tuần tự bên trong ⚠️
  [Hết Phase 1]
  B4  xu_ly_mis_di             → Đọc 2 ZIP, mỗi ZIP đọc tuần tự bên trong 🐢 CHẬM NHẤT
  [Phase 2 — ThreadPool ×2]
    B5  doi_chieu_di           → groupby + cumcount (vectorized)           ✅
    B7  doi_chieu_den          → groupby + cumcount (vectorized)           ✅
  [Hết Phase 2]
  xuat_excel                   → write_row từng dòng + ghi 2 sheet lớn    🐢 CHẬM
```

> **Lưu ý quan trọng:** `main()` ĐÃ có ThreadPoolExecutor cho Phase 1 (B2+B3+B6 song song)
> và Phase 2 (B5+B7 song song). Các nâng cấp A1 tập trung vào TRONG NỘI BỘ từng bước.

### Bottleneck chính (xếp theo tác động)

1. **Ghi 2 sheet lớn MIS_DI_KHOP + MIS_DEN_KHOP vào Excel** — có thể 1–2M dòng mỗi sheet,
   write_row benchmark thực tế: ~85s/500k dòng → 🐢 CHẬM NHẤT trong toàn bộ quy trình
2. **B4 đọc 2 ZIP MIS_DI tuần tự bên trong** — `frames = [_doc_zip(p) for p in zip_paths]` (b4 dòng 81)
3. **B6 đọc 2 ZIP MIS_DEN tuần tự bên trong** — cùng pattern với B4 (b6 dòng 42)

> **Đã loại khỏi bottleneck list (benchmark bác bỏ):**
> - `write_row` vs `write_column`: benchmark 0% khác biệt → giữ `constant_memory=True` + `write_row`
> - `engine='c'`: pandas đã mặc định dùng C engine → 0.02s khác biệt = nhiễu đo lường

---

## 2. UPGRADE A — Tăng tốc xử lý

### A1 — Song song hoá đọc 2 ZIP bên trong B4 và B6

**File cần sửa:** `modules/b4_xu_ly_mis_di.py`, `modules/b6_xu_ly_mis_den.py`

**Vấn đề hiện tại (b4 dòng 81, b6 dòng 42):**
```python
frames = [_doc_zip(p) for p in zip_paths]   # đọc zip1 xong mới đọc zip2
```

**Sửa thành (áp dụng cho cả b4 và b6):**
```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=2) as ex:
    futures = [ex.submit(_doc_zip, p) for p in zip_paths]
    frames  = [f.result() for f in futures]
```

> **Ghi chú GIL:** `pyzipper` AES decryption là CPU-bound, bị GIL giới hạn.
> Lợi ích chủ yếu từ phần **đọc đĩa (I/O) chồng lên nhau** khi chờ giải nén ZIP:
> - SSD: ~15–20% cải thiện  
> - HDD: ~20–30% cải thiện  
> Ước lượng tiết kiệm ~0.5–1.5 phút tổng cộng cho B4+B6.
>
> `b2_xu_ly_gl02.py` chỉ nhận 1 file ZIP duy nhất (`zip_path: str`) → A1 KHÔNG áp dụng cho b2.  
> `b3_xu_ly_gw.py` không dùng pyzipper → A1 KHÔNG áp dụng cho b3.

---

### A2 — ĐÃ LOẠI BỎ (write_column vs write_row — benchmark 0% cải thiện)

~~Tối ưu `_viet_sheet()` bỏ `constant_memory`, dùng `write_column`~~

**Benchmark thực tế đã bác bỏ hoàn toàn:**
```
200k rows × 18 cols:
  write_row  + constant_memory=True:  34.07s
  write_column + no constant_memory:  34.49s  → 0.99x (không cải thiện)

500k rows × 18 cols:
  write_row:   85.5s
  write_column: 85.1s  → 1.0x (không cải thiện)
```

**Nguyên nhân kỹ thuật:** Bottleneck là O(N×COLS) Python object overhead — duyệt theo hàng
hay theo cột đều phải truy cập cùng số lần. Bỏ `constant_memory=True` không có lợi tốc độ
nhưng tăng nguy cơ OOM với 2.4M dòng.

**→ Giữ nguyên `constant_memory=True` và `write_row` như hiện tại.**

---

### A3 — ĐÃ LOẠI BỎ (engine='c' — benchmark 0% + on_bad_lines nguy hiểm)

~~Thêm `engine='c'` và `on_bad_lines='skip'` vào `pd.read_csv`~~

**Benchmark thực tế đã bác bỏ:**
```
read_csv default (C engine ngầm):  3.12s / 1.2M dòng
read_csv engine='c' tường minh:   3.10s / 1.2M dòng
Khác biệt: 0.02s = nhiễu đo lường
```

pandas đã mặc định dùng C engine — thêm `engine='c'` không thay đổi gì.

**`on_bad_lines='skip'` KHÔNG nên thêm:** Với dữ liệu ngân hàng, bỏ qua dòng lỗi im lặng
có thể che giấu giao dịch thực sự bị hỏng → dẫn đến kết quả đối chiếu sai mà không có cảnh báo.

---

### A_NEW — CSV output cho sheet lớn (tiết kiệm 4–5 phút)

**File cần sửa:** `main.py` — module-level constant + hàm `xuat_excel()`

**Benchmark thực tế:**
```
500k dòng × 18 cột:
  Ghi Excel (write_row + border):  85.5s
  Ghi CSV (pandas to_csv):          2.5s
  Speedup: 34x nhanh hơn
```

**Phân tích sheet theo kích thước:**
| Sheet | Kích thước ước tính | Đề xuất |
|---|---|---|
| MIS_DI_KHOP | Có thể 1–2M dòng (phần khớp ≈90%+) | **→ CSV** |
| MIS_DEN_KHOP | Tương tự | **→ CSV** |
| NPO_DI_THUA | Vài trăm–vài ngàn | Excel OK |
| MIS_DI_THUA | Nhỏ | Excel OK |
| TIMEOUT_KHONG_KENH | Nhỏ | Excel OK |
| NPO_DEN_THUA | Nhỏ | Excel OK |
| MIS_DEN_THUA | Nhỏ | Excel OK |
| RAW_GW | Vài nghìn dòng | Excel OK |
| TONG_KET | <20 dòng | Excel OK |

**Ngưỡng:** `CSV_THRESHOLD = 50_000` dòng
- Dưới 50k → Excel ghi dưới ~10s, chấp nhận được
- Trên 50k → chuyển sang CSV, tiết kiệm đáng kể

**Sửa `main.py`:**
```python
# Thêm constant (đặt sau các _COLS_* constants):
CSV_THRESHOLD = 50_000

# Sửa xuat_excel():
def xuat_excel(output_path: str, session_id: str,
               df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
               df_timeout, df_mis_den_khop, df_npo_den_thua,
               df_mis_den_thua, df_gw_raw):

    output_dir  = os.path.dirname(output_path)
    ngay_str    = os.path.basename(output_path).replace('doi_chieu_', '').replace('.xlsx', '')
    df_gw_clean = df_gw_raw.drop(columns=['KEY_GW'], errors='ignore') if df_gw_raw is not None else None

    sheets = [
        ('TONG_KET',           None,                                            '#FFFFFF'),
        ('MIS_DI_KHOP',        _clean(df_mis_di_khop,  _COLS_MIS_DI, 'MIS_DI_KHOP'),  _XANH_LA),
        ('NPO_DI_THUA',        _clean(df_npo_di_thua,  _COLS_NPO,    'NPO_DI_THUA'),  _DO),
        ('MIS_DI_THUA',        _clean(df_mis_di_thua,  _COLS_MIS_DI, 'MIS_DI_THUA'),  _DO),
        ('TIMEOUT_KHONG_KENH', _clean(df_timeout,       _COLS_MIS_DI, 'TIMEOUT'),      _CAM),
        ('MIS_DEN_KHOP',       _clean(df_mis_den_khop, _COLS_MIS_DEN,'MIS_DEN_KHOP'), _XANH_LA),
        ('NPO_DEN_THUA',       _clean(df_npo_den_thua, _COLS_NPO,    'NPO_DEN_THUA'), _DO),
        ('MIS_DEN_THUA',       _clean(df_mis_den_thua, _COLS_MIS_DEN,'MIS_DEN_THUA'), _DO),
        ('RAW_GW',             df_gw_clean,                                     _XANH_LAM),
    ]

    workbook = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False,
                                                   'constant_memory': True})
    csv_files_created = []

    for sheet_name, df, color in sheets:
        # Sheet lớn → ghi CSV riêng, để note trong Excel
        if (df is not None and len(df) > CSV_THRESHOLD
                and sheet_name in ('MIS_DI_KHOP', 'MIS_DEN_KHOP')):
            csv_path = os.path.join(output_dir, f'{sheet_name}_{ngay_str}.csv')
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            csv_files_created.append((sheet_name, csv_path))
            ws = workbook.add_worksheet(sheet_name)
            ws.set_tab_color(color)
            ws.write(0, 0, f'[Du lieu lon - xem file: {os.path.basename(csv_path)}]')
            ws.write(1, 0, f'Tong so dong: {len(df):,}')
            print(f'[CSV] {sheet_name}: {len(df):,} dong -> {csv_path}')
            continue

        ws = workbook.add_worksheet(sheet_name)
        ws.set_tab_color(color)
        if sheet_name == 'TONG_KET':
            _viet_tong_ket(
                workbook, ws, session_id,
                len(df_mis_di_khop)  if df_mis_di_khop  is not None else 0,
                _tong_tien(df_mis_di_khop,   'SO_TIEN'),
                len(df_npo_di_thua)  if df_npo_di_thua  is not None else 0,
                _tong_tien(df_npo_di_thua,   'CRAMOUNT'),
                len(df_mis_di_thua)  if df_mis_di_thua  is not None else 0,
                _tong_tien(df_mis_di_thua,   'SO_TIEN'),
                len(df_timeout)      if df_timeout      is not None else 0,
                _tong_tien(df_timeout,       'SO_TIEN'),
                len(df_mis_den_khop) if df_mis_den_khop is not None else 0,
                _tong_tien(df_mis_den_khop,  'SO_TIEN'),
                len(df_npo_den_thua) if df_npo_den_thua is not None else 0,
                _tong_tien(df_npo_den_thua,  'DRAMOUNT'),
                len(df_mis_den_thua) if df_mis_den_thua is not None else 0,
                _tong_tien(df_mis_den_thua,  'SO_TIEN'),
            )
        else:
            _viet_sheet(workbook, ws, df, color)

    workbook.close()
    print(f'\n[DONE] Excel: {output_path}')
    for name, path in csv_files_created:
        print(f'       CSV  : {path}  ({name})')
```

> **Lưu ý:** `_clean()` cần cập nhật signature để nhận `label` — xem A_CLEAN bên dưới.

---

### A_CLEAN — Thêm warning vào `_clean()` khi thiếu cột

**File cần sửa:** `main.py` — hàm `_clean()` (dòng 55–60)

**Vấn đề hiện tại:** `_clean()` im lặng khi cột trong `cols` không tồn tại trong df → output Excel
thiếu cột mà không có cảnh báo nếu schema dữ liệu nguồn thay đổi.

**Sửa thành:**
```python
def _clean(df: pd.DataFrame, cols: list, label: str = '') -> pd.DataFrame:
    if df is None or len(df) == 0:
        return df
    existing = [c for c in cols if c in df.columns]
    missing  = [c for c in cols if c not in df.columns]
    if missing:
        print(f'[WARN] _clean({label}): thieu cot {missing}')
    return df[existing].copy()
```

> Các caller của `_clean()` truyền thêm tham số `label` — đã có trong snippet A_NEW ở trên.

---

### A5 — Thêm progress bar để theo dõi tiến trình (UX)

**Thêm thư viện:** `pip install tqdm`

```python
# main.py — trong xuat_excel(), bao bọc vòng lặp sheets:
from tqdm import tqdm

for sheet_name, df, color in tqdm(sheets, desc='Ghi Excel', unit='sheet'):
    ...
```

---

### Tổng hợp lợi ích

| # | Giải pháp | Ước lượng tiết kiệm | Độ phức tạp | Nguồn |
|---|---|---|---|---|
| A1 | ThreadPool ZIP B4+B6 | ~0.5–1.5 phút | Thấp | Phân tích I/O |
| ~~A2~~ | ~~write_column~~ | ~~0~~ **(bác bỏ)** | — | Benchmark thực tế |
| ~~A3~~ | ~~engine='c'~~ | ~~0~~ **(bác bỏ)** | — | Benchmark thực tế |
| **A_NEW** | CSV cho MIS_DI_KHOP + MIS_DEN_KHOP | **~4–5 phút** | Thấp | Benchmark 34x |
| A_CLEAN | `_clean()` warning | 0 (chất lượng kết quả) | Rất thấp | Review |
| A5 | tqdm progress bar | 0 (UX only) | Rất thấp | — |
| **Tổng** | | **~4.5–6.5 phút** | | |

> **Kết quả dự kiến:** Từ 10.7 phút → còn khoảng **4–6 phút** cho 2.4 triệu dòng.  
> **Lưu ý:** Nếu MIS_DI_KHOP và MIS_DEN_KHOP < 50k dòng (do lọc session khắt khe), A_NEW
> không kích hoạt và thời gian ghi Excel vốn đã nhỏ — trong trường hợp đó A1 là cải tiến chính.

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
**Tương thích trình duyệt:** Chrome/Edge hỗ trợ đầy đủ. Firefox KHÔNG hỗ trợ `webkitdirectory`
trong drag-and-drop (chỉ hỗ trợ qua click chọn file).

---

### Web-B1 — Cài đặt thư viện mới

```
pip install flask flask-socketio eventlet tqdm
```

Thêm vào `requirements.txt`:
```
flask>=3.0
flask-socketio>=5.3
eventlet>=0.35
tqdm
```

---

### Web-B2 — Cấu trúc thư mục mới

```
DOI-CHIEU-ACH/
├── main.py                    (sửa: A_NEW, A_CLEAN, Web-B5 main_from_dir)
├── web_app.py                 ← MỚI: Flask server
├── config.py
├── modules/
│   ├── b4_xu_ly_mis_di.py    (sửa: A1 + vá thread-safety TPAY)
│   ├── b6_xu_ly_mis_den.py   (sửa: A1)
│   └── ... (các module khác không đổi)
├── templates/
│   └── index.html             ← MỚI: Giao diện kéo thả
├── static/
│   └── style.css              ← MỚI (tuỳ chọn)
├── uploads/                   ← MỚI: thư mục tạm nhận file từ user
├── output/
└── START_WEB.bat              ← MỚI: khởi động web server
```

---

### Web-B3 — `web_app.py` — Flask Server chính

> ⚠️ **Thêm `ping_timeout=300` và `ping_interval=25`** vào SocketIO để tránh ngắt connection
> khi upload file lớn (hàng trăm MB) qua LAN chậm.

```python
"""
web_app.py — Web UI cho DOI-CHIEU-ACH
Chay: python web_app.py
Truy cap tu LAN: http://<IP_MAY_CHU>:8080
"""
import os
import uuid
import threading
from flask import Flask, request, jsonify, render_template, send_file
from flask_socketio import SocketIO
import config
from main import main_from_dir

app = Flask(__name__)
app.config['SECRET_KEY'] = 'doi_chieu_ach_secret'
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB

socketio = SocketIO(app, async_mode='eventlet',
                    cors_allowed_origins='*',
                    ping_timeout=300,      # ← tránh ngắt connection khi upload lớn
                    ping_interval=25)

UPLOAD_DIR = os.path.abspath('./uploads')
OUTPUT_DIR = os.path.abspath('./output')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'Khong co file nao duoc gui len'}), 400

    for f in files:
        filename = os.path.basename(f.filename)
        f.save(os.path.join(job_dir, filename))

    ngay = request.form.get('ngay_doi_chieu', '').strip()

    thread = threading.Thread(
        target=_run_processing,
        args=(job_id, job_dir, ngay),
        daemon=True
    )
    thread.start()

    return jsonify({'job_id': job_id, 'message': 'Dang xu ly...'})


def _run_processing(job_id: str, input_dir: str, ngay: str):
    def emit_log(msg: str):
        socketio.emit('log', {'job_id': job_id, 'msg': msg}, namespace='/')

    try:
        emit_log(f'[{job_id}] Bat dau xu ly...')
        output_path = main_from_dir(
            input_dir=input_dir,
            output_dir=OUTPUT_DIR,
            ngay=ngay or None,
            log_callback=emit_log
        )
        socketio.emit('done', {
            'job_id': job_id,
            'download_url': f'/download/{os.path.basename(output_path)}'
        }, namespace='/')
    except Exception as e:
        socketio.emit('error', {'job_id': job_id, 'msg': str(e)}, namespace='/')


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(path):
        return 'File khong ton tai', 404
    return send_file(path, as_attachment=True)


if __name__ == '__main__':
    import socket
    ip = socket.gethostbyname(socket.gethostname())
    print(f'\nWeb UI chay tai: http://{ip}:8080')
    print('   Tu may khac trong LAN, truy cap dia chi tren.')
    socketio.run(app, host='0.0.0.0', port=8080, debug=False)
```

---

### Web-B4 — VÁ THREAD-SAFETY B4 (BẮT BUỘC TRƯỚC KHI DEPLOY WEB UI)

**Vấn đề nghiêm trọng** — `b4_xu_ly_mis_di.py` dòng 6:
```python
from config import ZIP_PASSWORD, TPAY_TU, TPAY_DEN   # import một lần lúc module load
```

`TPAY_TU` và `TPAY_DEN` được import **một lần** khi server khởi động. Nếu 2 user gửi job
cùng lúc với ngày khác nhau, B4 của cả 2 job đều lọc theo cùng `TPAY_TU`/`TPAY_DEN` từ lúc
khởi động → **lọc sai ngày, kết quả đối chiếu sai**.

**Vá `b4_xu_ly_mis_di.py`:**
```python
# Thay:
from config import ZIP_PASSWORD, TPAY_TU, TPAY_DEN

# Thành:
from config import ZIP_PASSWORD

# Thay signature của xu_ly_mis_di:
def xu_ly_mis_di(zip_paths, dict_gw_count, session_id,
                  tpay_tu=None, tpay_den=None):
    import config
    _tpay_tu  = tpay_tu  if tpay_tu  is not None else config.TPAY_TU
    _tpay_den = tpay_den if tpay_den is not None else config.TPAY_DEN
    # ... dùng _tpay_tu, _tpay_den thay TPAY_TU, TPAY_DEN ở mọi chỗ trong hàm
```

**Trong `main_from_dir()` của main.py — truyền tpay xuống b4:**
```python
from datetime import timedelta
tpay_tu  = (ngay_dt - timedelta(days=1)).replace(hour=23, minute=0, second=0)
tpay_den = ngay_dt.replace(hour=23, minute=0, second=0)

mis_di_final, df_timeout = xu_ly_mis_di(
    mis_di_files, dict_gw_count, session_id,
    tpay_tu=tpay_tu, tpay_den=tpay_den
)
```

**Trong `main()` CLI — không truyền → dùng config default (backward compatible):**
```python
mis_di_final, df_timeout = xu_ly_mis_di(mis_di_files, dict_gw_count, session_id)
```

---

### Web-B5 — Refactor `main.py`: thêm `main_from_dir()`, refactor `main()` gọi lại

**Chiến lược:**
- Tách logic xử lý ra `main_from_dir()`, sau đó `main()` CLI chỉ parse args rồi gọi lại — tránh trùng lặp code
- Ngày đối chiếu truyền qua **tham số hàm**, KHÔNG mutate `config` module-level → thread-safe khi nhiều job web chạy đồng thời với ngày khác nhau

```python
# main.py — THÊM hàm mới (đặt trước hàm main()):

def main_from_dir(input_dir: str, output_dir: str,
                  ngay: str = None, log_callback=None) -> str:
    """
    Phiên bản của main() dùng cho Web UI.
    - input_dir: thư mục đã có file upload
    - output_dir: nơi lưu file kết quả
    - ngay: 'dd/mm/yyyy' hoặc None (lấy từ config)
    - log_callback: hàm(msg: str) để emit log real-time
    Trả về: đường dẫn file output .xlsx
    Thread-safe: ngày đối chiếu tính local, KHÔNG sửa config module-level.
    """
    from datetime import datetime, timedelta

    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    # Tính ngày local thay vì mutate config global
    if ngay:
        ngay_dt      = datetime.strptime(ngay.strip(), '%d/%m/%Y')
        ngay_str_cfg = ngay.strip()
    else:
        ngay_dt      = config.NGAY_DT
        ngay_str_cfg = config.NGAY_DOI_CHIEU

    # Tính tpay để truyền tường minh xuống b4 (thread-safe)
    tpay_tu  = (ngay_dt - timedelta(days=1)).replace(hour=23, minute=0, second=0)
    tpay_den = ngay_dt.replace(hour=23, minute=0, second=0)

    os.makedirs(output_dir, exist_ok=True)
    log(f'Ngay doi chieu: {ngay_str_cfg}')

    session_id    = doc_session(input_dir)
    gl02_files    = _tim_file(input_dir, 'GL02*.zip')
    gw_path       = _tim_gw_xlsx(input_dir)
    mis_di_files  = _tim_file(input_dir, '*_DI_*.zip')
    mis_den_files = _tim_file(input_dir, '*_DEN_*.zip')

    if not gl02_files:
        raise FileNotFoundError('Khong tim thay GL02*.zip')
    if len(mis_di_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DI zip, chi tim thay {len(mis_di_files)}')
    if len(mis_den_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DEN zip, chi tim thay {len(mis_den_files)}')

    log(f'Tim thay: GL02={len(gl02_files)}, DI={len(mis_di_files)}, DEN={len(mis_den_files)}')

    with ThreadPoolExecutor(max_workers=3) as ex:
        f_gl02    = ex.submit(xu_ly_gl02,    gl02_files[0])
        f_gw      = ex.submit(xu_ly_gw,      gw_path, session_id)
        f_mis_den = ex.submit(xu_ly_mis_den, mis_den_files, session_id, ngay_dt)
        npo_di, npo_den          = f_gl02.result()
        dict_gw_count, df_gw_raw = f_gw.result()
        df_mis_den               = f_mis_den.result()

    mis_di_final, df_timeout = xu_ly_mis_di(
        mis_di_files, dict_gw_count, session_id,
        tpay_tu=tpay_tu, tpay_den=tpay_den
    )

    with ThreadPoolExecutor(max_workers=2) as ex:
        f_di  = ex.submit(doi_chieu_di,  npo_di,  mis_di_final)
        f_den = ex.submit(doi_chieu_den, npo_den, df_mis_den)
        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua    = f_di.result()
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua = f_den.result()

    output_path = os.path.join(output_dir, f'doi_chieu_{ngay_dt.strftime("%Y%m%d")}.xlsx')
    xuat_excel(
        output_path, session_id,
        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua,
        df_timeout,
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua,
        df_gw_raw,
    )
    log(f'Hoan thanh: {output_path}')
    return output_path


# main.py — REFACTOR main() để gọi main_from_dir():
def main():
    parser = argparse.ArgumentParser(description='Doi chieu ACH GL02 vs MIS')
    parser.add_argument('--input',  default=config.INPUT_DIR)
    parser.add_argument('--output', default=config.OUTPUT_DIR)
    parser.add_argument('--date',   default=None, help='dd/mm/yyyy')
    args = parser.parse_args()

    print('=' * 60)
    print(f'DOI CHIEU ACH  —  Ngay: {args.date or config.NGAY_DOI_CHIEU}')
    print('=' * 60)

    main_from_dir(
        input_dir=args.input,
        output_dir=args.output,
        ngay=args.date,
    )
```

---

### Web-B6 — `templates/index.html` — Giao diện kéo thả

```html
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Doi Chieu ACH</title>
<script src="https://cdn.socket.io/4.7.5/socket.io.min.js"></script>
<style>
  body { font-family: Arial, sans-serif; max-width: 720px; margin: 40px auto; }
  h1   { color: #1a56a0; }
  #dropzone {
    border: 3px dashed #1a56a0; border-radius: 12px;
    padding: 60px 20px; text-align: center; cursor: pointer;
    background: #f0f5ff; transition: background 0.2s;
  }
  #dropzone.drag-over { background: #c8daff; }
  #browser-note { color: #888; font-size: 12px; margin-top: 6px; }
  #log  { background:#1e1e1e; color:#d4d4d4; padding:16px;
          border-radius:8px; height:260px; overflow-y:auto;
          font-family:monospace; font-size:13px; margin-top:20px; }
  #btn-download { display:none; margin-top:16px; padding:10px 24px;
                  background:#1a56a0; color:#fff; border:none;
                  border-radius:6px; font-size:15px; cursor:pointer; }
  input[type=text] { padding:6px 10px; border:1px solid #ccc;
                     border-radius:4px; width:160px; }
</style>
</head>
<body>
<h1>Doi Chieu ACH — GL02 vs MIS</h1>

<p>Ngay doi chieu (dd/mm/yyyy):
  <input type="text" id="ngay" placeholder="11/06/2026">
  <small>(de trong = lay tu config)</small>
</p>

<div id="dropzone" onclick="document.getElementById('fileInput').click()">
  Keo tha <strong>folder du lieu</strong> vao day<br>
  hoac click de chon cac file
  <input type="file" id="fileInput" multiple webkitdirectory style="display:none">
</div>
<div id="browser-note">Keo tha folder: chi ho tro Chrome/Edge. Firefox: dung click chon file.</div>

<div id="log"></div>
<button id="btn-download" onclick="window.location.href=downloadUrl">
  Tai file ket qua
</button>

<script>
  const socket = io();
  const log    = document.getElementById('log');
  const btn    = document.getElementById('btn-download');
  let downloadUrl = '';
  let currentJobId = '';

  function addLog(msg) {
    log.innerHTML += msg + '\n';
    log.scrollTop  = log.scrollHeight;
  }

  socket.on('log',   d => { if (d.job_id === currentJobId) addLog(d.msg); });
  socket.on('done',  d => {
    if (d.job_id !== currentJobId) return;
    addLog('HOAN THANH!');
    downloadUrl = d.download_url;
    btn.style.display = 'inline-block';
  });
  socket.on('error', d => {
    if (d.job_id !== currentJobId) return;
    addLog('LOI: ' + d.msg);
  });

  function uploadFiles(files) {
    if (!files.length) return;
    btn.style.display = 'none';
    log.innerHTML = '';
    addLog(`Dang tai len ${files.length} file...`);

    const fd   = new FormData();
    const ngay = document.getElementById('ngay').value.trim();
    if (ngay) fd.append('ngay_doi_chieu', ngay);
    for (const f of files) fd.append('files', f, f.webkitRelativePath || f.name);

    fetch('/upload', { method: 'POST', body: fd })
      .then(r => r.json())
      .then(d => {
        currentJobId = d.job_id;
        addLog(`Job ID: ${d.job_id} — Bat dau xu ly...`);
      })
      .catch(e => addLog('Loi upload: ' + e));
  }

  const zone = document.getElementById('dropzone');
  zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', ()  => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    uploadFiles([...e.dataTransfer.files]);
  });

  document.getElementById('fileInput').addEventListener('change', e => {
    uploadFiles([...e.target.files]);
  });
</script>
</body>
</html>
```

---

### Web-B7 — `START_WEB.bat` — Khởi động nhanh

```batch
@echo off
chcp 65001 > nul
echo ====================================
echo  DOI CHIEU ACH - WEB UI
echo ====================================
cd /d "%~dp0"
python web_app.py
pause
```

---

## 4. Kiến trúc tổng thể sau nâng cấp

```
┌─────────────────────────────────────────────────────────┐
│                      MÁY CHỦ                           │
│                                                         │
│  START_WEB.bat                                          │
│       │                                                 │
│  web_app.py (Flask + SocketIO, port 8080)               │
│       │   ping_timeout=300, ping_interval=25            │
│       ├── /upload    → lưu file vào uploads/job_id/     │
│       ├── /download  → trả file output về client        │
│       └── SocketIO   → emit log real-time               │
│                │                                        │
│         main_from_dir()  ← refactor từ main()           │
│                │  (ngày tính local, thread-safe)         │
│    ┌───────────┼───────────┐                            │
│   B2          B3          B6     (ThreadPool × 3)       │
│    │           │           │                            │
│    │           │      B6 nội bộ: (ThreadPool × 2)      │
│    └───────────┼───────────┘                            │
│     B4 (ThreadPool × 2 bên trong, tpay tham số hoá)    │
│    ┌───────────┴───────────┐                            │
│   B5                      B7     (ThreadPool × 2)       │
│    └───────────┬───────────┘                            │
│           xuat_excel()                                   │
│             ├── MIS_DI_KHOP > 50k → CSV (2.5s/500k)    │
│             ├── MIS_DEN_KHOP > 50k → CSV (2.5s/500k)   │
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
│  → Click tải .xlsx       │
└──────────────────────────┘
```

---

## 5. Thứ tự triển khai & ước lượng công việc

| Bước | Nội dung | Ưu tiên | Ước công |
|---|---|---|---|
| **1** | A_NEW — CSV output cho MIS_DI_KHOP + MIS_DEN_KHOP | 🔴 ~4–5 phút tiết kiệm | 45 phút |
| **2** | A1 — ThreadPool ZIP bên trong B4 + B6 | 🟡 ~10–20% đọc dữ liệu | 30 phút |
| **3** | A_CLEAN — `_clean()` warning thiếu cột | 🟡 Chất lượng kết quả | 10 phút |
| **4** | Web-B4 — Vá thread-safety B4 (tpay tham số hoá) | 🔴 Bắt buộc trước Web UI | 30 phút |
| **5** | Web-B5 — `main_from_dir()` + refactor `main()` | 🔴 Cần trước Web server | 45 phút |
| **6** | Web-B3 — `web_app.py` + ping_timeout | 🔴 Core Web UI | 60 phút |
| **7** | Web-B6 — `templates/index.html` | 🔴 Core Web UI | 45 phút |
| **8** | A5 — tqdm progress bar | 🟢 UX | 15 phút |
| **9** | Web-B7 — `START_WEB.bat` + cập nhật `requirements.txt` | 🟢 | 10 phút |
| **10** | Test tổng thể trên dữ liệu thật (đo timer từng bước) | 🔴 Bắt buộc | 60 phút |
| | **Tổng** | | **~6 giờ** |

---

## 6. Chi tiết code cần thêm / sửa

### Tóm tắt file cần đụng

| File | Loại thay đổi | Nội dung |
|---|---|---|
| `main.py` | Sửa + Thêm | A_NEW: `CSV_THRESHOLD` + logic CSV trong `xuat_excel()`; A_CLEAN: sửa `_clean()`; Web-B5: thêm `main_from_dir()`, refactor `main()` |
| `modules/b4_xu_ly_mis_di.py` | Sửa | A1: ThreadPool 2 ZIP; Web-B4: tham số hoá `tpay_tu`/`tpay_den` |
| `modules/b6_xu_ly_mis_den.py` | Sửa | A1: ThreadPool 2 ZIP |
| `modules/b2_xu_ly_gl02.py` | **Không đổi** | A3 đã loại bỏ |
| `modules/b3_xu_ly_gw.py` | **Không đổi** | |
| `config.py` | **Không đổi** | |
| `web_app.py` | Tạo mới | Flask server với `ping_timeout=300` |
| `templates/index.html` | Tạo mới | Web UI kéo thả |
| `START_WEB.bat` | Tạo mới | Launcher |
| `requirements.txt` | Sửa | Thêm flask, flask-socketio, eventlet, tqdm |

---

## 7. Kiểm thử & đo lường kết quả

### Đo tốc độ trước/sau

```python
import time
t0 = time.perf_counter()
# ... bước xử lý ...
print(f'[TIMER] {label}: {time.perf_counter() - t0:.1f}s')
```

### Checklist kiểm thử

- [ ] **A_NEW**: Chạy với dữ liệu thật — log `[CSV]` xuất hiện nếu sheet > 50k dòng; mở CSV kiểm tra đủ dòng, không bị lệch cột
- [ ] **A_NEW**: Số dòng trong CSV khớp với kết quả cũ (không mất dữ liệu)
- [ ] **A1**: Log `[B4]` và `[B6]` — kiểm tra `tong truoc timeout` đúng số dòng
- [ ] **A_CLEAN**: Chạy thử với DataFrame thiếu cột — xác nhận warning `[WARN] _clean(...)` xuất hiện
- [ ] **Thread-safety B4**: 2 job khác ngày cùng lúc → kết quả không bị trộn lẫn ngày
- [ ] **B Web UI**: Mở `http://<IP>:8080` từ máy khác trong LAN — thấy giao diện
- [ ] **B Upload**: Kéo thả folder → thấy log chạy real-time → tải được file .xlsx
- [ ] **B File lớn**: Upload 2.4tr dòng qua LAN → không timeout (ping_timeout=300)

### Đo tốc độ mục tiêu

| Giai đoạn | Hiện tại (ước tính) | Sau nâng cấp |
|---|---|---|
| Đọc B2+B3+B6 (song song) | ~3–4 phút | ~2.5–3 phút (A1) |
| Đọc B4 MIS_DI | ~3–4 phút | ~2.5–3 phút (A1) |
| Đối chiếu B5+B7 | ~30 giây | ~30 giây (không đổi) |
| Ghi Excel/CSV | ~4–6 phút | **~15–30 giây** (A_NEW nếu kích hoạt) |
| **Tổng** | **~10.7 phút** | **~5–7 phút** |

> **Ghi chú:** Nếu MIS_DI_KHOP + MIS_DEN_KHOP đều < 50k dòng, A_NEW không kích hoạt.
> Cần đo timer thực tế sau bước 1 để xác định mức tiết kiệm chính xác.

---

*Kế hoạch lập bởi Claude Sonnet 4.6 — 19/06/2026 (v2 cập nhật từ benchmark + review thực tế)*
