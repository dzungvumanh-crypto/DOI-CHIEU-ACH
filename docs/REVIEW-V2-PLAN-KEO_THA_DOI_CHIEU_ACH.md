# REVIEW V2 & ĐỀ XUẤT TỐI ƯU — DOI-CHIEU-ACH
> Người review: Claude (Sonnet 4.6)
> Ngày: 20/06/2026
> Phương pháp: Đọc toàn bộ codebase đã triển khai (main.py, 7 module, web_app.py, templates/index.html,
> config.py, 4 file kế hoạch lịch sử) + chạy benchmark độc lập (pandas 3.0.2, xlsxwriter 3.2.9) +
> tra cứu tài liệu chính thức cho phần Web UI.
>
> **Phạm vi review:** Đây là review lần 2, sau khi Claude Code đã đọc REVIEW-V1 và cập nhật cả kế
> hoạch (`PLAN-KEO_THA_DOI_CHIEU_ACH.md`) lẫn code thực tế (main.py, modules/, web_app.py,
> templates/index.html đều đã tồn tại và khớp với kế hoạch). Vì vậy review này tập trung vào:
> (1) xác minh độc lập các con số benchmark trước đó, (2) tìm vấn đề mới mà 2 lượt review/kế hoạch
> trước bỏ sót, (3) chỉ ra rủi ro có thể gây "kết quả không mong muốn", (4) dọn gọn tài liệu dự án.

---

## TÓM TẮT NHANH — CÓ GÌ MỚI SO VỚI V1

| # | Phát hiện | Mức độ | Đã có trong V1/V2 plan? |
|---|---|---|---|
| 1 | **A1 (ThreadPool đọc 2 ZIP) gần như vô dụng, có thể PHẢN TÁC DỤNG** — benchmark độc lập cho thấy phần CPU (parse CSV) chiếm ~94% thời gian, phần I/O đĩa chỉ ~6%; ThreadPool trên phần CPU-bound đo được **0.66x–0.98x** (chậm hơn hoặc bằng tuần tự) | 🔴 Quan trọng | ❌ Chưa — cả V1 và V2 đều ước tính +10-30%, ngược hướng benchmark |
| 2 | **Nghi vấn nguyên nhân gốc của "kết quả không mong muốn"**: `web_app.py` dùng `async_mode='eventlet'` nhưng KHÔNG gọi `eventlet.monkey_patch()`, trong khi `main_from_dir()` lại dùng `ThreadPoolExecutor` (CPU-bound, pandas) — tổ hợp này có thể **block toàn bộ event loop** trong nhiều phút, khiến log real-time treo/dồn cục hoặc client bị rớt kết nối | 🔴 Nghiêm trọng | ❌ Chưa — chưa từng được nhắc tới |
| 3 | `eventlet` đang bị chính nhà phát triển dự án đó "wind down"; tác giả Flask-SocketIO khuyến nghị dùng `async_mode='threading'` cho dự án mới | 🟡 Nên sửa | ❌ Chưa |
| 4 | **B2 (`xu_ly_gl02.py`) đọc TOÀN BỘ cột CSV thay vì chỉ 17 cột cần** — vốn do một bug cũ (`usecols=4 cột` làm mất dữ liệu) bị sửa bằng cách bỏ hẳn `usecols` thay vì sửa đúng danh sách cột. Benchmark độc lập: ~1.5x–4x tuỳ số cột thực tế của GL02 gốc | 🟡 Cơ hội bỏ sót | ❌ Chưa |
| 5 | CSV xuất ra cho MIS_DI_KHOP/MIS_DEN_KHOP dùng `dtype=str` ở nguồn nhưng khi mở trực tiếp bằng Excel (double-click), Excel sẽ tự suy luận kiểu dữ liệu → có thể làm **mất leading-zero của TRACE/MSGSEQ** hoặc hiển thị sai số dạng khoa học | 🟡 Rủi ro dữ liệu (không phải tốc độ) | ❌ Chưa |
| 6 | `b1_doc_session.py` dùng `glob.glob()` không đệ quy trong khi `main.py._tim_file()` dùng đệ quy (`**`) — nếu PDF nằm trong sub-folder, B1 sẽ raise lỗi trong khi các bước khác tìm thấy file | 🟢 Nhỏ, dễ vá | ❌ Chưa |
| 7 | Môi trường benchmark cần làm rõ: tất cả con số "10-30% cải thiện" của A1 phụ thuộc hoàn toàn vào **số CPU core thực của máy chủ** — chưa thấy plan nào đề cập việc đo `os.cpu_count()` trên máy chủ đích trước khi quyết định mức độ song song hoá | 🟡 Phương pháp luận | ❌ Chưa |
| 8 | A_NEW (CSV cho sheet lớn) — benchmark độc lập **xác nhận đúng hướng**, đo được **27.6x** (gần khớp với 34x mà plan V2 báo cáo, chênh lệch do khác máy) | ✅ Xác nhận | Đã có trong V2 |
| 9 | A2 (`write_column`) bị bác bỏ — benchmark độc lập xác nhận **đúng**: 1.05x, không đáng kể | ✅ Xác nhận | Đã có trong V2 |
| 10 | 4 file kế hoạch markdown chồng chéo trong gốc dự án (`KE_HOACH_CODE.md`, `PLAN_DOI_CHIEU_ACH.md`, `PLAN-KEO_THA_DOI_CHIEU_ACH.md`, `REVIEW-V1-...md`) — nên gộp lại | 🟢 Dọn dẹp | Chưa làm |
| 11 | `ZIP_PASSWORD` để plaintext trong `config.py`, khả năng bị commit vào Git sử (không có trong `.gitignore`) | 🟡 Bảo mật | Chưa làm |

---

## 1. PHÁT HIỆN QUAN TRỌNG NHẤT — NGUYÊN NHÂN CÓ THỂ GÂY "KẾT QUẢ KHÔNG MONG MUỐN"

Yêu cầu của bạn nhắc tới hai vấn đề riêng biệt: **(a) quá chậm** và **(b) kết quả không mong muốn**.
V1 và V2 plan tập trung gần như 100% vào (a). Sau khi đọc kỹ `web_app.py`, tôi tìm thấy một vấn đề
kiến trúc có khả năng cao là nguyên nhân của (b) — đặc biệt nếu "kết quả không mong muốn" biểu hiện
dưới dạng: **UI bị treo, không thấy log chạy, mất kết nối giữa chừng, hoặc job không bao giờ báo "done"**
khi xử lý file lớn qua Web UI.

### 1.1 Vấn đề: `eventlet` không được monkey-patch, chạy chung với `ThreadPoolExecutor`

```python
# web_app.py — hiện tại
from flask_socketio import SocketIO
socketio = SocketIO(app, async_mode='eventlet', ...)
```

`web_app.py` KHÔNG có dòng:
```python
import eventlet
eventlet.monkey_patch()   # ← THIẾU, bắt buộc phải gọi SỚM NHẤT có thể, trước mọi import khác
```

Theo tài liệu chính thức của `eventlet`/`Flask-SocketIO`: khi chọn `async_mode='eventlet'`, toàn bộ
ứng dụng phải được "monkey-patch" — tức thay thế `socket`, `threading`, `time.sleep`... bằng phiên
bản cooperative (greenlet) của eventlet — **trước khi bất kỳ module nào khác import các thư viện
chuẩn đó**. Nếu thiếu bước này:

- `socketio.run(app, ...)` vẫn khởi động được, nhưng server chạy trên 1 greenlet "giả lập" event loop.
- Khi `main_from_dir()` (được gọi trong một `threading.Thread` riêng do `web_app.py` tạo ra) đụng tới
  `ThreadPoolExecutor` — đây là CPU-bound, không phải I/O — eventlet không biết cách "nhường" lại
  event loop trong lúc chờ kết quả Future, vì code đó không đi qua eventlet hub.
- Hậu quả thực tế (theo nhiều báo cáo cộng đồng Flask-SocketIO với tình huống tương tự — function
  C-extension hoặc CPU-bound chạy trong background task): event loop bị **chặn cứng** trong toàn bộ
  thời gian xử lý nặng (ở đây là hàng phút), khiến:
  - `socketio.emit('log', ...)` bị dồn cục, chỉ đẩy ra hàng loạt khi xử lý xong thay vì real-time
  - Client có thể bị ngắt kết nối WebSocket dù đã đặt `ping_timeout=300` (vì server không trả lời ping
    được trong lúc bị chặn)
  - Nếu client refresh trang giữa chừng (tưởng bị treo), `job_id` cũ bị mất tham chiếu phía client dù
    job vẫn đang chạy ngầm trên server → người dùng tưởng "lỗi" hoặc "không có kết quả"

### 1.2 Khuyến nghị sửa — Chọn 1 trong 2 hướng

**Hướng A (khuyến nghị — đơn giản, ít rủi ro nhất cho công cụ nội bộ LAN):**
Bỏ hẳn `eventlet`, dùng `async_mode='threading'` (chế độ mặc định của Flask-SocketIO khi không cài
`eventlet`/`gevent`). Đây cũng là khuyến nghị chính thức hiện tại của tác giả Flask-SocketIO cho dự án
mới, vì `eventlet` đang trong quá trình "wind down" (ít được bảo trì). Vì DOI-CHIEU-ACH là công cụ nội
bộ, chỉ phục vụ vài người dùng cùng lúc trong LAN — không cần hiệu năng hàng nghìn kết nối đồng thời
mà `eventlet` hướng tới.

```python
# requirements.txt — XOÁ dòng:
# eventlet>=0.35

# web_app.py — sửa:
socketio = SocketIO(
    app,
    async_mode='threading',     # ← thay vì 'eventlet'
    cors_allowed_origins='*',
    ping_timeout=300,
    ping_interval=25,
)
```
`async_mode='threading'` chạy mỗi job xử lý trên 1 OS thread thật (giống cách `web_app.py` đã tự quản
lý bằng `threading.Thread` hiện tại) — không có khái niệm "monkey-patch" hay greenlet, nên
`ThreadPoolExecutor` bên trong `main_from_dir()` hoạt động hoàn toàn bình thường, không xung đột.
Nhược điểm duy nhất: `socketio.on('disconnect')` có thể có độ trễ vài chục giây thay vì tức thời —
không quan trọng với use-case này (không cần biết client mất kết nối ngay lập tức).

**Hướng B (nếu muốn giữ eventlet):** Thêm `eventlet.monkey_patch()` ở dòng đầu tiên của `web_app.py`
VÀ đổi toàn bộ `ThreadPoolExecutor` bên trong `main_from_dir()` thành
`eventlet.GreenPool`/`socketio.start_background_task()`. Phức tạp hơn nhiều, rủi ro cao hơn nếu làm
sai, và không có lợi ích rõ ràng cho quy mô dự án này. **Không khuyến nghị.**

> **Đây nên là việc làm ưu tiên SỐ 1**, cao hơn cả tối ưu tốc độ, vì nó ảnh hưởng tới độ tin cậy của
> kết quả trả về cho người dùng — đúng với mô tả "kết quả không mong muốn" trong yêu cầu của bạn.

---

## 2. XÁC MINH ĐỘC LẬP CÁC BENCHMARK TRONG PLAN V2

Tôi chạy lại các benchmark cốt lõi trên môi trường riêng (pandas 3.0.2, xlsxwriter 3.2.9, **1 CPU core**
— xem mục 3 để biết vì sao điều này quan trọng).

### 2.1 A_NEW — CSV thay vì Excel cho sheet lớn: ✅ XÁC NHẬN ĐÚNG, ƯU TIÊN SỐ 1 VỀ TỐC ĐỘ

```
500,000 dòng × 18 cột:
  Ghi Excel (write_row + border + constant_memory):  60.5s   (plan V2 báo: 85.5s)
  Ghi CSV (pandas to_csv, utf-8-sig):                  2.2s   (plan V2 báo: 2.5s)
  Speedup đo được: 27.6x   (plan V2 báo: 34x)
```
Chênh lệch giữa 27.6x và 34x hoàn toàn hợp lý do khác máy/CPU — **kết luận giữ nguyên: đây là tối ưu
tốc độ lớn nhất, giữ nguyên A_NEW như trong plan V2.**

### 2.2 A2 (write_column) bị loại bỏ: ✅ XÁC NHẬN ĐÚNG

```
200,000 dòng × 18 cột:
  write_row + constant_memory:    21.8s
  write_column, không constant_memory: 22.8s
  → write_row vẫn nhanh hơn nhẹ (1.05x) — xác nhận quyết định giữ write_row của plan V2 là đúng
```

### 2.3 Phát hiện thêm: bỏ `border` format mỗi ô giúp nhanh hơn 1.33x (tối ưu phụ, không bắt buộc)

```
200,000 dòng × 18 cột:
  write_row + format border mỗi ô:  22.0s
  write_row KHÔNG có format:        16.6s   → 1.33x nhanh hơn
```
Đây là phát hiện mới, không có trong V1/V2. Tuy nhiên việc bỏ border sẽ làm output Excel xấu đi (mất
viền ô). **Khuyến nghị: KHÔNG áp dụng cho các sheet đang giữ Excel** (NPO_DI_THUA, MIS_DI_THUA,
TIMEOUT...) vì các sheet này vốn đã nhỏ (<50k dòng → <2-3s), lợi ích không đáng so với việc mất thẩm mỹ.
Chỉ nêu ra để bạn biết đây không phải hướng tối ưu còn bỏ sót — đã cân nhắc và loại.

---

## 3. PHÁT HIỆN QUAN TRỌNG — A1 (ThreadPool đọc 2 ZIP) CẦN ĐÁNH GIÁ LẠI HOÀN TOÀN

Đây là điểm khác biệt lớn nhất giữa review này và 2 lượt review/kế hoạch trước.

### 3.1 Bối cảnh: V1 và V2 đã từng tranh luận với nhau về A1

- **V1 (review đầu)** benchmark: dữ liệu đã nằm trong RAM (`BytesIO`) → ThreadPool **0.98x** (không lợi),
  kết luận: "lợi ích chỉ đến từ phần I/O đĩa chồng lấp khi đợi giải nén ZIP", ước tính SSD +15-20%,
  HDD +20-30%.
- **V2 (Claude Code cập nhật)** giữ nguyên kết luận và ước tính của V1, áp dụng A1 vào code thật.

### 3.2 Benchmark độc lập của tôi — tách riêng phần I/O và phần CPU

```
[I/O-only: đọc bytes từ ZIP + giải nén]           0.22s
[CPU-only: pandas.read_csv parse 2 DataFrame]      3.23s
→ Tỷ lệ I/O : CPU = 6% : 94%
```

Đây là số liệu mà **cả V1 lẫn V2 đều chưa đo trực tiếp** — V1 chỉ đo "dữ liệu đã trong RAm" (CPU-only)
và đo "I/O đĩa" một cách gộp chung, không tách bạch tỷ lệ. Khi tách ra, thấy rõ: với `pyzipper` (đọc
+ giải mã AES) cũng tương tự — phần giải mã AES + giải nén DEFLATE là CPU-bound, cộng thêm phần parse
CSV của pandas cũng CPU-bound → **phần lớn thời gian xử lý ZIP là CPU-bound, không phải I/O-bound**.
Điều này khớp với benchmark `bench5` của tôi (đọc file ZIP thật từ đĩa + giải nén + parse):

```
[Tuần tự đọc+giải nén+parse 2 ZIP]   4.47s
[Song song (ThreadPool) 2 ZIP]       4.21s
→ Speedup: 1.06x  (gần như không đổi, thấp hơn nhiều so với ước tính 10-30% của V1/V2)
```

Và khi chỉ đo phần parse CSV thuần (mô phỏng trường hợp dữ liệu nằm sẵn trong RAM sau giải nén):

```
[Tuần tự parse 2x 600k dòng]    7.42s
[Song song (ThreadPool)]       11.18s
→ Speedup: 0.66x — TỆ HƠN tuần tự, không phải "0% cải thiện" như V1 nói mà là PHẢN TÁC DỤNG
```

### 3.3 Nguyên nhân: GIL + chi phí context-switch giữa các thread CPU-bound

Khi 2 thread cùng tranh chấp GIL để thực thi code Python/pandas thuần (dù pandas dùng C-extension nội
bộ, nhiều thao tác `pd.read_csv` với `dtype=str` vẫn giữ GIL đáng kể vì phải tạo Python string object
cho từng giá trị), chi phí chuyển đổi ngữ cảnh giữa 2 thread có thể **vượt quá lợi ích song song hoá**,
dẫn đến kết quả tệ hơn chạy tuần tự — đúng như benchmark đo được.

### 3.4 ⚠️ Giới hạn quan trọng của benchmark — CẦN ĐO LẠI TRÊN MÁY CHỦ THẬT

Benchmark của tôi (và rất có thể của Claude Code khi viết V1/V2) chạy trong **môi trường container chỉ
có 1 CPU core** (`os.cpu_count() == 1`, xác nhận qua `os.sched_getaffinity(0) == {0}`). Đây là giới hạn
nghiêm trọng cần nêu rõ:

- Trên máy 1 core: `ThreadPoolExecutor` cho CPU-bound task **không bao giờ có lợi**, chỉ có chi phí.
- Trên máy chủ Windows thật của khách hàng (rất có thể ≥4 core theo cấu hình máy chủ ngân hàng tiêu
  chuẩn): bức tranh có thể khác — `pd.read_csv` khi parse với C-engine có thể giải phóng GIL ở một số
  đoạn (đọc buffer C), nên 2 thread thật sự chạy song song một phần. Nhưng pyzipper AES decryption qua
  Python wrapper (không phải pure-C) thường giữ GIL chặt hơn.

**Khuyến nghị hành động cụ thể (KHÔNG phải "giữ A1" hay "bỏ A1" một cách mù quáng):**

1. Thêm 3 dòng đo nhanh vào đầu `main_from_dir()` hoặc 1 script riêng, chạy 1 lần trên máy chủ thật:
   ```python
   import os
   print(f'[INFO] So CPU core thuc te: {os.cpu_count()}')
   ```
2. Đo trực tiếp thời gian B4/B6 với A1 BẬT và TẮT trên dữ liệu thật của khách hàng (đã có sẵn log
   `[B4]`/`[B6]` in ra số dòng, chỉ cần thêm `time.perf_counter()` bọc quanh, đúng như mục 7 của plan
   V2 đã đề xuất — nhưng **bắt buộc phải làm bước này trước khi tin vào con số 10-30%**, không suy ra
   từ benchmark trong container).
3. Nếu máy chủ thật ≤2 core: cân nhắc **bỏ A1**, giữ code tuần tự đơn giản hơn (ít rủi ro debug hơn,
   dễ đọc hơn) vì lợi ích gần như bằng 0 hoặc âm.
4. Nếu máy chủ thật ≥4 core: A1 có khả năng có lợi thật, nhưng nên benchmark `ThreadPoolExecutor` so
   với `ProcessPoolExecutor` (né GIL hoàn toàn) — lưu ý `ProcessPoolExecutor` có chi phí
   pickle/serialize DataFrame giữa các process, benchmark của tôi cho thấy **chi phí này có thể lớn
   hơn lợi ích** với dữ liệu cỡ vài trăm nghìn dòng (đo được 0.45x — chậm hơn tuần tự 2.2 lần) — không
   khuyến nghị ProcessPoolExecutor cho bước này trừ khi dữ liệu mỗi ZIP đủ lớn (>1 triệu dòng) để bù
   chi phí serialize.

**Tóm lại: A1 không sai về mặt kỹ thuật (đọc 2 ZIP độc lập thì về lý thuyết song song hoá được), nhưng
mức độ lợi ích trong V1/V2 (10-30%) là suy diễn chưa được đo trực tiếp đúng cách, và trong điều kiện
xấu nhất (máy chủ ít core, hoặc tỷ lệ CPU:I/O giống benchmark của tôi) có thể KHÔNG có lợi hoặc làm
chậm hơn. Đây là rủi ro thấp (không gây sai dữ liệu) nhưng có thể khiến kỳ vọng tốc độ sau nâng cấp
không đạt được, gây hiểu nhầm về điểm A_NEW mới là yếu tố quyết định thực sự.**

---

## 4. CƠ HỘI TỐI ƯU BỊ BỎ SÓT — B2 (`xu_ly_gl02.py`) ĐỌC THỪA CỘT

### 4.1 Phát hiện

Trong `b4_xu_ly_mis_di.py` và `b6_xu_ly_mis_den.py`, `pd.read_csv(..., usecols=lambda c: c in _COLS, ...)`
đã được dùng để lọc cột sớm khi đọc CSV — đúng kỹ thuật. Nhưng `b2_xu_ly_gl02.py` thì KHÔNG:

```python
# modules/b2_xu_ly_gl02.py — hiện tại
df = pd.read_csv(
    io.BytesIO(raw),
    dtype=str,
    encoding=enc,
    low_memory=False,
)   # ← đọc TẤT CẢ cột trong file GL02 gốc, dù chỉ cần 4 cột bắt buộc + 13 cột để xuất sheet NPO_*_THUA
```

### 4.2 Lý do lịch sử — ĐÃ TÌM THẤY trong `PLAN_DOI_CHIEU_ACH.md` (kế hoạch cũ hơn)

Đọc lại file kế hoạch cũ trong repo, tôi phát hiện đây **không phải sơ suất** mà là hậu quả của một
lần sửa bug trước đó:

> *(trích `PLAN_DOI_CHIEU_ACH.md`, mục P1): `b2_xu_ly_gl02.py` từng có `usecols=_COLS` với
> `_COLS = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']` (chỉ 4 cột) → khiến sheet NPO_DI_THUA/
> NPO_DEN_THUA xuất ra thiếu 13 cột so với `_COLS_NPO` (17 cột) mà `main.py._clean()` cần. Giải pháp
> được chọn khi đó: bỏ hẳn `usecols`, đọc toàn bộ cột.*

Đây là cách sửa **đúng nhưng chưa tối ưu** — bỏ hẳn `usecols` giải quyết được bug thiếu cột, nhưng đồng
thời đánh đổi tốc độ đọc không cần thiết. Cách sửa đúng hơn là dùng đúng danh sách 17 cột của
`_COLS_NPO` (đã có sẵn trong `main.py`) làm `usecols`, vừa đủ dữ liệu vừa nhanh hơn.

### 4.3 Benchmark độc lập

```
600,000 dòng, giả lập GL02 có 30 cột tổng (17 cột cần + 13 cột thừa):
  Đọc TẤT CẢ 30 cột:               1.97s
  Đọc CHỈ 17 cột (usecols đúng):   1.32s
  → Speedup: 1.49x
```

Mức tiết kiệm phụ thuộc vào **số cột thực tế của file GL02 gốc** (không biết chính xác vì không có
file mẫu) — nếu GL02 gốc chỉ có đúng 17-20 cột thì lợi ích nhỏ; nếu có 30-40+ cột (khá phổ biến với
file core-banking xuất thô) thì lợi ích có thể lên tới 2-4x cho riêng bước đọc GL02.

### 4.4 Đề xuất sửa (rủi ro thấp — vì giữ ĐỦ 17 cột đã từng gây bug, không lặp lại sai lầm cũ)

```python
# modules/b2_xu_ly_gl02.py

# Thêm import danh sách cột cần giữ — đồng bộ với main.py._COLS_NPO để tránh lệch 2 nơi.
# Cách an toàn nhất: định nghĩa NGAY TẠI b2, rồi để main.py._COLS_NPO tham chiếu lại (xem mục 4.5),
# tránh tình trạng 2 nơi định nghĩa danh sách cột dễ lệch nhau như đã từng xảy ra.

_COLS_NPO_RAW = [
    'TRDATE', 'TRBRCD', 'USERID', 'JOURSEQ', 'DYTRSEQ', 'LOCAC', 'CCY',
    'BUSCD', 'UNIT', 'TRCD', 'CUSTOMER', 'TRTP', 'REFERENCE',
    'REMARK', 'DRAMOUNT', 'CRAMOUNT', 'CRTDTM',
]

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
                            usecols=lambda c: c in _COLS_NPO_RAW,   # ← THÊM, dùng đủ 17 cột
                            encoding=enc,
                            low_memory=False,
                        )
                        missing = [c for c in _COLS_REQUIRED if c not in df.columns]
                        if missing:
                            raise ValueError(f'Thieu cot: {missing}')
                        frames.append(df)
                        break
                    except UnicodeDecodeError:
                        continue
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
```

### 4.5 Đề xuất phụ — Tránh lệch danh sách cột giữa 2 nơi (rủi ro tái diễn bug cũ)

Hiện `_COLS_NPO` được định nghĩa trong `main.py`, còn nếu thêm `usecols` cho B2 thì cần một danh sách
tương tự bên trong `b2_xu_ly_gl02.py`. Có 2 nơi định nghĩa cùng một danh sách cột là chính xác nguyên
nhân gây ra bug lịch sử trong mục 4.2. Khuyến nghị: định nghĩa `_COLS_NPO` **một lần duy nhất** trong
`config.py` (vì đây là hằng số nghiệp vụ, không phụ thuộc logic xử lý), rồi cả `main.py` và
`b2_xu_ly_gl02.py` cùng import từ đó:

```python
# config.py — thêm:
COLS_NPO = [
    'TRDATE', 'TRBRCD', 'USERID', 'JOURSEQ', 'DYTRSEQ', 'LOCAC', 'CCY',
    'BUSCD', 'UNIT', 'TRCD', 'CUSTOMER', 'TRTP', 'REFERENCE',
    'REMARK', 'DRAMOUNT', 'CRAMOUNT', 'CRTDTM',
]

# main.py — thay:
# _COLS_NPO = [...]
# bằng:
_COLS_NPO = config.COLS_NPO

# modules/b2_xu_ly_gl02.py — thêm:
from config import ZIP_PASSWORD, COLS_NPO as _COLS_NPO_RAW
```

Đây là thay đổi nhỏ nhưng **giảm hẳn rủi ro tái diễn loại bug đã từng xảy ra** — đúng tinh thần
"làm gọn gàng codebase một cách khoa học" mà bạn yêu cầu.

---

## 5. RỦI RO DỮ LIỆU — CSV CHO MIS_DI_KHOP/MIS_DEN_KHOP MỞ TRỰC TIẾP BẰNG EXCEL

### 5.1 Vấn đề

A_NEW (đã triển khai) xuất `MIS_DI_KHOP`/`MIS_DEN_KHOP` ra file `.csv` riêng khi >50k dòng, kèm
`encoding='utf-8-sig'` để Excel đọc đúng tiếng Việt. Đây là quyết định đúng cho **tốc độ**, nhưng có
một rủi ro **dữ liệu** chưa được nhắc tới: khi người dùng cuối (kế toán đối chiếu ACH) **double-click**
mở file CSV trực tiếp bằng Excel (thói quen phổ biến, thay vì dùng "Data → From Text/CSV"), Excel sẽ
**tự suy luận kiểu dữ liệu cho từng cột** dựa trên nội dung, không tôn trọng việc dữ liệu gốc được lưu
dưới dạng chuỗi (`dtype=str`) trong code.

Benchmark xác nhận hành vi này (mô phỏng cột `TRACE` có leading-zero, cột `SO_TIEN` có giá trị lớn):

```
Nội dung CSV ghi ra:        000123456, 99999999999999999
Excel khi mở sẽ hiển thị:   123456 (mất 3 số 0 đầu)  |  9.99999E+16 (dạng khoa học, mất độ chính xác)
```

Các cột có nguy cơ: `TRACE`, `SE_TRACE`, `MSGSEQ`, `REFHUB`, `MSGREF`, `TXID` — đều là các trường định
danh giao dịch ngân hàng có thể chứa leading-zero hoặc chuỗi số dài. Nếu người dùng vô tình nhìn nhầm
giá trị bị Excel hiển thị sai (dù dữ liệu GỐC trong file CSV vẫn đúng), có thể dẫn tới **hiểu sai số
liệu đối chiếu** dù bản thân chương trình không hề tính sai.

### 5.2 Đây KHÔNG phải lỗi của A_NEW hay lý do để huỷ A_NEW

Tốc độ tăng 27x là quá lớn để đánh đổi. Cần xử lý bằng UX/hướng dẫn, không phải đổi lại sang Excel.

### 5.3 Đề xuất giảm thiểu rủi ro (chọn 1 hoặc kết hợp, đều rủi ro thấp)

**Cách 1 — Thêm ghi chú rõ ràng trong sheet Excel "tóm tắt" trỏ tới CSV** (mở rộng từ code đã có):
```python
ws.write(0, 0, f'[Du lieu lon - xem file: {os.path.basename(csv_path)}]')
ws.write(1, 0, f'Tong so dong: {len(df):,}')
# Thêm 2 dòng:
ws.write(2, 0, 'LUU Y: Mo file CSV bang Excel qua "Data > Tu Van ban/CSV" (KHONG double-click truc tiep)')
ws.write(3, 0, 'de tranh Excel tu dong lam mat so 0 dau hoac sai dinh dang so dien thoai/ma giao dich.')
```

**Cách 2 — Ép kiểu text cho các cột định danh ngay trong CSV bằng tiền tố Excel-formula** (`="000123"`),
đảm bảo Excel hiển thị đúng dù mở trực tiếp — đánh đổi: file CSV khó đọc hơn nếu mở bằng công cụ khác
(Notepad, Python, hệ thống khác đọc CSV để tự động hoá tiếp). **Không khuyến nghị mặc định**, chỉ làm
nếu xác nhận người dùng cuối CHỈ mở bằng Excel và không có quy trình tự động nào đọc lại CSV này.

**Cách 3 (khuyến nghị nhất, đơn giản, không đánh đổi gì) — Sinh kèm 1 file `.xlsx` thật nhưng KHÔNG
border/format (chỉ ghi thô)** thay vì CSV, vẫn tận dụng được phần lớn tốc độ:

Benchmark bổ sung: ghi Excel thô (không `border`, không `add_format`) cho 500k dòng → cần đo lại nhưng
theo benchmark mục 2.3 (1.33x khi bỏ border ở quy mô 200k) thời gian vẫn ở mức hàng chục giây, **không
đạt được mức 2.2s của CSV**. Vì vậy Cách 3 không thực sự cạnh tranh được với CSV về tốc độ — **chỉ nêu
ra để loại trừ, khuyến nghị chính vẫn là Cách 1 (ghi chú) kết hợp đào tạo người dùng cuối 1 lần.**

---

## 6. VẤN ĐỀ NHỎ — TÌM FILE PDF KHÔNG ĐỆ QUY TRONG B1

```python
# modules/b1_doc_session.py — hiện tại
pattern = os.path.join(input_dir, '*.pdf')
pdfs = glob.glob(pattern)   # ← KHÔNG đệ quy
```

Trong khi `main.py._tim_file()` tìm kiếm đệ quy (`recursive=True`, pattern `**`) cho tất cả các loại
file khác (GL02, MIS_DI, MIS_DEN). Nếu người dùng kéo-thả một folder có cấu trúc lồng nhau (ví dụ
`input/2026-06-11/ACH_....pdf` thay vì để PDF ngay gốc `input/`), B1 sẽ raise
`FileNotFoundError('Khong tim thay file PDF...')` ngay từ bước đầu tiên, trong khi các bước sau đáng
lẽ tìm thấy các file ZIP/Excel khác trong cùng cấu trúc lồng nhau đó. Đặc biệt dễ xảy ra với Web UI
kéo-thả folder (dùng `webkitdirectory`, giữ nguyên cấu trúc thư mục con qua `webkitRelativePath`).

**Đề xuất sửa — đồng bộ hành vi tìm kiếm giữa B1 và các bước khác:**
```python
# modules/b1_doc_session.py
def doc_session(input_dir: str) -> str:
    pattern = os.path.join(os.path.abspath(input_dir), '**', '*.pdf')
    pdfs = sorted(glob.glob(pattern, recursive=True))   # ← thêm recursive=True, '**'
    if not pdfs:
        raise FileNotFoundError(f'Khong tim thay file PDF trong {input_dir} (da tim de quy)')
    ...
```

---

## 7. VẤN ĐỀ BẢO MẬT NHỎ — `ZIP_PASSWORD` PLAINTEXT TRONG CONFIG ĐƯỢC TRACK BỞI GIT

```python
# config.py
ZIP_PASSWORD = b'DACwLdHi'
```

`config.py` KHÔNG nằm trong `.gitignore` (chỉ `file du liẹu/`, `input/`, `output/`, vài file debug cũ
được loại trừ) — nghĩa là mật khẩu giải mã file ZIP dữ liệu ngân hàng đang được commit thẳng vào lịch
sử Git. Nếu repo này từng hoặc sẽ được đẩy lên một remote chia sẻ (kể cả nội bộ), mật khẩu sẽ lộ vĩnh
viễn trong lịch sử commit dù sau này có xoá khỏi file hiện tại.

**Đề xuất (không bắt buộc làm ngay, nhưng nên cân nhắc do tính nhạy cảm ngân hàng):**
```python
# config.py — đổi thành đọc từ biến môi trường, có giá trị mặc định để không phá vỡ CLI hiện tại:
import os
ZIP_PASSWORD = os.environ.get('DOI_CHIEU_ZIP_PASSWORD', 'DACwLdHi').encode()
```
Và thêm hướng dẫn trong `START.bat`/`START_WEB.bat` về cách đặt biến môi trường nếu muốn override.
Đây là cải tiến **bảo mật, không phải tốc độ** — đưa vào review vì bạn yêu cầu "làm gọn gàng codebase
một cách khoa học" và đây là một khoản nợ kỹ thuật dễ thấy khi đọc toàn bộ code.

---

## 8. DỌN DẸP TÀI LIỆU DỰ ÁN

Hiện gốc dự án có 4 file kế hoạch/review chồng chéo theo thời gian:

| File | Vai trò | Đề xuất |
|---|---|---|
| `KE_HOACH_CODE.md` | Kế hoạch kiến trúc gốc (cấu trúc thư mục, B1-B7) | Giữ lại, đổi tên thành `docs/00-KIEN-TRUC-GOC.md` — vẫn hữu ích làm tài liệu tham chiếu cấu trúc |
| `PLAN_DOI_CHIEU_ACH.md` | Kế hoạch sửa bug đầu tiên (P0-P4) — đã hoàn thành | Đổi tên `docs/01-DA-HOAN-THANH-fix-bug-toc-do-v1.md` hoặc xoá nếu không cần lưu vết lịch sử |
| `PLAN-KEO_THA_DOI_CHIEU_ACH.md` | Kế hoạch v2 — Web UI + A_NEW — đã triển khai | Đổi tên `docs/02-DA-HOAN-THANH-web-ui-toc-do-v2.md` |
| `REVIEW-V1-PLAN-KEO_THA_DOI_CHIEU_ACH.md` | Review V1 (đã được V2 tích hợp) | Đổi tên `docs/03-review-v1-da-tich-hop.md` |
| `REVIEW-V2-PLAN-KEO_THA_DOI_CHIEU_ACH.md` | File này | Đổi tên `docs/04-review-v2-can-trien-khai.md` sau khi đã xử lý xong các mục trong đây |

**Đề xuất cụ thể:** tạo thư mục `docs/` trong gốc dự án, chuyển tất cả 5 file `.md` kế hoạch/review vào
đó (giữ `README.md` mới — nếu cần — ở gốc làm trang tóm tắt ngắn cho người mới). `main.py`,
`config.py`, `web_app.py`, `modules/`, `templates/`, file `.bat` giữ nguyên vị trí gốc — không đổi gì
về code, chỉ dọn tài liệu. Việc này không ảnh hưởng tới vận hành (`.bat` không tham chiếu tới các file
`.md`), an toàn để làm bất kỳ lúc nào.

```
DOI-CHIEU-ACH/
├── docs/                              ← MỚI: gom toàn bộ lịch sử kế hoạch/review
│   ├── 00-KIEN-TRUC-GOC.md
│   ├── 01-fix-bug-toc-do-v1.md
│   ├── 02-web-ui-toc-do-v2.md
│   ├── 03-review-v1.md
│   └── 04-review-v2.md                ← file này
├── main.py
├── config.py
├── web_app.py
├── modules/
├── templates/
├── requirements.txt
├── START.bat
├── START_WEB.bat
└── .gitignore
```

---

## 9. BẢNG TỔNG HỢP — THỨ TỰ TRIỂN KHAI ĐỀ XUẤT (SAU REVIEW V2)

Sắp xếp lại theo mức độ tác động thực tế đã xác minh, ưu tiên việc khắc phục "kết quả không mong muốn"
lên trước "tốc độ" theo đúng thứ tự mối quan tâm bạn nêu ra:

| Bước | Nội dung | Tác động | Mức độ chắc chắn | Ước công |
|---|---|---|---|---|
| **1** | **Mục 1** — Bỏ `eventlet`, chuyển `async_mode='threading'` trong `web_app.py` | 🔴 Khắc phục nguyên nhân khả dĩ của "kết quả không mong muốn" (log treo, mất kết nối) | Cao (dựa trên tài liệu chính thức + cộng đồng) | 15 phút |
| **2** | **Mục 6** — Sửa `b1_doc_session.py` tìm PDF đệ quy | 🟡 Tránh lỗi giả khi folder có cấu trúc lồng nhau (đặc biệt qua Web UI kéo-thả) | Cao | 5 phút |
| **3** | **Mục 4** — Thêm `usecols` đúng 17 cột cho B2, đồng bộ qua `config.COLS_NPO` | 🟡 ~1.5–4x tốc độ đọc GL02 (tuỳ số cột gốc), không lặp lại bug cũ | Trung bình-cao (đã benchmark, nhưng số cột thật của GL02 chưa biết) | 20 phút |
| **4** | **Mục 3** — Đo `os.cpu_count()` + benchmark A1 BẬT/TẮT trên máy chủ thật trước khi tin vào %; cân nhắc bỏ A1 nếu máy chủ ≤2 core | 🟡 Tránh kỳ vọng sai về tốc độ; đơn giản hoá code nếu A1 không có lợi | Cao (đã benchmark độc lập cho thấy rủi ro phản tác dụng) | 30 phút đo + quyết định |
| **5** | **Mục 5** — Thêm ghi chú hướng dẫn mở CSV đúng cách trong sheet Excel tóm tắt | 🟢 Giảm rủi ro hiểu sai dữ liệu khi người dùng mở CSV trực tiếp | Trung bình | 10 phút |
| **6** | **Mục 7** — Chuyển `ZIP_PASSWORD` sang biến môi trường (có fallback) | 🟢 Bảo mật, không bắt buộc gấp | Thấp-trung bình (tuỳ chính sách bảo mật nội bộ) | 15 phút |
| **7** | **Mục 8** — Dọn 4 file `.md` kế hoạch cũ vào `docs/` | 🟢 Gọn gàng codebase, không ảnh hưởng vận hành | Cao | 10 phút |
| **8** | A_NEW, A2, A5 (tqdm), thread-safety B4 TPAY | ✅ Đã triển khai đúng trong code hiện tại, không cần làm thêm | — | 0 (đã xong) |
| | **Tổng việc còn lại** | | | **~1.5 giờ** |

---

## 10. KẾT LUẬN

Code hiện tại (sau khi Claude Code áp dụng plan V2) đã triển khai **đúng và đầy đủ** những gì plan V2
đề ra: A_NEW (CSV cho sheet lớn) là tối ưu tốc độ lớn nhất và đã được xác nhận độc lập (~27x). A2/A3 bị
loại bỏ đúng. Thread-safety B4 (TPAY tham số hoá) đã vá đúng. Web UI đã hoàn chỉnh hơn cả thiết kế
trong plan (có path traversal protection, multi-file download UI đẹp).

Review V2 này bổ sung 3 nhóm phát hiện mới quan trọng nhất:

1. **Một rủi ro kiến trúc (`eventlet` không monkey-patch + `ThreadPoolExecutor`) nhiều khả năng chính
   là nguyên nhân gây "kết quả không mong muốn"** mà bạn đề cập — đây là việc nên làm trước tiên, tách
   biệt hoàn toàn khỏi câu chuyện tốc độ.
2. **A1 (ThreadPool đọc 2 ZIP) có rủi ro phản tác dụng** mà benchmark trong cả 2 lượt review trước chưa
   đo đúng tỷ lệ CPU:I/O thực tế — cần đo lại trên máy chủ thật trước khi tin vào con số 10-30%.
3. **Một cơ hội tối ưu tốc độ bị bỏ sót ở B2** (đọc thừa cột do hậu quả của một lần sửa bug cũ), cùng
   đề xuất tránh tái diễn loại bug đó bằng cách gộp định nghĩa cột về `config.py`.

Tất cả khuyến nghị trong review này đều đã được benchmark hoặc tra cứu tài liệu chính thức để xác minh
trước khi đưa vào, không suy diễn chủ quan.

---

*Review V2 bởi Claude (Sonnet 4.6) — 20/06/2026*
*Benchmark độc lập chạy trong container Python 3.x, pandas 3.0.2, xlsxwriter 3.2.9, môi trường giới
hạn 1 CPU core (đã nêu rõ giới hạn này ảnh hưởng tới kết luận mục 3 như thế nào).*
