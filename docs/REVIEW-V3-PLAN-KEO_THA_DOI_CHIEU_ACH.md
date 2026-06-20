# REVIEW V3 & ĐỀ XUẤT TỐI ƯU — DOI-CHIEU-ACH
> Người review: Claude (Sonnet 4.6)
> Ngày: 20/06/2026
> Phương pháp: Đối chiếu `PLAN-KEO_THA_DOI_CHIEU_ACH.md` (v3, do Claude Code viết lại sau khi đọc
> REVIEW-V2) với **mã nguồn thực tế** trong `DOI-CHIEU-ACH-main__2_.zip` đính kèm cùng lần upload này
> — bằng `diff` từng dòng, không chỉ đọc văn bản kế hoạch. Đồng thời kiểm thử lại độ an toàn của các
> đề xuất mới trong UPGRADE C (đặc biệt C3) bằng các test case biên (edge case).

---

## TÓM TẮT NHANH — KẾT LUẬN QUAN TRỌNG NHẤT CỦA REVIEW NÀY

> ## ⚠️ PHÁT HIỆN SỐ 1: PLAN V3 CHỈ LÀ VĂN BẢN — CODE THỰC TẾ CHƯA THAY ĐỔI MỘT DÒNG NÀO
>
> Tôi đã giải nén lại `DOI-CHIEU-ACH-main__2_.zip` đính kèm lần này và `diff` từng byte với codebase đã
> đọc ở lần review trước (V2). **Kết quả: 100% giống hệt nhau** — không có file code nào (`main.py`,
> `web_app.py`, `config.py`, `modules/*.py`, `templates/index.html`) bị thay đổi. Cụ thể đã xác minh
> trực tiếp bằng `grep`:
>
> | Hạng mục trong plan v3 ghi "✅ ĐÃ TRIỂN KHAI" hoặc cần sửa | Trạng thái CODE THỰC TẾ |
> |---|---|
> | C1 — `web_app.py` dùng `async_mode='threading'` | ❌ **VẪN dùng `async_mode='eventlet'`**, `requirements.txt` vẫn còn `eventlet>=0.35` |
> | C2 — `b1_doc_session.py` tìm PDF đệ quy | ❌ **VẪN dùng `glob.glob(pattern)` không đệ quy** |
> | C3 — `config.py` có `COLS_NPO`, `b2_xu_ly_gl02.py` có `usecols` | ❌ **`config.py` chưa có `COLS_NPO`; `b2_xu_ly_gl02.py` chưa có `usecols`** |
> | C4 — Ghi chú "mở CSV đúng cách" trong `main.py` | ❌ **Chưa có 2 dòng cảnh báo** |
> | C5 — `ZIP_PASSWORD` đọc từ biến môi trường | ❌ **VẪN là `ZIP_PASSWORD = b'DACwLdHi'` plaintext** |
> | C6 — Gom file `.md` vào `docs/` | ❌ **Chưa tạo thư mục `docs/`, 4 file `.md` vẫn ở gốc** |
>
> **Điều này có nghĩa là rủi ro nghiêm trọng nhất — vấn đề `eventlet` có thể gây treo Web UI khi xử lý
> dữ liệu lớn — VẪN CÒN NGUYÊN trong code đang chạy thực tế**, dù plan đã mô tả đúng cách sửa. Một kế
> hoạch đúng trên giấy không tự động trở thành một sửa chữa trong sản phẩm. Đây là việc quan trọng nhất
> cần làm ngay, không phải việc tiếp theo cần lên kế hoạch thêm.

| # | Đánh giá | Kết luận |
|---|---|---|
| Chất lượng lập luận kỹ thuật trong plan v3 | ✅ Tốt — tích hợp đúng và đầy đủ toàn bộ phát hiện của REVIEW-V2, không có claim sai mới | Giữ nguyên |
| C1-C6 (nội dung đề xuất) | ✅ Đúng kỹ thuật, đã kiểm tra lại độc lập, không phát hiện rủi ro mới | Thông qua, sẵn sàng code |
| **Trạng thái triển khai thực tế** | ❌ **0/6 hạng mục UPGRADE C đã được áp dụng vào code** | **Cần code ngay, không cần lên kế hoạch thêm** |
| C3 (usecols cho B2) — kiểm tra sâu thêm | ✅ Xác nhận an toàn qua test case biên (4 cột bắt buộc nằm trong 17 cột; encoding lỗi vẫn raise đúng exception để fallback) | Thông qua |
| Vấn đề mới phát hiện trong review này | 1 vấn đề nhỏ (mục 3) — thiếu test case cho `usecols` + encoding fallback trong checklist | Bổ sung khuyến nghị |

---

## 1. XÁC MINH ĐỘC LẬP: CODE THỰC TẾ VS. PLAN V3

### 1.1 Phương pháp xác minh

```bash
# So sánh toàn bộ source code trong zip mới với codebase đã review ở V2
diff -rq DOI-CHIEU-ACH-main__2_.zip(giải nén) codebase_đã_đọc_ở_V2/
→ Kết quả: "Only in .../V2: REVIEW-V2-PLAN-KEO_THA_DOI_CHIEU_ACH.md"
→ Nghĩa là: KHÔNG có file code nào khác biệt. Zip lần này = zip lần trước, y hệt.
```

### 1.2 Xác minh trực tiếp từng claim bằng `grep` trên code thật

```bash
$ grep -n "async_mode" web_app.py
web_app.py:21:    async_mode='eventlet',

$ grep -n "recursive" modules/b1_doc_session.py
(không có kết quả — glob.glob() vẫn không đệ quy)

$ grep -n "COLS_NPO" config.py
(không có kết quả)

$ grep -n "usecols" modules/b2_xu_ly_gl02.py
(không có kết quả)

$ grep -n "LUU Y" main.py
(không có kết quả)

$ grep -n "ZIP_PASSWORD" config.py
config.py:3:ZIP_PASSWORD   = b'DACwLdHi'
```

Tất cả đều xác nhận: **đây vẫn là codebase ở trạng thái sau khi áp dụng plan V2 (A_NEW, A1, A_CLEAN,
A5, Web UI cơ bản), chưa áp dụng bất kỳ mục nào của UPGRADE C (plan V3).**

### 1.3 Vì sao điều này quan trọng hơn là một ghi chú thủ tục

Bạn đã hỏi 3 lần liên tiếp về "kết quả không mong muốn" và "tối ưu tốc độ". REVIEW-V2 xác định nguyên
nhân khả dĩ nhất của "kết quả không mong muốn" là tổ hợp `eventlet` (chưa monkey-patch) +
`ThreadPoolExecutor` (CPU-bound) trong `web_app.py`. Plan V3 mô tả đúng cách sửa (C1) — nhưng **nếu chỉ
dừng ở việc viết thêm một bản kế hoạch mà không sửa code, vấn đề gốc rễ vẫn còn nguyên trong sản phẩm
đang chạy**. Khuyến nghị rõ ràng: bước tiếp theo nên là **yêu cầu sửa trực tiếp 6 file đã liệt kê ở
mục 7 của plan v3**, không phải viết thêm một vòng kế hoạch/review nữa.

---

## 2. RÀ SOÁT KỸ THUẬT NỘI DUNG UPGRADE C (C1–C6) — KHÔNG PHÁT HIỆN SAI SÓT MỚI

Tôi đọc lại toàn bộ nội dung kỹ thuật của 6 đề xuất trong UPGRADE C bằng con mắt phản biện (không giả
định plan đúng), tìm sai sót logic hoặc rủi ro mới. Kết luận: **các đề xuất C1, C2, C4, C5, C6 đúng và
an toàn, không có gì cần sửa thêm.** Riêng C3 tôi kiểm thử sâu hơn — xem mục 3.

### 2.1 C1 — `eventlet` → `threading`: đúng

Đã tự kiểm chứng độc lập trong REVIEW-V2 qua tài liệu chính thức Flask-SocketIO/eventlet. Plan v3 trích
dẫn đúng, code đề xuất đúng cú pháp (`async_mode='threading'`, xoá `eventlet>=0.35` khỏi
`requirements.txt`). Không có gì cần sửa.

### 2.2 C2 — B1 tìm PDF đệ quy: đúng, nhất quán với `main.py._tim_file()`

```python
# Đề xuất trong plan, đối chiếu với main.py._tim_file() đang dùng:
abs_dir = os.path.abspath(input_dir)
return sorted(glob.glob(os.path.join(abs_dir, '**', pattern), recursive=True))
```
Cách viết trong C2 (`os.path.join(os.path.abspath(input_dir), '**', '*.pdf')`,
`recursive=True`) khớp 100% với pattern đã dùng trong `_tim_file()` — đúng tinh thần "nhất quán hành vi
tìm kiếm" mà REVIEW-V2 đề xuất. Không có gì cần sửa.

### 2.3 C4 — Ghi chú CSV: đúng, nội dung rõ ràng

Hai dòng cảnh báo ngắn gọn, đúng tiếng Việt không dấu (nhất quán với toàn bộ phần còn lại của
`main.py` vốn không dùng dấu trong log/print). Không có gì cần sửa.

### 2.4 C5 — `ZIP_PASSWORD` qua biến môi trường: đúng, có fallback an toàn

```python
ZIP_PASSWORD = os.environ.get('DOI_CHIEU_ZIP_PASSWORD', 'DACwLdHi').encode()
```
Giữ được khả năng chạy ngay không cần cấu hình gì thêm (CLI cũ vẫn hoạt động), đồng thời cho phép
override khi cần. Đúng nguyên tắc "không phá vỡ tương thích ngược" đã áp dụng nhất quán trong toàn bộ
dự án (giống cách `tpay_tu`/`tpay_den` có fallback `None` trong B4). Không có gì cần sửa.

### 2.5 C6 — Gom file `.md` vào `docs/`: đúng, không ảnh hưởng vận hành

Đã tự xác minh lại: `START.bat` và `START_WEB.bat` không tham chiếu bất kỳ file `.md` nào (chỉ gọi
`python main.py` / `python web_app.py`), nên việc di chuyển file tài liệu an toàn tuyệt đối. Không có
gì cần sửa.

---

## 3. KIỂM THỬ SÂU C3 — `usecols` CHO B2 + GỘP `COLS_NPO` VỀ `config.py`

Đây là đề xuất duy nhất có khả năng ảnh hưởng tới **tính đúng đắn của dữ liệu** (không chỉ tốc độ), nên
tôi kiểm thử kỹ hơn 2 kịch bản biên mà plan v3 chưa đề cập.

### 3.1 Kịch bản 1 — Liệu 4 cột bắt buộc (`_COLS_REQUIRED`) có nằm trong 17 cột (`_COLS_NPO`) không?

Nếu thêm `usecols=lambda c: c in _COLS_NPO` mà vô tình một trong 4 cột `_COLS_REQUIRED` (`TRBRCD`,
`REFERENCE`, `DRAMOUNT`, `CRAMOUNT`) không nằm trong `_COLS_NPO`, đoạn code validate
`missing = [c for c in _COLS_REQUIRED if c not in df.columns]` sẽ luôn `raise ValueError` — làm gãy
toàn bộ chương trình ngay cả với dữ liệu hợp lệ.

```python
_COLS_REQUIRED = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']
_COLS_NPO = ['TRDATE','TRBRCD','USERID','JOURSEQ','DYTRSEQ','LOCAC','CCY','BUSCD','UNIT',
             'TRCD','CUSTOMER','TRTP','REFERENCE','REMARK','DRAMOUNT','CRAMOUNT','CRTDTM']
missing = [c for c in _COLS_REQUIRED if c not in _COLS_NPO]
# → [] (rỗng) — an toàn, cả 4 cột đều có trong danh sách 17 cột
```
**Kết quả: AN TOÀN.** Cũng đã đối chiếu lại bằng cách liệt kê toàn bộ `df['...']` được dùng trong suốt
`b2_xu_ly_gl02.py` (`CRAMOUNT`, `DRAMOUNT`, `REFERENCE`, `TRBRCD`, cộng cột phái sinh `SO_TRACE` được
tạo sau khi đọc nên không bị ảnh hưởng) — toàn bộ đều nằm trong `_COLS_NPO`. Không thiếu cột nào.

### 3.2 Kịch bản 2 — `usecols` + vòng lặp thử encoding (`utf-8-sig` → `cp1252`) có xung đột không?

`_doc_zip()` hiện thử đọc bằng `utf-8-sig` trước, nếu `UnicodeDecodeError` thì thử `cp1252`. Câu hỏi:
nếu encoding sai khiến header bị đọc sai (mojibake) thay vì raise lỗi ngay, liệu `usecols` có "âm thầm"
trả về DataFrame rỗng cột mà không kích hoạt cơ chế fallback, khiến lỗi bị nuốt im lặng?

Đã kiểm thử trực tiếp 2 trường hợp:
```python
# TH1: usecols filter ra danh sách cột không khớp hoàn toàn (nhưng decode thành công)
# → pd.read_csv KHÔNG raise lỗi, chỉ trả về DataFrame với các cột khớp được (có thể 0 cột)
# → Nhưng bước validate "missing = [...]" NGAY SAU ĐÓ vẫn hoạt động đúng:
#   nếu DataFrame có 0 cột do header bị đọc sai, "missing" sẽ liệt kê đủ 4 cột _COLS_REQUIRED
#   → vẫn raise ValueError đúng như thiết kế ban đầu, KHÔNG có silent failure.

# TH2: encoding sai nghiêm trọng (binary không decode được)
# → pd.read_csv raise UnicodeDecodeError NGAY LẬP TỨC, trước khi usecols kịp xử lý gì
# → vòng lặp for-enc vẫn fallback sang cp1252 đúng như cũ, KHÔNG bị ảnh hưởng bởi usecols
```
**Kết quả: AN TOÀN.** Cơ chế validate `missing = [...]` đang đóng vai trò lưới an toàn cuối cùng, hoạt
động độc lập với việc có `usecols` hay không. Việc thêm `usecols` không tạo ra đường nào để lỗi dữ liệu
bị bỏ qua một cách im lặng.

### 3.3 Khuyến nghị bổ sung (nhỏ) — Thêm 2 dòng vào checklist kiểm thử của plan v3

Plan v3 mục 8 đã có dòng kiểm thử *"NPO_DI_THUA và NPO_DEN_THUA vẫn đủ 17 cột sau khi thêm usecols"* —
đúng hướng nhưng nên cụ thể hoá thêm 2 trường hợp biên đã kiểm chứng ở trên, để người thực thi (hoặc
Claude Code ở lượt sau) không cần tự suy luận lại:

```
- [ ] C3: Test với file GL02 mẫu có ít hơn 17 cột (một số cột trong _COLS_NPO không tồn tại)
      → xác nhận chương trình vẫn chạy đúng cho các cột có sẵn, KHÔNG raise lỗi giả
      (vì usecols chỉ lọc cột TỒN TẠI, không bắt buộc phải đủ cả 17 cột — chỉ 4 cột
      _COLS_REQUIRED mới bắt buộc phải có)
- [ ] C3: Test với file GL02 cố tình sai encoding hoàn toàn (vd: mở bằng latin1 nhưng file
      thực chất utf-8 có ký tự tiếng Việt) → xác nhận vẫn raise UnicodeDecodeError và
      fallback sang cp1252 đúng như hành vi hiện tại, không bị usecols che giấu lỗi
```

---

## 4. RÀ SOÁT BỔ SUNG — CÓ ĐIỂM NÀO KHÁC CHƯA ĐƯỢC PLAN V3 NHẮC TỚI KHÔNG?

Tôi rà lại toàn bộ codebase một lần nữa (không chỉ những điểm đã nêu ở V1/V2) để tìm xem còn sót gì.
Kết luận: **không phát hiện vấn đề mới nào đáng kể.** Các module B3, B5, B7 vẫn giữ nguyên logic đã
được xác nhận đúng (vectorized `groupby`/`cumcount`, không có vòng lặp Python theo từng dòng). Cấu trúc
thư mục ngoài 5 file `.md` (đã nêu ở C6) không có file rác (`__pycache__`, file debug cũ) — đúng như đã
xác nhận ở REVIEW-V2, vẫn đúng ở thời điểm này.

Một quan sát nhỏ, không phải lỗi nhưng đáng ghi nhận cho minh bạch: `web_app.py` thu thập file kết quả
để trả về cho client bằng cách so khớp tên file theo điều kiện
`if base.replace('doi_chieu_', '') in fname` — đây là so khớp chuỗi con (substring), về lý thuyết nếu
2 job chạy cùng ngày nhưng khác `job_id` thì tên file CSV output (`MIS_DI_KHOP_{ngay_str}.csv`) sẽ
trùng nhau giữa các job, dẫn tới ghi đè lẫn nhau trên đĩa (file CSV/Excel output đặt tên theo
`output/doi_chieu_YYYYMMDD.xlsx`, không có `job_id` trong tên). Đây là rủi ro **chỉ xảy ra khi 2 người
dùng web cùng xử lý cùng một ngày đối chiếu cùng lúc** — một kịch bản hẹp, có khả năng xảy ra thấp với
quy mô vài người dùng nội bộ nhưng đáng để biết. **Không đề xuất sửa ngay** (độ phức tạp/lợi ích không
tương xứng với một dự án nội bộ quy mô nhỏ), chỉ ghi nhận để bạn cân nhắc nếu sau này mở rộng số người
dùng đồng thời.

---

## 5. KHUYẾN NGHỊ HÀNH ĐỘNG — RÚT GỌN, TẬP TRUNG VÀO THỰC THI

Khác với REVIEW-V2 (cần đề xuất nội dung sửa), nhiệm vụ chính của REVIEW-V3 này là **xác nhận kế hoạch
đã đúng và đẩy sang giai đoạn thực thi**. Không cần thêm một vòng lập kế hoạch nữa.

| Bước | Nội dung | Trạng thái | Việc cần làm tiếp |
|---|---|---|---|
| 1 | Nội dung kỹ thuật của UPGRADE C (C1-C6) | ✅ Đã rà soát kỹ, không có sai sót | Không cần sửa kế hoạch nữa |
| 2 | Áp dụng C1 vào `web_app.py` + `requirements.txt` | ❌ Chưa làm trong code | **Làm ngay — ưu tiên cao nhất** |
| 3 | Áp dụng C2 vào `modules/b1_doc_session.py` | ❌ Chưa làm trong code | Làm ngay (5 phút, rủi ro thấp) |
| 4 | Áp dụng C3 vào `config.py` + `modules/b2_xu_ly_gl02.py` + `main.py` | ❌ Chưa làm trong code | Làm ngay, đã xác nhận an toàn ở mục 3 |
| 5 | Áp dụng C4 vào `main.py` (ghi chú CSV) | ❌ Chưa làm trong code | Làm ngay (10 phút) |
| 6 | Áp dụng C5 vào `config.py` | ❌ Chưa làm trong code | Làm khi thuận tiện (không khẩn) |
| 7 | Áp dụng C6 (gom file `.md`) | ❌ Chưa làm | Làm khi thuận tiện (không khẩn) |
| 8 | Đo `os.cpu_count()` + benchmark A1 trên máy chủ thật | Chưa có số liệu thực tế từ máy chủ khách hàng | Cần thực hiện trên máy chủ thật, không thể làm thay trong môi trường review |
| 9 | Chạy checklist kiểm thử đầy đủ (đã có trong plan v3 mục 8, bổ sung 2 dòng ở mục 3.3 review này) | Chưa chạy | Sau khi code xong bước 2-7 |

**Đề xuất cụ thể cho bước tiếp theo:** nếu bạn muốn, hãy yêu cầu trực tiếp "hãy sửa code theo UPGRADE C"
(thay vì "hãy rà soát kế hoạch thêm lần nữa") — vì nội dung kỹ thuật đã được xác nhận đúng qua 2 vòng
review độc lập (V2 và V3), bước có giá trị nhất tiếp theo là biến kế hoạch thành code thật trong
`DOI-CHIEU-ACH-main/`, để vấn đề "kết quả không mong muốn" (rủi ro treo Web UI do `eventlet`) được khắc
phục trong sản phẩm đang chạy, không chỉ trên giấy.

---

## 6. KẾT LUẬN

Bản kế hoạch v3 (`PLAN-KEO_THA_DOI_CHIEU_ACH.md`) do Claude Code viết lại có chất lượng tốt: tích hợp
đúng và đầy đủ toàn bộ phát hiện từ REVIEW-V2, trình bày trung thực (đánh dấu rõ "❌ CẦN LÀM" thay vì
ngụy biện là đã xong), không có sai sót kỹ thuật mới. Sau khi tự kiểm thử sâu thêm đề xuất C3 (rủi ro
cao nhất về mặt dữ liệu trong số 6 đề xuất), xác nhận an toàn qua 2 kịch bản biên.

Phát hiện quan trọng nhất của lần review này không nằm ở nội dung kỹ thuật của kế hoạch — mà ở khoảng
cách giữa **kế hoạch đã viết đúng** và **mã nguồn thực tế chưa được sửa**. Tôi đã xác minh bằng `diff`
trực tiếp: 0 trong 6 hạng mục UPGRADE C đã được áp dụng vào code. Rủi ro nghiêm trọng nhất (`eventlet`
gây treo Web UI khi xử lý dữ liệu lớn) vẫn còn nguyên trong sản phẩm đang chạy.

**Khuyến nghị duy nhất và rõ ràng nhất của REVIEW-V3: chuyển từ giai đoạn "lập kế hoạch/review" sang
giai đoạn "thực thi code" cho 6 hạng mục UPGRADE C, bắt đầu từ C1.**

---

*Review V3 bởi Claude (Sonnet 4.6) — 20/06/2026*
*Phương pháp xác minh: `diff -rq` toàn bộ source code giữa 2 lần upload zip, `grep` trực tiếp từng
claim trong plan v3 đối chiếu với code thật, kiểm thử độc lập 2 kịch bản biên cho đề xuất C3.*
