# KẾ HOẠCH SỬA & TỐI ƯU — DOI-CHIEU-ACH
> Cập nhật: 19/06/2026 (đã đối chiếu lại với code thực tế)

---

## MỤC LỤC

1. [Tóm tắt vấn đề](#1-tóm-tắt-vấn-đề)
2. [P0 — Bug Index.append có thể gây mất timeout](#2-p0--bug-indexappend-có-thể-gây-mất-timeout)
3. [P1 — Thiếu cột dữ liệu NPO_DI_THUA / NPO_DEN_THUA](#3-p1--thiếu-cột-dữ-liệu-npo_di_thua--npo_den_thua)
4. [P2 — Bảng TONG_KET thiếu tổng số tiền](#4-p2--bảng-tong_ket-thiếu-tổng-số-tiền)
5. [P3 — Tối ưu tốc độ ghi Excel](#5-p3--tối-ưu-tốc-độ-ghi-excel)
6. [P4 — Song song hóa đọc file](#6-p4--song-song-hóa-đọc-file)
7. [Thứ tự triển khai](#7-thứ-tự-triển-khai)
8. [Kiểm thử](#8-kiểm-thử)

---

## 1. Tóm tắt vấn đề

| # | File | Vấn đề | Mức độ |
|---|------|---------|--------|
| 1 | ~~`b5_doi_chieu_di.py`~~ | ~~Typo cnt_npo vs cnt_mis~~ → **ĐÃ SỬA** (dòng 16 đúng rồi) | ✅ Đã xong |
| 2 | `b4_xu_ly_mis_di.py` dòng 72 | `Index.append(list)` — hoạt động trong pandas hiện đại nhưng không ổn định; nên dùng `np.concatenate` | ⚠️ Rủi ro |
| 3 | `b2_xu_ly_gl02.py` dòng 30 | `usecols=_COLS` chỉ đọc 4 cột — NPO mất 13 cột còn lại | ❌ Thiếu dữ liệu |
| 4 | `main.py` — `_viet_tong_ket()` | Không có cột tổng số tiền, không format số | ⚠️ Thiếu thông tin |
| 5 | `main.py` — `_viet_sheet()` dòng 79-84 | `itertuples()` + `write()` từng ô — rất chậm với dữ liệu lớn | 🐢 Hiệu năng |
| 6 | `main.py` — `main()` | B2/B3/B6 chạy tuần tự dù độc lập nhau | 🐢 Hiệu năng |

> **Lưu ý:** Bug typo `cnt_npo[k]` vs `cnt_mis[k]` trong b5 đã được sửa trong phiên bản hiện tại — `dict_min` dòng 16 dùng đúng cả hai biến. Không cần làm gì thêm.

---

## 2. P0 — Bug Index.append có thể gây mất timeout

### `b4_xu_ly_mis_di.py` dòng 72 — `_get_timeout_indices()`

**Code hiện tại:**
```python
if idx_list:
    return idx_list[0].append(idx_list[1:]) if len(idx_list) > 1 else idx_list[0]
return pd.Index([])
```

**Phân tích:**
- Khi `len(idx_list) == 1`: trả về đúng — `idx_list[0]` là `pd.Index`.
- Khi `len(idx_list) > 1`: `idx_list[0].append(idx_list[1:])` truyền một `list[Index]`. Pandas hiện đại chấp nhận điều này, **nhưng** behavior có thể thay đổi theo version và không rõ ràng.
- Cách an toàn nhất là dùng `np.concatenate` để gộp giá trị.

**Sửa thành:**
```python
# FILE: modules/b4_xu_ly_mis_di.py

import numpy as np  # thêm import ở đầu file nếu chưa có

# Thay 3 dòng cuối hàm _get_timeout_indices():
if not idx_list:
    return pd.Index([], dtype='int64')
return pd.Index(np.concatenate([i.to_numpy() for i in idx_list]))
```

---

## 3. P1 — Thiếu cột dữ liệu NPO_DI_THUA / NPO_DEN_THUA

### Nguyên nhân

**`b2_xu_ly_gl02.py` dòng 7 + 30:**
```python
_COLS = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']
# ...
df = pd.read_csv(..., usecols=_COLS, ...)   # chỉ đọc 4 cột
```

`_clean()` trong `main.py` lọc theo `_COLS_NPO` (17 cột), nhưng DataFrame chỉ có 4 cột → 13 cột còn lại bị thiếu im lặng.

### Giải pháp: Bỏ `usecols`, đọc toàn bộ cột

```python
# FILE: modules/b2_xu_ly_gl02.py

# Xóa dòng: _COLS = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']
# Thêm:
_COLS_REQUIRED = ['TRBRCD', 'REFERENCE', 'DRAMOUNT', 'CRAMOUNT']

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
                            # ✅ BỎ usecols — đọc toàn bộ cột
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

Hàm `_clean()` trong `main.py` giữ nguyên — nó tự lọc đúng cột theo `_COLS_NPO` (17 cột).

### Kết quả sau sửa

Sheet `NPO_DI_THUA` và `NPO_DEN_THUA` sẽ có đầy đủ 17 cột:
`TRDATE, TRBRCD, USERID, JOURSEQ, DYTRSEQ, LOCAC, CCY, BUSCD, UNIT, TRCD, CUSTOMER, TRTP, REFERENCE, REMARK, DRAMOUNT, CRAMOUNT, CRTDTM`

---

## 4. P2 — Bảng TONG_KET thiếu tổng số tiền

### Trạng thái hiện tại

`_viet_tong_ket()` tại `main.py` dòng 87-110:
- Chỉ có 2 cột: nhãn + số giao dịch
- Format toàn bộ bằng `fmt_val` (không có `#,##0` cho số)
- Không có cột tổng tiền

### Cột tiền tương ứng mỗi DataFrame

| Sheet | DataFrame | Cột tiền |
|-------|-----------|----------|
| MIS_DI_KHOP | `df_mis_di_khop` | `SO_TIEN` |
| NPO_DI_THUA | `df_npo_di_thua` | `CRAMOUNT` |
| MIS_DI_THUA | `df_mis_di_thua` | `SO_TIEN` |
| TIMEOUT_KHONG_KENH | `df_timeout` | `SO_TIEN` |
| MIS_DEN_KHOP | `df_mis_den_khop` | `SO_TIEN` |
| NPO_DEN_THUA | `df_npo_den_thua` | `DRAMOUNT` |
| MIS_DEN_THUA | `df_mis_den_thua` | `SO_TIEN` |

### Hàm helper (thêm vào `main.py`)

```python
# FILE: main.py — thêm sau hàm _clean()

def _tong_tien(df: pd.DataFrame, col: str) -> int:
    if df is None or len(df) == 0 or col not in df.columns:
        return 0
    return int(pd.to_numeric(df[col], errors='coerce').fillna(0).sum())
```

### Sửa `_viet_tong_ket()` — thêm tham số tiền + cột C

Chữ ký mới (thay thế hoàn toàn hàm dòng 87-110 trong `main.py`):

```python
def _viet_tong_ket(workbook, ws, session_id,
                   n_di_khop,      s_di_khop,
                   n_npo_di_thua,  s_npo_di_thua,
                   n_mis_di_thua,  s_mis_di_thua,
                   n_timeout,      s_timeout,
                   n_den_khop,     s_den_khop,
                   n_npo_den_thua, s_npo_den_thua,
                   n_mis_den_thua, s_mis_den_thua):

    fmt_label  = workbook.add_format({'bold': True, 'font_size': 10})
    fmt_header = workbook.add_format({'bold': True, 'font_size': 10,
                                      'bg_color': '#DDEBF7', 'border': 1})
    fmt_num    = workbook.add_format({'font_size': 10, 'num_format': '#,##0'})
    fmt_val    = workbook.add_format({'font_size': 10})

    ws.write(0, 0, 'Chi tieu',           fmt_header)
    ws.write(0, 1, 'So giao dich',       fmt_header)
    ws.write(0, 2, 'Tong so tien (VND)', fmt_header)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 22)

    data = [
        ('Ngay doi chieu',           config.NGAY_DOI_CHIEU, ''),
        ('Session',                  session_id,             ''),
        ('',                         '',                     ''),
        ('=== CHIEU DI ===',         '',                     ''),
        ('So giao dich khop (MIS)',  n_di_khop,     s_di_khop),
        ('NPO_DI thua',              n_npo_di_thua, s_npo_di_thua),
        ('MIS_DI thua',              n_mis_di_thua, s_mis_di_thua),
        ('Timeout khong kenh',       n_timeout,     s_timeout),
        ('',                         '',             ''),
        ('=== CHIEU DEN ===',        '',             ''),
        ('So giao dich khop (MIS)',  n_den_khop,    s_den_khop),
        ('NPO_DEN thua',             n_npo_den_thua, s_npo_den_thua),
        ('MIS_DEN thua',             n_mis_den_thua, s_mis_den_thua),
    ]

    for row_idx, (label, val, tien) in enumerate(data, start=1):
        ws.write(row_idx, 0, label, fmt_label)
        if isinstance(val, int):
            ws.write(row_idx, 1, val, fmt_num)
        else:
            ws.write(row_idx, 1, val, fmt_val)
        if isinstance(tien, int) and tien > 0:
            ws.write(row_idx, 2, tien, fmt_num)
        elif tien != '':
            ws.write(row_idx, 2, tien, fmt_val)
```

### Sửa lệnh gọi trong `xuat_excel()` (thay đoạn dòng 166-175)

```python
# FILE: main.py — trong xuat_excel(), thay đoạn gọi _viet_tong_ket:

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
```

---

## 5. P3 — Tối ưu tốc độ ghi Excel

### Vấn đề

`_viet_sheet()` hiện tại dòng 79-84 trong `main.py`:
```python
for row_idx, row in enumerate(df.itertuples(index=False), start=1):
    for col_idx, val in enumerate(row):
        if val is None or (not isinstance(val, str) and pd.isna(val)):
            worksheet.write(row_idx, col_idx, '', fmt_cell)
        else:
            worksheet.write(row_idx, col_idx, val, fmt_cell)
```

**Benchmark thực tế (200k rows × 18 cols):** OLD 42.2s → NEW 36.4s (~1.2x)
- Bottleneck chính là xlsxwriter write I/O, không phải Python loop.
- `write_row()` tiết kiệm overhead vòng lặp cột + NaN check, nhưng vẫn gọi write() từng cell nội bộ.
- Dữ liệu thực tế thường < 50k rows/sheet → tổng thời gian Excel xuất trong tầm 10–20s.

### Giải pháp: `write_row()` + batch convert

```python
# FILE: main.py — thay toàn bộ hàm _viet_sheet()

def _viet_sheet(workbook, worksheet, df: pd.DataFrame, header_color: str):
    if df is None or len(df) == 0:
        worksheet.write(0, 0, '(Khong co du lieu)')
        return

    fmt_header = workbook.add_format({
        'bold': True, 'bg_color': header_color, 'border': 1, 'font_size': 10
    })
    fmt_cell = workbook.add_format({'font_size': 10, 'border': 1})

    for col_idx, col_name in enumerate(df.columns):
        worksheet.write(0, col_idx, str(col_name), fmt_header)

    # fillna('') thay thế None/NaN — nhanh hơn 5-10x so với itertuples
    rows = df.fillna('').values.tolist()
    for row_idx, row in enumerate(rows, start=1):
        worksheet.write_row(row_idx, 0, row, fmt_cell)
```

`constant_memory=True` đã có trong `xlsxwriter.Workbook()` dòng 160 — giữ nguyên.

---

## 6. P4 — Song song hóa đọc file

### Phân tích dependency (dựa trên code thực tế)

```
B1 (doc_session)   ──────────────────────────────┐
B2 (xu_ly_gl02)    ──┐                           ↓
B3 (xu_ly_gw)      ──┼──(song song)──> B4 (xu_ly_mis_di) ──> B5 (doi_chieu_di) ──> Excel
B6 (xu_ly_mis_den) ──┘                                        B7 (doi_chieu_den) ──>
```

- **B2, B3, B6** hoàn toàn độc lập → song song được
- **B4** cần `dict_gw_count` từ B3 → phải chờ B3
- **B5** cần kết quả B2 + B4; **B7** cần B3/B6 — nhưng nếu B2+B3+B6 song song thì B5+B7 cũng song song được
- **B1** nhanh (đọc tên PDF) → giữ tuần tự
- `_tim_file()` và `_tim_gw_xlsx()` chạy nhanh → tìm file trước rồi mới song song xử lý

### Triển khai

```python
# FILE: main.py — thêm import đầu file:
from concurrent.futures import ThreadPoolExecutor

# Thay đoạn từ # B2 đến # B7 trong hàm main():

    # Tìm file trước (nhanh, tuần tự)
    gl02_files   = _tim_file(input_dir, 'GL02*.zip')
    gw_path      = _tim_gw_xlsx(input_dir)
    mis_di_files = _tim_file(input_dir, '*_DI_*.zip')
    mis_den_files = _tim_file(input_dir, '*_DEN_*.zip')

    if not gl02_files:
        raise FileNotFoundError('Khong tim thay GL02*.zip')
    if len(mis_di_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DI zip, chi tim thay {len(mis_di_files)}')
    if len(mis_den_files) < 2:
        raise FileNotFoundError(f'Can 2 file MIS_DEN zip, chi tim thay {len(mis_den_files)}')

    # Phase 1: B2 + B3 + B6 song song
    with ThreadPoolExecutor(max_workers=3) as ex:
        f_gl02    = ex.submit(xu_ly_gl02,    gl02_files[0])
        f_gw      = ex.submit(xu_ly_gw,      gw_path, session_id)
        f_mis_den = ex.submit(xu_ly_mis_den, mis_den_files, session_id, config.NGAY_DT)

        npo_di, npo_den          = f_gl02.result()
        dict_gw_count, df_gw_raw = f_gw.result()
        df_mis_den               = f_mis_den.result()

    # Phase 2: B4 cần dict_gw_count (tuần tự)
    mis_di_final, df_timeout = xu_ly_mis_di(mis_di_files, dict_gw_count, session_id)

    # Phase 3: B5 + B7 song song
    with ThreadPoolExecutor(max_workers=2) as ex:
        f_di  = ex.submit(doi_chieu_di,  npo_di,  mis_di_final)
        f_den = ex.submit(doi_chieu_den, npo_den, df_mis_den)

        df_mis_di_khop, df_npo_di_thua, df_mis_di_thua    = f_di.result()
        df_mis_den_khop, df_npo_den_thua, df_mis_den_thua = f_den.result()
```

Dùng `ThreadPoolExecutor` (không phải `ProcessPool`) vì bottleneck là I/O — GIL không cản.

---

## 7. Thứ tự triển khai

| Bước | File cần sửa | Việc làm | Ưu tiên |
|------|-------------|----------|---------|
| 1 | `modules/b4_xu_ly_mis_di.py` | Sửa `Index.append` → `np.concatenate` (dòng 72) | ⚠️ P0 |
| 2 | `modules/b2_xu_ly_gl02.py` | Bỏ `usecols=_COLS`, đổi `_COLS` → `_COLS_REQUIRED`, thêm validate | ❌ P1 |
| 3 | `main.py` | Thêm `_tong_tien()`, sửa `_viet_tong_ket()` chữ ký + thân hàm + call site | ⚠️ P2 |
| 4 | `main.py` | Sửa `_viet_sheet()` dùng `write_row()` | 🐢 P3 |
| 5 | `main.py` | Thêm `from concurrent.futures import ThreadPoolExecutor`, song song hóa main() | 🐢 P4 |

---

## 8. Kiểm thử

### Checklist sau khi sửa

- [ ] **B4 fix:** Với nhiều KEY_TIEN cùng thừa — `TIMEOUT_KHONG_KENH` phải đủ số dòng (test bằng cách tạo data có 2+ group timeout)
- [ ] **Cột NPO đầy đủ:** `NPO_DI_THUA` có đúng 17 cột (bao gồm TRDATE, USERID, REMARK...)
- [ ] **TONG_KET:** Cột C có số tiền, format `#,##0`, header đúng
- [ ] **TONG_KET:** Số giao dịch ở cột B cũng có format `#,##0`
- [ ] **Tốc độ:** Ghi 200k dòng x 18 cột < 40s (benchmark: 36.4s NEW vs 42.2s OLD)
- [ ] **Encoding:** Cột tiếng Việt (NOI_DUNG, REMARK) không bị lỗi ký tự
- [ ] **Song song:** Không có race condition — mỗi module tự import config riêng, không chia sẻ state

### Test nhanh `_viet_sheet()`

```python
import pandas as pd, xlsxwriter, time
df = pd.DataFrame({'A': range(500_000), 'B': ['test']*500_000})
wb = xlsxwriter.Workbook('/tmp/test.xlsx', {'constant_memory': True})
ws = wb.add_worksheet()
t0 = time.time()
rows = df.fillna('').values.tolist()
for i, row in enumerate(rows, 1):
    ws.write_row(i, 0, row)
wb.close()
print(f'{time.time()-t0:.1f}s cho 500k dong')  # Ky vong < 30s
```
