import re
import glob
import os


def doc_session(input_dir: str) -> str:
    """
    Tim file PDF trong input_dir, lay so session tu ten file.
    Ten file dang: ACH_20260612_VBAAVNVN_NRT_15882_N03_1.pdf
    """
    pattern = os.path.join(input_dir, '*.pdf')
    pdfs = glob.glob(pattern)
    if not pdfs:
        raise FileNotFoundError(f'Khong tim thay file PDF trong {input_dir}')

    pdf_path = pdfs[0]
    ten_file = os.path.basename(pdf_path)
    m = re.search(r'_NRT_(\d+)_', ten_file)
    if not m:
        raise ValueError(f'Khong the lay session tu ten file: {ten_file}')

    session_id = m.group(1)
    print(f'[B1] Session: {session_id}  (tu file: {ten_file})')
    return session_id
