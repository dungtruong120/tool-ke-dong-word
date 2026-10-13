import os
import re
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Cấu hình số dòng kẻ nét liền theo yêu cầu:
SO_DONG_TRAC_NGHIEM = 2       # Trắc nghiệm: 2 dòng
SO_DONG_DUNG_SAI = 6          # Đúng / Sai: 6 dòng
SO_DONG_TRA_LOI_NGAN = 8      # Trả lời ngắn: 8 dòng

def tao_bang_dong_ke(so_dong):
    """
    Tạo bảng gồm `so_dong` hàng:
    - Viền dưới nét đơn rõ nét màu xám #707070.
    - Chiều cao dòng ~8mm chuẩn tay viết học sinh.
    """
    rows_xml = ""
    for _ in range(so_dong):
        rows_xml += (
            '    <w:tr>\n'
            '      <w:trPr>\n'
            '        <w:trHeight w:val="440" w:hRule="atLeast"/>\n'  # Chiều cao mỗi dòng kẻ (~8mm)
            '      </w:trPr>\n'
            '      <w:tc>\n'
            '        <w:tcPr>\n'
            '          <w:tcW w:w="9360" w:type="dxa"/>\n'
            '          <w:tcBorders>\n'
            '            <w:top w:val="none"/>\n'
            '            <w:left w:val="none"/>\n'
            '            <w:bottom w:val="single" w:sz="8" w:space="0" w:color="707070"/>\n'
            '            <w:right w:val="none"/>\n'
            '          </w:tcBorders>\n'
            '          <w:tcMar>\n'
            '            <w:top w:w="0" w:type="dxa"/>\n'
            '            <w:bottom w:w="0" w:type="dxa"/>\n'
            '          </w:tcMar>\n'
            '        </w:tcPr>\n'
            '        <w:p>\n'
            '          <w:pPr>\n'
            '            <w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/>\n'
            '          </w:pPr>\n'
            '        </w:p>\n'
            '      </w:tc>\n'
            '    </w:tr>\n'
        )

    tbl_xml = (
        f'<w:tbl {nsdecls("w")}>\n'
        f'  <w:tblPr>\n'
        f'    <w:tblW w:w="9360" w:type="dxa"/>\n'
        f'    <w:tblCellMar>\n'
        f'      <w:top w:w="0" w:type="dxa"/>\n'
        f'      <w:bottom w:w="0" w:type="dxa"/>\n'
        f'      <w:left w:w="0" w:type="dxa"/>\n'
        f'      <w:right w:w="0" w:type="dxa"/>\n'
        f'    </w:tblCellMar>\n'
        f'    <w:tblBorders>\n'
        f'      <w:top w:val="none"/>\n'
        f'      <w:left w:val="none"/>\n'
        f'      <w:bottom w:val="none"/>\n'
        f'      <w:right w:val="none"/>\n'
        f'      <w:insideH w:val="none"/>\n'
        f'      <w:insideV w:val="none"/>\n'
        f'    </w:tblBorders>\n'
        f'  </w:tblPr>\n'
        f'  <w:tblGrid>\n'
        f'    <w:gridCol w:w="9360"/>\n'
        f'  </w:tblGrid>\n'
        f'{rows_xml}'
        f'</w:tbl>'
    )
    return parse_xml(tbl_xml)

def chen_dong_ke(target_element, so_dong):
    """
    Chèn bảng dòng kẻ và đệm khoảng cách hợp lý trên/dưới để không dính sát câu kế tiếp
    """
    tbl_elm = tao_bang_dong_ke(so_dong)
    target_element.addnext(tbl_elm)

    # Thêm 1 đoạn đệm nhỏ (khoảng cách sau 10pt) để tách rời dòng kẻ khỏi câu hỏi tiếp theo
    p_dem_duoi = parse_xml(
        f'<w:p {nsdecls("w")}>\n'
        f'  <w:pPr>\n'
        f'    <w:spacing w:before="0" w:after="200" w:line="160" w:lineRule="exact"/>\n'
        f'  </w:pPr>\n'
        f'</w:p>'
    )
    tbl_elm.addnext(p_dem_duoi)

def lay_text(el):
    texts = el.xpath('.//w:t')
    return "".join([t.text for t in texts if t.text]).strip()

def lay_style(el):
    val = el.xpath('.//w:pPr/w:pStyle/@w:val')
    return val[0] if val else ""

def la_tieu_de_cau(el):
    """Xác định paragraph có phải là điểm bắt đầu câu hỏi mới"""
    style = lay_style(el).lower()
    txt = lay_text(el)
    
    if any(k in style for k in ["cau", "question", "heading", "bai"]):
        return True
    if len(el.xpath('.//w:numPr')) > 0:
        return True
    if re.match(r"^(Câu\s+\d+|Bài\s+\d+)", txt, re.IGNORECASE):
        return True
    return False

def phan_loai_cau(ds_elements):
    """
    Phân loại chính xác dạng câu hỏi:
    1. Trắc nghiệm 4 đáp án: A., B., C., D. -> 2 dòng
    2. Đúng / Sai: a), b), c), d) -> 6 dòng
    3. Trả lời ngắn / Tự luận: còn lại -> 8 dòng
    """
    toan_bo_text = " ".join([lay_text(e) for e in ds_elements])

    # 1. Kiểm tra TRẮC NGHIỆM 4 ĐÁP ÁN (A., B., C., D. in hoa)
    co_A = bool(re.search(r"(?:^|\s|[^\w])A[\.\)]", toan_bo_text))
    co_B = bool(re.search(r"(?:^|\s|[^\w])B[\.\)]", toan_bo_text))
    co_C = bool(re.search(r"(?:^|\s|[^\w])C[\.\)]", toan_bo_text))
    co_D = bool(re.search(r"(?:^|\s|[^\w])D[\.\)]", toan_bo_text))
    dem_tn = sum([co_A, co_B, co_C, co_D])
    if dem_tn >= 3:
        return SO_DONG_TRAC_NGHIEM, "Trắc nghiệm"

    # 2. Kiểm tra ĐÚNG / SAI (a), b), c), d) chữ thường)
    co_a = bool(re.search(r"(?:^|\s|[^\w])a\)", toan_bo_text))
    co_b = bool(re.search(r"(?:^|\s|[^\w])b\)", toan_bo_text))
    co_c = bool(re.search(r"(?:^|\s|[^\w])c\)", toan_bo_text))
    co_d = bool(re.search(r"(?:^|\s|[^\w])d\)", toan_bo_text))
    dem_ds = sum([co_a, co_b, co_c, co_d])
    if dem_ds >= 3:
        return SO_DONG_DUNG_SAI, "Đúng / Sai"

    # 3. Còn lại là TRẢ LỜI NGẮN / TỰ LUẬN
    return SO_DONG_TRA_LOI_NGAN, "Trả lời ngắn"

def xu_ly_file_docx(input_path, output_path):
    print(f"-> Đang xử lý: {os.path.basename(input_path)}...")
    doc = Document(input_path)
    body = doc._body._body

    elements = [e for e in body if e.tag.endswith(('p', 'tbl'))]

    nhom_cac_cau = []
    cau_hien_tai = []

    for el in elements:
        txt = lay_text(el)
        if txt.lower().startswith("phần") or txt.lower().startswith("tổng hợp"):
            if cau_hien_tai:
                nhom_cac_cau.append(cau_hien_tai)
                cau_hien_tai = []
            continue

        if la_tieu_de_cau(el):
            if cau_hien_tai:
                nhom_cac_cau.append(cau_hien_tai)
            cau_hien_tai = [el]
        else:
            if cau_hien_tai:
                cau_hien_tai.append(el)

    if cau_hien_tai:
        nhom_cac_cau.append(cau_hien_tai)

    tong_so_cau = len(nhom_cac_cau)
    dem_tn, dem_ds, dem_tln = 0, 0, 0

    for cau in reversed(nhom_cac_cau):
        so_dong, loai = phan_loai_cau(cau)
        if loai == "Trắc nghiệm":
            dem_tn += 1
        elif loai == "Đúng / Sai":
            dem_ds += 1
        else:
            dem_tln += 1

        el_cuoi_cung = cau[-1]
        chen_dong_ke(el_cuoi_cung, so_dong)

    doc.save(output_path)
    print(f"   => Xong! Tổng cộng {tong_so_cau} câu:")
    print(f"      - Trắc nghiệm: {dem_tn} câu (mỗi câu 2 dòng)")
    print(f"      - Đúng / Sai:  {dem_ds} câu (mỗi câu 6 dòng)")
    print(f"      - Trả lời ngắn: {dem_tln} câu (mỗi câu 8 dòng)")

def main():
    thu_muc = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(thu_muc, "input")
    output_dir = os.path.join(thu_muc, "output")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    files = [f for f in os.listdir(input_dir) if f.endswith(".docx") and not f.startswith("~$")]
    if not files:
        print("Chưa có file .docx nào trong thư mục 'input'!")
        return
        
    for f in files:
        in_p = os.path.join(input_dir, f)
        out_p = os.path.join(output_dir, f.replace(".docx", "_co_dong_ke.docx"))
        xu_ly_file_docx(in_p, out_p)
        
    print("\nHoàn tất 100%!")

if __name__ == "__main__":
    main()