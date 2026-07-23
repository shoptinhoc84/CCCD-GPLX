import io
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="Công cụ Photo / Ghép giấy tờ in A4",
    layout="centered",
    page_icon="🖨️",
)

st.title("🖨️ Photo Ghép CCCD + GPLX lên 1 Trang A4")
st.write(
    "Hỗ trợ tự động cắt bỏ phần thừa và xếp **CCCD (hàng trên) & GPLX (hàng dưới)** nằm ngang trên cùng 1 trang A4 để in/photocopy."
)


def crop_card(image_np):
    """Cắt thẻ chuẩn xác dựa trên nhận diện viền chữ nhật và tỷ lệ khung hình chuẩn."""
    h_img, w_img, _ = image_np.shape

    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # Threshold tự động (Otsu)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    contours, _ = cv2.findContours(
        thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE
    )

    best_rect = None
    max_area = 0

    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area > (w_img * h_img * 0.10):
            x, y, w, h = cv2.boundingRect(cnt)
            aspect_ratio = float(w) / h

            if 1.2 <= aspect_ratio <= 1.9 and (w < w_img * 0.98):
                if area > max_area:
                    max_area = area
                    best_rect = (x, y, w, h)

    if best_rect:
        x, y, w, h = best_rect
        x = max(0, x - 3)
        y = max(0, y - 3)
        w = min(w_img - x, w + 6)
        h = min(h_img - y, h + 6)
        return image_np[y : y + h, x : x + w]

    # Dự phòng bằng Canny Edge Detection
    edges = cv2.Canny(blur, 30, 150)
    contours, _ = cv2.findContours(
        edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if contours:
        valid_cnts = [
            c
            for c in contours
            if cv2.contourArea(c) < (w_img * h_img * 0.95)
            and cv2.contourArea(c) > (w_img * h_img * 0.08)
        ]
        if valid_cnts:
            c = max(valid_cnts, key=cv2.contourArea)
            x, y, w, h = cv2.boundingRect(c)
            return image_np[y : y + h, x : x + w]

    return image_np


def add_card_row_to_doc(doc, pil_front, pil_back, label=""):
    """Thêm 1 hàng gồm 2 mặt nằm ngang vào file Word."""
    if label:
        p_title = doc.add_paragraph()
        run_title = p_title.add_run(label)
        run_title.bold = True
        run_title.font.size = Pt(12)
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(4)

    buf_front = io.BytesIO()
    pil_front.save(buf_front, format="PNG")
    buf_front.seek(0)

    buf_back = io.BytesIO()
    pil_back.save(buf_back, format="PNG")
    buf_back.seek(0)

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)

    p_left = cell_left.paragraphs[0]
    p_left.add_run().add_picture(buf_front, width=Inches(3.37))

    p_right = cell_right.paragraphs[0]
    p_right.add_run().add_picture(buf_back, width=Inches(3.37))

    # Bỏ đường viền của bảng
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = docx.oxml.OxmlElement("w:tcBorders")
            for border_name in ["top", "left", "bottom", "right"]:
                border = docx.oxml.OxmlElement(f"w:{border_name}")
                border.set(docx.oxml.ns.qn("w:val"), "none")
                tcBorders.append(border)
            tcPr.append(tcBorders)


def create_docx_combined(cards_data):
    """Tạo file Word (.docx) chứa cả CCCD và GPLX."""
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for item in cards_data:
        add_card_row_to_doc(
            doc, item["front"], item["back"], label=item["title"]
        )

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# Giao diện tải ảnh
st.subheader("1. Ảnh Căn cước công dân (CCCD)")
col1, col2 = st.columns(2)
with col1:
    cccd_front_file = st.file_uploader(
        "Mặt trước CCCD", type=["jpg", "jpeg", "png"], key="cccd_front"
    )
with col2:
    cccd_back_file = st.file_uploader(
        "Mặt sau CCCD", type=["jpg", "jpeg", "png"], key="cccd_back"
    )

st.markdown("---")
st.subheader("2. Ảnh Giấy phép lái xe (GPLX)")
col3, col4 = st.columns(2)
with col3:
    gplx_front_file = st.file_uploader(
        "Mặt trước GPLX", type=["jpg", "jpeg", "png"], key="gplx_front"
    )
with col4:
    gplx_back_file = st.file_uploader(
        "Mặt sau GPLX", type=["jpg", "jpeg", "png"], key="gplx_back"
    )

if cccd_front_file and cccd_back_file and gplx_front_file and gplx_back_file:
    with st.spinner("Đang tự động tách nền & xếp bố cục..."):
        # Process CCCD
        pil_cccd_f = Image.fromarray(
            crop_card(
                np.array(Image.open(cccd_front_file).convert("RGB"))
            )
        )
        pil_cccd_b = Image.fromarray(
            crop_card(np.array(Image.open(cccd_back_file).convert("RGB")))
        )

        # Process GPLX
        pil_gplx_f = Image.fromarray(
            crop_card(
                np.array(Image.open(gplx_front_file).convert("RGB"))
            )
        )
        pil_gplx_b = Image.fromarray(
            crop_card(np.array(Image.open(gplx_back_file).convert("RGB")))
        )

    st.markdown("---")
    st.subheader("📸 Kết quả cắt từng mặt")
    r1_c1, r1_c2, r2_c1, r2_c2 = st.columns(4)
    r1_c1.image(pil_cccd_f, caption="CCCD Trước", use_container_width=True)
    r1_c2.image(pil_cccd_b, caption="CCCD Sau", use_container_width=True)
    r2_c1.image(pil_gplx_f, caption="GPLX Trước", use_container_width=True)
    r2_c2.image(pil_gplx_b, caption="GPLX Sau", use_container_width=True)

    # Ghép Ảnh PNG A4
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "white")

    target_w = 510
    gap_x = 50
    x_front = (a4_w - (target_w * 2 + gap_x)) // 2
    x_back = x_front + target_w + gap_x

    # Ghép CCCD (Hàng 1)
    r_cccd_f = target_w / pil_cccd_f.width
    cccd_f_res = pil_cccd_f.resize(
        (target_w, int(pil_cccd_f.height * r_cccd_f))
    )
    r_cccd_b = target_w / pil_cccd_b.width
    cccd_b_res = pil_cccd_b.resize(
        (target_w, int(pil_cccd_b.height * r_cccd_b))
    )

    y_cccd = 150
    canvas.paste(cccd_f_res, (x_front, y_cccd))
    canvas.paste(cccd_b_res, (x_back, y_cccd))

    # Ghép GPLX (Hàng 2)
    r_gplx_f = target_w / pil_gplx_f.width
    gplx_f_res = pil_gplx_f.resize(
        (target_w, int(pil_gplx_f.height * r_gplx_f))
    )
    r_gplx_b = target_w / pil_gplx_b.width
    gplx_b_res = pil_gplx_b.resize(
        (target_w, int(pil_gplx_b.height * r_gplx_b))
    )

    y_gplx = y_cccd + cccd_f_res.height + 120
    canvas.paste(gplx_f_res, (x_front, y_gplx))
    canvas.paste(gplx_b_res, (x_back, y_gplx))

    img_io = io.BytesIO()
    canvas.save(img_io, format="PNG")
    img_io.seek(0)

    # Ghép File Word (.docx)
    cards_data = [
        {"title": "1. CĂN CƯỚC CÔNG DÂN", "front": pil_cccd_f, "back": pil_cccd_b},
        {"title": "2. GIẤY PHÉP LÁI XE", "front": pil_gplx_f, "back": pil_gplx_b},
    ]
    docx_io = create_docx_combined(cards_data)

    st.markdown("---")
    st.subheader("📄 Xem trước Trang in A4 Hoàn chỉnh")
    st.image(
        canvas, width=450, caption="Trang A4 chứa cả CCCD và GPLX sẵn sàng để in"
    )

    btn1, btn2 = st.columns(2)
    with btn1:
        st.download_button(
            label="📝 Tải file WORD (.docx)",
            data=docx_io,
            file_name="CCCD_va_GPLX_in_A4.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with btn2:
        st.download_button(
            label="🖼️ Tải file ÁNH (.png)",
            data=img_io,
            file_name="CCCD_va_GPLX_in_A4.png",
            mime="image/png",
            use_container_width=True,
        )
