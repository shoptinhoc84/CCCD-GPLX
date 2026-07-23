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
    "Tự động tách nền, cắt viền và ghép **CCCD (hàng trên) & GPLX (hàng dưới)** nằm ngang lên trang A4 để in/photocopy."
)


def resize_image_if_large(pil_img, max_dim=1200):
    """Giảm kích thước ảnh đầu vào để xử lý siêu nhanh & tránh quá tải RAM."""
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        new_w = int(w * scale)
        new_h = int(h * scale)
        return pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    return pil_img


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


def add_card_row_to_doc(doc, pil_front, pil_back, is_last=False):
    """Thêm 1 hàng gồm 2 mặt nằm ngang vào file Word (Không chữ)."""
    buf_front = io.BytesIO()
    pil_front.save(buf_front, format="JPEG", quality=90)
    buf_front.seek(0)

    buf_back = io.BytesIO()
    pil_back.save(buf_back, format="JPEG", quality=90)
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

    # Thêm khoảng trống giữa 2 hàng thẻ
    if not is_last:
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(20)
        p_space.paragraph_format.space_after = Pt(0)


def create_docx_combined(cards_data):
    """Tạo file Word (.docx) chứa thuần ảnh CCCD và GPLX (không tiêu đề)."""
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for idx, item in enumerate(cards_data):
        is_last = idx == len(cards_data) - 1
        add_card_row_to_doc(doc, item["front"], item["back"], is_last=is_last)

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
    with st.spinner("Đang tự động tách nền & tối ưu tốc độ..."):
        # Đọc & thu nhỏ kích thước nếu quá nặng
        raw_cccd_f = resize_image_if_large(
            Image.open(cccd_front_file).convert("RGB")
        )
        raw_cccd_b = resize_image_if_large(
            Image.open(cccd_back_file).convert("RGB")
        )
        raw_gplx_f = resize_image_if_large(
            Image.open(gplx_front_file).convert("RGB")
        )
        raw_gplx_b = resize_image_if_large(
            Image.open(gplx_back_file).convert("RGB")
        )

        # Tách nền cắt viền
        pil_cccd_f = Image.fromarray(crop_card(np.array(raw_cccd_f)))
        pil_cccd_b = Image.fromarray(crop_card(np.array(raw_cccd_b)))
        pil_gplx_f = Image.fromarray(crop_card(np.array(raw_gplx_f)))
        pil_gplx_b = Image.fromarray(crop_card(np.array(raw_gplx_b)))

        # Ghép File Word (.docx)
        cards_data = [
            {"front": pil_cccd_f, "back": pil_cccd_b},
            {"front": pil_gplx_f, "back": pil_gplx_b},
        ]
        docx_bytes = create_docx_combined(cards_data).getvalue()

        # Ghép Trang Ảnh A4
        a4_w, a4_h = 1240, 1754
        canvas = Image.new("RGB", (a4_w, a4_h), "white")

        target_w = 510
        gap_x = 50
        x_front = (a4_w - (target_w * 2 + gap_x)) // 2
        x_back = x_front + target_w + gap_x

        # CCCD
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

        # GPLX
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
        canvas.save(img_io, format="JPEG", quality=90)
        img_bytes = img_io.getvalue()

    st.markdown("---")
    st.subheader("📸 Kết quả cắt từng mặt")
    r1_c1, r1_c2, r2_c1, r2_c2 = st.columns(4)
    r1_c1.image(pil_cccd_f, caption="CCCD Trước", use_container_width=True)
    r1_c2.image(pil_cccd_b, caption="CCCD Sau", use_container_width=True)
    r2_c1.image(pil_gplx_f, caption="GPLX Trước", use_container_width=True)
    r2_c2.image(pil_gplx_b, caption="GPLX Sau", use_container_width=True)

    st.markdown("---")
    st.subheader("📄 Xem trước Trang in A4")
    st.image(
        canvas, width=450, caption="Trang A4 chứa cả CCCD và GPLX sẵn sàng để in"
    )

    btn1, btn2 = st.columns(2)
    with btn1:
        st.download_button(
            label="📝 Tải file WORD (.docx)",
            data=docx_bytes,
            file_name="CCCD_va_GPLX_in_A4.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with btn2:
        st.download_button(
            label="🖼️ Tải file ÁNH (.jpg)",
            data=img_bytes,
            file_name="CCCD_va_GPLX_in_A4.jpg",
            mime="image/jpeg",
            use_container_width=True,
        )
