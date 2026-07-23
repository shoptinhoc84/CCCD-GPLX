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

st.title("🖨️ Photo Ghép Giấy Tờ Tự Động (Kéo thả 1 Lần)")
st.write(
    "Tải lên **cùng lúc 4 ảnh** (2 mặt CCCD + 2 mặt GPLX). Hệ thống sẽ **tự nhận diện từng mặt**, cắt viền và xếp lên trang A4."
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


def classify_card(pil_img):
    """Phân loại ảnh thuộc loại nào: 'cccd_front', 'cccd_back', 'gplx_front', 'gplx_back'."""
    img_np = np.array(pil_img)
    h, w, _ = img_np.shape

    # Chuyển sang HSV để phân tích màu sắc
    hsv = cv2.cvtColor(img_np, cv2.COLOR_RGB2HSV)

    # 1. Phát hiện GPLX dựa trên sắc màu vàng nhạt/kem đặc trưng
    # Màu vàng trong HSV có dải H từ 15 đến 35
    yellow_mask = cv2.inRange(hsv, (15, 30, 150), (35, 255, 255))
    yellow_ratio = np.sum(yellow_mask > 0) / (h * w)

    is_gplx = yellow_ratio > 0.25

    # 2. Phân biệt mặt trước / mặt sau
    # Mặt sau CCCD chứa chip điện tử hình chữ nhật màu vàng hoặc dấu vân tay tối màu
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    # Mặt trước thường có ảnh chân dung bên trái (vùng tối/độ lệch chuẩn cao)
    left_side = gray[:, : int(w * 0.4)]
    right_side = gray[:, int(w * 0.6) :]

    # Kiểm tra vết chip vàng mặt sau CCCD (xét nửa bên trái)
    chip_mask = cv2.inRange(
        hsv[: int(h * 0.8), : int(w * 0.5)], (15, 100, 100), (30, 255, 255)
    )
    has_chip = np.sum(chip_mask > 0) > (h * w * 0.005)

    if is_gplx:
        # Mặt trước GPLX thường chứa ảnh chân dung màu sắc phức tạp hơn mặt sau
        if np.std(left_side) > np.std(right_side) + 5:
            return "gplx_front"
        else:
            return "gplx_back"
    else:
        if has_chip:
            return "cccd_back"
        elif np.std(left_side) > np.std(right_side) + 3:
            return "cccd_front"
        else:
            return "cccd_back"


def add_card_row_to_doc(doc, pil_front, pil_back, is_last=False):
    """Thêm 1 hàng 2 mặt nằm ngang vào file Word (Không chữ)."""
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

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = docx.oxml.OxmlElement("w:tcBorders")
            for border_name in ["top", "left", "bottom", "right"]:
                border = docx.oxml.OxmlElement(f"w:{border_name}")
                border.set(docx.oxml.ns.qn("w:val"), "none")
                tcBorders.append(border)
            tcPr.append(tcBorders)

    if not is_last:
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(20)
        p_space.paragraph_format.space_after = Pt(0)


def create_docx_combined(cards_data):
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


# Giao diện cho phép chọn 1 lúc nhiều ảnh
uploaded_files = st.file_uploader(
    "📁 Chọn hoặc Kéo thả 4 ảnh cùng lúc (2 mặt CCCD + 2 mặt GPLX):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 4:
        st.warning(
            f"⚠️ Bạn đã tải lên {len(uploaded_files)}/4 ảnh. Vui lòng chọn đủ 4 ảnh để ghép."
        )
    else:
        with st.spinner("Đang tự động nhận diện từng loại giấy tờ..."):
            classified = {}

            # Đọc & cắt viền từng ảnh trước
            cropped_images = []
            for f in uploaded_files[:4]:
                raw_img = resize_image_if_large(
                    Image.open(f).convert("RGB")
                )
                cropped_np = crop_card(np.array(raw_img))
                cropped_pil = Image.fromarray(cropped_np)

                # Phân loại
                card_type = classify_card(cropped_pil)
                classified[card_type] = cropped_pil

            # Kiểm tra xem có đủ 4 loại mặt không
            required_keys = [
                "cccd_front",
                "cccd_back",
                "gplx_front",
                "gplx_back",
            ]
            missing_keys = [k for k in required_keys if k not in classified]

            # Nếu nhận diện tự động bị nhầm do ảnh chụp quá mờ, tự động gán theo thứ tự
            if missing_keys:
                sorted_files = [
                    Image.fromarray(crop_card(np.array(resize_image_if_large(Image.open(f).convert("RGB")))))
                    for f in uploaded_files[:4]
                ]
                classified = {
                    "cccd_front": sorted_files[0],
                    "cccd_back": sorted_files[1],
                    "gplx_front": sorted_files[2],
                    "gplx_back": sorted_files[3],
                }

        st.markdown("---")
        st.subheader("📸 Kết quả tự động phân loại & cắt viền")
        r1_c1, r1_c2, r2_c1, r2_c2 = st.columns(4)
        r1_c1.image(
            classified["cccd_front"],
            caption="CCCD Mặt trước",
            use_container_width=True,
        )
        r1_c2.image(
            classified["cccd_back"],
            caption="CCCD Mặt sau",
            use_container_width=True,
        )
        r2_c1.image(
            classified["gplx_front"],
            caption="GPLX Mặt trước",
            use_container_width=True,
        )
        r2_c2.image(
            classified["gplx_back"],
            caption="GPLX Mặt sau",
            use_container_width=True,
        )

        # Ghép File Word (.docx)
        cards_data = [
            {
                "front": classified["cccd_front"],
                "back": classified["cccd_back"],
            },
            {
                "front": classified["gplx_front"],
                "back": classified["gplx_back"],
            },
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
        pil_cccd_f, pil_cccd_b = (
            classified["cccd_front"],
            classified["cccd_back"],
        )
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
        pil_gplx_f, pil_gplx_b = (
            classified["gplx_front"],
            classified["gplx_back"],
        )
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
        st.subheader("📄 Xem trước Trang in A4")
        st.image(
            canvas, width=450, caption="Trang A4 ghép hoàn chỉnh sẵn sàng để in"
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
