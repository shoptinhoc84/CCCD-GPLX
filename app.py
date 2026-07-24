import io
import cv2
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="Công cụ Photo / Ghép giấy tờ in A4",
    layout="centered",
    page_icon="🖨️",
)

st.title("🖨️ Ghép giấy tờ A4 (Ngang - 2 bộ/trang)")
st.write(
    "Hỗ trợ tự động xếp **Mặt trước - Mặt sau nằm ngang**, tối ưu **2 bộ giấy tờ trên 1 trang A4**."
)


def crop_card(image_np):
    """Cắt thẻ chuẩn xác dựa trên nhận diện viền chữ nhật."""
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
        if area > (w_img * h_img * 0.12):
            x, y, w, h = cv2.boundingRect(cnt)
            aspect_ratio = float(w) / h

            if 1.2 <= aspect_ratio <= 1.9 and (w < w_img * 0.98):
                if area > max_area:
                    max_area = area
                    best_rect = (x, y, w, h)

    if best_rect:
        x, y, w, h = best_rect
        return image_np[
            max(0, y - 3) : min(h_img, y + h + 6),
            max(0, x - 3) : min(w_img, x + w + 6),
        ]

    return image_np


def add_card_row_to_word(doc, pil_front, pil_back):
    """Thêm 1 bộ (Mặt trước + Mặt sau nằm ngang) vào file Word dưới dạng Bảng 1 dòng 2 cột."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Chuyển ảnh sang Byte Stream
    buf_front, buf_back = io.BytesIO(), io.BytesIO()
    pil_front.save(buf_front, format="PNG")
    pil_back.save(buf_back, format="PNG")
    buf_front.seek(0)
    buf_back.seek(0)

    # Ô 1: Mặt trước
    cell_f = table.cell(0, 0)
    p_f = cell_f.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f.add_run().add_picture(buf_front, width=Inches(3.3))

    # Ô 2: Mặt sau
    cell_b = table.cell(0, 1)
    p_b = cell_b.paragraphs[0]
    p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b.add_run().add_picture(buf_back, width=Inches(3.3))

    # Khoảng cách giữa các bộ
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(24)


def create_multi_docx(card_pairs):
    """Tạo file Word (.docx) chứa 2 bộ nằm trên 1 trang A4."""
    doc = docx.Document()

    # Cấu hình lề A4
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for i, (pil_front, pil_back) in enumerate(card_pairs):
        # Sau mỗi 2 bộ thì ngắt sang trang A4 mới
        if i > 0 and i % 2 == 0:
            doc.add_page_break()

        add_card_row_to_word(doc, pil_front, pil_back)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def create_a4_canvas_horizontal(card_pairs_chunk):
    """Tạo file ảnh PNG khổ A4 chứa tối đa 2 bộ nằm ngang."""
    a4_w, a4_h = 1240, 1754  # Khổ A4 chuẩn pixel
    canvas = Image.new("RGB", (a4_w, a4_h), "white")

    card_w = 540  # Chiều rộng 1 mặt thẻ trên canvas
    spacing_x = 40  # Khoảng cách giữa mặt trước và mặt sau
    start_x = (a4_w - (card_w * 2 + spacing_x)) // 2

    y_positions = [180, 920]  # Tọa độ Y cho Bộ 1 và Bộ 2

    for idx, (pil_front, pil_back) in enumerate(card_pairs_chunk):
        if idx >= 2:
            break

        y_pos = y_positions[idx]

        # Resize mặt trước
        ratio_f = card_w / pil_front.width
        f_resized = pil_front.resize((card_w, int(pil_front.height * ratio_f)))

        # Resize mặt sau
        ratio_b = card_w / pil_back.width
        b_resized = pil_back.resize((card_w, int(pil_back.height * ratio_b)))

        # Dán vào Canvas A4
        canvas.paste(f_resized, (start_x, y_pos))
        canvas.paste(b_resized, (start_x + card_w + spacing_x, y_pos))

    return canvas


# Giao diện ứng dụng
doc_type = st.radio(
    "📋 Chọn loại giấy tờ cần xử lý:",
    ["Căn cước công dân (CCCD)", "Giấy phép lái xe (GPLX)"],
    horizontal=True,
)

uploaded_files = st.file_uploader(
    "📸 Chọn TẤT CẢ ảnh giấy tờ (Hệ thống ghép từng cặp 1-2, 3-4,...):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    num_files = len(uploaded_files)

    if num_files < 2:
        st.warning("⚠️ Vui lòng chọn ít nhất 2 ảnh (mặt trước và mặt sau).")
    else:
        card_pairs = []
        a4_canvases = []

        with st.spinner("Đang tự động lọc viền & xếp 2 bộ / trang A4..."):
            for i in range(0, num_files - (num_files % 2), 2):
                img_front_raw = Image.open(uploaded_files[i]).convert("RGB")
                img_back_raw = Image.open(uploaded_files[i + 1]).convert("RGB")

                cropped_front = crop_card(np.array(img_front_raw))
                cropped_back = crop_card(np.array(img_back_raw))

                card_pairs.append(
                    (Image.fromarray(cropped_front), Image.fromarray(cropped_back))
                )

            # Ghép ảnh A4 (Mỗi trang chứa tối đa 2 bộ)
            for i in range(0, len(card_pairs), 2):
                chunk = card_pairs[i : i + 2]
                canvas = create_a4_canvas_horizontal(chunk)
                a4_canvases.append(canvas)

        st.success(
            f"✅ Đã xử lý thành công **{len(card_pairs)}** bộ {doc_type.split()[0]}!"
        )

        st.markdown("---")
        st.subheader("📄 Xem trước mẫu A4 (Ngang 2 bộ)")

        cols = st.columns(min(len(a4_canvases), 3))
        for idx, canvas in enumerate(a4_canvases):
            cols[idx % 3].image(
                canvas, caption=f"Trang A4 thứ {idx + 1}", use_container_width=True
            )

        # Xuất file Word
        docx_io = create_multi_docx(card_pairs)

        st.markdown("---")
        st.subheader("📥 Tải về tập tin")

        btn_col1, btn_col2 = st.columns(2)

        with btn_col1:
            st.download_button(
                label=f"📝 Tải WORD ({len(card_pairs)} bộ - Nằm ngang) (.docx)",
                data=docx_io,
                file_name=f"{doc_type.split()[0]}_A4_Ngang.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with btn_col2:
            img_io = io.BytesIO()
            a4_canvases[0].save(img_io, format="PNG")
            img_io.seek(0)

            st.download_button(
                label="🖼️ Tải ÁNH Trang 1 (.png)",
                data=img_io,
                file_name=f"{doc_type.split()[0]}_Trang1.png",
                mime="image/png",
                use_container_width=True,
            )
