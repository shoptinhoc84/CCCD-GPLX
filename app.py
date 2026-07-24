import io
import cv2
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="Công cụ Photo / Ghép giấy tờ hàng loạt in A4",
    layout="centered",
    page_icon="🖨️",
)

st.title("🖨️ Lọc viền & Ghép nhiều giấy tờ in A4")
st.write(
    "Tải lên **nhiều ảnh cùng lúc** (bao gồm mặt trước & mặt sau). "
    "Hệ thống sẽ tự động ghép từng cặp (1-2, 3-4,...) thành các trang A4 hoàn chỉnh."
)


def crop_card(image_np):
    """Cắt thẻ chuẩn xác dựa trên nhận diện viền chữ nhật và tỷ lệ khung hình chuẩn."""
    h_img, w_img, _ = image_np.shape

    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # Threshold tự động (Otsu) để tách biệt thẻ và nền
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

            # Tỷ lệ thẻ chuẩn (CCCD / GPLX PET) dao động từ 1.2 - 1.9
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
            and cv2.contourArea(c) > (w_img * h_img * 0.1)
        ]
        if valid_cnts:
            c = max(valid_cnts, key=cv2.contourArea)
            x, y, w, h = cv2.boundingRect(c)
            return image_np[y : y + h, x : x + w]

    return image_np


def create_multi_docx(card_pairs):
    """Tạo file Word (.docx) chứa nhiều bộ giấy tờ, mỗi bộ nằm trên 1 trang A4 riêng."""
    doc = docx.Document()

    # Cấu hình lề trang A4
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    for i, (pil_front, pil_back) in enumerate(card_pairs):
        if i > 0:
            doc.add_page_break()  # Sang trang mới cho bộ tiếp theo

        # Chuyển PIL Image sang Byte Stream
        buf_front = io.BytesIO()
        pil_front.save(buf_front, format="PNG")
        buf_front.seek(0)

        buf_back = io.BytesIO()
        pil_back.save(buf_back, format="PNG")
        buf_back.seek(0)

        # Mặt trước
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run()
        run1.add_picture(buf_front, width=Inches(3.37))

        # Khoảng cách giữa 2 mặt
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(18)

        # Mặt sau
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run()
        run2.add_picture(buf_back, width=Inches(3.37))

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def create_a4_canvas(pil_front, pil_back):
    """Tạo ảnh A4 ghép mặt trước và mặt sau."""
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "white")
    target_w = 800

    ratio_f = target_w / pil_front.width
    pil_front_resized = pil_front.resize(
        (target_w, int(pil_front.height * ratio_f))
    )

    ratio_b = target_w / pil_back.width
    pil_back_resized = pil_back.resize(
        (target_w, int(pil_back.height * ratio_b))
    )

    x_pos = (a4_w - target_w) // 2
    canvas.paste(pil_front_resized, (x_pos, 180))
    canvas.paste(
        pil_back_resized, (x_pos, 180 + pil_front_resized.height + 80)
    )

    return canvas


# Giao diện ứng dụng
doc_type = st.radio(
    "📋 Chọn loại giấy tờ cần xử lý:",
    ["Căn cước công dân (CCCD)", "Giấy phép lái xe (GPLX)"],
    horizontal=True,
)

uploaded_files = st.file_uploader(
    "📸 Chọn TẤT CẢ ảnh giấy tờ (Cho phép chọn nhiều ảnh cùng lúc):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    num_files = len(uploaded_files)

    if num_files < 2:
        st.warning("⚠️ Vui lòng chọn ít nhất 2 ảnh (mặt trước và mặt sau).")
    else:
        if num_files % 2 != 0:
            st.info(
                f"ℹ️ Bạn đã tải lên **{num_files}** ảnh (số lẻ). Hệ thống sẽ xử lý **{num_files - 1}** ảnh đủ cặp đầu tiên."
            )

        card_pairs = []
        a4_canvases = []

        with st.spinner("Đang tự động lọc viền & ghép các bộ giấy tờ..."):
            # Lặp qua từng cặp ảnh (2 ảnh / bộ)
            for i in range(0, num_files - (num_files % 2), 2):
                file_front = uploaded_files[i]
                file_back = uploaded_files[i + 1]

                img_front_raw = Image.open(file_front).convert("RGB")
                img_back_raw = Image.open(file_back).convert("RGB")

                cropped_front = crop_card(np.array(img_front_raw))
                cropped_back = crop_card(np.array(img_back_raw))

                pil_front = Image.fromarray(cropped_front)
                pil_back = Image.fromarray(cropped_back)

                card_pairs.append((pil_front, pil_back))

                # Tạo trang A4 tương ứng
                canvas = create_a4_canvas(pil_front, pil_back)
                a4_canvases.append(canvas)

        st.success(
            f"✅ Đã xử lý thành công **{len(card_pairs)}** bộ {doc_type.split()[0]}!"
        )

        st.markdown("---")
        st.subheader("📄 Xem trước kết quả")

        # Hiển thị xem trước danh sách các trang A4
        cols = st.columns(min(len(a4_canvases), 3))
        for idx, canvas in enumerate(a4_canvases):
            col_idx = idx % 3
            cols[col_idx].image(
                canvas,
                caption=f"Bộ {idx + 1} (A4)",
                use_container_width=True,
            )

        # Xuất file Word chung cho tất cả các trang
        docx_io = create_multi_docx(card_pairs)

        st.markdown("---")
        st.subheader("📥 Tải về tập tin đã ghép")

        btn_col1, btn_col2 = st.columns(2)

        with btn_col1:
            st.download_button(
                label=f"📝 Tải WORD chứa {len(card_pairs)} bộ (.docx)",
                data=docx_io,
                file_name=f"{doc_type.split()[0]}_HangLoat.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with btn_col2:
            # Tải ảnh trang A4 đầu tiên (hoặc bộ đầu tiên)
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
