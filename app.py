import io
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# 1. CẤU HÌNH TRANG & CUSTOM CSS (DESIGN SYSTEM)
st.set_page_config(
    page_title="Photo Doc Pro - Ghép Giấy Tờ Tự Động",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Nhúng CSS tùy chỉnh giao diện
st.markdown(
    """
    <style>
    /* Tổng thể nền & font */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Header Styling */
    .main-title {
        color: #0f172a;
        font-weight: 800;
        font-size: 2.2rem;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #475569;
        font-size: 1.05rem;
        margin-bottom: 2rem;
    }
    
    /* Card Container */
    .card-box {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    
    /* Badge trạng thái */
    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 10px;
    }
    .badge-cccd { background-color: #e0f2fe; color: #0369a1; }
    .badge-gplx { background-color: #fef3c7; color: #b45309; }
    
    /* Nút tải về Custom */
    div.stDownloadButton > button {
        width: 100%;
        border-radius: 8px;
        height: 48px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.2s ease;
    }
    div.stDownloadButton > button:first-child {
        background-color: #2563eb;
        color: white;
        border: none;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #1d4ed8;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
    }
    
    /* Tối ưu Sidebar */
    .css-1d37w0e {
        background-color: #ffffff;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# 2. CÁC HÀM XỬ LÝ ẢNH & TỐI ƯU
def resize_image_if_large(pil_img, max_dim=1200):
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        return pil_img.resize(
            (int(w * scale), int(h * scale)), Image.Resampling.LANCZOS
        )
    return pil_img


def crop_card(image_np):
    """Cắt viền tự động dùng thuật toán nhận diện chữ nhật chuẩn."""
    h_img, w_img, _ = image_np.shape
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    contours, _ = cv2.findContours(
        thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE
    )
    best_rect, max_area = None, 0

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
        x, y = max(0, x - 3), max(0, y - 3)
        w, h = min(w_img - x, w + 6), min(h_img - y, h + 6)
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
    """Phân loại tự động 4 mặt giấy tờ."""
    img_np = np.array(pil_img)
    h, w, _ = img_np.shape
    hsv = cv2.cvtColor(img_np, cv2.COLOR_RGB2HSV)

    yellow_mask = cv2.inRange(hsv, (15, 30, 150), (35, 255, 255))
    is_gplx = (np.sum(yellow_mask > 0) / (h * w)) > 0.25

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    left_side = gray[:, : int(w * 0.4)]
    right_side = gray[:, int(w * 0.6) :]

    chip_mask = cv2.inRange(
        hsv[: int(h * 0.8), : int(w * 0.5)], (15, 100, 100), (30, 255, 255)
    )
    has_chip = np.sum(chip_mask > 0) > (h * w * 0.005)

    if is_gplx:
        return (
            "gplx_front"
            if np.std(left_side) > np.std(right_side) + 5
            else "gplx_back"
        )
    else:
        if has_chip:
            return "cccd_back"
        return (
            "cccd_front"
            if np.std(left_side) > np.std(right_side) + 3
            else "cccd_back"
        )


def create_docx_combined(cards_data, space_between=20):
    """Tạo file Word chuyên nghiệp, chuẩn lề in A4."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for idx, item in enumerate(cards_data):
        buf_f, buf_b = io.BytesIO(), io.BytesIO()
        item["front"].save(buf_f, format="JPEG", quality=92)
        item["back"].save(buf_b, format="JPEG", quality=92)
        buf_f.seek(0)
        buf_b.seek(0)

        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        table.cell(0, 0).paragraphs[0].add_run().add_picture(
            buf_f, width=Inches(3.37)
        )
        table.cell(0, 1).paragraphs[0].add_run().add_picture(
            buf_b, width=Inches(3.37)
        )

        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcBorders = docx.oxml.OxmlElement("w:tcBorders")
                for b_name in ["top", "left", "bottom", "right"]:
                    border = docx.oxml.OxmlElement(f"w:{b_name}")
                    border.set(docx.oxml.ns.qn("w:val"), "none")
                    tcBorders.append(border)
                tcPr.append(tcBorders)

        if idx < len(cards_data) - 1:
            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(space_between)
            p_space.paragraph_format.space_after = Pt(0)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# 3. SIDEBAR TÙY CHỈNH NÂNG CAO
with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/96/print.png", width=64
    )
    st.markdown("### ⚙️ Cấu hình Trang In")

    space_val = st.slider(
        "Khoảng cách giữa CCCD & GPLX (px):",
        min_value=60,
        max_value=200,
        value=120,
        step=10,
    )

    draw_border = st.checkbox("Thêm viền mỏng quanh thẻ khi in", value=False)

    st.markdown("---")
    st.markdown("### 💡 Hướng dẫn nhanh")
    st.info(
        "1. Kéo thả **cùng lúc 4 ảnh** (2 mặt CCCD + 2 mặt GPLX) vào ô bên phải.\n"
        "2. Hệ thống tự nhận diện và cắt bỏ phần thừa.\n"
        "3. Tải file **Word (.docx)** hoặc **Ảnh (.jpg)** về in ngay!"
    )

# 4. GIAO DIỆN CHÍNH (MAIN CONTENT)
st.markdown(
    '<div class="main-title">📄 Smart Photo Paper A4</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Công cụ tự động cắt viền & xếp bố cục in A4 cho CCCD và GPLX</div>',
    unsafe_allow_html=True,
)

# Zone Tải ảnh
uploaded_files = st.file_uploader(
    "📥 Tải lên 4 ảnh giấy tờ (Thả 4 file cùng lúc):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 4:
        st.warning(
            f"⚠️ Bạn mới tải lên **{len(uploaded_files)}/4** ảnh. Vui lòng tải đủ 4 mặt để tự động xếp A4."
        )
    else:
        with st.spinner("✨ Đang phân tích màu sắc, tách nền và xếp trang..."):
            classified = {}

            for f in uploaded_files[:4]:
                raw_img = resize_image_if_large(
                    Image.open(f).convert("RGB")
                )
                cropped_np = crop_card(np.array(raw_img))
                cropped_pil = Image.fromarray(cropped_np)

                # Vẽ viền nếu người dùng bật checkbox sidebar
                if draw_border:
                    img_border = Image.new(
                        "RGB",
                        (cropped_pil.width + 4, cropped_pil.height + 4),
                        "#cbd5e1",
                    )
                    img_border.paste(cropped_pil, (2, 2))
                    cropped_pil = img_border

                card_type = classify_card(cropped_pil)
                classified[card_type] = cropped_pil

            # Fallback nếu thiếu key
            required_keys = [
                "cccd_front",
                "cccd_back",
                "gplx_front",
                "gplx_back",
            ]
            if any(k not in classified for k in required_keys):
                sorted_files = [
                    Image.fromarray(
                        crop_card(
                            np.array(
                                resize_image_if_large(
                                    Image.open(f).convert("RGB")
                                )
                            )
                        )
                    )
                    for f in uploaded_files[:4]
                ]
                classified = {
                    "cccd_front": sorted_files[0],
                    "cccd_back": sorted_files[1],
                    "gplx_front": sorted_files[2],
                    "gplx_back": sorted_files[3],
                }

        # BẢNG HIỂN THỊ KẾT QUẢ CẮT
        st.markdown("### 📸 Kết quả tự động phân loại")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(
                '<span class="status-badge badge-cccd">CCCD - Trước</span>',
                unsafe_allow_html=True,
            )
            st.image(classified["cccd_front"], use_container_width=True)

        with col2:
            st.markdown(
                '<span class="status-badge badge-cccd">CCCD - Sau</span>',
                unsafe_allow_html=True,
            )
            st.image(classified["cccd_back"], use_container_width=True)

        with col3:
            st.markdown(
                '<span class="status-badge badge-gplx">GPLX - Trước</span>',
                unsafe_allow_html=True,
            )
            st.image(classified["gplx_front"], use_container_width=True)

        with col4:
            st.markdown(
                '<span class="status-badge badge-gplx">GPLX - Sau</span>',
                unsafe_allow_html=True,
            )
            st.image(classified["gplx_back"], use_container_width=True)

        # TẠO FILE WORD & KHUNG A4
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

        # Canvas A4
        a4_w, a4_h = 1240, 1754
        canvas = Image.new("RGB", (a4_w, a4_h), "#ffffff")

        target_w = 510
        gap_x = 50
        x_front = (a4_w - (target_w * 2 + gap_x)) // 2
        x_back = x_front + target_w + gap_x

        # Paste CCCD
        pil_cccd_f, pil_cccd_b = (
            classified["cccd_front"],
            classified["cccd_back"],
        )
        cccd_f_res = pil_cccd_f.resize(
            (target_w, int(pil_cccd_f.height * (target_w / pil_cccd_f.width)))
        )
        cccd_b_res = pil_cccd_b.resize(
            (target_w, int(pil_cccd_b.height * (target_w / pil_cccd_b.width)))
        )

        y_cccd = 160
        canvas.paste(cccd_f_res, (x_front, y_cccd))
        canvas.paste(cccd_b_res, (x_back, y_cccd))

        # Paste GPLX
        pil_gplx_f, pil_gplx_b = (
            classified["gplx_front"],
            classified["gplx_back"],
        )
        gplx_f_res = pil_gplx_f.resize(
            (target_w, int(pil_gplx_f.height * (target_w / pil_gplx_f.width)))
        )
        gplx_b_res = pil_gplx_b.resize(
            (target_w, int(pil_gplx_b.height * (target_w / pil_gplx_b.width)))
        )

        y_gplx = y_cccd + cccd_f_res.height + space_val
        canvas.paste(gplx_f_res, (x_front, y_gplx))
        canvas.paste(gplx_b_res, (x_back, y_gplx))

        img_io = io.BytesIO()
        canvas.save(img_io, format="JPEG", quality=92)
        img_bytes = img_io.getvalue()

        st.markdown("---")

        # XEM TRƯỚC VÀ NÚT TẢI
        p_col1, p_col2 = st.columns([1.2, 1])

        with p_col1:
            st.markdown("### 📄 Bản xem trước trang A4")
            st.image(
                canvas,
                use_container_width=True,
                caption="Trang A4 sẵn sàng để in",
            )

        with p_col2:
            st.markdown("### 📥 Tải file kết quả")
            st.write(
                "File xuất ra đã được căn chỉnh lề chuẩn A4 thực tế. Chọn định dạng bạn muốn:"
            )

            st.download_button(
                label="📝 Tải file WORD (.docx)",
                data=docx_bytes,
                file_name="Photo_CCCD_GPLX_A4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            st.write("")

            st.download_button(
                label="🖼️ Tải file ÁNH (.jpg)",
                data=img_bytes,
                file_name="Photo_CCCD_GPLX_A4.jpg",
                mime="image/jpeg",
            )
