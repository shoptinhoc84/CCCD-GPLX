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

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .main-title { color: #0f172a; font-weight: 800; font-size: 2.2rem; margin-bottom: 0.2rem; }
    .sub-title { color: #475569; font-size: 1.05rem; margin-bottom: 2rem; }
    .status-badge {
        display: inline-block; padding: 4px 12px; border-radius: 20px;
        font-size: 0.85rem; font-weight: 600; margin-bottom: 10px;
    }
    .badge-cccd { background-color: #e0f2fe; color: #0369a1; }
    .badge-gplx { background-color: #fef3c7; color: #b45309; }
    
    div.stDownloadButton > button {
        width: 100%; border-radius: 8px; height: 48px; font-weight: 600; font-size: 1rem;
        transition: all 0.2s ease;
    }
    div.stDownloadButton > button:first-child { background-color: #2563eb; color: white; border: none; }
    div.stDownloadButton > button:first-child:hover { background-color: #1d4ed8; }
    </style>
""",
    unsafe_allow_html=True,
)


# 2. CÁC HÀM XỬ LÝ ẢNH & NHẬN DIỆN THÔNG MINH
def resize_image_if_large(pil_img, max_dim=1200):
    """Giảm kích thước ảnh đầu vào để xử lý siêu nhanh & tránh quá tải RAM."""
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        return pil_img.resize(
            (int(w * scale), int(h * scale)), Image.Resampling.LANCZOS
        )
    return pil_img


def crop_card(image_np):
    """Cắt viền tự động sử dụng thuật toán nhận diện hình chữ nhật chuẩn."""
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
    """Phân loại tự động các mặt giấy tờ (CCCD / GPLX)."""
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


def create_docx_dynamic(cards_rows, space_between=20):
    """Tạo file Word (.docx) chuẩn lề A4, không đường viền bảng."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for idx, item in enumerate(cards_rows):
        buf_f, buf_b = io.BytesIO(), io.BytesIO()
        item["front"].save(buf_f, format="JPEG", quality=92)
        if item["back"]:
            item["back"].save(buf_b, format="JPEG", quality=92)
            buf_b.seek(0)
        buf_f.seek(0)

        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        table.cell(0, 0).paragraphs[0].add_run().add_picture(
            buf_f, width=Inches(3.37)
        )
        if item["back"]:
            table.cell(0, 1).paragraphs[0].add_run().add_picture(
                buf_b, width=Inches(3.37)
            )

        # Xóa đường viền của bảng
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcBorders = docx.oxml.OxmlElement("w:tcBorders")
                for border_name in ["top", "left", "bottom", "right"]:
                    border = docx.oxml.OxmlElement(f"w:{border_name}")
                    border.set(docx.oxml.ns.qn("w:val"), "none")
                    tcBorders.append(border)
                tcPr.append(tcBorders)

        if idx < len(cards_rows) - 1:
            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(space_between)
            p_space.paragraph_format.space_after = Pt(0)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# 3. SIDEBAR CẤU HÌNH
with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/96/print.png", width=64
    )
    st.markdown("### ⚙️ Cấu hình Trang In")
    space_val = st.slider(
        "Khoảng cách giữa các hàng (px):", 60, 200, 120, 10
    )
    draw_border = st.checkbox("Thêm viền mỏng quanh thẻ khi in", value=False)
    st.markdown("---")
    st.info(
        "💡 **Hướng dẫn:** Kéo thả từ **1 đến 4 ảnh** bất kỳ (CCCD hoặc GPLX). Hệ thống sẽ tự cắt viền và xếp đều lên trang A4!"
    )

# 4. GIAO DIỆN CHÍNH (MAIN CONTENT)
st.markdown(
    '<div class="main-title">📄 Smart Photo Paper A4</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Tự động cắt viền & xếp trang in A4 linh hoạt cho CCCD và GPLX</div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "📥 Tải ảnh lên (Có thể chọn từ 1 đến 4 ảnh cùng lúc):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    with st.spinner("✨ Đang tự động tách nền và sắp xếp bố cục..."):
        processed_imgs = []

        # Tách nền & cắt viền từng ảnh
        for f in uploaded_files:
            raw_img = resize_image_if_large(Image.open(f).convert("RGB"))
            cropped_np = crop_card(np.array(raw_img))
            cropped_pil = Image.fromarray(cropped_np)

            if draw_border:
                img_border = Image.new(
                    "RGB",
                    (cropped_pil.width + 4, cropped_pil.height + 4),
                    "#cbd5e1",
                )
                img_border.paste(cropped_pil, (2, 2))
                cropped_pil = img_border

            ctype = classify_card(cropped_pil)
            processed_imgs.append({"type": ctype, "img": cropped_pil})

        # Phân nhóm CCCD và GPLX
        cccd_list = [
            x for x in processed_imgs if x["type"].startswith("cccd")
        ]
        gplx_list = [
            x for x in processed_imgs if x["type"].startswith("gplx")
        ]

        cards_rows = []

        # Nếu có CCCD
        if cccd_list:
            f_img = next(
                (x["img"] for x in cccd_list if x["type"] == "cccd_front"),
                cccd_list[0]["img"],
            )
            b_img = next(
                (x["img"] for x in cccd_list if x["type"] == "cccd_back"),
                cccd_list[1]["img"] if len(cccd_list) > 1 else None,
            )
            cards_rows.append({"label": "CCCD", "front": f_img, "back": b_img})

        # Nếu có GPLX
        if gplx_list:
            f_img = next(
                (x["img"] for x in gplx_list if x["type"] == "gplx_front"),
                gplx_list[0]["img"],
            )
            b_img = next(
                (x["img"] for x in gplx_list if x["type"] == "gplx_back"),
                gplx_list[1]["img"] if len(gplx_list) > 1 else None,
            )
            cards_rows.append({"label": "GPLX", "front": f_img, "back": b_img})

        # Dự phòng gán cặp ảnh nếu phân loại nhầm
        if not cards_rows:
            for i in range(0, len(processed_imgs), 2):
                f_img = processed_imgs[i]["img"]
                b_img = (
                    processed_imgs[i + 1]["img"]
                    if i + 1 < len(processed_imgs)
                    else None
                )
                cards_rows.append(
                    {"label": "Giấy tờ", "front": f_img, "back": b_img}
                )

    # Hiển thị các ảnh đã cắt
    st.markdown("### 📸 Kết quả tự động tách viền")
    cols = st.columns(len(processed_imgs))
    for idx, item in enumerate(processed_imgs):
        with cols[idx]:
            b_class = (
                "badge-cccd" if "cccd" in item["type"] else "badge-gplx"
            )
            st.markdown(
                f'<span class="status-badge {b_class}">{item["type"].upper()}</span>',
                unsafe_allow_html=True,
            )
            st.image(item["img"], use_container_width=True)

    # Ghép Canvas A4
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "#ffffff")
    target_w = 510
    gap_x = 50
    x_front = (a4_w - (target_w * 2 + gap_x)) // 2
    x_back = x_front + target_w + gap_x

    current_y = 160
    for row in cards_rows:
        # Mặt trước
        f_res = row["front"].resize(
            (
                target_w,
                int(row["front"].height * (target_w / row["front"].width)),
            )
        )
        canvas.paste(f_res, (x_front, current_y))

        # Mặt sau (nếu có)
        if row["back"]:
            b_res = row["back"].resize(
                (
                    target_w,
                    int(row["back"].height * (target_w / row["back"].width)),
                )
            )
            canvas.paste(b_res, (x_back, current_y))

        current_y += f_res.height + space_val

    # Xuất file Word & Ảnh
    docx_bytes = create_docx_dynamic(
        cards_rows, space_between=space_val
    ).getvalue()

    img_io = io.BytesIO()
    canvas.save(img_io, format="JPEG", quality=92)
    img_bytes = img_io.getvalue()

    st.markdown("---")
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
        st.write("File đã được xếp chuẩn lề in thực tế:")

        st.download_button(
            label="📝 Tải file WORD (.docx)",
            data=docx_bytes,
            file_name="Photo_GiayTo_A4.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.write("")
        st.download_button(
            label="🖼️ Tải file ÁNH (.jpg)",
            data=img_bytes,
            file_name="Photo_GiayTo_A4.jpg",
            mime="image/jpeg",
        )
