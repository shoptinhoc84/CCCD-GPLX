import io
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image, ImageOps
import streamlit as st

# 1. CẤU HÌNH TRANG
st.set_page_config(
    page_title="Photo Doc Pro - Ghép Giấy Tờ A4",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .main-title { color: #0f172a; font-weight: 800; font-size: 1.8rem; margin-bottom: 0.2rem; }
    .sub-title { color: #475569; font-size: 0.95rem; margin-bottom: 1.5rem; }
    
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


# 2. XỬ LÝ ẢNH & TÁCH THẺ BẢO TOÀN NỘI DUNG
def load_and_fix_orientation(file):
    """Sửa lỗi xoay góc EXIF từ camera điện thoại & nén dung lượng chống tràn RAM."""
    img = Image.open(file)
    img = ImageOps.exif_transpose(img).convert("RGB")
    w, h = img.size
    max_dim = 1200
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        img = img.resize(
            (int(w * scale), int(h * scale)), Image.Resampling.LANCZOS
        )
    return img


def smart_crop_and_split(pil_img):
    """
    Tự động phát hiện và cắt viền thẻ với khoảng đệm an toàn (padding = 25px)
    để đảm bảo không bao giờ bị xém vào Quốc huy, QR code hay chữ.
    """
    img_np = np.array(pil_img)
    h_img, w_img, _ = img_np.shape
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    thresh_inv = cv2.bitwise_not(thresh)
    edges = cv2.Canny(blur, 30, 150)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    combined_mask = cv2.dilate(
        cv2.bitwise_or(edges, thresh_inv), kernel, iterations=1
    )

    contours, _ = cv2.findContours(
        combined_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE
    )

    card_boxes = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area > (w_img * h_img * 0.04):
            x, y, w, h = cv2.boundingRect(cnt)
            aspect_ratio = float(w) / h
            if (
                1.2 <= aspect_ratio <= 1.85
                and w < w_img * 0.98
                and h < h_img * 0.98
            ):
                card_boxes.append((x, y, w, h, area))

    card_boxes = sorted(card_boxes, key=lambda b: b[4], reverse=True)
    filtered = []
    for box in card_boxes:
        x, y, w, h, a = box
        overlap = False
        for fx, fy, fw, fh, fa in filtered:
            dx = max(0, min(x + w, fx + fw) - max(x, fx))
            dy = max(0, min(y + h, fy + fh) - max(y, fy))
            if (dx * dy) > 0.4 * min(w * h, fw * fh):
                overlap = True
                break
        if not overlap:
            filtered.append(box)

    filtered = sorted(filtered, key=lambda b: b[1])

    if len(filtered) >= 1:
        results = []
        # Tăng lề an toàn lên 25px để bảo toàn nguyên vẹn 4 góc thẻ
        pad = 25
        for x, y, w, h, _ in filtered:
            x_pad = max(0, x - pad)
            y_pad = max(0, y - pad)
            w_pad = min(w_img - x_pad, w + (pad * 2))
            h_pad = min(h_img - y_pad, h + (pad * 2))
            cropped = Image.fromarray(
                img_np[y_pad : y_pad + h_pad, x_pad : x_pad + w_pad]
            )
            results.append(cropped)
        return results

    if (h_img / float(w_img)) > 1.15:
        half_h = h_img // 2
        top_half = Image.fromarray(img_np[0:half_h, :])
        bottom_half = Image.fromarray(img_np[half_h:h_img, :])
        return [top_half, bottom_half]

    return [pil_img]


def standardize_card_size(pil_img, target_w=856):
    """Resize giữ nguyên tỷ lệ khung hình thực tế của thẻ, không làm méo hay mất chi tiết."""
    w, h = pil_img.size
    target_h = int(h * (target_w / float(w)))
    return pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)


def create_docx_dynamic(cards_rows, space_between=20):
    """Tạo file Word (.docx) chuẩn lề in A4."""
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

        cell_f = table.cell(0, 0).paragraphs[0].add_run()
        cell_f.add_picture(buf_f, width=Inches(3.37))

        if item["back"]:
            cell_b = table.cell(0, 1).paragraphs[0].add_run()
            cell_b.add_picture(buf_b, width=Inches(3.37))

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
    st.image("https://img.icons8.com/fluency/96/print.png", width=64)
    st.markdown("### ⚙️ Cấu hình Trang In")
    space_val = st.slider("Khoảng cách giữa các hàng (px):", 60, 200, 120, 10)
    draw_border = st.checkbox("Thêm viền mỏng quanh thẻ khi in", value=False)

# 4. GIAO DIỆN CHÍNH
st.markdown(
    '<div class="main-title">📄 Smart Photo Paper A4</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Cắt viền an toàn & ghép trang A4 chuyên nghiệp</div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "📥 Chụp trực tiếp hoặc chọn ảnh từ thiết bị:",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    with st.spinner("✨ Đang tự động xử lý và cắt viền an toàn..."):
        images_cropped = []
        for f in uploaded_files:
            raw_img = load_and_fix_orientation(f)
            extracted_cards = smart_crop_and_split(raw_img)

            for card in extracted_cards:
                card_std = standardize_card_size(card)

                if draw_border:
                    img_border = Image.new(
                        "RGB",
                        (card_std.width + 4, card_std.height + 4),
                        "#cbd5e1",
                    )
                    img_border.paste(card_std, (2, 2))
                    card_std = img_border

                images_cropped.append(card_std)

    st.markdown("---")
    st.markdown("### 📸 Các mặt thẻ đã được bóc tách nguyên vẹn")

    cols = st.columns(len(images_cropped))
    for idx, img in enumerate(images_cropped):
        with cols[idx]:
            st.caption(f"Mặt {idx+1}")
            st.image(img, use_container_width=True)

    swap = False
    if len(images_cropped) >= 2:
        swap = st.checkbox("🔄 Đảo vị trí Trái ↔ Phải (Đổi vị trí 2 mặt)")

    if swap:
        new_imgs = []
        for i in range(0, len(images_cropped), 2):
            pair = images_cropped[i : i + 2]
            if len(pair) == 2:
                new_imgs.extend([pair[1], pair[0]])
            else:
                new_imgs.extend(pair)
        images_cropped = new_imgs

    cards_rows = []
    for i in range(0, len(images_cropped), 2):
        f_img = images_cropped[i]
        b_img = images_cropped[i + 1] if i + 1 < len(images_cropped) else None
        cards_rows.append({"front": f_img, "back": b_img})

    # DỰNG KHUNG A4 CHUẨN TỶ LỆ
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "#ffffff")

    target_w = 510
    gap_x = 50
    x_front = (a4_w - (target_w * 2 + gap_x)) // 2
    x_back = x_front + target_w + gap_x

    current_y = 160
    for row in cards_rows:
        # Giữ tỷ lệ chiều cao theo chiều rộng thực tế
        f_h = int(row["front"].height * (target_w / float(row["front"].width)))
        f_res = row["front"].resize((target_w, f_h), Image.Resampling.LANCZOS)
        canvas.paste(f_res, (x_front, current_y))

        if row["back"]:
            b_h = int(row["back"].height * (target_w / float(row["back"].width)))
            b_res = row["back"].resize((target_w, b_h), Image.Resampling.LANCZOS)
            canvas.paste(b_res, (x_back, current_y))

        current_y += f_h + space_val

    docx_bytes = create_docx_dynamic(
        cards_rows, space_between=space_val
    ).getvalue()

    img_io = io.BytesIO()
    canvas.save(img_io, format="JPEG", quality=92)
    img_bytes = img_io.getvalue()

    st.markdown("---")
    st.markdown("### 📄 Bản xem trước trang A4")
    st.image(canvas, width=500, caption="Trang A4 chuẩn tỉ lệ in ấn")

    st.markdown("### 📥 Tải file kết quả")
    btn1, btn2 = st.columns(2)
    with btn1:
        st.download_button(
            label="📝 Tải file WORD (.docx)",
            data=docx_bytes,
            file_name="Photo_GiayTo_A4.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with btn2:
        st.download_button(
            label="🖼️ Tải file ÁNH (.jpg)",
            data=img_bytes,
            file_name="Photo_GiayTo_A4.jpg",
            mime="image/jpeg",
            use_container_width=True,
        )
