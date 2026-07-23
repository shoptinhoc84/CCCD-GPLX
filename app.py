import io
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# 1. CẤU HÌNH TRANG & CUSTOM CSS
st.set_page_config(
    page_title="Photo Doc Pro - Ghép Giấy Tờ A4",
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


# 2. XỬ LÝ & TÁCH ẢNH DÍNH LIỀN
def resize_image_if_large(pil_img, max_dim=1200):
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        return pil_img.resize(
            (int(w * scale), int(h * scale)), Image.Resampling.LANCZOS
        )
    return pil_img


def crop_card(image_np):
    """Cắt viền tự động, không lẹm chữ."""
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
        x, y = max(0, x - 8), max(0, y - 8)
        w, h = min(w_img - x, w + 16), min(h_img - y, h + 16)
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
            x, y = max(0, x - 5), max(0, y - 5)
            w, h = min(w_img - x, w + 10), min(h_img - y, h + 10)
            return image_np[y : y + h, x : x + w]

    return image_np


def process_uploaded_image(pil_img):
    """
    Tự động phát hiện nếu ảnh tải lên là dạng 2 mặt dính liền dọc
    (như màn hình VNeID/GPLX) thì tách thành 2 ảnh rời, ngược lại cắt viền bình thường.
    """
    w, h = pil_img.size
    aspect_ratio = h / float(w)

    # Nếu ảnh là dạng dọc dài (tỷ lệ cao/rộng > 1.1) -> Bức ảnh chứa 2 thẻ dính liền
    if aspect_ratio > 1.1:
        img_np = np.array(pil_img)
        gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        _, thresh = cv2.threshold(
            blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        contours, _ = cv2.findContours(
            thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE
        )

        cards = []
        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area > (w * h * 0.08):
                x, y, cw, ch = cv2.boundingRect(cnt)
                c_ratio = float(cw) / ch
                if 1.2 <= c_ratio <= 1.9 and cw < w * 0.98:
                    cards.append((y, x, cw, ch, cnt))

        # Nếu tìm thấy 2 khung thẻ chữ nhật riêng biệt
        if len(cards) >= 2:
            # Sắp xếp từ trên xuống dưới
            cards.sort(key=lambda item: item[0])
            top_card = cards[0]
            bottom_card = cards[1]

            y1, x1, w1, h1 = top_card[0], top_card[1], top_card[2], top_card[3]
            y2, x2, w2, h2 = (
                bottom_card[0],
                bottom_card[1],
                bottom_card[2],
                bottom_card[3],
            )

            img_top = Image.fromarray(
                img_np[
                    max(0, y1 - 5) : min(h, y1 + h1 + 5),
                    max(0, x1 - 5) : min(w, x1 + w1 + 5),
                ]
            )
            img_bottom = Image.fromarray(
                img_np[
                    max(0, y2 - 5) : min(h, y2 + h2 + 5),
                    max(0, x2 - 5) : min(w, x2 + w2 + 5),
                ]
            )
            return [img_top, img_bottom]
        else:
            # Fallback: Tách đôi ảnh theo chiều ngang (nửa trên và nửa dưới)
            half_h = h // 2
            top_half = img_np[0:half_h, :]
            bottom_half = img_np[half_h:h, :]

            top_cropped = crop_card(top_half)
            bottom_cropped = crop_card(bottom_half)

            return [
                Image.fromarray(top_cropped),
                Image.fromarray(bottom_cropped),
            ]
    else:
        # Ảnh thẻ đơn lẻ bình thường
        cropped_np = crop_card(np.array(pil_img))
        return [Image.fromarray(cropped_np)]


def create_docx_dynamic(cards_rows, space_between=20):
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
    st.markdown("---")
    st.info(
        "💡 **Tính năng mới:** Hỗ trợ tải **ảnh chụp VNeID / Bản quét dính liền 2 mặt dọc**, hệ thống tự động tách làm 2 thẻ rời!"
    )

# 4. GIAO DIỆN CHÍNH
st.markdown(
    '<div class="main-title">📄 Smart Photo Paper A4</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Tự động cắt viền, tách ảnh dính 2 mặt & ghép A4</div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "📥 Tải ảnh lên (Ảnh rời hoặc Ảnh chụp màn hình VNeID 2 mặt):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    with st.spinner("✨ Đang tự động tách mặt thẻ & cắt viền..."):
        images_cropped = []
        for f in uploaded_files:
            raw_img = resize_image_if_large(Image.open(f).convert("RGB"))
            # Tự nhận diện tách 2 mặt nếu là ảnh dọc dài (VNeID/GPLX điện tử)
            extracted_cards = process_uploaded_image(raw_img)

            for card in extracted_cards:
                if draw_border:
                    img_border = Image.new(
                        "RGB", (card.width + 4, card.height + 4), "#cbd5e1"
                    )
                    img_border.paste(card, (2, 2))
                    card = img_border
                images_cropped.append(card)

    st.markdown("---")
    st.markdown("### 📸 Các mặt thẻ đã được tách & cắt viền")
    cols = st.columns(len(images_cropped))
    for idx, img in enumerate(images_cropped):
        with cols[idx]:
            st.caption(f"Mặt {idx+1}")
            st.image(img, use_container_width=True)

    # Nút đổi vị trí
    swap = False
    if len(images_cropped) >= 2:
        swap = st.checkbox("🔄 Đảo vị trí Trái ↔ Phải")

    if swap:
        new_imgs = []
        for i in range(0, len(images_cropped), 2):
            pair = images_cropped[i : i + 2]
            if len(pair) == 2:
                new_imgs.extend([pair[1], pair[0]])
            else:
                new_imgs.extend(pair)
        images_cropped = new_imgs

    # Ghép hàng 2 mặt song song
    cards_rows = []
    for i in range(0, len(images_cropped), 2):
        f_img = images_cropped[i]
        b_img = images_cropped[i + 1] if i + 1 < len(images_cropped) else None
        cards_rows.append({"front": f_img, "back": b_img})

    # TẠO CANVAS A4
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "#ffffff")
    target_w = 510
    gap_x = 50
    x_front = (a4_w - (target_w * 2 + gap_x)) // 2
    x_back = x_front + target_w + gap_x

    current_y = 160
    for row in cards_rows:
        f_res = row["front"].resize(
            (
                target_w,
                int(row["front"].height * (target_w / row["front"].width)),
            )
        )
        canvas.paste(f_res, (x_front, current_y))

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
