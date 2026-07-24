import io
import gc
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# ---------------------------------------------------------
# CẤU HÌNH TRANG & GIAO DIỆN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Dàn Trang A4 Siêu Tốc",
    layout="wide",
    page_icon="🖨️",
)

st.markdown(
    """
    <style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .stDownloadButton button {
        width: 100%;
        background-color: #16A34A !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 1.05rem !important;
        border-radius: 8px !important;
        padding: 0.75rem 1rem !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .stDownloadButton button:hover {
        background-color: #15803D !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="main-header">🖨️ Xử Lý & Dàn Trang Giấy Tờ A4 Super Fast</div>', unsafe_allow_html=True)
st.caption("Tự động cắt viền • Tự động sắp xếp Mặt Trước/Mặt Sau • Tối ưu 2 bộ/trang A4")

# Quản lý Session State
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "swap_sides" not in st.session_state:
    st.session_state["swap_sides"] = False

def clear_all_files():
    st.session_state["uploader_key"] += 1
    st.session_state["swap_sides"] = False

def toggle_swap():
    st.session_state["swap_sides"] = not st.session_state["swap_sides"]

# ---------------------------------------------------------
# THUẬT TOÁN NHẬN BIẾT MẶT TRƯỚC / MẶT SAU & XỬ LÝ ẢNH
# ---------------------------------------------------------
def is_front_side(image_np):
    """
    Kiểm tra xem ảnh có phải mặt trước hay không dựa vào vị trí ảnh chân dung 
    (Mặt trước CCCD/GPLX luôn có khu vực khuôn mặt ở phía dưới bên trái).
    """
    try:
        h, w, _ = image_np.shape
        # Cắt vùng góc dưới bên trái (nơi chứa ảnh chân dung)
        face_crop = image_np[int(h * 0.35):int(h * 0.95), int(w * 0.05):int(w * 0.45)]
        
        # Dùng Haar Cascade phát hiện khuôn mặt nhanh
        gray_crop = cv2.cvtColor(face_crop, cv2.COLOR_RGB2GRAY)
        face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        faces = face_cascade.detectMultiScale(gray_crop, scaleFactor=1.1, minNeighbors=3, minSize=(30, 30))
        
        if len(faces) > 0:
            return True
            
        # Nếu không thấy mặt (ảnh mờ), kiểm tra vân tay (mặt sau CCCD có 2 ô vân tay màu đỏ/ xám đặc trưng)
        # Hoặc kiểm tra độ đậm vùng mã QR/Mã vạch ở góc dưới
        return False
    except Exception:
        return True

def optimize_image_size(pil_img, max_dim=1600):
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        return pil_img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    return pil_img

def crop_card_fast(image_np):
    h_img, w_img, _ = image_np.shape
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    best_rect, max_area = None, 0
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area > (w_img * h_img * 0.10):
            x, y, w, h = cv2.boundingRect(cnt)
            aspect_ratio = float(w) / h
            if 1.2 <= aspect_ratio <= 1.95 and (w < w_img * 0.98):
                if area > max_area:
                    max_area = area
                    best_rect = (x, y, w, h)

    if best_rect:
        x, y, w, h = best_rect
        return image_np[max(0, y - 4):min(h_img, y + h + 8), max(0, x - 4):min(w_img, x + w + 8)]
    return image_np

def add_card_row_to_word(doc, pil_front, pil_back):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    buf_front, buf_back = io.BytesIO(), io.BytesIO()
    pil_front.save(buf_front, format="PNG", optimize=True)
    pil_back.save(buf_back, format="PNG", optimize=True)
    buf_front.seek(0); buf_back.seek(0)

    cell_f = table.cell(0, 0)
    p_f = cell_f.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f.add_run().add_picture(buf_front, width=Inches(3.35))

    cell_b = table.cell(0, 1)
    p_b = cell_b.paragraphs[0]
    p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b.add_run().add_picture(buf_back, width=Inches(3.35))

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(20)

def create_multi_docx(card_pairs):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    for i, (pil_front, pil_back) in enumerate(card_pairs):
        if i > 0 and i % 2 == 0:
            doc.add_page_break()
        add_card_row_to_word(doc, pil_front, pil_back)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def create_a4_canvas_horizontal(card_pairs_chunk):
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "white")
    card_w, spacing_x = 540, 40
    start_x = (a4_w - (card_w * 2 + spacing_x)) // 2
    y_positions = [180, 920]

    for idx, (pil_front, pil_back) in enumerate(card_pairs_chunk):
        if idx >= 2: break
        y_pos = y_positions[idx]
        ratio_f = card_w / pil_front.width
        f_resized = pil_front.resize((card_w, int(pil_front.height * ratio_f)), Image.Resampling.LANCZOS)
        ratio_b = card_w / pil_back.width
        b_resized = pil_back.resize((card_w, int(pil_back.height * ratio_b)), Image.Resampling.LANCZOS)

        canvas.paste(f_resized, (start_x, y_pos))
        canvas.paste(b_resized, (start_x + card_w + spacing_x, y_pos))

    return canvas

# ---------------------------------------------------------
# BỐ CỤC CHÍNH
# ---------------------------------------------------------
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.subheader("1. Tải Ảnh & Xuất File")
    
    doc_type = st.radio(
        "📋 Loại giấy tờ:",
        ["CCCD", "GPLX"],
        horizontal=True
    )

    col_up_title, col_btn_clear = st.columns([0.65, 0.35])
    with col_up_title:
        st.write("📥 Tải lên tất cả ảnh:")
    with col_btn_clear:
        st.button("🧹 Làm mới", on_click=clear_all_files, use_container_width=True)

    uploaded_files = st.file_uploader(
        "Tải lên tất cả ảnh:",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state['uploader_key']}",
        label_visibility="collapsed"
    )

    download_area = st.container()

with col_right:
    st.subheader("2. Xem Trước Bản In A4")
    preview_area = st.container()

# ---------------------------------------------------------
# THỰC THI XỬ LÝ DỮ LIỆU
# ---------------------------------------------------------
if uploaded_files:
    num_files = len(uploaded_files)

    if num_files < 2:
        col_left.warning("⚠️ Vui lòng chọn ít nhất 2 ảnh.")
    else:
        num_pairs = num_files // 2
        card_pairs = []
        a4_canvases = []

        with col_left:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx, i in enumerate(range(0, num_pairs * 2, 2)):
                status_text.text(f"⏳ Tự động kiểm tra & ghép Bộ {idx + 1}/{num_pairs}...")
                
                raw_img1 = Image.open(uploaded_files[i]).convert("RGB")
                raw_img2 = Image.open(uploaded_files[i + 1]).convert("RGB")

                opt_1 = optimize_image_size(raw_img1)
                opt_2 = optimize_image_size(raw_img2)

                crop_1 = crop_card_fast(np.array(opt_1))
                crop_2 = crop_card_fast(np.array(opt_2))

                # TỰ ĐỘNG PHÁT HIỆN MẶT TRƯỚC / MẶT SAU
                img1_is_front = is_front_side(crop_1)
                img2_is_front = is_front_side(crop_2)

                if not img1_is_front and img2_is_front:
                    # Đảo vị trí nếu ảnh 1 là mặt sau, ảnh 2 mới là mặt trước
                    pil_f = Image.fromarray(crop_2)
                    pil_b = Image.fromarray(crop_1)
                else:
                    pil_f = Image.fromarray(crop_1)
                    pil_b = Image.fromarray(crop_2)

                # Nếu người dùng bấm nút đảo thủ công
                if st.session_state["swap_sides"]:
                    pil_f, pil_b = pil_b, pil_f

                card_pairs.append((pil_f, pil_b))

                del raw_img1, raw_img2, opt_1, opt_2, crop_1, crop_2
                gc.collect()

                progress_bar.progress((idx + 1) / num_pairs)

            status_text.empty()
            progress_bar.empty()

        for i in range(0, len(card_pairs), 2):
            chunk = card_pairs[i:i + 2]
            canvas = create_a4_canvas_horizontal(chunk)
            a4_canvases.append(canvas)

        docx_io = create_multi_docx(card_pairs)

        with download_area:
            st.success(f"⚡ Đã ghép xong **{len(card_pairs)} bộ**!")

            # Nút Đảo Mặt Thủ Công (Nhanh 1-click)
            st.button("🔄 Đảo Vị Trí Mặt Trước ↔ Mặt Sau", on_click=toggle_swap, use_container_width=True)

            st.download_button(
                label=f"📝 TẢI FILE WORD NGAY (.docx)",
                data=docx_io,
                file_name=f"{doc_type}_In_A4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            
            img_io = io.BytesIO()
            a4_canvases[0].save(img_io, format="PNG", optimize=True)
            img_io.seek(0)

            st.download_button(
                label="🖼️ TẢI ÁNH PNG TRANG 1",
                data=img_io,
                file_name=f"{doc_type}_Trang1.png",
                mime="image/png",
            )

        with preview_area:
            if len(a4_canvases) > 1:
                tabs = st.tabs([f"Trang #{i + 1}" for i in range(len(a4_canvases))])
                for i, tab in enumerate(tabs):
                    with tab:
                        st.image(a4_canvases[i], use_container_width=True)
            else:
                st.image(a4_canvases[0], use_container_width=True)
