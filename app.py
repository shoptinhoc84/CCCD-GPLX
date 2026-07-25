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
    page_title="Hệ Thống Xử Lý & Dàn Trang Giấy Tờ",
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

# Quản lý Session State
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "swap_dict" not in st.session_state:
    st.session_state["swap_dict"] = {}  # Lưu trạng thái đảo mặt riêng cho từng bộ

def clear_all_files():
    st.session_state["uploader_key"] += 1
    st.session_state["swap_dict"] = {}

def toggle_swap_pair(pair_index):
    """Đảo vị trí riêng cho duy nhất 1 bộ chỉ định."""
    current_state = st.session_state["swap_dict"].get(pair_index, False)
    st.session_state["swap_dict"][pair_index] = not current_state

# ---------------------------------------------------------
# THUẬT TOÁN XỬ LÝ ẢNH CHUNG & DÀN TRANG A4
# ---------------------------------------------------------
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
# THUẬT TOÁN BỐ CỤC CHUYÊN CẮT 2 MẶT TỪ ẢNH MẪU VNeID
# ---------------------------------------------------------
def crop_vneid_gplx_two_sides(pil_img):
    """
    Tự động tìm kiếm vùng 2 thẻ GPLX xếp dọc trên màn hình VNeID
    và tách thành mặt trước, mặt sau độc lập.
    """
    img_np = np.array(pil_img)
    h_img, w_img, _ = img_np.shape
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    
    card_boxes = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area > (w_img * h_img * 0.05):
            x, y, w, h = cv2.boundingRect(cnt)
            aspect_ratio = float(w) / h
            # Tỷ lệ thẻ GPLX chuẩn (~ 1.4 - 1.75)
            if 1.2 <= aspect_ratio <= 1.95 and w < w_img * 0.98:
                card_boxes.append((x, y, w, h))

    # Lọc trùng lặp bằng cách ghép các khung tương đương
    filtered_boxes = []
    card_boxes = sorted(card_boxes, key=lambda b: b[1])  # Sắp xếp từ trên xuống dưới
    for box in card_boxes:
        if not any(abs(box[1] - f[1]) < 30 for f in filtered_boxes):
            filtered_boxes.append(box)

    # Nếu tìm thấy đúng từ 2 thẻ trở lên
    if len(filtered_boxes) >= 2:
        # Lấy khung trên (Mặt trước) và khung dưới (Mặt sau)
        b1, b2 = filtered_boxes[0], filtered_boxes[1]
        
        # Cắt lề 2px để nét hơn
        front_crop = img_np[max(0, b1[1]-2):min(h_img, b1[1]+b1[3]+2), max(0, b1[0]-2):min(w_img, b1[0]+b1[2]+2)]
        back_crop = img_np[max(0, b2[1]-2):min(h_img, b2[1]+b2[3]+2), max(0, b2[0]-2):min(w_img, b2[0]+b2[2]+2)]
        
        return Image.fromarray(front_crop), Image.fromarray(back_crop)
    
    # Dự phòng: Nếu thuật toán Contour không bắt được viền (do nền sáng quá/thiếu tương phản),
    # dùng thuật toán cắt cố định theo vùng hiển thị chuẩn VNeID GPLX.
    y1, y2 = int(h_img * 0.11), int(h_img * 0.38)
    y3, y4 = int(h_img * 0.39), int(h_img * 0.65)
    x1, x2 = int(w_img * 0.04), int(w_img * 0.96)
    
    front_crop = img_np[y1:y2, x1:x2]
    back_crop = img_np[y3:y4, x1:x2]
    
    return Image.fromarray(front_crop), Image.fromarray(back_crop)


# ---------------------------------------------------------
# CẤU TRÚC GIAO DIỆN TABS
# ---------------------------------------------------------
tab1, tab2 = st.tabs(["🖨️ Dàn Trang A4 (Hàng Loạt)", "✂️ Tách 2 Mặt GPLX VNeID"])

# =========================================================
# TAB 1: DÀN TRANG A4 TỰ ĐỘNG
# =========================================================
with tab1:
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
        st.subheader("2. Xem Trước Bản In A4 & Tuỳ Chỉnh")
        preview_area = st.container()

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
                    status_text.text(f"⏳ Đang xử lý Bộ {idx + 1}/{num_pairs}...")
                    
                    raw_img1 = Image.open(uploaded_files[i]).convert("RGB")
                    raw_img2 = Image.open(uploaded_files[i + 1]).convert("RGB")

                    opt_1 = optimize_image_size(raw_img1)
                    opt_2 = optimize_image_size(raw_img2)

                    crop_1 = crop_card_fast(np.array(opt_1))
                    crop_2 = crop_card_fast(np.array(opt_2))

                    pil_f = Image.fromarray(crop_1)
                    pil_b = Image.fromarray(crop_2)

                    if st.session_state["swap_dict"].get(idx, False):
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
                    label="🖼️ TẢI ẢNH PNG TRANG 1",
                    data=img_io,
                    file_name=f"{doc_type}_Trang1.png",
                    mime="image/png",
                )

            with preview_area:
                st.write("🔄 **Tuỳ chỉnh vị trí cho riêng từng bộ (nếu bị ngược):**")
                btn_cols = st.columns(min(len(card_pairs), 4))
                for p_idx in range(len(card_pairs)):
                    is_swapped = st.session_state["swap_dict"].get(p_idx, False)
                    btn_label = f"🔄 Đảo Bộ {p_idx + 1}" + (" (Đã đảo)" if is_swapped else "")
                    btn_cols[p_idx % 4].button(
                        btn_label, 
                        key=f"btn_swap_{p_idx}", 
                        on_click=toggle_swap_pair, 
                        args=(p_idx,),
                        use_container_width=True
                    )

                st.markdown("---")

                if len(a4_canvases) > 1:
                    tabs = st.tabs([f"Trang #{i + 1}" for i in range(len(a4_canvases))])
                    for i, tab in enumerate(tabs):
                        with tab:
                            st.image(a4_canvases[i], use_container_width=True)
                else:
                    st.image(a4_canvases[0], use_container_width=True)


# =========================================================
# TAB 2: CẮT TÁCH 2 MẶT TỪ ẢNH VNeID GPLX (TÍNH NĂNG MỚI)
# =========================================================
with tab2:
    st.subheader("✂️ Cắt Tách Mặt Trước & Mặt Sau Từ Ảnh GPLX VNeID")
    st.caption("Tải lên ảnh chụp màn hình ứng dụng VNeID / Dịch vụ công (chứa 2 mặt xếp dọc), hệ thống sẽ tự động tách rời 2 mặt.")

    col_vneid_up, col_vneid_res = st.columns([1, 1], gap="medium")

    with col_vneid_up:
        uploaded_vneid = st.file_uploader(
            " Tải ảnh chụp màn hình GPLX VNeID:",
            type=["jpg", "jpeg", "png"],
            key="vneid_uploader"
        )
        if uploaded_vneid:
            raw_vneid_img = Image.open(uploaded_vneid).convert("RGB")
            st.image(raw_vneid_img, caption="Ảnh gốc đã tải lên", use_container_width=True)

    with col_vneid_res:
        if uploaded_vneid:
            with st.spinner("⏳ Đang nhận diện và cắt mặt trước & sau..."):
                img_front, img_back = crop_vneid_gplx_two_sides(raw_vneid_img)

                col_f, col_b = st.columns(2)
                
                # Mặt trước
                with col_f:
                    st.markdown("**1. Mặt trước**")
                    st.image(img_front, use_container_width=True)
                    buf_f = io.BytesIO()
                    img_front.save(buf_f, format="PNG")
                    st.download_button(
                        label="⬇️ Tải Mặt Trước",
                        data=buf_f.getvalue(),
                        file_name="GPLX_MatTruoc.png",
                        mime="image/png",
                        use_container_width=True
                    )

                # Mặt sau
                with col_b:
                    st.markdown("**2. Mặt sau**")
                    st.image(img_back, use_container_width=True)
                    buf_b = io.BytesIO()
                    img_back.save(buf_b, format="PNG")
                    st.download_button(
                        label="⬇️ Tải Mặt Sau",
                        data=buf_b.getvalue(),
                        file_name="GPLX_MatSau.png",
                        mime="image/png",
                        use_container_width=True
                    )
