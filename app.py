import io
import gc
import cv2
import re
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# Sử dụng Pytesseract nhẹ nhàng, không load model AI nặng vào RAM
import pytesseract

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

st.markdown('<div class="main-header">🖨️ Xử Lý Giấy Tờ & Tự Động Điền Đơn Học GPLX</div>', unsafe_allow_html=True)

# Quản lý Session State
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "swap_dict" not in st.session_state:
    st.session_state["swap_dict"] = {}

def clear_all_files():
    st.session_state["uploader_key"] += 1
    st.session_state["swap_dict"] = {}

def toggle_swap_pair(pair_index):
    current_state = st.session_state["swap_dict"].get(pair_index, False)
    st.session_state["swap_dict"][pair_index] = not current_state

# ---------------------------------------------------------
# THUẬT TOÁN XỬ LÝ ẢNH CHUNG (TAB 1 & TAB 2)
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

def create_single_cropped_docx(pil_cropped_img):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    buf = io.BytesIO()
    pil_cropped_img.save(buf, format="PNG", optimize=True)
    buf.seek(0)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(buf, width=Inches(3.5))

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

def crop_vneid_combined_block(pil_img):
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
            if 1.2 <= aspect_ratio <= 1.95 and w < w_img * 0.98:
                card_boxes.append((x, y, w, h))

    filtered_boxes = []
    card_boxes = sorted(card_boxes, key=lambda b: b[1])
    for box in card_boxes:
        if not any(abs(box[1] - f[1]) < 30 for f in filtered_boxes):
            filtered_boxes.append(box)

    if len(filtered_boxes) >= 2:
        b1, b2 = filtered_boxes[0], filtered_boxes[1]
        min_x = min(b1[0], b2[0])
        max_x = max(b1[0] + b1[2], b2[0] + b2[2])
        min_y = min(b1[1], b2[1])
        max_y = max(b1[1] + b1[3], b2[1] + b2[3])

        crop_np = img_np[
            max(0, min_y - 6) : min(h_img, max_y + 14), 
            max(0, min_x - 6) : min(w_img, max_x + 6)
        ]
        return Image.fromarray(crop_np)

    y1, y2 = int(h_img * 0.11), int(h_img * 0.67)
    x1, x2 = int(w_img * 0.035), int(w_img * 0.965)
    
    crop_np = img_np[y1:y2, x1:x2]
    return Image.fromarray(crop_np)

# ---------------------------------------------------------
# THUẬT TOÁN TAB 3: TRÍCH XUẤT CỰC NHẸ (KHÔNG BỊ TRÀN RAM)
# ---------------------------------------------------------
def process_auto_batch_ocr(list_uploaded_files):
    data = {
        "ho_ten": "", "ngay_sinh": "", "so_cccd": "", "ngay_cap_cccd": "",
        "noi_cap_cccd": "Cục Cảnh sát quản lý hành chính về trật tự xã hội",
        "so_gplx": "", "hang_gplx": "", "noi_cap_gplx": "", "ngay_cap_gplx": "",
        "hang_dang_ky": "A1", "vi_pham": "Không", "sdt": ""
    }

    logs = []

    for idx, file in enumerate(list_uploaded_files):
        img = Image.open(file).convert("RGB")
        
        # Tối ưu kích thước ảnh trước khi OCR để tốn cực ít bộ nhớ
        img_opt = optimize_image_size(img, max_dim=1200)

        try:
            full_text = pytesseract.image_to_string(img_opt, lang='eng+vie').upper()
        except Exception:
            try:
                full_text = pytesseract.image_to_string(img_opt, lang='eng').upper()
            except Exception:
                full_text = ""

        lines = [line.strip() for line in full_text.split('\n') if line.strip()]

        is_gplx = ("GIẤY PHÉP LÁI XE" in full_text) or ("DRIVER" in full_text) or ("CLASS" in full_text)
        is_cccd = ("CĂN CƯỚC" in full_text) or ("CITIZEN" in full_text) or ("SOCIALIST" in full_text) or ("CAN CUOC" in full_text)

        if is_gplx:
            logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **GPLX**")
            gplx_match = re.search(r'\b\d{12}\b', full_text)
            if gplx_match:
                data["so_gplx"] = gplx_match.group(0)

            hang_match = re.search(r'(?:HẠNG|CLASS)[:\s]*([A-Z0-9]+)', full_text)
            if hang_match:
                data["hang_gplx"] = hang_match.group(1)

            dates_g = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
            if dates_g:
                data["ngay_cap_gplx"] = dates_g[0]

        else:
            if is_cccd:
                logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **CCCD / Căn cước**")
            else:
                logs.append(f"📸 Ảnh #{idx+1}: Quét dữ liệu bổ sung...")

            id_match = re.search(r'\b\d{12}\b', full_text)
            if id_match and not data["so_cccd"]:
                data["so_cccd"] = id_match.group(0)

            date_matches = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
            if date_matches:
                if not data["ngay_sinh"]:
                    data["ngay_sinh"] = date_matches[0]
                if len(date_matches) > 1 and not data["ngay_cap_cccd"]:
                    data["ngay_cap_cccd"] = date_matches[1]

            # Rút trích Họ tên
            for line in lines:
                clean_l = line.strip()
                if clean_l.isupper() and len(clean_l) > 5 and not any(char.isdigit() for char in clean_l):
                    if "CONG HOA" not in clean_l and "CAN CUOC" not in clean_l and "VIET NAM" not in clean_l:
                        if not data["ho_ten"]:
                            data["ho_ten"] = clean_l
                            break

        # Giải phóng bộ nhớ liên tục
        del img, img_opt
        gc.collect()

    return data, logs

def generate_don_de_nghi_docx(d):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_head.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r1.bold = True; r1.font.size = Pt(12)
    r2 = p_head.add_run("Độc lập – Tự do – Hạnh phúc\n")
    r2.bold = True; r2.font.size = Pt(13)
    p_head.paragraph_format.space_after = Pt(12)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_title.add_run("ĐƠN ĐỀ NGHỊ\nHỌC, DỰ SÁT HẠCH ĐỂ CẤP GIẤY PHÉP LÁI XE\n")
    rt.bold = True; rt.font.size = Pt(14)
    p_title.paragraph_format.space_after = Pt(12)

    p_kg = doc.add_paragraph()
    p_kg.paragraph_format.left_indent = Inches(0.5)
    p_kg.add_run("Kính gửi:\n").bold = True
    p_kg.add_run("- Sở Xây dựng tỉnh Vĩnh Long;\n")
    p_kg.add_run("  Phòng Cảnh sát giao thông - Công an tỉnh Vĩnh Long;\n")
    p_kg.add_run("- Trung tâm GDNN và SHLX Nguyễn Trình.\n")
    p_kg.paragraph_format.space_after = Pt(10)

    p_body = doc.add_paragraph()
    p_body.paragraph_format.line_spacing = 1.25
    
    p_body.add_run("Tôi là (CHỮ IN HOA): ").font.size = Pt(12)
    r_name = p_body.add_run(f"{d['ho_ten'].upper()}\n")
    r_name.bold = True; r_name.font.size = Pt(12)

    p_body.add_run(f"Ngày tháng năm sinh: {d['ngay_sinh']}\n").font.size = Pt(12)
    p_body.add_run(f"Số Căn cước công dân/Căn cước hoặc Hộ chiếu: {d['so_cccd']}\n").font.size = Pt(12)
    p_body.add_run(f"Cấp ngày: {d['ngay_cap_cccd']} ; nơi cấp: {d['noi_cap_cccd']}\n").font.size = Pt(12)
    
    p_body.add_run(f"Đã có giấy phép lái xe số: {d['so_gplx']} hạng {d['hang_gplx']} do {d['noi_cap_gplx']} cấp ngày: {d['ngay_cap_gplx']}\n").font.size = Pt(12)
    
    p_body.add_run("Đề nghị cho tôi được học, dự sát hạch để cấp giấy phép lái xe hạng: ").font.size = Pt(12)
    r_h = p_body.add_run(f"{d['hang_dang_ky']}\n")
    r_h.bold = True; r_h.font.size = Pt(12)

    is_vp_co = "[ X ] Có   [   ] Không" if d['vi_pham'] == "Có" else "[   ] Có   [ X ] Không"
    p_body.add_run(f"Vi phạm hành chính trong lĩnh vực giao thông đường bộ với hình thức tước quyền sử dụng giấy phép lái xe: {is_vp_co}\n\n").font.size = Pt(12)

    p_body.add_run("Xin gửi kèm theo:\n").bold = True
    p_body.add_run("- 01 Giấy khám sức khỏe của người lái xe do cơ sở y tế có thẩm quyền cấp theo quy định.\n")
    p_body.add_run("- 01 bản sao thẻ căn cước công dân hoặc hộ chiếu còn thời hạn.\n")
    p_body.add_run("- 01 bản sao giấy phép lái xe bằng thẻ PET (nếu có).\n")
    p_body.add_run("- 06 ảnh thẻ 3×4 phông nền xanh (Không đeo kính, không để tóc chờm vào mắt).\n\n")

    p_body.add_run("Tôi xin cam đoan những điều ghi trên là đúng sự thật, nếu sai tôi xin hoàn toàn chịu trách nhiệm.\n").font.size = Pt(12)

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.add_run("Vĩnh Long, ngày ... tháng ... năm 20...\n").font.size = Pt(12)
    p_sign.add_run("NGƯỜI LÀM ĐƠN\n").bold = True
    p_sign.add_run("(Ký và ghi rõ họ, tên)\n\n\n\n").italic = True
    
    r_sub_name = p_sign.add_run(f"{d['ho_ten'].upper()}\n")
    r_sub_name.bold = True
    p_sign.add_run(f"Số điện thoại: {d['sdt']}").font.size = Pt(11)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# ---------------------------------------------------------
# CẤU TRÚC GIAO DIỆN TABS
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs([
    "🖨️ Dàn Trang A4 (Hàng Loạt)", 
    "✂️ Cắt Ảnh Khung GPLX VNeID", 
    "📝 Đơn Đề Nghị Học GPLX (Tự Động)"
])

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
# TAB 2: CẮT KHUNG VNeID CHUẨN MẪU
# =========================================================
with tab2:
    st.subheader("✂️ Cắt Trọn Khung Bằng Lái Xe GPLX VNeID")
    st.caption("Tải lên ảnh màn hình VNeID, hệ thống sẽ tự động cắt khung chứa 2 thẻ GPLX.")

    col_vneid_up, col_vneid_res = st.columns([1, 1], gap="medium")

    with col_vneid_up:
        uploaded_vneid = st.file_uploader(
            "📥 Tải ảnh chụp màn hình GPLX VNeID:",
            type=["jpg", "jpeg", "png"],
            key="vneid_uploader"
        )
        if uploaded_vneid:
            raw_vneid_img = Image.open(uploaded_vneid).convert("RGB")
            st.image(raw_vneid_img, caption="Ảnh gốc VNeID", use_container_width=True)

    with col_vneid_res:
        if uploaded_vneid:
            with st.spinner("⏳ Đang nhận diện và cắt khung ảnh..."):
                cropped_gplx_img = crop_vneid_combined_block(raw_vneid_img)

                buf = io.BytesIO()
                cropped_gplx_img.save(buf, format="PNG", optimize=True)
                buf_png = buf.getvalue()

                docx_cropped_io = create_single_cropped_docx(cropped_gplx_img)

                st.markdown("### 📥 Tải về kết quả:")
                
                col_dl_word, col_dl_png = st.columns(2)
                with col_dl_word:
                    st.download_button(
                        label="📝 TẢI FILE WORD (.docx)",
                        data=docx_cropped_io,
                        file_name="GPLX_VNeID_Cropped.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col_dl_png:
                    st.download_button(
                        label="🖼️ TẢI ẢNH PNG (.png)",
                        data=buf_png,
                        file_name="GPLX_VNeID_Cropped.png",
                        mime="image/png",
                        use_container_width=True
                    )

                st.markdown("---")
                st.markdown("### 👁️ Xem trước ảnh đã cắt:")
                st.image(cropped_gplx_img, caption="Khung ảnh GPLX 2 mặt", use_container_width=True)


# =========================================================
# TAB 3: TRÍCH XUẤT CỰC NHẸ BẰNG TESSERACT (KHÔNG SẬP)
# =========================================================
with tab3:
    st.subheader("📝 Tự Động Phân Loại & Điền File Word 'Đơn Đề Nghị Học GPLX'")
    st.caption("Tải lên tất cả ảnh giấy tờ (CCCD / GPLX / Ảnh VNeID...), hệ thống sẽ tự động phân loại và trích xuất thông tin.")

    col3_up, col3_form = st.columns([0.45, 0.55], gap="medium")

    with col3_up:
        st.markdown("#### 1. Tải tất cả ảnh giấy tờ vào đây:")
        up_batch = st.file_uploader(
            "📥 Chọn hoặc kéo thả tất cả ảnh (CCCD, GPLX...):", 
            type=["jpg", "jpeg", "png"], 
            accept_multiple_files=True,
            key="ocr_batch_uploader"
        )
        
        btn_ocr = st.button("🔍 Đọc & Phân Loại Tự Động", use_container_width=True, type="primary")

        if "form_data" not in st.session_state:
            st.session_state["form_data"] = {
                "ho_ten": "", "ngay_sinh": "", "so_cccd": "", "ngay_cap_cccd": "",
                "noi_cap_cccd": "Cục Cảnh sát quản lý hành chính về trật tự xã hội",
                "so_gplx": "", "hang_gplx": "", "noi_cap_gplx": "", "ngay_cap_gplx": "",
                "hang_dang_ky": "A1", "vi_pham": "Không", "sdt": ""
            }

        if btn_ocr:
            if not up_batch:
                st.warning("⚠️ Vui lòng tải lên ít nhất 1 ảnh giấy tờ.")
            else:
                with st.spinner("⏳ Đang nhận diện & phân loại giấy tờ..."):
                    extracted, logs = process_auto_batch_ocr(up_batch)
                    
                    for k, v in extracted.items():
                        if v:
                            st.session_state["form_data"][k] = v
                    
                    st.success("✅ Đã hoàn tất phân loại!")
                    for log in logs:
                        st.info(log)

    with col3_form:
        st.markdown("#### 2. Kiểm tra thông tin điền tự động:")
        
        fd = st.session_state["form_data"]
        
        c1, c2 = st.columns(2)
        with c1:
            fd["ho_ten"] = st.text_input("Họ và tên (IN HOA):", value=fd["ho_ten"])
            fd["so_cccd"] = st.text_input("Số CCCD/CMND:", value=fd["so_cccd"])
            fd["ngay_cap_cccd"] = st.text_input("CCCD Cấp ngày:", value=fd["ngay_cap_cccd"])
        with c2:
            fd["ngay_sinh"] = st.text_input("Ngày tháng năm sinh:", value=fd["ngay_sinh"])
            fd["noi_cap_cccd"] = st.text_input("Nơi cấp CCCD:", value=fd["noi_cap_cccd"])
            fd["sdt"] = st.text_input("Số điện thoại:", value=fd["sdt"])

        st.markdown("---")
        st.markdown("**Thông tin GPLX đã có (nếu tìm thấy trong ảnh):**")
        c3, c4 = st.columns(2)
        with c3:
            fd["so_gplx"] = st.text_input("Số GPLX hiện có:", value=fd["so_gplx"])
            fd["noi_cap_gplx"] = st.text_input("Do nơi nào cấp:", value=fd["noi_cap_gplx"])
        with c4:
            fd["hang_gplx"] = st.text_input("Hạng GPLX hiện có:", value=fd["hang_gplx"])
            fd["ngay_cap_gplx"] = st.text_input("GPLX Cấp ngày:", value=fd["ngay_cap_gplx"])

        st.markdown("---")
        c5, c6 = st.columns(2)
        with c5:
            fd["hang_dang_ky"] = st.selectbox("Đề nghị học, dự sát hạch hạng:", ["A1", "A2", "A", "B1", "B2", "B", "C1", "C"], index=0)
        with c6:
            fd["vi_pham"] = st.radio("Có bị tước quyền sử dụng GPLX không?", ["Không", "Có"], horizontal=True)

        st.session_state["form_data"] = fd

        st.markdown("<br>", unsafe_allow_html=True)
        
        docx_don_io = generate_don_de_nghi_docx(fd)
        
        st.download_button(
            label="📝 TẢI FILE WORD 'ĐƠN ĐỀ NGHỊ HỌC SÁT HẠCH' (.docx)",
            data=docx_don_io,
            file_name=f"Don_De_Nghi_Hoc_GPLX_{fd['ho_ten'].replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
