import io
import cv2
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# Thử import easyocr, nếu chưa cài sẽ báo hướng dẫn
try:
    import easyocr
    # Khởi tạo reader (cached để không bị load lại nhiều lần)
    @st.cache_resource
    def load_ocr():
        return easyocr.Reader(['vi', 'en'], gpu=False)
    reader = load_ocr()
except ImportError:
    reader = None

st.set_page_config(
    page_title="Hệ thống tự động phân loại & Ghép giấy tờ A4",
    layout="wide",
    page_icon="🖨️",
)

st.title("🖨️ Phân loại Smart & Ghép giấy tờ in A4")
st.write(
    "Thả **tất cả ảnh** (CCCD, GPLX, mặt trước, mặt sau lẫn lộn) vào 1 ô duy nhất. "
    "Hệ thống sẽ tự nhận diện loại giấy tờ, tự lọc viền và xếp theo đúng cặp!"
)

if reader is None:
    st.error("⚠️ Vui lòng cài đặt thư viện `easyocr` bằng lệnh: `pip install easyocr` để sử dụng tính năng tự nhận dạng.")

def crop_card(image_np):
    """Tách nền và cắt viền giấy tờ."""
    h_img, w_img, _ = image_np.shape
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

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
        return image_np[max(0, y-3):min(h_img, y+h+6), max(0, x-3):min(w_img, x+w+6)]

    return image_np

def classify_and_detect_side(image_np):
    """
    Tự động nhận dạng:
    - Loại giấy tờ: 'CCCD' hoặc 'GPLX'
    - Mặt: 'front' hoặc 'back'
    """
    if reader is None:
        return "UNKNOWN", "front"

    # Đọc text trong ảnh bằng EasyOCR
    results = reader.readtext(image_np, detail=0)
    text_combined = " ".join(results).lower()

    # Logic nhận diện Loại giấy tờ
    doc_type = "CCCD"
    if "bằng lái" in text_text or "giấy phép lái xe" in text_combined or "driving licence" in text_combined or "gplx" in text_combined:
        doc_type = "GPLX"
    elif "căn cước" in text_combined or "mã số" in text_combined or "chứng minh" in text_combined or "identity card" in text_combined:
        doc_type = "CCCD"

    # Logic nhận diện Mặt trước / Mặt sau
    side = "front"
    # Dấu hiệu mặt sau: Có thông tin đặc điểm nhân dạng, ngày cấp, ngón trỏ, hoặc dấu vân tay / mã QR lớn
    if any(k in text_text for k in ["đặc điểm nhân dạng", "ngày, tháng, năm cấp", "có giá trị đến", "thâm quyến", "cục trưởng"]):
        side = "back"
    elif any(k in text_text for k in ["họ và tên", "ngày sinh", "giới tính", "quốc tịch", "date of birth"]):
        side = "front"

    return doc_type, side

def create_multi_docx(grouped_pairs):
    """Tạo file Word (.docx) gom tất cả CCCD & GPLX đã phân loại."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    first_page = True
    for doc_type, pairs in grouped_pairs.items():
        for pil_front, pil_back in pairs:
            if not first_page:
                doc.add_page_break()
            first_page = False

            # Tiêu đề nhóm
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(f"GIẤY TỜ: {doc_type.upper()}")
            r_title.bold = True
            r_title.font.size = Pt(14)

            # Đưa ảnh vào BytesIO
            buf_f, buf_b = io.BytesIO(), io.BytesIO()
            pil_front.save(buf_f, format="PNG")
            pil_back.save(buf_b, format="PNG")
            buf_f.seek(0); buf_b.seek(0)

            # Mặt trước
            p1 = doc.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.add_run().add_picture(buf_f, width=Inches(3.37))

            doc.add_paragraph().paragraph_format.space_after = Pt(12)

            # Mặt sau
            p2 = doc.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.add_run().add_picture(buf_b, width=Inches(3.37))

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# --- GIAO DIỆN CHÍNH ---
uploaded_files = st.file_uploader(
    "📥 Tải lên TẤT CẢ ảnh giấy tờ (CCCD + GPLX lẫn lộn):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) < 2:
        st.warning("⚠️ Vui lòng chọn ít nhất 2 ảnh để ghép.")
    else:
        with st.spinner("🤖 Đang quét OCR, tự động nhận dạng loại giấy tờ và ghép mặt trước/sau..."):
            processed_data = []

            for file in uploaded_files:
                img_raw = Image.open(file).convert("RGB")
                np_img = np.array(img_raw)
                
                # Cắt viền
                cropped_np = crop_card(np_img)
                cropped_pil = Image.fromarray(cropped_np)
                
                # Nhận dạng OCR
                doc_type, side = classify_and_detect_side(cropped_np)
                
                processed_data.append({
                    'file_name': file.name,
                    'image': cropped_pil,
                    'type': doc_type,
                    'side': side
                })

        # Phân loại thành các nhóm
        grouped_docs = {}
        # Tách danh sách theo loại giấy tờ
        for item in processed_data:
            dt = item['type']
            if dt not in grouped_docs:
                grouped_docs[dt] = {'fronts': [], 'backs': []}
            
            if item['side'] == 'front':
                grouped_docs[dt]['fronts'].append(item['image'])
            else:
                grouped_docs[dt]['backs'].append(item['image'])

        # Ghép cặp (Front + Back)
        final_pairs = {}
        total_pairs = 0

        for dt, data in grouped_docs.items():
            pairs = []
            f_list, b_list = data['fronts'], data['backs']
            
            # Ưu tiên ghép front[i] với back[i]
            min_len = min(len(f_list), len(b_list))
            for i in range(min_len):
                pairs.append((f_list[i], b_list[i]))
            
            if pairs:
                final_pairs[dt] = pairs
                total_pairs += len(pairs)

        st.success(f"🎉 Đã tự động phân loại và ghép được **{total_pairs} bộ giấy tờ** hoàn chỉnh!")

        # Hiển thị kết quả đã phân loại
        for dt, pairs in final_pairs.items():
            st.markdown(f"### 📋 Loại: **{dt}** ({len(pairs)} bộ)")
            cols = st.columns(min(len(pairs), 3))
            for idx, (f, b) in enumerate(pairs):
                with cols[idx % 3]:
                    st.caption(f"{dt} - Bộ {idx+1}")
                    st.image(f, caption="Mặt trước", use_container_width=True)
                    st.image(b, caption="Mặt sau", use_container_width=True)

        # Download File Word tổng hợp
        if total_pairs > 0:
            docx_file = create_multi_docx(final_pairs)
            st.markdown("---")
            st.download_button(
                label=f"📝 Tải file WORD tổng hợp tất cả {total_pairs} bộ giấy tờ (.docx)",
                data=docx_file,
                file_name="GiayTo_DaPhanLoai_InA4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
