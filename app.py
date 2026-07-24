import io
import cv2
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

# Tải EasyOCR (dùng cache để không bị tải lại model khi rerun app)
try:
    import easyocr

    @st.cache_resource
    def load_ocr():
        return easyocr.Reader(["vi", "en"], gpu=False)

    reader = load_ocr()
except Exception as e:
    reader = None

st.set_page_config(
    page_title="Hệ thống phân loại & Ghép giấy tờ A4 hàng loạt",
    layout="wide",
    page_icon="🖨️",
)

st.title("🖨️ Phân loại Smart & Ghép giấy tờ in A4")
st.write(
    "Tải lên **tất cả ảnh** (CCCD, GPLX, mặt trước, mặt sau lẫn lộn) vào 1 ô duy nhất. "
    "Hệ thống sẽ tự nhận diện loại giấy tờ, tự lọc viền và xếp theo đúng cặp!"
)

if reader is None:
    st.error(
        "⚠️ Chưa thể khởi tạo thư viện `easyocr`. "
        "Vui lòng đảm bảo đã thêm `easyocr` vào `requirements.txt` và `libgl1`, `libglib2.0-0` vào `packages.txt`."
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


def classify_and_detect_side(image_np):
    """Tự động nhận dạng:

    - Loại giấy tờ: 'CCCD' hoặc 'GPLX'
    - Mặt: 'front' hoặc 'back'
    """
    if reader is None:
        return "CCCD", "front"

    # Đọc text trong ảnh bằng EasyOCR
    results = reader.readtext(image_np, detail=0)
    text_combined = " ".join(results).lower()

    # Logic nhận diện Loại giấy tờ
    doc_type = "CCCD"
    gplx_keywords = [
        "bằng lái",
        "giấy phép lái xe",
        "driving licence",
        "gplx",
        "phép lái xe",
    ]
    if any(k in text_combined for k in gplx_keywords):
        doc_type = "GPLX"
    elif any(
        k in text_combined
        for k in ["căn cước", "mã số", "chứng minh", "identity card"]
    ):
        doc_type = "CCCD"

    # Logic nhận diện Mặt trước / Mặt sau
    side = "front"
    back_keywords = [
        "đặc điểm nhân dạng",
        "ngày, tháng, năm cấp",
        "có giá trị đến",
        "thâm quyến",
        "cục trưởng",
        "ký, ghi rõ",
    ]
    front_keywords = [
        "họ và tên",
        "ngày sinh",
        "giới tính",
        "quốc tịch",
        "date of birth",
    ]

    if any(k in text_combined for k in back_keywords):
        side = "back"
    elif any(k in text_combined for k in front_keywords):
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
            buf_f.seek(0)
            buf_b.seek(0)

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
        st.warning("⚠️ Vui lòng chọn ít nhất 2 ảnh để thực hiện ghép cặp.")
    else:
        with st.spinner(
            "🤖 Đang quét OCR, tự động phân loại giấy tờ và ghép mặt trước/sau..."
        ):
            processed_data = []

            for file in uploaded_files:
                img_raw = Image.open(file).convert("RGB")
                np_img = np.array(img_raw)

                # 1. Tách viền
                cropped_np = crop_card(np_img)
                cropped_pil = Image.fromarray(cropped_np)

                # 2. Nhận diện OCR
                doc_type, side = classify_and_detect_side(cropped_np)

                processed_data.append({
                    "file_name": file.name,
                    "image": cropped_pil,
                    "type": doc_type,
                    "side": side,
                })

        # Gom nhóm dữ liệu theo Loại giấy tờ
        grouped_docs = {}
        for item in processed_data:
            dt = item["type"]
            if dt not in grouped_docs:
                grouped_docs[dt] = {"fronts": [], "backs": []}

            if item["side"] == "front":
                grouped_docs[dt]["fronts"].append(item["image"])
            else:
                grouped_docs[dt]["backs"].append(item["image"])

        # Tự động ghép cặp (Mặt trước + Mặt sau)
        final_pairs = {}
        total_pairs = 0

        for dt, data in grouped_docs.items():
            pairs = []
            f_list, b_list = data["fronts"], data["backs"]

            # Lần lượt ghép front[i] với back[i]
            min_len = min(len(f_list), len(b_list))
            for i in range(min_len):
                pairs.append((f_list[i], b_list[i]))

            if pairs:
                final_pairs[dt] = pairs
                total_pairs += len(pairs)

        if total_pairs > 0:
            st.success(
                f"🎉 Đã tự động phân loại và ghép thành công **{total_pairs} bộ giấy tờ**!"
            )

            # Hiển thị kết quả xem trước
            for dt, pairs in final_pairs.items():
                st.markdown(f"### 📋 Loại: **{dt}** ({len(pairs)} bộ)")
                cols = st.columns(min(len(pairs), 3))
                for idx, (f, b) in enumerate(pairs):
                    with cols[idx % 3]:
                        st.caption(f"{dt} - Bộ {idx + 1}")
                        st.image(
                            f, caption="Mặt trước", use_container_width=True
                        )
                        st.image(
                            b, caption="Mặt sau", use_container_width=True
                        )

            # Tạo nút tải file Word
            docx_file = create_multi_docx(final_pairs)
            st.markdown("---")
            st.download_button(
                label=f"📝 Tải file WORD tổng hợp {total_pairs} bộ giấy tờ (.docx)",
                data=docx_file,
                file_name="GiayTo_DaPhanLoai_InA4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.warning(
                "⚠️ Chưa ghép được bộ hoàn chỉnh nào. Vui lòng đảm bảo bạn có đủ cả mặt trước và mặt sau."
            )
