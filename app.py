import io
import cv2
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
import numpy as np
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="Công cụ Photo / Ghép giấy tờ in A4",
    layout="centered",
    page_icon="🖨️",
)

st.title("🖨️ Lọc viền & Ghép giấy tờ in A4")
st.write(
    "Hỗ trợ tự động cắt bỏ phần thừa và xếp **Mặt trước - Mặt sau nằm ngang song song** sẵn sàng để in/photocopy."
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


def create_docx_horizontal(pil_front, pil_back):
    """Tạo file Word (.docx) chứa 2 mặt giấy tờ nằm ngang hàng chuẩn khổ A4."""
    doc = docx.Document()

    # Cấu hình lề trang A4 (Lề trên/dưới/trái/phải ~ 1.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Byte stream chuyển đổi ảnh
    buf_front = io.BytesIO()
    pil_front.save(buf_front, format="PNG")
    buf_front.seek(0)

    buf_back = io.BytesIO()
    pil_back.save(buf_back, format="PNG")
    buf_back.seek(0)

    # Sử dụng Bảng (Table) 1 hàng 2 cột không viền để đặt 2 mặt nằm ngang
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Kích thước chuẩn thực tế chiều rộng thẻ ~ 3.37 inches (8.56 cm)
    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)

    p_left = cell_left.paragraphs[0]
    p_left.add_run().add_picture(buf_front, width=Inches(3.37))

    p_right = cell_right.paragraphs[0]
    p_right.add_run().add_picture(buf_back, width=Inches(3.37))

    # Bỏ đường viền của bảng
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = docx.oxml.OxmlElement("w:tcBorders")
            for border_name in ["top", "left", "bottom", "right"]:
                border = docx.oxml.OxmlElement(f"w:{border_name}")
                border.set(docx.oxml.ns.qn("w:val"), "none")
                tcBorders.append(border)
            tcPr.append(tcBorders)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# Giao diện ứng dụng
doc_type = st.radio(
    "📋 Chọn loại giấy tờ cần xử lý:",
    ["Căn cước công dân (CCCD)", "Giấy phép lái xe (GPLX)"],
    horizontal=True,
)

upload_method = st.radio(
    "📸 Phương thức tải ảnh:",
    ["Tải file từ máy tính/điện thoại", "Chụp trực tiếp bằng Camera"],
    horizontal=True,
)

col1, col2 = st.columns(2)
front_file, back_file = None, None

if upload_method == "Tải file từ máy tính/điện thoại":
    with col1:
        front_file = st.file_uploader(
            f"Mặt trước {doc_type.split()[0]}",
            type=["jpg", "jpeg", "png"],
            key="front_file",
        )
    with col2:
        back_file = st.file_uploader(
            f"Mặt sau {doc_type.split()[0]}",
            type=["jpg", "jpeg", "png"],
            key="back_file",
        )
else:
    with col1:
        front_file = st.camera_input(
            f"Chụp mặt trước {doc_type.split()[0]}", key="front_cam"
        )
    with col2:
        back_file = st.camera_input(
            f"Chụp mặt sau {doc_type.split()[0]}", key="back_cam"
        )

if front_file and back_file:
    img_front_raw = Image.open(front_file).convert("RGB")
    img_back_raw = Image.open(back_file).convert("RGB")

    np_front = np.array(img_front_raw)
    np_back = np.array(img_back_raw)

    with st.spinner("Đang xử lý và tách nền tự động..."):
        cropped_front = crop_card(np_front)
        cropped_back = crop_card(np_back)

    pil_front = Image.fromarray(cropped_front)
    pil_back = Image.fromarray(cropped_back)

    st.markdown("---")
    st.subheader("📸 Kết quả sau khi lọc phần thừa")
    c1, c2 = st.columns(2)
    c1.image(pil_front, caption="Mặt trước (đã cắt)", use_container_width=True)
    c2.image(pil_back, caption="Mặt sau (đã cắt)", use_container_width=True)

    # 1. Tạo file Ảnh PNG A4 (Nằm ngang hàng)
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "white")

    # Kích thước mỗi mặt thẻ (Rộng ~ 530px)
    target_w = 530
    ratio_f = target_w / pil_front.width
    pil_front_resized = pil_front.resize(
        (target_w, int(pil_front.height * ratio_f))
    )

    ratio_b = target_w / pil_back.width
    pil_back_resized = pil_back.resize(
        (target_w, int(pil_back.height * ratio_b))
    )

    # Tọa độ ghép song song 2 mặt nằm ngang trên A4
    gap = 60  # Khoảng cách giữa 2 thẻ
    y_pos = 250  # Vị trí hàng ngang từ trên xuống
    x_front = (a4_w - (target_w * 2 + gap)) // 2
    x_back = x_front + target_w + gap

    canvas.paste(pil_front_resized, (x_front, y_pos))
    canvas.paste(pil_back_resized, (x_back, y_pos))

    img_io = io.BytesIO()
    canvas.save(img_io, format="PNG")
    img_io.seek(0)

    # 2. Tạo file Word (.docx) nằm ngang
    docx_io = create_docx_horizontal(pil_front, pil_back)

    st.markdown("---")
    st.subheader("📄 Xem trước & Tải file về in")
    st.image(
        canvas, width=420, caption="Bố cục photo ngang hàng sẵn sàng để in A4"
    )

    btn_col1, btn_col2 = st.columns(2)

    with btn_col1:
        st.download_button(
            label="📝 Tải file WORD (.docx)",
            data=docx_io,
            file_name=f"{doc_type.split()[0]}_Photo_Ngang.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with btn_col2:
        st.download_button(
            label="🖼️ Tải file ÁNH (.png)",
            data=img_io,
            file_name=f"{doc_type.split()[0]}_Photo_Ngang.png",
            mime="image/png",
            use_container_width=True,
        )
