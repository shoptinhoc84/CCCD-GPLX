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
# CẤU HÌNH TRANG & GIAO DIỆN (UI/UX CHUYÊN NGHIỆP)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Tự Động Xử Lý & Dàn Trang A4",
    layout="wide",
    page_icon="🖨️",
    initial_sidebar_state="collapsed",
)

# Thêm CSS tùy chỉnh cho giao diện phẳng, đẹp mắt
st.markdown(
    """
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #4B5563;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stDownloadButton button {
        width: 100%;
        background-color: #2563EB !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 0.6rem 1rem !important;
    }
    .stDownloadButton button:hover {
        background-color: #1D4ED8 !important;
    }
    .card-box {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-header">🖨️ Công Cụ Xử Lý & Dàn Trang Giấy Tờ A4</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-header">Tự động tách viền • Xếp mặt trước/sau nằm ngang • 2 bộ/trang A4 chuẩn in ấn</div>',
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# THUẬT TOÁN XỬ LÝ ẢNH NHẸ & NHANH (CHỐNG OOM / SẬP WEB)
# ---------------------------------------------------------
def optimize_image_size(pil_img, max_dim=1600):
    """Giảm độ phân giải ảnh siêu nhanh nếu quá lớn để chống ngốn RAM."""
    w, h = pil_img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        new_w, new_h = int(w * scale), int(h * scale)
        return pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    return pil_img


def crop_card_fast(image_np):
    """Tách viền thẻ nhanh, tiết kiệm CPU/RAM."""
    h_img, w_img, _ = image_np.shape

    # Chuyển xám và làm mịn nhẹ
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # Threshold Otsu
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(
        thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE
    )

    best_rect = None
    max_area = 0

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
        return image_np[
            max(0, y - 4) : min(h_img, y + h + 8),
            max(0, x - 4) : min(w_img, x + w + 8),
        ]

    return image_np


# ---------------------------------------------------------
# XỬ LÝ FILE WORD & CANVAS A4 (NGANG - 2 BỘ / TRANG)
# ---------------------------------------------------------
def add_card_row_to_word(doc, pil_front, pil_back):
    """Thêm 1 bộ (Mặt trước + Mặt sau nằm ngang) vào Word dạng Bảng 1 dòng 2 cột."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    buf_front, buf_back = io.BytesIO(), io.BytesIO()
    pil_front.save(buf_front, format="PNG", optimize=True)
    pil_back.save(buf_back, format="PNG", optimize=True)
    buf_front.seek(0)
    buf_back.seek(0)

    # Cột 1: Mặt trước
    cell_f = table.cell(0, 0)
    p_f = cell_f.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f.add_run().add_picture(buf_front, width=Inches(3.35))

    # Cột 2: Mặt sau
    cell_b = table.cell(0, 1)
    p_b = cell_b.paragraphs[0]
    p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b.add_run().add_picture(buf_back, width=Inches(3.35))

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(20)


def create_multi_docx(card_pairs):
    """Tạo file Word (.docx) chứa 2 bộ / 1 trang A4 chuẩn."""
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
    """Tạo file ảnh A4 (PNG) chứa tối đa 2 bộ nằm ngang."""
    a4_w, a4_h = 1240, 1754
    canvas = Image.new("RGB", (a4_w, a4_h), "white")

    card_w = 540
    spacing_x = 40
    start_x = (a4_w - (card_w * 2 + spacing_x)) // 2
    y_positions = [180, 920]

    for idx, (pil_front, pil_back) in enumerate(card_pairs_chunk):
        if idx >= 2:
            break

        y_pos = y_positions[idx]

        ratio_f = card_w / pil_front.width
        f_resized = pil_front.resize(
            (card_w, int(pil_front.height * ratio_f)), Image.Resampling.LANCZOS
        )

        ratio_b = card_w / pil_back.width
        b_resized = pil_back.resize(
            (card_w, int(pil_back.height * ratio_b)), Image.Resampling.LANCZOS
        )

        canvas.paste(f_resized, (start_x, y_pos))
        canvas.paste(b_resized, (start_x + card_w + spacing_x, y_pos))

    return canvas


# ---------------------------------------------------------
# THIẾT KẾ GIAO DIỆN & LUỒNG XỬ LÝ (FLOW)
# ---------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Cấu hình hệ thống")
    doc_type = st.radio(
        "📋 Loại giấy tờ xử lý:",
        ["Căn cước công dân (CCCD)", "Giấy phép lái xe (GPLX)"],
    )
    st.info(
        "💡 **Quy tắc ghép cặp:**\nFiles sẽ được ghép lần lượt theo thứ tự: "
        "\n• Ảnh 1 + Ảnh 2 $\rightarrow$ Bộ 1"
        "\n• Ảnh 3 + Ảnh 4 $\rightarrow$ Bộ 2..."
    )

uploaded_files = st.file_uploader(
    "📥 Tải lên toàn bộ ảnh (Chọn nhiều ảnh cùng lúc):",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    help="Giữ phím Ctrl (hoặc Chọn nhiều ảnh từ điện thoại) để tải toàn bộ mặt trước/sau.",
)

if uploaded_files:
    num_files = len(uploaded_files)

    if num_files < 2:
        st.warning("⚠️ Vui lòng tải lên ít nhất 2 ảnh (mặt trước và mặt sau).")
    else:
        # Thống kê KPI nhanh
        num_pairs = num_files // 2
        col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
        col_kpi1.metric("📸 Tổng số ảnh đã chọn", f"{num_files} ảnh")
        col_kpi2.metric("📋 Số bộ hoàn chỉnh", f"{num_pairs} bộ")
        col_kpi3.metric(
            "📄 Số trang A4 dự kiến", f"{(num_pairs + 1) // 2} trang"
        )

        if num_files % 2 != 0:
            st.caption(
                f"ℹ️ Bạn đã tải lên số lẻ ({num_files} ảnh). Hệ thống sẽ ghép đủ {num_pairs * 2} ảnh đầu tiên."
            )

        card_pairs = []
        a4_canvases = []

        # Thanh tiến trình chuyên nghiệp
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, i in enumerate(range(0, num_pairs * 2, 2)):
            status_text.text(
                f"⏳ Đang xử lý Bộ {idx + 1}/{num_pairs}: Lọc viền & Tối ưu hóa kích thước..."
            )

            # Tải ảnh & Giảm kích thước trước để cứu RAM
            raw_f = Image.open(uploaded_files[i]).convert("RGB")
            raw_b = Image.open(uploaded_files[i + 1]).convert("RGB")

            opt_f = optimize_image_size(raw_f)
            opt_b = optimize_image_size(raw_b)

            # Tách viền
            crop_f = crop_card_fast(np.array(opt_f))
            crop_b = crop_card_fast(np.array(opt_b))

            pil_f = Image.fromarray(crop_f)
            pil_b = Image.fromarray(crop_b)

            card_pairs.append((pil_f, pil_b))

            # Giải phóng RAM ngay lập tức
            del raw_f, raw_b, opt_f, opt_b, crop_f, crop_b
            gc.collect()

            progress_bar.progress((idx + 1) / num_pairs)

        status_text.text("⚡ Đang tạo file Word và kết xuất ảnh A4...")

        # Tạo trang A4 (Mỗi trang chứa 2 bộ)
        for i in range(0, len(card_pairs), 2):
            chunk = card_pairs[i : i + 2]
            canvas = create_a4_canvas_horizontal(chunk)
            a4_canvases.append(canvas)

        status_text.empty()
        progress_bar.empty()
        st.success(
            f"✅ Đã xử lý hoàn tất **{len(card_pairs)}** bộ {doc_type.split()[0]}!"
        )

        # ---------------------------------------------------------
        # KHU VỰC XEM TRƯỚC & TẢI FILE (PREVIEW & DOWNLOAD)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader("📄 Xem trước bản thiết kế A4 (Bố cục Ngang)")

        # Hiển thị các trang A4 trong tab chuyên nghiệp
        if len(a4_canvases) > 1:
            tabs = st.tabs([f"Trang A4 #{i + 1}" for i in range(len(a4_canvases))])
            for i, tab in enumerate(tabs):
                with tab:
                    st.image(
                        a4_canvases[i],
                        caption=f"Mẫu in Trang {i + 1} (Gồm {len(card_pairs[i * 2:(i + 1) * 2])} bộ)",
                        use_container_width=True,
                    )
        else:
            st.image(
                a4_canvases[0],
                caption="Mẫu in Trang 1 (Gồm 2 bộ nằm ngang)",
                use_container_width=True,
            )

        # Xuất file Word
        docx_io = create_multi_docx(card_pairs)

        st.markdown("---")
        st.subheader("📥 Tải File Hoàn Thiện Để In")

        d_col1, d_col2 = st.columns(2)

        with d_col1:
            st.download_button(
                label=f"📝 Tải File WORD (.docx) — {len(card_pairs)} Bộ",
                data=docx_io,
                file_name=f"{doc_type.split()[0]}_BocucNgang_A4.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        with d_col2:
            img_io = io.BytesIO()
            a4_canvases[0].save(img_io, format="PNG", optimize=True)
            img_io.seek(0)

            st.download_button(
                label="🖼️ Tải Ảnh PNG Trang 1",
                data=img_io,
                file_name=f"{doc_type.split()[0]}_Trang1.png",
                mime="image/png",
            )
