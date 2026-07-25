import streamlit as st
import cv2
import numpy as np
from PIL import Image
import pytesseract
import re
import gc

# ---------------------------------------------------------
# HÀM XỬ LÝ VÀ TỐI ƯU ẢNH
# ---------------------------------------------------------

def optimize_image_size(pil_img, max_dim=1200):
    """Giảm kích thước ảnh nhẹ nhàng để tránh đơ app nhưng giữ đủ nét."""
    width, height = pil_img.size
    if max(width, height) > max_dim:
        scale = max_dim / float(max(width, height))
        new_width = int(width * scale)
        new_height = int(height * scale)
        return pil_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
    return pil_img

def preprocess_for_digits(pil_img):
    """Ảnh tăng tương phản chuyên biệt cho việc quét SỐ (CCCD / GPLX / Ngày tháng)."""
    img_np = np.array(pil_img)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    blur = cv2.GaussianBlur(enhanced, (3, 3), 0)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(thresh)

def process_auto_batch_ocr(list_uploaded_files):
    data = {
        "ho_ten": "", "ngay_sinh": "", "so_cccd": "", "ngay_cap_cccd": "",
        "noi_cap_cccd": "Cục Cảnh sát quản lý hành chính về trật tự xã hội",
        "so_gplx": "", "hang_gplx": "", "noi_cap_gplx": "", "ngay_cap_gplx": "",
        "hang_dang_ky": "A1", "vi_pham": "Không", "sdt": ""
    }

    logs = []

    # Danh sách từ khóa nhiễu được bổ sung triệt để
    stop_words_name = [
        "CONG HOA", "XHCN", "VIET NAM", "DOC LAP", "TU DO", "HANH PHUC",
        "CAN CUOC", "CONG DAN", "SOCIALIST", "REPUBLIC", "IDENTITY", "CARD",
        "GIAY PHEP", "LAI XE", "DRIVER", "LICENSE", "HO TEN", "FULL NAME",
        "NGAY SINH", "DATE OF BIRTH", "QUOC TICH", "NATIONALITY", "NOI DANG KY",
        "EXPIRE", "CSGT", "CUC CANH SAT", "SEX", "GIOI TINH", "QUE QUAN",
        "NOI TRU", "PLACE OF ORIGIN", "PLACE OF RESIDENCE", "PERSONAL IDENTIFICATION"
    ]

    for idx, file in enumerate(list_uploaded_files):
        img_raw = Image.open(file).convert("RGB")
        img_opt = optimize_image_size(img_raw, max_dim=1200)
        
        # Quét 1: Dùng ảnh GỐC để trích xuất CHỮ TIẾNG VIỆT (Tên, Ngày tháng)
        text_raw = pytesseract.image_to_string(img_opt, lang='vie+eng')
        
        # Quét 2: Dùng ảnh NHỊ PHÂN để trích xuất SỐ chuẩn xác
        img_proc = preprocess_for_digits(img_opt)
        text_digits = pytesseract.image_to_string(img_proc, lang='eng')

        full_text = text_raw + "\n" + text_digits
        full_text_upper = full_text.upper()
        lines = [line.strip() for line in text_raw.split('\n') if line.strip()]

        is_gplx = ("GIẤY PHÉP LÁI XE" in full_text_upper) or ("DRIVER" in full_text_upper) or ("SỐ/NO" in full_text_upper)
        is_cccd = ("CĂN CƯỚC" in full_text_upper) or ("CAN CUOC" in full_text_upper) or ("CITIZEN" in full_text_upper) or ("SỐ / NO" in full_text_upper)

        if is_gplx:
            logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **GPLX**")
            
            # Quét số GPLX (12 chữ số)
            gplx_match = re.search(r'\b\d{12}\b', full_text)
            if gplx_match:
                data["so_gplx"] = gplx_match.group(0)

            # Quét hạng GPLX
            hang_match = re.search(r'HẠNG[/\s]*CLASS[:\s]*([A-Z0-9]{1,3})', full_text_upper)
            if hang_match:
                found_hang = hang_match.group(1).strip()
                if found_hang in ["A1", "A2", "A3", "A4", "A", "B1", "B2", "B", "C", "D", "E", "FC", "FE"]:
                    data["hang_gplx"] = found_hang
            else:
                for h in ["A1", "A2", "B2", "B1", "A", "C", "D"]:
                    if f"HẠNG {h}" in full_text_upper or f"CLASS {h}" in full_text_upper:
                        data["hang_gplx"] = h
                        break

            # Quét ngày cấp GPLX
            dates_g = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
            if dates_g:
                data["ngay_cap_gplx"] = dates_g[0]

        else:
            if is_cccd:
                logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **CCCD / Căn cước**")
            else:
                logs.append(f"📸 Ảnh #{idx+1}: Quét dữ liệu bổ sung...")

            # 1. Số CCCD (12 chữ số)
            id_match = re.search(r'\b\d{12}\b', full_text)
            if id_match and not data["so_cccd"]:
                data["so_cccd"] = id_match.group(0)

            # 2. Ngày sinh & Ngày cấp
            date_matches = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
            if date_matches:
                if not data["ngay_sinh"]:
                    data["ngay_sinh"] = date_matches[0]
                if len(date_matches) > 1 and not data["ngay_cap_cccd"]:
                    data["ngay_cap_cccd"] = date_matches[1]

            # 3. Lọc lấy Họ Tên chuẩn xác
            for line in lines:
                line_clean = line.strip()
                line_upper = line_clean.upper()
                
                # Kiểm tra từ khóa nhiễu
                if any(stop_word in line_upper for stop_word in stop_words_name):
                    continue
                
                # Bắt các dòng viết hoa toàn bộ (bao gồm cả chữ cái Tiếng Việt có dấu)
                # Loại bỏ các ký tự đặc biệt hoặc số
                clean_name = re.sub(r'[^A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚƯÝĐ\s]', '', line_upper).strip()
                
                if len(clean_name) >= 5 and len(clean_name.split()) >= 2:
                    if not data["ho_ten"]:
                        data["ho_ten"] = clean_name
                        break

        # Giải phóng bộ nhớ
        del img_raw, img_opt, img_proc
        gc.collect()

    return data, logs


# ---------------------------------------------------------
# GIAO DIỆN STREAMLIT
# ---------------------------------------------------------

st.set_page_config(page_title="Nhận diện CCCD & GPLX", layout="wide")

st.title("🪪 Ứng Dụng Nhận Diện CCCD & GPLX Tự Động")
st.write("Tải lên ảnh CCCD hoặc Giấy phép lái xe để tự động trích xuất thông tin.")

uploaded_files = st.file_uploader(
    "Tải lên 1 hoặc nhiều ảnh (CCCD, GPLX):", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 Bắt đầu Quét & Nhận Diện"):
        with st.spinner("Đang xử lý hình ảnh và nhận diện dữ liệu..."):
            extracted_data, logs = process_auto_batch_ocr(uploaded_files)

        st.success("Xử lý hoàn tất!")

        # Hiển thị nhật ký quét
        st.subheader("📋 Nhật ký xử lý:")
        for log in logs:
            st.write(log)

        # Hiển thị dữ liệu thu được dưới dạng Form
        st.subheader("📝 Kết Quả Trích Xuất:")
        col1, col2 = st.columns(2)

        with col1:
            st.text_input("Họ và Tên", value=extracted_data["ho_ten"])
            st.text_input("Ngày sinh", value=extracted_data["ngay_sinh"])
            st.text_input("Số CCCD", value=extracted_data["so_cccd"])
            st.text_input("Ngày cấp CCCD", value=extracted_data["ngay_cap_cccd"])
            st.text_input("Nơi cấp CCCD", value=extracted_data["noi_cap_cccd"])

        with col2:
            st.text_input("Số GPLX", value=extracted_data["so_gplx"])
            st.text_input("Hạng GPLX", value=extracted_data["hang_gplx"])
            st.text_input("Ngày cấp GPLX", value=extracted_data["ngay_cap_gplx"])
            st.text_input("Hạng đăng ký", value=extracted_data["hang_dang_ky"])
            st.text_input("Nơi cấp GPLX", value=extracted_data["noi_cap_gplx"])
