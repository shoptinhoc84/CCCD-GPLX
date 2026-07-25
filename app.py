import streamlit as st
import cv2
import numpy as np
from PIL import Image
import pytesseract
import re
import gc

# ---------------------------------------------------------
# HÀM TỐI ƯU VÀ XỬ LÝ OCR CHÍNH XÁC
# ---------------------------------------------------------

def optimize_image_size(pil_img, max_dim=1200):
    width, height = pil_img.size
    if max(width, height) > max_dim:
        scale = max_dim / float(max(width, height))
        new_width = int(width * scale)
        new_height = int(height * scale)
        return pil_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
    return pil_img

def preprocess_for_ocr(pil_img):
    """Tiền xử lý ảnh làm rõ nét chữ cho Tesseract."""
    img_np = np.array(pil_img)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    
    # Nâng độ tương phản bằng CLAHE nhẹ
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    
    return Image.fromarray(enhanced)

def extract_dates(text):
    """Trích xuất ngày tháng linh hoạt (chấp nhận /, ., - hoặc khoảng trắng)."""
    # Regex tìm ngày tháng dạng DD/MM/YYYY, DD.MM.YYYY, DD-MM-YYYY
    pattern = r'\b(\d{2})[/.\-\s](\d{2})[/.\-\s](\d{4})\b'
    matches = re.findall(pattern, text)
    formatted_dates = [f"{m[0]}/{m[1]}/{m[2]}" for m m in matches]
    return formatted_dates

def process_auto_batch_ocr(list_uploaded_files):
    data = {
        "ho_ten": "", "ngay_sinh": "", "so_cccd": "", "ngay_cap_cccd": "",
        "noi_cap_cccd": "Cục Cảnh sát quản lý hành chính về trật tự xã hội",
        "so_gplx": "", "hang_gplx": "", "noi_cap_gplx": "", "ngay_cap_gplx": "",
        "hang_dang_ky": "A1", "vi_pham": "Không", "sdt": ""
    }

    logs = []

    # Danh sách từ khóa cấm xuất hiện trong Tên (chứa cả từ đơn nhiễu)
    forbidden_words = [
        "CONG", "HOA", "XA", "HOI", "CHU", "NGHIA", "VIET", "NAM",
        "DOC", "LAP", "TU", "DO", "HANH", "PHUC", "CAN", "CUOC",
        "CONG", "DAN", "SOCIALIST", "REPUBLIC", "IDENTITY", "CARD",
        "GIAY", "PHEP", "LAI", "XE", "DRIVER", "LICENSE", "FULL", "NAME",
        "DATE", "BIRTH", "NATIONALITY", "EXPIRE", "CSGT", "CUC", "CANH", "SAT",
        "SEX", "GIOI", "TINH", "QUE", "QUAN", "NOI", "TRU", "ORIGIN", "RESIDENCE"
    ]

    for idx, file in enumerate(list_uploaded_files):
        img_raw = Image.open(file).convert("RGB")
        img_opt = optimize_image_size(img_raw, max_dim=1200)
        img_proc = preprocess_for_ocr(img_opt)
        
        # Quét OCR với cả 2 ngôn ngữ
        raw_text = pytesseract.image_to_string(img_proc, lang='vie+eng')
        
        # Tách dòng
        lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
        full_text_upper = raw_text.upper()

        is_gplx = ("GIẤY PHÉP LÁI XE" in full_text_upper) or ("DRIVER" in full_text_upper) or ("SỐ/NO" in full_text_upper)
        is_cccd = ("CĂN CƯỚC" in full_text_upper) or ("CAN CUOC" in full_text_upper) or ("CITIZEN" in full_text_upper) or ("SỐ / NO" in full_text_upper)

        if is_gplx:
            logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **GPLX**")
            
            # Quét số GPLX (12 chữ số)
            gplx_match = re.search(r'\b\d{12}\b', raw_text)
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
            dates_g = extract_dates(raw_text)
            if dates_g:
                data["ngay_cap_gplx"] = dates_g[0]

        else:
            if is_cccd:
                logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **CCCD / Căn cước**")
            else:
                logs.append(f"📸 Ảnh #{idx+1}: Quét dữ liệu bổ sung...")

            # 1. Số CCCD (12 chữ số)
            id_match = re.search(r'\b\d{12}\b', raw_text)
            if id_match and not data["so_cccd"]:
                data["so_cccd"] = id_match.group(0)

            # 2. Quét Ngày sinh & Ngày cấp
            date_matches = extract_dates(raw_text)
            if date_matches:
                if not data["ngay_sinh"]:
                    data["ngay_sinh"] = date_matches[0]
                if len(date_matches) > 1 and not data["ngay_cap_cccd"]:
                    data["ngay_cap_cccd"] = date_matches[1]

            # 3. Thuật toán trích xuất Họ Tên chuẩn xác
            for i, line in enumerate(lines):
                line_upper = line.upper()
                
                # Cách 1: Tìm dòng đứng ngay sau chữ "HỌ VÀ TÊN" / "FULL NAME"
                if any(kw in line_upper for kw in ["HỌ VÀ TÊN", "HO VA TEN", "FULL NAME", "HỌ TÊN"]):
                    # Nếu tên nằm cùng dòng với nhãn
                    after_label = re.sub(r'.*(HỌ VÀ TÊN|HO VA TEN|FULL NAME|HỌ TÊN)[:\s]*', '', line_upper)
                    clean_after = re.sub(r'[^A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚƯÝĐ\s]', '', after_label).strip()
                    if len(clean_after) >= 5 and len(clean_after.split()) >= 2:
                        data["ho_ten"] = clean_after
                        break
                    
                    # Nếu tên nằm ở dòng tiếp theo ngay phía dưới
                    elif i + 1 < len(lines):
                        next_line = lines[i+1].upper()
                        clean_next = re.sub(r'[^A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚƯÝĐ\s]', '', next_line).strip()
                        # Kiểm tra xem dòng tiếp theo có bị dính từ cấm không
                        words = clean_next.split()
                        if len(clean_next) >= 5 and len(words) >= 2:
                            if not any(w in forbidden_words for w in words):
                                data["ho_ten"] = clean_next
                                break

            # Cách 2 (Dự phòng): Lọc từng dòng nếu Cách 1 không tìm thấy
            if not data["ho_ten"]:
                for line in lines:
                    line_upper = line.upper()
                    clean_line = re.sub(r'[^A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚƯÝĐ\s]', '', line_upper).strip()
                    words = clean_line.split()
                    
                    if len(clean_line) >= 5 and len(words) >= 2:
                        # Bỏ qua nếu có bất kỳ từ cấm nào trong dòng
                        if any(w in forbidden_words for w in words):
                            continue
                        data["ho_ten"] = clean_line
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
