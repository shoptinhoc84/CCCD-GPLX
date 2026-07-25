import cv2
import numpy as np
from PIL import Image
import pytesseract
import re
import gc

def preprocess_for_ocr(pil_img):
    """Xử lý ảnh nhanh, nhẹ RAM và tối ưu cho Tesseract."""
    img_np = np.array(pil_img)
    
    # Chuyển ảnh xám
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    
    # Tăng độ tương phản nhẹ bằng CLAHE (Nhanh và KHÔNG tốn RAM như detailEnhance)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    
    # Khử nhiễu nhẹ
    blur = cv2.GaussianBlur(enhanced, (3, 3), 0)
    
    # Nhị phân hóa Otsu
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
    stop_words_name = [
        "CONG HOA", "XHCN", "VIET NAM", "DOC LAP", "TU DO", "HANH PHUC",
        "CAN CUOC", "CONG DAN", "SOCIALIST", "REPUBLIC", "IDENTITY", "CARD",
        "GIAY PHEP", "LAI XE", "DRIVER", "LICENSE", "HO TEN", "FULL NAME",
        "NGAY SINH", "DATE OF BIRTH", "QUOC TICH", "NATIONALITY", "NOI DANG KY",
        "EXPIRE", "CSGT", "CUC CANH SAT"
    ]

    for idx, file in enumerate(list_uploaded_files):
        img_raw = Image.open(file).convert("RGB")
        
        # Hạ kích thước tối đa xuống khoảng 1000px để giảm tải cho Tesseract
        img_opt = optimize_image_size(img_raw, max_dim=1000) 
        img_proc = preprocess_for_ocr(img_opt)
        
        # Chỉ quét 1 lần với ảnh đã qua tiền xử lý (vừa nhận diện tiếng Việt tốt, vừa lấy số chuẩn)
        full_text = pytesseract.image_to_string(img_proc, lang='eng+vie').upper()
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]

        is_gplx = ("GIẤY PHÉP LÁI XE" in full_text) or ("DRIVER" in full_text) or ("SỐ/NO" in full_text)
        is_cccd = ("CĂN CƯỚC" in full_text) or ("CAN CUOC" in full_text) or ("CITIZEN" in full_text) or ("SỐ / NO" in full_text)

        if is_gplx:
            logs.append(f"📸 Ảnh #{idx+1}: Nhận diện là **GPLX**")
            
            gplx_match = re.search(r'\b\d{12}\b', full_text)
            if gplx_match:
                data["so_gplx"] = gplx_match.group(0)

            hang_match = re.search(r'HẠNG[/\s]*CLASS[:\s]*([A-Z0-9]{1,3})', full_text)
            if hang_match:
                found_hang = hang_match.group(1).strip()
                if found_hang in ["A1", "A2", "A3", "A4", "A", "B1", "B2", "B", "C", "D", "E", "FC", "FE"]:
                    data["hang_gplx"] = found_hang
            else:
                for h in ["A1", "A2", "B2", "B1", "A", "C", "D"]:
                    if f"HẠNG {h}" in full_text or f"CLASS {h}" in full_text or f"HẠNG/CLASS {h}" in full_text:
                        data["hang_gplx"] = h
                        break

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

            for line in lines:
                clean_l = re.sub(r'[^A-Z\s]', '', line).strip()
                if clean_l.isupper() and len(clean_l) > 5 and len(clean_l.split()) >= 2:
                    if not any(stop_word in clean_l for stop_word in stop_words_name):
                        if not data["ho_ten"]:
                            data["ho_ten"] = clean_l
                            break

        # GIẢI PHÓNG BỘ NHỚ TRIỆT ĐỂ
        del img_raw, img_opt, img_proc
        gc.collect()

    return data, logs
