"""Dàn trang CCCD/GPLX chuẩn A4 và cắt thẻ tự động / thủ công.

Chạy: streamlit run app_cccd_nang_cap.py
Phụ thuộc: streamlit, opencv-python-headless, Pillow, numpy, python-docx, reportlab
"""
from __future__ import annotations

import hashlib
import io
import math
from typing import Optional

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageOps
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader

# Dùng đơn vị milimét để in đúng kích thước thẻ ID-1 / CCCD / GPLX.
A4_MM = (210.0, 297.0)
CARD_MM = (85.6, 54.0)
DPI = 300
CARD_ASPECT = CARD_MM[0] / CARD_MM[1]


def read_image(data: bytes) -> Image.Image:
    with Image.open(io.BytesIO(data)) as im:
        im = ImageOps.exif_transpose(im)
        return im.convert("RGB")


def rotate_image(pil_img: Image.Image, degrees: int) -> Image.Image:
    if degrees % 360 == 0:
        return pil_img
    return pil_img.rotate(degrees, expand=True, resample=Image.Resampling.BICUBIC)


def _resize_for_detection(img: np.ndarray, max_side: int = 1650):
    h, w = img.shape[:2]
    scale = min(1.0, max_side / max(h, w))
    if scale < 1:
        img = cv2.resize(img, (round(w * scale), round(h * scale)), interpolation=cv2.INTER_AREA)
    return img, scale


def _order_quad(points) -> np.ndarray:
    """Trả các đỉnh theo thứ tự: trái-trên, phải-trên, phải-dưới, trái-dưới."""
    pts = np.asarray(points, np.float32).reshape(4, 2)
    sums = pts.sum(axis=1)
    diffs = pts[:, 0] - pts[:, 1]
    ids = [int(np.argmin(sums)), int(np.argmax(diffs)), int(np.argmax(sums)), int(np.argmin(diffs))]
    if len(set(ids)) == 4:
        return pts[ids]
    # Chống trường hợp góc xoay 45° khiến cách cộng/trừ trùng đỉnh.
    center = pts.mean(axis=0)
    theta = np.arctan2(pts[:, 1] - center[1], pts[:, 0] - center[0])
    ordered = pts[np.argsort(theta)]
    start = int(np.argmin(ordered.sum(axis=1)))
    return np.roll(ordered, -start, axis=0)


def _quad_geometry(quad: np.ndarray):
    tl, tr, br, bl = quad
    top = np.linalg.norm(tr - tl)
    bottom = np.linalg.norm(br - bl)
    left = np.linalg.norm(bl - tl)
    right = np.linalg.norm(br - tr)
    a = max(float(top), float(bottom))
    b = max(float(left), float(right))
    ratio = max(a, b) / max(1e-6, min(a, b))
    return a, b, ratio


def _candidate_edges(rgb: np.ndarray):
    """Nhiều bộ lọc hỗ trợ nền sáng, nền tối, nền hoa văn và thẻ bóng."""
    lab = cv2.cvtColor(rgb, cv2.COLOR_RGB2LAB)
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    smooth = cv2.bilateralFilter(enhanced, 7, 55, 55)
    kernels = [cv2.getStructuringElement(cv2.MORPH_RECT, (k, k)) for k in (3, 7)]
    masks = []
    for layer in (smooth, gray, clahe.apply(lab[:, :, 1]), clahe.apply(lab[:, :, 2])):
        for t1, t2 in ((30, 90), (65, 160)):
            edge = cv2.Canny(layer, t1, t2)
            for kernel in kernels:
                masks.append(cv2.morphologyEx(edge, cv2.MORPH_CLOSE, kernel, iterations=2))
    # Otsu hoạt động tốt với thẻ trắng trên bàn sẫm hoặc ngược lại.
    for layer in (smooth, gray):
        _, binary = cv2.threshold(layer, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        for b in (binary, cv2.bitwise_not(binary)):
            masks.append(cv2.morphologyEx(b, cv2.MORPH_CLOSE, kernels[1]))
    # Đường viền theo sự thay đổi độ sáng: hữu ích với ảnh chụp trên nền bàn.
    grad = cv2.morphologyEx(smooth, cv2.MORPH_GRADIENT, kernels[0])
    _, grad_bin = cv2.threshold(grad, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    masks.append(cv2.morphologyEx(grad_bin, cv2.MORPH_CLOSE, kernels[1]))
    return masks


def detect_card_candidates(pil_img: Image.Image, max_candidates: int = 10):
    """Trả [(quad theo pixel ảnh gốc, confidence 0..1), ...].

    Confidence là chỉ số gợi ý dựa trên hình học, KHÔNG phải xác suất chắc chắn.
    """
    orig = np.asarray(pil_img.convert("RGB"))
    scaled, scale = _resize_for_detection(orig)
    h, w = scaled.shape[:2]
    image_area = float(h * w)
    found = []
    for mask in _candidate_edges(scaled):
        contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        # Bỏ qua những đường viền quá nhỏ để tránh tốn tài nguyên.
        contours = sorted(contours, key=cv2.contourArea, reverse=True)[:160]
        for cnt in contours:
            contour_area = float(cv2.contourArea(cnt))
            portion = contour_area / image_area
            if not 0.035 < portion < 0.95:
                continue
            perimeter = cv2.arcLength(cnt, True)
            if perimeter < 1:
                continue
            hull = cv2.convexHull(cnt)
            if cv2.contourArea(hull) < 1:
                continue
            for epsilon in (0.015, 0.025, 0.04):
                approx = cv2.approxPolyDP(hull, epsilon * cv2.arcLength(hull, True), True)
                if len(approx) == 4 and cv2.isContourConvex(approx):
                    quad = _order_quad(approx.reshape(4, 2))
                    break
            else:
                # Bù khi viền bị rạn do lóa; chỉ chấp nhận khung chữ nhật xoay
                # nếu đường biên của vật thực sự bám tương đối sát hình chữ nhật.
                rect = cv2.minAreaRect(hull)
                rarea = float(rect[1][0] * rect[1][1])
                if rarea <= 0 or cv2.contourArea(hull) / rarea < 0.82:
                    continue
                quad = _order_quad(cv2.boxPoints(rect))
            area = float(abs(cv2.contourArea(quad)))
            frac = area / image_area
            if not 0.035 < frac < 0.93:
                continue
            a, b, ratio = _quad_geometry(quad)
            if not 1.16 <= ratio <= 2.30:
                continue
            min_border = float(np.min(np.stack((quad[:, 0], quad[:, 1], w - quad[:, 0], h - quad[:, 1]))))
            if min_border < min(w, h) * 0.007:
                continue
            hull_area = float(cv2.contourArea(hull))
            coverage = min(1.0, hull_area / max(area, 1.0))
            if coverage < 0.67:
                continue
            # Ưu tiên thẻ ID tỉ lệ ~1.586, có 4 cạnh đều và chiếm diện tích đủ lớn.
            aspect_score = math.exp(-2.8 * abs(math.log(ratio / CARD_ASPECT)))
            geom_score = min(1.0, coverage)
            center = quad.mean(axis=0)
            dist = math.hypot((center[0] - w / 2) / w, (center[1] - h / 2) / h)
            center_score = max(0.0, 1.0 - dist)
            # Không quá thưởng cho viền lớn (tránh cắt luôn tấm nền / mặt bàn).
            size_score = min(1.0, frac / 0.30)
            quality = (0.49 * aspect_score + 0.18 * geom_score + 0.22 * size_score + 0.11 * center_score)
            if quality < 0.52:
                continue
            quad_orig = quad / scale
            found.append((quad_orig.astype(np.float32), float(quality), float(frac)))
    # Loại các viền lồng nhau của cùng một thẻ.
    found.sort(key=lambda item: (item[1], item[2]), reverse=True)
    dedup = []
    for candidate in found:
        quad = candidate[0]
        box = cv2.boundingRect(quad)
        a_box = [box[0], box[1], box[0] + box[2], box[1] + box[3]]
        matched = False
        for prev in dedup:
            p = cv2.boundingRect(prev[0])
            b_box = [p[0], p[1], p[0] + p[2], p[1] + p[3]]
            inter_w = max(0, min(a_box[2], b_box[2]) - max(a_box[0], b_box[0]))
            inter_h = max(0, min(a_box[3], b_box[3]) - max(a_box[1], b_box[1]))
            area_inter = inter_w * inter_h
            area_union = (a_box[2] - a_box[0]) * (a_box[3] - a_box[1]) + (b_box[2] - b_box[0]) * (b_box[3] - b_box[1]) - area_inter
            if area_inter / max(1, area_union) > 0.52:
                matched = True
                break
        if not matched:
            dedup.append(candidate)
            if len(dedup) >= max_candidates:
                break
    return [(quad, confidence) for quad, confidence, _ in dedup]


def crop_quad(pil_img: Image.Image, points, output_width: int = 1280) -> Image.Image:
    pts = _order_quad(points)
    w, h, _ = _quad_geometry(pts)
    landscape = w >= h
    # Ảnh thẻ xoay dọc: nắn thành khung dọc trước, sau đó quay ngang.
    out_w, out_h = ((output_width, round(output_width / CARD_ASPECT)) if landscape
                    else (round(output_width / CARD_ASPECT), output_width))
    dst = np.array([[0, 0], [out_w - 1, 0], [out_w - 1, out_h - 1], [0, out_h - 1]], np.float32)
    mat = cv2.getPerspectiveTransform(pts, dst)
    warped = cv2.warpPerspective(np.asarray(pil_img.convert("RGB")), mat, (out_w, out_h),
                                 flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    if not landscape:
        warped = cv2.rotate(warped, cv2.ROTATE_90_CLOCKWISE)
    return Image.fromarray(warped)


def auto_crop(pil_img: Image.Image):
    matches = detect_card_candidates(pil_img)
    if not matches:
        return pil_img.copy(), None, 0.0
    quad, quality = matches[0]
    return crop_quad(pil_img, quad), quad, quality


def crop_rectangle(pil_img: Image.Image, left_pct, top_pct, right_pct, bottom_pct) -> Image.Image:
    w, h = pil_img.size
    l, r = sorted((float(left_pct), float(right_pct)))
    t, b = sorted((float(top_pct), float(bottom_pct)))
    x1, x2 = max(0, int(w * l / 100)), min(w, int(w * r / 100))
    y1, y2 = max(0, int(h * t / 100)), min(h, int(h * b / 100))
    if x2 <= x1 + 3 or y2 <= y1 + 3:
        raise ValueError("Vùng cắt quá nhỏ; tăng khoảng cách giữa hai biên.")
    return pil_img.crop((x1, y1, x2, y2))


def draw_detected_outline(pil_img: Image.Image, quad=None, percentage_corners=None) -> Image.Image:
    preview = pil_img.copy()
    preview.thumbnail((700, 700), Image.Resampling.LANCZOS)
    sx, sy = preview.width / pil_img.width, preview.height / pil_img.height
    if percentage_corners is not None:
        quad = np.array([[x * pil_img.width / 100, y * pil_img.height / 100]
                         for x, y in percentage_corners], np.float32)
    if quad is not None:
        poly = [(int(x * sx), int(y * sy)) for x, y in _order_quad(quad)]
        painter = ImageDraw.Draw(preview)
        painter.line(poly + [poly[0]], fill=(0, 230, 95), width=max(3, preview.width // 180))
        for idx, (x, y) in enumerate(poly):
            r = max(6, preview.width // 75)
            painter.ellipse((x - r, y - r, x + r, y + r), fill=(255, 180, 30), outline=(0, 0, 0), width=2)
            painter.text((x + r + 3, y - r - 3), str(idx + 1), fill=(0, 0, 0), stroke_width=2, stroke_fill=(255,255,255))
    return preview


def fit_into_card(pil_img: Image.Image, size_px: tuple[int, int]) -> Image.Image:
    """Không ép méo ảnh; dùng viền trắng nếu crop thủ công chưa đúng tỷ lệ."""
    card = Image.new("RGB", size_px, "white")
    shrunk = ImageOps.contain(pil_img.convert("RGB"), size_px, Image.Resampling.LANCZOS)
    card.paste(shrunk, ((size_px[0] - shrunk.width) // 2, (size_px[1] - shrunk.height) // 2))
    return card


def create_a4_canvas(card_pairs, pairs_per_page: int = 3, dpi: int = DPI) -> Image.Image:
    """Sắp xếp thẻ theo kích thước THỰC 85.6 x 54 mm trên A4 210 x 297 mm."""
    if not 1 <= len(card_pairs) <= pairs_per_page <= 4:
        raise ValueError("Số bộ trên trang phải từ 1 đến 4.")
    unit = dpi / 25.4
    px = lambda mm: round(mm * unit)
    page = Image.new("RGB", (px(A4_MM[0]), px(A4_MM[1])), "white")
    card_size = (px(CARD_MM[0]), px(CARD_MM[1]))
    gap_x_mm = 12.0
    x1 = px((A4_MM[0] - CARD_MM[0] * 2 - gap_x_mm) / 2)
    x2 = px((A4_MM[0] + gap_x_mm) / 2)
    gap_y_mm = {1: 0, 2: 30.0, 3: 19.0, 4: 11.5}[pairs_per_page]
    total_h_mm = pairs_per_page * CARD_MM[1] + (pairs_per_page - 1) * gap_y_mm
    y_start = (A4_MM[1] - total_h_mm) / 2
    for row_idx, (front, back) in enumerate(card_pairs):
        y = px(y_start + row_idx * (CARD_MM[1] + gap_y_mm))
        page.paste(fit_into_card(front, card_size), (x1, y))
        page.paste(fit_into_card(back, card_size), (x2, y))
    return page


def make_pages(card_pairs, pairs_per_page: int = 3, dpi: int = DPI):
    return [create_a4_canvas(card_pairs[i: i + pairs_per_page], pairs_per_page, dpi)
            for i in range(0, len(card_pairs), pairs_per_page)]


def _apply_cell_margins_zero(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = tc_pr.find(qn('w:tcMar'))
    if mar is None:
        mar = OxmlElement('w:tcMar')
        tc_pr.append(mar)
    for edge in ('top', 'left', 'bottom', 'right'):
        node = mar.find(qn('w:' + edge))
        if node is None:
            node = OxmlElement('w:' + edge)
            mar.append(node)
        node.set(qn('w:w'), '0')
        node.set(qn('w:type'), 'dxa')


def _set_cell_border_none(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = borders.find(qn('w:' + side))
        if el is None:
            el = OxmlElement('w:' + side)
            borders.append(el)
        el.set(qn('w:val'), 'nil')


def create_multi_docx(card_pairs, pairs_per_page: int = 3) -> io.BytesIO:
    """Bảng cố định 3 (hoặc 4) hàng thẻ, không phát sinh tràn trang ngoài ý muốn."""
    if not card_pairs:
        raise ValueError("Chưa có ảnh để dàn trang.")
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin, sec.right_margin = Mm(12), Mm(12)
    gap_y_mm = {2: 30.0, 3: 19.0, 4: 11.5}[pairs_per_page]
    total_h_mm = pairs_per_page * CARD_MM[1] + (pairs_per_page - 1) * gap_y_mm
    sec.top_margin, sec.bottom_margin = Mm((A4_MM[1] - total_h_mm) / 2), Mm(12)
    doc.styles['Normal'].paragraph_format.space_after = Pt(0)
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)

    # Một bảng DUY NHẤT trên mỗi trang tránh khoảng trống phát sinh giữa nhiều bảng.
    for page_index, start in enumerate(range(0, len(card_pairs), pairs_per_page)):
        if page_index:
            pb = doc.add_paragraph()
            pb.paragraph_format.space_after = Pt(0)
            pb.add_run().add_break(WD_BREAK.PAGE)
        subset = card_pairs[start:start + pairs_per_page]
        table = doc.add_table(rows=2 * len(subset) - 1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.columns[0].width = Mm(92)
        table.columns[1].width = Mm(92)
        for row_i, row in enumerate(table.rows):
            row.height = Mm(CARD_MM[1] if row_i % 2 == 0 else gap_y_mm)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            tr_pr = row._tr.get_or_add_trPr()
            no_split = OxmlElement('w:cantSplit')
            tr_pr.append(no_split)
            for col_idx, cell in enumerate(row.cells):
                cell.width = Mm(92)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                _apply_cell_margins_zero(cell)
                _set_cell_border_none(cell)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if row_i % 2 == 0:
                    img = subset[row_i // 2][col_idx]
                    target_px = (round(85.6 * DPI / 25.4), round(54 * DPI / 25.4))
                    composed = fit_into_card(img, target_px)
                    buf = io.BytesIO()
                    composed.save(buf, format='PNG')
                    buf.seek(0)
                    p.add_run().add_picture(buf, width=Mm(85.6), height=Mm(54))
        # Word có thể thêm một dòng rỗng sau bảng cuối, nhưng tổng chiều cao bảng
        # tối đa 4 x 54 + 3 x 11.5 = 250.5 mm, nhỏ hơn vùng in 267 mm.
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def create_pdf(pages: list[Image.Image]) -> bytes:
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4, pageCompression=1)
    for page in pages:
        jpg = io.BytesIO()
        page.convert('RGB').save(jpg, format='JPEG', quality=95, subsampling=0)
        jpg.seek(0)
        c.drawImage(ImageReader(jpg), 0, 0, width=A4[0], height=A4[1], mask='auto')
        c.showPage()
    c.save()
    return buf.getvalue()


def export_png(pil_img: Image.Image) -> bytes:
    buf = io.BytesIO()
    pil_img.save(buf, format='PNG', optimize=True)
    return buf.getvalue()


def export_pages_zip(pages: list[Image.Image], prefix: str) -> bytes:
    from zipfile import ZipFile, ZIP_DEFLATED
    out = io.BytesIO()
    with ZipFile(out, 'w', compression=ZIP_DEFLATED) as archive:
        for idx, page in enumerate(pages, start=1):
            archive.writestr(f'{prefix}_Trang_{idx}.png', export_png(page))
    return out.getvalue()


def crop_vneid_combined_block(pil_img: Image.Image):
    """Chỉ tự cắt nếu phát hiện rõ hai thẻ; không dùng tọa độ đoán theo % chiều cao."""
    matches = detect_card_candidates(pil_img, max_candidates=12)
    if len(matches) < 2:
        return pil_img.copy(), False, []
    # Lấy 2 thẻ khác vị trí, diện tích đáng kể và kích cỡ tương đối tương đồng.
    w, h = pil_img.size
    for i, (q1, score1) in enumerate(matches):
        box1 = cv2.boundingRect(q1)
        area1 = cv2.contourArea(q1)
        for q2, score2 in matches[i + 1:]:
            box2 = cv2.boundingRect(q2)
            area2 = cv2.contourArea(q2)
            if min(area1, area2) / max(area1, area2) < 0.35:
                continue
            xdist = abs((box1[0] + box1[2] / 2) - (box2[0] + box2[2] / 2))
            if xdist > w * 0.3:
                continue
            x_min = max(0, math.floor(min(q1[:, 0].min(), q2[:, 0].min()) - w * 0.012))
            x_max = min(w, math.ceil(max(q1[:, 0].max(), q2[:, 0].max()) + w * 0.012))
            y_min = max(0, math.floor(min(q1[:, 1].min(), q2[:, 1].min()) - h * 0.01))
            y_max = min(h, math.ceil(max(q1[:, 1].max(), q2[:, 1].max()) + h * 0.01))
            return pil_img.crop((x_min, y_min, x_max, y_max)), True, [q1, q2]
    return pil_img.copy(), False, []


def create_single_cropped_docx(pil_img: Image.Image) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(15), Mm(15)
    sec.left_margin, sec.right_margin = Mm(12), Mm(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    target_w = min(180.0, pil_img.width / max(1, pil_img.height) * 220)
    p.add_run().add_picture(io.BytesIO(export_png(pil_img)), width=Mm(target_w))
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def main():
    import streamlit as st

    st.set_page_config(page_title='CCCD & GPLX | Crop & In A4', page_icon='🪪', layout='wide')
    st.markdown('''
        <style>
        .block-container {padding-top:1.4rem; max-width:1480px;}
        .stDownloadButton button[kind="primary"] {font-weight:700;}
        </style>
        ''', unsafe_allow_html=True)
    st.title('🪪 Xử lý CCCD / GPLX & in A4')
    st.caption('Crop 4 góc + nắn phối cảnh • 3 bộ/trang A4 mặc định • PDF/Word/PNG • Xử lý trong ứng dụng, không dùng API OCR')

    if 'upload_cycle' not in st.session_state:
        st.session_state.upload_cycle = 0
    if 'swaps' not in st.session_state:
        st.session_state.swaps = {}

    @st.cache_data(show_spinner=False, max_entries=128)
    def cached_auto(image_bytes: bytes, angle: int):
        original = rotate_image(read_image(image_bytes), angle)
        cropped, quad, confidence = auto_crop(original)
        return cropped, quad, confidence

    tab_layout, tab_vneid = st.tabs(['🖨️ Dàn trang CCCD / GPLX', '📱 Crop 2 mặt từ VNeID'])
    with tab_layout:
        c1, c2 = st.columns([1, 1], gap='large')
        with c1:
            doc_type = st.radio('Loại giấy tờ', ['CCCD', 'GPLX'], horizontal=True)
            pairs_per_page = st.select_slider('Số bộ trên một trang A4', options=[2, 3, 4], value=3,
                                              help='Một bộ = 2 mặt. Mặc định 3 bộ (6 mặt) trên một trang.')
            if st.button('🧹 Xóa danh sách / Làm mới'):
                st.session_state.upload_cycle += 1
                st.session_state.swaps = {}
                st.rerun()
            files = st.file_uploader('Tải ảnh theo thứ tự: trước 1, sau 1, trước 2, sau 2, ...',
                                     type=['jpg','jpeg','png','webp'], accept_multiple_files=True,
                                     key=f'files_{st.session_state.upload_cycle}')
            st.caption('Mỗi 2 ảnh kế tiếp = 1 bộ. Nếu thiếu mặt sau, hệ thống để trống mặt đó.')
        with c2:
            st.info('**Chuẩn in:** khổ A4 đứng (210 × 297 mm). Mỗi CCCD/GPLX được đặt trong khung 85,6 × 54 mm, không kéo méo tỷ lệ. PDF phù hợp để nộp/in; Word có thể chỉnh từng ảnh.')

        if files:
            st.subheader('1. Kiểm tra ảnh đã cắt')
            st.caption('Các ảnh cắt không chắc chắn sẽ mở phần chỉnh sửa. Chọn "Chỉnh 4 góc" nếu nền phức tạp hoặc thẻ bị nghiêng.')
            processed = []
            for idx, uploaded in enumerate(files):
                data = uploaded.getvalue()
                file_key = f'{st.session_state.upload_cycle}_{idx}_{hashlib.md5(data).hexdigest()[:10]}'
                angle = st.session_state.get('rotation_' + file_key, 0)
                raw = rotate_image(read_image(data), angle)
                cropped_auto, quad, confidence = cached_auto(data, angle)
                verified = quad is not None and confidence >= 0.73
                status = '✅ Đã tìm khung' if verified else ('⚠️ Cần kiểm tra khung' if quad is not None else '⚠️ Không tìm thấy viền thẻ')
                with st.expander(f'Ảnh {idx + 1} · {uploaded.name} · {status}', expanded=not verified):
                    col_settings, col_preview = st.columns([1, 1.2], gap='medium')
                    with col_settings:
                        angle_new = st.selectbox('Xoay ảnh', [0,90,180,270], format_func=lambda v:f'{v}°',
                                                 key='rotation_' + file_key)
                        if angle_new != angle:
                            st.rerun()
                        choice = st.selectbox('Chế độ cắt',
                            ['Tự động (4 góc)', 'Cắt khung chữ nhật (%)', 'Chỉnh 4 góc (%)', 'Giữ nguyên ảnh'],
                            key='mode_' + file_key)
                        result = cropped_auto
                        drawn_points = quad
                        if choice == 'Giữ nguyên ảnh':
                            result = raw.copy()
                            drawn_points = None
                        elif choice == 'Cắt khung chữ nhật (%)':
                            x_default = (max(0, int(quad[:,0].min() / raw.width * 100)),
                                         min(100, int(quad[:,0].max() / raw.width * 100))) if quad is not None else (5,95)
                            y_default = (max(0, int(quad[:,1].min() / raw.height * 100)),
                                         min(100, int(quad[:,1].max() / raw.height * 100))) if quad is not None else (5,95)
                            if x_default[0] >= x_default[1]: x_default = (5,95)
                            if y_default[0] >= y_default[1]: y_default = (5,95)
                            xv = st.slider('X trái ↔ phải (%)', 0, 100, x_default, key='x_' + file_key)
                            yv = st.slider('Y trên ↕ dưới (%)', 0, 100, y_default, key='y_' + file_key)
                            try:
                                result = crop_rectangle(raw, xv[0], yv[0], xv[1], yv[1])
                                drawn_points = np.float32([[xv[0]*raw.width/100, yv[0]*raw.height/100],
                                    [xv[1]*raw.width/100, yv[0]*raw.height/100],
                                    [xv[1]*raw.width/100, yv[1]*raw.height/100],
                                    [xv[0]*raw.width/100, yv[1]*raw.height/100]])
                            except ValueError as exc:
                                st.error(str(exc))
                        elif choice == 'Chỉnh 4 góc (%)':
                            default = (quad / [raw.width, raw.height] * 100).tolist() if quad is not None else [
                                [8,20],[92,20],[92,80],[8,80]]
                            rows = [{'Góc':v, 'X (%)':round(float(pt[0]),1), 'Y (%)':round(float(pt[1]),1)}
                                    for v,pt in zip(['1 · Trái trên','2 · Phải trên','3 · Phải dưới','4 · Trái dưới'], default)]
                            st.caption('Sửa tọa độ 4 góc (tính theo % ảnh gốc). X tăng từ trái sang phải; Y tăng từ trên xuống dưới.')
                            changed = st.data_editor(rows, key='corners_' + file_key, hide_index=True,
                                use_container_width=True, num_rows='fixed', disabled=['Góc'],
                                column_config={
                                    'X (%)': st.column_config.NumberColumn('X (%)', min_value=0.0, max_value=100.0, step=0.5),
                                    'Y (%)': st.column_config.NumberColumn('Y (%)', min_value=0.0, max_value=100.0, step=0.5)})
                            try:
                                pts = np.float32([[float(p['X (%)']) * raw.width / 100,
                                                   float(p['Y (%)']) * raw.height / 100] for p in changed])
                                if cv2.contourArea(_order_quad(pts)) < raw.width * raw.height * 0.001:
                                    raise ValueError('Bốn góc không tạo thành vùng cắt hợp lệ.')
                                result = crop_quad(raw, pts)
                                drawn_points = pts
                            except (ValueError, TypeError, KeyError) as exc:
                                st.error(f'Tọa độ chưa hợp lệ: {exc}')
                        elif quad is None:
                            st.warning('Chưa xác định được khung. Hãy chọn Cắt khung chữ nhật hoặc Chỉnh 4 góc để loại bỏ nền.')
                        else:
                            st.caption(f'Chỉ số hình học của vùng đã nhận dạng: {confidence:.0%}. Cần xem lại nếu viền xanh không khớp mép thẻ.')
                    with col_preview:
                        pic1, pic2 = st.columns(2)
                        with pic1:
                            st.image(draw_detected_outline(raw, drawn_points), caption='Ảnh gốc / viền nhận dạng', use_container_width=True)
                        with pic2:
                            st.image(result, caption='Kết quả crop', use_container_width=True)
                    processed.append(result)
            st.subheader('2. Ghép mặt trước / mặt sau')
            pairs = []
            for ix in range(0, len(processed), 2):
                j = ix // 2
                a = processed[ix]
                b = processed[ix + 1] if ix + 1 < len(processed) else Image.new('RGB', a.size, 'white')
                cols = st.columns([1.5,1,1])
                cols[0].write(f'**Bộ {j + 1}** · ảnh {ix + 1} & {ix + 2 if ix + 1 < len(processed) else "trống"}')
                if cols[1].button('🔄 Đảo hai mặt', key=f'swap_{j}'):
                    st.session_state.swaps[j] = not st.session_state.swaps.get(j, False)
                if st.session_state.swaps.get(j):
                    a,b = b,a
                    cols[2].caption('Đã đảo mặt')
                pairs.append((a,b))
            pages = make_pages(pairs, pairs_per_page)
            st.success(f'Đã xếp {len(pairs)} bộ vào {len(pages)} trang A4, tối đa {pairs_per_page} bộ/trang.')
            if len(files) % 2:
                st.warning('Số ảnh lẻ: mặt sau của bộ cuối được để trắng.')
            left, right = st.columns([1,1], gap='large')
            with left:
                st.subheader('3. Tải kết quả')
                st.download_button('📄 TẢI PDF A4 (khuyến nghị)', data=create_pdf(pages),
                    file_name=f'{doc_type}_A4_{len(pairs)}_bo.pdf', mime='application/pdf',
                    type='primary', use_container_width=True)
                st.download_button('📝 TẢI WORD A4 (.docx)',
                    data=create_multi_docx(pairs, pairs_per_page).getvalue(),
                    file_name=f'{doc_type}_A4_{len(pairs)}_bo.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', use_container_width=True)
                if len(pages) == 1:
                    st.download_button('🖼️ TẢI PNG A4', data=export_png(pages[0]),
                        file_name=f'{doc_type}_A4_Trang_1.png', mime='image/png', use_container_width=True)
                else:
                    st.download_button('🗂️ TẢI ZIP TẤT CẢ TRANG PNG',
                        data=export_pages_zip(pages, doc_type), file_name=f'{doc_type}_TatCaTrang.zip',
                        mime='application/zip', use_container_width=True)
            with right:
                st.subheader('Xem trước bản A4')
                if len(pages) > 1:
                    page_index = st.selectbox('Chọn trang', range(len(pages)), format_func=lambda i:f'Trang {i+1}')
                else:
                    page_index = 0
                st.image(pages[page_index], caption=f'A4 · Trang {page_index + 1}/{len(pages)} · 300 DPI',
                         use_container_width=True)

    with tab_vneid:
        st.subheader('Crop khung 2 mặt GPLX / CCCD trên ảnh VNeID')
        f = st.file_uploader('Tải ảnh chụp màn hình VNeID', type=['jpg','jpeg','png','webp'], key='vneid_input')
        if f is not None:
            original = read_image(f.getvalue())
            auto, ok, outline = crop_vneid_combined_block(original)
            if not ok:
                st.warning('Chưa tìm đủ 2 khung. Không cắt theo vị trí đoán sẵn; hãy điều chỉnh thủ công.')
            mode = st.radio('Cách cắt', ['Tự động', 'Thủ công theo vùng (%)'], horizontal=True,
                            index=0 if ok else 1, key='vneid_mode')
            result = auto
            if mode == 'Thủ công theo vùng (%)':
                left, right = st.slider('Trái / phải (%)', 0, 100, (3,97), key='vneid_x')
                top, bottom = st.slider('Trên / dưới (%)', 0, 100, (8,80), key='vneid_y')
                result = crop_rectangle(original,left,top,right,bottom)
            col_ori, col_result = st.columns(2)
            with col_ori:
                st.image(original, caption='Ảnh màn hình gốc', use_container_width=True)
            with col_result:
                st.image(result, caption='Vùng xuất (kiểm tra trước khi tải)', use_container_width=True)
            st.download_button('🖼️ TẢI PNG KHUNG VNeID', data=export_png(result),
                file_name='VNeID_Khung_2_Mat.png', mime='image/png')
            st.download_button('📝 TẢI WORD KHUNG VNeID', data=create_single_cropped_docx(result),
                file_name='VNeID_Khung_2_Mat.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    main()
