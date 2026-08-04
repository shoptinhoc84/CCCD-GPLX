# =========================================================
# TAB 1: DÀN TRANG A4 TỰ ĐỘNG (Xử lý từ 1 ảnh trở lên)
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

        # Tính toán số cặp (Nếu lẻ 1 ảnh, tự tạo ảnh trắng làm mặt còn lại)
        num_pairs = (num_files + 1) // 2
        card_pairs = []
        a4_canvases = []

        with col_left:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx in range(num_pairs):
                status_text.text(f"⏳ Đang xử lý Bộ {idx + 1}/{num_pairs}...")
                
                i1 = idx * 2
                i2 = i1 + 1

                # Đọc ảnh 1
                raw_img1 = Image.open(uploaded_files[i1]).convert("RGB")
                opt_1 = optimize_image_size(raw_img1)
                crop_1 = crop_card_fast(np.array(opt_1))
                pil_f = Image.fromarray(crop_1)

                # Đọc ảnh 2 (Nếu thiếu ảnh thứ 2 thì tạo ảnh trắng)
                if i2 < num_files:
                    raw_img2 = Image.open(uploaded_files[i2]).convert("RGB")
                    opt_2 = optimize_image_size(raw_img2)
                    crop_2 = crop_card_fast(np.array(opt_2))
                    pil_b = Image.fromarray(crop_2)
                    del raw_img2, opt_2, crop_2
                else:
                    # Tạo ảnh trắng có kích thước tương đương ảnh thứ nhất
                    pil_b = Image.new("RGB", pil_f.size, (255, 255, 255))

                if st.session_state["swap_dict"].get(idx, False):
                    pil_f, pil_b = pil_b, pil_f

                card_pairs.append((pil_f, pil_b))

                del raw_img1, opt_1, crop_1
                gc.collect()

                progress_bar.progress((idx + 1) / num_pairs)

            status_text.empty()
            progress_bar.empty()

        # Tạo trang A4 preview
        for i in range(0, len(card_pairs), 2):
            chunk = card_pairs[i:i + 2]
            canvas = create_a4_canvas_horizontal(chunk)
            a4_canvases.append(canvas)

        docx_io = create_multi_docx(card_pairs)

        with download_area:
            if num_files % 2 != 0:
                st.warning("⚠️ Số lượng ảnh lẻ, hệ thống đã tự ghép ảnh trống làm mặt còn lại.")
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
