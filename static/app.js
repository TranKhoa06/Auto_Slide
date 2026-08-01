document.addEventListener('DOMContentLoaded', () => {
    let pdfFile = null;
    let currentSessionId = "";
    let currentPdfPath = "";
    let scriptData = [];

    // Setup Dropzones
    function setupDropzone(id, fileType, callback) {
        const dropzone = document.getElementById(id);
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = fileType;
        fileInput.style.display = 'none';
        dropzone.appendChild(fileInput);

        dropzone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) {
                const file = e.target.files[0];
                callback(file);
                const p = dropzone.querySelector('p');
                p.innerText = `Đã chọn: ${file.name}`;
                p.style.color = 'var(--primary-color)';
                dropzone.style.borderColor = 'var(--primary-color)';
            }
        });
    }

    setupDropzone('pdf-dropzone', '.pdf,.pptx', (f) => pdfFile = f);

    // Sync slider
    document.getElementById('slide-range').addEventListener('input', (e) => {
        document.getElementById('slide-count').innerText = e.target.value;
    });

    const loadingOverlay = document.getElementById('loading-overlay');
    const loadingText = document.getElementById('loading-text');

    function showLoading(msg) {
        loadingText.innerText = msg;
        loadingOverlay.style.display = 'flex';
    }
    function hideLoading() {
        loadingOverlay.style.display = 'none';
    }

    // STEP 1: Generate Script & Images
    document.getElementById('btn-generate-script').addEventListener('click', async () => {
        if (!pdfFile) return Swal.fire('Lỗi', 'Vui lòng chọn file PDF hoặc PPTX', 'warning');
        const apiKey = document.getElementById('api-key-input').value.trim();
        if (!apiKey) return Swal.fire('Lỗi', 'Vui lòng nhập Gemini API Key', 'warning');
        
        showLoading('✨ Đang phân tích tài liệu & Soạn kịch bản...');
        
        try {
            // 1. Soạn kịch bản
            let fd = new FormData();
            fd.append('source_file', pdfFile);
            fd.append('slide_count', document.getElementById('slide-range').value);
            fd.append('api_key', apiKey);

            let res = await fetch('/generate_script', { method: 'POST', body: fd });
            if (!res.ok) throw new Error("Lỗi tạo kịch bản");
            let data = await res.json();
            
            currentSessionId = data.session_id;
            currentPdfPath = data.pdf_path;
            
            // 2. Lấy hình ảnh (Trích xuất & AI)
            showLoading('🎨 Đang trích xuất & tạo ảnh AI...');
            let imgFd = new FormData();
            imgFd.append('session_id', currentSessionId);
            imgFd.append('script', data.script);
            imgFd.append('pdf_path', currentPdfPath);
            
            let imgRes = await fetch('/generate_images', { method: 'POST', body: imgFd });
            if (!imgRes.ok) throw new Error("Lỗi tạo ảnh");
            let imgData = await imgRes.json();
            
            scriptData = JSON.parse(imgData.script);
            renderReviewUI();
            
            // Chuyển UI
            document.getElementById('step1-box').classList.remove('active-step');
            document.getElementById('step2-box').style.display = 'block';
            
            hideLoading();
            Swal.fire('Thành công', 'Kịch bản và hình ảnh đã sẵn sàng!', 'success');
            
        } catch (e) {
            hideLoading();
            Swal.fire('Lỗi', e.message, 'error');
        }
    });

    function renderReviewUI() {
        const container = document.getElementById('review-container');
        container.innerHTML = '';
        
        scriptData.forEach((slide, idx) => {
            const card = document.createElement('div');
            card.className = 'slide-review-card';
            
            // Left: Text content
            const textDiv = document.createElement('div');
            textDiv.innerHTML = `
                <div style="font-weight:600; color:var(--primary-color)">Slide ${idx+1}</div>
                <input type="text" style="width:100%; background:transparent; border:none; border-bottom:1px solid #333; color:white; font-size:18px; margin-top:10px" value="${slide.title}">
                <textarea rows="4">${slide.bullets ? slide.bullets.join('\\n') : ''}</textarea>
                <textarea rows="2" style="font-size:12px; opacity:0.8" placeholder="Notes...">${slide.speaker_notes}</textarea>
            `;
            
            // Right: Image content
            const imgDiv = document.createElement('div');
            const imgSrc = slide.image_url || '';
            const isAI = slide.image_source === 'ai_generated';
            
            imgDiv.innerHTML = `
                <div style="font-size:12px; margin-bottom:5px; color:#aaa">Minh họa (${isAI ? 'AI tạo' : 'Từ PDF'})</div>
                <img src="${imgSrc}" class="slide-img-preview" id="img-${idx}" onerror="this.src=''">
                ${isAI ? `<div class="img-actions">
                    <button class="btn-small" onclick="regenerateImage(${idx})">Tạo lại</button>
                    <button class="btn-small" onclick="removeImage(${idx})">Xóa</button>
                </div>` : ''}
            `;
            
            card.appendChild(textDiv);
            card.appendChild(imgDiv);
            container.appendChild(card);
        });
    }

    // Global funcs for inline onclick
    window.regenerateImage = function(idx) {
        const slide = scriptData[idx];
        const randomSeed = Math.floor(Math.random() * 100000);
        const prompt = encodeURIComponent(`Minimalist illustration for presentation slide about ${slide.title}. Modern corporate style, clean background, highly detailed`);
        const imgUrl = `https://image.pollinations.ai/prompt/${prompt}?width=800&height=600&nologo=true&seed=${randomSeed}`;
        
        document.getElementById(`img-${idx}`).src = imgUrl;
        scriptData[idx].image_url = imgUrl;
    }
    window.removeImage = function(idx) {
        document.getElementById(`img-${idx}`).src = '';
        scriptData[idx].image_url = '';
    }

    // BƯỚC 2: TẠO SLIDE
    document.getElementById('btn-generate-pptx').addEventListener('click', async () => {
        showLoading('✨ Đang lắp ráp & Căn chỉnh Template Natively...');
        
        try {
            let fd = new FormData();
            fd.append('session_id', currentSessionId);
            fd.append('script', JSON.stringify(scriptData));
            
            let res = await fetch('/generate_pptx', { method: 'POST', body: fd });
            if (!res.ok) throw new Error("Lỗi khi tạo PPTX");
            
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'AutoSlide_Premium.pptx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            
            hideLoading();
            confetti({ particleCount: 150, spread: 80, origin: { y: 0.6 } });
            Swal.fire('Hoàn tất!', 'Slide đã được tạo thành công!', 'success');
            
        } catch(e) {
            hideLoading();
            Swal.fire('Lỗi', e.message, 'error');
        }
    });
});
