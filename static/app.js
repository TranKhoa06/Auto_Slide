document.addEventListener('DOMContentLoaded', () => {
    const pdfDropzone = document.getElementById('pdf-dropzone');
    const templateDropzone = document.getElementById('template-dropzone');
    let pdfFile = null;
    let templateFile = null;

    // Helper function to handle dropzone logic
    function setupDropzone(dropzone, fileType, callback) {
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.style.borderColor = 'var(--primary-color)';
            dropzone.style.background = 'rgba(139, 92, 246, 0.05)';
            dropzone.style.transform = 'scale(1.02)';
        });

        dropzone.addEventListener('dragleave', () => {
            dropzone.style.borderColor = 'var(--border-color)';
            dropzone.style.background = '#f8fafc';
            dropzone.style.transform = 'scale(1)';
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.style.borderColor = 'var(--border-color)';
            dropzone.style.background = '#f8fafc';
            dropzone.style.transform = 'scale(1)';
            
            if (e.dataTransfer.files.length) {
                const file = e.dataTransfer.files[0];
                if (file.name.toLowerCase().endsWith(fileType)) {
                    callback(file);
                    dropzone.querySelector('p').innerText = `Đã chọn: ${file.name}`;
                    dropzone.querySelector('p').style.color = 'var(--primary-color)';
                } else {
                    Swal.fire({
                        icon: 'error',
                        title: 'Sai định dạng!',
                        text: `Vui lòng chọn file ${fileType.toUpperCase()}`
                    });
                }
            }
        });

        // Click to upload
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = fileType === '.pdf' ? 'application/pdf' : '.pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation';
        fileInput.style.display = 'none';
        dropzone.appendChild(fileInput);

        dropzone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) {
                const file = e.target.files[0];
                callback(file);
                dropzone.querySelector('p').innerText = `Đã chọn: ${file.name}`;
                dropzone.querySelector('p').style.color = 'var(--primary-color)';
            }
        });
    }

    // Setup Dropzones
    setupDropzone(pdfDropzone, '.pdf', (file) => { pdfFile = file; });
    setupDropzone(templateDropzone, '.pptx', (file) => { templateFile = file; });

    // Slider sync
    const slider = document.querySelector('input[type="range"]');
    const slideCountLabel = document.getElementById('slide-count');
    slider.addEventListener('input', (e) => {
        slideCountLabel.innerText = e.target.value;
    });

    // Bắn pháo hoa
    function fireConfetti() {
        var duration = 3 * 1000;
        var animationEnd = Date.now() + duration;
        var defaults = { startVelocity: 30, spread: 360, ticks: 60, zIndex: 0 };

        function randomInRange(min, max) {
            return Math.random() * (max - min) + min;
        }

        var interval = setInterval(function() {
            var timeLeft = animationEnd - Date.now();
            if (timeLeft <= 0) {
                return clearInterval(interval);
            }
            var particleCount = 50 * (timeLeft / duration);
            confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.1, 0.3), y: Math.random() - 0.2 } });
            confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.7, 0.9), y: Math.random() - 0.2 } });
        }, 250);
    }

    // Variables to store session
    let currentSessionId = "";

    // Elements
    const btnScript = document.getElementById('btn-generate-script');
    const btnPptx = document.getElementById('btn-generate-pptx');
    const statusDiv = document.getElementById('ai-status');
    const statusText = statusDiv.querySelector('p');
    const scriptContainer = document.getElementById('script-container');
    const scriptTextarea = document.getElementById('script-textarea');
    const templateContainer = document.getElementById('template-container');

    // BƯỚC 1: TẠO KỊCH BẢN
    btnScript.addEventListener('click', async function() {
        if (!pdfFile) {
            Swal.fire('Thiếu dữ liệu', 'Vui lòng tải lên tài liệu học tập (PDF)!', 'warning');
            return;
        }

        const slideCount = document.querySelector('input[type="range"]').value;
        const formData = new FormData();
        formData.append('pdf_file', pdfFile);
        formData.append('slide_count', slideCount);
        formData.append('api_key', '');

        // Loading UI
        this.style.display = 'none';
        statusDiv.style.display = 'block';
        statusText.innerText = "✨ Đang đọc tài liệu & Soạn kịch bản...";

        try {
            const response = await fetch('/generate_script', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.detail || 'Lỗi xử lý từ máy chủ');
            }

            const data = await response.json();
            currentSessionId = data.session_id;
            
            // Hiển thị kịch bản
            scriptTextarea.value = data.script;
            scriptContainer.style.display = 'block';
            templateContainer.style.display = 'block';
            
            // Chuyển nút bấm sang Bước 2
            statusDiv.style.display = 'none';
            btnPptx.style.display = 'block';

            Swal.fire({
                title: 'Kịch bản đã sẵn sàng!',
                text: 'Hãy kiểm tra nội dung và tải lên Mẫu thiết kế để tạo Slide.',
                icon: 'success',
                confirmButtonColor: '#10b981'
            });

        } catch (error) {
            Swal.fire('Lỗi', error.message, 'error');
            this.style.display = 'block';
            statusDiv.style.display = 'none';
        }
    });

    // BƯỚC 2: TẠO SLIDE
    btnPptx.addEventListener('click', async function() {
        if (!templateFile) {
            Swal.fire('Thiếu dữ liệu', 'Vui lòng tải lên Mẫu thiết kế PPTX!', 'warning');
            return;
        }

        const formData = new FormData();
        formData.append('session_id', currentSessionId);
        formData.append('script', scriptTextarea.value);
        formData.append('template_file', templateFile);
        formData.append('api_key', '');

        // Loading UI
        this.style.display = 'none';
        statusDiv.style.display = 'block';
        
        let step = 0;
        const messages = [
            "✨ Quét Không gian & Nhận diện Bố cục...",
            "✨ Đang Lắp ráp Kịch bản vào Không gian...",
            "✨ Tự động Căn chỉnh Chữ (Auto-Fit)...",
            "✨ Gần xong rồi, đang lưu file..."
        ];
        const msgInterval = setInterval(() => {
            step++;
            if(step < messages.length) statusText.innerText = messages[step];
        }, 3500);

        try {
            const response = await fetch('/generate_pptx', {
                method: 'POST',
                body: formData
            });
            clearInterval(msgInterval);

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.detail || 'Lỗi xử lý từ máy chủ');
            }

            statusText.innerText = "🎉 Hoàn tất! Đang tải file xuống...";
            
            // Tải file về
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'AutoSlide_Result.pptx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);

            fireConfetti();
            Swal.fire('Thành công!', 'Slide siêu việt của bạn đã hoàn thành.', 'success');

        } catch (error) {
            clearInterval(msgInterval);
            Swal.fire('Lỗi', error.message, 'error');
            statusText.innerText = "❌ Đã có lỗi xảy ra.";
        } finally {
            setTimeout(() => {
                this.style.display = 'block';
                statusDiv.style.display = 'none';
                statusText.innerText = "✨...";
            }, 3000);
        }
    });
});
