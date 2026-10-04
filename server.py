@app.route("/")
def trang_chu():
    return """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>All Thủy Điện — Trợ lý dữ liệu</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #2563eb;
            --secondary: #7c3aed;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --bg-main: #f8fafc;
            --bg-card: #ffffff;
            --text-dark: #1e293b;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.05);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 16px;
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #eff6ff 0%, #faf5ff 100%);
            min-height: 100vh;
            color: var(--text-dark);
        }

        /* === HEADER === */
        .header {
            background: rgba(255,255,255,0.9);
            backdrop-filter: blur(10px);
            padding: 16px 24px;
            box-shadow: var(--shadow-sm);
            display: flex;
            align-items: center;
            gap: 14px;
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .logo {
            width: 44px;
            height: 44px;
            border-radius: 12px;
            background: linear-gradient(135deg, #2563eb, #7c3aed);
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 20px;
        }
        .header-text h1 { font-size: 18px; font-weight: 700; }
        .header-text p { font-size: 13px; color: var(--text-muted); margin-top: 2px; }

        /* === STEP PROGRESS === */
        .progress-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px 24px 0;
        }
        .progress-bar {
            display: flex;
            justify-content: space-between;
            position: relative;
        }
        .progress-bar::before {
            content: '';
            position: absolute;
            top: 18px;
            left: 10%;
            right: 10%;
            height: 2px;
            background: var(--border);
            z-index: 1;
        }
        .progress-step {
            display: flex;
            flex-direction: column;
            align-items: center;
            position: relative;
            z-index: 2;
            flex: 1;
        }
        .step-circle {
            width: 36px;
            height: 36px;
            border-radius: 50%;
            background: white;
            border: 2px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 600;
            font-size: 14px;
            margin-bottom: 6px;
            transition: all 0.3s ease;
        }
        .step-circle.active {
            background: var(--primary);
            border-color: var(--primary);
            color: white;
            box-shadow: 0 0 0 4px rgba(37,99,235,0.15);
        }
        .step-text {
            font-size: 12px;
            color: var(--text-muted);
            text-align: center;
            transition: color 0.3s;
        }
        .step-text.active {
            color: var(--primary);
            font-weight: 500;
        }

        /* === MAIN LAYOUT === */
        .main-container {
            max-width: 1200px;
            margin: 20px auto;
            padding: 0 24px 40px;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 24px;
        }

        /* === CARD === */
        .card {
            background: white;
            border-radius: var(--radius-lg);
            padding: 24px;
            box-shadow: var(--shadow-md);
            display: flex;
            flex-direction: column;
            height: calc(100vh - 200px);
            min-height: 600px;
        }
        .card-header {
            display: flex;
            align-items: center;
            gap: 10px;
            padding-bottom: 14px;
            margin-bottom: 18px;
            border-bottom: 2px solid transparent;
        }
        .card-header.blue { border-bottom-color: #dbeafe; }
        .card-header.purple { border-bottom-color: #f3e8ff; }
        .card-icon { font-size: 22px; }
        .card-title { font-size: 17px; font-weight: 700; }
        .card-title.blue { color: var(--primary); }
        .card-title.purple { color: var(--secondary); }

        /* === UPLOAD AREA === */
        .upload-area {
            border: 2px dashed var(--border);
            border-radius: var(--radius-md);
            padding: 32px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-bottom: 16px;
        }
        .upload-area:hover {
            border-color: var(--primary);
            background: #eff6ff;
            transform: translateY(-2px);
        }
        .upload-icon { font-size: 36px; margin-bottom: 10px; }
        .upload-text { color: var(--text-dark); font-weight: 500; }
        .upload-note { font-size: 12px; color: var(--text-muted); margin-top: 4px; }

        .file-info {
            display: none;
            align-items: center;
            gap: 10px;
            padding: 12px 16px;
            background: #ecfdf5;
            border-radius: var(--radius-sm);
            margin-bottom: 16px;
        }
        .file-info.show { display: flex; }
        .file-name { flex: 1; font-size: 14px; font-weight: 500; }
        .file-remove {
            border: none;
            background: none;
            color: var(--danger);
            font-size: 20px;
            cursor: pointer;
            padding: 0 6px;
            transition: transform 0.2s;
        }
        .file-remove:hover { transform: scale(1.2); }

        /* === QUICK BUTTONS === */
        .quick-actions {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
            margin-bottom: 18px;
        }
        .quick-btn {
            padding: 12px 14px;
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            background: white;
            cursor: pointer;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.25s ease;
            text-align: left;
        }
        .quick-btn:hover {
            border-color: var(--primary);
            background: #eff6ff;
            color: var(--primary);
            transform: translateY(-1px);
        }

        /* === CHAT AREA === */
        .chat-container {
            flex: 1;
            overflow-y: auto;
            padding: 4px 8px;
            margin-bottom: 16px;
        }
        .chat-container::-webkit-scrollbar { width: 4px; }
        .chat-container::-webkit-scrollbar-thumb {
            background: var(--border);
            border-radius: 4px;
        }

        .message {
            margin-bottom: 18px;
            display: flex;
            max-width: 96%;
            animation: msgIn 0.3s ease forwards;
            opacity: 0;
        }
        @keyframes msgIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .message.user { justify-content: flex-end; margin-left: auto; }
        .message.ai { justify-content: flex-start; margin-right: auto; }

        .bubble {
            padding: 14px 18px;
            border-radius: var(--radius-lg);
            line-height: 1.6;
            font-size: 14px;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .message.user .bubble {
            background: linear-gradient(135deg, #dbeafe, #e0e7ff);
            border-bottom-right-radius: 6px;
            color: #1e40af;
        }
        .message.ai .bubble {
            background: #f8fafc;
            border-bottom-left-radius: 6px;
            border: 1px solid var(--border);
        }
        .bubble.warning {
            background: #fffbeb;
            border-left: 3px solid var(--warning);
            color: #92400e;
        }
        .bubble.error {
            background: #fef2f2;
            border-left: 3px solid var(--danger);
            color: #b91c1c;
        }

        .file-tag {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 5px 12px;
            background: #dbeafe;
            color: #1d4ed8;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-bottom: 8px;
        }

        /* === DOWNLOAD BUTTONS === */
        .download-group {
            display: flex;
            gap: 10px;
            margin-top: 14px;
            flex-wrap: wrap;
        }
        .download-btn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 8px 16px;
            border-radius: 20px;
            text-decoration: none;
            font-size: 13px;
            font-weight: 600;
            transition: all 0.2s ease;
        }
        .download-btn:hover { transform: translateY(-2px); }
        .dl-word { background: #dbeafe; color: #1d4ed8; }
        .dl-excel { background: #d1fae5; color: #047857; }
        .dl-pdf { background: #fee2e2; color: #b91c1c; }

        /* === INPUT AREA === */
        .input-wrapper {
            display: flex;
            gap: 10px;
            align-items: flex-end;
        }
        textarea {
            flex: 1;
            min-height: 48px;
            max-height: 120px;
            padding: 12px 18px;
            border: 1px solid var(--border);
            border-radius: 24px;
            font-size: 14px;
            font-family: inherit;
            resize: none;
            outline: none;
            transition: border-color 0.2s, box-shadow 0.2s;
        }
        textarea:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(37,99,235,0.1);
        }
        .card.purple textarea:focus {
            border-color: var(--secondary);
            box-shadow: 0 0 0 3px rgba(124,58,237,0.1);
        }
        .send-btn {
            width: 44px;
            height: 44px;
            border-radius: 50%;
            border: none;
            background: var(--primary);
            color: white;
            cursor: pointer;
            font-size: 18px;
            transition: all 0.2s ease;
            flex-shrink: 0;
        }
        .send-btn:hover {
            background: #1d4ed8;
            transform: scale(1.08);
        }
        .send-btn:disabled {
            background: #cbd5e1;
            cursor: not-allowed;
            transform: none;
        }
        .card.purple .send-btn { background: var(--secondary); }
        .card.purple .send-btn:hover { background: #6d28d9; }

        /* === RESPONSIVE === */
        @media (max-width: 900px) {
            .main-container { grid-template-columns: 1fr; }
            .card { height: auto; min-height: 500px; }
        }
    </style>
</head>
<body>
    <!-- HEADER -->
    <header class="header">
        <div class="logo">⚡</div>
        <div class="header-text">
            <h1>All Thủy Điện</h1>
            <p>Xử lý dữ liệu thông minh & Tạo báo cáo tự động</p>
        </div>
    </header>

    <!-- PROGRESS STEPS -->
    <div class="progress-container">
        <div class="progress-bar">
            <div class="progress-step">
                <div class="step-circle active" id="s1">1</div>
                <div class="step-text active" id="st1">Tải tệp lên</div>
            </div>
            <div class="progress-step">
                <div class="step-circle" id="s2">2</div>
                <div class="step-text" id="st2">Phân tích</div>
            </div>
            <div class="progress-step">
                <div class="step-circle" id="s3">3</div>
                <div class="step-text" id="st3">Xem kết quả</div>
            </div>
            <div class="progress-step">
                <div class="step-circle" id="s4">4</div>
                <div class="step-text" id="st4">Tải báo cáo</div>
            </div>
        </div>
    </div>

    <!-- MAIN CONTENT -->
    <div class="main-container">
        <!-- LEFT CARD: PROCESS & REPORT -->
        <div class="card">
            <div class="card-header blue">
                <span class="card-icon">📊</span>
                <h2 class="card-title blue">Xử lý dữ liệu & Tạo báo cáo</h2>
            </div>

            <div class="upload-area" id="uploadZone">
                <div class="upload-icon">📎</div>
                <div class="upload-text">Nhấn để chọn hoặc kéo thả tệp</div>
                <div class="upload-note">Hỗ trợ: .docx .xlsx .txt .pdf</div>
            </div>
            <input type="file" id="fileInput" accept=".docx,.xlsx,.txt,.pdf" style="display:none;">

            <div class="file-info" id="fileDisplay">
                <span>📄</span>
                <span class="file-name" id="fileName"></span>
                <button class="file-remove" onclick="clearFile()">✕</button>
            </div>

            <div class="quick-actions">
                <button class="quick-btn" onclick="sendQuick('Sắp xếp dữ liệu theo tên thiết bị')">
                    📋 Sắp xếp dữ liệu
                </button>
                <button class="quick-btn" onclick="sendQuick('Tính thành tiền = số lượng × đơn giá')">
                    💰 Tính thành tiền
                </button>
                <button class="quick-btn" onclick="sendQuick('Lập báo cáo tổng hợp dữ liệu')">
                    📑 Lập báo cáo
                </button>
                <button class="quick-btn" onclick="sendQuick('Kiểm tra tình trạng và danh sách thiết bị')">
                    🔍 Kiểm tra thiết bị
                </button>
            </div>

            <div class="chat-container" id="reportChat">
                <div class="message ai">
                    <div class="bubble">
                        👋 Xin chào! Hãy tải tệp dữ liệu lên hoặc chọn một yêu cầu nhanh để bắt đầu.
                    </div>
                </div>
            </div>

            <div class="input-wrapper">
                <textarea id="reportInput" placeholder="Nhập yêu cầu... (Enter = Gửi, Shift+Enter = Xuống dòng)" 
                    onkeydown="handleReportKey(event)"></textarea>
                <button class="send-btn" id="reportBtn" onclick="sendReport()">➤</button>
            </div>
        </div>

        <!-- RIGHT CARD: FREE CHAT -->
        <div class="card purple">
            <div class="card-header purple">
                <span class="card-icon">💬</span>
                <h2 class="card-title purple">Trò chuyện với AI</h2>
            </div>

            <div class="chat-container" id="freeChat">
                <div class="message ai">
                    <div class="bubble">
                        👋 Tôi là AI trợ lý thông minh. Bạn có thể hỏi tôi bất kỳ điều gì nhé!
                    </div>
                </div>
            </div>

            <div class="input-wrapper">
                <textarea id="chatInput" placeholder="Đặt câu hỏi cho AI..." 
                    onkeydown="handleChatKey(event)"></textarea>
                <button class="send-btn" id="chatBtn" onclick="sendChat()">➤</button>
            </div>
        </div>
    </div>

    <script>
        let uploadedContent = "";
        let uploadedFileName = "";

        document.addEventListener('DOMContentLoaded', () => {
            document.getElementById('uploadZone').addEventListener('click', () => {
                document.getElementById('fileInput').click();
            });
            document.getElementById('fileInput').addEventListener('change', handleFileSelect);
        });

        // === Progress Steps ===
        function setStep(n) {
            for (let i = 1; i <= 4; i++) {
                const circle = document.getElementById('s'+i);
                const text = document.getElementById('st'+i);
                if (i <= n) {
                    circle.classList.add('active');
                    text.classList.add('active');
                } else {
                    circle.classList.remove('active');
                    text.classList.remove('active');
                }
            }
        }

        // === File Handling ===
        function handleFileSelect(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            uploadedFileName = file.name;
            const formData = new FormData();
            formData.append('file', file);

            fetch('/api/upload', { method: 'POST', body: formData })
                .then(res => res.json())
                .then(data => {
                    if (data.status === 'ok') {
                        uploadedContent = data.content;
                        document.getElementById('fileName').textContent = uploadedFileName;
                        document.getElementById('fileDisplay').classList.add('show');
                        setStep(2);
                    } else {
                        addReportMsg('ai', '❌ ' + (data.error || 'Lỗi tải tệp'));
                    }
                })
                .catch(err => addReportMsg('ai', '❌ Lỗi kết nối: ' + err.message));
        }

        function clearFile() {
            uploadedContent = "";
            uploadedFileName = "";
            document.getElementById('fileDisplay').classList.remove('show');
            document.getElementById('fileInput').value = "";
            setStep(1);
        }

        // === Display Helper ===
        function cleanHtmlTags(text) {
            return text.replace(/<span\b[^>]*>/gi, '').replace(/<\/span>/gi, '');
        }

        function addReportMsg(type, content, files = null) {
            const container = document.getElementById('reportChat');
            const msgDiv = document.createElement('div');
            msgDiv.className = 'message ' + type;

            let displayContent = cleanHtmlTags(content);
            if (type === 'user' && uploadedFileName) {
                displayContent = `<span class="file-tag">📄 ${uploadedFileName}</span>` + 
                    (content.trim() ? '\n' + displayContent : ' Phân tích nội dung tệp');
            }

            let downloadLinks = '';
            if (files && (files.word || files.excel || files.pdf)) {
                downloadLinks = '<div class="download-group">';
                if (files.word) downloadLinks += `<a href="${files.word}" class="download-btn dl-word" target="_blank">📄 Word</a>`;
                if (files.excel) downloadLinks += `<a href="${files.excel}" class="download-btn dl-excel" target="_blank">📊 Excel</a>`;
                if (files.pdf) downloadLinks += `<a href="${files.pdf}" class="download-btn dl-pdf" target="_blank">📕 PDF</a>`;
                downloadLinks += '</div>';
                setStep(4);
            }

            msgDiv.innerHTML = `<div class="bubble">${displayContent}${downloadLinks}</div>`;
            container.appendChild(msgDiv);
            container.scrollTop = container.scrollHeight;

            if (type === 'ai') {
                setStep(3);
            }
        }

        function addChatMsg(type, content) {
            const container = document.getElementById('freeChat');
            const msgDiv = document.createElement('div');
            msgDiv.className = 'message ' + type;
            const safeText = cleanHtmlTags(content)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
            msgDiv.innerHTML = `<div class="bubble">${safeText}</div>`;
            container.appendChild(msgDiv);
            container.scrollTop = container.scrollHeight;
        }

        // === Quick Buttons ===
        function sendQuick(text) {
            document.getElementById('reportInput').value = text;
            sendReport();
        }

        // === Keyboard Handling ===
        function handleReportKey(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendReport();
            }
        }
        function handleChatKey(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendChat();
            }
        }

        // === Send Report Request ===
        async function sendReport() {
            const input = document.getElementById('reportInput');
            const btn = document.getElementById('reportBtn');
            const message = input.value.trim();

            if (!message && !uploadedContent) return;

            // Show user message
            addReportMsg('user', message || 'Phân tích nội dung tệp');
            
            input.value = '';
            btn.disabled = true;
            btn.textContent = '⏳';

            try {
                const res = await fetch('/api/chat-bao-cao', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        message: message || 'Phân tích và xử lý nội dung tệp',
                        file_content: uploadedContent
                    })
                });

                const data = await res.json();
                let reply = data.reply || '';
                
                // Style warning/error messages
                const container = document.getElementById('reportChat');
                const lastMsg = container.lastElementChild;
                if (lastMsg && lastMsg.classList.contains('ai')) {
                    const bubble = lastMsg.querySelector('.bubble');
                    if (reply.includes('quota') || reply.includes('hết hạn')) {
                        bubble.classList.add('warning');
                    } else if (reply.includes('❌')) {
                        bubble.classList.add('error');
                    }
                }

                addReportMsg('ai', reply, { word: data.word, excel: data.excel, pdf: data.pdf });

                // Only clear file if NOT quota error
                if (!reply.includes('quota')) {
                    clearFile();
                }
            } catch (err) {
                addReportMsg('ai', '❌ Lỗi kết nối: ' + err.message);
            } finally {
                btn.disabled = false;
                btn.textContent = '➤';
            }
        }

        // === Send Free Chat ===
        async function sendChat() {
            const input = document.getElementById('chatInput');
            const btn = document.getElementById('chatBtn');
            const message = input.value.trim();
            if (!message) return;

            addChatMsg('user', message);
            input.value = '';
            btn.disabled = true;
            btn.textContent = '⏳';

            try {
                const res = await fetch('/api/chat-tu-do', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ message: message })
                });
                const data = await res.json();
                addChatMsg('ai', data.reply || '❌ Không có phản hồi');
            } catch (err) {
                addChatMsg('ai', '❌ Lỗi kết nối: ' + err.message);
            } finally {
                btn.disabled = false;
                btn.textContent = '➤';
            }
        }
    </script>
</body>
</html>
"""
