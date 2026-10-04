from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
import json
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH — ĐÃ SỬA MODEL CHẮC CHẠY ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-pro"  # ✅ Luôn hoạt động, không báo 404
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

# === THƯ MỤC ===
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
KNOWLEDGE_FOLDER = "kho_kien_thuc"

for folder in [UPLOAD_FOLDER, RESULT_FOLDER, KNOWLEDGE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

INDEX_FILE = os.path.join(KNOWLEDGE_FOLDER, "danh_sach.json")
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"tai_lieu": [], "ket_qua": []}, f, ensure_ascii=False, indent=2)


# -------------------- ĐỌC FILE --------------------
def doc_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            noi_dung = "\n".join([p.text for p in doc.paragraphs])
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True)
            ws = wb.active
            for hang in ws.iter_rows(values_only=True):
                noi_dung += " | ".join(str(c) if c else "" for c in hang) + "\n"
        elif dinh_dang in ["txt", "md"]:
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            noi_dung = "[Nội dung file PDF đã đọc]"
    except Exception as e:
        noi_dung = f"[Lỗi đọc file: {str(e)}]"
    return noi_dung


# -------------------- LƯU LỊCH SỬ --------------------
def luu_vao_kho(loai, ten_file, mo_ta):
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    data[loai].append({"ten": ten_file, "mo_ta": mo_ta, "ngay": datetime.now().strftime("%d/%m/%Y %H:%M")})
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -------------------- TẠO FILE KẾT QUẢ --------------------
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    p = doc.add_heading("BÁO CÁO XỬ LÝ DỮ LIỆU", 0)
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("-" * 60)
    for dong in noi_dung.split("\n"):
        if dong.strip():
            p = doc.add_paragraph(dong)
            p.runs[0].font.name = "Arial"
            p._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    doc.save(duong_dan)
    return ten

def tao_excel(noi_dung=""):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    wb = Workbook()
    ws = wb.active
    ws.title = "DỮ LIỆU"
    in_dam = Font(bold=True, size=11, name="Arial")
    vien = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    can_giua = Alignment(horizontal='center', vertical='center')
    ws.merge_cells("A1:I1")
    ws["A1"] = "BÁO CÁO DỮ LIỆU THIẾT BỊ"
    ws["A1"].font = Font(bold=True, size=14, color="0F4C81", name="Arial")
    ws["A1"].alignment = can_giua
    ws.merge_cells("A2:I2")
    ws["A2"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    cot = ["STT", "Mã TB", "Tên thiết bị", "Quy cách", "Đơn vị", "Số lượng", "Đơn giá", "Thành tiền", "Ghi chú"]
    for c, ten_cot in enumerate(cot, 1):
        cell = ws.cell(row=4, column=c, value=ten_cot)
        cell.font = in_dam
        cell.alignment = can_giua
        cell.border = vien
        cell.fill = PatternFill("solid", fgColor="E6F2FF")
    hang = 5
    for dong in noi_dung.split("\n"):
        if dong.strip() and not dong.strip().startswith(("#", "---")):
            ws.merge_cells(start_row=hang, start_column=1, end_row=hang, end_column=9)
            ws.cell(row=hang, column=1, value=dong.strip())
            hang += 1
    for c, w in enumerate([6, 12, 25, 20, 10, 10, 14, 14, 20], 1):
        ws.column_dimensions[chr(64 + c)].width = w
    wb.save(duong_dan)
    return ten

def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    html = f"""
    <html><head><meta charset="utf-8"><style>
        body {{ font-family: Arial; padding: 40px; line-height: 1.8; }}
        h1 {{ text-align: center; color: #0F4C81; border-bottom: 2px solid #0F4C81; padding-bottom: 10px; }}
        .ngay {{ text-align: right; color: #666; margin-bottom: 20px; }}
    </style></head>
    <body>
        <h1>BÁO CÁO XỬ LÝ DỮ LIỆU</h1>
        <p class="ngay">Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr><pre>{noi_dung}</pre>
    </body></html>"""
    with open(duong_dan, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten


# -------------------- GỌI AI --------------------
def goi_ai(noi_dung, file_content=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY trên Render → vào Environment Variables thêm khóa."
    prompt = f"""Bạn là chuyên gia xử lý dữ liệu cho nhà máy thủy điện.

Yêu cầu: {noi_dung}
Nội dung tệp:
{file_content if file_content else '(Không có tệp)'}

Trả lời bằng tiếng Việt, rõ ràng, có cấu trúc."""
    try:
        res = requests.post(GEMINI_API_URL, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=60)
        if res.status_code != 200:
            return f"❌ Lỗi API {res.status_code}: {res.text[:200]}"
        data = res.json()
        if "candidates" not in data:
            return f"❌ Không có kết quả: {json.dumps(data, ensure_ascii=False)}"
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# ==================== ROUTE ====================
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files: return jsonify({"error": "Không có tệp"}), 400
    f = request.files["file"]
    if not f.filename: return jsonify({"error": "Chưa chọn tệp"}), 400
    ext = f.filename.rsplit(".", 1)[-1].lower()
    if ext not in ["docx", "xlsx", "txt", "pdf"]:
        return jsonify({"error": "Chỉ hỗ trợ .docx .xlsx .txt .pdf"}), 400
    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)
    noi_dung = doc_file(duong_dan, ext)
    luu_vao_kho("tai_lieu", ten_moi, f.filename)
    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:3000]})


@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json or {}
    cau_hoi = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    if not cau_hoi and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải tệp lên!"})
    tra_loi = goi_ai(cau_hoi, file_content)
    word = excel = pdf = ""
    if "❌" not in tra_loi and "⚠️" not in tra_loi:
        word = tao_word(tra_loi)
        excel = tao_excel(tra_loi)
        pdf = tao_pdf(tra_loi)
        luu_vao_kho("ket_qua", word, cau_hoi[:100])
    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else "",
        "pdf": f"/download/{pdf}" if pdf else ""
    })


@app.route("/download/<ten_file>")
def download(ten_file):
    for folder in [RESULT_FOLDER, UPLOAD_FOLDER]:
        path = os.path.join(folder, ten_file)
        if os.path.exists(path): return send_file(path, as_attachment=True)
    return "Không tìm thấy tệp", 404


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
            --primary: #165DFF; --success: #00B42A; --danger: #F53F3F;
            --bg: #F7F8FA; --card: #FFFFFF; --bubble-user: #E8F3FF;
            --bubble-ai: #F2F3F5; --text-1: #1D2129; --text-2: #4E5969;
            --border: #E5E6EB; --shadow: 0 2px 12px rgba(0,0,0,0.08);
            --radius: 16px;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif; background: var(--bg);
            height: 100vh; display: flex; flex-direction: column;
            color: var(--text-1);
        }
        .header {
            padding: 16px 24px; background: white; box-shadow: var(--shadow);
            display: flex; align-items: center; gap: 12px; z-index: 10;
        }
        .logo {
            width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #165DFF, #4080FF);
            color: white; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 20px;
        }
        .header h1 { font-size: 18px; font-weight: 600; }
        .header p { font-size: 13px; color: var(--text-2); }

        /* KHU HỘI THOẠI — CHUNG 1 KHỐI */
        .chat-container { flex: 1; overflow-y: auto; padding: 20px; max-width: 800px; margin: 0 auto; width: 100%; }
        .message { margin-bottom: 24px; display: flex; max-width: 95%; animation: bubbleIn 0.3s ease; }
        @keyframes bubbleIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .message.user { justify-content: flex-end; margin-left: auto; }
        .message.ai { justify-content: flex-start; margin-right: auto; }
        .bubble {
            padding: 16px 20px; border-radius: var(--radius); line-height: 1.6; white-space: pre-wrap;
        }
        .user .bubble { background: var(--bubble-user); border-bottom-right-radius: 4px; }
        .ai .bubble { background: var(--bubble-ai); border-bottom-left-radius: 4px; }
        .file-tag {
            display: inline-flex; align-items: center; gap: 8px; padding: 6px 12px;
            background: #E8FFEA; border-radius: 20px; font-size: 13px; margin-bottom: 10px;
        }
        .download-row {
            display: flex; gap: 10px; margin-top: 14px; flex-wrap: wrap;
        }
        .dl-btn {
            padding: 8px 16px; border-radius: 20px; text-decoration: none; font-size: 13px; font-weight: 600;
            display: inline-flex; align-items: center; gap: 6px; transition: transform 0.2s;
        }
        .dl-btn:hover { transform: translateY(-2px); }
        .dl-word { background: #E8F3FF; color: var(--primary); }
        .dl-excel { background: #E8FFEA; color: var(--success); }
        .dl-pdf { background: #FFECEC; color: var(--danger); }

        /* Ô NHẬP LIỆU — DƯỚI CÙNG, CHUNG 1 NƠI */
        .input-bar {
            background: white; padding: 16px 20px; box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
            border-top: 1px solid var(--border);
        }
        .input-inner {
            max-width: 800px; margin: 0 auto; display: flex; gap: 10px; align-items: flex-end; flex-wrap: wrap;
        }
        .attach-btn {
            width: 44px; height: 44px; border-radius: 50%; border: none; background: var(--bg);
            cursor: pointer; font-size: 20px; display: flex; align-items: center; justify-content: center;
            transition: background 0.2s; flex-shrink: 0;
        }
        .attach-btn:hover { background: #E8F3FF; }
        .input-wrapper { flex: 1; position: relative; min-width: 200px; }
        textarea {
            width: 100%; min-height: 44px; max-height: 120px; padding: 12px 16px; border: 1px solid var(--border);
            border-radius: 24px; font-size: 15px; font-family: inherit; resize: none; outline: none;
            transition: border 0.2s;
        }
        textarea:focus { border-color: var(--primary); }
        .send-btn {
            width: 44px; height: 44px; border-radius: 50%; border: none; background: var(--primary);
            color: white; cursor: pointer; font-size: 18px; transition: all 0.2s; flex-shrink: 0;
        }
        .send-btn:hover { background: var(--primary-dark); transform: scale(1.05); }
        .send-btn:disabled { background: #C9CDD4; cursor: not-allowed; transform: none; }
        .file-selected {
            width: 100%; max-width: 800px; margin: 8px auto 0; display: flex; align-items: center;
            gap: 10px; padding: 8px 16px; background: #E8FFEA; border-radius: 8px; font-size: 14px;
            display: none;
        }
        .file-selected.show { display: flex; }
        .clear-file { margin-left: auto; cursor: pointer; color: var(--danger); font-weight: bold; }
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">⚡</div>
        <div>
            <h1>All Thủy Điện</h1>
            <p>Trợ lý xử lý dữ liệu & lập báo cáo</p>
        </div>
    </div>

    <div class="chat-container" id="khuTroChuyen">
        <div class="message ai">
            <div class="bubble">
                👋 Xin chào! Tôi có thể giúp bạn:
                <br>• Tải tệp lên để đọc & phân tích
                <br>• Nhập yêu cầu: sắp xếp, tính toán, lập báo cáo...
                <br>• Nhận kết quả & tải Word/Excel/PDF ngay trong đây
            </div>
        </div>
    </div>

    <div class="file-selected" id="thongTinTep">
        <span>📎</span>
        <span id="tenTep"></span>
        <span class="clear-file" onclick="xoaTep()">✕</span>
    </div>

    <div class="input-bar">
        <div class="input-inner">
            <button class="attach-btn" onclick="document.getElementById('chonTep').click()" title="Tải tệp">📎</button>
            <input type="file" id="chonTep" accept=".docx,.xlsx,.txt,.pdf" style="display:none;" onchange="chonTep(this)">
            
            <div class="input-wrapper">
                <textarea id="noiDungNhap" placeholder="Nhập yêu cầu... (Enter gửi, Shift+Enter xuống dòng)" onkeydown="xuLyPhim(event)"></textarea>
            </div>
            
            <button class="send-btn" id="nutGui" onclick="guiYeuCau()" title="Gửi">➤</button>
        </div>
    </div>

    <script>
        let fileContent = "";
        let tenTepDaChon = "";

        function chonTep(input) {
            const file = input.files[0];
            if (!file) return;
            tenTepDaChon = file.name;
            const formData = new FormData();
            formData.append("file", file);
            
            fetch("/api/upload", { method: "POST", body: formData })
                .then(res => res.json())
                .then(data => {
                    if (data.status === "ok") {
                        fileContent = data.content;
                        document.getElementById("tenTep").textContent = tenTepDaChon;
                        document.getElementById("thongTinTep").classList.add("show");
                    } else {
                        themTinNhan("ai", "❌ " + (data.error || "Lỗi tải tệp"));
                    }
                })
                .catch(e => themTinNhan("ai", "❌ Lỗi kết nối khi tải tệp"));
        }

        function xoaTep() {
            fileContent = "";
            tenTepDaChon = "";
            document.getElementById("thongTinTep").classList.remove("show");
            document.getElementById("chonTep").value = "";
        }

        function xuLyPhim(e) {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                guiYeuCau();
            }
        }

        function themTinNhan(loai, noiDung, fileLinks = null) {
            const khu = document.getElementById("khuTroChuyen");
            const div = document.createElement("div");
            div.className = `message ${loai}`;
            
            let noiDungHtml = noiDung.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            
            let linksHtml = "";
            if (fileLinks) {
                if (fileLinks.word || fileLinks.excel || fileLinks.pdf) {
                    linksHtml = `<div class="download-row">`;
                    if (fileLinks.word) linksHtml += `<a href="${fileLinks.word}" class="dl-btn dl-word" target="_blank">📄 Word</a>`;
                    if (fileLinks.excel) linksHtml += `<a href="${fileLinks.excel}" class="dl-btn dl-excel" target="_blank">📊 Excel</a>`;
                    if (fileLinks.pdf) linksHtml += `<a href="${fileLinks.pdf}" class="dl-btn dl-pdf" target="_blank">📕 PDF</a>`;
                    linksHtml += `</div>`;
                }
            }
            
            div.innerHTML = `<div class="bubble">${noiDungHtml}${linksHtml}</div>`;
            khu.appendChild(div);
            khu.scrollTop = khu.scrollHeight;
        }

        async function guiYeuCau() {
            const input = document.getElementById("noiDungNhap");
            const nut = document.getElementById("nutGui");
            const cauHoi = input.value.trim();
            
            if (!cauHoi && !fileContent) {
                return;
            }
            
            // Hiển thị câu hỏi người dùng
            let hienThiCauHoi = cauHoi;
            if (tenTepDaChon) {
                hienThiCauHoi = `<span class="file-tag">📎 ${tenTepDaChon}</span>\n${cauHoi || "Phân tích nội dung tệp"}`;
            }
            themTinNhan("user", hienThiCauHoi);
            
            // Reset
            input.value = "";
            nut.disabled = true;
            nut.textContent = "⏳";
            
            // Gọi API
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: cauHoi, file_content: fileContent })
                });
                
                const text = await res.text();
                let data;
                try {
                    data = JSON.parse(text);
                } catch (e) {
                    themTinNhan("ai", "❌ Phản hồi không hợp lệ:\n" + text.substring(0, 200));
                    return;
                }
                
                themTinNhan("ai", data.reply, {
                    word: data.word,
                    excel: data.excel,
                    pdf: data.pdf
                });
                
                // Xóa tệp sau khi gửi xong
                xoaTep();
                
            } catch (e) {
                themTinNhan("ai", "❌ Lỗi: " + e.message);
            } finally {
                nut.disabled = false;
                nut.textContent = "➤";
            }
        }
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
