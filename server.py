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

# ==================== CẤU HÌNH BIẾN MÔI TRƯỜNG ====================
# Gemini — ưu tiên chính
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-2.0-flash-exp"
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

# Groq / Llama 3 (META) — miễn phí, tốc độ nhanh
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "llama-3.3-70b-versatile"

# OpenAI GPT
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-3.5-turbo"

UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok())

# Lưu trạng thái Google Sheets
sheets_state = {}

# ==================== GỌI DỊCH VỤ AI ====================
def goi_gemini(prompt, file_content="", he_thong=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY trên Render → vào Environment thêm khóa từ aistudio.google.com/apikey"
    full_prompt = f"""{he_thong or "Trả lời bằng tiếng Việt rõ ràng, có cấu trúc."}

Yêu cầu: {prompt}
Nội dung tham khảo:
{file_content[:4000] if file_content else '(Không có tệp)'}"""
    try:
        res = requests.post(
            GEMINI_API_URL,
            json={"contents": [{"parts": [{"text": full_prompt}]}]},
            timeout=120
        )
        if res.status_code == 429:
            return "⚠️ Gemini hết quota → tạo khóa mới tại aistudio.google.com/apikey"
        if res.status_code != 200:
            return f"❌ Lỗi Gemini {res.status_code}"
        data = res.json()
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as e:
        return f"❌ Lỗi kết nối Gemini: {str(e)}"


def goi_groq(prompt, he_thong=""):
    if not GROQ_API_KEY:
        return "⚠️ Chưa đặt GROQ_API_KEY → lấy tại console.groq.com/keys"
    try:
        res = requests.post(
            GROQ_API_URL,
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": GROQ_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt, rõ ràng, dễ hiểu."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 4000
            },
            timeout=120
        )
        if res.status_code != 200:
            return f"❌ Lỗi Groq {res.status_code}"
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ Lỗi kết nối Groq: {str(e)}"


def goi_gpt(prompt, he_thong=""):
    if not OPENAI_API_KEY:
        return "⚠️ Chưa đặt OPENAI_API_KEY → lấy tại platform.openai.com/api-keys"
    try:
        res = requests.post(
            OPENAI_API_URL,
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": OPENAI_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Bạn là chuyên gia soạn thảo văn bản, viết chuẩn mực, rõ ràng."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7
            },
            timeout=120
        )
        if res.status_code != 200:
            return f"❌ Lỗi GPT {res.status_code}"
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ Lỗi kết nối GPT: {str(e)}"


# ==================== XỬ LÝ GOOGLE SHEETS ====================
def doc_google_sheet(sheet_url):
    try:
        # Trích xuất ID từ URL
        if "docs.google.com/spreadsheets/d/" not in sheet_url:
            return None, "❌ Link Google Sheets không đúng định dạng"
        sheet_id = sheet_url.split("/d/")[1].split("/")[0]
        # Sử dụng export CSV để đọc nhanh, không cần OAuth phức tạp
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        res = requests.get(csv_url, timeout=30)
        if res.status_code != 200:
            return None, "❌ Không đọc được Sheet → Kiểm tra quyền truy cập: Chia sẻ thành 'Bất kỳ ai có link'"
        return res.text, None
    except Exception as e:
        return None, f"❌ Lỗi đọc Sheet: {str(e)}"


# ==================== TẠO TỆP BÁO CÁO ====================
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    p = doc.add_heading("BÁO CÁO", 0)
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("-" * 60)
    for dong in noi_dung.split("\n"):
        if dong.strip():
            p = doc.add_paragraph(dong.strip())
            for run in p.runs:
                run.font.name = "Arial"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    doc.save(duong_dan)
    return ten

def tao_excel(noi_dung=""):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    wb = Workbook()
    ws = wb.active
    ws.title = "BÁO CÁO"
    ws["A1"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    hang = 3
    for dong in noi_dung.split("\n"):
        if dong.strip():
            ws.cell(row=hang, column=1, value=dong.strip())
            hang += 1
    wb.save(duong_dan)
    return ten

def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    html = f"""<html><head><meta charset="utf-8"><style>
    body {{ font-family: Arial; padding: 30px; line-height: 1.6; }}
    h1 {{ text-align: center; color: #0F4C81; }}
    </style></head>
    <body><h1>BÁO CÁO</h1><p>Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p><hr><pre>{noi_dung[:6000]}</pre></body></html>"""
    with open(duong_dan, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten


# ==================== ROUTE API ====================
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có tệp"}), 400
    f = request.files["file"]
    ext = f.filename.rsplit(".", 1)[-1].lower()
    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)
    
    noi_dung = ""
    if ext == "xlsx":
        wb = load_workbook(duong_dan, data_only=True, read_only=True)
        ws = wb.active
        for hang in ws.iter_rows(values_only=True):
            noi_dung += " | ".join(str(c) if c else "" for c in hang) + "\n"
        wb.close()
    elif ext == "txt":
        with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
            noi_dung = f.read()
    
    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:5000]})


@app.route("/api/connect-sheet", methods=["POST"])
def connect_sheet():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "Vui lòng dán link Google Sheets"}), 400
    noi_dung, loi = doc_google_sheet(url)
    if loi:
        return jsonify({"error": loi}), 400
    sheet_id = url.split("/d/")[1].split("/")[0]
    sheets_state[sheet_id] = noi_dung
    return jsonify({"status": "ok", "id": sheet_id, "content": noi_dung[:5000]})


@app.route("/api/chat-data", methods=["POST"])
def chat_data():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    sheet_content = data.get("sheet_content", "")
    full_data = f"{file_content}\n---\n{sheet_content}" if (file_content or sheet_content) else ""
    
    if not msg and not full_data:
        return jsonify({"reply": "Vui lòng nhập yêu cầu, tải tệp hoặc dán link Google Sheets!"})
    
    tra_loi = goi_gemini(msg, full_data, he_thong="""Bạn là chuyên gia phân tích dữ liệu và tạo báo cáo cho nhà máy thủy điện. 
    Trả lời rõ ràng, có cấu trúc, tóm tắt số liệu quan trọng. Xuất kết quả phù hợp dạng bảng khi có dữ liệu.""")
    
    word = excel = pdf = ""
    if "❌" not in tra_loi and "⚠️" not in tra_loi:
        word = tao_word(tra_loi)
        excel = tao_excel(tra_loi)
        pdf = tao_pdf(tra_loi)
    
    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else "",
        "pdf": f"/download/{pdf}" if pdf else ""
    })


@app.route("/api/chat-doc", methods=["POST"])
def chat_doc():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    if not msg:
        return jsonify({"reply": "Vui lòng nhập yêu cầu soạn thảo!"})
    # Ưu tiên GPT, dự phòng Groq
    tra_loi = goi_gpt(msg, he_thong="Bạn là chuyên gia soạn thảo văn bản hành chính, hợp đồng, thư từ. Viết chuẩn mực, đầy đủ, đúng văn phong Việt Nam.")
    if "⚠️" in tra_loi or "❌" in tra_loi:
        tra_loi = goi_groq(msg, he_thong="Bạn là chuyên gia soạn thảo văn bản hành chính, hợp đồng, thư từ. Viết chuẩn mực, đầy đủ, đúng văn phong Việt Nam.")
    return jsonify({"reply": tra_loi})


@app.route("/api/chat-tender", methods=["POST"])
def chat_tender():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    if not msg:
        return jsonify({"reply": "Vui lòng nhập yêu cầu về quy trình đấu thầu!"})
    # Ưu tiên Gemini, dự phòng Groq
    tra_loi = goi_gemini(msg, he_thong="Bạn là chuyên gia tư vấn quy trình đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, hồ sơ cần chuẩn bị, lưu ý pháp lý.")
    if "⚠️" in tra_loi or "❌" in tra_loi:
        tra_loi = goi_groq(msg, he_thong="Bạn là chuyên gia tư vấn quy trình đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, hồ sơ cần chuẩn bị, lưu ý pháp lý.")
    return jsonify({"reply": tra_loi})


@app.route("/api/chat-equip", methods=["POST"])
def chat_equip():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    if not msg and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải danh sách thiết bị!"})
    # Ưu tiên Groq/Llama, dự phòng Gemini
    tra_loi = goi_groq(f"{msg}\nDữ liệu tham khảo:\n{file_content[:3000]}", he_thong="Bạn là chuyên gia quản lý thiết bị nhà máy. Phân loại, theo dõi tình trạng, đề xuất kế hoạch bảo trì, tính tuổi thọ thiết bị.")
    if "⚠️" in tra_loi or "❌" in tra_loi:
        tra_loi = goi_gemini(msg, file_content, he_thong="Bạn là chuyên gia quản lý thiết bị nhà máy. Phân loại, theo dõi tình trạng, đề xuất kế hoạch bảo trì, tính tuổi thọ thiết bị.")
    return jsonify({"reply": tra_loi})


@app.route("/download/<ten_file>")
def download(ten_file):
    for folder in [RESULT_FOLDER, UPLOAD_FOLDER]:
        path = os.path.join(folder, ten_file)
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    return "Không tìm thấy tệp", 404


@app.route("/")
def trang_chu():
    return r"""
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hệ thống hỗ trợ toàn diện</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
--c1: #2563eb; --c1-light: #dbeafe; --c1-bg: #f0f7ff;
--c2: #16a34a; --c2-light: #dcfce7; --c2-bg: #f0fdf4;
--c3: #9333ea; --c3-light: #f3e8ff; --c3-bg: #faf5ff;
--c4: #f59e0b; --c4-light: #fef3c7; --c4-bg: #fffbeb;
--border: #e2e8f0; --shadow: 0 4px 20px rgba(0,0,0,0.06);
--radius: 16px;
}
* { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Inter', sans-serif; }
body { background: linear-gradient(135deg, #f0f7ff, #faf5ff); min-height: 100vh; padding: 20px; }
.header { text-align: center; margin-bottom: 24px; }
.header h1 { font-size: 24px; font-weight: 700; color: #1e293b; }
.header p { color: #64748b; margin-top: 4px; }
.grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; max-width: 1900px; margin: 0 auto; }
@media (max-width: 1400px) { .grid { grid-template-columns: repeat(2, 1fr); } }
@media (max-width: 768px) { .grid { grid-template-columns: 1fr; } }
.card { background: white; border-radius: var(--radius); padding: 20px; box-shadow: var(--shadow); display: flex; flex-direction: column; height: calc(100vh - 140px); min-height: 600px; }
.card-c1 { border-top: 4px solid var(--c1); }
.card-c2 { border-top: 4px solid var(--c2); }
.card-c3 { border-top: 4px solid var(--c3); }
.card-c4 { border-top: 4px solid var(--c4); }
.card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 2px solid; }
.card-c1 .card-head { border-bottom-color: var(--c1-light); }
.card-c2 .card-head { border-bottom-color: var(--c2-light); }
.card-c3 .card-head { border-bottom-color: var(--c3-light); }
.card-c4 .card-head { border-bottom-color: var(--c4-light); }
.card-icon { font-size: 22px; }
.card-title { font-size: 16px; font-weight: 700; }
.card-c1 .card-title { color: var(--c1); }
.card-c2 .card-title { color: var(--c2); }
.card-c3 .card-title { color: var(--c3); }
.card-c4 .card-title { color: var(--c4); }
.card-ai { font-size: 11px; color: #94a3b8; margin-left: auto; }
.upload-zone { border: 2px dashed var(--border); border-radius: 12px; padding: 14px; text-align: center; cursor: pointer; margin-bottom: 10px; transition: 0.2s; }
.card-c1 .upload-zone:hover { border-color: var(--c1); background: var(--c1-bg); }
.upload-zone p { font-size: 13px; color: #64748b; }
.file-bar { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-radius: 8px; margin-bottom: 10px; font-size: 13px; }
.card-c1 .file-bar { background: var(--c1-light); }
.file-bar button { margin-left: auto; background: none; border: none; font-size: 18px; cursor: pointer; color: #ef4444; }
.sheet-bar { display: flex; gap: 8px; margin-bottom: 10px; }
.sheet-bar input { flex: 1; padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 13px; }
.sheet-bar button { padding: 10px 14px; border: none; border-radius: 8px; background: var(--c1); color: white; cursor: pointer; font-weight: 500; }
.quick-btns { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 12px; }
.q-btn { padding: 10px; border: 1px solid var(--border); border-radius: 8px; background: white; cursor: pointer; font-size: 12px; transition: 0.2s; }
.card-c1 .q-btn:hover { border-color: var(--c1); background: var(--c1-bg); }
.card-c2 .q-btn:hover { border-color: var(--c2); background: var(--c2-bg); }
.card-c3 .q-btn:hover { border-color: var(--c3); background: var(--c3-bg); }
.card-c4 .q-btn:hover { border-color: var(--c4); background: var(--c4-bg); }
.chat-area { flex: 1; overflow-y: auto; padding: 4px; margin-bottom: 12px; }
.msg { margin-bottom: 14px; max-width: 96%; animation: fadeIn 0.3s forwards; opacity: 0; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
.msg.user { margin-left: auto; }
.msg.ai { margin-right: auto; }
.bubble { padding: 12px 16px; border-radius: 14px; line-height: 1.5; font-size: 13px; white-space: pre-wrap; word-break: break-word; }
.msg.user .bubble { background: linear-gradient(135deg, #e0f2fe, #e0e7ff); border-bottom-right-radius: 6px; }
.msg.ai .bubble { background: #f8fafc; border: 1px solid var(--border); border-bottom-left-radius: 6px; }
.bubble.warn { background: #fffbeb; border-left: 3px solid #f59e0b; color: #92400e; }
.bubble.err { background: #fef2f2; border-left: 3px solid #ef4444; color: #b91c1c; }
.dl-group { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
.dl-btn { display: inline-flex; align-items: center; gap: 4px; padding: 6px 12px; border-radius: 20px; text-decoration: none; font-size: 12px; font-weight: 600; }
.dl-word { background: #dbeafe; color: #1d4ed8; }
.dl-excel { background: #dcfce7; color: #15803d; }
.dl-pdf { background: #fee2e2; color: #b91c1c; }
.input-row { display: flex; gap: 8px; align-items: flex-end; }
textarea { flex: 1; min-height: 44px; max-height: 100px; padding: 10px 16px; border: 1px solid var(--border); border-radius: 22px; font-size: 13px; resize: none; outline: none; }
textarea:focus { border-color: var(--c1); box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
.card-c2 textarea:focus { border-color: var(--c2); box-shadow: 0 0 0 3px rgba(22,163,74,0.1); }
.card-c3 textarea:focus { border-color: var(--c3); box-shadow: 0 0 0 3px rgba(147,51,234,0.1); }
.card-c4 textarea:focus { border-color: var(--c4); box-shadow: 0 0 0 3px rgba(245,158,11,0.1); }
.send-btn { width: 40px; height: 40px; border-radius: 50%; border: none; color: white; cursor: pointer; font-size: 16px; flex-shrink: 0; }
.send-btn:disabled { opacity: 0.5; cursor: not-allowed; }
.card-c1 .send-btn { background: var(--c1); }
.card-c2 .send-btn { background: var(--c2); }
.card-c3 .send-btn { background: var(--c3); }
.card-c4 .send-btn { background: var(--c4); }
.send-btn:hover { transform: scale(1.05); }
.hidden { display: none !important; }
</style>
</head>
<body>
<div class="header">
<h1>⚡ All Thủy Điện — Hệ thống hỗ trợ toàn diện</h1>
<p>Xử lý dữ liệu · Soạn thảo văn bản · Đấu thầu · Quản lý thiết bị</p>
</div>
<div class="grid">

<!-- CỘT 1: XỬ LÝ DỮ LIỆU & TẠO BÁO CÁO → GEMINI + GOOGLE SHEETS -->
<div class="card card-c1">
<div class="card-head">
<span class="card-icon">📊</span>
<h3 class="card-title">Xử lý dữ liệu & Tạo báo cáo</h3>
<span class="card-ai">Powered by Gemini</span>
</div>

<div class="upload-zone" id="uploadZone1" onclick="document.getElementById('fileInput1').click()">
<p>📎 Nhấn chọn hoặc kéo thả tệp (.xlsx, .txt)</p>
</div>
<input type="file" id="fileInput1" accept=".xlsx,.txt" class="hidden">
<div class="file-bar hidden" id="fileBar1">
<span id="fName1"></span>
<button onclick="clearFile1()">✕</button>
</div>

<div class="sheet-bar">
<input type="text" id="sheetUrl" placeholder="🔗 Dán link Google Sheets...">
<button onclick="connectSheet()">Kết nối</button>
</div>
<div class="file-bar hidden" id="sheetBar">
<span>✅ Google Sheets đã kết nối</span>
<button onclick="clearSheet()">✕</button>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickData('Sắp xếp dữ liệu theo thứ tự')">📋 Sắp xếp dữ liệu</button>
<button class="q-btn" onclick="quickData('Tính tổng thành tiền = số lượng × đơn giá')">💰 Tính thành tiền</button>
<button class="q-btn" onclick="quickData('Lập báo cáo tổng hợp đầy đủ')">📑 Lập báo cáo</button>
<button class="q-btn" onclick="quickData('Phân tích số liệu và đưa ra nhận xét')">📈 Phân tích số liệu</button>
</div>

<div class="chat-area" id="chat1">
<div class="msg ai"><div class="bubble">👋 Tải tệp, dán link Google Sheets hoặc nhập yêu cầu để bắt đầu phân tích nhé!</div></div>
</div>

<div class="input-row">
<textarea id="input1" placeholder="Nhập yêu cầu... (Enter = Gửi)" onkeydown="handleKey(event, sendData)"></textarea>
<button class="send-btn" id="btn1" onclick="sendData()">➤</button>
</div>
</div>

<!-- CỘT 2: SOẠN THẢO VĂN BẢN → GPT / CLAUDE -->
<div class="card card-c2">
<div class="card-head">
<span class="card-icon">✍️</span>
<h3 class="card-title">Soạn thảo văn bản</h3>
<span class="card-ai">Powered by GPT</span>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickDoc('Soạn thảo công văn gửi cấp trên')">📝 Công văn</button>
<button class="q-btn" onclick="quickDoc('Soạn thảo hợp đồng mua bán thiết bị')">📄 Hợp đồng</button>
<button class="q-btn" onclick="quickDoc('Viết báo cáo tiến độ thực hiện dự án')">📈 Báo cáo tiến độ</button>
<button class="q-btn" onclick="quickDoc('Soạn thảo thư mời họp biên bản cuộc họp')">📋 Thư & Biên bản</button>
</div>

<div class="chat-area" id="chat2">
<div class="msg ai"><div class="bubble">👋 Tôi sẽ giúp bạn soạn thảo văn bản chuẩn mực, đúng thể thức Việt Nam. Bạn cần viết gì?</div></div>
</div>

<div class="input-row">
<textarea id="input2" placeholder="Bạn cần soạn thảo gì...?" onkeydown="handleKey(event, sendDoc)"></textarea>
<button class="send-btn" id="btn2" onclick="sendDoc()">➤</button>
</div>
</div>

<!-- CỘT 3: QUY TRÌNH ĐẤU THẦU → GEMINI / LLAMA -->
<div class="card card-c3">
<div class="card-head">
<span class="card-icon">🏆</span>
<h3 class="card-title">Quy trình đấu thầu</h3>
<span class="card-ai">Powered by Gemini + Llama</span>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickTender('Giải thích quy trình đấu thầu từ bước chuẩn bị đến kết quả')">📋 Toàn bộ quy trình</button>
<button class="q-btn" onclick="quickTender('Danh mục hồ sơ cần chuẩn bị cho mời thầu')">📑 Hồ sơ mời thầu</button>
<button class="q-btn" onclick="quickTender('Lưu ý pháp lý và rủi ro thường gặp')">⚖️ Pháp lý & Rủi ro</button>
<button class="q-btn" onclick="quickTender('Mẫu biểu mẫu thông dụng trong đấu thầu')">📄 Biểu mẫu</button>
</div>

<div class="chat-area" id="chat3">
<div class="msg ai"><div class="bubble">👋 Tôi hướng dẫn chi tiết quy trình đấu thầu theo quy định Việt Nam. Bạn cần hỗ trợ về bước nào?</div></div>
</div>

<div class="input-row">
<textarea id="input3" placeholder="Hỏi về quy trình đấu thầu...?" onkeydown="handleKey(event, sendTender)"></textarea>
<button class="send-btn" id="btn3" onclick="sendTender()">➤</button>
</div>
</div>

<!-- CỘT 4: QUẢN LÝ THIẾT BỊ → LLAMA / DOLA -->
<div class="card card-c4">
<div class="card-head">
<span class="card-icon">🔧</span>
<h3 class="card-title">Quản lý thiết bị</h3>
<span class="card-ai">Powered by Llama</span>
</div>

<div class="upload-zone" id="uploadZone4" onclick="document.getElementById('fileInput4').click()">
<p>📎 Tải danh sách thiết bị (.xlsx, .txt)</p>
</div>
<input type="file" id="fileInput4" accept=".xlsx,.txt" class="hidden">
<div class="file-bar hidden" id="fileBar4">
<span id="fName4"></span>
<button onclick="clearFile4()">✕</button>
</div>

<div class="quick-btns">
<button class="q-btn" onclick="quickEquip('Phân loại thiết bị theo nhóm chức năng')">📊 Phân loại thiết bị</button>
<button class="q-btn" onclick="quickEquip('Đề xuất kế hoạch bảo trì định kỳ')">🛠️ Kế hoạch bảo trì</button>
<button class="q-btn" onclick="quickEquip('Đánh giá tình trạng và rủi ro thiết bị')">⚠️ Đánh giá rủi ro</button>
<button class="q-btn" onclick="quickEquip('Tính tuổi thọ còn lại và đề xuất thay thế')">🔄 Tuổi thọ & Thay thế</button>
</div>

<div class="chat-area" id="chat4">
<div class="msg ai"><div class="bubble">👋 Tải danh sách thiết bị hoặc nhập yêu cầu quản lý — tôi sẽ phân tích và đề xuất kế hoạch chi tiết nhé!</div></div>
</div>

<div class="input-row">
<textarea id="input4" placeholder="Nhập yêu cầu quản lý thiết bị...?" onkeydown="handleKey(event, sendEquip)"></textarea>
<button class="send-btn" id="btn4" onclick="sendEquip()">➤</button>
</div>
</div>

</div>

<script>
let file1Content = "";
let file4Content = "";
let sheetContent = "";

document.getElementById('fileInput1').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    const fd = new FormData(); fd.append('file', f);
    fetch('/api/upload', {method:'POST', body:fd}).then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            file1Content = d.content;
            document.getElementById('fName1').textContent = d.name;
            document.getElementById('fileBar1').classList.remove('hidden');
        }
    });
});
document.getElementById('fileInput4').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    const fd = new FormData(); fd.append('file', f);
    fetch('/api/upload', {method:'POST', body:fd}).then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            file4Content = d.content;
            document.getElementById('fName4').textContent = d.name;
            document.getElementById('fileBar4').classList.remove('hidden');
        }
    });
});
function clearFile1(){ file1Content=''; document.getElementById('fileBar1').classList.add('hidden'); document.getElementById('fileInput1').value=''; }
function clearFile4(){ file4Content=''; document.getElementById('fileBar4').classList.add('hidden'); document.getElementById('fileInput4').value=''; }
function clearSheet(){ sheetContent=''; document.getElementById('sheetBar').classList.add('hidden'); document.getElementById('sheetUrl').value=''; }

function connectSheet(){
    const url = document.getElementById('sheetUrl').value.trim();
    if(!url) return;
    fetch('/api/connect-sheet', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({url})})
    .then(r=>r.json()).then(d=>{
        if(d.status==='ok'){
            sheetContent = d.content;
            document.getElementById('sheetBar').classList.remove('hidden');
        } else alert(d.error);
    });
}

function handleKey(e, fn){ if(e.key==='Enter' && !e.shiftKey){ e.preventDefault(); fn(); } }
function addMsg(chatId, type, text, links=''){
    const c = document.getElementById(chatId);
    const d = document.createElement('div'); d.className = 'msg '+type;
    let cls = '';
    if(text.includes('⚠️')) cls=' warn';
    else if(text.includes('❌')) cls=' err';
    d.innerHTML = `<div class="bubble${cls}">${text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}${links}</div>`;
    c.appendChild(d); c.scrollTop = c.scrollHeight;
}
function dlLinks(d){
    let h='';
    if(d.word) h += `<a href="${d.word}" class="dl-btn dl-word" target="_blank">📄 Word</a>`;
    if(d.excel) h += `<a href="${d.excel}" class="dl-btn dl-excel" target="_blank">📊 Excel</a>`;
    if(d.pdf) h += `<a href="${d.pdf}" class="dl-btn dl-pdf" target="_blank">📕 PDF</a>`;
    return h ? `<div class="dl-group">${h}</div>` : '';
}

async function sendData(){
    const i = document.getElementById('input1');
    const m = i.value.trim(); if(!m && !file1Content && !sheetContent) return;
    addMsg('chat1','user',m||'Phân tích dữ liệu'); i.value='';
    document.getElementById('btn1').disabled=true; document.getElementById('btn1').textContent='⏳';
    const r = await fetch('/api/chat-data', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m, file_content:file1Content, sheet_content:sheetContent})});
    const d = await r.json();
    addMsg('chat1','ai',d.reply, dlLinks(d));
    document.getElementById('btn1').disabled=false; document.getElementById('btn1').textContent='➤';
}
function quickData(t){ document.getElementById('input1').value=t; sendData(); }

async function sendDoc(){
    const i = document.getElementById('input2');
    const m = i.value.trim(); if(!m) return;
    addMsg('chat2','user',m); i.value='';
    document.getElementById('btn2').disabled=true; document.getElementById('btn2').textContent='⏳';
    const r = await fetch('/api/chat-doc', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m})});
    const d = await r.json(); addMsg('chat2','ai',d.reply);
    document.getElementById('btn2').disabled=false; document.getElementById('btn2').textContent='➤';
}
function quickDoc(t){ document.getElementById('input2').value=t; sendDoc(); }

async function sendTender(){
    const i = document.getElementById('input3');
    const m = i.value.trim(); if(!m) return;
    addMsg('chat3','user',m); i.value='';
    document.getElementById('btn3').disabled=true; document.getElementById('btn3').textContent='⏳';
    const r = await fetch('/api/chat-tender', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m})});
    const d = await r.json(); addMsg('chat3','ai',d.reply);
    document.getElementById('btn3').disabled=false; document.getElementById('btn3').textContent='➤';
}
function quickTender(t){ document.getElementById('input3').value=t; sendTender(); }

async function sendEquip(){
    const i = document.getElementById('input4');
    const m = i.value.trim(); if(!m && !file4Content) return;
    addMsg('chat4','user',m||'Phân tích danh sách thiết bị'); i.value='';
    document.getElementById('btn4').disabled=true; document.getElementById('btn4').textContent='⏳';
    const r = await fetch('/api/chat-equip', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({message:m, file_content:file4Content})});
    const d = await r.json(); addMsg('chat4','ai',d.reply);
    document.getElementById('btn4').disabled=false; document.getElementById('btn4').textContent='➤';
}
function quickEquip(t){ document.getElementById('input4').value=t; sendEquip(); }
</script>
</body>
</html>
"""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
