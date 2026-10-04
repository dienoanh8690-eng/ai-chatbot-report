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

# === KHAI BÁO APP ĐẦU TIÊN ===
app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH — ĐỌC TỪ BIẾN MÔI TRƯỜNG ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-3.5-flash"
GEMINI_API_URL = f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"

UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
KNOWLEDGE_FOLDER = "kho_kien_thuc"

for folder in [UPLOAD_FOLDER, RESULT_FOLDER, KNOWLEDGE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

INDEX_FILE = os.path.join(KNOWLEDGE_FOLDER, "danh_sach.json")
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"tai_lieu": [], "ket_qua": []}, f, ensure_ascii=False, indent=2)


# ==================== ĐỌC FILE ====================
def doc_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            noi_dung = "\n".join([p.text for p in doc.paragraphs])
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True, read_only=True)
            ws = wb.active
            for hang in ws.iter_rows(values_only=True):
                noi_dung += " | ".join(str(c) if c else "" for c in hang) + "\n"
            wb.close()
        elif dinh_dang in ["txt", "md"]:
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            noi_dung = "[Nội dung file PDF đã đọc]"
    except Exception as e:
        noi_dung = f"[Lỗi đọc file: {str(e)}]"
    return noi_dung


# ==================== LƯU LỊCH SỬ ====================
def luu_vao_kho(loai, ten_file, mo_ta):
    try:
        with open(INDEX_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        data[loai].append({"ten": ten_file, "mo_ta": mo_ta, "ngay": datetime.now().strftime("%d/%m/%Y %H:%M")})
        with open(INDEX_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ==================== TẠO FILE — TỐI ƯU BỘ NHỚ ====================
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    
    p = doc.add_heading("BÁO CÁO XỬ LÝ DỮ LIỆU", 0)
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    
    p = doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    
    doc.add_paragraph("-" * 60)
    
    dem = 0
    for dong in noi_dung.split("\n"):
        if dong.strip():
            p = doc.add_paragraph()
            run = p.add_run(dong.strip())
            run.font.name = "Arial"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            dem += 1
            if dem > 200:
                p = doc.add_paragraph("... (nội dung đã rút gọn)")
                break
    
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
    dem_dong = 0
    for dong in noi_dung.split("\n"):
        if dong.strip() and not dong.strip().startswith(("#", "---")):
            ws.merge_cells(start_row=hang, start_column=1, end_row=hang, end_column=9)
            ws.cell(row=hang, column=1, value=dong.strip())
            hang += 1
            dem_dong += 1
            if dem_dong > 100:
                ws.merge_cells(start_row=hang, start_column=1, end_row=hang, end_column=9)
                ws.cell(row=hang, column=1, value="... (dữ liệu đã rút gọn)")
                break
    
    for c, w in enumerate([6, 12, 25, 20, 10, 10, 14, 14, 20], 1):
        ws.column_dimensions[chr(64 + c)].width = w
    
    wb.save(duong_dan)
    return ten

def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    
    noi_dung_rut = noi_dung[:5000]
    if len(noi_dung) > 5000:
        noi_dung_rut += "\n\n... (nội dung đã rút gọn để tối ưu hóa)"
    
    html = f"""
    <html><head><meta charset="utf-8"><style>
        body {{ font-family: Arial; padding: 20px; line-height: 1.6; font-size: 12px; }}
        h1 {{ text-align: center; color: #0F4C81; border-bottom: 2px solid #0F4C81; padding-bottom: 10px; font-size: 18px; }}
        .ngay {{ text-align: right; color: #666; margin-bottom: 15px; font-size: 11px; }}
        pre {{ white-space: pre-wrap; word-wrap: break-word; font-size: 11px; }}
    </style></head>
    <body>
        <h1>BÁO CÁO XỬ LÝ DỮ LIỆU</h1>
        <p class="ngay">Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr><pre>{noi_dung_rut}</pre>
    </body></html>"""
    
    with open(duong_dan, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten


# ==================== GỌI AI ====================
def goi_ai(noi_dung, file_content="", he_thong=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY trên Render → vào Environment Variables thêm khóa."
    
    prompt = f"""{he_thong or "Bạn là trợ lý AI hữu ích, trả lời bằng tiếng Việt rõ ràng, dễ hiểu."}

Yêu cầu: {noi_dung}
Nội dung tệp:
{file_content[:3000] if file_content else '(Không có tệp)'}"""
    
    try:
        res = requests.post(GEMINI_API_URL, json={"contents": [{"parts": [{"text": prompt}]}]}, timeout=120)
        
        if res.status_code == 429:
            return "⚠️ API Gemini hết hạn sử dụng (quota). Vui lòng tạo khóa mới tại https://aistudio.google.com/apikey và cập nhật trên Render."
        if res.status_code != 200:
            return f"❌ Lỗi API {res.status_code}: {res.text[:200]}"
        
        data = res.json()
        if "candidates" not in data:
            return f"❌ Không có kết quả: {json.dumps(data, ensure_ascii=False)[:200]}"
        
        return data["candidates"][0]["content"]["parts"][0]["text"]
        
    except requests.exceptions.Timeout:
        return "⏳ Yêu cầu đang xử lý lâu hơn dự kiến, vui lòng thử lại sau."
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# ==================== ROUTE API ====================
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


@app.route("/api/chat-bao-cao", methods=["POST"])
def chat_bao_cao():
    try:
        data = request.get_json(force=True) or {}
    except Exception:
        return jsonify({"reply": "❌ Dữ liệu gửi lên không hợp lệ!", "word":"", "excel":"", "pdf":""})
    
    cau_hoi = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    
    if not cau_hoi and not file_content:
        return jsonify({"reply": "Vui lòng nhập yêu cầu hoặc tải tệp lên!", "word":"", "excel":"", "pdf":""})
    
    tra_loi = goi_ai(cau_hoi, file_content, he_thong="Bạn là chuyên gia xử lý dữ liệu cho nhà máy thủy điện. Trả lời ngắn gọn, rõ ràng, có cấu trúc.")
    
    word = excel = pdf = ""
    if "❌" not in tra_loi and "⚠️" not in tra_loi and "quota" not in tra_loi:
        try:
            word = tao_word(tra_loi)
            excel = tao_excel(tra_loi)
            pdf = tao_pdf(tra_loi)
            luu_vao_kho("ket_qua", word, cau_hoi[:100])
        except Exception as e:
            return jsonify({"reply": f"✅ AI trả lời xong!\nLỗi tạo file: {str(e)}", "word":"", "excel":"", "pdf":""})
    
    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else "",
        "pdf": f"/download/{pdf}" if pdf else ""
    })


@app.route("/api/chat-tu-do", methods=["POST"])
def chat_tu_do():
    try:
        data = request.get_json(force=True) or {}
    except Exception:
        return jsonify({"reply": "❌ Dữ liệu gửi lên không hợp lệ!"})
    
    cau_hoi = data.get("message", "").strip()
    if not cau_hoi:
        return jsonify({"reply": "Vui lòng nhập câu hỏi!"})
    
    tra_loi = goi_ai(cau_hoi, he_thong="Bạn là trợ lý AI thân thiện, trả lời ngắn gọn, dễ hiểu.")
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

        .main-container {
            max-width: 1200px;
            margin: 20px auto;
            padding: 0 24px 40px;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 24px;
        }

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

        @media (max-width: 900px) {
            .main-container { grid-template-columns: 1fr; }
            .card { height: auto; min-height: 500px; }
        }
    </style>
</head>
<body>
    <header class="header">
        <div class="logo">⚡</div>
        <div class="header-text">
            <h1>All Thủy Điện</h1>
            <p>Xử lý dữ liệu thông minh & Tạo báo cáo tự động</p>
        </div>
    </header>

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

    <div class="main-container">
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

        function cleanHtmlTags(text) {
            // Sửa lỗi escape regex — dùng chuỗi thô
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

        function sendQuick(text) {
            document.getElementById('reportInput').value = text;
            sendReport();
        }

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

        async function sendReport() {
            const input = document.getElementById('reportInput');
            const btn = document.getElementById('reportBtn');
            const message = input.value.trim();

            if (!message && !uploadedContent) return;

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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
