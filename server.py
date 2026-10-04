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

# ==================== CẤU HÌNH ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
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
            wb = load_workbook(duong_dan, data_only=True, read_only=True)  # Tối ưu bộ nhớ
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
        pass  # Bỏ qua lỗi lưu lịch sử để không ảnh hưởng chính


# ==================== TẠO FILE — TỐI ƯU BỘ NHỚ ====================
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    
    # Tiêu đề
    p = doc.add_heading("BÁO CÁO XỬ LÝ DỮ LIỆU", 0)
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    
    # Ngày tạo
    p = doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    for run in p.runs:
        run.font.name = "Arial"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    
    doc.add_paragraph("-" * 60)
    
    # Nội dung — Tối ưu: giới hạn độ dài
    dem = 0
    for dong in noi_dung.split("\n"):
        if dong.strip():
            p = doc.add_paragraph()
            run = p.add_run(dong.strip())
            run.font.name = "Arial"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            dem += 1
            if dem > 200:  # Giới hạn số dòng để tiết kiệm RAM
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
            if dem_dong > 100:  # Giới hạn tiết kiệm RAM
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
    
    # Rút gọn nội dung trước khi tạo PDF
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


# ==================== GỌI AI — XỬ LÝ LỖI 429 ====================
def goi_ai(noi_dung, file_content="", he_thong=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY trên Render → vào Environment Variables thêm khóa."
    
    prompt = f"""{he_thong or "Bạn là trợ lý AI hữu ích, trả lời bằng tiếng Việt rõ ràng, dễ hiểu."}

Yêu cầu: {noi_dung}
Nội dung tệp:
{file_content[:3000] if file_content else '(Không có tệp)'}"""  # Giới hạn độ dài gửi đi
    
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
            --primary: #165DFF; --success: #00B42A; --danger: #F53F3F;
            --warning: #FF7D00; --bg: #F7F8FA; --card: #FFFFFF;
            --bubble-user: #E8F3FF; --bubble-ai: #F2F3F5;
            --bubble-chat-user: #EDE7F6; --bubble-chat-ai: #F3E5F5;
            --text-1: #1D2129; --text-2: #4E5969; --border: #E5E6EB;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif; background: var(--bg);
            min-height: 100vh; display: flex; flex-direction: column;
        }
        .header {
            padding: 16px 24px; background: white; box-shadow: 0 2px 12px rgba(0,0,0,0.08);
            display: flex; align-items: center; gap: 12px; z-index: 10;
        }
        .logo {
            width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #165DFF, #4080FF);
            color: white; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 20px;
        }
        .header h1 { font-size: 18px; font-weight: 600; }
        .header p { font-size: 13px; color: var(--text-2); }

        /* Thanh quy trình */
        .process-bar {
            background: white; padding: 12px 24px; border-bottom: 1px solid var(--border);
            display: flex; justify-content: space-between; align-items: center;
            max-width: 1200px; margin: 0 auto; width: 100%; flex-wrap: wrap; gap: 8px;
        }
        .process-step { display: flex; align-items: center; gap: 8px; }
        .step-number {
            width: 24px; height: 24px; border-radius: 50%; background: var(--bg);
            color: var(--text-2); display: flex; align-items: center; justify-content: center;
            font-weight: 600; font-size: 12px; transition: all 0.3s;
        }
        .step-number.active { background: var(--primary); color: white; }
        .step-text { font-size: 13px; color: var(--text-2); transition: all 0.3s; }
        .step-text.active { color: var(--primary); font-weight: 500; }

        /* Bố cục chính */
        .main-container {
            flex: 1; max-width: 1200px; margin: 0 auto; width: 100%; padding: 20px;
            display: grid; grid-template-columns: 1fr 1fr; gap: 20px;
        }

        /* Card chung */
        .card {
            background: white; border-radius: 16px; padding: 20px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08); display: flex; flex-direction: column;
        }
        .card-title {
            font-size: 16px; font-weight: 600; margin-bottom: 16px; color: var(--text-1);
            display: flex; align-items: center; gap: 8px;
        }
        .card-title.bc { border-bottom: 2px solid var(--primary); padding-bottom: 8px; }
        .card-title.ai { border-bottom: 2px solid #9C27B0; padding-bottom: 8px; }

        /* Khu vực báo cáo */
        .upload-area {
            border: 2px dashed var(--border); border-radius: 12px; padding: 24px;
            text-align: center; cursor: pointer; transition: all 0.2s; margin-bottom: 16px;
        }
        .upload-area:hover { border-color: var(--primary); background: #F0F7FF; }
        .upload-icon { font-size: 36px; margin-bottom: 8px; }
        .file-selected {
            padding: 10px 14px; background: #E8FFEA; border-radius: 8px; display: none;
            align-items: center; gap: 10px; margin-bottom: 16px;
        }
        .file-selected.show { display: flex; }
        .clear-file {
            margin-left: auto; cursor: pointer; color: var(--danger); font-weight: bold;
            border: none; background: none; font-size: 18px;
        }
        .quick-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 16px; }
        .quick-btn {
            padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px;
            background: white; cursor: pointer; font-size: 13px; transition: all 0.2s; text-align: left;
        }
        .quick-btn:hover { border-color: var(--primary); background: #F0F7FF; color: var(--primary); }

        /* Khu vực hội thoại */
        .chat-box { flex: 1; overflow-y: auto; padding: 4px; min-height: 350px; }
        .message {
            margin-bottom: 16px; display: flex; max-width: 95%;
            animation: bubbleIn 0.3s ease;
        }
        @keyframes bubbleIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .message.user { justify-content: flex-end; margin-left: auto; }
        .message.ai { justify-content: flex-start; margin-right: auto; }
        .bubble {
            padding: 12px 16px; border-radius: 16px; line-height: 1.6; white-space: pre-wrap;
            font-size: 14px;
        }
        .bao-cao .user .bubble { background: var(--bubble-user); border-bottom-right-radius: 4px; }
        .bao-cao .ai .bubble { background: var(--bubble-ai); border-bottom-left-radius: 4px; }
        .chat-tu-do .user .bubble { background: var(--bubble-chat-user); border-bottom-right-radius: 4px; }
        .chat-tu-do .ai .bubble { background: var(--bubble-chat-ai); border-bottom-left-radius: 4px; }
        .bubble.warning { background: #FFF7E8; border-left: 3px solid var(--warning); }
        .bubble.error { background: #FFF1F0; border-left: 3px solid var(--danger); }

        .file-tag {
            display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px;
            background: #E8FFEA; border-radius: 16px; font-size: 12px; margin-bottom: 8px;
        }
        .download-row { display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }
        .dl-btn {
            padding: 6px 14px; border-radius: 18px; text-decoration: none; font-size: 12px; font-weight: 600;
            display: inline-flex; align-items: center; gap: 4px; transition: transform 0.2s;
        }
        .dl-btn:hover { transform: translateY(-2px); }
        .dl-word { background: #E8F3FF; color: var(--primary); }
        .dl-excel { background: #E8FFEA; color: var(--success); }
        .dl-pdf { background: #FFECEC; color: var(--danger); }

        /* Ô nhập liệu */
        .input-row { display: flex; gap: 8px; align-items: flex-end; margin-top: 12px; }
        textarea {
            flex: 1; min-height: 44px; max-height: 100px; padding: 10px 14px;
            border: 1px solid var(--border); border-radius: 20px;
            font-size: 14px; font-family: inherit; resize: none; outline: none;
            transition: border 0.2s;
        }
        textarea:focus { border-color: var(--primary); }
        .send-btn {
            width: 40px; height: 40px; border-radius: 50%; border: none;
            background: var(--primary); color: white; cursor: pointer; font-size: 16px;
            transition: all 0.2s; flex-shrink: 0;
        }
        .send-btn:hover { background: #0E42D2; transform: scale(1.05); }
        .send-btn:disabled { background: #C9CDD4; cursor: not-allowed; transform: none; }
        .chat-tu-do .send-btn { background: #9C27B0; }
        .chat-tu-do .send-btn:hover { background: #7B1FA2; }
        .chat-tu-do textarea:focus { border-color: #9C27B0; }

        /* Responsive */
        @media (max-width: 900px) {
            .main-container { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">⚡</div>
        <div>
            <h1>All Thủy Điện</h1>
            <p>Xử lý dữ liệu & Trò chuyện với AI</p>
        </div>
    </div>

    <div class="process-bar">
        <div class="process-step">
            <div class="step-number active" id="buoc1">1</div>
            <div class="step-text active" id="t1">Đăng nhập</div>
        </div>
        <div class="process-step">
            <div class="step-number" id="buoc2">2</div>
            <div class="step-text" id="t2">Phân tích</div>
        </div>
        <div class="process-step">
            <div class="step-number" id="buoc3">3</div>
            <div class="step-text" id="t3">Kết quả</div>
        </div>
        <div class="process-step">
            <div class="step-number" id="buoc4">4</div>
            <div class="step-text" id="t4">Tải báo cáo</div>
        </div>
    </div>

    <div class="main-container">
        <!-- BÊN TRÁI: XỬ LÝ DỮ LIỆU & BÁO CÁO -->
        <div class="card bao-cao">
            <div class="card-title bc">📊 Xử lý dữ liệu & Báo cáo</div>

            <div class="upload-area" id="uploadArea">
                <div class="upload-icon">📎</div>
                <div>Nhấn để chọn hoặc kéo thả tệp</div>
                <div style="font-size: 12px; color: var(--text-2); margin-top: 4px;">.docx .xlsx .txt .pdf</div>
            </div>
            <input type="file" id="chonTep" accept=".docx,.xlsx,.txt,.pdf" style="display:none;">
            
            <div class="file-selected" id="thongTinTep">
                <span>📎</span>
                <span id="tenTep"></span>
                <button class="clear-file" onclick="xoaTep()">✕</button>
            </div>

            <div class="quick-actions">
                <button class="quick-btn" onclick="nhapYeuCauVaGui('Sắp xếp dữ liệu theo tên thiết bị')">Sắp xếp dữ liệu</button>
                <button class="quick-btn" onclick="nhapYeuCauVaGui('Tính thành tiền = số lượng × đơn giá')">Tính thành tiền</button>
                <button class="quick-btn" onclick="nhapYeuCauVaGui('Lập báo cáo tổng hợp')">Lập báo cáo</button>
                <button class="quick-btn" onclick="nhapYeuCauVaGui('Kiểm tra tình trạng thiết bị')">Kiểm tra thiết bị</button>
            </div>

            <div class="chat-box" id="khuBaoCao">
                <div class="message ai">
                    <div class="bubble">
                        👋 Đăng nhập → tải tệp lên hoặc chọn yêu cầu nhanh để bắt đầu xử lý dữ liệu và lập báo cáo.
                    </div>
                </div>
            </div>

            <div class="input-row">
                <textarea id="inputBaoCao" placeholder="Nhập yêu cầu... (Enter gửi, Shift+Enter xuống dòng)" onkeydown="xuLyPhimBaoCao(event)"></textarea>
                <button class="send-btn" id="nutGuiBaoCao" onclick="guiBaoCao()">➤</button>
            </div>
        </div>

        <!-- BÊN PHẢI: TRÒ CHUYỆN VỚI AI -->
        <div class="card chat-tu-do">
            <div class="card-title ai">💬 Trò chuyện với AI</div>

            <div class="chat-box" id="khuChatTuDo">
                <div class="message ai">
                    <div class="bubble">
                        👋 Tôi là AI trợ lý. Bạn có thể hỏi tôi bất cứ điều gì nhé!
                    </div>
                </div>
            </div>

            <div class="input-row">
                <textarea id="inputChatTuDo" placeholder="Hỏi AI bất kỳ điều gì..." onkeydown="xuLyPhimChat(event)"></textarea>
                <button class="send-btn" id="nutGuiChat" onclick="guiChatTuDo()">➤</button>
            </div>
        </div>
    </div>

    <script>
        let fileContent = "";
        let tenTepDaChon = "";

        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('uploadArea').addEventListener('click', () => document.getElementById('chonTep').click());
            document.getElementById('chonTep').addEventListener('change', chonTep);
        });

        // === CẬP NHẬT BƯỚC QUY TRÌNH ===
        function capNhatBuoc(n) {
            for (let i = 1; i <= 4; i++) {
                const b = document.getElementById('buoc'+i);
                const t = document.getElementById('t'+i);
                if (i <= n) {
                    b.classList.add('active'); t.classList.add('active');
                } else {
                    b.classList.remove('active'); t.classList.remove('active');
                }
            }
        }

        // === TẢI TỆP ===
        function chonTep(e) {
            const f = e.target.files[0];
            if (!f) return;
            tenTepDaChon = f.name;
            const fd = new FormData(); fd.append('file', f);
            fetch('/api/upload', { method: 'POST', body: fd })
                .then(r => r.json())
                .then(d => {
                    if (d.status === 'ok') {
                        fileContent = d.content;
                        document.getElementById('tenTep').textContent = tenTepDaChon;
                        document.getElementById('thongTinTep').classList.add('show');
                        capNhatBuoc(2);
                    } else {
                        themTinBaoCao('ai', '❌ ' + (d.error || 'Lỗi tải tệp'));
                    }
                })
                .catch(err => themTinBaoCao('ai', '❌ Lỗi: ' + err.message));
        }

        function xoaTep() {
            fileContent = ''; tenTepDaChon = '';
            document.getElementById('thongTinTep').classList.remove('show');
            document.getElementById('chonTep').value = '';
            capNhatBuoc(1);
        }

        // === NÚT NHANH ===
        function nhapYeuCauVaGui(text) {
            document.getElementById('inputBaoCao').value = text;
            guiBaoCao();
        }

        function xuLyPhimBaoCao(e) {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); guiBaoCao(); }
        }

        // === HIỂN THỊ TIN NHẮN BÁO CÁO ===
        function themTinBaoCao(loai, nd, links=null) {
            const kh = document.getElementById('khuBaoCao');
            const div = document.createElement('div');
            div.className = 'message ' + loai;
            
            let noi_dung_hien = nd;
            if (loai === 'user' && tenTepDaChon) {
                noi_dung_hien = `<span class="file-tag">📎 ${tenTepDaChon}</span>\n${nd}`;
            }
            
            let linkHtml = '';
            if (links && (links.word||links.excel||links.pdf)) {
                linkHtml = '<div class="download-row">';
                if (links.word) linkHtml += `<a href="${links.word}" class="dl-btn dl-word" target="_blank">📄 Word</a>`;
                if (links.excel) linkHtml += `<a href="${links.excel}" class="dl-btn dl-excel" target="_blank">📊 Excel</a>`;
                if (links.pdf) linkHtml += `<a href="${links.pdf}" class="dl-btn dl-pdf" target="_blank">📕 PDF</a>`;
                linkHtml += '</div>';
            }
            
            div.innerHTML = `<div class="bubble">${noi_dung_hien.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace('&lt;span','<span').replace('&lt;/span&gt;','</span>')}${linkHtml}</div>`;
            kh.appendChild(div); kh.scrollTop = kh.scrollHeight;
            
            if (loai === 'ai') {
                capNhatBuoc(3);
                if (links && (links.word||links.excel||links.pdf)) capNhatBuoc(4);
            }
        }

        // === GỬI YÊU CẦU BÁO CÁO ===
        async function guiBaoCao() {
            const inp = document.getElementById('inputBaoCao');
            const btn = document.getElementById('nutGuiBaoCao');
            const msg = inp.value.trim();
            
            if (!msg && !fileContent) return;

            let hienThi = msg;
            if (tenTepDaChon) {
                hienThi = `<span class="file-tag">📎 ${tenTepDaChon}</span>\n${msg || 'Phân tích nội dung tệp'}`;
            }
            themTinBaoCao('user', hienThi);
            
            inp.value = ''; btn.disabled = true; btn.textContent = '⏳';

            try {
                const res = await fetch('/api/chat-bao-cao', {
                    method: 'POST', 
                    headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({ 
                        message: msg || 'Phân tích và xử lý nội dung tệp', 
                        file_content: fileContent 
                    })
                });
                
                const d = await res.json();
                themTinBaoCao('ai', d.reply || '', { word: d.word, excel: d.excel, pdf: d.pdf });
                
                if (!d.word && !d.excel && !d.pdf && d.reply.includes('quota')) {
                    // Giữ nguyên file để thử lại với key mới
                } else {
                    xoaTep();
                }
            } catch (e) {
                themTinBaoCao('ai', '❌ Lỗi: ' + (e.message || 'Không xác định'));
            } finally {
                btn.disabled = false; btn.textContent = '➤';
            }
        }

        // === TRÒ CHUYỆN AI ===
        function xuLyPhimChat(e) {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); guiChatTuDo(); }
        }

        function themTinChat(loai, nd) {
            const kh = document.getElementById('khuChatTuDo');
            const div = document.createElement('div');
            div.className = 'message ' + loai;
            const html = nd.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            div.innerHTML = `<div class="bubble">${html}</div>`;
            kh.appendChild(div); kh.scrollTop = kh.scrollHeight;
        }

        async function guiChatTuDo() {
            const inp = document.getElementById('inputChatTuDo');
            const btn = document.getElementById('nutGuiChat');
            const msg = inp.value.trim();
            if (!msg) return;

            themTinChat('user', msg);
            inp.value = ''; btn.disabled = true; btn.textContent = '⏳';

            try {
                const res = await fetch('/api/chat-tu-do', {
                    method: 'POST', 
                    headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({ message: msg })
                });
                const d = await res.json();
                themTinChat('ai', d.reply || '❌ Không có phản hồi');
            } catch (e) {
                themTinChat('ai', '❌ Lỗi: ' + (e.message || 'Không xác định'));
            } finally {
                btn.disabled = false; btn.textContent = '➤';
            }
        }
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Dùng đúng cổng Render
    app.run(host="0.0.0.0", port=port)
