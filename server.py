from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from xhtml2pdf import pisa
from flask_cors import CORS

# === KHỞI TẠO APP — ĐÚNG VỊ TRÍ NÀY ===
app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH ====================
HUGGINGFACE_TOKEN = os.environ.get("HUGGINGFACE_TOKEN", "")
ZALO_BOT_TOKEN = os.environ.get("ZALO_BOT_TOKEN", "")
FB_PAGE_TOKEN = os.environ.get("FB_PAGE_TOKEN", "")
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "baocao_ai_2026")
# ===================================================

THU_MUC_FILES = "generated_files"
os.makedirs(THU_MUC_FILES, exist_ok=True)

# -------------------- TẠO FILE --------------------
def tao_word(noi_dung):
    ten_file = f"{THU_MUC_FILES}/bao_cao_{uuid.uuid4().hex[:8]}.docx"
    doc = Document()
    doc.add_heading("BÁO CÁO TỔNG HỢP", 0)
    doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("=" * 40)
    for doan in noi_dung.split("\n"):
        if doan.strip():
            doc.add_paragraph(doan)
    doc.save(ten_file)
    return ten_file

def tao_excel(du_lieu=None):
    ten_file = f"{THU_MUC_FILES}/so_lieu_{uuid.uuid4().hex[:8]}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "DuLieu"
    ws.append(["Chỉ tiêu", "Giá trị"])
    ws.append(["Ngày tạo", datetime.now().strftime("%d/%m/%Y")])
    if du_lieu:
        for k, v in du_lieu.items():
            ws.append([k, v])
    wb.save(ten_file)
    return ten_file

def tao_pdf(noi_dung):
    ten_file = f"{THU_MUC_FILES}/bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    html = f"""
    <html>
        <body style="font-family: DejaVu Sans; padding: 20px;">
            <h1 style="text-align:center; color:#2c3e50;">BÁO CÁO AI TỔNG HỢP</h1>
            <p style="text-align:right;">Ngày: {datetime.now().strftime('%d/%m/%Y')}</p>
            <hr>
            <p>{noi_dung.replace(chr(10), '<br>')}</p>
        </body>
    </html>
    """
    with open(ten_file, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten_file

# -------------------- GỌI AI HUGGING FACE --------------------
def goi_ai(noi_dung_nguoi_dung):
    prompt = f"""Bạn là trợ lý phân tích dữ liệu, tạo báo cáo, soạn văn bản.
Trả lời bằng tiếng Việt, rõ ràng, có cấu trúc.
Yêu cầu: {noi_dung_nguoi_dung}"""
    
    try:
        headers = {
            "Authorization": f"Bearer {HUGGINGFACE_TOKEN}",
            "Content-Type": "application/json"
        }
        data = {
            "inputs": prompt,
            "parameters": {
                "max_new_tokens": 1024,
                "temperature": 0.7
            }
        }
        
        api_url = "https://api-inference.huggingface.co/models/Qwen/Qwen2.5-7B-Instruct"
        resp = requests.post(api_url, headers=headers, json=data, timeout=120)
        result = resp.json()
        
        if isinstance(result, list) and len(result) > 0:
            return result[0].get("generated_text", "Không nhận được phản hồi.")
        elif isinstance(result, dict):
            return result.get("generated_text", str(result))
        return str(result)
        
    except Exception as e:
        return f"Lỗi hệ thống: {str(e)}. Vui lòng thử lại sau ít phút."

# -------------------- GỬI TIN NHẮN --------------------
def gui_zalo(nguoi_dung_id, noi_dung):
    if not ZALO_BOT_TOKEN: return
    try:
        requests.post(
            "https://openapi.zalo.me/v2.0/message/text",
            headers={"access_token": ZALO_BOT_TOKEN},
            json={"user_id": nguoi_dung_id, "text": noi_dung}
        )
    except: pass

def gui_fb(tin_nhan_id, noi_dung):
    if not FB_PAGE_TOKEN: return
    try:
        requests.post(
            f"https://graph.facebook.com/v18.0/me/messages?access_token={FB_PAGE_TOKEN}",
            json={"recipient": {"id": tin_nhan_id}, "message": {"text": noi_dung}}
        )
    except: pass

# -------------------- WEBHOOK ZALO --------------------
@app.route("/webhooks/zalo", methods=["POST"])
def zalo_webhook():
    data = request.json
    if data.get("event_type") == "user_send_text":
        nd = data["message"]["text"]
        ai_tra_loi = goi_ai(nd)
        
        link_file = ""
        if any(tu in nd.lower() for tu in ["báo cáo", "tổng hợp", "file", "word", "excel", "pdf"]):
            fw = tao_word(ai_tra_loi)
            fe = tao_excel()
            fp = tao_pdf(ai_tra_loi)
            link_file = f"""

📄 Tải file:
Word: https://ten-dich-vu.onrender.com/download/{os.path.basename(fw)}
Excel: https://ten-dich-vu.onrender.com/download/{os.path.basename(fe)}
PDF: https://ten-dich-vu.onrender.com/download/{os.path.basename(fp)}"""
        
        gui_zalo(data["sender_id"], ai_tra_loi + link_file)
    return "OK"

# -------------------- WEBHOOK FACEBOOK --------------------
@app.route("/webhooks/messenger", methods=["GET", "POST"])
def fb_webhook():
    if request.method == "GET":
        token_nhan = request.args.get("hub.verify_token", "")
        if token_nhan == FB_VERIFY_TOKEN:
            return request.args.get("hub.challenge", "")
        return "Xác minh thất bại", 403
    
    data = request.json
    for entry in data.get("entry", []):
        for change in entry.get("messaging", []):
            if "message" in change and not change["message"].get("is_echo"):
                nd = change["message"]["text"]
                ai_tra_loi = goi_ai(nd)
                
                link_file = ""
                if any(tu in nd.lower() for tu in ["báo cáo", "tổng hợp", "file", "word", "excel", "pdf"]):
                    fw = tao_word(ai_tra_loi)
                    fe = tao_excel()
                    fp = tao_pdf(ai_tra_loi)
                    link_file = f"""

📄 Tải file:
Word: https://ten-dich-vu.onrender.com/download/{os.path.basename(fw)}
Excel: https://ten-dich-vu.onrender.com/download/{os.path.basename(fe)}
PDF: https://ten-dich-vu.onrender.com/download/{os.path.basename(fp)}"""
                
                gui_fb(change["sender"]["id"], ai_tra_loi + link_file)
    return "OK"

# -------------------- TẢI FILE --------------------
@app.route("/download/<ten_file>")
def tai_file(ten_file):
    duong_dan = os.path.join(THU_MUC_FILES, ten_file)
    if os.path.exists(duong_dan):
        return send_file(duong_dan, as_attachment=True)
    return "File không tồn tại", 404

# -------------------- WEB CHAT CHO WEBSITE --------------------
@app.route("/api/chat", methods=["POST"])
def web_chat():
    data = request.json
    nd = data.get("message", "")
    ai_tra_loi = goi_ai(nd)
    
    file_links = {}
    if any(tu in nd.lower() for tu in ["báo cáo", "tổng hợp", "file", "word", "excel", "pdf"]):
        file_links["word"] = f"/download/{os.path.basename(tao_word(ai_tra_loi))}"
        file_links["excel"] = f"/download/{os.path.basename(tao_excel())}"
        file_links["pdf"] = f"/download/{os.path.basename(tao_pdf(ai_tra_loi))}"
    
    return jsonify({"reply": ai_tra_loi, "files": file_links})

@app.route("/")
def trang_chu():
    return """
    <html>
    <body style="font-family:Arial;max-width:600px;margin:30px auto;padding:0 20px;">
        <h1>🤖 Trợ lý AI Tạo Báo Cáo</h1>
        <p>Nhập yêu cầu: tổng hợp số liệu, soạn văn bản, tạo báo cáo...</p>
        <textarea id="input" rows="4" style="width:100%;padding:10px;margin:10px 0;"></textarea>
        <button onclick="gui()" style="padding:10px 25px;background:#2563eb;color:white;border:none;border-radius:5px;cursor:pointer;">Gửi</button>
        <div id="ketqua" style="margin-top:20px;white-space:pre-wrap;"></div>
        <script>
        async function gui(){
            const val = document.getElementById("input").value;
            document.getElementById("ketqua").innerHTML = "⏳ Đang xử lý...";
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: {"Content-Type":"application/json"},
                    body: JSON.stringify({message: val})
                });
                const d = await res.json();
                let html = d.reply || "Không có phản hồi";
                if(d.files){
                    html += "<br><br>📄 Tải file: ";
                    html += `<a href='${d.files.word}'>Word</a> | `;
                    html += `<a href='${d.files.excel}'>Excel</a> | `;
                    html += `<a href='${d.files.pdf}'>PDF</a>`;
                }
                document.getElementById("ketqua").innerHTML = html;
            } catch(e) {
                document.getElementById("ketqua").innerHTML = "❌ Lỗi kết nối: " + e;
            }
        }
        </script>
    </body>
    </html>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
