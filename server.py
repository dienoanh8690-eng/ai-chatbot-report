from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH BIẾN MÔI TRƯỜNG ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
ZALO_BOT_TOKEN = os.environ.get("ZALO_BOT_TOKEN", "")
FB_PAGE_TOKEN = os.environ.get("FB_PAGE_TOKEN", "")
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "baocao_ai_2026")
# ====================================================================

THU_MUC_FILES = "generated_files"
os.makedirs(THU_MUC_FILES, exist_ok=True)


# -------------------- TẠO FILE WORD --------------------
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


# -------------------- TẠO FILE EXCEL --------------------
def tao_excel(noi_dung=""):
    ten_file = f"{THU_MUC_FILES}/bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo cáo"
    ws.append(["Nội dung báo cáo"])
    ws.append([noi_dung])
    ws.append(["Ngày tạo", datetime.now().strftime("%d/%m/%Y %H:%M")])
    wb.save(ten_file)
    return ten_file


# -------------------- TẠO FILE PDF --------------------
def tao_pdf(noi_dung):
    ten_file = f"{THU_MUC_FILES}/bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: DejaVu Sans; padding: 20px; }}
            h1 {{ text-align: center; color: #2563eb; }}
        </style>
    </head>
    <body>
        <h1>BÁO CÁO AI TẠO</h1>
        <p><strong>Ngày:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <p>{noi_dung.replace(chr(10), '<br>')}</p>
    </body>
    </html>
    """
    with open(ten_file, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten_file


# -------------------- GỌI GOOGLE GEMINI AI --------------------
def goi_ai(noi_dung_nguoi_dung):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY. Lấy miễn phí tại: aistudio.google.com/apikey"
    
    prompt = f"""Bạn là trợ lý AI thông minh, giúp người dùng tạo báo cáo, tóm tắt thông tin, soạn văn bản, trả lời câu hỏi.
Hãy trả lời bằng tiếng Việt rõ ràng, mạch lạc, dễ hiểu, có cấu trúc phù hợp.

Câu hỏi / Yêu cầu: {noi_dung_nguoi_dung}"""
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={GEMINI_API_KEY}"
        
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }],
            "generationConfig": {
                "maxOutputTokens": 1024,
                "temperature": 0.7
            }
        }
        
        response = requests.post(url, json=payload, timeout=60)
        
        if response.status_code == 200:
            result = response.json()
            try:
                return result["candidates"][0]["content"]["parts"][0]["text"]
            except (KeyError, IndexError):
                return "Không nhận được nội dung phản hồi từ AI"
        else:
            return f"❌ Lỗi API: Mã {response.status_code} - {response.text[:200]}"
            
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"


# -------------------- WEBHOOK ZALO --------------------
@app.route("/webhooks/zalo", methods=["POST"])
def zalo_webhook():
    data = request.json
    return jsonify({"status": "ok"})


# -------------------- WEBHOOK FACEBOOK --------------------
@app.route("/webhooks/messenger", methods=["GET", "POST"])
def fb_webhook():
    if request.method == "GET":
        token_nhan = request.args.get("hub.verify_token", "")
        if token_nhan == FB_VERIFY_TOKEN:
            return request.args.get("hub.challenge", "")
        return "Xác minh thất bại", 403
    return jsonify({"status": "ok"})


# -------------------- TẢI FILE --------------------
@app.route("/download/<ten_file>")
def tai_file(ten_file):
    duong_dan = os.path.join(THU_MUC_FILES, ten_file)
    if os.path.exists(duong_dan):
        return send_file(duong_dan, as_attachment=True)
    return "File không tồn tại", 404


# -------------------- CHAT WEB --------------------
@app.route("/api/chat", methods=["POST"])
def web_chat():
    data = request.json
    noi_dung = data.get("message", "").strip()
    if not noi_dung:
        return jsonify({"reply": "Vui lòng nhập nội dung câu hỏi!"})
    
    phan_hoi = goi_ai(noi_dung)
    
    word_link = ""
    excel_link = ""
    pdf_link = ""
    
    if "❌" not in phan_hoi and "⚠️" not in phan_hoi:
        word_file = tao_word(phan_hoi)
        excel_file = tao_excel(phan_hoi)
        pdf_file = tao_pdf(phan_hoi)
        word_link = f"/download/{os.path.basename(word_file)}"
        excel_link = f"/download/{os.path.basename(excel_file)}"
        pdf_link = f"/download/{os.path.basename(pdf_file)}"
    
    return jsonify({
        "reply": phan_hoi,
        "word": word_link,
        "excel": excel_link,
        "pdf": pdf_link
    })


# -------------------- TRANG CHỦ --------------------
@app.route("/")
def trang_chu():
    return """
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Trợ lý AI Tạo Báo Cáo</title>
        <style>
            * { box-sizing: border-box; margin: 0; padding: 0; font-family: Arial, sans-serif; }
            body { max-width: 700px; margin: 40px auto; padding: 0 20px; background: #f8fafc; }
            h1 { text-align: center; color: #1e40af; margin-bottom: 10px; }
            .mota { text-align: center; color: #64748b; margin-bottom: 30px; }
            .khung { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
            textarea { width: 100%; height: 100px; padding: 12px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 15px; resize: vertical; margin-bottom: 15px; }
            button { background: #2563eb; color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; }
            button:hover { background: #1d4ed8; }
            button:disabled { background: #94a3b8; cursor: not-allowed; }
            .ketqua { margin-top: 25px; padding: 15px; background: #f0fdf4; border-radius: 8px; border-left: 4px solid #22c55e; white-space: pre-wrap; }
            .tai { margin-top: 15px; }
            .tai a { display: inline-block; margin-right: 15px; color: #2563eb; text-decoration: none; font-weight: bold; }
            .tai a:hover { text-decoration: underline; }
        </style>
    </head>
    <body>
        <h1>🤖 Trợ lý AI Tạo Báo Cáo</h1>
        <p class="mota">Nhập yêu cầu: tổng hợp số liệu, soạn văn bản, tạo báo cáo...</p>
        <div class="khung">
            <textarea id="input" placeholder="Ví dụ: Tạo báo cáo tổng kết tháng 10..."></textarea>
            <br>
            <button id="nutgui" onclick="gui()">Gửi</button>
            <div id="ketqua" class="ketqua" style="display:none;"></div>
            <div id="tai" class="tai" style="display:none;"></div>
        </div>

        <script>
        async function gui() {
            const input = document.getElementById("input");
            const nut = document.getElementById("nutgui");
            const ketqua = document.getElementById("ketqua");
            const tai = document.getElementById("tai");
            
            const noi_dung = input.value.trim();
            if (!noi_dung) return;
            
            nut.disabled = true;
            nut.innerText = "Đang xử lý...";
            ketqua.style.display = "block";
            ketqua.innerText = "⏳ Đang gửi đến AI, vui lòng chờ...";
            tai.style.display = "none";
            
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: noi_dung })
                });
                const data = await res.json();
                
                ketqua.innerText = data.reply || "Không có phản hồi";
                
                if (data.word || data.excel || data.pdf) {
                    tai.style.display = "block";
                    let html = "<strong>Tải file:</strong> ";
                    if (data.word) html += `<a href="${data.word}" target="_blank">Word</a> `;
                    if (data.excel) html += `<a href="${data.excel}" target="_blank">Excel</a> `;
                    if (data.pdf) html += `<a href="${data.pdf}" target="_blank">PDF</a>`;
                    tai.innerHTML = html;
                }
            } catch (e) {
                ketqua.innerText = "❌ Lỗi kết nối: " + e;
            }
            
            nut.disabled = false;
            nut.innerText = "Gửi";
        }
        </script>
    </body>
    </html>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
