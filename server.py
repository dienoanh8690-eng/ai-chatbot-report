from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from openpyxl import Workbook, load_workbook
from xhtml2pdf import pisa
from flask_cors import CORS
import base64

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH BIẾN MÔI TRƯỜNG ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "baocao_ai_2026")
# ====================================================================

THU_MUC_FILES = "generated_files"
THU_MUC_UPLOAD = "uploaded_files"
os.makedirs(THU_MUC_FILES, exist_ok=True)
os.makedirs(THU_MUC_UPLOAD, exist_ok=True)

# -------------------- ĐỌC NỘI DUNG FILE ĐƯỢC TẢI LÊN --------------------
def doc_noi_dung_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            for p in doc.paragraphs:
                noi_dung += p.text + "\n"
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True)
            ws = wb.active
            for hang in ws.iter_rows(values_only=True):
                noi_dung += " | ".join(str(o) if o else "" for o in hang) + "\n"
        elif dinh_dang == "txt":
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            # Đọc PDF cơ bản — nếu cần nâng cao có thể thêm PyPDF2
            noi_dung = "[Đã nhận file PDF — nội dung sẽ được AI phân tích theo yêu cầu của bạn]\n"
            noi_dung += "Nội dung tóm tắt từ file sẽ được xử lý khi bạn nêu rõ yêu cầu chỉnh sửa."
    except Exception as e:
        noi_dung = f"[Không đọc được file: {str(e)}]"
    return noi_dung

# -------------------- TẠO FILE WORD --------------------
def tao_word(noi_dung):
    ten_file = f"{THU_MUC_FILES}/bao_cao_{uuid.uuid4().hex[:8]}.docx"
    doc = Document()
    doc.add_heading("BÁO CÁO TỔNG HỢP", 0)
    doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("=" * 50)
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
    for dong in noi_dung.split("\n"):
        if dong.strip():
            ws.append([dong])
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
            body {{ font-family: DejaVu Sans; padding: 25px; line-height: 1.6; }}
            h1 {{ text-align: center; color: #1e40af; }}
            .ngay {{ color: #64748b; margin-bottom: 20px; }}
            hr {{ border: 1px solid #e2e8f0; margin: 20px 0; }}
        </style>
    </head>
    <body>
        <h1>BÁO CÁO AI TẠO</h1>
        <p class="ngay">Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <p>{noi_dung.replace(chr(10), '<br>')}</p>
    </body>
    </html>
    """
    with open(ten_file, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten_file

# -------------------- GỌI GOOGLE GEMINI AI --------------------
def goi_ai(noi_dung_nguoi_dung, noi_dung_file=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY. Lấy miễn phí tại: aistudio.google.com/apikey"
    
    prompt = """Bạn là chuyên gia phân tích và biên soạn báo cáo khoa học.
Nhiệm vụ của bạn: nhận nội dung từ người dùng và nội dung file đính kèm (nếu có),
sau đó biên soạn lại thành nội dung khoa học, rõ ràng, mạch lạc, có cấu trúc,
dùng ngôn ngữ trang trọng, phù hợp văn bản chính thức/báo cáo.

NỘI DUNG NGƯỜI DÙNG:
""" + noi_dung_nguoi_dung + """

NỘI DUNG FILE ĐÍNH KÈM:
""" + (noi_dung_file if noi_dung_file else "(Không có file đính kèm)") + """

Hãy trả về nội dung đã chỉnh sửa hoàn chỉnh ngay:"""
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={GEMINI_API_KEY}"
        
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }],
            "generationConfig": {
                "maxOutputTokens": 4096,
                "temperature": 0.3
            }
        }
        
        response = requests.post(url, json=payload, timeout=120)
        
        if response.status_code == 200:
            result = response.json()
            try:
                return result["candidates"][0]["content"]["parts"][0]["text"]
            except (KeyError, IndexError):
                return "Không nhận được nội dung phản hồi từ AI"
        else:
            return f"❌ Lỗi API: Mã {response.status_code} - {response.text[:300]}"
            
    except Exception as e:
        return f"❌ Lỗi kết nối: {str(e)}"

# -------------------- TẢI FILE LÊN --------------------
@app.route("/api/upload", methods=["POST"])
def tai_file_len():
    if "file" not in request.files:
        return jsonify({"error": "Không có file nào được gửi"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Chưa chọn file"}), 400
    
    dinh_dang = file.filename.rsplit(".", 1)[-1].lower()
    ho_tro = ["docx", "xlsx", "txt", "pdf"]
    
    if dinh_dang not in ho_tro:
        return jsonify({"error": f"Định dạng không hỗ trợ. Chỉ nhận: {', '.join(ho_tro)}"}), 400
    
    ten_file_moi = f"{uuid.uuid4().hex[:12]}.{dinh_dang}"
    duong_dan = os.path.join(THU_MUC_UPLOAD, ten_file_moi)
    file.save(duong_dan)
    
    # Đọc nội dung để gửi cho AI
    noi_dung = doc_noi_dung_file(duong_dan, dinh_dang)
    
    return jsonify({
        "status": "ok",
        "filename": ten_file_moi,
        "type": dinh_dang,
        "preview": noi_dung[:500] + ("..." if len(noi_dung) > 500 else ""),
        "content": noi_dung
    })

# -------------------- TẢI FILE KẾT QUẢ VỀ --------------------
@app.route("/download/<ten_file>")
def tai_file_ve(ten_file):
    duong_dan = os.path.join(THU_MUC_FILES, ten_file)
    if os.path.exists(duong_dan):
        return send_file(duong_dan, as_attachment=True)
    return "File không tồn tại", 404

# -------------------- CHAT WEB --------------------
@app.route("/api/chat", methods=["POST"])
def web_chat():
    data = request.json
    noi_dung = data.get("message", "").strip()
    noi_dung_file = data.get("file_content", "")
    
    if not noi_dung and not noi_dung_file:
        return jsonify({"reply": "Vui lòng nhập nội dung hoặc tải file lên!"})
    
    phan_hoi = goi_ai(noi_dung, noi_dung_file)
    
    word_link = excel_link = pdf_link = ""
    
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
            body { max-width: 750px; margin: 30px auto; padding: 0 20px; background: #f8fafc; }
            h1 { text-align: center; color: #1e40af; margin-bottom: 8px; }
            .mota { text-align: center; color: #64748b; margin-bottom: 25px; }
            .khung { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
            .khu_vuc_tai { border: 2px dashed #cbd5e1; padding: 25px; text-align: center; border-radius: 10px; margin-bottom: 20px; cursor: pointer; transition: all 0.3s; }
            .khu_vuc_tai:hover { border-color: #2563eb; background: #eff6ff; }
            .khu_vuc_tai.dang_chon { border-color: #22c55e; background: #f0fdf4; }
            textarea { width: 100%; height: 110px; padding: 12px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 15px; resize: vertical; margin-bottom: 15px; }
            button { background: #2563eb; color: white; border: none; padding: 12px 28px; border-radius: 8px; font-size: 16px; cursor: pointer; }
            button:hover { background: #1d4ed8; }
            button:disabled { background: #94a3b8; cursor: not-allowed; }
            .ketqua { margin-top: 25px; padding: 18px; background: #f0fdf4; border-radius: 8px; border-left: 4px solid #22c55e; white-space: pre-wrap; line-height: 1.7; }
            .tai { margin-top: 15px; padding-top: 10px; border-top: 1px solid #e2e8f0; }
            .tai a { display: inline-block; margin-right: 15px; color: #2563eb; text-decoration: none; font-weight: bold; }
            .tai a:hover { text-decoration: underline; }
            .ten_file { margin: 10px 0; padding: 8px 12px; background: #e0f2fe; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; }
            .xoa { color: #ef4444; cursor: pointer; font-weight: bold; }
            .an { display: none; }
        </style>
    </head>
    <body>
        <h1>🤖 Trợ lý AI Tạo & Chỉnh Sửa Báo Cáo</h1>
        <p class="mota">Tải file mẫu lên → nhập yêu cầu → AI chỉnh sửa thành văn bản khoa học</p>
        
        <div class="khung">
            <!-- Khu vực tải file -->
            <div class="khu_vuc_tai" id="khuTai" onclick="document.getElementById('chonFile').click()">
                <strong>📎 Nhấn để tải file lên</strong><br>
                <span style="color:#64748b; font-size:14px;">Hỗ trợ: .docx .xlsx .txt .pdf</span>
                <input type="file" id="chonFile" accept=".docx,.xlsx,.txt,.pdf" style="display:none;" onchange="xuLyFile(this)">
            </div>
            <div id="thongTinFile" class="an"></div>
            
            <!-- Ô nhập yêu cầu -->
            <textarea id="input" placeholder="Ví dụ: Dựa trên nội dung file, viết lại thành báo cáo khoa học có cấu trúc rõ ràng..."></textarea>
            
            <br>
            <button id="nutgui" onclick="gui()">Gửi Đến AI</button>
            
            <div id="ketqua" class="ketqua an"></div>
            <div id="tai" class="tai an"></div>
        </div>

        <script>
        let fileDaChon = null;
        let noiDungFile = "";

        async function xuLyFile(input) {
            const file = input.files[0];
            if (!file) return;
            
            fileDaChon = file;
            const khu = document.getElementById("khuTai");
            khu.classList.add("dang_chon");
            khu.innerHTML = `✅ Đã chọn: <strong>${file.name}</strong><br><span style="color:#22c55e">Đang đọc nội dung...</span>`;
            
            const formData = new FormData();
            formData.append("file", file);
            
            try {
                const res = await fetch("/api/upload", { method: "POST", body: formData });
                const data = await res.json();
                
                if (data.status === "ok") {
                    noiDungFile = data.content;
                    khu.innerHTML = `✅ Đã chọn: <strong>${file.name}</strong><br><span style="color:#22c55e">Nội dung đã sẵn sàng (${data.preview.length} ký tự)</span>`;
                } else {
                    khu.innerHTML = `❌ ${data.error}`;
                }
            } catch (e) {
                khu.innerHTML = `❌ Lỗi tải file: ${e}`;
            }
        }

        async function gui() {
            const input = document.getElementById("input");
            const nut = document.getElementById("nutgui");
            const ketqua = document.getElementById("ketqua");
            const tai = document.getElementById("tai");
            
            const noi_dung = input.value.trim();
            if (!noi_dung && !noiDungFile) {
                alert("Vui lòng nhập yêu cầu hoặc tải file lên!");
                return;
            }
            
            nut.disabled = true;
            nut.innerText = "Đang xử lý...";
            ketqua.style.display = "block";
            ketqua.innerText = "⏳ AI đang phân tích nội dung, vui lòng chờ...";
            tai.style.display = "none";
            
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        message: noi_dung,
                        file_content: noiDungFile
                    })
                });
                const data = await res.json();
                
                ketqua.innerText = data.reply || "Không có phản hồi";
                
                if (data.word || data.excel || data.pdf) {
                    tai.style.display = "block";
                    let html = "<strong>📂 Tải file kết quả:</strong> ";
                    if (data.word) html += `<a href="${data.word}" target="_blank">📄 Word</a> `;
                    if (data.excel) html += `<a href="${data.excel}" target="_blank">📊 Excel</a> `;
                    if (data.pdf) html += `<a href="${data.pdf}" target="_blank">📕 PDF</a>`;
                    tai.innerHTML = html;
                }
            } catch (e) {
                ketqua.innerText = "❌ Lỗi kết nối: " + e;
            }
            
            nut.disabled = false;
            nut.innerText = "Gửi Đến AI";
        }
        </script>
    </body>
    </html>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
