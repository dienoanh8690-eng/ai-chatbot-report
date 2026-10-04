from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
import json
from datetime import datetime
from docx import Document
from openpyxl import Workbook, load_workbook
from xhtml2pdf import pisa
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ==================== CẤU HÌNH ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# === THƯ MỤC LƯU DỮ LIỆU ===
UPLOAD_FOLDER = "tai_lieu_tai_len"
RESULT_FOLDER = "ket_qua_xuat_ra"
KNOWLEDGE_FOLDER = "kho_kien_thuc"

for folder in [UPLOAD_FOLDER, RESULT_FOLDER, KNOWLEDGE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

INDEX_FILE = os.path.join(KNOWLEDGE_FOLDER, "danh_sach.json")
if not os.path.exists(INDEX_FILE):
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump({"tai_lieu": [], "ket_qua": []}, f, ensure_ascii=False, indent=2)


# -------------------- ĐỌC NỘI DUNG FILE --------------------
def doc_file(duong_dan, dinh_dang):
    noi_dung = ""
    try:
        if dinh_dang == "docx":
            doc = Document(duong_dan)
            noi_dung = "\n".join([p.text for p in doc.paragraphs])
        elif dinh_dang == "xlsx":
            wb = load_workbook(duong_dan, data_only=True)
            ws = wb.active
            noi_dung = "\n".join([" | ".join(str(c) if c else "" for c in row) for row in ws.iter_rows(values_only=True)])
        elif dinh_dang in ["txt", "md"]:
            with open(duong_dan, "r", encoding="utf-8", errors="ignore") as f:
                noi_dung = f.read()
        elif dinh_dang == "pdf":
            noi_dung = "[File PDF - nội dung đã lưu tham khảo]"
    except Exception as e:
        noi_dung = f"[Lỗi đọc: {str(e)}]"
    return noi_dung


# -------------------- LƯU VÀO KHO KIẾN THỨC --------------------
def luu_vao_kho(loai, ten_file, mo_ta):
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    data[loai].append({
        "ten": ten_file,
        "mo_ta": mo_ta,
        "ngay": datetime.now().strftime("%d/%m/%Y %H:%M")
    })
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -------------------- TẠO FILE WORD --------------------
def tao_word(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    doc = Document()
    doc.add_heading("BÁO CÁO", 0)
    doc.add_paragraph(f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph("-" * 50)
    for dong in noi_dung.split("\n"):
        if dong.strip():
            doc.add_paragraph(dong)
    doc.save(duong_dan)
    return ten


# -------------------- TẠO FILE EXCEL --------------------
def tao_excel(noi_dung=""):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    wb = Workbook()
    ws = wb.active
    ws.append(["Nội dung báo cáo"])
    for dong in noi_dung.split("\n"):
        if dong.strip():
            ws.append([dong])
    ws.append(["Ngày tạo", datetime.now().strftime("%d/%m/%Y %H:%M")])
    wb.save(duong_dan)
    return ten


# -------------------- TẠO FILE PDF --------------------
def tao_pdf(noi_dung):
    ten = f"bao_cao_{uuid.uuid4().hex[:8]}.pdf"
    duong_dan = os.path.join(RESULT_FOLDER, ten)
    html = f"""
    <html><head><meta charset="utf-8"><style>
        body {{ font-family: Arial; padding: 20px; }}
        h1 {{ text-align: center; color: #0f4c81; }}
    </style></head>
    <body>
        <h1>BÁO CÁO</h1>
        <p>Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <p>{noi_dung.replace(chr(10), '<br>')}</p>
    </body></html>
    """
    with open(duong_dan, "wb") as f:
        pisa.CreatePDF(html, dest=f)
    return ten


# -------------------- GỌI AI --------------------
def goi_ai(noi_dung, file_content=""):
    if not GEMINI_API_KEY:
        return "⚠️ Chưa đặt GEMINI_API_KEY"
    
    prompt = f"""Bạn là trợ lý chuyên về thủy điện. Trả lời rõ ràng, dễ hiểu, chính xác.

NỘI DUNG YÊU CẦU:
{noi_dung}

NỘI DUNG FILE ĐÍNH KÈM:
{file_content if file_content else '(Không có file)'}
"""
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": 2048, "temperature": 0.3}
        }
        res = requests.post(url, json=payload, timeout=120)
        if res.status_code == 200:
            return res.json()["candidates"][0]["content"]["parts"][0]["text"]
        return f"❌ Lỗi API {res.status_code}"
    except Exception as e:
        return f"❌ Lỗi: {str(e)}"


# ==================== ROUTE ====================

# Tải file lên
@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Không có file"}), 400
    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "Chưa chọn file"}), 400
    
    ext = f.filename.rsplit(".", 1)[-1].lower()
    if ext not in ["docx", "xlsx", "txt", "pdf"]:
        return jsonify({"error": "Chỉ hỗ trợ: .docx .xlsx .txt .pdf"}), 400
    
    ten_moi = f"{uuid.uuid4().hex[:10]}.{ext}"
    duong_dan = os.path.join(UPLOAD_FOLDER, ten_moi)
    f.save(duong_dan)
    
    noi_dung = doc_file(duong_dan, ext)
    luu_vao_kho("tai_lieu", ten_moi, f.filename)
    
    return jsonify({"status": "ok", "name": f.filename, "content": noi_dung[:2000]})


# Chat
@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json
    cau_hoi = data.get("message", "").strip()
    file_content = data.get("file_content", "")
    
    if not cau_hoi and not file_content:
        return jsonify({"reply": "Vui lòng nhập nội dung hoặc tải file lên!"})
    
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


# Tải file về
@app.route("/download/<ten_file>")
def download(ten_file):
    for folder in [RESULT_FOLDER, UPLOAD_FOLDER]:
        path = os.path.join(folder, ten_file)
        if os.path.exists(path):
            return send_file(path, as_attachment=True)
    return "Không tìm thấy file", 404


# Trang chủ
@app.route("/")
def trang_chu():
    return """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>All thủy điện</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
        body { max-width: 700px; margin: 30px auto; padding: 0 20px; background: #f0f7ff; }
        h1 { text-align: center; color: #0f4c81; margin-bottom: 30px; }
        .box { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        
        /* Khu tải file */
        .upload-area { border: 2px dashed #94b8d9; padding: 30px; text-align: center; border-radius: 10px; cursor: pointer; margin-bottom: 20px; }
        .upload-area:hover { border-color: #0f4c81; background: #e6f2ff; }
        .upload-area.active { border-color: #22c55e; background: #f0fdf4; }
        .file-info { margin: 10px 0; padding: 10px; background: #e6ffed; border-radius: 6px; display: none; }
        
        textarea { width: 100%; height: 100px; padding: 12px; border: 1px solid #b3d1e8; border-radius: 8px; font-size: 15px; margin-bottom: 15px; }
        button { background: #0f4c81; color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; width: 100%; }
        button:hover { background: #0d3c68; }
        button:disabled { background: #94b8d9; cursor: not-allowed; }
        
        .result { margin-top: 20px; padding: 15px; background: #f0f9ff; border-radius: 8px; border-left: 4px solid #0f4c81; white-space: pre-wrap; line-height: 1.6; display: none; }
        .download { margin-top: 15px; padding-top: 15px; border-top: 1px solid #cce0f0; display: none; }
        .download a { display: inline-block; margin-right: 15px; color: #0f4c81; font-weight: bold; text-decoration: none; }
        .download a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <h1>⚡ All thủy điện</h1>
    <div class="box">
        <!-- Tải file -->
        <div class="upload-area" id="khuTai" onclick="document.getElementById('chonFile').click()">
            <strong>📎 Tải file tài liệu lên</strong><br>
            <span style="color:#666; font-size:13px;">.docx .xlsx .txt .pdf</span>
            <input type="file" id="chonFile" accept=".docx,.xlsx,.txt,.pdf" style="display:none;" onchange="xuLyFile(this)">
        </div>
        <div class="file-info" id="thongTinFile">✅ Đã chọn: <span id="tenFile"></span></div>
        
        <!-- Ô nhập -->
        <textarea id="cauhoi" placeholder="Nhập yêu cầu hoặc câu hỏi ở đây..."></textarea>
        
        <!-- Nút gửi -->
        <button id="nutGui" onclick="gui()">Gửi & Phân tích</button>
        
        <!-- Kết quả -->
        <div class="result" id="ketQua"></div>
        
        <!-- Tải file -->
        <div class="download" id="linkTai">
            <strong>Tải kết quả:</strong><br>
            <a id="linkWord" href="#" target="_blank">📄 Word</a>
            <a id="linkExcel" href="#" target="_blank">📊 Excel</a>
            <a id="linkPdf" href="#" target="_blank">📕 PDF</a>
        </div>
    </div>

    <script>
        let fileContent = "";

        async function xuLyFile(input) {
            const file = input.files[0];
            if (!file) return;
            
            const khu = document.getElementById("khuTai");
            const thongTin = document.getElementById("thongTinFile");
            
            khu.classList.add("active");
            khu.innerHTML = "⏳ Đang đọc file...";
            
            const formData = new FormData();
            formData.append("file", file);
            
            try {
                const res = await fetch("/api/upload", { method: "POST", body: formData });
                const data = await res.json();
                
                if (data.status === "ok") {
                    fileContent = data.content;
                    khu.innerHTML = "✅ Sẵn sàng nhận file";
                    thongTin.style.display = "block";
                    document.getElementById("tenFile").textContent = data.name;
                } else {
                    khu.innerHTML = "❌ " + data.error;
                }
            } catch (e) {
                khu.innerHTML = "❌ Lỗi: " + e;
            }
        }

        async function gui() {
            const cauhoi = document.getElementById("cauhoi").value.trim();
            const nut = document.getElementById("nutGui");
            const ketQua = document.getElementById("ketQua");
            const linkTai = document.getElementById("linkTai");
            
            if (!cauhoi && !fileContent) {
                alert("Vui lòng nhập yêu cầu hoặc tải file lên!");
                return;
            }
            
            nut.disabled = true;
            nut.textContent = "Đang xử lý...";
            ketQua.style.display = "block";
            ketQua.textContent = "⏳ AI đang phân tích, vui lòng chờ...";
            linkTai.style.display = "none";
            
            try {
                const res = await fetch("/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ message: cauhoi, file_content: fileContent })
                });
                const data = await res.json();
                
                ketQua.textContent = data.reply || "Không có phản hồi";
                
                if (data.word || data.excel || data.pdf) {
                    linkTai.style.display = "block";
                    if (data.word) document.getElementById("linkWord").href = data.word;
                    if (data.excel) document.getElementById("linkExcel").href = data.excel;
                    if (data.pdf) document.getElementById("linkPdf").href = data.pdf;
                }
            } catch (e) {
                ketQua.textContent = "❌ Lỗi kết nối: " + e;
            }
            
            nut.disabled = false;
            nut.textContent = "Gửi & Phân tích";
        }
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
