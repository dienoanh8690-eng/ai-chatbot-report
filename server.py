from flask import Flask, request, jsonify, send_file
import requests
import os
import uuid
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ========== CẤU HÌNH — Đặt trên Render → Environment ==========
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"

AI_API_KEY = os.environ.get("AI_API_KEY", "").strip()
AIML_URL = "https://api.aimlapi.com/v1/chat/completions"
AIML_MODEL = "bytedance/dola-seed-2-0-pro"

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "llama-3.1-8b-instant"

RESULT_FOLDER = "bao_cao_xuat_ra"
os.makedirs(RESULT_FOLDER, exist_ok=True)

# ========== HÀM GỌI AI ==========
def goi_gemini(prompt, he_thong=""):
    if not GEMINI_API_KEY:
        return "Gemini", None, "⚠️ Chưa đặt GEMINI_API_KEY"
    try:
        full = f"{he_thong}\nYêu cầu: {prompt}" if he_thong else prompt
        res = requests.post(GEMINI_URL, json={"contents": [{"parts": [{"text": full}]}]}, timeout=30)
        if res.status_code == 200:
            data = res.json()
            return "Gemini", data["candidates"][0]["content"]["parts"][0]["text"], None
        return "Gemini", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "Gemini", None, f"Lỗi: {str(e)}"

def goi_aiml(prompt, he_thong=""):
    if not AI_API_KEY:
        return "DOLA", None, "⚠️ Chưa đặt AI_API_KEY"
    try:
        res = requests.post(
            AIML_URL,
            headers={"Authorization": f"Bearer {AI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": AIML_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 1024
            },
            timeout=30
        )
        if res.status_code == 200:
            return "DOLA", res.json()["choices"][0]["message"]["content"], None
        return "DOLA", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "DOLA", None, f"Lỗi: {str(e)}"

def goi_groq(prompt, he_thong=""):
    if not GROQ_API_KEY:
        return "Llama", None, "⚠️ Chưa đặt GROQ_API_KEY"
    try:
        res = requests.post(
            GROQ_URL,
            headers={"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": GROQ_MODEL,
                "messages": [
                    {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 1024
            },
            timeout=30
        )
        if res.status_code == 200:
            return "Llama", res.json()["choices"][0]["message"]["content"], None
        return "Llama", None, f"Lỗi {res.status_code}"
    except Exception as e:
        return "Llama", None, f"Lỗi: {str(e)}"

def goi_ai_tu_dong(prompt, danh_sach, he_thong=""):
    loi = []
    for ham in danh_sach:
        ten, kq, err = ham(prompt, he_thong)
        if kq:
            return f"✅ [{ten}]\n{kq}"
        loi.append(f"{ten}: {err}")
    return "❌ Tất cả AI đều không trả lời:\n" + "\n".join(loi)

# ========== TẠO FILE XUẤT ==========
def tao_word(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        path = os.path.join(RESULT_FOLDER, ten)
        doc = Document()
        doc.add_heading("BÁO CÁO", 0)
        doc.add_paragraph(f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph("-" * 50)
        for dong in noi_dung.split("\n"):
            if dong.strip():
                doc.add_paragraph(dong.strip())
        doc.save(path)
        return ten
    except Exception as e:
        print(f"Lỗi Word: {e}")
        return ""

def tao_excel(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
        path = os.path.join(RESULT_FOLDER, ten)
        wb = Workbook()
        ws = wb.active
        ws.append([f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}"])
        ws.append([])
        for dong in noi_dung.split("\n"):
            if dong.strip():
                ws.append([dong.strip()])
        wb.save(path)
        return ten
    except Exception as e:
        print(f"Lỗi Excel: {e}")
        return ""

# ========== API ROUTES ==========
@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    loai = data.get("loai", "chung")
    
    if not msg:
        return jsonify({"reply": "Vui lòng nhập nội dung!"})
    
    cau_hinh = {
        "chung": ("Trợ lý tổng hợp, trả lời rõ ràng dễ hiểu.", [goi_gemini, goi_aiml, goi_groq]),
        "vanban": ("Chuyên gia soạn thảo văn bản hành chính, chuẩn mực Việt Nam.", [goi_gemini, goi_aiml, goi_groq]),
        "baocao": ("Chuyên gia phân tích, lập báo cáo có cấu trúc rõ ràng.", [goi_gemini, goi_aiml, goi_groq]),
        "dauthau": ("Chuyên gia tư vấn đấu thầu theo pháp luật Việt Nam, hướng dẫn chi tiết.", [goi_gemini, goi_aiml, goi_groq]),
    }
    
    he_thong, ds_ai = cau_hinh.get(loai, cau_hinh["chung"])
    tra_loi = goi_ai_tu_dong(msg, ds_ai, he_thong)
    
    word = tao_word(tra_loi) if "✅" in tra_loi else ""
    excel = tao_excel(tra_loi) if "✅" in tra_loi else ""
    
    return jsonify({
        "reply": tra_loi,
        "word": f"/download/{word}" if word else "",
        "excel": f"/download/{excel}" if excel else ""
    })

@app.route("/download/<ten_file>")
def download(ten_file):
    path = os.path.join(RESULT_FOLDER, ten_file)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "Không tìm thấy tệp", 404

@app.route("/")
def trang_chu():
    return '''
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hỗ trợ công việc</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
body{background:#f0f7ff;padding:20px;max-width:900px;margin:0 auto}
h1{text-align:center;color:#1e40af;margin-bottom:5px;font-size:22px}
p.desc{text-align:center;color:#64748b;margin-bottom:25px}
.tab{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
.tab-btn{padding:10px 16px;border:none;border-radius:20px;background:#e2e8f0;cursor:pointer;font-size:14px;transition:all .2s}
.tab-btn.active{background:#2563eb;color:#fff;font-weight:bold}
.box{background:#fff;border-radius:16px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.08)}
.chat{height:400px;overflow-y:auto;margin-bottom:15px;padding:10px;border:1px solid #e2e8f0;border-radius:12px}
.msg{margin-bottom:12px;max-width:90%}
.msg.user{margin-left:auto;text-align:right}
.bubble{padding:12px 16px;border-radius:16px;white-space:pre-wrap;line-height:1.5}
.msg.user .bubble{background:#dbeafe;color:#1e40af;border-bottom-right-radius:4px}
.msg.ai .bubble{background:#f1f5f9;color:#1e293b;border-bottom-left-radius:4px}
.in-row{display:flex;gap:10px;align-items:flex-end}
textarea{flex:1;padding:12px 16px;border:1px solid #e2e8f0;border-radius:24px;font-size:14px;resize:none;height:48px;outline:none}
textarea:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.send{width:44px;height:44px;border-radius:50%;border:none;background:#2563eb;color:#fff;cursor:pointer;font-size:18px}
.send:disabled{opacity:.5;cursor:not-allowed}
.dl{margin-top:12px;display:flex;gap:10px;flex-wrap:wrap}
.dl a{padding:8px 16px;border-radius:20px;text-decoration:none;font-size:13px;font-weight:bold}
.dl-word{background:#dbeafe;color:#1d4ed8}
.dl-excel{background:#dcfce7;color:#15803d}
.quick{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:15px}
.quick button{padding:8px 12px;border:1px solid #e2e8f0;background:#fafafa;border-radius:8px;cursor:pointer;font-size:13px}
.quick button:hover{border-color:#2563eb;background:#eff6ff}
</style>
</head>
<body>
<h1>⚡ All Thủy Điện — Hỗ trợ công việc</h1>
<p class="desc">Trò chuyện · Soạn văn bản · Lập báo cáo · Tư vấn đấu thầu</p>

<div class="tab">
<button class="tab-btn active" data-loai="chung">💬 Trò chuyện chung</button>
<button class="tab-btn" data-loai="vanban">✍️ Soạn văn bản</button>
<button class="tab-btn" data-loai="baocao">📊 Lập báo cáo</button>
<button class="tab-btn" data-loai="dauthau">🏆 Đấu thầu</button>
</div>

<div class="box">
<div class="quick" id="quickBtns">
<button onclick="goi('Giải thích khái niệm về nhà máy thủy điện')">Thủy điện cơ bản</button>
<button onclick="goi('Soạn công văn gửi cấp trên')">Công văn mẫu</button>
<button onclick="goi('Lập báo cáo tổng hợp công việc tháng')">Báo cáo tháng</button>
<button onclick="goi('Tóm tắt quy trình đấu thầu')">Quy trình đấu thầu</button>
</div>
<div class="chat" id="chatBox">
<div class="msg ai"><div class="bubble">👋 Xin chào! Tôi sẵn sàng hỗ trợ bạn. Chọn chức năng hoặc nhập yêu cầu nhé!</div></div>
</div>
<div class="in-row">
<textarea id="inputMsg" placeholder="Nhập nội dung cần hỗ trợ..." onkeydown="enterGui(event)"></textarea>
<button class="send" id="btnSend" onclick="gui()">➤</button>
</div>
</div>

<script>
let loaiHienTai = "chung";
document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.onclick = () => {
        document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
        btn.classList.add("active");
        loaiHienTai = btn.dataset.loai;
    };
});

function enterGui(e){ if(e.key==="Enter" && !e.shiftKey){e.preventDefault();gui();} }
function goi(nd){ document.getElementById("inputMsg").value=nd;gui(); }

async function gui(){
    const i = document.getElementById("inputMsg");
    const m = i.value.trim(); if(!m) return;
    themMsg("user", m); i.value="";
    const b = document.getElementById("btnSend");
    b.disabled=true; b.textContent="⏳";
    
    const res = await fetch("/api/chat", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({message:m, loai:loaiHienTai})
    });
    const d = await res.json();
    themMsg("ai", d.reply, d.word, d.excel);
    b.disabled=false; b.textContent="➤";
}

function themMsg(loai, nd, word, excel){
    const c = document.getElementById("chatBox");
    const div = document.createElement("div");
    div.className = "msg "+loai;
    let link = "";
    if(word || excel){
        link = '<div class="dl">';
        if(word) link += '<a href="'+word+'" class="dl-word" target="_blank">📄 Tải Word</a>';
        if(excel) link += '<a href="'+excel+'" class="dl-excel" target="_blank">📊 Tải Excel</a>';
        link += "</div>";
    }
    div.innerHTML = '<div class="bubble">'+nd.replace(/&/g,"&amp;").replace(/</g,"&lt;")+"</div>"+link;
    c.appendChild(div);
    c.scrollTop = c.scrollHeight;
}
</script>
</body>
</html>
'''

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
