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

# ==================================================
# CẤU HÌNH — Điền trên Render → Environment Variables
# ==================================================
# Router AI / AIML API — lấy khóa tại: https://aimlapi.com
ROUTER_API_KEY = os.environ.get("AI_API_KEY", "").strip()
ROUTER_URL = "https://api.aimlapi.com/v1/chat/completions"
# Các model dự phòng — tự động chuyển nếu 1 cái bận
MODEL_LIST = [
    "google/gemini-2.5-flash",
    "bytedance/dola-seed-2-0-pro",
    "meta-llama/llama-3.2-1b-instruct",
    "mistralai/mistral-7b-instruct-v0.3"
]

RESULT_FOLDER = "bao_cao_xuat_ra"
os.makedirs(RESULT_FOLDER, exist_ok=True)

# ==================================================
# HÀM GỌI ROUTER AI — TỰ ĐỘNG CHUYỂN MODEL
# ==================================================
def goi_router(prompt, he_thong=""):
    if not ROUTER_API_KEY:
        return None, "⚠️ Chưa đặt AI_API_KEY trên Render"
    
    loi = []
    for model in MODEL_LIST:
        try:
            res = requests.post(
                ROUTER_URL,
                headers={
                    "Authorization": f"Bearer {ROUTER_API_KEY}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": model,
                    "messages": [
                        {"role": "system", "content": he_thong or "Trả lời bằng tiếng Việt rõ ràng, dễ hiểu."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.7,
                    "max_tokens": 1024
                },
                timeout=30
            )
            if res.status_code == 200:
                return f"✅ [{model.split('/')[-1]}]\n" + res.json()["choices"][0]["message"]["content"], None
            loi.append(f"{model}: Lỗi {res.status_code}")
        except Exception as e:
            loi.append(f"{model}: {str(e)}")
    
    return None, "❌ Tất cả đều không trả lời:\n" + "\n".join(loi)

# ==================================================
# TẠO FILE WORD & EXCEL
# ==================================================
def tao_word(noi_dung):
    try:
        ten = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        path = os.path.join(RESULT_FOLDER, ten)
        doc = Document()
        doc.add_heading("BÁO CÁO", 0)
        doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph("-" * 50)
        for dong in noi_dung.split("\n"):
            if dong.strip():
                doc.add_paragraph(dong.strip())
        doc.save(path)
        return ten
    except Exception as e:
        print(f"Lỗi tạo Word: {e}")
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
        print(f"Lỗi tạo Excel: {e}")
        return ""

# ==================================================
# API
# ==================================================
@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    msg = data.get("message", "").strip()
    loai = data.get("loai", "chung")

    if not msg:
        return jsonify({"reply": "Vui lòng nhập nội dung cần hỗ trợ!"})

    cau_hinh = {
        "chung": "Bạn là trợ lý AI tổng hợp. Trả lời rõ ràng, dễ hiểu, bằng tiếng Việt.",
        "vanban": "Bạn là chuyên gia soạn thảo văn bản hành chính, công văn, hợp đồng theo chuẩn Việt Nam. Viết trang trọng, đúng thể thức.",
        "baocao": "Bạn là chuyên gia phân tích và lập báo cáo. Trình bày có cấu trúc rõ ràng, tóm tắt số liệu chính.",
        "dauthau": "Bạn là chuyên gia tư vấn đấu thầu theo pháp luật Việt Nam. Hướng dẫn chi tiết từng bước, hồ sơ, lưu ý pháp lý."
    }

    he_thong = cau_hinh.get(loai, cau_hinh["chung"])
    tra_loi, loi = goi_router(msg, he_thong)
    
    if not tra_loi:
        return jsonify({"reply": loi})

    word = tao_word(tra_loi)
    excel = tao_excel(tra_loi)

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
*{margin:0;padding:0;box-sizing:border-box;font-family:"Segoe UI",Arial,sans-serif}
body{background:linear-gradient(135deg,#eff6ff,#faf5ff);min-height:100vh;padding:20px}
.container{max-width:850px;margin:0 auto}
h1{text-align:center;color:#1e40af;margin-bottom:5px;font-size:22px}
.desc{text-align:center;color:#64748b;margin-bottom:25px;font-size:14px}
.tabs{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px;justify-content:center}
.tab{padding:10px 18px;border:none;border-radius:24px;background:#e2e8f0;cursor:pointer;font-size:14px;transition:all .2s}
.tab.active{background:#2563eb;color:#fff;font-weight:600;box-shadow:0 2px 8px rgba(37,99,235,.2)}
.card{background:#fff;border-radius:20px;padding:24px;box-shadow:0 4px 20px rgba(37,99,235,.08)}
.quick{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px}
.q-btn{padding:9px 14px;border:1px solid #dbeafe;background:#f0f7ff;border-radius:12px;cursor:pointer;font-size:13px;color:#1e40af;transition:all .2s}
.q-btn:hover{background:#dbeafe;transform:translateY(-1px)}
.chat{height:420px;overflow-y:auto;padding:16px;border:1px solid #e0e7ff;border-radius:16px;margin-bottom:16px;background:#fafbff}
.msg{margin-bottom:18px;max-width:92%;animation:fadeIn .3s forwards;opacity:0}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.msg.user{margin-left:auto}
.msg.ai{margin-right:auto}
.bubble{padding:14px 18px;border-radius:18px;line-height:1.6;white-space:pre-wrap;font-size:14px}
.msg.user .bubble{background:linear-gradient(135deg,#dbeafe,#e0e7ff);border-bottom-right-radius:6px;color:#1e3a8a}
.msg.ai .bubble{background:#f1f5f9;border-bottom-left-radius:6px;color:#1e293b}
.bubble.err{background:#fef2f2;border-left:3px solid #ef4444;color:#b91c1c}
.bubble.warn{background:#fffbeb;border-left:3px solid #f59e0b;color:#92400e}
.input-row{display:flex;gap:12px;align-items:flex-end}
textarea{flex:1;padding:14px 20px;border:1px solid #e2e8f0;border-radius:24px;font-size:15px;resize:none;height:52px;outline:none;line-height:1.4}
textarea:focus{border-color:#2563eb;box-shadow:0 0 0 4px rgba(37,99,235,.1)}
.send{width:52px;height:52px;border-radius:50%;border:none;background:#2563eb;color:#fff;cursor:pointer;font-size:20px;transition:all .2s}
.send:disabled{opacity:.5;cursor:not-allowed;transform:none}
.send:hover:not(:disabled){transform:scale(1.1)}
.dl{margin-top:14px;display:flex;gap:12px;flex-wrap:wrap}
.dl a{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border-radius:24px;text-decoration:none;font-size:14px;font-weight:600;transition:transform .2s}
.dl-word{background:#dbeafe;color:#1d4ed8}
.dl-excel{background:#dcfce7;color:#15803d}
.dl a:hover{transform:scale(1.05)}
</style>
</head>
<body>
<div class="container">
<h1>⚡ All Thủy Điện — Hỗ trợ công việc</h1>
<p class="desc">Router AI · Gemini+Dola+Llama · Xuất Word/Excel</p>

<div class="tabs">
<button class="tab active" data-loai="chung">💬 Trò chuyện chung</button>
<button class="tab" data-loai="vanban">✍️ Soạn văn bản</button>
<button class="tab" data-loai="baocao">📊 Lập báo cáo</button>
<button class="tab" data-loai="dauthau">🏆 Đấu thầu</button>
</div>

<div class="card">
<div class="quick">
<button class="q-btn" onclick="goiNhanh('Giải thích khái niệm về nhà máy thủy điện')">Thủy điện cơ bản</button>
<button class="q-btn" onclick="goiNhanh('Soạn công văn gửi cấp trên báo cáo tiến độ')">Công văn mẫu</button>
<button class="q-btn" onclick="goiNhanh('Lập báo cáo tổng hợp công việc tháng có cấu trúc')">Báo cáo tháng</button>
<button class="q-btn" onclick="goiNhanh('Tóm tắt quy trình đấu thầu theo luật định')">Quy trình đấu thầu</button>
</div>

<div class="chat" id="chatKhu">
<div class="msg ai"><div class="bubble">👋 Xin chào! Tôi dùng Router AI — tự động chuyển đổi giữa Gemini, Dola, Llama khi cần. Nhập yêu cầu nhé!</div></div>
</div>

<div class="input-row">
<textarea id="inputCau" placeholder="Nhập nội dung cần hỗ trợ..." onkeydown="xuLyEnter(event)"></textarea>
<button class="send" id="nutGui" onclick="guiCau()">➤</button>
</div>
</div>
</div>

<script>
let loaiHienTai = "chung";

document.querySelectorAll(".tab").forEach(tab => {
    tab.onclick = () => {
        document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
        tab.classList.add("active");
        loaiHienTai = tab.dataset.loai;
    };
});

function xuLyEnter(e){ if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); guiCau(); } }
function goiNhanh(nd){ document.getElementById("inputCau").value = nd; guiCau(); }

async function guiCau(){
    const input = document.getElementById("inputCau");
    const noiDung = input.value.trim();
    if(!noiDung) return;

    themTinNhan("user", noiDung);
    input.value = "";

    const nut = document.getElementById("nutGui");
    nut.disabled = true; nut.textContent = "⏳";

    const phanHoi = await fetch("/api/chat", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({message: noiDung, loai: loaiHienTai})
    });

    const duLieu = await phanHoi.json();
    themTinNhan("ai", duLieu.reply, duLieu.word, duLieu.excel);

    nut.disabled = false; nut.textContent = "➤";
}

function themTinNhan(loai, noiDung, linkWord, linkExcel){
    const khu = document.getElementById("chatKhu");
    const div = document.createElement("div");
    div.className = "msg " + loai;

    let classBubble = "";
    if(noiDung.includes("❌")) classBubble = " err";
    else if(noiDung.includes("⚠️")) classBubble = " warn";

    let linkTai = "";
    if(linkWord || linkExcel){
        linkTai = '<div class="dl">';
        if(linkWord) linkTai += '<a href="'+linkWord+'" class="dl-word" target="_blank">📄 Tải Word</a>';
        if(linkExcel) linkTai += '<a href="'+linkExcel+'" class="dl-excel" target="_blank">📊 Tải Excel</a>';
        linkTai += "</div>";
    }

    const anToan = noiDung.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    div.innerHTML = '<div class="bubble'+classBubble+'">'+anToan+"</div>"+linkTai;
    khu.appendChild(div);
    khu.scrollTop = khu.scrollHeight;
}
</script>
</body>
</html>
'''

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
