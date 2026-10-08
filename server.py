import os
import re
import uuid
from datetime import datetime
import requests
from flask import Flask, request, jsonify, send_from_directory, session
from flask_cors import CORS
from docx import Document
from openpyxl import Workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
CORS(app)

# ==================================================
# CẤU HÌNH — ĐÃ ĐẶT SẴN THÔNG TIN CỦA BẠN
# ==================================================
def env(name, default=""):
    val = os.environ.get(name, default)
    return val.strip() if isinstance(val, str) else val

# === API Keys ===
GEMINI_API_KEY = env("GEMINI_API_KEY")
GEMINI_MODEL = env("GEMINI_MODEL", "gemini-2.0-flash")

GROQ_API_KEY = env("GROQ_API_KEY")
GROQ_MODEL = env("GROQ_MODEL", "llama-3.3-70b-versatile")

OPENROUTER_API_KEY = env("OPENROUTER") or env("OPENROUTER_API_KEY")
OPENROUTER_MODEL = env("OPENROUTER_MODEL", "openrouter/auto")

# === Đăng nhập — ĐÃ KHỚP THÔNG TIN BẠN CUNG CẤP ===
ADMIN_USER = env("ADMIN_USER", "chotvjp").strip()
ADMIN_PASS = env("ADMIN_PASS", "Do@058690").strip()
SECRET_KEY = env("SECRET_KEY") or os.urandom(32).hex()

app.secret_key = SECRET_KEY
app.config.update(
    SESSION_COOKIE_SECURE=bool(os.environ.get("RENDER")),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=12 * 3600
)

TIMEOUT = 120
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

DEFAULT_SYSTEM_PROMPT = "Trả lời bằng tiếng Việt rõ ràng, chính xác, dễ hiểu, phù hợp lĩnh vực thủy điện."

# ==================================================
# HÀM GỌI AI
# ==================================================
def call_openai_style(api_url, api_key, model, messages, max_tokens=2048):
    if not api_key:
        return None, "chưa đặt API key"
    try:
        resp = requests.post(
            api_url,
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": messages, "max_tokens": max_tokens, "temperature": 0.7},
            timeout=TIMEOUT
        )
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"].strip(), None
    except Exception as e:
        return None, str(e)

def call_gemini(prompt, system_prompt=DEFAULT_SYSTEM_PROMPT):
    if not GEMINI_API_KEY:
        return None, "chưa đặt GEMINI_API_KEY"
    try:
        url = f"https://generativelanguage.googleapis.com/v1/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
        full_prompt = f"Hướng dẫn hệ thống: {system_prompt}\n\nNội dung: {prompt}"
        resp = requests.post(
            url,
            headers={"Content-Type": "application/json"},
            json={"contents": [{"parts": [{"text": full_prompt}]}]},
            timeout=TIMEOUT
        )
        resp.raise_for_status()
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"].strip(), None
    except Exception as e:
        return None, str(e)

def call_groq(messages):
    return call_openai_style(
        "https://api.groq.com/openai/v1/chat/completions",
        GROQ_API_KEY, GROQ_MODEL, messages
    )

def call_openrouter(messages):
    return call_openai_style(
        "https://openrouter.ai/api/v1/chat/completions",
        OPENROUTER_API_KEY, OPENROUTER_MODEL, messages
    )

def chat_with_fallback(user_prompt, system_prompt=DEFAULT_SYSTEM_PROMPT):
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    attempts = [
        ("Gemini", lambda: call_gemini(user_prompt, system_prompt)),
        ("Groq", lambda: call_groq(messages)),
        ("OpenRouter", lambda: call_openrouter(messages)),
    ]
    for name, fn in attempts:
        text, err = fn()
        if text:
            return f"✅ [{name}]\n{text}", text
    return "❌ Tất cả AI đều không trả lời — kiểm tra lại API key", ""

# ==================================================
# XUẤT FILE
# ==================================================
def clean_text(text):
    return re.sub(r"[*_`#]", "", text)

def export_word(content):
    try:
        doc = Document()
        doc.add_heading("BÁO CÁO", 0)
        doc.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph("")
        for para in content.split("\n"):
            if para.strip():
                doc.add_paragraph(clean_text(para))
        fname = f"bao_cao_{uuid.uuid4().hex[:8]}.docx"
        path = os.path.join(OUTPUT_FOLDER, fname)
        doc.save(path)
        return fname
    except Exception as e:
        print(f"Lỗi tạo Word: {e}", flush=True)
        return None

def export_excel(content):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Báo cáo"
        ws["A1"] = f"Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws.append(["Nội dung"])
        for line in content.split("\n"):
            if line.strip():
                ws.append([clean_text(line)])
        fname = f"bao_cao_{uuid.uuid4().hex[:8]}.xlsx"
        path = os.path.join(OUTPUT_FOLDER, fname)
        wb.save(path)
        return fname
    except Exception as e:
        print(f"Lỗi tạo Excel: {e}", flush=True)
        return None

# ==================================================
# ĐĂNG NHẬP
# ==================================================
@app.before_request
def check_auth():
    public = ["/", "/login", "/healthz"]
    if any(request.path.startswith(p) for p in public):
        return
    if not session.get("authenticated"):
        return jsonify({"error": "Vui lòng đăng nhập"}), 401

@app.route("/login", methods=["POST"])
def login():
    data = request.get_json(silent=True) or {}
    u = data.get("username", "").strip()
    p = data.get("password", "").strip()
    # In ra log để kiểm tra
    print(f"Đăng nhập thử: user={u!r}, pass={p!r}", flush=True)
    print(f"Đúng: user={ADMIN_USER!r}, pass={ADMIN_PASS!r}", flush=True)
    if u == ADMIN_USER and p == ADMIN_PASS:
        session["authenticated"] = True
        return jsonify({"ok": True})
    return jsonify({"error": "Sai thông tin đăng nhập"}), 401

@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})

@app.route("/healthz")
def healthz():
    return "OK"

# ==================================================
# API
# ==================================================
@app.route("/api/chat", methods=["POST"])
def api_chat():
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "").strip()
    sys_prompt = data.get("system_prompt", DEFAULT_SYSTEM_PROMPT)
    if not prompt:
        return jsonify({"error": "Thiếu nội dung hỏi"}), 400
    display_text, raw_text = chat_with_fallback(prompt, sys_prompt)
    return jsonify({"reply": display_text, "raw": raw_text})

@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(silent=True) or {}
    content = data.get("content", "")
    kind = data.get("type", "word")
    if not content:
        return jsonify({"error": "Không có nội dung để xuất"}), 400
    fname = export_word(content) if kind == "word" else export_excel(content)
    if not fname:
        return jsonify({"error": "Lỗi tạo file"}), 500
    return jsonify({"url": f"/download/{fname}"})

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

# ==================================================
# GIAO DIỆN
# ==================================================
@app.route("/")
def index():
    return '''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Thủy Điện — Hệ thống hỗ trợ</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:Segoe UI,Roboto,sans-serif}
:root{--p:#2563eb;--s:#16a34a;--g:#f59e0b;--d:#1e293b}
body{background:linear-gradient(135deg,#eff6ff,#f0fdf4);min-height:100vh;padding:20px}
.header{text-align:center;margin-bottom:24px}
.header h1{color:var(--p);font-size:24px;margin-bottom:4px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:20px;max-width:1400px;margin:0 auto}
.card{background:#fff;border-radius:16px;padding:20px;box-shadow:0 4px 12px rgba(0,0,0,.06);display:flex;flex-direction:column;height:680px}
.card h2{font-size:16px;color:var(--p);margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #e2e8f0}
.chat-box{flex:1;overflow-y:auto;padding:12px;background:#f8fafc;border-radius:12px;margin-bottom:12px}
.msg{margin-bottom:12px;max-width:90%}
.msg.user{margin-left:auto}
.msg.ai{margin-right:auto}
.msg-bubble{padding:12px 16px;border-radius:16px;line-height:1.5;font-size:14px}
.msg.user .msg-bubble{background:var(--p);color:#fff;border-bottom-right-radius:4px}
.msg.ai .msg-bubble{background:#e2e8f0;color:var(--d);border-bottom-left-radius:4px}
.input-row{display:flex;gap:8px}
textarea{flex:1;padding:12px 16px;border:1px solid #e2e8f0;border-radius:12px;font-size:14px;outline:none;resize:none;height:60px}
textarea:focus{border-color:var(--p)}
button{padding:12px 20px;border:none;border-radius:12px;font-size:14px;font-weight:600;cursor:pointer}
.btn-primary{background:var(--p);color:#fff}
.btn-success{background:var(--s);color:#fff}
.btn-warning{background:var(--g);color:#fff}
.login-wrap{position:fixed;inset:0;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;z-index:999}
.login-box{background:#fff;padding:32px;border-radius:16px;box-shadow:0 8px 24px rgba(0,0,0,.15);width:90%;max-width:400px;text-align:center}
.login-box h2{color:var(--p);margin-bottom:20px}
.login-box input{width:100%;padding:12px 16px;border:1px solid #e2e8f0;border-radius:12px;font-size:14px;margin-bottom:12px}
.hidden{display:none!important}
.contact{font-size:13px;line-height:1.8;color:#475569}
</style>
</head>
<body>

<div id="loginScreen" class="login-wrap">
  <div class="login-box">
    <h2>🔐 Đăng nhập</h2>
    <input type="text" id="user" placeholder="Tên đăng nhập" value="chotvjp">
    <input type="password" id="pass" placeholder="Mật khẩu" onkeydown="event.key==='Enter'&&doLogin()">
    <button class="btn-primary" style="width:100%;margin-top:8px" onclick="doLogin()">Đăng nhập</button>
    <p id="loginErr" style="color:red;margin-top:10px;display:none">Sai thông tin đăng nhập</p>
  </div>
</div>

<div id="mainApp" class="hidden">
  <div class="header">
    <h1>⚡ All Thủy Điện — Hệ thống hỗ trợ</h1>
    <p>Soạn thảo văn bản · Tạo báo cáo · Quản lý thiết bị</p>
  </div>

  <div class="grid">
    <div class="card">
      <h2>✍️ Soạn thảo văn bản</h2>
      <div class="chat-box" id="chat1"></div>
      <textarea id="input1" placeholder="Nhập yêu cầu soạn thảo..."></textarea>
      <div class="input-row">
        <button class="btn-primary" onclick="sendMsg(1)">Gửi</button>
        <button class="btn-success" onclick="exportFile(1,'word')">Tải Word</button>
        <button class="btn-warning" onclick="exportFile(1,'excel')">Tải Excel</button>
      </div>
    </div>

    <div class="card">
      <h2>📊 Quy trình & Báo cáo</h2>
      <div class="chat-box" id="chat2"></div>
      <textarea id="input2" placeholder="Nhập yêu cầu phân tích..."></textarea>
      <div class="input-row">
        <button class="btn-primary" onclick="sendMsg(2)">Gửi</button>
        <button class="btn-success" onclick="exportFile(2,'word')">Tải Word</button>
        <button class="btn-warning" onclick="exportFile(2,'excel')">Tải Excel</button>
      </div>
    </div>

    <div class="card">
      <h2>📌 Thông tin liên hệ</h2>
      <div class="contact" style="flex:1">
        <p><strong>Admin:</strong> Chotvjp</p>
        <p><strong>ĐT:</strong> 0973020486</p>
        <p><strong>Gmail:</strong> hoangdien86ncc@gmail.com</p>
        <p><strong>Bộ phận:</strong> Quản lý kỹ thuật</p>
        <p><strong>Đơn vị:</strong> Công ty Cổ phần Thủy điện Nậm Chiến</p>
        <p><strong>Địa chỉ:</strong> TK5 - Mường La - Sơn La</p>
        <p><strong>Website:</strong> namchien.vn</p>
        <hr style="margin:16px 0;border:none;border-top:1px solid #e2e8f0">
        <p style="color:#64748b;font-size:12px">© 2026 — All Thủy Điện</p>
      </div>
      <button class="btn-primary" style="margin-top:8px" onclick="doLogout()">Đăng xuất</button>
    </div>
  </div>
</div>

<script>
let lastReply = {1:"",2:""};

async function doLogin(){
  const u = document.getElementById("user").value.trim();
  const p = document.getElementById("pass").value.trim();
  const res = await fetch("/login", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({username:u,password:p})
  });
  if(res.ok){
    document.getElementById("loginScreen").classList.add("hidden");
    document.getElementById("mainApp").classList.remove("hidden");
  }else{
    document.getElementById("loginErr").style.display = "block";
  }
}

async function doLogout(){
  await fetch("/logout", {method:"POST"});
  location.reload();
}

function addMsg(col,role,text){
  const box = document.getElementById(`chat${col}`);
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.innerHTML = `<div class="msg-bubble">${text.replace(/\n/g,"<br>")}</div>`;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
  if(role === "ai") lastReply[col] = text;
}

async function sendMsg(col){
  const ta = document.getElementById(`input${col}`);
  const text = ta.value.trim();
  if(!text) return;
  addMsg(col,"user",text);
  ta.value = "";
  addMsg(col,"ai","⏳ Đang xử lý...");
  
  const sys = col===1 
    ? "Bạn là trợ lý soạn thảo văn bản chính thức, chuẩn mực tiếng Việt."
    : "Bạn là chuyên gia thủy điện, phân tích dữ liệu, lập quy trình, báo cáo kỹ thuật.";
  
  const res = await fetch("/api/chat", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({prompt:text, system_prompt:sys})
  });
  const data = await res.json();
  
  const box = document.getElementById(`chat${col}`);
  box.lastChild.remove();
  addMsg(col,"ai", data.reply || "❌ Lỗi không nhận được phản hồi");
}

async function exportFile(col,type){
  if(!lastReply[col]) return alert("Chưa có nội dung để xuất");
  const res = await fetch("/api/export", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({content:lastReply[col], type:type})
  });
  const data = await res.json();
  if(data.url){
    window.open(data.url,"_blank");
  }else{
    alert("Lỗi xuất file: " + (data.error||""));
  }
}
</script>
</body>
</html>
'''

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
