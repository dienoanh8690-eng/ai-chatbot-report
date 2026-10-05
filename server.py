import os
import requests
from flask import Flask

app = Flask(__name__)

# ===================== CẤU HÌNH KHÓA =====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
AI_API_KEY = os.environ.get("AI_API_KEY", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "").strip()

TIMEOUT = 15
TEST_PROMPT = "Chào, trả lời ngắn: tôi là AI nào?"
# ==========================================================

def kiem_tra_gemini():
    ten = "🔵 Gemini"
    if not GEMINI_API_KEY:
        return ten, False, "Chưa đặt khóa"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"
    try:
        res = requests.post(url, json={
            "contents": [{"parts": [{"text": TEST_PROMPT}]}]
        }, timeout=TIMEOUT)
        if res.status_code == 200:
            txt = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
            return ten, True, f"✅ HOẠT ĐỘNG → {txt[:60]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:50]}"

def kiem_tra_groq():
    ten = "🟢 Groq/Llama"
    if not GROQ_API_KEY:
        return ten, False, "Chưa đặt khóa"
    url = "https://api.groq.com/openai/v1/chat/completions"
    try:
        res = requests.post(url, headers={
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json"
        }, json={
            "model": "llama-3.1-8b-instant",
            "messages": [{"role": "user", "content": TEST_PROMPT}],
            "max_tokens": 100
        }, timeout=TIMEOUT)
        if res.status_code == 200:
            txt = res.json()["choices"][0]["message"]["content"].strip()
            return ten, True, f"✅ HOẠT ĐỘNG → {txt[:60]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:50]}"

def kiem_tra_aiml():
    ten = "🟣 AIML/DOLA"
    if not AI_API_KEY:
        return ten, False, "Chưa đặt khóa"
    url = "https://api.aimlapi.com/v1/chat/completions"
    try:
        res = requests.post(url, headers={
            "Authorization": f"Bearer {AI_API_KEY}",
            "Content-Type": "application/json"
        }, json={
            "model": "bytedance/dola-seed-2-0-pro",
            "messages": [{"role": "user", "content": TEST_PROMPT}],
            "max_tokens": 100
        }, timeout=TIMEOUT)
        if res.status_code == 200:
            txt = res.json()["choices"][0]["message"]["content"].strip()
            return ten, True, f"✅ HOẠT ĐỘNG → {txt[:60]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:50]}"

def kiem_tra_openai():
    ten = "🔴 OpenAI/GPT"
    if not OPENAI_API_KEY:
        return ten, False, "Chưa đặt khóa"
    url = "https://api.openai.com/v1/chat/completions"
    try:
        res = requests.post(url, headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json"
        }, json={
            "model": "gpt-4o-mini",
            "messages": [{"role": "user", "content": TEST_PROMPT}],
            "max_tokens": 100
        }, timeout=TIMEOUT)
        if res.status_code == 200:
            txt = res.json()["choices"][0]["message"]["content"].strip()
            return ten, True, f"✅ HOẠT ĐỘNG → {txt[:60]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:50]}"

def kiem_tra_claude():
    ten = "🟠 Claude"
    if not CLAUDE_API_KEY:
        return ten, False, "Chưa đặt khóa"
    url = "https://api.anthropic.com/v1/messages"
    try:
        res = requests.post(url, headers={
            "x-api-key": CLAUDE_API_KEY,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }, json={
            "model": "claude-3-haiku-20240307",
            "max_tokens": 100,
            "messages": [{"role": "user", "content": TEST_PROMPT}]
        }, timeout=TIMEOUT)
        if res.status_code == 200:
            txt = res.json()["content"][0]["text"].strip()
            return ten, True, f"✅ HOẠT ĐỘNG → {txt[:60]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:50]}"

@app.route("/")
def kiem_tra_tat_ca():
    ham_kiem_tra = [
        kiem_tra_gemini,
        kiem_tra_groq,
        kiem_tra_aiml,
        kiem_tra_openai,
        kiem_tra_claude,
    ]

    html = """
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Kiểm tra kết nối AI</title>
        <style>
            body { font-family: Arial; padding: 30px; max-width: 800px; margin: 0 auto; background: #f5f7fa; }
            h1 { color: #1e40af; text-align: center; }
            .item { padding: 15px; margin: 10px 0; border-radius: 10px; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
            .ok { border-left: 5px solid #22c55e; }
            .fail { border-left: 5px solid #ef4444; }
            .stats { margin-top: 20px; padding: 15px; background: #e0e7ff; border-radius: 10px; font-weight: bold; }
            .note { margin-top: 30px; padding: 15px; background: #fef3c7; border-radius: 10px; }
        </style>
    </head>
    <body>
        <h1>🔍 KIỂM TRA KẾT NỐI CÁC AI</h1>
    """

    tong = 0
    dung = 0
    for ham in ham_kiem_tra:
        ten, ok, msg = ham()
        tong += 1
        if ok: dung += 1
        status_class = "ok" if ok else "fail"
        html += f'<div class="item {status_class}"><strong>{ten}</strong><br>{msg}</div>'

    html += f"""
        <div class="stats">
            Tổng số: {tong} AI | Hoạt động: {dung} | Lỗi: {tong - dung}
        </div>
        <div class="note">
            💡 <strong>Bước tiếp theo:</strong> Copy toàn bộ nội dung trang này gửi mình → mình sẽ viết code chính chỉ dùng AI đã chạy được!
        </div>
    </body>
    </html>
    """
    return html

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
