import os
import requests

# ===================== CẤU HÌNH KHÓA =====================
# Điền khóa của bạn vào đây, hoặc để trống nếu chưa có
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
AI_API_KEY = os.environ.get("AI_API_KEY", "").strip()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "").strip()

TIMEOUT = 15
TEST_PROMPT = "Chào, trả lời ngắn gọn: tôi là AI nào?"
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
            return ten, True, f"✅ OK → {txt[:50]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:40]}"

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
            return ten, True, f"✅ OK → {txt[:50]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:40]}"

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
            return ten, True, f"✅ OK → {txt[:50]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:40]}"

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
            return ten, True, f"✅ OK → {txt[:50]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:40]}"

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
            return ten, True, f"✅ OK → {txt[:50]}"
        return ten, False, f"❌ Lỗi {res.status_code}"
    except Exception as e:
        return ten, False, f"❌ Lỗi: {str(e)[:40]}"

# ===================== CHẠY KIỂM TRA =====================
if __name__ == "__main__":
    print("=" * 60)
    print("🔍 KIỂM TRA KẾT NỐI CÁC AI")
    print("=" * 60)

    ham_kiem_tra = [
        kiem_tra_gemini,
        kiem_tra_groq,
        kiem_tra_aiml,
        kiem_tra_openai,
        kiem_tra_claude,
    ]

    tong = 0
    dung = 0
    ket_qua = []

    for ham in ham_kiem_tra:
        ten, ok, msg = ham()
        tong += 1
        if ok:
            dung += 1
        ket_qua.append((ten, ok, msg))
        print(f"{ten:15} | {msg}")

    print("=" * 60)
    print(f"Kết quả: {dung}/{tong} AI hoạt động")
    print("=" * 60)

    if dung == 0:
        print("\n💡 Hướng dẫn: Điền khóa vào biến tương ứng ở đầu file")
        print("   hoặc đặt trên Render → Environment")
    else:
        print(f"\n✅ Có {dung} AI hoạt động — mình sẽ xây code chính dựa trên các AI này!")
        print("   Các AI chạy được:", ", ".join([ten for ten, ok, _ in ket_qua if ok]))
