import streamlit as st
import requests
import random

st.set_page_config(page_title="Vibe Check", page_icon="✨", layout="centered")

# ── Styling (matches app theme) ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Manrope', sans-serif;
}

.vibe-header {
    text-align: center;
    padding: 2rem 0 0.5rem 0;
}

.vibe-title {
    font-size: 2.4rem;
    font-weight: 800;
    color: #1a1a2e;
    margin-bottom: 0.2rem;
}

.vibe-subtitle {
    font-size: 1rem;
    color: #666;
    margin-bottom: 2rem;
}

.mood-card {
    background: #f8f9ff;
    border-radius: 16px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    border: 1px solid #e8eaf6;
}

.result-label {
    font-size: 0.85rem;
    font-weight: 600;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 0.5rem;
}

.vibe-message {
    font-size: 1.25rem;
    font-weight: 700;
    color: #1a1a2e;
    margin-bottom: 1.5rem;
    text-align: center;
    line-height: 1.4;
}

div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 0.6rem 2rem;
    font-family: 'Manrope', sans-serif;
    font-weight: 700;
    font-size: 1rem;
    width: 100%;
    transition: opacity 0.2s;
}

div[data-testid="stButton"] > button:hover {
    opacity: 0.88;
}

.gif-container {
    display: flex;
    justify-content: center;
    margin-top: 1rem;
}

.footer-note {
    text-align: center;
    font-size: 0.78rem;
    color: #bbb;
    margin-top: 2.5rem;
}
</style>
""", unsafe_allow_html=True)

# ── PS-flavored mood messages ─────────────────────────────────────────────────
MOOD_RESPONSES = {
    "happy":      "That energy? Bottle it. Your utilization is probably looking great too. 🚀",
    "stressed":   "Breathe. Scope creep is temporary. Delivery cred is forever. 💪",
    "tired":      "You've shipped more before 9am than most do all week. Rest is earned. 😴",
    "excited":    "New SOW energy. Love to see it. Don't let Finance slow you down. ⚡",
    "overwhelmed":"One ticket at a time. You didn't build the backlog in a day. 🧘",
    "frustrated": "If it were easy, they wouldn't need PS. Keep going. 🔥",
    "motivated":  "Main character behavior. Your team feels it. Keep leading. 🌟",
    "bored":      "That's just your pipeline being too clean. Go find a complex customer. 😏",
    "anxious":    "The data is in. The report is built. You've got this. Exhale. 🌬️",
    "grateful":   "That's the good stuff. Hold onto it on the hard days. 🙏",
}

DEFAULT_RESPONSE = "Wherever you're at — you showed up today. That counts. ✨"

# ── Giphy search terms per mood ───────────────────────────────────────────────
MOOD_GIF_TERMS = {
    "happy":      "celebrating happy dance",
    "stressed":   "its fine everything is fine",
    "tired":      "sleepy tired coffee",
    "excited":    "lets go excited pumped",
    "overwhelmed":"deep breath you got this",
    "frustrated": "keep going push through",
    "motivated":  "champion winner trophy",
    "bored":      "waiting bored tapping fingers",
    "anxious":    "nervous but okay",
    "grateful":   "thank you grateful wholesome",
}

DEFAULT_GIF_TERM = "you got this motivational"

# ── Giphy fetch ───────────────────────────────────────────────────────────────
def get_gif(search_term: str) -> str | None:
    try:
        api_key = st.secrets.get("GIPHY_API_KEY", "")
        if not api_key:
            return None
        url = "https://api.giphy.com/v1/gifs/search"
        params = {
            "api_key": api_key,
            "q": search_term,
            "limit": 20,
            "rating": "pg",
            "lang": "en",
        }
        r = requests.get(url, params=params, timeout=5)
        data = r.json().get("data", [])
        if not data:
            return None
        gif = random.choice(data)
        return gif["images"]["downsized_large"]["url"]
    except Exception:
        return None

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="vibe-header">
    <div class="vibe-title">✨ Vibe Check</div>
    <div class="vibe-subtitle">Rough day? Great day? Tell us. We've got a GIF for that.</div>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="mood-card">', unsafe_allow_html=True)

    mood_options = ["Select your mood..."] + [m.capitalize() for m in MOOD_RESPONSES.keys()]
    selected = st.selectbox("How are you feeling right now?", mood_options, label_visibility="collapsed")

    custom_mood = st.text_input(
        "Or describe it in your own words",
        placeholder="e.g. 'surviving', 'thriving', 'in back-to-back calls'...",
        label_visibility="visible"
    )

    st.markdown('</div>', unsafe_allow_html=True)

    go = st.button("Check My Vibe →")

if go:
    mood_key = selected.lower() if selected != "Select your mood..." else None
    if custom_mood.strip():
        mood_key = None  # use default message for freeform

    message = MOOD_RESPONSES.get(mood_key, DEFAULT_RESPONSE)
    gif_term = MOOD_GIF_TERMS.get(mood_key, custom_mood.strip() or DEFAULT_GIF_TERM)

    st.markdown(f'<div class="vibe-message">{message}</div>', unsafe_allow_html=True)

    gif_url = get_gif(gif_term)

    if gif_url:
        st.markdown('<div class="gif-container">', unsafe_allow_html=True)
        st.image(gif_url, use_container_width=False, width=420)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("💡 Add your GIPHY_API_KEY to Streamlit secrets to unlock GIFs. Get one free at developers.giphy.com", icon="🎬")

    st.button("Give me another one 🎲")

st.markdown('<div class="footer-note">Powered by Giphy · PS Tools · Not a substitute for PTO</div>', unsafe_allow_html=True)
