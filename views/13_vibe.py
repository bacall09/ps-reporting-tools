import streamlit as st
import requests
import random

from components import inject_css
inject_css()





st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        .vibe-message {
            font-size: 1.2rem;
            font-weight: 700;
            color: inherit;
            margin: 1rem 0;
            text-align: center;
            line-height: 1.5;
            padding: 1rem 1.5rem;
            border-radius: 10px;
            border: 1px solid rgba(128,128,128,0.2);
        }
        .gif-container { display: flex; justify-content: center; margin: 1rem 0 1.5rem 0; }
        .footer-note { text-align: center; font-size: 11px; opacity: 0.35; margin-top: 2rem; }
    </style>
""", unsafe_allow_html=True)

MOOD_RESPONSES = {
    "happy":       "That energy? Bottle it. Your utilization is probably looking great too. 🚀",
    "stressed":    "Breathe. Scope creep is temporary. Delivery cred is forever. 💪",
    "tired":       "You've shipped more before 9am than most do all week. Rest is earned. 😴",
    "excited":     "New SOW energy. Love to see it. Don't let Finance slow you down. ⚡",
    "overwhelmed": "One ticket at a time. You didn't build the backlog in a day. 🧘",
    "frustrated":  "If it were easy, they wouldn't need PS. Keep going. 🔥",
    "motivated":   "Main character behavior. Your team feels it. Keep leading. 🌟",
    "bored":       "That's just your pipeline being too clean. Go find a complex customer. 😏",
    "anxious":     "The data is in. The report is built. You've got this. Exhale. 🌬️",
    "grateful":    "That's the good stuff. Hold onto it on the hard days. 🙏",
}

DEFAULT_RESPONSE = "Wherever you're at — you showed up today. That counts. ✨"

MOOD_GIF_TERMS = {
    "happy":       "happy celebration office",
    "stressed":    "deep breath relax calm",
    "tired":       "tired sleepy office worker",
    "excited":     "excited pumped lets go",
    "overwhelmed": "you got this team support",
    "frustrated":  "persistence keep going determination",
    "motivated":   "motivated champion success",
    "bored":       "waiting patiently office",
    "anxious":     "its going to be okay reassuring",
    "grateful":    "thank you appreciation wholesome",
}

DEFAULT_GIF_TERM = "you got this keep going"

# Words that should never appear in a GIF title or slug on a professional site
BLOCK_TERMS = [
    "sexy", "hot", "bikini", "drunk", "beer", "wine", "alcohol", "party",
    "fight", "angry", "rage", "cry", "crying", "sad", "fail", "failure",
    "wtf", "omg", "damn", "hell", "ass", "rude", "kiss", "love", "flirt",
]

def _is_safe(gif: dict) -> bool:
    """Return True if GIF title and slug contain no blocked terms."""
    title = gif.get("title", "").lower()
    slug  = gif.get("slug", "").lower()
    combined = title + " " + slug
    return not any(term in combined for term in BLOCK_TERMS)

def get_gif(search_term, seed=0):
    try:
        api_key = st.secrets.get("GIPHY_API_KEY", "")
        if not api_key:
            return None
        r = requests.get(
            "https://api.giphy.com/v1/gifs/search",
            params={
                "api_key": api_key,
                "q": search_term,
                "limit": 50,        # fetch more so filtering still leaves good options
                "rating": "g",      # strictest Giphy rating — general audiences only
                "lang": "en",
            },
            timeout=5,
        )
        data = r.json().get("data", [])
        # Filter to safe GIFs only
        safe = [g for g in data if _is_safe(g)]
        if not safe:
            safe = data  # fallback to unfiltered if nothing passes (shouldn't happen)
        if not safe:
            return None
        rng = random.Random(seed)
        return rng.choice(safe)["images"]["downsized_large"]["url"]
    except Exception:
        return None

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px'>
    <h1 style='color:#fff;margin:0;font-size:26px;font-weight:700'>✨ Vibe Check</h1>
    <p style='color:#a0aec0;margin:8px 0 0;font-size:13px'>
        Rough day? Great day? Tell us. We've got a GIF for that.
    </p>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "vibe_result" not in st.session_state:
    st.session_state["vibe_result"] = None
if "vibe_seed" not in st.session_state:
    st.session_state["vibe_seed"] = 0

# ── Show result at top ────────────────────────────────────────────────────────
if st.session_state["vibe_result"]:
    result = st.session_state["vibe_result"]

    if result["gif_url"]:
        st.markdown('<div class="gif-container">', unsafe_allow_html=True)
        st.image(result["gif_url"], use_container_width=False, width=420)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="vibe-message">{result["message"]}</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🎲 Give me another one", use_container_width=True):
            st.session_state["vibe_seed"] += 1
            st.session_state["vibe_result"]["gif_url"] = get_gif(
                result["gif_term"], seed=st.session_state["vibe_seed"]
            )
            st.rerun()
    with col2:
        if st.button("↩ Check a different vibe", use_container_width=True):
            st.session_state["vibe_result"] = None
            st.session_state["vibe_seed"]   = 0
            st.rerun()

    st.markdown('<hr style="border:none;border-top:1px solid rgba(128,128,128,0.15);margin:24px 0">', unsafe_allow_html=True)

# ── Input form ────────────────────────────────────────────────────────────────
mood_options = ["Select your mood..."] + [m.capitalize() for m in MOOD_RESPONSES.keys()]
selected     = st.selectbox("How are you feeling right now?", mood_options)
custom_mood  = st.text_input("Or describe it in your own words",
                              placeholder="e.g. 'surviving', 'thriving', 'in back-to-back calls'...")

if st.button("Check My Vibe →", type="primary", use_container_width=True):
    mood_key = selected.lower() if selected != "Select your mood..." else None
    if custom_mood.strip():
        mood_key = None

    message  = MOOD_RESPONSES.get(mood_key, DEFAULT_RESPONSE)
    gif_term = MOOD_GIF_TERMS.get(mood_key, custom_mood.strip() or DEFAULT_GIF_TERM)
    seed     = random.randint(0, 999)

    st.session_state["vibe_result"] = {"message": message, "gif_url": get_gif(gif_term, seed), "gif_term": gif_term}
    st.session_state["vibe_seed"]   = seed
    st.rerun()

st.markdown('<div class="footer-note">Powered by Giphy · PS Tools · Not a substitute for PTO</div>', unsafe_allow_html=True)
