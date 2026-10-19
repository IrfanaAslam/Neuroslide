import streamlit as st
import json, os, io, time, tempfile
import pandas as pd

from ai_engine.ollama_client import generate_slides, check_ollama, list_models, analyze_csv_with_ai
from ppt_generator.create_ppt import build_presentation, THEMES
from pdf_generator.create_pdf import build_pdf
from data_processor.analyze_data import analyze_dataframe
from data_processor.charts import generate_all_charts, pick_chart_for_slide

st.set_page_config(page_title="NeuraSlide v5", page_icon="🧠", layout="wide",
                   initial_sidebar_state="expanded")

# ═══════════════════════════════════════════════════════════════════════════════
#  GLOBAL DARK CSS — every element forced dark
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Inter:wght@300;400;500;600&display=swap');

/* ── Reset & Base ── */
html, body, [class*="css"], * { font-family: 'Inter', sans-serif !important; box-sizing: border-box; }
.stApp { background: #080810 !important; color: #e8e6ff !important; }
header[data-testid="stHeader"] { background: rgba(8,8,16,0.95) !important; border-bottom: 1px solid rgba(124,58,237,0.15); }
section[data-testid="stSidebar"] { background: #06060e !important; border-right: 1px solid rgba(124,58,237,0.15) !important; }
.block-container { padding-top: 1.5rem !important; }

/* ── ALL INPUTS ── */
input, textarea, select,
.stTextInput input, .stTextArea textarea,
[data-baseweb="input"] input, [data-baseweb="textarea"] textarea,
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea {
  background: #0e0e22 !important; color: #ddd8ff !important;
  border: 1px solid rgba(124,58,237,0.35) !important;
  border-radius: 8px !important; caret-color: #a78bfa !important;
  font-size: 0.88rem !important;
}
input:focus, textarea:focus,
.stTextInput input:focus, .stTextArea textarea:focus {
  border-color: rgba(124,58,237,0.75) !important;
  box-shadow: 0 0 0 3px rgba(124,58,237,0.12) !important; outline: none !important;
}
input::placeholder, textarea::placeholder { color: rgba(180,170,240,0.3) !important; }

/* ── ALL labels ── */
label, p, [data-testid="stWidgetLabel"] p, .stSelectbox label,
.stTextInput label, .stTextArea label, .stSlider label,
.stCheckbox label, .stRadio label {
  color: rgba(180,170,240,0.65) !important; font-size: 0.78rem !important;
}

/* ── Selectbox ── */
.stSelectbox > div > div, [data-baseweb="select"] > div,
[data-testid="stSelectbox"] > div > div {
  background: #0e0e22 !important; border: 1px solid rgba(124,58,237,0.35) !important;
  border-radius: 8px !important; color: #ddd8ff !important;
}
[data-baseweb="select"] svg, [data-testid="stSelectbox"] svg { fill: #a78bfa !important; }
[data-baseweb="menu"], [role="listbox"],
ul[data-baseweb="menu"], [data-baseweb="popover"] { background: #12122a !important;
  border: 1px solid rgba(124,58,237,0.3) !important; border-radius: 8px !important; }
[role="option"] { color: #ddd8ff !important; font-size: 0.86rem !important; }
[role="option"]:hover, [aria-selected="true"] { background: rgba(124,58,237,0.22) !important; }

/* ── Slider ── */
[data-testid="stSlider"] > div > div > div { background: #7c3aed !important; }
[data-testid="stSlider"] [role="slider"] {
  background: #a78bfa !important; border: 2px solid #7c3aed !important;
}
[data-testid="stSlider"] p { color: rgba(167,139,250,0.8) !important; }

/* ── Checkbox ── */
[data-baseweb="checkbox"] span { border-color: rgba(124,58,237,0.5) !important; border-radius: 4px !important; }
[data-baseweb="checkbox"] [aria-checked="true"] span {
  background: #7c3aed !important; border-color: #7c3aed !important;
}
[data-baseweb="checkbox"] p { color: #c4c0f0 !important; font-size: 0.84rem !important; }

/* ── File uploader ── */
[data-testid="stFileUploadDropzone"] {
  background: #0e0e22 !important; border: 2px dashed rgba(124,58,237,0.4) !important;
  border-radius: 10px !important;
}
[data-testid="stFileUploadDropzone"] p, [data-testid="stFileUploadDropzone"] span,
[data-testid="stFileUploadDropzone"] small { color: #9490d0 !important; }

/* ── Expander ── */
[data-testid="stExpander"] {
  background: #0e0e22 !important; border: 1px solid rgba(124,58,237,0.2) !important;
  border-radius: 10px !important;
}
[data-testid="stExpander"] summary { color: #c4c0f0 !important; }
[data-testid="stExpander"] summary:hover { color: #a78bfa !important; }
[data-testid="stExpander"] svg { fill: #a78bfa !important; }
[data-testid="stExpander"] > div > div { border-top: 1px solid rgba(124,58,237,0.1) !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] iframe { background: #0e0e22 !important; border-radius: 8px; }

/* ── Primary Buttons ── */
.stButton > button {
  background: linear-gradient(135deg, #7c3aed 0%, #9333ea 100%) !important;
  color: white !important; border: none !important; border-radius: 8px !important;
  font-weight: 600 !important; font-size: 0.88rem !important;
  padding: 0.55rem 1.2rem !important; letter-spacing: 0.02em !important;
  transition: all 0.2s ease !important;
  box-shadow: 0 2px 12px rgba(124,58,237,0.3) !important;
}
.stButton > button:hover {
  transform: translateY(-1px) !important;
  box-shadow: 0 6px 24px rgba(124,58,237,0.45) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

/* ── Download Buttons ── */
[data-testid="stDownloadButton"] button {
  background: rgba(124,58,237,0.1) !important;
  border: 1px solid rgba(124,58,237,0.4) !important;
  color: #a78bfa !important; border-radius: 8px !important;
  font-weight: 600 !important; transition: all 0.2s !important;
}
[data-testid="stDownloadButton"] button:hover {
  background: rgba(124,58,237,0.25) !important; color: white !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
  background: rgba(255,255,255,0.03) !important; border-radius: 10px !important;
  padding: 5px !important; gap: 4px !important; border: 1px solid rgba(124,58,237,0.12);
}
.stTabs [data-baseweb="tab"] {
  border-radius: 7px !important; color: rgba(200,195,255,0.45) !important;
  font-size: 0.86rem !important; font-weight: 500 !important;
}
.stTabs [aria-selected="true"] {
  background: #7c3aed !important; color: white !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.2rem !important; }

/* ── Progress ── */
[data-testid="stProgressBar"] > div > div {
  background: linear-gradient(90deg, #7c3aed, #a855f7, #38bdf8) !important;
  border-radius: 4px !important;
}

/* ── Alerts ── */
[data-testid="stAlert"] {
  background: rgba(124,58,237,0.08) !important;
  border: 1px solid rgba(124,58,237,0.25) !important;
  border-radius: 8px !important; color: #c4c0f0 !important;
}
[data-testid="stAlert"][data-type="error"] {
  background: rgba(239,68,68,0.08) !important; border-color: rgba(239,68,68,0.3) !important;
}
[data-testid="stAlert"][data-type="success"] {
  background: rgba(16,185,129,0.08) !important; border-color: rgba(16,185,129,0.3) !important;
}

/* ── Divider ── */
hr { border-color: rgba(124,58,237,0.15) !important; margin: 1rem 0 !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #0e0e22; }
::-webkit-scrollbar-thumb { background: rgba(124,58,237,0.4); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: rgba(124,58,237,0.7); }

/* ════════════════════════ CUSTOM COMPONENTS ═══════════════════════════════ */
.hero {
  font-family: 'Syne', sans-serif !important; font-size: 2.55rem; font-weight: 800;
  letter-spacing: -0.03em; line-height: 1.08;
  background: linear-gradient(135deg, #a78bfa 0%, #7c3aed 40%, #38bdf8 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
}
.hero-sub { color: rgba(200,195,255,0.42); font-size: 0.9rem; margin: 0.2rem 0 1rem; line-height: 1.6; }

.stat-row { display: flex; gap: 0.5rem; flex-wrap: wrap; margin: 0.6rem 0; }
.stat-chip {
  background: rgba(124,58,237,0.08); border: 1px solid rgba(124,58,237,0.22);
  border-radius: 20px; padding: 4px 12px; font-size: 0.76rem; color: #a78bfa;
  font-weight: 500;
}
.stat-chip b { color: #c4c0f0; }

.panel {
  background: rgba(14,14,34,0.8); border: 1px solid rgba(124,58,237,0.12);
  border-radius: 12px; padding: 1.1rem; margin-bottom: 0.7rem;
}
.panel-accent {
  background: rgba(124,58,237,0.06); border: 1px solid rgba(124,58,237,0.2);
  border-radius: 12px; padding: 1rem; margin-bottom: 0.7rem;
}
.section-label {
  font-size: 0.67rem; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.14em; color: rgba(200,195,255,0.28); margin-bottom: 0.3rem;
}
.slide-card {
  background: rgba(14,14,34,0.7); border: 1px solid rgba(124,58,237,0.12);
  border-radius: 10px; padding: 0.82rem; margin-bottom: 0.38rem; transition: all 0.18s;
}
.slide-card:hover { border-color: rgba(167,139,250,0.4); transform: translateY(-1px); }
.slide-type { display: inline-block; background: rgba(124,58,237,0.2); color: #a78bfa;
  font-size: 0.6rem; padding: 1px 6px; border-radius: 4px; margin-bottom: 3px; font-weight: 600; text-transform: uppercase; }
.slide-num { color: rgba(167,139,250,0.65); font-size: 0.63rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; }
.slide-title { font-family: 'Syne', sans-serif; font-weight: 700; font-size: 0.85rem; margin: 2px 0 3px; color: #e8e6ff; }
.slide-body { color: rgba(200,195,255,0.42); font-size: 0.72rem; line-height: 1.5; }
.slide-ex { background: rgba(124,58,237,0.08); border-left: 2px solid rgba(167,139,250,0.5);
  padding: 3px 7px; font-size: 0.68rem; color: rgba(167,139,250,0.85); margin-top: 4px; border-radius: 0 5px 5px 0; }
.slide-stat { color: rgba(245,158,11,0.8); font-size: 0.68rem; margin-top: 3px; font-style: italic; }

.badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.72rem; font-weight: 500; }
.badge-ok  { background: rgba(16,185,129,0.12); color: #10b981; border: 1px solid rgba(16,185,129,0.25); }
.badge-err { background: rgba(239,68,68,0.12); color: #ef4444; border: 1px solid rgba(239,68,68,0.25); }
.badge-warn{ background: rgba(245,158,11,0.12); color: #f59e0b; border: 1px solid rgba(245,158,11,0.25); }

.tip { background: rgba(56,189,248,0.06); border: 1px solid rgba(56,189,248,0.2);
  border-radius: 8px; padding: 0.58rem 0.9rem; font-size: 0.78rem; color: rgba(200,195,255,0.65); margin: 0.35rem 0; }
.tip b { color: #38bdf8; }

.dl-row { display: flex; gap: 0.6rem; margin-top: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# ── State ──────────────────────────────────────────────────────────────────────
for k, v in [("slides",None),("pptx",None),("pdf",None),("charts",{}),("qtopic",None)]:
    if k not in st.session_state: st.session_state[k] = v

# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.3rem;'
                'background:linear-gradient(135deg,#a78bfa,#38bdf8);'
                '-webkit-background-clip:text;-webkit-text-fill-color:transparent;">'
                '🧠 NeuraSlide v5</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:rgba(200,195,255,.35);font-size:.72rem;margin-bottom:.5rem;">'
                'Professional AI Presentation Studio</div>', unsafe_allow_html=True)
    st.divider()

    st.markdown('<div class="section-label">🔌 AI Engine</div>', unsafe_allow_html=True)
    ok, _ = check_ollama()
    if ok:
        st.markdown('<span class="badge badge-ok">✓ Ollama Connected</span>', unsafe_allow_html=True)
        models = list_models()
        model = st.selectbox("Model", models or ["mistral"], label_visibility="collapsed")
    else:
        st.markdown('<span class="badge badge-err">✗ Ollama Offline</span>', unsafe_allow_html=True)
        st.markdown('<div style="background:rgba(239,68,68,.06);border:1px solid rgba(239,68,68,.18);'
                    'border-radius:8px;padding:.55rem;font-size:.73rem;margin-top:.35rem;color:#fca5a5;">'
                    'Run: <code>ollama serve</code><br>'
                    'Then: <code>ollama pull mistral</code></div>', unsafe_allow_html=True)
        model = "mistral"

    st.divider()
    st.markdown('<div class="section-label">🎨 Design</div>', unsafe_allow_html=True)
    theme = st.selectbox("Theme", list(THEMES.keys()), label_visibility="collapsed")

    st.markdown('<div class="section-label" style="margin-top:.5rem;">📊 Content</div>',
                unsafe_allow_html=True)
    n_slides   = st.slider("Number of slides", 5, 20, 10, label_visibility="visible")
    difficulty = st.selectbox("Content depth", ["Beginner","Intermediate","Advanced","Expert"])
    language   = st.selectbox("Language", ["English","Urdu","Arabic","French","Spanish","German","Chinese"])

    st.divider()
    st.markdown('<div class="section-label">⚙️ Options</div>', unsafe_allow_html=True)
    add_wm  = st.checkbox("🔐 Watermark text")
    wm_text = st.text_input("Watermark", "CONFIDENTIAL", label_visibility="collapsed") if add_wm else ""
    add_pdf  = st.checkbox("📄 PDF speaker notes", True)
    add_refs = st.checkbox("📚 References slide")
    add_imgs = st.checkbox("🖼️ Real photos (internet)", True)

    st.divider()
    st.markdown("""
    <div class="tip"><b>Tips for best results:</b><br>
    • 8–12 slides = best quality/speed<br>
    • Be specific: "ML for healthcare 2024"<br>
    • mistral = fastest & most reliable<br>
    • Uncheck 🖼️ if offline</div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="hero">AI Presentation Studio</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">10 professional layouts · Real Unsplash photos · Example boxes · '
            'Stat cards · Process flows · Comparison splits · Speaker notes · CSV data charts</div>',
            unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📝  Topic / Prompt", "📊  CSV / Excel Data", "✏️  Edit Slides"])

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — TOPIC
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns([3, 2], gap="large")

    with c1:
        st.markdown('<div class="section-label">Your topic or detailed prompt</div>', unsafe_allow_html=True)
        default_val = st.session_state.qtopic or ""
        topic = st.text_area("Topic", value=default_val,
            placeholder="Examples:\n• 'Machine Learning Basics for Students'\n"
                         "• 'Impact of Climate Change on Pakistan's Agriculture'\n"
                         "• 'Digital Marketing Strategies for E-commerce 2024'\n"
                         "• 'How Neural Networks Work — Visual Guide'",
            height=125, label_visibility="collapsed")

        st.markdown('<div class="section-label" style="margin-top:.6rem;">⚡ Quick Topics</div>',
                    unsafe_allow_html=True)
        topics = [
            "Machine Learning","Neural Networks","Blockchain","Digital Marketing",
            "Climate Change","Cybersecurity","Data Science","Space Exploration",
            "Python Programming","Startup Strategy","Cloud Computing","IoT",
            "Deep Learning","Quantum Computing","EV Technology","Renewable Energy",
        ]
        qc = st.columns(4)
        for i, q in enumerate(topics):
            if qc[i%4].button(q, key=f"qt{i}", use_container_width=True):
                st.session_state.qtopic = q; st.rerun()

    with c2:
        st.markdown(f"""<div class="panel-accent">
        <div style="font-family:Syne,sans-serif;font-weight:700;font-size:.82rem;margin-bottom:.5rem;color:#c4c0f0;">
        Current Configuration</div>
        <div class="stat-row">
          <span class="stat-chip">🎨 <b>{theme.split(' ',1)[-1][:16]}</b></span>
          <span class="stat-chip">📊 <b>{n_slides} slides</b></span>
        </div>
        <div class="stat-row">
          <span class="stat-chip">🎯 <b>{difficulty}</b></span>
          <span class="stat-chip">🌍 <b>{language}</b></span>
        </div>
        <div class="stat-row">
          <span class="stat-chip">🖼️ <b>{'Photos on' if add_imgs else 'Gradients'}</b></span>
          <span class="stat-chip">📄 <b>{'PDF ✓' if add_pdf else 'No PDF'}</b></span>
          {'<span class="stat-chip">🔐 <b>'+wm_text+'</b></span>' if add_wm and wm_text else ''}
        </div></div>""", unsafe_allow_html=True)

        st.markdown("""<div class="panel">
        <div style="font-family:Syne,sans-serif;font-weight:700;font-size:.8rem;color:#a78bfa;margin-bottom:.45rem;">
        📐 Slide Layouts Used</div>
        <div style="font-size:.74rem;color:rgba(200,195,255,.52);line-height:1.95;">
        🎬 Full-bleed title + image<br>
        📋 Agenda with numbered items<br>
        ◧ Text + photo side panel<br>
        🖼️ Banner photo + 2-col bullets<br>
        📊 Big stat hero (42pt number)<br>
        💬 Expert pull-quote<br>
        ⚖️ Split/comparison columns<br>
        🔄 Numbered process steps<br>
        🌅 Big visual image slide<br>
        🃏 4-card grid layout
        </div></div>""", unsafe_allow_html=True)

    st.markdown("---")
    gen_col, _ = st.columns([2,1])
    with gen_col:
        generate_clicked = st.button("🚀  Generate Professional Presentation",
                                      use_container_width=True, type="primary")

    if generate_clicked:
        if not topic.strip():
            st.error("⚠️ Please enter a topic or click a Quick Topic button.")
        elif not ok:
            st.error("⚠️ Ollama is not running. Start it with: `ollama serve`")
        else:
            prog = st.progress(0); stat = st.empty()
            try:
                stat.markdown("**🧠 Generating slide content with AI...** *(may take 30–120 sec)*")
                prog.progress(5)

                def _prog_cb(done, total_slides):
                    p = int(5 + (done/total_slides)*50)
                    prog.progress(min(p,55))
                    stat.markdown(f"**✍️ Generated {done}/{total_slides} slides...**")

                slides = generate_slides(topic, model, n_slides, difficulty, language,
                                         add_refs, progress_cb=_prog_cb)
                prog.progress(58)

                stat.markdown(f"**🖼️ Building {len(slides)}-slide presentation with images...**")
                pptx = build_presentation(slides, theme, wm_text, topic, fetch_images=add_imgs)
                st.session_state.slides = slides
                st.session_state.pptx   = pptx
                st.session_state.qtopic = None
                prog.progress(85)

                if add_pdf:
                    stat.markdown("**📄 Generating PDF speaker notes...**")
                    st.session_state.pdf = build_pdf(slides, topic, theme, wm_text)

                prog.progress(100); stat.empty(); prog.empty()
                st.success(f"✅ {len(slides)}-slide presentation ready! Scroll down to preview and download.")
                st.balloons()
            except Exception as e:
                stat.empty(); prog.empty()
                st.error(f"❌ Error: {e}")
                st.info("💡 Try: fewer slides (6–8), or smaller model like `mistral` or `llama3.2:1b`")

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — CSV
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-label">Upload your data file</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv","xlsx","xls"],
                                label_visibility="collapsed")
    if uploaded:
        try:
            df = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
            num_c = df.select_dtypes(include="number").columns.tolist()
            cat_c = df.select_dtypes(include=["object","category"]).columns.tolist()

            st.markdown(f"""<div class="panel">
            <div class="stat-row">
              <span class="stat-chip">📁 <b>{uploaded.name[:28]}</b></span>
              <span class="stat-chip">📏 <b>{df.shape[0]:,} rows</b></span>
              <span class="stat-chip">📊 <b>{df.shape[1]} cols</b></span>
              <span class="stat-chip">🔢 <b>{len(num_c)} numeric</b></span>
              <span class="stat-chip">🔤 <b>{len(cat_c)} text</b></span>
            </div></div>""", unsafe_allow_html=True)

            with st.expander("👀 Preview first 8 rows"):
                st.dataframe(df.head(8), use_container_width=True)

            ca, cb = st.columns(2)
            with ca: csv_title = st.text_input("Presentation title", f"Analysis — {uploaded.name[:30]}")
            with cb: n_charts  = st.slider("Chart slides", 1, min(9, max(2, len(num_c)+1)), min(5, max(2, len(num_c))))

            if st.button("🚀  Analyze & Build Data Presentation", use_container_width=True):
                if not ok: st.error("Ollama not running!")
                else:
                    prog=st.progress(0); stat=st.empty(); tmpdir=tempfile.mkdtemp()
                    try:
                        stat.markdown("**📊 Analyzing dataset...**"); prog.progress(8)
                        analysis = analyze_dataframe(df)
                        stat.markdown("**📈 Generating charts...**"); prog.progress(18)
                        charts = generate_all_charts(df, tmpdir)
                        stat.markdown("**🧠 AI generating insights...**"); prog.progress(30)
                        slides = analyze_csv_with_ai(analysis, model, csv_title, n_slides)

                        # Map charts → slides
                        chart_map={}; used=set()
                        for idx, s in enumerate(slides):
                            hint = s.get("chart_hint","none")
                            if hint and hint != "none":
                                p = pick_chart_for_slide({k:v for k,v in charts.items() if k not in used}, hint)
                                if p:
                                    chart_map[idx+1] = p
                                    for k,v in charts.items():
                                        if v==p: used.add(k); break
                        extra = {k:v for k,v in charts.items() if k not in used}
                        for i,(cn,cp) in enumerate(list(extra.items())[:n_charts]):
                            chart_map[len(slides)+i+1] = cp

                        prog.progress(62)
                        stat.markdown("**🎨 Building presentation...**")
                        pptx = build_presentation(slides, theme, wm_text, csv_title, chart_map, add_imgs)
                        st.session_state.slides=slides; st.session_state.pptx=pptx
                        st.session_state.charts=charts
                        if add_pdf: st.session_state.pdf=build_pdf(slides,csv_title,theme,wm_text)
                        prog.progress(100); stat.empty(); prog.empty()
                        st.success(f"✅ {len(slides)} insight slides + {len(chart_map)} chart slides generated!")
                        st.balloons()
                        if charts:
                            st.markdown('<div class="section-label" style="margin-top:1rem;">📈 All Charts Generated</div>',unsafe_allow_html=True)
                            cc=st.columns(3)
                            for i,(n,p) in enumerate(charts.items()):
                                cc[i%3].image(p, caption=n.upper(), use_container_width=True)
                    except Exception as e:
                        stat.empty(); prog.empty(); st.error(f"❌ {e}")
        except Exception as e:
            st.error(f"Could not read file: {e}")
    else:
        st.markdown("""<div class="panel" style="text-align:center;padding:2.5rem 1.5rem;">
        <div style="font-size:2.2rem;margin-bottom:.5rem;">📊</div>
        <div style="font-family:Syne,sans-serif;font-weight:700;font-size:.95rem;margin-bottom:.3rem;color:#c4c0f0;">
        Upload CSV or Excel</div>
        <div style="color:rgba(200,195,255,.38);font-size:.8rem;line-height:1.6;">
        AI reads your actual column names and real numbers,<br>
        generates data-driven insights, and embeds 9 chart types into slides.
        </div></div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — EDIT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    if not st.session_state.slides:
        st.markdown("""<div class="panel" style="text-align:center;padding:2.5rem;">
        <div style="font-size:2rem;">✏️</div>
        <div style="font-family:Syne,sans-serif;font-weight:700;font-size:.92rem;margin:.3rem 0;color:#c4c0f0;">
        No slides yet</div>
        <div style="color:rgba(200,195,255,.38);font-size:.8rem;">
        Generate a presentation first, then edit every slide field here.</div>
        </div>""", unsafe_allow_html=True)
    else:
        slides = st.session_state.slides
        st.markdown(f'<div style="font-family:Syne,sans-serif;font-weight:700;font-size:.92rem;'
                    f'margin-bottom:.6rem;color:#c4c0f0;">✏️ {len(slides)} Slides — expand any to edit</div>',
                    unsafe_allow_html=True)
        types = ["title","agenda","content","statistic","quote","split","comparison","process",
                 "timeline","bigimage","conclusion","references"]
        edited = []
        for i, s in enumerate(slides):
            label = f"Slide {i+1} [{s.get('type','content')}] — {s.get('title','')[:48]}"
            with st.expander(label, expanded=(i==0)):
                r1, r2 = st.columns(2)
                with r1:
                    et  = st.text_input("Title",    value=s.get("title",""),    key=f"et{i}")
                    esb = st.text_input("Subtitle", value=s.get("subtitle",""), key=f"esb{i}")
                    ety = st.selectbox("Slide type", types,
                          index=types.index(s.get("type","content")) if s.get("type","content") in types else 2,
                          key=f"ety{i}")
                    ek  = st.text_input("Image keyword", value=s.get("image_keyword",""), key=f"ek{i}")
                with r2:
                    est = st.text_input("Statistic", value=s.get("stat",""),  key=f"est{i}")
                    en  = st.text_input("Notes",     value=s.get("notes",""), key=f"en{i}")
                    eq  = st.text_input("Quote",     value=s.get("quote",""), key=f"eq{i}")
                ec  = st.text_area("Content (bullets)", value=s.get("content",""),       height=80, key=f"ec{i}")
                ecl = st.text_area("Left column",       value=s.get("content_left",""),  height=65, key=f"ecl{i}")
                ecr = st.text_area("Right column",      value=s.get("content_right",""), height=65, key=f"ecr{i}")
                esp = st.text_area("Steps",             value=s.get("steps",""),         height=65, key=f"esp{i}")
                ee  = st.text_area("Example",           value=s.get("example",""),       height=65, key=f"ee{i}")
                edited.append({"type":ety,"title":et,"subtitle":esb,"content":ec,
                                "content_left":ecl,"content_right":ecr,"steps":esp,
                                "example":ee,"stat":est,"quote":eq,"notes":en,"image_keyword":ek})

        st.markdown("---")
        if st.button("💾  Save Edits & Rebuild Presentation", use_container_width=True):
            st.session_state.slides = edited
            prog=st.progress(0); stat=st.empty()
            stat.markdown("**🔨 Rebuilding...**"); prog.progress(25)
            st.session_state.pptx = build_presentation(edited, theme, wm_text,
                                                        "Edited Presentation", fetch_images=add_imgs)
            prog.progress(80)
            if add_pdf:
                st.session_state.pdf = build_pdf(edited,"Edited Presentation",theme,wm_text)
            prog.progress(100); stat.empty(); prog.empty()
            st.success("✅ Presentation rebuilt with your edits!")

# ═══════════════════════════════════════════════════════════════════════════════
#  PREVIEW + DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.slides:
    st.markdown("---")
    pc1, pc2 = st.columns([2,1])
    with pc1:
        st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.1rem;'
                    'background:linear-gradient(135deg,#a78bfa,#38bdf8);'
                    '-webkit-background-clip:text;-webkit-text-fill-color:transparent;">'
                    '📊 Slide Preview</div>', unsafe_allow_html=True)
    with pc2:
        st.markdown(f'<div style="color:rgba(200,195,255,.45);font-size:.78rem;text-align:right;'
                    f'padding-top:4px;">{len(st.session_state.slides)} slides generated</div>',
                    unsafe_allow_html=True)

    prev_cols = st.columns(3)
    for i, s in enumerate(st.session_state.slides):
        with prev_cols[i%3]:
            ex = f'<div class="slide-ex">📌 {s["example"][:68]}...</div>' if s.get("example") else ""
            st_ = f'<div class="slide-stat">📊 {s["stat"][:58]}</div>' if s.get("stat") else ""
            img = (f'<div style="font-size:.62rem;color:rgba(124,58,237,.55);margin-top:3px;">'
                   f'🖼️ {s.get("image_keyword","")}</div>') if s.get("image_keyword") else ""
            st.markdown(f"""<div class="slide-card">
            <span class="slide-type">{s.get('type','content')}</span>
            <div class="slide-num">SLIDE {i+1}</div>
            <div class="slide-title">{s.get('title','')[:54]}</div>
            <div class="slide-body">{str(s.get('content',''))[:82]}{'...' if len(str(s.get('content','')))>82 else ''}</div>
            {ex}{st_}{img}</div>""", unsafe_allow_html=True)

    # Downloads
    st.markdown("---")
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:700;font-size:.95rem;'
                'margin-bottom:.6rem;color:#c4c0f0;">📥 Download Files</div>',
                unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        if st.session_state.pptx:
            st.download_button(
                "⬇️  PowerPoint (.pptx)",
                st.session_state.pptx, "neuroslide.pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
    with dl2:
        if st.session_state.pdf:
            st.download_button(
                "⬇️  PDF Speaker Notes",
                st.session_state.pdf, "neuroslide_notes.pdf",
                "application/pdf", use_container_width=True)
    with dl3:
        if st.session_state.slides:
            st.download_button(
                "⬇️  Raw JSON Data",
                json.dumps(st.session_state.slides, indent=2, ensure_ascii=False),
                "neuroslide_data.json", "application/json",
                use_container_width=True)
