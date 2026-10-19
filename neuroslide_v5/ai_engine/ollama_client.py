import requests, json, re, time

OLLAMA_URL = "http://localhost:11434"

# ── Connection ────────────────────────────────────────────────────────────────
def check_ollama():
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
        return (True, "Connected") if r.status_code == 200 else (False, "Bad response")
    except Exception as e:
        return False, str(e)

def list_models():
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
        return [m["name"] for m in r.json().get("models", [])]
    except:
        return []

# ── Streaming call ─────────────────────────────────────────────────────────────
def _stream(prompt, model, system="", timeout=600):
    payload = {
        "model": model, "prompt": prompt, "stream": True,
        "options": {"temperature": 0.4, "num_predict": 8000, "top_p": 0.85, "repeat_penalty": 1.1}
    }
    if system:
        payload["system"] = system
    out = []
    with requests.post(f"{OLLAMA_URL}/api/generate", json=payload,
                       stream=True, timeout=(10, timeout)) as r:
        r.raise_for_status()
        for line in r.iter_lines():
            if line:
                try:
                    c = json.loads(line)
                    out.append(c.get("response", ""))
                    if c.get("done"): break
                except: pass
    return "".join(out)

# ── Bulletproof JSON repair ────────────────────────────────────────────────────
def _to_bullet_str(val):
    if isinstance(val, list):
        parts = []
        for item in val:
            s = str(item).strip().strip('"\'').lstrip("•●▪▸▹-–— \u2022\u25cf\u25aa\u2023").strip()
            if s: parts.append(f"• {s}")
        return "\n".join(parts)
    if isinstance(val, str):
        return val
    return str(val) if val else ""

def _repair(text):
    # 1. Strip fences & normalize quotes
    text = re.sub(r"```(?:json)?", "", text).strip().replace("```", "")
    text = (text.replace("\u2018","'").replace("\u2019","'")
                .replace("\u201c",'"').replace("\u201d",'"')
                .replace("\u2022","•").replace("\u25cf","•"))

    # 2. Convert any [ ... ] field values to bullet strings BEFORE json parse
    def array_to_str(m):
        key, inner = m.group(1), m.group(2)
        items = re.findall(r'(?:•\s*)?["\']([^"\']{3,})["\']', inner)
        if not items:
            items = re.findall(r'•\s*([^\n•"\']{3,})', inner)
        if not items:
            items = re.findall(r'"([^"]{3,})"', inner)
        bullets = "\n".join(f"• {i.strip()}" for i in items if i.strip())
        return f'"{key}": "{bullets}"'
    text = re.sub(r'"(\w+)"\s*:\s*\[([^\]]{0,800})\]', array_to_str, text, flags=re.DOTALL)

    # 3. Remove trailing commas
    text = re.sub(r",\s*([\]}])", r"\1", text)

    # 4. Escape unescaped newlines/tabs inside string values
    def fix_nl(t):
        res=[]; in_s=False; esc=False
        for ch in t:
            if esc: res.append(ch); esc=False; continue
            if ch == '\\': esc=True; res.append(ch); continue
            if ch == '"': in_s = not in_s
            if in_s and ch == '\n': res.append('\\n')
            elif in_s and ch == '\t': res.append(' ')
            else: res.append(ch)
        return ''.join(res)
    text = fix_nl(text)

    # 5. Try to parse
    s, e = text.find("["), text.rfind("]")
    if s != -1 and e != -1 and e > s:
        for attempt in [text[s:e+1], re.sub(r",\s*([\]}])",r"\1",text[s:e+1])]:
            try:
                data = json.loads(attempt)
                if isinstance(data, list) and data:
                    for sl in data:
                        if isinstance(sl, dict):
                            for k in ["content","content_left","content_right","steps",
                                      "example","stat","quote","notes","image_keyword",
                                      "subtitle","title"]:
                                if k in sl: sl[k] = _to_bullet_str(sl[k])
                    return [s for s in data if isinstance(s,dict) and s.get("title")]
            except: pass

    # 6. Object-by-object fallback
    slides = []
    for obj in re.findall(r'\{[^{}]{20,}\}', text, re.DOTALL):
        try:
            o = json.loads(obj)
            for k in ["content","example","stat","quote","notes"]:
                if k in o: o[k] = _to_bullet_str(o[k])
            if o.get("title"): slides.append(o)
        except: pass
    if slides: return slides
    raise ValueError(f"Could not parse AI response.\nPreview: {text[:400]}")

# ── System prompt (injected once) ─────────────────────────────────────────────
_SYS = (
    "You are a professional presentation content generator. "
    "Output ONLY a valid JSON array — no markdown, no code fences, no prose. "
    "Every field must be a plain STRING. "
    "The 'content' field must be: '• Point one\\n• Point two\\n• Point three' "
    "— NEVER a JSON array. Use \\n between bullets, not actual newlines."
)

# ── Slide type guide injected into prompt ──────────────────────────────────────
_TYPE_GUIDE = """
Slide types and when to use them:
- "title"      : First slide only. Has subtitle and stat.
- "agenda"     : Second slide. Lists what will be covered (content = topic list).
- "content"    : Standard bullet slide. Fill content + example + stat.
- "statistic"  : Data-heavy slide. Must have a compelling stat with real number.
- "quote"      : Expert insight. Must have a real attributed quote.
- "split"      : Two-sided comparison. Fill content_left AND content_right.
- "process"    : Step-by-step flow. Fill the steps field.
- "timeline"   : Chronological events. Fill the steps field.
- "bigimage"   : Visual impact slide. Minimal text, strong image keyword.
- "conclusion" : Final slide. Numbered takeaways in content.
"""

def _build_prompt(topic, n, diff, lang, add_refs, is_cont=False, prev=None):
    ref = "\n- One final slide: type='references'" if add_refs else ""
    prev_note = f"\nAlready generated: {prev}. Continue from where you left off, no repetition." if prev else ""
    intro = (f"Continue the presentation on '{topic}'.{prev_note}\nGenerate {n} MORE slides."
             if is_cont else f"Create a {n}-slide professional presentation on: \"{topic}\"")
    return f"""{intro}
Level: {diff} | Language: {lang}

{_TYPE_GUIDE}

Each slide is a JSON object with ALL these string fields:
{{
  "type":          "title|agenda|content|statistic|quote|split|process|timeline|bigimage|conclusion|references",
  "title":         "Compelling slide heading",
  "subtitle":      "Supporting line (used in title/section slides)",
  "content":       "• Point one\\n• Point two\\n• Point three\\n• Point four",
  "content_left":  "• Left col point 1\\n• Left col point 2  (split slides only)",
  "content_right": "• Right col point 1\\n• Right col point 2  (split slides only)",
  "steps":         "Step 1: action\\nStep 2: action\\nStep 3: action  (process/timeline only)",
  "example":       "Concrete real-world example with specific names/numbers (2-3 sentences)",
  "stat":          "Statistic with real number + source e.g. '73% of firms use AI (McKinsey 2024)'",
  "quote":         "Expert quote with name e.g. 'AI is electricity. — Andrew Ng'",
  "notes":         "One presenter note sentence",
  "image_keyword": "3 specific English words for Unsplash photo"
}}

CRITICAL RULES:
1. content must be "• text\\n• text" STRING — NEVER a JSON array
2. Every slide needs a strong specific image_keyword
3. stat must contain a real number and a source
4. example must be concrete and specific with company names or real data
5. For split slides fill BOTH content_left and content_right
6. For process/timeline slides fill the steps field
7. Use a VARIED MIX of types — avoid repeating the same type consecutively
8. Make content SMART and DETAILED — not generic fluff{ref}

Output ONLY the JSON array starting with [ and ending with ]"""


def generate_slides(topic, model, slide_count=10, difficulty="Intermediate",
                    language="English", add_refs=False, progress_cb=None):
    """Generate slides in batches of 6 with retries."""
    BATCH = 6
    MAX_RETRIES = 2
    all_slides = []
    remaining = slide_count

    while remaining > 0:
        b = min(BATCH, remaining)
        is_last = remaining <= BATCH
        prev_titles = [s.get("title","") for s in all_slides] if all_slides else None

        prompt = _build_prompt(topic, b, difficulty, language,
                               add_refs and is_last, bool(all_slides), prev_titles)

        for attempt in range(MAX_RETRIES + 1):
            try:
                raw = _stream(prompt, model, _SYS, timeout=600)
                batch = _repair(raw)
                if batch:
                    all_slides.extend(batch)
                    if progress_cb:
                        progress_cb(len(all_slides), slide_count)
                    remaining -= len(batch)
                    break
            except Exception as e:
                if attempt == MAX_RETRIES:
                    raise RuntimeError(f"Failed after {MAX_RETRIES+1} attempts: {e}")
                time.sleep(2)

    # Post-process: ensure conclusion is last
    if all_slides and all_slides[-1].get("type") not in ("conclusion","references"):
        all_slides[-1]["type"] = "conclusion"

    # Ensure type="title" is first
    if all_slides and all_slides[0].get("type") != "title":
        all_slides[0]["type"] = "title"

    return all_slides[:slide_count]


def analyze_csv_with_ai(analysis, model, title, slide_count=8):
    import json as _json
    prompt = f"""Analyze the dataset "{title}" and create {slide_count} data-driven slides.

ACTUAL DATA:
- Shape: {analysis.get('shape','')}
- Columns: {analysis.get('columns',[])}
- Statistics: {_json.dumps(analysis.get('numeric_summary',{}))[:1000]}
- Top values: {_json.dumps(analysis.get('top_values',{}))[:700]}
- Correlations: {_json.dumps(analysis.get('correlations',{}))[:400]}

Requirements:
- Use REAL column names from the data above
- Use REAL numbers from the statistics above
- Every stat field must reference actual numbers from the dataset
- chart_hint: one of "bar","line","pie","scatter","histogram","box","area","heatmap","none"

Slide structure:
- Slide 1: type="title" for this dataset
- Slide 2: type="content" — dataset overview (real shape, columns, data types)
- Slides 3-{slide_count-2}: individual insights using real data
- Slide {slide_count-1}: type="statistic" — most compelling finding
- Slide {slide_count}: type="conclusion" — recommendations

Each slide: type, title, subtitle, content (STRING bullets), example, stat (REAL number), notes, image_keyword, chart_hint

CRITICAL: content = "• text\\n• text" STRING not array.
Output ONLY the JSON array."""

    raw = _stream(prompt, model, _SYS, timeout=480)
    return _repair(raw)
