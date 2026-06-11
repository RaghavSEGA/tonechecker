"""
narrative_qa.py — SNAKE Narrative Integrity Validator
Multi-agent QA for narrative-driven game scripts.
Two separate AWS credential sets:
  - SES  (OTP email):  AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY / AWS_SES_REGION
  - Bedrock (LLM):     AWS_BEDROCK_ACCESS_KEY_ID / AWS_BEDROCK_SECRET_ACCESS_KEY / AWS_BEDROCK_REGION
"""
from __future__ import annotations
import streamlit as st
import json, re, hashlib, hmac, time, os, io, csv, textwrap, random, base64
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from typing import Any

try:
    from anthropic import AnthropicBedrock, Anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

try:
    import boto3; HAS_BOTO3 = True
except ImportError:
    HAS_BOTO3 = False

try:
    import pypdf; HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from docx import Document as DocxDocument; HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl; HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(page_title="SNAKE Narrative QA", page_icon="\U0001F40D",
                    layout="wide", initial_sidebar_state="expanded")

# ── SEGA dark-theme CSS ─────────────────────────────────────────────────────
st.markdown("""
<style>
:root{--bg0:#040A1C;--bg1:#0D1B2A;--bg2:#0A1628;--bdr:#1A2A4A;--txt:#E0E0E0;
--mut:#8899BB;--blu:#00AADD;--gld:#F5C218;--grn:#00BB66;--red:#CC2244;}
html,body,[data-testid="stAppViewContainer"],[data-testid="stApp"]{background:var(--bg0)!important;color:var(--txt);}
[data-testid="stSidebar"]{background:var(--bg2)!important;}
[data-testid="stSidebar"] *{color:var(--txt)!important;}
header[data-testid="stHeader"]{background:transparent!important;}footer{visibility:hidden;}
h1,h2,h3,h4,h5,h6{color:#fff!important;}
.nqa-card{background:var(--bg1);border:1px solid var(--bdr);border-radius:12px;padding:1.5rem;margin-bottom:1rem;}
.nqa-score-card{background:var(--bg1);border:1px solid var(--bdr);border-radius:12px;padding:1.2rem;text-align:center;}
.nqa-score-card .sv{font-size:2.4rem;font-weight:800;}.nqa-score-card .sl{font-size:.85rem;color:var(--mut);margin-top:4px;}
.sev-c{display:inline-block;padding:3px 10px;border-radius:12px;background:var(--red);color:#fff;font-size:.8rem;font-weight:600;}
.sev-w{display:inline-block;padding:3px 10px;border-radius:12px;background:var(--gld);color:#000;font-size:.8rem;font-weight:600;}
.sev-i{display:inline-block;padding:3px 10px;border-radius:12px;background:var(--blu);color:#fff;font-size:.8rem;font-weight:600;}
.ag-row{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--bdr);}
.ag-row:last-child{border-bottom:none;}
.ag-dot{width:14px;height:14px;border-radius:50%;flex-shrink:0;}
.dot-pending{background:#333;}.dot-running{background:var(--gld);animation:pulse 1s infinite;}
.dot-complete{background:var(--grn);}.dot-error{background:var(--red);}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
.ag-name{font-weight:600;min-width:170px;}.ag-st{color:var(--mut);font-size:.85rem;}
.char-tag{display:inline-block;padding:4px 14px;border-radius:20px;background:var(--bg1);border:1px solid var(--blu);margin:4px;font-size:.85rem;color:var(--blu);font-weight:500;}
.aline{padding:6px 12px;margin:3px 0;border-radius:6px;border-left:3px solid transparent;font-family:monospace;font-size:.9rem;line-height:1.5;}
.al-ok{border-left-color:var(--grn);background:rgba(0,187,102,.04);}
.al-c{border-left-color:var(--red);background:rgba(204,34,68,.08);}
.al-w{border-left-color:var(--gld);background:rgba(245,194,24,.06);}
.al-i{border-left-color:var(--blu);background:rgba(0,170,221,.06);}
.hero-t{font-size:2.2rem;font-weight:700;color:#fff;margin-bottom:0;}
.hero-s{font-size:1rem;color:var(--mut);margin-top:2px;}
.stButton>button{background:var(--blu)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;transition:all .2s!important;}
.stButton>button:hover{background:#0088BB!important;transform:translateY(-1px);}
textarea{background:var(--bg1)!important;color:var(--txt)!important;border:1px solid var(--bdr)!important;border-radius:8px!important;}
[data-testid="stExpander"]{border:1px solid var(--bdr)!important;border-radius:10px!important;background:var(--bg1)!important;}
.fl-c{background:rgba(204,34,68,.55);color:#fff;padding:0 3px;border-radius:3px;font-weight:600;}
.fl-w{background:rgba(245,194,24,.45);color:#fff;padding:0 3px;border-radius:3px;font-weight:600;}
.fl-i{background:rgba(0,170,221,.40);color:#fff;padding:0 3px;border-radius:3px;font-weight:600;}
.line-quote{font-family:monospace;font-size:.8rem;background:rgba(255,255,255,.04);padding:4px 8px;border-radius:4px;margin-top:3px;white-space:pre-wrap;word-break:break-word;}
.ag-chip{display:inline-block;padding:1px 9px;border-radius:10px;font-size:.72rem;font-weight:700;letter-spacing:.02em;}

/* ── tabs: quiet labels, blue active, no stock red underline ─────────────── */
.stTabs [data-baseweb="tab-list"]{gap:2px;border-bottom:1px solid var(--bdr);}
.stTabs [data-baseweb="tab"]{color:var(--mut);font-weight:600;padding:8px 18px;border-radius:8px 8px 0 0;background:transparent;}
.stTabs [data-baseweb="tab"]:hover{color:var(--txt);background:rgba(0,170,221,.06);}
.stTabs [aria-selected="true"]{color:var(--blu)!important;}
.stTabs [data-baseweb="tab-highlight"]{background-color:var(--blu);height:2px;}
.stTabs [data-baseweb="tab-border"]{background-color:var(--bdr);}

/* ── selects & popover menus (sample picker etc.) ────────────────────────── */
[data-baseweb="select"] > div{background:var(--bg1)!important;border-color:var(--bdr)!important;color:var(--txt)!important;border-radius:8px!important;}
[data-baseweb="select"] > div:hover{border-color:var(--blu)!important;}
[data-baseweb="popover"] [data-baseweb="menu"],[data-baseweb="popover"] ul{background:var(--bg1)!important;border:1px solid var(--bdr)!important;border-radius:8px!important;}
li[role="option"]{background:var(--bg1)!important;color:var(--txt)!important;}
li[role="option"]:hover,li[aria-selected="true"][role="option"]{background:rgba(0,170,221,.12)!important;color:#fff!important;}
[data-baseweb="tag"]{background:rgba(0,170,221,.18)!important;border:1px solid rgba(0,170,221,.4)!important;color:var(--blu)!important;border-radius:10px!important;}
[data-baseweb="tag"] span{color:var(--blu)!important;}

/* ── pills (filter chips) ─────────────────────────────────────────────────── */
.stPills [data-testid="stBaseButton-pills"],.stPills [data-testid="stBaseButton-pillsActive"]{border-radius:14px!important;font-size:.8rem!important;padding:2px 12px!important;}
.stPills [data-testid="stBaseButton-pills"]{background:var(--bg1)!important;border:1px solid var(--bdr)!important;color:var(--mut)!important;}
.stPills [data-testid="stBaseButton-pills"]:hover{border-color:var(--blu)!important;color:var(--txt)!important;}
.stPills [data-testid="stBaseButton-pillsActive"]{background:rgba(0,170,221,.16)!important;border:1px solid var(--blu)!important;color:var(--blu)!important;}

/* ── file uploader dropzone ──────────────────────────────────────────────── */
[data-testid="stFileUploaderDropzone"]{background:var(--bg1)!important;border:1px dashed var(--bdr)!important;border-radius:10px!important;}
[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--blu)!important;}
[data-testid="stFileUploaderDropzone"] small,[data-testid="stFileUploaderDropzone"] span{color:var(--mut)!important;}

/* ── download buttons: outline style to read as secondary actions ────────── */
.stDownloadButton>button{background:transparent!important;border:1px solid var(--blu)!important;color:var(--blu)!important;border-radius:8px!important;font-weight:600!important;}
.stDownloadButton>button:hover{background:rgba(0,170,221,.12)!important;transform:translateY(-1px);}

/* ── score cards: hover lift + severity-tinted top edge via inline color ─── */
.nqa-score-card{transition:transform .15s,border-color .15s;border-top:2px solid var(--bdr);}
.nqa-score-card:hover{transform:translateY(-2px);border-color:var(--blu);}

/* ── hero: gradient rule beneath the title, quiet otherwise ──────────────── */
.hero-t{letter-spacing:.01em;}
.hero-s{padding-bottom:10px;border-bottom:2px solid transparent;
  border-image:linear-gradient(90deg,var(--blu),var(--gld) 45%,transparent 70%) 1;}

/* ── widget labels: muted, consistent ────────────────────────────────────── */
[data-testid="stWidgetLabel"] p{color:var(--mut)!important;font-size:.82rem!important;font-weight:600!important;}

/* ── expander hover ──────────────────────────────────────────────────────── */
[data-testid="stExpander"]:hover{border-color:var(--blu)!important;}
[data-testid="stExpander"] summary{color:var(--txt)!important;font-weight:600;}
[data-testid="stExpander"] summary:hover{color:var(--blu)!important;}

/* ── inputs: blue focus instead of stock red ─────────────────────────────── */
textarea:focus,input:focus{border-color:var(--blu)!important;box-shadow:0 0 0 1px var(--blu)!important;}
:focus-visible{outline:2px solid var(--blu)!important;outline-offset:2px;}

/* ── dark scrollbars ─────────────────────────────────────────────────────── */
*::-webkit-scrollbar{width:10px;height:10px;}
*::-webkit-scrollbar-track{background:var(--bg0);}
*::-webkit-scrollbar-thumb{background:var(--bdr);border-radius:5px;}
*::-webkit-scrollbar-thumb:hover{background:var(--blu);}

/* ── reduced motion ──────────────────────────────────────────────────────── */
@media (prefers-reduced-motion: reduce){
  *,*::before,*::after{animation-duration:.01ms!important;transition-duration:.01ms!important;}
}
</style>""", unsafe_allow_html=True)


# ── SNAKE style-guide (full canonical source of truth) ──────────────────────
STYLE_GUIDE_TEXT = """\
SNAKE: Style Guide — For Production & Writing Teams

VOICE & PROSE
SNAKE is prestige dramatic thriller in a fighting game. On the tonal spectrum, the benchmark is the HBO Watchmen TV series and The Batman rather than Mortal Kombat: grounded characters, social texture, earned emotion. Every line reveals character, advances plot, or both. Nobody speaks in exposition. The complex characters, their relationships and heart-felt emotion drive the narrative, which features a rich layer of compelling moral ambiguity.

Register: Contemporary and colloquial. Characters interrupt, deflect, and lie. Profanity is earned — when it lands, it means something. Frequency varies by character; do not flatten everyone to the same register.
Prose style (VO/narration): Direct, economical, present-tense observation. No flourishes.
Avoid throughout: over-explanation, self-pity, characters stating their own themes, on the nose or cliched dialogue.
Language and translation rules: Do not soften profanity — match the exact register. Preserve incomplete sentences and interruptions. Character speech rhythms are not interchangeable. Humor is character-specific. Regional Spanish variants matter (Cielo is Paraguayan Spanish, not Castilian or Mexican).

PROTAGONIST VOICES
CIELO SALINAS-WALKER — Physical, direct, emotionally contained. Humor is self-deprecating and delivered flat. Speaks in short declaratives. Code-switches into Paraguayan Spanish naturally ("abu," "carajo," "Quique"). Uses understatement as armor. Does not philosophize — gives simple answers and means them. Does not ask for help easily. His anger is a gearshift, not a boil.
EZECHIEL BROOKS — Low-affect, steady, quietly sardonic. British. Spends words carefully. Delivers menace with perfect calm: "Been a pleasure, mate." Does not unpack his own trauma. Does not preach. His one full emotional eruption ("BECAUSE YOU'D END UP FUCKING DEAD!") lands only because nothing like it has come before — this character cannot raise his voice prior to that moment. Preserve British phrasing.
FORTUNA DE LA ROSSA — Precise, dry, always two moves ahead. Most architecturally constructed voice in the game — her sentences are built, not improvised. Inner-voice observation mode is clinical notation: "Touching the face: a classic tell." Does not gloat. Does not explain her reasoning until she chooses to. Emotional exposure is rare and costly. Her confrontation with Rebecca must preserve both women's dignity.
STELLA BRIDGES — Upbeat surface, genuine steel underneath. Quips under pressure as a coping mechanism, but the quips are good. Speaks in plain American English, no affectations. Faster and more impulsive than the other protagonists. Does not perform toughness — she simply is tough.

LEGACY CHARACTER VOICES
PAI CHAN — Composed authority with a blade of wit. Slightly formal register. "Not bad" is her standing ovation. "Kung fu" means holistic mastery; do not substitute.
AKIRA YUKI — Philosophical without being pompous. Simple, weighted observations. "You are ten years too early" — preserve exactly in all languages.
SARAH BRYANT — Cold, competent, mission-focused. Steely and direct. Treats Stella as a junior agent, not a surrogate daughter.
VANESSA LEWIS — Two distinct registers. As "April/Sharon": warm, maternal. As herself: all-business, clipped, military. The contrast is a plot reveal.
JACKY BRYANT — Confident, emotionally open. Carries grief for Sarah close to the surface. Not comic relief.
GOH HINOGAMI — Patrician, unhurried, serenely certain of his own superiority. Never raises his voice. His patience is the threat.

SUPPORTING CHARACTER VOICES
ENRIQUE — Fast, high-energy, one beat too loud. His switch to "Concordat enforcer" mode is a performance.
JASPER — Easygoing with sardonic undertow. Longer sentences than Cielo. Self-deprecation as defense.
DIVINA — Direct, quick, not to be patronized. Driest humor in Cielo's group.
CHUCK — Old-school, gruff, enormous love expressed as toughness.
CISCO PERFETTI — Continental European authority. Never raises his voice. Delivers menace as courtesy. Most frightening when most relaxed.
REBECCA — Smooth, measured, always building toward something. Better liar than almost everyone except Fortuna. Warm surface, calculated core.
HABITO — Even, warm, occasionally self-effacing. The emotional anchor of Fortuna's story.
SANTO (THE DON) — Warm, direct, profane when pushed. Commands rooms without effort.
NICOLAS — Conflicted soldier. Restrained, deliberate sentences. A good man trapped by circumstance.
LUCAS — Wry, permissive, genuinely likeable. His positivity is a professional survival skill.
WAI-LEUNG CHIU — Young, volatile, performing toughness to cover shame. His cruelty comes from insecurity.
EDEUN — Fast, boundary-free, chaotic exterior concealing real loss.
REN FUJIMA — Controlled ambition dressed as principle. Speaks with precision.
SUMIRE — Haughty, formal, lethal. Speaks as if everyone is beneath her attention.
CHERRY — Street-smart, direct, protective of her girls.
MR. CHIU — Quiet, dignified, carrying grief without displaying it.
"""

# ── Character profiles dict ─────────────────────────────────────────────────
CHARACTER_PROFILES: dict[str, str] = {
    "CIELO SALINAS-WALKER": "Physical, direct, emotionally contained. Humor is self-deprecating and delivered flat. Speaks in short declaratives. Code-switches into Paraguayan Spanish naturally. Uses understatement as armor. Does not philosophize. His anger is a gearshift, not a boil.",
    "EZECHIEL BROOKS": "Low-affect, steady, quietly sardonic. British. Spends words carefully. Delivers menace with perfect calm. Does not unpack his own trauma. Cannot raise his voice prior to his one eruption moment. Preserve British phrasing.",
    "FORTUNA DE LA ROSSA": "Precise, dry, always two moves ahead. Sentences are built, not improvised. Inner-voice is clinical notation. Does not gloat. Does not explain her reasoning until she chooses to. Emotional exposure is rare and costly.",
    "STELLA BRIDGES": "Upbeat surface, genuine steel underneath. Quips under pressure. Speaks in plain American English. Faster and more impulsive. Does not perform toughness — she simply is tough.",
    "PAI CHAN": "Composed authority with a blade of wit. Slightly formal register. 'Not bad' is her standing ovation.",
    "AKIRA YUKI": "Philosophical without being pompous. Simple, weighted observations. 'You are ten years too early' — preserve exactly.",
    "SARAH BRYANT": "Cold, competent, mission-focused. Steely and direct. Treats Stella as a junior agent, not a surrogate daughter.",
    "VANESSA LEWIS": "Two distinct registers. As 'April/Sharon': warm, maternal. As herself: all-business, clipped, military.",
    "JACKY BRYANT": "Confident, emotionally open. Carries grief for Sarah close to the surface. Not comic relief.",
    "GOH HINOGAMI": "Patrician, unhurried, serenely certain of his own superiority. Never raises his voice. His patience is the threat.",
    "ENRIQUE": "Fast, high-energy, one beat too loud. His switch to 'Concordat enforcer' mode is a performance.",
    "JASPER": "Easygoing with sardonic undertow. Longer sentences than Cielo. Self-deprecation as defense.",
    "DIVINA": "Direct, quick, not to be patronized. Driest humor in Cielo's group.",
    "CHUCK": "Old-school, gruff, enormous love expressed as toughness.",
    "CISCO PERFETTI": "Continental European authority. Never raises his voice. Delivers menace as courtesy. Most frightening when most relaxed.",
    "REBECCA": "Smooth, measured, always building toward something. Better liar than almost everyone except Fortuna. Warm surface, calculated core.",
    "HABITO": "Even, warm, occasionally self-effacing. The emotional anchor of Fortuna's story.",
    "SANTO (THE DON)": "Warm, direct, profane when pushed. Commands rooms without effort.",
    "NICOLAS": "Conflicted soldier. Restrained, deliberate sentences. A good man trapped by circumstance.",
    "LUCAS": "Wry, permissive, genuinely likeable. His positivity is a professional survival skill.",
    "WAI-LEUNG CHIU": "Young, volatile, performing toughness to cover shame. His cruelty comes from insecurity.",
    "EDEUN": "Fast, boundary-free, chaotic exterior concealing real loss.",
    "REN FUJIMA": "Controlled ambition dressed as principle. Speaks with precision.",
    "SUMIRE": "Haughty, formal, lethal. Speaks as if everyone is beneath her attention.",
    "CHERRY": "Street-smart, direct, protective of her girls.",
    "MR. CHIU": "Quiet, dignified, carrying grief without displaying it.",
}

# ── Sample scripts (with intentional issues) ────────────────────────────────
SAMPLE_SCRIPTS: dict[str, str] = {
    "Sample 1: Cielo & Ezechiel \u2014 Bar Scene": textwrap.dedent("""\
        INT. THE BROKEN CROWN \u2014 NIGHT
        A dim bar. Rain against glass. Cielo sits; Ezechiel slides in.

        EZECHIEL
        Long night.

        CIELO
        I've been contemplating the fundamental nature of loyalty,
        you know? The epistemological implications of our
        predicament are rather unfortunate.

        EZECHIEL
        (SHOUTING)
        THAT'S THE STUPIDEST THING I'VE EVER HEARD, MATE! WAKE UP!

        CIELO
        That plan's dead in the water, hermano.

        EZECHIEL
        (calmly)
        Been a pleasure, mate.
    """),
    "Sample 2: Fortuna & Rebecca \u2014 The Confrontation": textwrap.dedent("""\
        INT. SANTO'S STUDY \u2014 EVENING
        Books on the walls. Single lamp. Fortuna enters; Rebecca waits.

        REBECCA
        (screaming)
        OH MY GOD FORTUNA I can't BELIEVE you went behind my back!

        FORTUNA
        (gloating)
        Ha! I KNEW you'd crack. I am simply smarter than you.
        Let me explain exactly how I figured it all out.

        REBECCA
        LOL you're so funny! Let's go get tacos!

        FORTUNA
        Totally! BFF road-trip! This is like college, right?

        HABITO
        Ladies, we've got bigger fish to fry. The synergies
        aren't aligning on this one.
    """),
}

# ── helper utilities ────────────────────────────────────────────────────────
def detect_characters(script: str) -> list[str]:
    upper = script.upper()
    found = []
    for name in CHARACTER_PROFILES:
        variants = [name]
        parts = name.split()
        if len(parts) > 1:
            variants.append(parts[0])
        if name == "SANTO (THE DON)":
            variants.extend(["SANTO", "THE DON"])
        for v in variants:
            if v in upper:
                found.append(name); break
    return sorted(set(found))


def _scan_balanced(raw: str, open_ch: str, close_ch: str):
    """Extract the first balanced {...} or [...] from raw, ignoring brackets
    that appear inside JSON string literals."""
    idx = raw.find(open_ch)
    if idx == -1:
        return None
    depth, in_str, esc = 0, False, False
    for i in range(idx, len(raw)):
        ch = raw[i]
        if in_str:
            if esc:           esc = False
            elif ch == "\\": esc = True
            elif ch == '"':   in_str = False
        else:
            if ch == '"':       in_str = True
            elif ch == open_ch:  depth += 1
            elif ch == close_ch:
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(raw[idx:i+1])
                    except Exception:
                        return None
    return None


def _repair_truncated_json(raw: str):
    """Best-effort recovery of JSON that was cut off mid-stream (e.g. the
    model hit max_tokens). Trims back to the last syntactically complete
    token, strips a dangling comma, then closes any open brackets."""
    start = raw.find("{")
    if start == -1:
        return None
    s = raw[start:]
    stack: list[str] = []
    in_str = esc = False
    last_good = 0
    for i, ch in enumerate(s):
        if in_str:
            if esc:            esc = False
            elif ch == "\\":  esc = True
            elif ch == '"':
                in_str = False
                last_good = i + 1
        else:
            if ch == '"':       in_str = True
            elif ch in "{[":    stack.append(ch)
            elif ch in "}]":
                if stack: stack.pop()
                last_good = i + 1
            elif ch not in ", :\n\t\r":
                last_good = i + 1          # number / true / false / null chars
    cand = s[:last_good].rstrip().rstrip(",")
    closers = "".join("}" if c == "{" else "]" for c in reversed(stack))
    try:
        return json.loads(cand + closers)
    except Exception:
        return None


def parse_llm_json(raw: str) -> Any:
    raw = raw.strip()

    def _ensure_dict(obj: Any) -> Any:
        """Unwrap a single-element list that wraps a dict (e.g. [{...}] -> {...}).
        Any other list is treated as a parse failure so callers always get a dict."""
        if isinstance(obj, list):
            if len(obj) == 1 and isinstance(obj[0], dict):
                return obj[0]
            return {"raw_text": raw, "parse_error": True}
        return obj

    # 1. whole string is valid JSON
    try:
        return _ensure_dict(json.loads(raw))
    except Exception:
        pass
    # 2. fenced ```json block
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", raw, re.S)
    if m:
        try:
            return _ensure_dict(json.loads(m.group(1)))
        except Exception:
            pass
    # 3. string-aware balanced extraction ({ first, then [)
    for co, cc in [("{", "}"), ("[", "]")]:
        obj = _scan_balanced(raw, co, cc)
        if obj is not None:
            return _ensure_dict(obj)
    # 4. truncation repair (model hit max_tokens mid-object)
    obj = _repair_truncated_json(raw)
    if obj is not None:
        obj["_truncation_repaired"] = True
        return _ensure_dict(obj)
    return {"raw_text": raw, "parse_error": True}


def extract_docx_text(file) -> str:
    """Extract plain text from a .docx file using python-docx."""
    if not HAS_DOCX:
        return "[python-docx not installed — run: pip install python-docx]"
    try:
        doc = DocxDocument(file)
        parts = []
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text)
        # Also pull text from tables
        for table in doc.tables:
            for row in table.rows:
                row_text = "	".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    parts.append(row_text)
        return "\n".join(parts)
    except Exception as exc:
        return f"[DOCX extraction error: {exc}]"


def extract_xlsx_text(file) -> str:
    """Extract sheet contents from an .xlsx file using openpyxl."""
    if not HAS_OPENPYXL:
        return "[openpyxl not installed — run: pip install openpyxl]"
    try:
        wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
        parts = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            parts.append(f"=== Sheet: {sheet_name} ===")
            for row in ws.iter_rows(values_only=True):
                row_str = "	".join(
                    str(cell) if cell is not None else "" for cell in row
                ).rstrip()
                if row_str.strip():
                    parts.append(row_str)
        return "\n".join(parts)
    except Exception as exc:
        return f"[XLSX extraction error: {exc}]"


# ── Script-table parser (CSV / XLSX with SceneID, Talker, Text columns) ────────

# Columns we recognise in SEGA-style script sheets
_SCRIPT_COLS = {"sceneid", "id", "talker", "text"}

def _is_script_table(headers: list[str]) -> bool:
    """Return True when the header row matches the expected script column set."""
    normed = {h.strip().lower() for h in headers if h}
    return _SCRIPT_COLS.issubset(normed)


def extract_script_table(rows: list[dict], use_jb_pass: bool = False) -> tuple[str, dict[int, int]]:
    """
    Convert a list of row-dicts (keyed by column header) from a SEGA script
    sheet into a clean script string suitable for NQA analysis.

    Returns (text, line_map) where line_map maps each 1-based output line
    number to the 0-based ordinal of the source data row it came from
    (synthetic lines such as scene separators have no entry). The map is what
    lets analysis results be written back into the original spreadsheet.

    Column semantics
    ----------------
    SceneID  – scene / event identifier; a change signals a new scene block
    ID       – unique line ID; empty rows are stage directions / notes
    Talker   – character name; empty = non-dialogue row
    Text     – original / source dialogue or stage direction
    JB PASS  – editorial revision of Text (may be empty)
    PURPOSE  – note explaining the revision (ignored during extraction)

    Parsing rules
    -------------
    * Rows starting with "?" or "<" are system/meta markers → skip entirely
    * Rows consisting only of dashes → skip (visual separators in the sheet)
    * Rows where Talker is empty but Text begins with INT./EXT. → scene heading
    * Rows where Talker is empty → stage direction, wrapped in [ ]
    * Rows with a Talker → dialogue line; format as  "TALKER: dialogue"
      - Strip inline VO cues after "//"  (e.g.  "//VO: phone call")
      - Use JB PASS column when use_jb_pass=True *and* the cell is non-empty
    """
    import re as _re

    def _clean_dialogue(raw: str) -> str:
        """Remove //VO: and //TB: director / translator notes from a line."""
        return _re.split(r"//(?:VO|TB|ALT)\s*:", raw, maxsplit=1)[0].strip()

    out:  list[str]        = []
    srcs: list[int | None] = []          # parallel: source row ordinal per line

    def _emit(line: str, source: int | None = None):
        out.append(line); srcs.append(source)

    current_scene: str = ""

    for ridx, row in enumerate(rows):
        scene_id = str(row.get("SceneID") or "").strip()
        talker   = str(row.get("Talker")  or "").strip()
        text     = str(row.get("Text")    or "").strip()
        jb       = str(row.get("JB PASS") or "").strip()

        if not text and not talker:
            continue                                       # truly blank row

        # ── system / meta markers ──
        if text.startswith("?") or text.startswith("<"):
            continue

        # ── visual separator lines ──
        if not text.replace("-", "").strip():
            continue

        # ── scene break ──
        if scene_id and scene_id != current_scene:
            current_scene = scene_id
            _emit(""); _emit("─" * 60); _emit(f"[{scene_id}]")

        # ── stage direction / scene heading ──
        if not talker:
            clean_text = _re.split(r"//(?:TB|VO|ALT)\s*:", text, maxsplit=1)[0].strip()
            if not clean_text:
                continue
            if clean_text.startswith(("INT.", "EXT.")):
                _emit(""); _emit(clean_text, ridx)
            else:
                _emit(f"  [{clean_text}]", ridx)
            continue

        # ── dialogue line ──
        dialogue = jb if (use_jb_pass and jb) else text
        _emit(f"{talker}: {_clean_dialogue(dialogue)}", ridx)

    line_map = {i + 1: s for i, s in enumerate(srcs) if s is not None}
    return "\n".join(out), line_map


def extract_csv_script(file, use_jb_pass: bool = False) -> tuple[str, bool, dict[int, int]]:
    """
    Parse an uploaded CSV file. If it looks like a SEGA script table
    (contains SceneID / Talker / Text headers) run extract_script_table;
    otherwise fall back to a plain UTF-8 decode.

    Returns (text, has_jb_pass, line_map). line_map is empty for plain CSVs.
    """
    import io as _io, csv as _csv

    raw = file.read()
    if isinstance(raw, bytes):
        # Try common encodings in order; script files from Windows tools are
        # often saved as cp1252 (which encodes em-dashes, curly quotes, etc.)
        for _enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                raw = raw.decode(_enc); break
            except (UnicodeDecodeError, AttributeError):
                continue
        else:
            raw = raw.decode("utf-8", errors="replace")

    reader = _csv.DictReader(_io.StringIO(raw))
    headers = reader.fieldnames or []
    has_jb  = "JB PASS" in headers

    if _is_script_table(headers):
        text, line_map = extract_script_table(list(reader), use_jb_pass=use_jb_pass)
        return text, has_jb, line_map

    # Plain CSV fallback
    return raw, False, {}


def extract_xlsx_script(file, use_jb_pass: bool = False) -> tuple[str, bool, dict[int, int]]:
    """
    Parse an uploaded XLSX file. If the first sheet matches the SEGA script
    table format, run extract_script_table; otherwise fall back to a raw
    tab-dump.

    Returns (text, has_jb_pass, line_map). line_map is empty in fallback mode.
    """
    if not HAS_OPENPYXL:
        return "[openpyxl not installed — run: pip install openpyxl]", False, {}
    try:
        wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows_iter  = ws.iter_rows(values_only=True)
        header_row = next(rows_iter, None)
        if header_row is None:
            return "", False, {}
        headers = [str(h).strip() if h is not None else "" for h in header_row]
        has_jb  = "JB PASS" in headers

        if _is_script_table(headers):
            row_dicts = [
                {headers[i]: (str(v).strip() if v is not None else "")
                 for i, v in enumerate(row)}
                for row in rows_iter
            ]
            text, line_map = extract_script_table(row_dicts, use_jb_pass=use_jb_pass)
            return text, has_jb, line_map

        # ── raw dump fallback (non-script XLSX) ──
        file.seek(0)
        wb2   = openpyxl.load_workbook(file, data_only=True, read_only=True)
        parts = []
        for sheet_name in wb2.sheetnames:
            parts.append(f"=== Sheet: {sheet_name} ===")
            for row in wb2[sheet_name].iter_rows(values_only=True):
                row_str = "\t".join(str(c) if c is not None else "" for c in row).rstrip()
                if row_str.strip():
                    parts.append(row_str)
        return "\n".join(parts), False, {}
    except Exception as exc:
        return f"[XLSX extraction error: {exc}]", False, {}


def get_severity_badge(severity: str) -> str:
    s = (severity or "info").lower().strip()
    if s in ("critical","hard","high","red"):   return '<span class="sev-c">CRITICAL</span>'
    if s in ("warning","medium","soft","yellow"): return '<span class="sev-w">WARNING</span>'
    return '<span class="sev-i">INFO</span>'


def score_color(score: int) -> str:
    if score >= 80: return "#00BB66"
    if score >= 60: return "#F5C218"
    return "#CC2244"


# ── analysis history persistence ─────────────────────────────────────────────
# History is stored per-user (keyed by sign-in email) in a JSON file in the
# home directory, so it survives page reloads and new sessions on the same
# container. Each entry records a hash of the script so repeat analyses of the
# same material can be compared draft-over-draft.
_HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".snake_nqa_history.json")

def _script_hash(script: str) -> str:
    """Stable content hash (whitespace-normalised) for draft comparison."""
    norm = "\n".join(l.strip() for l in script.splitlines() if l.strip())
    return hashlib.md5(norm.encode("utf-8", errors="replace")).hexdigest()[:10]

def _script_label(script: str) -> str:
    """Human-readable label: first [SCENE_ID] tag if present, else first line."""
    for ln in script.splitlines():
        ln = ln.strip()
        if not ln or set(ln) <= set("\u2500-\u2014 "):
            continue
        m = re.match(r"\[([A-Za-z0-9_]+)\]", ln)
        if m:
            return m.group(1)[:34]
        return (ln[:32] + "\u2026") if len(ln) > 32 else ln
    return "untitled"

def load_history(email: str) -> list[dict]:
    try:
        with open(_HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get(email, [])
    except Exception:
        return []

def append_history(email: str, entry: dict) -> None:
    data: dict = {}
    try:
        with open(_HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        pass
    data.setdefault(email, []).append(entry)
    data[email] = data[email][-50:]          # keep last 50 per user
    try:
        with open(_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass                                  # read-only FS: degrade silently



# ═══════════════════════════════════════════════════════════════════════════════
# AUTHENTICATION — Email OTP via AWS SES
# SES creds:     AWS_ACCESS_KEY_ID  /  AWS_SECRET_ACCESS_KEY  /  AWS_SES_REGION
# Bedrock creds: AWS_BEDROCK_ACCESS_KEY_ID / AWS_BEDROCK_SECRET_ACCESS_KEY / AWS_BEDROCK_REGION
# ═══════════════════════════════════════════════════════════════════════════════

ALLOWED_DOMAIN    = "segaamerica.com"
OTP_EXPIRY_SECS   = 600          # 10 minutes
TOKEN_EXPIRY_DAYS = 7
MAX_OTP_ATTEMPTS  = 5

def _generate_otp() -> str:
    return f"{random.SystemRandom().randint(0, 999999):06d}"

def _send_otp_email(email: str, code: str) -> bool:
    try:
        ses = boto3.client("ses",
            region_name=st.secrets.get("AWS_SES_REGION", "us-east-1"),
            aws_access_key_id=st.secrets["AWS_ACCESS_KEY_ID"],
            aws_secret_access_key=st.secrets["AWS_SECRET_ACCESS_KEY"])
        ses.send_email(
            Source=st.secrets.get("EMAIL_FROM", "streamlit_otp@segaamerica.com"),
            Destination={"ToAddresses": [email]},
            Message={"Subject": {"Data": "SNAKE NQA \u2014 Verification Code", "Charset": "UTF-8"},
                     "Body": {"Text": {"Data": f"Your code is:\n\n    {code}\n\nExpires in 10 minutes.", "Charset": "UTF-8"}}})
        return True
    except Exception as exc:
        st.error(f"Failed to send email: {exc}")
        return False

def _sign_token(email: str) -> str:
    key = st.secrets.get("COOKIE_SIGNING_KEY", "dev-key-change-me").encode()
    ts = str(int(time.time()))
    payload = f"{email}|{ts}"
    sig = hmac.new(key, payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}|{sig}".encode()).decode()

def _verify_token(token: str) -> str | None:
    try:
        raw = base64.urlsafe_b64decode(token.encode()).decode()
        parts = raw.split("|")
        if len(parts) != 3: return None
        email, ts_str, sig = parts
        key = st.secrets.get("COOKIE_SIGNING_KEY", "dev-key-change-me").encode()
        expected = hmac.new(key, f"{email}|{ts_str}".encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected): return None
        if time.time() - int(ts_str) > TOKEN_EXPIRY_DAYS * 86400: return None
        return email
    except Exception:
        return None

def _check_auth() -> bool:
    if st.session_state.get("auth_verified"): return True
    token = st.query_params.get("t")
    if token:
        email = _verify_token(token)
        if email:
            st.session_state["auth_verified"] = True
            st.session_state["auth_email"] = email
            return True
    return False

def _render_login():
    # Inject CSS so the block container is vertically centered.
    # The old approach of wrapping st.columns in an opening/closing st.markdown div
    # does not work — Streamlit renders each st.markdown as a sibling element, not a
    # wrapper, so the div just became a blank 70 vh block sitting above the form.
    st.markdown("""
    <style>
    .block-container {
        padding-top: calc(50vh - 220px) !important;
        padding-bottom: 2rem !important;
    }
    </style>
    """, unsafe_allow_html=True)
    _, col_c, _ = st.columns([1, 2, 1])
    with col_c:
        st.markdown(
            '<div class="nqa-card" style="text-align:center;padding:2.5rem;">'
            '<span style="font-size:3.5rem;">\U0001F40D</span><br/>'
            '<span style="font-size:1.6rem;font-weight:700;color:#fff;">SNAKE Narrative QA</span><br/>'
            '<span style="color:#8899BB;font-size:.9rem;">Sign in with your SEGA email</span>'
            '</div>', unsafe_allow_html=True)
        email = st.text_input("Email", placeholder="you@segaamerica.com", key="login_email")
        if st.button("\U0001F4E7 Send Verification Code", use_container_width=True):
            if not email or not email.strip().endswith(f"@{ALLOWED_DOMAIN}"):
                st.error(f"Only @{ALLOWED_DOMAIN} addresses are allowed.")
            elif not HAS_BOTO3:
                st.error("boto3 is required for OTP email.")
            else:
                code = _generate_otp()
                if _send_otp_email(email.strip(), code):
                    st.session_state.update(otp_code=code, otp_time=time.time(),
                                            otp_email=email.strip(), otp_attempts=0)
                    st.success("Code sent! Check your inbox.")
        if st.session_state.get("otp_code"):
            code_in = st.text_input("6-digit code", max_chars=6, key="login_code")
            if st.button("\u2705 Verify", use_container_width=True):
                att = st.session_state.get("otp_attempts", 0)
                if att >= MAX_OTP_ATTEMPTS:
                    st.error("Too many attempts. Request a new code.")
                elif time.time() - st.session_state.get("otp_time", 0) > OTP_EXPIRY_SECS:
                    st.error("Code expired. Request a new code.")
                elif code_in == st.session_state["otp_code"]:
                    em = st.session_state["otp_email"]
                    st.session_state.update(auth_verified=True, auth_email=em)
                    for k in ("otp_code","otp_time","otp_email","otp_attempts"):
                        st.session_state.pop(k, None)
                    st.query_params["t"] = _sign_token(em)
                    st.rerun()
                else:
                    st.session_state["otp_attempts"] = att + 1
                    rem = MAX_OTP_ATTEMPTS - st.session_state["otp_attempts"]
                    st.error(f"Incorrect code. {rem} attempt{'s' if rem!=1 else ''} left.")

# ── auth gate ───────────────────────────────────────────────────────────────
if not _check_auth():
    _render_login()
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# LLM CLIENT (uses Bedrock credentials — separate from SES)
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_MODEL = "us.anthropic.claude-sonnet-4-6"

@st.cache_resource
def get_client():
    try:
        return AnthropicBedrock(
            aws_access_key=st.secrets["AWS_BEDROCK_ACCESS_KEY_ID"],
            aws_secret_key=st.secrets["AWS_BEDROCK_SECRET_ACCESS_KEY"],
            aws_region=st.secrets.get("AWS_BEDROCK_REGION", "us-east-1"))
    except Exception: pass
    try:
        return Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except Exception: pass
    return None

def call_agent(client, model: str, system_prompt: str, user_prompt: str,
               temperature: float = 0.15, cache_system: bool = False) -> str:
    """Invoke the model. When cache_system=True the system prompt is marked
    with cache_control so identical prefixes (style guide + profiles + script)
    are cached across the four agents and across re-runs within the cache TTL,
    cutting input cost ~90% on hits. Falls back to an uncached call if the
    region/SDK rejects cache_control."""
    sys_param: Any = system_prompt
    if cache_system:
        sys_param = [{"type": "text", "text": system_prompt,
                      "cache_control": {"type": "ephemeral"}}]
    try:
        resp = client.messages.create(model=model, max_tokens=8192,
            temperature=temperature, system=sys_param,
            messages=[{"role": "user", "content": user_prompt}])
        return resp.content[0].text
    except Exception as exc:
        if cache_system:
            try:    # cache_control unsupported here → plain retry
                resp = client.messages.create(model=model, max_tokens=8192,
                    temperature=temperature, system=system_prompt,
                    messages=[{"role": "user", "content": user_prompt}])
                return resp.content[0].text
            except Exception as exc2:
                return json.dumps({"error": str(exc2)})
        return json.dumps({"error": str(exc)})

def relevant_profiles(characters: list[str]) -> str:
    lines = [f"### {c}\n{CHARACTER_PROFILES[c]}\n" for c in characters if c in CHARACTER_PROFILES]
    return "\n".join(lines) if lines else "(No specific profiles matched.)"


# ── agent system prompts ────────────────────────────────────────────────────

_VK_SYS = """\
You are VOICE KEEPER for the game SNAKE. Check every dialogue line against CHARACTER PROFILES.
Check: vocabulary fingerprint, sentence-construction, emotional register, code-switching, tonal arc.
RULES: nobody speaks in exposition; profanity is earned; speech rhythms are NOT interchangeable; avoid over-explanation, self-pity, clich\u00e9.
Report the most impactful issues only — maximum 25, prioritised by severity.
You cannot know full character backstories, future reveals, or writer intent; apparent inconsistencies may be deliberate (a reveal, growth, deception). When a flag depends on what a character could plausibly know or say, begin the explanation with "Verify with writer:" and frame it as a question to confirm, not a definitive error.
OUTPUT ONLY a single valid JSON object — no prose, no markdown fences:
{"voice_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact full line>","flagged_phrase":"<exact substring of line_text that is the specific concern>","character":"<NAME>","issue_type":"<vocabulary_deviation|register_mismatch|emotional_register|code_switching|tonal_arc|exposition>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","direction":"<suggestion, do NOT rewrite>"}]}"""

_LW_SYS = """\
You are LORE WARDEN for SNAKE. Identify any contradiction with established canon, world rules, character histories or timeline.
Check: factual contradictions, timeline violations, world-rule violations, character knowledge state, relationship state, absent character references.
Report the most impactful issues only — maximum 25, prioritised by severity.
You cannot know full character backstories, future reveals, or writer intent; apparent inconsistencies may be deliberate (a reveal, growth, deception). When a flag depends on what a character could plausibly know or say, begin the explanation with "Verify with writer:" and frame it as a question to confirm, not a definitive error.
OUTPUT ONLY a single valid JSON object — no prose, no markdown fences:
{"lore_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact full line>","flagged_phrase":"<exact substring of line_text that is the specific concern>","issue_type":"<factual_contradiction|timeline_violation|world_rule|knowledge_state|relationship_state|absent_character>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","resolution_options":"<fixes>"}]}"""

_DS_SYS = """\
You are DIALECT SENTINEL for SNAKE. Protect linguistic texture.
Check: faction/region slang, idiom coherence (no real-world idioms that break setting), profanity-system integrity, politeness markers, cultural register, regional Spanish (Cielo = Paraguayan).
Report the most impactful issues only — maximum 25, prioritised by severity.
You cannot know full character backstories, future reveals, or writer intent; apparent inconsistencies may be deliberate (a reveal, growth, deception). When a flag depends on what a character could plausibly know or say, begin the explanation with "Verify with writer:" and frame it as a question to confirm, not a definitive error.
OUTPUT ONLY a single valid JSON object — no prose, no markdown fences:
{"dialect_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact full line>","flagged_phrase":"<exact substring of line_text that is the specific concern>","character":"<NAME>","issue_type":"<slang_inconsistency|idiom_violation|profanity_misuse|formality_error|cultural_register|language_variant>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","suggestion":"<direction>"}]}"""

_TC_SYS = """\
You are TONE CARTOGRAPHER for SNAKE. Analyse scene-level emotional trajectory.
Check: tonal whiplash, escalation/de-escalation curves, thematic consistency, clich\u00e9s, pacing.
Keep tone_map to at most 12 segments — group adjacent scenes that share a tone rather than listing every scene.
Keep all text fields concise (1-2 sentences). Report the most impactful issues only — maximum 15, prioritised by severity.
For each issue also include the line_number and exact line_text of ONE representative line from that segment (use the numeric prefix shown in the script).
You cannot know full character backstories, future reveals, or writer intent; apparent inconsistencies may be deliberate (a reveal, growth, deception). When a flag depends on what a character could plausibly know or say, begin the explanation with "Verify with writer:" and frame it as a question to confirm, not a definitive error.
OUTPUT ONLY a single valid JSON object — no prose, no markdown fences:
{"tone_score":<0-100>,"tone_map":[{"scene_segment":"<beat>","tone_description":"<e.g. restrained tension>","intended_effect":"<player feeling>"}],"issues":[{"scene_segment":"<part>","line_number":<int representative line>,"line_text":"<exact representative line>","issue_type":"<tonal_whiplash|flat_escalation|thematic_inconsistency|cliche|pacing>","severity":"<critical|warning|info>","explanation":"<2-3 sent>"}]}"""

_ARB_SYS = """\
You are ARBITER for SNAKE narrative QA. You receive outputs from Voice Keeper, Lore Warden, Dialect Sentinel, Tone Cartographer.
1. Merge & deduplicate overlapping flags.  2. Rank by severity & narrative impact.
3. Identify cross-agent conflicts.  4. Detect patterns (character flagged repeatedly = possible intentional evolution).
Top 20 items in priority_list maximum.
Preserve any "Verify with writer:" framing from agents — these are verification requests, not confirmed errors; do not upgrade them to definitive claims.
OUTPUT ONLY a single valid JSON object — no prose, no markdown fences:
{"overall_score":<0-100>,"critical_count":<int>,"warning_count":<int>,"info_count":<int>,
"priority_list":[{"rank":<int>,"agent_source":"<name>","severity":"<lvl>","summary":"<1 sent>","recommendation":"<action>"}],
"narrative_health_summary":"<2-3 paragraph assessment>",
"character_bible_updates":[{"character":"<NAME>","suggested_update":"<reason & suggestion>"}]}"""

# ── agent runners ───────────────────────────────────────────────────────────

def _numbered_script(script: str) -> str:
    """Prefix every line with its 1-based number so agents report exact,
    verifiable line references."""
    return "\n".join(f"{i:>4} | {l}" for i, l in enumerate(script.split("\n"), 1))

def _shared_system(script: str, characters: list[str]) -> str:
    """Context shared verbatim by all four line agents. Kept identical (and
    placed in the system block with cache_control) so one cache entry serves
    every agent call for this script."""
    return ("You are one of several specialised narrative-QA agents for the game SNAKE. "
            "Your specific role and output schema follow in the user message.\n\n"
            f"## STYLE GUIDE\n{STYLE_GUIDE_TEXT}\n\n"
            f"## DETECTED CHARACTER PROFILES\n{relevant_profiles(characters)}\n\n"
            "## SCRIPT TO ANALYSE\n"
            "Each line is prefixed with its line number and ' | '.\n"
            f"```\n{_numbered_script(script)}\n```\n\n"
            "When reporting an issue: line_number MUST be the numeric prefix of the "
            "flagged line exactly as shown; line_text MUST be the exact text AFTER "
            "the ' | ' separator (never include the number prefix); flagged_phrase "
            "MUST be copied verbatim from within line_text.")

def _run_line_agent(client, model, role_prompt, script, characters, temp):
    return parse_llm_json(call_agent(
        client, model,
        _shared_system(script, characters),                 # cached prefix
        role_prompt + "\n\nReturn ONLY valid JSON matching your schema.",
        temp, cache_system=True))

def run_voice_keeper(client, model, script, characters, temp):
    return _run_line_agent(client, model, _VK_SYS, script, characters, temp)

def run_lore_warden(client, model, script, characters, temp):
    return _run_line_agent(client, model, _LW_SYS, script, characters, temp)

def run_dialect_sentinel(client, model, script, characters, temp):
    return _run_line_agent(client, model, _DS_SYS, script, characters, temp)

def run_tone_cartographer(client, model, script, characters, temp):
    return _run_line_agent(client, model, _TC_SYS, script, characters, temp)

def run_arbiter(client, model, agent_results: dict, temp):
    prompt = (f"## AGENT OUTPUTS\n```json\n{json.dumps(agent_results, indent=2, default=str)}\n```\n\n"
              "Merge, deduplicate, rank. Return ONLY valid JSON per your schema.")
    return parse_llm_json(call_agent(client, model, _ARB_SYS, prompt, temp))



# ═══════════════════════════════════════════════════════════════════════════════
# RENDERING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

AGENT_META = {
    "Voice Keeper":      {"icon": "\U0001F3AD", "key": "voice_keeper"},
    "Lore Warden":       {"icon": "\U0001F4DC", "key": "lore_warden"},
    "Dialect Sentinel":  {"icon": "\U0001F5E3", "key": "dialect_sentinel"},
    "Tone Cartographer": {"icon": "\U0001F3B5", "key": "tone_cartographer"},
    "Arbiter":           {"icon": "\u2696",     "key": "arbiter"},
}

def render_agent_progress(statuses: dict[str, str], timings: dict | None = None) -> str:
    rows = []
    for name, meta in AGENT_META.items():
        sv = statuses.get(name, "pending")
        lb = {"pending":"Waiting\u2026","running":"Analysing\u2026",
              "complete":"Done \u2714","error":"Error \u2718"}.get(sv, sv)
        if timings and name in timings and sv in ("complete", "error"):
            lb += f' <span style="color:#556;">{timings[name]:.0f}s</span>'
        rows.append(
            f'<div class="ag-row"><div class="ag-dot dot-{sv}"></div>'
            f'<div class="ag-name">{meta["icon"]} {name}</div>'
            f'<div class="ag-st">{lb}</div></div>')
    return '<div class="nqa-card">' + "".join(rows) + "</div>"

def render_score_cards(arb, vk, lw, ds, tc):
    scores = [("Overall", arb.get("overall_score","?")),
              ("Voice", vk.get("voice_score","?")),
              ("Lore", lw.get("lore_score","?")),
              ("Dialect", ds.get("dialect_score","?")),
              ("Tone", tc.get("tone_score","?"))]
    cols = st.columns(5)
    for col, (label, val) in zip(cols, scores):
        v = int(val) if isinstance(val, (int, float)) else 0
        clr = score_color(v)
        col.markdown(
            f'<div class="nqa-score-card"><div class="sv" style="color:{clr}">{val}</div>'
            f'<div class="sl">{label}</div></div>', unsafe_allow_html=True)

def _collect_all_issues(results: dict) -> list[dict]:
    out = []
    for ak in ("voice_keeper","lore_warden","dialect_sentinel","tone_cartographer"):
        for iss in results.get(ak, {}).get("issues", []):
            c = dict(iss); c["_agent"] = ak; out.append(c)
    return out

_AGENT_COLORS = {
    "voice_keeper":      "#C77DFF",
    "lore_warden":       "#00BB66",
    "dialect_sentinel":  "#00AADD",
    "tone_cartographer": "#F5C218",
}

def _esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _hl_phrase(line_text: str, phrase: str, severity: str) -> str:
    """Return escaped line_text with the flagged phrase wrapped in a
    severity-coloured highlight span. Falls back to the plain escaped line
    when the phrase is missing or not found."""
    lt = _esc(line_text)
    ph = _esc((phrase or "").strip())
    if not ph:
        return lt
    cls = {"critical": "fl-c", "warning": "fl-w"}.get((severity or "").lower(), "fl-i")
    i = lt.lower().find(ph.lower())
    if i == -1:
        return lt
    return lt[:i] + f'<span class="{cls}">' + lt[i:i+len(ph)] + "</span>" + lt[i+len(ph):]

def _agent_chip(agent_key: str) -> str:
    clr = _AGENT_COLORS.get(agent_key, "#8899BB")
    nm = agent_key.replace("_", " ").title()
    return f'<span class="ag-chip" style="background:{clr}22;color:{clr};border:1px solid {clr}66;">{nm}</span>'

def _norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())

def _attach_issues_to_lines(lines: list[str], all_iss: list[dict]):
    """Map each issue to a script line. An agent-supplied line_number is only
    trusted when the issue's line_text actually matches that line; otherwise
    we search the whole script for the quoted text. Returns (line_map,
    unmatched) where unmatched holds issues that could not be pinned."""
    norm_lines = [_norm_text(l) for l in lines]
    lm: dict[int, list[dict]] = {}
    unmatched: list[dict] = []

    def _matches(nl: str, nt: str) -> bool:
        return bool(nt) and bool(nl) and (nt in nl or nl in nt)

    for iss in all_iss:
        ln = iss.get("line_number")
        nt = _norm_text(iss.get("line_text", ""))
        attached = False
        # 1. trust line_number only if the text corroborates (or no text given)
        if isinstance(ln, int) and 1 <= ln <= len(lines):
            if not nt or _matches(norm_lines[ln - 1], nt):
                lm.setdefault(ln, []).append(iss); attached = True
        # 2. otherwise locate by quoted text anywhere in the script
        if not attached and nt:
            for idx, nl in enumerate(norm_lines, 1):
                if _matches(nl, nt):
                    lm.setdefault(idx, []).append(iss); attached = True; break
        if not attached:
            unmatched.append(iss)
    return lm, unmatched

def _sev_bucket(s: str) -> str:
    s = (s or "").lower()
    if s in ("critical", "hard", "high"):   return "Critical"
    if s in ("warning", "medium", "soft"):  return "Warning"
    return "Info"

def render_annotated_script(script: str, results: dict):
    all_iss = _collect_all_issues(results)

    # ── filter bar ──────────────────────────────────────────────────────────
    agents_present = sorted({i.get("_agent", "") for i in all_iss if i.get("_agent")})
    agent_labels = {a: a.replace("_", " ").title() for a in agents_present}
    _has_pills = hasattr(st, "pills")
    fc1, fc2, fc3 = st.columns([1.3, 1.7, 2.6])
    with fc1:
        only_flagged = st.toggle("Flagged lines only", value=False, key="as_only_flagged",
                                 help="Hide clean lines and show just the flagged ones with a line of context.")
    with fc2:
        if _has_pills:
            sev_sel = st.pills("Severity", ["Critical", "Warning", "Info"],
                               selection_mode="multi",
                               default=["Critical", "Warning", "Info"], key="as_sev")
        else:
            sev_sel = st.multiselect("Severity", ["Critical", "Warning", "Info"],
                                     default=["Critical", "Warning", "Info"], key="as_sev")
    with fc3:
        if not agents_present:
            ag_sel = []
        elif _has_pills:
            ag_sel = st.pills("Agent", agents_present, selection_mode="multi",
                              default=agents_present,
                              format_func=lambda a: agent_labels.get(a, a), key="as_agent")
        else:
            ag_sel = st.multiselect("Agent", agents_present,
                                    default=agents_present,
                                    format_func=lambda a: agent_labels.get(a, a), key="as_agent")

    # apply filters
    filt = [i for i in all_iss
            if _sev_bucket(i.get("severity")) in (sev_sel or ["Critical", "Warning", "Info"])
            and (i.get("_agent") in ag_sel if ag_sel else True)]
    n_hidden = len(all_iss) - len(filt)
    if n_hidden:
        st.caption(f"{len(filt)} issue{'s' if len(filt) != 1 else ''} shown \u00B7 {n_hidden} hidden by filters")

    # legend
    leg = " ".join(
        f'<span class="ag-chip" style="background:{c}22;color:{c};border:1px solid {c}66;">'
        f'{AGENT_META[n]["icon"]} {n}</span>'
        for n, c in [("Voice Keeper", _AGENT_COLORS["voice_keeper"]),
                     ("Lore Warden", _AGENT_COLORS["lore_warden"]),
                     ("Dialect Sentinel", _AGENT_COLORS["dialect_sentinel"]),
                     ("Tone Cartographer", _AGENT_COLORS["tone_cartographer"])])
    st.markdown(
        f'<div style="margin-bottom:10px;">{leg} &nbsp;&nbsp;'
        f'<span class="sev-c">Critical</span> <span class="sev-w">Warning</span> '
        f'<span class="sev-i">Info</span> '
        f'<span style="color:#667;font-size:.78rem;">\u2014 line background = highest severity; '
        f'highlighted text = flagged phrase</span></div>',
        unsafe_allow_html=True)

    lines = script.split("\n")
    lm, unmatched = _attach_issues_to_lines(lines, filt)

    SEV_RANK = {"critical": 0, "hard": 0, "high": 0,
                "warning": 1, "medium": 1, "soft": 1}

    def _issue_note(i: dict) -> str:
        b = get_severity_badge(i.get("severity", "info"))
        ex = _esc(i.get("explanation") or i.get("summary") or "")
        dr = _esc(i.get("direction") or i.get("suggestion") or i.get("resolution_options") or "")
        chip = _agent_chip(i.get("_agent", ""))
        seg = i.get("scene_segment")
        seg_html = f'<span style="color:#8899BB;font-size:.78rem;">[{_esc(str(seg))}]</span> ' if seg else ""
        return (f'<div style="margin:2px 0 6px 36px;padding:6px 12px;background:rgba(255,255,255,.03);'
                f'border-radius:6px;font-size:.82rem;">'
                f'{b} {chip} {seg_html}{ex}'
                + (f'<br/><em style="color:var(--blu);">\u27A4 {dr}</em>' if dr else "")
                + "</div>")

    def _render_line(idx: int) -> list[str]:
        line = lines[idx - 1]
        issues = sorted(lm.get(idx, []),
                        key=lambda i: SEV_RANK.get((i.get("severity") or "").lower(), 2))
        out = []
        if issues:
            top = issues[0]
            w = {0: "critical", 1: "warning"}.get(
                SEV_RANK.get((top.get("severity") or "").lower(), 2), "info")
            cls = f"aline al-{w[0]}"
            sl = _hl_phrase(line, top.get("flagged_phrase", ""), top.get("severity", ""))
        else:
            cls = "aline al-ok" if line.strip() else "aline"
            sl = _esc(line)
        out.append(f'<div class="{cls}"><span style="color:#555;margin-right:8px;">{idx:>3}</span>{sl}</div>')
        for i in issues:
            out.append(_issue_note(i))
        return out

    hp = ['<div style="font-family:monospace;font-size:.88rem;line-height:1.6;">']
    if only_flagged:
        if not lm:
            hp.append('<div style="color:#667;padding:12px;">No flagged lines match the current filters.</div>')
        else:
            # flagged lines plus one line of context, with gap markers
            visible: set[int] = set()
            for ln in lm:
                visible.update({ln - 1, ln, ln + 1})
            visible = {v for v in visible if 1 <= v <= len(lines)}
            prev = None
            for idx in sorted(visible):
                if prev is not None and idx > prev + 1:
                    hp.append('<div style="color:#445;text-align:center;font-size:.8rem;'
                              'padding:2px 0;">\u22EF</div>')
                hp.extend(_render_line(idx))
                prev = idx
    else:
        for idx in range(1, len(lines) + 1):
            hp.extend(_render_line(idx))
    hp.append("</div>")
    st.markdown("".join(hp), unsafe_allow_html=True)

    # scene-level / unpinned issues — visible instead of silently dropped
    if unmatched:
        st.markdown(
            f'<div style="margin-top:14px;"><strong style="color:#F5C218;">'
            f'\U0001F4CC Scene-level &amp; unpinned issues ({len(unmatched)})</strong> '
            f'<span style="color:#667;font-size:.78rem;">\u2014 these flags could not be '
            f'matched to a single script line (scene-wide notes, or the agent\u2019s line '
            f'reference didn\u2019t match)</span></div>', unsafe_allow_html=True)
        blk = []
        for i in unmatched:
            qt = i.get("line_text", "")
            quote_html = (f'<div class="line-quote">{_hl_phrase(qt, i.get("flagged_phrase",""), i.get("severity",""))}</div>'
                          if qt else "")
            blk.append(_issue_note(i).replace('margin:2px 0 6px 36px;', 'margin:6px 0;') + quote_html)
        st.markdown("".join(blk), unsafe_allow_html=True)

def render_issues_table(issues: list[dict], columns: list[str]):
    if not issues:
        st.info("No issues found by this agent. \U0001F389"); return
    hdr = "".join(f"<th style='padding:8px 12px;text-align:left;border-bottom:1px solid #1A2A4A;color:#00AADD;'>{c.replace('_',' ')}</th>" for c in columns)
    rw = ""
    for iss in issues:
        cells = ""
        for c in columns:
            k = c.lower().replace(" ", "_")
            if k == "line":
                num = iss.get("line_number", "")
                qt = _hl_phrase(iss.get("line_text", ""), iss.get("flagged_phrase", ""),
                                iss.get("severity", ""))
                v_html = (f'<span style="color:#667;font-size:.75rem;">line {num}</span>'
                          f'<div class="line-quote">{qt}</div>') if qt.strip() else \
                         f'<span style="color:#667;">line {num}</span>'
                cells += f"<td style='padding:8px 12px;border-bottom:1px solid #0D1B2A;font-size:.85rem;min-width:240px;'>{v_html}</td>"
                continue
            if k == "direction":
                v = iss.get("direction") or iss.get("suggestion") or iss.get("resolution_options") or ""
            else:
                v = iss.get(k, "")
            cells += f"<td style='padding:8px 12px;border-bottom:1px solid #0D1B2A;font-size:.85rem;'>{get_severity_badge(str(v)) if k=='severity' else _esc(str(v))}</td>"
        rw += f"<tr>{cells}</tr>"
    st.markdown(
        f'<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;background:#0A1628;border-radius:8px;">'
        f"<tr>{hdr}</tr>{rw}</table></div>", unsafe_allow_html=True)

def render_character_reports(characters: list[str], results: dict):
    all_iss = _collect_all_issues(results)
    for char in characters:
        short = char.split()[0]
        ci = [i for i in all_iss if short.lower() in (i.get("character", "") or "").lower()
              or short.lower() in (i.get("line_text", "") or "").lower()]
        with st.expander(f"\U0001F464 {char}  ({len(ci)} issue{'s' if len(ci)!=1 else ''})"):
            pr = CHARACTER_PROFILES.get(char, "Profile not found.")
            st.markdown(
                f'<div class="nqa-card" style="border-left:3px solid #00AADD;">'
                f'<strong style="color:#00AADD;">Canonical Voice Profile</strong><br/>'
                f'<span style="color:#ccc;">{_esc(pr)}</span></div>', unsafe_allow_html=True)
            if not ci:
                st.success("No issues for this character."); continue
            # card per issue: quoted line with highlight, then verdict
            for i in ci:
                b = get_severity_badge(i.get("severity", "info"))
                chip = _agent_chip(i.get("_agent", ""))
                num = i.get("line_number", "?")
                qt = _hl_phrase(i.get("line_text", ""), i.get("flagged_phrase", ""),
                                i.get("severity", ""))
                ex = _esc(i.get("explanation") or "")
                dr = _esc(i.get("direction") or i.get("suggestion") or i.get("resolution_options") or "")
                it = _esc(str(i.get("issue_type", "")))
                st.markdown(
                    f'<div class="nqa-card" style="padding:10px 16px;margin:6px 0;">'
                    f'{b} {chip} <span style="color:#8899BB;font-size:.8rem;">{it} \u00B7 line {num}</span>'
                    f'<div class="line-quote">{qt}</div>'
                    f'<div style="color:#ccc;font-size:.85rem;margin-top:6px;">{ex}</div>'
                    + (f'<em style="color:var(--blu);font-size:.85rem;">\u27A4 {dr}</em>' if dr else "")
                    + "</div>", unsafe_allow_html=True)

def _issues_by_source_row(script: str, results: dict, line_map: dict) -> dict[int, list[dict]]:
    """Group attached issues by their ORIGINAL spreadsheet row ordinal using
    the extraction line_map (output line number -> 0-based data-row index)."""
    all_iss = _collect_all_issues(results)
    lines = script.split("\n")
    lm, _ = _attach_issues_to_lines(lines, all_iss)
    by_row: dict[int, list[dict]] = {}
    for ln, iss_list in lm.items():
        ridx = line_map.get(ln) if isinstance(line_map, dict) else None
        # JSON round-trips turn int keys into strings — accept both
        if ridx is None and isinstance(line_map, dict):
            ridx = line_map.get(str(ln))
        if ridx is None:
            continue
        by_row.setdefault(int(ridx), []).extend(iss_list)
    return by_row

def _row_annotation(iss_list: list[dict]) -> tuple[str, str, str]:
    """Collapse a row's issues into (worst severity, agent list, notes)."""
    _order = {"Critical": 0, "Warning": 1, "Info": 2}
    worst = min((_sev_bucket(i.get("severity")) for i in iss_list),
                key=lambda b: _order[b])
    agents = ", ".join(sorted({(i.get("_agent") or "").replace("_", " ").title()
                               for i in iss_list if i.get("_agent")}))
    notes = []
    for i in iss_list:
        ag = (i.get("_agent") or "").replace("_", " ").title()
        ex = (i.get("explanation") or "").strip()
        dr = (i.get("direction") or i.get("suggestion") or i.get("resolution_options") or "").strip()
        note = f"[{_sev_bucket(i.get('severity'))}/{ag}] {ex}"
        if dr:
            note += f" \u27A4 {dr}"
        notes.append(note)
    return worst, agents, " | ".join(notes)

_WB_HEADERS = ["NQA Severity", "NQA Agents", "NQA Note"]

def build_writeback_csv(src_bytes: bytes, by_row: dict[int, list[dict]]) -> bytes:
    """Return the original CSV with NQA columns appended, aligned by row."""
    raw = src_bytes
    if isinstance(raw, bytes):
        for _enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                raw = raw.decode(_enc); break
            except (UnicodeDecodeError, AttributeError):
                continue
        else:
            raw = raw.decode("utf-8", errors="replace")
    rows = list(csv.reader(io.StringIO(raw)))
    if not rows:
        return src_bytes
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]      # pad ragged rows
    rows[0].extend(_WB_HEADERS)
    for ridx, iss_list in by_row.items():
        target = ridx + 1                                    # +1 skips header
        if 1 <= target < len(rows):
            rows[target].extend(_row_annotation(iss_list))
    for i, r in enumerate(rows):
        if len(r) < width + 3:
            rows[i] = r + [""] * (width + 3 - len(r))
    buf = io.StringIO()
    csv.writer(buf).writerows(rows)
    return buf.getvalue().encode("utf-8-sig")               # Excel-friendly BOM

def build_writeback_xlsx(src_bytes: bytes, by_row: dict[int, list[dict]]) -> bytes | None:
    """Return the original XLSX with NQA columns appended on the first sheet."""
    if not HAS_OPENPYXL:
        return None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(src_bytes))
        ws = wb[wb.sheetnames[0]]
        base = ws.max_column
        for off, h in enumerate(_WB_HEADERS, 1):
            ws.cell(row=1, column=base + off, value=h)
        for ridx, iss_list in by_row.items():
            sev, agents, note = _row_annotation(iss_list)
            target = ridx + 2                                # +1 header, +1 1-based
            ws.cell(row=target, column=base + 1, value=sev)
            ws.cell(row=target, column=base + 2, value=agents)
            ws.cell(row=target, column=base + 3, value=note)
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    except Exception:
        return None

def build_csv_export(results: dict) -> str:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Agent","Line#","Line Text","Flagged Phrase","Character","Severity","Type","Explanation","Direction"])
    for ak in ("voice_keeper","lore_warden","dialect_sentinel","tone_cartographer"):
        for i in results.get(ak, {}).get("issues", []):
            w.writerow([ak, i.get("line_number",""), i.get("line_text",""),
                        i.get("flagged_phrase",""), i.get("character",""),
                        i.get("severity",""), i.get("issue_type",""),
                        i.get("explanation",""),
                        i.get("direction", i.get("suggestion", i.get("resolution_options","")))])
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        '<div style="text-align:center;padding:12px 0;">'
        '<span style="font-size:2.5rem;">\U0001F40D</span><br/>'
        '<strong style="font-size:1.3rem;color:#fff;">SNAKE NQA</strong><br/>'
        '<span style="color:#8899BB;font-size:.85rem;">Narrative Integrity Validator v1.0</span>'
        '</div>', unsafe_allow_html=True)
    st.divider()

    model_choice = DEFAULT_MODEL          # Claude Sonnet 4.6 on AWS Bedrock
    st.markdown(
        '<div style="font-size:.8rem;color:#8899BB;">Model<br/>'
        '<strong style="color:#00AADD;">Claude Sonnet 4.6</strong> '
        '<span style="color:#556;">on AWS Bedrock</span></div>',
        unsafe_allow_html=True)
    temperature = st.slider("Temperature", 0.0, 0.5, 0.15, 0.05,
        help="Controls randomness in the analysis. 0 = fully deterministic \u2014 the "
             "same script yields nearly identical flags every run (best for QA). "
             "Higher values explore subtler interpretations but vary more between runs.")

    st.divider()
    with st.expander("📖 Style Guide Reference"):
        st.caption("The canonical SNAKE style guide \u2014 every agent receives this verbatim with each analysis.")
        _sg = STYLE_GUIDE_TEXT.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        st.markdown(
            f'<div style="max-height:320px;overflow-y:auto;font-size:.78rem;line-height:1.5;'
            f'color:#ccc;white-space:pre-wrap;background:#0D1B2A;padding:10px;border-radius:8px;">{_sg}</div>',
            unsafe_allow_html=True)

    st.divider()
    st.markdown("**\U0001F4DC Analysis History**")
    _email_sb = st.session_state.get("auth_email", "")
    history = load_history(_email_sb) if _email_sb else st.session_state.get("history", [])
    if history:
        for _pos in range(len(history) - 1, max(-1, len(history) - 11), -1):
            h = history[_pos]
            sc = h.get("score", "?")
            # draft-over-draft delta vs the previous run of the same script
            delta_html = ""
            _hsh = h.get("script_hash")
            if _hsh:
                for _q in range(_pos - 1, -1, -1):
                    if history[_q].get("script_hash") == _hsh:
                        _prev = history[_q].get("score")
                        if isinstance(_prev, (int, float)) and isinstance(sc, (int, float)):
                            _d = sc - _prev
                            if _d > 0:   delta_html = f' <span style="color:var(--grn);font-weight:700;">\u25B2+{_d}</span>'
                            elif _d < 0: delta_html = f' <span style="color:var(--red);font-weight:700;">\u25BC{_d}</span>'
                            else:        delta_html = ' <span style="color:#777;">\u3003</span>'
                        break
            _lbl = h.get("script_label", "")
            st.markdown(
                f'<div style="padding:6px 8px;margin:4px 0;background:#0D1B2A;border-radius:6px;font-size:.78rem;">'
                f'<strong style="color:{score_color(int(sc) if isinstance(sc,(int,float)) else 0)}">{sc}/100</strong>{delta_html}'
                + (f' &middot; <span style="color:#ccc;">{_lbl}</span>' if _lbl else "")
                + f'<br/><span style="color:#556;">{h.get("timestamp","")}'
                + (f' &middot; {", ".join(h.get("characters", [])[:3])}' if h.get("characters") else "")
                + '</span></div>', unsafe_allow_html=True)
        st.caption("\u25B2\u25BC = change vs the previous run of the same script.")
    else:
        st.caption("No analyses yet for this account.")

    st.divider()
    st.markdown(
        f'<span style="color:#8899BB;font-size:.8rem;">Signed in as<br/>'
        f'<strong style="color:#00AADD;">{st.session_state.get("auth_email","")}</strong></span>',
        unsafe_allow_html=True)
    if st.button("\U0001F6AA Sign Out", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.query_params.clear()
        st.rerun()

    st.divider()
    st.caption("Built for SEGA\u2019s narrative team.\nPowered by Claude on AWS Bedrock.\nAgents advise, humans decide.")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(
    '<div class="hero-t">\U0001F40D SNAKE Narrative Integrity Validator</div>'
    '<div class="hero-s">Multi-agent quality assurance for narrative-driven game scripts</div><br/>',
    unsafe_allow_html=True)

# ── script input ────────────────────────────────────────────────────────────
st.markdown('<div class="nqa-card">', unsafe_allow_html=True)
st.subheader("Script Input")
tab_paste, tab_upload = st.tabs(["\u270D\uFE0F Paste Script", "\U0001F4C1 Upload File"])
script_text = ""
with tab_paste:
    sample_keys = ["\u2014 Custom \u2014"] + list(SAMPLE_SCRIPTS.keys())
    sample_choice = st.selectbox("Load a sample script", sample_keys)
    default_val = SAMPLE_SCRIPTS.get(sample_choice, "")
    script_text = st.text_area("Paste script", value=default_val, height=320, key="si_paste")
with tab_upload:
    _missing = []
    if not HAS_PYPDF:    _missing.append("pypdf (for PDF)")
    if not HAS_DOCX:     _missing.append("python-docx (for DOCX)")
    if not HAS_OPENPYXL: _missing.append("openpyxl (for XLSX)")
    if _missing:
        st.caption(f"⚠️ Optional packages not installed: {', '.join(_missing)}")
    uploaded = st.file_uploader(
        "Upload .txt, .pdf, .docx, .xlsx, or .csv",
        type=["txt", "pdf", "docx", "xlsx", "csv"],
    )
    if uploaded is not None:
        fname = uploaded.name.lower()

        # ── JB PASS toggle (shown before extraction so it gates the parse) ──
        # We need to peek at headers without consuming the buffer, so we do a
        # lightweight pre-check for CSV/XLSX before the full parse.
        _show_jb_toggle = False
        if fname.endswith((".csv", ".xlsx")):
            import io as _io, csv as _csv
            _preview = uploaded.read(4096)
            uploaded.seek(0)          # reset for the real read below
            if isinstance(_preview, bytes):
                _preview = _preview.decode("utf-8", errors="replace")
            _first_line = _preview.splitlines()[0] if _preview else ""
            if "JB PASS" in _first_line:
                _show_jb_toggle = True

        use_jb = False
        if _show_jb_toggle:
            use_jb = st.toggle(
                "Use **JB PASS** column (editorial revisions)",
                value=False,
                help="When on, dialogue is taken from the 'JB PASS' column instead of "
                     "'Text'. Lines without a JB PASS entry keep the original text.",
            )

        if fname.endswith(".pdf") and HAS_PYPDF:
            script_text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(uploaded).pages)
        elif fname.endswith(".pdf"):
            st.warning("pypdf is not installed; falling back to raw read. Install with: pip install pypdf")
            script_text = uploaded.read().decode("utf-8", errors="replace")
        elif fname.endswith(".docx"):
            script_text = extract_docx_text(uploaded)
        elif fname.endswith(".xlsx"):
            _src_bytes = uploaded.getvalue()
            script_text, _has_jb, _line_map = extract_xlsx_script(uploaded, use_jb_pass=use_jb)
            st.session_state["upload_source"] = {
                "bytes": _src_bytes, "name": uploaded.name, "kind": "xlsx",
                "line_map": _line_map, "extracted_text": script_text} if _line_map else None
        elif fname.endswith(".csv"):
            _src_bytes = uploaded.getvalue()
            script_text, _has_jb, _line_map = extract_csv_script(uploaded, use_jb_pass=use_jb)
            st.session_state["upload_source"] = {
                "bytes": _src_bytes, "name": uploaded.name, "kind": "csv",
                "line_map": _line_map, "extracted_text": script_text} if _line_map else None
        else:
            script_text = uploaded.read().decode("utf-8", errors="replace")
            st.session_state["upload_source"] = None

        if fname.endswith((".csv", ".xlsx")) and _show_jb_toggle:
            col_label = "JB PASS" if use_jb else "Text"
            st.caption(f"📄 Extracted using **{col_label}** column — toggle above to switch")

        st.text_area("Extracted (editable)", script_text, height=320, key="si_upload")
st.markdown("</div>", unsafe_allow_html=True)

# ── character detection ─────────────────────────────────────────────────────
characters: list[str] = []
if script_text.strip():
    characters = detect_characters(script_text)
    if characters:
        tags = " ".join(f'<span class="char-tag">{c}</span>' for c in characters)
        st.markdown(
            f'<div class="nqa-card"><strong style="color:#8899BB;">Detected Characters</strong><br/>{tags}</div>',
            unsafe_allow_html=True)

# ── run / clear ─────────────────────────────────────────────────────────────
if script_text.strip():
    c1, c2, _ = st.columns([2,2,6])
    run_btn = c1.button("\U0001F50D Run Full Analysis", type="primary", use_container_width=True)
    if c2.button("\U0001F5D1 Clear", use_container_width=True):
        st.session_state.pop("results", None); st.rerun()

    if run_btn:
        client = get_client()
        if client is None:
            st.error("No LLM credentials. Add AWS_BEDROCK_* or ANTHROPIC_API_KEY to secrets.")
            st.stop()
        statuses = {n:"pending" for n in AGENT_META}
        ph = st.empty()
        ph.markdown(render_agent_progress(statuses), unsafe_allow_html=True)
        ar: dict[str, Any] = {}

        _agent_fns = {
            "Voice Keeper":      lambda: run_voice_keeper(client, model_choice, script_text, characters, temperature),
            "Lore Warden":       lambda: run_lore_warden(client, model_choice, script_text, characters, temperature),
            "Dialect Sentinel":  lambda: run_dialect_sentinel(client, model_choice, script_text, characters, temperature),
            "Tone Cartographer": lambda: run_tone_cartographer(client, model_choice, script_text, characters, temperature),
        }
        _timings: dict[str, float] = {}
        _t0 = time.time()

        def _record(agent_name: str, result):
            key = AGENT_META[agent_name]["key"]
            if not isinstance(result, dict):
                result = {"parse_error": True, "raw_text": str(result)}
            ar[key] = result
            statuses[agent_name] = "complete" if not result.get("parse_error") else "error"
            _timings[agent_name] = time.time() - _t0
            ph.markdown(render_agent_progress(statuses, _timings), unsafe_allow_html=True)

        # All four agents run fully concurrently. The shared system block is
        # still cache_control-marked, so re-runs of the same script within the
        # cache TTL get cached-input pricing; within a single first run the
        # four concurrent calls each pay full input cost (the cache entry is
        # only readable after the first request finishes). This trades some
        # first-run cost for the fastest possible wall-clock time.
        for n in _agent_fns:
            statuses[n] = "running"
        ph.markdown(render_agent_progress(statuses, _timings), unsafe_allow_html=True)
        with ThreadPoolExecutor(max_workers=4) as _ex:
            _futs = {_ex.submit(fn): n for n, fn in _agent_fns.items()}
            for _fut in as_completed(_futs):
                _n = _futs[_fut]
                try:
                    _res = _fut.result()
                except Exception as _e:
                    _res = {"parse_error": True, "raw_text": str(_e)}
                _record(_n, _res)

        statuses["Arbiter"] = "running"
        ph.markdown(render_agent_progress(statuses, _timings), unsafe_allow_html=True)
        _t_arb = time.time()
        ar["arbiter"] = run_arbiter(client, model_choice, ar, temperature)
        _timings["Arbiter"] = time.time() - _t_arb
        _arb = ar["arbiter"]
        statuses["Arbiter"] = (
            "complete" if isinstance(_arb, dict) and not _arb.get("parse_error") else "error"
        )
        ph.markdown(render_agent_progress(statuses, _timings), unsafe_allow_html=True)

        st.session_state["results"] = ar
        st.session_state["result_script"] = script_text
        st.session_state["result_characters"] = characters
        # keep the source-file mapping only if the analysed text is exactly
        # what was extracted (edits invalidate row alignment)
        _us = st.session_state.get("upload_source")
        if _us and _us.get("extracted_text") == script_text:
            st.session_state["result_source"] = _us
        else:
            st.session_state.pop("result_source", None)
        _arb_hist = ar.get("arbiter", {})
        if not isinstance(_arb_hist, dict):
            _arb_hist = {}
        _hist_entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "score": _arb_hist.get("overall_score", "?"),
            "characters": characters,
            "script_hash": _script_hash(script_text),
            "script_label": _script_label(script_text),
            "critical": _arb_hist.get("critical_count", 0),
            "warning": _arb_hist.get("warning_count", 0),
        }
        _email_run = st.session_state.get("auth_email", "")
        if _email_run:
            append_history(_email_run, _hist_entry)
        st.session_state.setdefault("history", []).append(_hist_entry)


# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
def _safe_dict(v: Any) -> dict:
    """Return v if it is a dict, else an empty dict. Guards against the LLM
    returning a JSON array instead of an object."""
    return v if isinstance(v, dict) else {}

if "results" in st.session_state:
    res = st.session_state["results"]
    vk = _safe_dict(res.get("voice_keeper")); lw = _safe_dict(res.get("lore_warden"))
    ds = _safe_dict(res.get("dialect_sentinel")); tc = _safe_dict(res.get("tone_cartographer"))
    arb = _safe_dict(res.get("arbiter")); chars = st.session_state.get("result_characters",[])
    scr = st.session_state.get("result_script","")

    st.divider()
    st.markdown("## \U0001F4CA Analysis Results")
    render_score_cards(arb, vk, lw, ds, tc)

    crit = arb.get("critical_count",0); warn = arb.get("warning_count",0)
    inf = arb.get("info_count",0); total = crit+warn+inf
    st.markdown(
        f'<div class="nqa-card" style="text-align:center;">'
        f'<strong style="font-size:1.1rem;">{total} issues found</strong> &nbsp; '
        f'<span class="sev-c">{crit} Critical</span> &nbsp; '
        f'<span class="sev-w">{warn} Warning</span> &nbsp; '
        f'<span class="sev-i">{inf} Info</span></div>', unsafe_allow_html=True)

    t1,t2,t3,t4,t5 = st.tabs(["\U0001F4DD Annotated Script","\U0001F3AD Agent Reports",
        "\U0001F464 Character Analysis","\U0001F4CA Arbiter Summary","\U0001F4BE Export"])

    with t1:
        render_annotated_script(scr, res)

    with t2:
        # surface parse failures so errors are debuggable instead of silent
        for _nm, _d in [("Voice Keeper", vk), ("Lore Warden", lw),
                        ("Dialect Sentinel", ds), ("Tone Cartographer", tc),
                        ("Arbiter", arb)]:
            if _d.get("parse_error"):
                st.error(f"\u26A0\uFE0F {_nm} returned a response that couldn't be parsed as JSON. "
                         "Its results are missing from this analysis \u2014 try re-running.")
                with st.expander(f"Raw {_nm} response (for debugging)"):
                    st.code(str(_d.get("raw_text", ""))[:6000], language="text")
            elif _d.get("_truncation_repaired"):
                st.warning(f"\u2702\uFE0F {_nm}'s response was cut off mid-stream and automatically "
                           "repaired \u2014 the tail end of its issue list may be missing.")
        with st.expander("\U0001F3AD Voice Keeper", expanded=True):
            st.markdown(f"**Voice Score: {vk.get('voice_score','?')}**/100")
            render_issues_table(vk.get("issues",[]),
                ["Line","Character","Issue_Type","Severity","Explanation","Direction"])
        with st.expander("\U0001F4DC Lore Warden"):
            st.markdown(f"**Lore Score: {lw.get('lore_score','?')}**/100")
            render_issues_table(lw.get("issues",[]),
                ["Line","Issue_Type","Severity","Explanation","Resolution_Options"])
        with st.expander("\U0001F5E3 Dialect Sentinel"):
            st.markdown(f"**Dialect Score: {ds.get('dialect_score','?')}**/100")
            render_issues_table(ds.get("issues",[]),
                ["Line","Character","Issue_Type","Severity","Explanation","Suggestion"])
        with st.expander("\U0001F3B5 Tone Cartographer"):
            st.markdown(f"**Tone Score: {tc.get('tone_score','?')}**/100")
            for seg in tc.get("tone_map",[]):
                st.markdown(
                    f'<div class="nqa-card" style="padding:10px 14px;margin:4px 0;">'
                    f'<strong style="color:#F5C218;">{seg.get("scene_segment","")}</strong><br/>'
                    f'<span style="color:#8899BB;">Tone:</span> {seg.get("tone_description","")}<br/>'
                    f'<span style="color:#8899BB;">Effect:</span> {seg.get("intended_effect","")}'
                    f'</div>', unsafe_allow_html=True)
            render_issues_table(tc.get("issues",[]),
                ["Scene_Segment","Issue_Type","Severity","Explanation"])

    with t3:
        render_character_reports(chars, res)

    with t4:
        nhs = arb.get("narrative_health_summary","")
        if nhs:
            st.markdown(
                f'<div class="nqa-card"><strong style="color:#00AADD;">Narrative Health Summary</strong>'
                f'<p style="color:#ccc;line-height:1.7;font-size:.95rem;">{nhs}</p></div>',
                unsafe_allow_html=True)
        for item in arb.get("priority_list",[]):
            b = get_severity_badge(item.get("severity","info"))
            src = item.get("agent_source","").replace("_"," ").title()
            st.markdown(
                f'<div class="nqa-card" style="padding:10px 16px;margin:6px 0;">'
                f'<strong style="color:#F5C218;">#{item.get("rank","?")}</strong> {b} '
                f'<strong style="color:#aaa;">[{src}]</strong> {item.get("summary","")}'
                + (f'<br/><em style="color:var(--blu);">\u27A4 {item.get("recommendation","")}</em>'
                   if item.get("recommendation") else "")
                + "</div>", unsafe_allow_html=True)
        for upd in arb.get("character_bible_updates",[]):
            st.markdown(
                f'<div class="nqa-card" style="border-left:3px solid #F5C218;padding:10px 16px;">'
                f'<strong style="color:#F5C218;">{upd.get("character","")}</strong><br/>'
                f'<span style="color:#ccc;">{upd.get("suggested_update","")}</span></div>',
                unsafe_allow_html=True)

    with t5:
        st.markdown("### Download Analysis")
        c1, c2, c3, _ = st.columns([2, 2, 3, 3])
        c1.download_button("\u2B07 CSV", build_csv_export(res),
            file_name=f"snake_nqa_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv", use_container_width=True)
        c2.download_button("\u2B07 JSON", json.dumps(res, indent=2, default=str),
            file_name=f"snake_nqa_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json", use_container_width=True)

        _src = st.session_state.get("result_source")
        if _src and _src.get("line_map"):
            _by_row = _issues_by_source_row(scr, res, _src["line_map"])
            _stem = os.path.splitext(_src["name"])[0]
            if _src["kind"] == "csv":
                _wb = build_writeback_csv(_src["bytes"], _by_row)
                c3.download_button("\u2B07 Annotated source CSV", _wb,
                    file_name=f"{_stem}_NQA.csv", mime="text/csv",
                    use_container_width=True,
                    help="Your original file with NQA Severity / Agents / Note "
                         "columns appended to the flagged rows.")
            else:
                _wb = build_writeback_xlsx(_src["bytes"], _by_row)
                if _wb:
                    c3.download_button("\u2B07 Annotated source XLSX", _wb,
                        file_name=f"{_stem}_NQA.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="Your original workbook with NQA Severity / Agents / Note "
                             "columns appended to the flagged rows.")
            st.caption(f"\U0001F4CE Write-back: {len(_by_row)} source row"
                       f"{'s' if len(_by_row) != 1 else ''} annotated in {_src['name']}")
        else:
            st.caption("\U0001F4A1 Write-back export (your original CSV/XLSX with NQA columns "
                       "added) is available when you analyse an uploaded script sheet without "
                       "editing the extracted text.")
        with st.expander("Preview JSON"):
            st.json(res)
