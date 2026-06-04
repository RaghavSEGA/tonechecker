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


def parse_llm_json(raw: str) -> Any:
    raw = raw.strip()
    try: return json.loads(raw)
    except Exception: pass
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", raw, re.S)
    if m:
        try: return json.loads(m.group(1))
        except Exception: pass
    for co, cc in [("{", "}"), ("[", "]")]:
        idx = raw.find(co)
        if idx != -1:
            d = 0
            for i in range(idx, len(raw)):
                if raw[i] == co: d += 1
                elif raw[i] == cc: d -= 1
                if d == 0:
                    try: return json.loads(raw[idx:i+1])
                    except Exception: break
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


def get_severity_badge(severity: str) -> str:
    s = (severity or "info").lower().strip()
    if s in ("critical","hard","high","red"):   return '<span class="sev-c">CRITICAL</span>'
    if s in ("warning","medium","soft","yellow"): return '<span class="sev-w">WARNING</span>'
    return '<span class="sev-i">INFO</span>'


def score_color(score: int) -> str:
    if score >= 80: return "#00BB66"
    if score >= 60: return "#F5C218"
    return "#CC2244"


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
               temperature: float = 0.15) -> str:
    try:
        resp = client.messages.create(model=model, max_tokens=4096,
            temperature=temperature, system=system_prompt,
            messages=[{"role":"user","content":user_prompt}])
        return resp.content[0].text
    except Exception as exc:
        return json.dumps({"error": str(exc)})

def relevant_profiles(characters: list[str]) -> str:
    lines = [f"### {c}\n{CHARACTER_PROFILES[c]}\n" for c in characters if c in CHARACTER_PROFILES]
    return "\n".join(lines) if lines else "(No specific profiles matched.)"


# ── agent system prompts ────────────────────────────────────────────────────

_VK_SYS = """\
You are VOICE KEEPER for the game SNAKE. Check every dialogue line against CHARACTER PROFILES.
Check: vocabulary fingerprint, sentence-construction, emotional register, code-switching, tonal arc.
RULES: nobody speaks in exposition; profanity is earned; speech rhythms are NOT interchangeable; avoid over-explanation, self-pity, clich\u00e9.
OUTPUT valid JSON ONLY:
{"voice_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact>","character":"<NAME>","issue_type":"<vocabulary_deviation|register_mismatch|emotional_register|code_switching|tonal_arc|exposition>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","direction":"<suggestion, do NOT rewrite>"}]}"""

_LW_SYS = """\
You are LORE WARDEN for SNAKE. Identify any contradiction with established canon, world rules, character histories or timeline.
Check: factual contradictions, timeline violations, world-rule violations, character knowledge state, relationship state, absent character references.
OUTPUT valid JSON ONLY:
{"lore_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact>","issue_type":"<factual_contradiction|timeline_violation|world_rule|knowledge_state|relationship_state|absent_character>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","resolution_options":"<fixes>"}]}"""

_DS_SYS = """\
You are DIALECT SENTINEL for SNAKE. Protect linguistic texture.
Check: faction/region slang, idiom coherence (no real-world idioms that break setting), profanity-system integrity, politeness markers, cultural register, regional Spanish (Cielo = Paraguayan).
OUTPUT valid JSON ONLY:
{"dialect_score":<0-100>,"issues":[{"line_number":<int>,"line_text":"<exact>","character":"<NAME>","issue_type":"<slang_inconsistency|idiom_violation|profanity_misuse|formality_error|cultural_register|language_variant>","severity":"<critical|warning|info>","explanation":"<2-3 sent>","suggestion":"<direction>"}]}"""

_TC_SYS = """\
You are TONE CARTOGRAPHER for SNAKE. Analyse scene-level emotional trajectory.
Check: tonal whiplash, escalation/de-escalation curves, thematic consistency, clich\u00e9s, pacing.
OUTPUT valid JSON ONLY:
{"tone_score":<0-100>,"tone_map":[{"scene_segment":"<beat>","tone_description":"<e.g. restrained tension>","intended_effect":"<player feeling>"}],"issues":[{"scene_segment":"<part>","issue_type":"<tonal_whiplash|flat_escalation|thematic_inconsistency|cliche|pacing>","severity":"<critical|warning|info>","explanation":"<2-3 sent>"}]}"""

_ARB_SYS = """\
You are ARBITER for SNAKE narrative QA. You receive outputs from Voice Keeper, Lore Warden, Dialect Sentinel, Tone Cartographer.
1. Merge & deduplicate overlapping flags.  2. Rank by severity & narrative impact.
3. Identify cross-agent conflicts.  4. Detect patterns (character flagged repeatedly = possible intentional evolution).
OUTPUT valid JSON ONLY:
{"overall_score":<0-100>,"critical_count":<int>,"warning_count":<int>,"info_count":<int>,
"priority_list":[{"rank":<int>,"agent_source":"<name>","severity":"<lvl>","summary":"<1 sent>","recommendation":"<action>"}],
"narrative_health_summary":"<2-3 paragraph assessment>",
"character_bible_updates":[{"character":"<NAME>","suggested_update":"<reason & suggestion>"}]}"""

# ── agent runners ───────────────────────────────────────────────────────────

def _build_user_prompt(script: str, characters: list[str]) -> str:
    return (f"## STYLE GUIDE\n{STYLE_GUIDE_TEXT}\n\n"
            f"## DETECTED CHARACTER PROFILES\n{relevant_profiles(characters)}\n\n"
            f"## SCRIPT TO ANALYSE\n```\n{script}\n```\n\n"
            "Return ONLY valid JSON matching your schema.")

def run_voice_keeper(client, model, script, characters, temp):
    return parse_llm_json(call_agent(client, model, _VK_SYS,
                                     _build_user_prompt(script, characters), temp))

def run_lore_warden(client, model, script, characters, temp):
    return parse_llm_json(call_agent(client, model, _LW_SYS,
                                     _build_user_prompt(script, characters), temp))

def run_dialect_sentinel(client, model, script, characters, temp):
    return parse_llm_json(call_agent(client, model, _DS_SYS,
                                     _build_user_prompt(script, characters), temp))

def run_tone_cartographer(client, model, script, characters, temp):
    return parse_llm_json(call_agent(client, model, _TC_SYS,
                                     _build_user_prompt(script, characters), temp))

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

def render_agent_progress(statuses: dict[str, str]) -> str:
    rows = []
    for name, meta in AGENT_META.items():
        sv = statuses.get(name, "pending")
        lb = {"pending":"Waiting\u2026","running":"Analysing\u2026",
              "complete":"Done \u2714","error":"Error \u2718"}.get(sv, sv)
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

def render_annotated_script(script: str, results: dict):
    all_iss = _collect_all_issues(results)
    lines = script.split("\n")
    lm: dict[int, list[dict]] = {}
    for iss in all_iss:
        ln = iss.get("line_number")
        if isinstance(ln, int) and 0 < ln <= len(lines):
            lm.setdefault(ln, []).append(iss)
    for iss in all_iss:
        lt = (iss.get("line_text") or "").strip()
        if lt and iss.get("line_number") is None:
            for idx, raw in enumerate(lines, 1):
                if lt.lower() in raw.lower():
                    lm.setdefault(idx, []).append(iss); break
    hp = ['<div style="font-family:monospace;font-size:.88rem;line-height:1.6;">']
    for idx, line in enumerate(lines, 1):
        issues = lm.get(idx, [])
        if issues:
            w = "info"
            for i in issues:
                s = (i.get("severity") or "").lower()
                if s in ("critical","hard","high"): w = "critical"; break
                if s in ("warning","medium","soft"): w = "warning"
            cls = f"aline al-{w[0]}"
        else:
            cls = "aline al-ok" if line.strip() else "aline"
        sl = line.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        hp.append(f'<div class="{cls}"><span style="color:#555;margin-right:8px;">{idx:>3}</span>{sl}</div>')
        for i in issues:
            b = get_severity_badge(i.get("severity","info"))
            ex = (i.get("explanation") or i.get("summary") or "").replace("<","&lt;")
            dr = (i.get("direction") or i.get("suggestion") or "").replace("<","&lt;")
            ag = i.get("_agent","").replace("_"," ").title()
            hp.append(
                f'<div style="margin:2px 0 6px 36px;padding:6px 12px;background:rgba(255,255,255,.03);'
                f'border-radius:6px;font-size:.82rem;">'
                f'{b} <strong style="color:#aaa;">[{ag}]</strong> {ex}'
                + (f'<br/><em style="color:var(--blu);">\u27A4 {dr}</em>' if dr else "")
                + "</div>")
    hp.append("</div>")
    st.markdown("".join(hp), unsafe_allow_html=True)

def render_issues_table(issues: list[dict], columns: list[str]):
    if not issues:
        st.info("No issues found by this agent. \U0001F389"); return
    hdr = "".join(f"<th style='padding:8px 12px;text-align:left;border-bottom:1px solid #1A2A4A;color:#00AADD;'>{c}</th>" for c in columns)
    rw = ""
    for iss in issues:
        cells = ""
        for c in columns:
            k = c.lower().replace(" ","_")
            v = iss.get(k, "")
            cells += f"<td style='padding:8px 12px;border-bottom:1px solid #0D1B2A;font-size:.85rem;'>{get_severity_badge(str(v)) if k=='severity' else str(v).replace('<','&lt;')}</td>"
        rw += f"<tr>{cells}</tr>"
    st.markdown(
        f'<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;background:#0A1628;border-radius:8px;">'
        f"<tr>{hdr}</tr>{rw}</table></div>", unsafe_allow_html=True)

def render_character_reports(characters: list[str], results: dict):
    all_iss = _collect_all_issues(results)
    for char in characters:
        short = char.split()[0]
        ci = [i for i in all_iss if short.lower() in (i.get("character","") or "").lower()
              or short.lower() in (i.get("line_text","") or "").lower()]
        with st.expander(f"\U0001F464 {char}  ({len(ci)} issue{'s' if len(ci)!=1 else ''})"):
            pr = CHARACTER_PROFILES.get(char, "Profile not found.")
            st.markdown(
                f'<div class="nqa-card" style="border-left:3px solid #00AADD;">'
                f'<strong style="color:#00AADD;">Canonical Voice Profile</strong><br/>'
                f'<span style="color:#ccc;">{pr}</span></div>', unsafe_allow_html=True)
            if ci: render_issues_table(ci, ["Severity","Issue_Type","Explanation","Direction"])
            else: st.success("No issues for this character.")

def build_csv_export(results: dict) -> str:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Agent","Line#","Character","Severity","Type","Explanation","Direction"])
    for ak in ("voice_keeper","lore_warden","dialect_sentinel","tone_cartographer"):
        for i in results.get(ak, {}).get("issues", []):
            w.writerow([ak, i.get("line_number",""), i.get("character",""),
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

    model_choice = st.selectbox("LLM model", [DEFAULT_MODEL,
        "us.anthropic.claude-sonnet-4-20250514",
        "anthropic.claude-3-5-sonnet-20241022-v2:0"], index=0)
    temperature = st.slider("Temperature", 0.0, 0.5, 0.15, 0.05)

    st.divider()
    with st.expander("\U0001F4D6 Style Guide Reference"):
        st.text_area("guide", STYLE_GUIDE_TEXT, height=300, disabled=True, label_visibility="collapsed")

    st.divider()
    st.markdown("**\U0001F4DC Analysis History**")
    history = st.session_state.get("history", [])
    if history:
        for h in reversed(history[-10:]):
            sc = h.get("score", "?")
            st.markdown(
                f'<div style="padding:6px;margin:4px 0;background:#0D1B2A;border-radius:6px;font-size:.8rem;">'
                f'<strong style="color:{score_color(int(sc) if isinstance(sc,(int,float)) else 0)}">{sc}/100</strong>'
                f' &middot; {", ".join(h.get("characters",[])[:3])} &middot; '
                f'<span style="color:#555;">{h.get("timestamp","")}</span></div>', unsafe_allow_html=True)
    else:
        st.caption("No analyses yet.")

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
        "Upload .txt, .pdf, .docx, or .xlsx",
        type=["txt", "pdf", "docx", "xlsx"],
    )
    if uploaded is not None:
        fname = uploaded.name.lower()
        if fname.endswith(".pdf") and HAS_PYPDF:
            script_text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(uploaded).pages)
        elif fname.endswith(".pdf"):
            st.warning("pypdf is not installed; falling back to raw read. Install with: pip install pypdf")
            script_text = uploaded.read().decode("utf-8", errors="replace")
        elif fname.endswith(".docx"):
            script_text = extract_docx_text(uploaded)
        elif fname.endswith(".xlsx"):
            script_text = extract_xlsx_text(uploaded)
        else:
            script_text = uploaded.read().decode("utf-8", errors="replace")
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

        for agent_name, runner in [
            ("Voice Keeper",      lambda: run_voice_keeper(client, model_choice, script_text, characters, temperature)),
            ("Lore Warden",       lambda: run_lore_warden(client, model_choice, script_text, characters, temperature)),
            ("Dialect Sentinel",  lambda: run_dialect_sentinel(client, model_choice, script_text, characters, temperature)),
            ("Tone Cartographer", lambda: run_tone_cartographer(client, model_choice, script_text, characters, temperature)),
        ]:
            statuses[agent_name] = "running"
            ph.markdown(render_agent_progress(statuses), unsafe_allow_html=True)
            key = AGENT_META[agent_name]["key"]
            ar[key] = runner()
            _res = ar[key]
            statuses[agent_name] = (
                "complete" if isinstance(_res, dict) and not _res.get("parse_error") else "error"
            )
            ph.markdown(render_agent_progress(statuses), unsafe_allow_html=True)

        statuses["Arbiter"] = "running"
        ph.markdown(render_agent_progress(statuses), unsafe_allow_html=True)
        ar["arbiter"] = run_arbiter(client, model_choice, ar, temperature)
        _arb = ar["arbiter"]
        statuses["Arbiter"] = (
            "complete" if isinstance(_arb, dict) and not _arb.get("parse_error") else "error"
        )
        ph.markdown(render_agent_progress(statuses), unsafe_allow_html=True)

        st.session_state["results"] = ar
        st.session_state["result_script"] = script_text
        st.session_state["result_characters"] = characters
        st.session_state.setdefault("history", []).append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "score": ar.get("arbiter",{}).get("overall_score","?"),
            "characters": characters})


# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
if "results" in st.session_state:
    res = st.session_state["results"]
    vk = res.get("voice_keeper",{}); lw = res.get("lore_warden",{})
    ds = res.get("dialect_sentinel",{}); tc = res.get("tone_cartographer",{})
    arb = res.get("arbiter",{}); chars = st.session_state.get("result_characters",[])
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
        with st.expander("\U0001F3AD Voice Keeper", expanded=True):
            st.markdown(f"**Voice Score: {vk.get('voice_score','?')}**/100")
            render_issues_table(vk.get("issues",[]),
                ["Line_Number","Character","Issue_Type","Severity","Explanation","Direction"])
        with st.expander("\U0001F4DC Lore Warden"):
            st.markdown(f"**Lore Score: {lw.get('lore_score','?')}**/100")
            render_issues_table(lw.get("issues",[]),
                ["Line_Number","Issue_Type","Severity","Explanation","Resolution_Options"])
        with st.expander("\U0001F5E3 Dialect Sentinel"):
            st.markdown(f"**Dialect Score: {ds.get('dialect_score','?')}**/100")
            render_issues_table(ds.get("issues",[]),
                ["Line_Number","Character","Issue_Type","Severity","Explanation","Suggestion"])
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
        c1, c2, _ = st.columns([2,2,6])
        c1.download_button("\u2B07 CSV", build_csv_export(res),
            file_name=f"snake_nqa_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv", use_container_width=True)
        c2.download_button("\u2B07 JSON", json.dumps(res, indent=2, default=str),
            file_name=f"snake_nqa_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json", use_container_width=True)
        with st.expander("Preview JSON"):
            st.json(res)
