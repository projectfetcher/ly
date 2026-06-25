import re
import time
import base64
import json
import hashlib
import logging
import os
from datetime import datetime, timedelta
from urllib.parse import urlparse, unquote, quote_plus

import requests
from bs4 import BeautifulSoup
import openpyxl
import pandas as pd

# Optional: load secrets from a local .env file if python-dotenv is installed.
# This is silent/no-op if the package or file isn't present.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ─────────────────────────────────────────────────────────────────────────────
#  OPTIONAL heavy deps
# ─────────────────────────────────────────────────────────────────────────────
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

SHEET_NAME       = "Sheet1"
DELAY_S          = 2.0
FETCH_CHAR_LIMIT = 120_000

MAX_PAGES       = 0   # 0 = unlimited
MAX_EMPTY_PAGES = 5
JOB_LIMIT       = 0   # 0 = no cap

OUTPUT_FILE        = "jobs_output.xlsx"
PROCESSED_IDS_FILE = "processed_jobs.csv"

# ── WordPress (secrets via environment variables — see header docstring) ────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE        = WP_URL.rstrip("/")
WP_JOBS_URL    = f"{WP_BASE}/job-listings"     # ✅
WP_COMPANY_URL = f"{WP_BASE}/companies"        # ✅
WP_MEDIA_URL   = f"{WP_BASE}/media"

# ── Mistral (secret via environment variable — see header docstring) ────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True   # set False to skip paraphrasing entirely

# ── Startup checks: warn (don't crash) if secrets are missing ───────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

# =============================================================================
#  KEYWORDS
# =============================================================================

SEARCH_KEYWORDS = [
    "",
   # "engineer", "developer", "manager", "finance", "sales", "HR",
   # "doctor", "construction", "logistics", "operations", "customer service",
   # "teacher", "chef", "lawyer", "graphic designer", "production manager",
   # "petroleum", "driver", "security", "researcher", "journalist",
   # "banker", "retail", "renewable energy",
]

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

import sys
_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 72)

# =============================================================================
#  ROTATING USER-AGENTS
# =============================================================================

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
]
_ua_idx = 0

def _next_headers() -> dict:
    global _ua_idx
    ua = USER_AGENTS[_ua_idx % len(USER_AGENTS)]
    _ua_idx += 1
    return {
        "User-Agent":       ua,
        "Accept":           "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language":  "en-US,en;q=0.9",
        "Accept-Encoding":  "gzip, deflate, br",
        "Cache-Control":    "no-cache",
        "X-Li-Lang":        "en_US",
        "X-Requested-With": "XMLHttpRequest",
    }

HEADERS = _next_headers()

# =============================================================================
#  DOMAIN LISTS
# =============================================================================

SKIP_CRAWL_DOMAINS = [
    "dhl.com","fedex.com","ups.com","amazon.com","amazon.jobs",
    "google.com","microsoft.com","apple.com","meta.com","ibm.com",
    "oracle.com","sap.com","accenture.com","deloitte.com","pwc.com",
    "kpmg.com","ey.com","mckinsey.com","bcg.com","bain.com",
    "citibank.com","hsbc.com","barclays.com","bnpparibas.com",
    "airbus.com","boeing.com","siemens.com","ge.com",
    "unilever.com","nestle.com","pg.com","shell.com","bp.com",
]
BAD_DOMAINS = [
    "linkedin.com","google.com","youtube.com","facebook.com",
    "twitter.com","x.com","instagram.com","t.co","example.com",
    "w3.org","sentry.io","schema.org",
]
NOISE_EMAIL_DOMAINS = [
    "example.com","sentry.io","google.com","w3.org",
    "schema.org","wixpress.com","squarespace.com",
]

# Known ATS domains — always valid apply URLs
ATS_DOMAINS = [
    "greenhouse.io","lever.co","workday.com","bamboohr.com",
    "smartrecruiters.com","taleo.net","icims.com","jazzhr.com",
    "recruitee.com","workable.com","ashbyhq.com","breezy.hr",
    "jobvite.com","pinpointhq.com","teamtailor.com","personio.de",
    "comeet.com","rippling.com","gusto.com","ats.com",
    "myworkdayjobs.com","ultipro.com","successfactors.com",
    "oraclecloudhq.com","careers-page.com","applytojob.com",
]

FAKE_LOCAL_RE  = re.compile(
    r"^(name|user|email|mail|yourname|your[-_.]?email|sample|test|info|hello"
    r"|noreply|no[-_.]?reply|admin|webmaster|support|contact|example)$", re.I)
FAKE_DOMAIN_RE = re.compile(
    r"^(domain|example|yoursite|yourdomain|yourbrand|company|mycompany"
    r"|website|yourcompany|mysite|placeholder|site)\.[a-z]{2,}$", re.I)

MONTH_MAP = {
    "jan":0,"feb":1,"mar":2,"apr":3,"may":4,"jun":5,
    "jul":6,"aug":7,"sep":8,"oct":9,"nov":10,"dec":11,
}

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer",
}

# Industry inference keywords (for company-site harvesting)
INDUSTRY_KEYWORDS = [
    ("Information Technology", ["software","technology","it services","tech company","saas","cloud computing","it solutions"]),
    ("Finance & Banking", ["bank","financial services","fintech","insurance","investment","asset management"]),
    ("Construction & Real Estate", ["construction","real estate","property development","contracting","engineering & construction"]),
    ("Oil & Gas / Energy", ["oil and gas","petroleum","energy","renewable energy","power generation"]),
    ("Healthcare", ["healthcare","hospital","clinic","pharmaceutical","medical"]),
    ("Retail & E-commerce", ["retail","e-commerce","ecommerce","shopping","fmcg"]),
    ("Manufacturing", ["manufacturing","factory","industrial","production facility"]),
    ("Hospitality & Tourism", ["hospitality","hotel","tourism","travel agency","resort"]),
    ("Education", ["education","school","university","training institute","academy"]),
    ("Logistics & Transportation", ["logistics","shipping","freight","transportation","supply chain"]),
    ("Telecommunications", ["telecom","telecommunications","mobile network","internet service provider"]),
    ("Consulting", ["consulting","advisory","professional services"]),
    ("Media & Entertainment", ["media","entertainment","broadcasting","publishing","advertising agency"]),
    ("Agriculture", ["agriculture","farming","agribusiness"]),
    ("Government & Non-Profit", ["government","ministry","non-profit","ngo","public sector"]),
]

# =============================================================================
#  ▶▶ LINKEDIN URL FILTER
# =============================================================================

def is_linkedin_url(value: str) -> bool:
    """Return True if the value is a LinkedIn URL (any subdomain)."""
    if not value:
        return False
    return bool(re.search(r"linkedin\.com", value, re.I))

def blank_if_linkedin(value: str) -> str:
    """Return empty string if value contains a LinkedIn URL, else return as-is."""
    return "" if is_linkedin_url(value) else value

# =============================================================================
#  v7: BAD COMPANY NAME / LOGIN-WALL HELPERS
# =============================================================================

BAD_COMPANY_NAMES = {
    "sign in", "linkedin", "join now", "log in", "login",
    "join linkedin", "welcome back", "",
}

def is_bad_company_name(name: str) -> bool:
    """Detect LinkedIn login-wall placeholders masquerading as a company name."""
    if not name:
        return True
    n = name.strip().lower()
    if n in BAD_COMPANY_NAMES:
        return True
    if "linkedin" in n and len(n) < 30:
        return True
    return False

def extract_company_from_job_url(job_url: str) -> str:
    """
    Fallback: LinkedIn job URLs are typically of the form
        .../jobs/view/<job-title-slug>-at-<company-slug>-<numeric-id>/
    Extract <company-slug> and turn it into a readable name.
    """
    if not job_url:
        return ""
    m = re.search(r"-at-([a-z0-9\-]+)-\d+/?(?:\?.*)?$", job_url, re.I)
    if not m:
        return ""
    slug = m.group(1).replace("-", " ").strip()
    if not slug:
        return ""
    # Title-case but keep common acronyms upper if all-caps in slug originally
    return " ".join(w.upper() if len(w) <= 3 and w.isalpha() and w == w else w.title()
                     for w in slug.split()).title()

# =============================================================================
#  MOJIBAKE / TEXT HELPERS
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False, is_email=False) -> str:
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan","None","NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url or is_email:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    # v7: keep Arabic ranges (U+0600-U+06FF Arabic, U+0750-U+077F Arabic Supplement)
    # so Arabic-language job descriptions aren't reduced to stray punctuation.
    text = re.sub(
        r"[^\x20-\x7E\n\u00C0-\u017F\u0600-\u06FF\u0750-\u077F\u2013\u2014\u2018-\u201D\u2022]",
        "", text
    )
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

# =============================================================================
#  v7: LANGUAGE DETECTION HELPERS
# =============================================================================

_ARABIC_CHAR_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F]")
_LATIN_CHAR_RE  = re.compile(r"[A-Za-z]")

def detect_text_language(text: str) -> str:
    """
    Very lightweight heuristic language detector for the purpose of deciding
    whether to run the English-oriented paraphrase prompts.
    Returns "ar" if the text is predominantly Arabic, otherwise "en".
    """
    if not text:
        return "en"
    arabic_chars = len(_ARABIC_CHAR_RE.findall(text))
    latin_chars  = len(_LATIN_CHAR_RE.findall(text))
    total = arabic_chars + latin_chars
    if total == 0:
        return "en"
    if arabic_chars / total >= 0.4:
        return "ar"
    return "en"

# =============================================================================
#  NLP TOOLS (lazy init)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")


def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    → ✅ ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity     : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ ⚠️  No valid paraphrase found → Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean


def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in clean.split("\n") if p.strip()]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraphs) {'─'*25}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    → ✅ ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)


def paraphrase_company(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY PARAPHRASE {'─'*43}")
    orig_wc = len(clean.split())
    print(f" │ Original ({orig_wc} words):")
    _print_wrapped(clean, prefix=" │    ")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company description professionally. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw     = len(result.split()) if result else 0
    sim    = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ Paraphrased ({rw} words, sim={sim:.3f}):")
        _print_wrapped(result, prefix=" │    ")
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if rw < 10:    reasons.append(f"too short ({rw} words, min=10)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean


def paraphrase_tagline(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text[:300])
    if not clean:
        return text

    print(f"\n ┌─ TAGLINE PARAPHRASE {'─'*43}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company tagline as a crisp, professional phrase. "
        f"Output ONLY the rewritten tagline (5–12 words). No explanation.\n\n"
        f"Original: {clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=35, temperature=0.75)
    result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
    wc     = len(result.split()) if result else 0

    print(f" │ Paraphrased : \"{result}\"")
    print(f" │ Words: {wc}")

    if result and 3 <= wc <= 15:
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if wc < 3:     reasons.append(f"too short ({wc} words, min=3)")
        if wc > 15:    reasons.append(f"too long ({wc} words, max=15)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean

# =============================================================================
#  DUPLICATE TRACKER
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "WP ID",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)

def load_processed_ids() -> tuple:
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    df   = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)

def make_job_id(job_url: str, title: str = "", company: str = "", idx: int = 0) -> str:
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    seed = f"{title}{company}{idx}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped"})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": wp_id})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  HELPERS
# =============================================================================

def should_skip_crawl(url: str) -> bool:
    if not url: return True
    return any(d in url.lower() for d in SKIP_CRAWL_DOMAINS)

def is_bad_url(url: str) -> bool:
    if not url or not url.startswith("http"): return True
    return any(d in url.lower() for d in BAD_DOMAINS)

def is_ats_url(url: str) -> bool:
    if not url: return False
    return any(d in url.lower() for d in ATS_DOMAINS)

def is_career_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in [
        "career","jobs","apply","vacanci","recruit","opening",
        "hiring","work-with","join-us","join_us","opportunities",
        "current-opening","job-listing","positions",
    ])

def is_contact_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in [
        "contact","about","reach","get-in","enquir","support",
        "about-us","about_us",
    ])

def is_about_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in [
        "about","who-we-are","our-story","company","our-team",
        "overview","mission","vision",
    ])

def make_absolute(href: str, root_url: str) -> str:
    if not href: return ""
    href = href.strip()
    if href.startswith("http"): return href
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"): return root_url.rstrip("/") + href
    return ""

def decode_html_entities(s: str) -> str:
    if not s: return ""
    for old, new in [("&amp;","&"),("&lt;","<"),("&gt;",">"),("&quot;",'"'),
                     ("&#39;","'"),("\\u0026","&"),("\\u003D","="),
                     ("\\u003A",":"),("\\u002F","/")]:
        s = s.replace(old, new)
    return s

def canonicalise_job_url(url: str) -> str:
    if not url: return ""
    m = re.search(r"/jobs/view/(\d+)", url)
    if m: return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"
    return re.sub(r"[?#].*$", "", url)

def _strip_li_tracking(url: str) -> str:
    if not url: return ""
    url = decode_html_entities(url)
    m = re.search(r"[?&]url=([^&]+)", url)
    if m:
        try:
            decoded = unquote(m.group(1))
            if "%" in decoded: decoded = unquote(decoded)
            if decoded.startswith("http") and "linkedin.com" not in decoded:
                return decoded
        except Exception:
            pass
    return url

def _title_similarity(a: str, b: str) -> float:
    if not a or not b: return 0.0
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def _follow_redirect_chain(url: str, max_hops: int = 4) -> str:
    """Follow up to max_hops HTTP redirects for ATS redirect chains."""
    current = url
    for _ in range(max_hops):
        try:
            r = requests.head(current, headers=_next_headers(),
                              allow_redirects=False, timeout=8)
            if r.status_code in (301, 302, 303, 307, 308):
                loc = r.headers.get("Location", "")
                if loc:
                    if not loc.startswith("http"):
                        parsed = urlparse(current)
                        loc = f"{parsed.scheme}://{parsed.netloc}{loc}"
                    current = loc
                    continue
        except Exception:
            pass
        break
    return current

# =============================================================================
#  HTTP  (with retry + back-off)
# =============================================================================

def fetch_page(url: str, follow_redirects: bool = True, retries: int = 3) -> str | None:
    for attempt in range(retries):
        try:
            time.sleep(0.3 + attempt * 1.5)
            r = requests.get(url, headers=_next_headers(),
                             allow_redirects=follow_redirects, timeout=20)
            if r.status_code == 429:
                wait = 30 + attempt * 30
                log.warning(f"Rate-limited (429) — sleeping {wait}s")
                time.sleep(wait); continue
            if r.status_code in (403, 999):
                log.warning(f"Blocked ({r.status_code}): {url}"); return None
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}"); return None
            text = r.text
            if len(text) > FETCH_CHAR_LIMIT: text = text[:FETCH_CHAR_LIMIT]
            return text
        except Exception as e:
            log.warning(f"fetch attempt {attempt+1} failed ({url}): {e}")
            time.sleep(2 + attempt * 2)
    return None

# =============================================================================
#  DATE HELPERS
# =============================================================================

def normalise_date_text(text: str) -> str:
    if not text: return ""
    fr_map = {"heure":"hour","heures":"hours","jour":"day","jours":"days",
              "semaine":"week","semaines":"weeks","mois":"month","an":"year","ans":"years"}
    m = re.search(r"il\s+y\s+a\s+(\d+)\s+([a-zéè]+)", text, re.I)
    if m:
        unit = fr_map.get(m.group(2).lower())
        if unit: return f"{m.group(1)} {unit} ago"
    if re.match(r"^hier$", text.strip(), re.I): return "1 day ago"
    if re.search(r"aujourd|today", text, re.I): return "0 days ago"
    return text

def resolve_posted_date(raw: str) -> str:
    if not raw: return ""
    text = normalise_date_text(raw)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text.strip()): return text.strip()
    try: return datetime.fromisoformat(text).strftime("%Y-%m-%d")
    except Exception: pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day" in unit: base -= timedelta(days=n)
        elif "week"in unit: base -= timedelta(weeks=n)
        elif "month"in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
        elif "year"in unit: base = base.replace(year=base.year - n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I): return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")

def try_parse_date(s: str) -> datetime | None:
    if not s: return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%B %d, %Y", "%d %B %Y"):
        try: return datetime.strptime(s.strip(), fmt)
        except Exception: pass
    try: return datetime.fromisoformat(s)
    except Exception: pass
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m:
        mon = MONTH_MAP.get(m.group(2)[:3].lower())
        if mon is not None: return datetime(int(m.group(3)), mon+1, int(m.group(1)))
    return None

def parse_deadline(soup: BeautifulSoup) -> str:
    full_text = soup.get_text()
    patterns = [
        r"closes?\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"closes?\s+on\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"apply\s+by\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"apply\s+by\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"applications?\s+close[sd]?\s*(?:on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"closing\s+date[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"closing\s+date[:\s]+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    ]
    now = datetime.now()
    for p in patterns:
        m = re.search(p, full_text, re.I)
        if m:
            d = try_parse_date(m.group(1))
            if d and d > now: return d.strftime("%Y-%m-%d")
    return ""

def estimate_deadline_from_posted(posted_text: str) -> str:
    if not posted_text: return ""
    text = normalise_date_text(posted_text)
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour" in unit:  base -= timedelta(hours=n)
        elif "day"in unit:  base -= timedelta(days=n)
        elif "week"in unit: base -= timedelta(weeks=n)
        elif "month"in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
    mo = base.month + 3; yr = base.year + (mo - 1) // 12; mo = (mo - 1) % 12 + 1
    return base.replace(year=yr, month=mo).strftime("%Y-%m-%d")

# =============================================================================
#  TEXT CLEANERS
# =============================================================================

def clean_description(raw: str) -> str:
    if not raw: return ""
    text = raw.replace("\u00a0"," ").replace("\u200b","").replace("\r\n","\n").replace("\r","\n")
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-z])", r"\1 \2", text)
    text = re.sub(r"([A-Za-z])(\d)", r"\1 \2", text)
    text = re.sub(r"([.,:;!?])([A-Za-z0-9])", r"\1 \2", text)
    text = re.sub(r"\s*[•·▪◦]\s*", "\n• ", text)
    text = re.sub(r"\n\s*[-–—]\s+", "\n• ", text)
    paragraphs = re.split(r"\n{2,}", text)
    cleaned = []
    for para in paragraphs:
        lines = [l.strip() for l in para.split("\n") if l.strip()]
        out = []
        for line in lines:
            if (not re.search(r"[.!?:;,]$", line)
                    and not re.match(r"^[A-Z\s]{3,30}$", line)
                    and len(line) > 8
                    and not re.match(r"^[•\-–]", line)
                    and not re.match(r"^\w+:$", line)):
                line += "."
            out.append(line)
        cleaned.append("\n".join(out))
    return re.sub(r" {2,}", " ", "\n\n".join(p for p in cleaned if p.strip())).strip()

# v7: known TLDs used to correctly truncate email domains, fixing cases like
# "hr.ksa@mammoet.comtak" → "hr.ksa@mammoet.com"
COMMON_EMAIL_TLDS = [
    "com","net","org","gov","edu","info","biz","co",
    "sa","ae","uk","us","in","eg","jo","qa","kw","bh","om","ye","iq","lb","sy",
    "ma","tn","dz","ly","sd","mu","mr","pk","tr","de","fr","it","es","nl","ch",
    "io","app","tech","online","store","site","me","name","mobi",
]
_TLD_RE_PART = "|".join(sorted(set(COMMON_EMAIL_TLDS), key=len, reverse=True))

def clean_email(raw: str) -> str:
    if not raw: return ""
    em = raw
    em = re.sub(r"^mailto:", "", em, flags=re.I)
    em = re.sub(r"\?.*$", "", em)
    for pat, rep in [(r"\\u003[Ee]",""),(r"\\u003[Cc]",""),(r"\\u0040","@"),
                     (r"\\u002[Ee]","."),(r"\\u0026",""),(r"u003[Ee]",""),
                     (r"u003[Cc]",""),(r"u0040","@"),(r"&amp;",""),(r"&lt;",""),
                     (r"&gt;",""),(r"&#64;","@"),(r"&#46;","."),(r"&nbsp;",""),
                     (r"%40","@"),(r"%2[Ee]","."),(r"%20",""),(r"[>]+$",""),
                     (r"[<]+$","")]:
        em = re.sub(pat, rep, em, flags=re.I)
    em = em.strip().lower()
    if not em or "@" not in em or "." not in em: return ""
    if not re.match(r"^[a-zA-Z0-9]", em): return ""
    at = em.rfind("@")
    if at == -1: return ""
    local, domain = em[:at], em[at+1:]
    if ".mu" in domain:
        domain = re.sub(r"\.mu.*", ".mu", domain, flags=re.I)
    elif ".uk" in domain:
        domain = re.sub(r"\.uk.*", ".uk", domain, flags=re.I)
    else:
        # v7: truncate right after the first recognized TLD so trailing junk
        # (e.g. "comtak" from "mammoet.comtak") is correctly removed.
        m = re.search(rf"^([a-z0-9.\-]+\.(?:{_TLD_RE_PART}))(?:[a-z0-9].*)?$", domain, re.I)
        if m:
            domain = m.group(1)
        else:
            domain = re.sub(r"(\.[a-z]{2,6})[a-z0-9\-_/?#+]*$", r"\1", domain, flags=re.I)
    em = local + "@" + domain
    if not em or "@" not in em or "." not in em: return ""
    if not re.match(r"^[a-zA-Z0-9]", em): return ""
    return em

def clean_application_link(raw: str) -> str:
    if not raw: return ""
    raw = raw.strip()
    # ── Block LinkedIn URLs in application links ──────────────────────────────
    if is_linkedin_url(raw):
        log.info(f"Blanking LinkedIn application URL: {raw}")
        return ""
    if "@" in raw and not raw.startswith("http"): return clean_email(raw)
    if raw.startswith("http"):
        url = raw
        if ".mu" in url.lower():
            def mu_replace(m):
                tld, path = m.group(1), m.group(2) or ""
                if path and re.match(r"^/[a-z0-9\-/]+$", path, re.I): return tld + path
                return tld
            url = re.sub(r"(\.mu)(\/[^\s]*)?$", mu_replace, url, flags=re.I)
        url = re.sub(r"#.*$", "", url)
        url = re.sub(r"(subject|applysubject|refno|applyref|applyhere|clickhere|applynow)(\?.*)?$","",url,flags=re.I)
        url = re.sub(r"[.,;:!?)]+$", "", url)
        return url.strip()
    return raw

def clean_logo_url(raw: str) -> str:
    if not raw: return ""
    raw = decode_html_entities(raw).strip()
    if not raw.startswith("http"): return ""
    return re.sub(r"[\"')\s]+$", "", raw)

def is_placeholder_logo(url: str) -> bool:
    """Detect LinkedIn ghost/placeholder/login-wall logos that shouldn't be trusted."""
    if not url: return True
    l = url.lower()
    return any(k in l for k in [
        "ghost", "placeholder", "static.licdn.com/aero-v1/sc/h/"
        "9c8pery4andzj6ohjkjp54ms4", "default-company-logo",
        # v7: LinkedIn login-wall favicon
        "static.licdn.com/scds/common/u/images/logos/favicons",
        "favicon",
    ])

# =============================================================================
#  EMAIL HELPERS
# =============================================================================

def extract_email_from_text(text: str) -> str:
    if not text: return ""
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    for raw_em in emails:
        em = clean_email(raw_em)
        if not em or "@" not in em: continue
        parts = em.split("@")
        if len(parts) != 2: continue
        if any(em.find(d) != -1 for d in NOISE_EMAIL_DOMAINS): continue
        if FAKE_LOCAL_RE.match(parts[0]) or FAKE_DOMAIN_RE.match(parts[1]): continue
        return em
    return ""

def scan_page_for_email(soup: BeautifulSoup, raw_html: str = "") -> str:
    for tag in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        em = clean_email(tag.get("href", ""))
        if not em: continue
        if any(d in em for d in NOISE_EMAIL_DOMAINS): continue
        parts = em.split("@")
        if len(parts) == 2 and not FAKE_LOCAL_RE.match(parts[0]) and not FAKE_DOMAIN_RE.match(parts[1]):
            return em
    for sel in ["footer","#footer",".footer","#contact",".contact"]:
        for tag in soup.select(sel):
            found = extract_email_from_text(tag.get_text())
            if found: return found
    body = extract_email_from_text(soup.get_text())
    if body: return body
    if raw_html:
        obs = re.findall(
            r"[a-zA-Z0-9._%+\-]+\s*[\[\(]?\s*at\s*[\]\)]?\s*[a-zA-Z0-9.\-]+"
            r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*[a-zA-Z]{2,}", raw_html, re.I)
        if obs:
            norm = re.sub(r"\s*[\[\(]?\s*at\s*[\]\)]?\s*", "@", obs[0], flags=re.I)
            norm = re.sub(r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*", ".", norm, flags=re.I)
            norm = re.sub(r"\s+", "", norm).lower()
            if "@" in norm and not FAKE_LOCAL_RE.match(norm.split("@")[0]): return norm
        found = extract_email_from_text(raw_html)
        if found: return found
    return ""

# =============================================================================
#  DECODE / FOLLOW LINKEDIN APPLY URL
# =============================================================================

def decode_linkedin_apply_url(raw: str) -> str:
    if not raw: return ""
    raw = decode_html_entities(raw)
    if raw.startswith("http") and "linkedin.com" not in raw: return raw
    m = re.search(r"[?&]url=([^&]+)", raw)
    if m:
        try:
            d = unquote(m.group(1))
            if "%" in d: d = unquote(d)
            if d.startswith("http") and "linkedin.com" not in d: return d
        except Exception: pass
    b64m = re.search(r"[?&]offsiteApplyUrl=([^&]+)", raw)
    if b64m:
        try:
            d2 = base64.b64decode(unquote(b64m.group(1))).decode("utf-8")
            p = json.loads(d2)
            if p and "url" in p: return p["url"]
        except Exception: pass
    return ""

def follow_linkedin_apply_button(soup: BeautifulSoup, job_url: str) -> str:
    for tag in soup.find_all("a", href=True):
        ctrl = tag.get("data-tracking-control-name", "")
        if "offsite" in ctrl.lower() or "apply" in ctrl.lower():
            r = decode_linkedin_apply_url(tag["href"])
            if r and not is_bad_url(r): return r
    for tag in soup.find_all("a", href=True):
        href = tag["href"]; text = tag.get_text().lower()
        if ("apply" in text or "/apply" in href) and "linkedin.com" not in href:
            if href.startswith("http") and not is_bad_url(href): return href
    return ""

# =============================================================================
#  JSON-LD PARSER
# =============================================================================

def _parse_jsonld(html: str) -> dict:
    result = {}
    for raw in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        try:
            data = json.loads(raw.strip())
        except Exception:
            continue
        if isinstance(data, list):
            data = next((d for d in data if isinstance(d, dict)), {})
        if not isinstance(data, dict): continue
        schema_type = data.get("@type", "")
        if schema_type == "JobPosting":
            org = data.get("hiringOrganization", {}) or {}
            result.update({
                "job_title":       data.get("title", ""),
                "job_description": data.get("description", ""),
                "date_posted":     data.get("datePosted", ""),
                "valid_through":   data.get("validThrough", ""),
                "employment_type": data.get("employmentType", ""),
                "salary":          _extract_salary_jsonld(data.get("baseSalary", {})),
                "company_name":    org.get("name", ""),
                "company_logo":    clean_logo_url(
                    org.get("logo", "") if isinstance(org.get("logo"), str) else
                    org.get("logo", {}).get("url", "") if isinstance(org.get("logo"), dict) else ""),
                "company_url":     org.get("sameAs", "") or org.get("url", ""),
                "company_website": org.get("sameAs", "") or org.get("url", ""),
                "apply_url":       (data.get("url", "") or
                                    (data.get("applicationContact", {}) or {}).get("url", "")),
                "location":        _extract_location_jsonld(data.get("jobLocation", {})),
            })
            addr = data.get("jobLocation", {})
            if isinstance(addr, list): addr = addr[0] if addr else {}
            place = addr.get("address", {}) if isinstance(addr, dict) else {}
            if isinstance(place, dict):
                city    = place.get("addressLocality", "")
                country = place.get("addressCountry", "")
                if city or country:
                    result["location"] = ", ".join(filter(None, [city, country]))
        elif schema_type in ("Organization", "Corporation", "LocalBusiness"):
            result.update({
                "company_name":     data.get("name", ""),
                "company_logo":     clean_logo_url(
                    data.get("logo", "") if isinstance(data.get("logo"), str) else
                    data.get("logo", {}).get("url", "") if isinstance(data.get("logo"), dict) else ""),
                "company_url":      data.get("sameAs", "") or data.get("url", ""),
                "company_website":  data.get("sameAs", "") or data.get("url", ""),
                "company_industry": data.get("industry", ""),
                "company_founded":  str(data.get("foundingDate", "") or ""),
                "company_address":  _extract_address_jsonld(data.get("address", {})),
                "company_about":    data.get("description", ""),
            })
    return result

def _extract_salary_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        val = obj.get("value", {}); currency = obj.get("currency", "")
        if isinstance(val, dict):
            lo = val.get("minValue", ""); hi = val.get("maxValue", ""); unit = val.get("unitText", "")
            parts = [str(x) for x in [lo, hi] if x]
            return f"{currency} {' - '.join(parts)} {unit}".strip()
        return f"{currency} {val}".strip()
    return ""

def _extract_location_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, list): obj = obj[0] if obj else {}
    if not isinstance(obj, dict): return str(obj)
    addr = obj.get("address", {})
    if isinstance(addr, dict):
        return ", ".join(filter(None, [
            addr.get("addressLocality", ""),
            addr.get("addressRegion", ""),
            addr.get("addressCountry", ""),
        ]))
    return str(addr)

def _extract_address_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        return ", ".join(filter(None, [
            obj.get("streetAddress", ""),
            obj.get("addressLocality", ""),
            obj.get("addressRegion", ""),
            obj.get("postalCode", ""),
            obj.get("addressCountry", ""),
        ]))
    return ""

# =============================================================================
#  EXTRACT COMPANY DATA FROM JOB PAGE
# =============================================================================

def extract_company_from_job_page(html: str, soup: BeautifulSoup) -> dict:
    result = {}
    ld = _parse_jsonld(html)
    if ld: result.update({k: v for k, v in ld.items() if v})

    def _meta(name_or_prop: str) -> str:
        tag = (soup.find("meta", attrs={"property": name_or_prop}) or
               soup.find("meta", attrs={"name": name_or_prop}))
        return (tag.get("content", "") if tag else "").strip()

    og_image = _meta("og:image")
    if og_image and not result.get("company_logo"):
        result["company_logo"] = clean_logo_url(og_image)

    def _sel(*selectors) -> str:
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    if not result.get("company_name"):
        result["company_name"] = _sel(
            ".topcard__org-name-link",
            ".job-details-jobs-unified-top-card__company-name",
            ".topcard__flavor",
        )

    if not result.get("company_logo"):
        for img_sel in [".artdeco-entity-image","img.company-logo",
                        ".jobs-unified-top-card__company-logo img",".topcard__logo img"]:
            img = soup.select_one(img_sel)
            if img:
                src = img.get("src","") or img.get("data-delayed-url","") or img.get("data-ghost-url","")
                src = clean_logo_url(src)
                if src and "ghost" not in src.lower() and "placeholder" not in src.lower():
                    result["company_logo"] = src; break

    for chip in soup.select(
        ".job-details-jobs-unified-top-card__job-insight,"
        ".jobs-unified-top-card__job-insight,"
        ".jobs-details__salary-main-rail-card"
    ):
        text = chip.get_text(strip=True)
        if not result.get("company_industry") and re.search(
                r"technology|consulting|financial|healthcare|education|"
                r"manufacturing|retail|media|energy|real estate|"
                r"construction|automotive|telecom|pharmaceutical", text, re.I):
            result["company_industry"] = text

    for script in soup.find_all("script"):
        txt = script.string or ""
        if not txt.strip(): continue
        if not result.get("company_website"):
            for pat in [r'"companyPageUrl"\s*:\s*"([^"]+)"',
                        r'"companyUrl"\s*:\s*"([^"]+)"',
                        r'"websiteUrl"\s*:\s*"([^"]+)"']:
                m = re.search(pat, txt)
                if m:
                    url = _strip_li_tracking(decode_html_entities(m.group(1)))
                    if url.startswith("http") and "linkedin.com" not in url:
                        result["company_website"] = url; break
        if not result.get("company_logo"):
            for pat in [r'"logoUrl"\s*:\s*"([^"]+)"',
                        r'"companyLogo"\s*:\s*"([^"]+)"',
                        r'"logo"\s*:\s*"([^"]+)"']:
                m = re.search(pat, txt)
                if m:
                    logo = clean_logo_url(decode_html_entities(m.group(1)))
                    if logo: result["company_logo"] = logo; break
        if not result.get("company_about"):
            for pat in [r'"tagline"\s*:\s*"([^"]+)"', r'"description"\s*:\s*"([^"]{20,})"']:
                m = re.search(pat, txt)
                if m:
                    about = decode_html_entities(m.group(1)).strip()
                    if about and len(about) > 20: result["company_about"] = about; break

    return result

# =============================================================================
#  v6: COMPANY-WEBSITE LOGO HUNTER
# =============================================================================

def _extract_logo_from_soup(soup: BeautifulSoup, base_url: str) -> str:
    """
    Search a single page's soup for the best candidate company logo.
    Priority: JSON-LD Organization.logo > og:image (if logo-like) >
              <link rel=icon/apple-touch-icon> (high-res) >
              <img> with 'logo' in class/id/alt/src in header/nav/footer.
    """
    # 1. JSON-LD organization logo
    for raw in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            str(soup), re.S | re.I):
        try:
            data = json.loads(raw.strip())
        except Exception:
            continue
        if isinstance(data, list):
            data = next((d for d in data if isinstance(d, dict)), {})
        if not isinstance(data, dict): continue
        logo_field = data.get("logo")
        if isinstance(logo_field, str) and logo_field:
            cand = clean_logo_url(make_absolute(logo_field, base_url) if not logo_field.startswith("http") else logo_field)
            if cand: return cand
        if isinstance(logo_field, dict):
            cand = clean_logo_url(logo_field.get("url", ""))
            if cand: return cand
        org = data.get("publisher") or data.get("hiringOrganization")
        if isinstance(org, dict):
            ologo = org.get("logo")
            if isinstance(ologo, str) and ologo:
                cand = clean_logo_url(make_absolute(ologo, base_url) if not ologo.startswith("http") else ologo)
                if cand: return cand
            if isinstance(ologo, dict):
                cand = clean_logo_url(ologo.get("url", ""))
                if cand: return cand

    # 2. <img> elements explicitly tagged as logo, scoped to header/nav/footer first
    LOGO_IMG_RE = re.compile(r"logo", re.I)
    scopes = []
    for sel in ["header", "nav", ".header", "#header", ".navbar", ".site-header"]:
        scopes.extend(soup.select(sel))
    scopes.append(soup)  # whole page fallback

    for scope in scopes:
        for img in scope.find_all("img"):
            attrs_blob = " ".join(filter(None, [
                img.get("class", []) and " ".join(img.get("class", [])),
                img.get("id", ""), img.get("alt", ""), img.get("src", ""),
                img.get("data-src", ""),
            ]))
            if LOGO_IMG_RE.search(attrs_blob):
                src = (img.get("src") or img.get("data-src") or
                       img.get("data-lazy-src") or img.get("data-original") or "")
                src = clean_logo_url(make_absolute(src, base_url) if src and not src.startswith("http") else src)
                if src and not is_placeholder_logo(src) and not src.lower().endswith((".svg",)) or \
                   (src and src.lower().endswith(".svg")):
                    if src:
                        return src

    # 3. og:image / twitter:image meta tags (often the logo or a brand image)
    for prop in ["og:image", "og:image:secure_url", "twitter:image", "twitter:image:src"]:
        tag = soup.find("meta", attrs={"property": prop}) or soup.find("meta", attrs={"name": prop})
        if tag:
            content = tag.get("content", "")
            if content:
                cand = clean_logo_url(make_absolute(content, base_url) if not content.startswith("http") else content)
                if cand and not is_placeholder_logo(cand):
                    return cand

    # 4. apple-touch-icon (usually high-res square brand mark)
    for rel in ["apple-touch-icon", "apple-touch-icon-precomposed"]:
        link = soup.find("link", rel=lambda r: r and rel in [x.lower() for x in (r if isinstance(r, list) else [r])])
        if link:
            href = link.get("href", "")
            if href:
                cand = clean_logo_url(make_absolute(href, base_url) if not href.startswith("http") else href)
                if cand: return cand

    # 5. standard favicon as last resort
    for rel in ["icon", "shortcut icon"]:
        link = soup.find("link", rel=lambda r: r and rel in [x.lower() for x in (r if isinstance(r, list) else [r])])
        if link:
            href = link.get("href", "")
            if href:
                cand = clean_logo_url(make_absolute(href, base_url) if not href.startswith("http") else href)
                if cand: return cand

    return ""


def get_company_website_logo(website_url: str, about_url: str = "") -> str:
    """
    Visit the company's own website (homepage, then about page) to find a
    proper, non-placeholder logo. Used as a fallback / override whenever
    LinkedIn's logo is missing or looks like a ghost/placeholder image.
    """
    if not website_url or should_skip_crawl(website_url):
        return ""
    root = website_url.rstrip("/")
    try:
        time.sleep(0.4)
        r = requests.get(root, headers=_next_headers(), timeout=12, allow_redirects=True)
        if r.status_code == 200:
            html = r.text[:FETCH_CHAR_LIMIT]
            soup = BeautifulSoup(html, "html.parser")
            logo = _extract_logo_from_soup(soup, root)
            if logo:
                log.info(f"company-site logo found on homepage: {logo}")
                return logo
    except Exception as e:
        log.debug(f"get_company_website_logo home error ({root}): {e}")

    if about_url and about_url != website_url:
        try:
            time.sleep(0.4)
            r = requests.get(about_url, headers=_next_headers(), timeout=12, allow_redirects=True)
            if r.status_code == 200:
                html = r.text[:FETCH_CHAR_LIMIT]
                soup = BeautifulSoup(html, "html.parser")
                logo = _extract_logo_from_soup(soup, about_url)
                if logo:
                    log.info(f"company-site logo found on about page: {logo}")
                    return logo
        except Exception as e:
            log.debug(f"get_company_website_logo about error ({about_url}): {e}")

    return ""

# =============================================================================
#  v6: DEEP COMPANY WEBSITE CRAWLER — apply URL + email + enrichment
# =============================================================================

def crawl_company_website_deep(website_url: str, job_title: str) -> dict:
    """
    Multi-layer apply-URL hunter + company enrichment (v6):
      1. Fetch company home page
      2. Discover careers/jobs page via link text + URL patterns + HEAD probes
      3. Fuzzy-match job title in anchor text AND surrounding context
      4. Find the real Apply button; follow ATS redirect chains; detect mailto:
      5. ALWAYS additionally harvest: logo, about/description, address,
         founded year, industry, phone, social links — from home, about,
         contact, and careers pages — regardless of whether an apply URL
         was found, so missing company fields can be filled in.
      6. Fall back gracefully at each layer.

    Returns: {
        "apply_url": str, "email": str, "method": str,
        "logo": str, "about": str, "address": str, "founded": str,
        "industry": str, "phone": str, "social_links": str,
        "about_url": str, "careers_url": str,
    }
    """
    result = {
        "apply_url": "", "email": "", "method": "",
        "logo": "", "about": "", "address": "", "founded": "",
        "industry": "", "phone": "", "social_links": "",
        "about_url": "", "careers_url": "",
    }
    if not website_url or should_skip_crawl(website_url):
        return result

    deadline    = time.time() + 25
    root        = website_url.rstrip("/")
    parsed_root = urlparse(root)
    root_domain = parsed_root.netloc

    def _get(url: str):
        if time.time() > deadline: return None, None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=_next_headers(), timeout=12, allow_redirects=True)
            if r.status_code != 200: return None, None
            html = r.text[:FETCH_CHAR_LIMIT]
            return html, BeautifulSoup(html, "html.parser")
        except Exception as e:
            log.debug(f"deep crawl fetch error ({url}): {e}")
            return None, None

    def _same_domain_links(soup: BeautifulSoup, base: str) -> list:
        out = []
        for tag in soup.find_all("a", href=True):
            href = make_absolute(tag.get("href", ""), base)
            if not href or not href.startswith("http"): continue
            if root_domain not in urlparse(href).netloc: continue
            out.append((href, tag.get_text(strip=True).lower()))
        return out

    def _harvest_enrichment(soup: BeautifulSoup, html: str, page_url: str):
        """Pull logo/about/address/founded/industry/phone/social from a page."""
        full_text = soup.get_text(" ", strip=True)

        if not result["logo"]:
            logo = _extract_logo_from_soup(soup, page_url)
            if logo and not is_placeholder_logo(logo):
                result["logo"] = logo

        if not result["email"]:
            em = scan_page_for_email(soup, html)
            if em: result["email"] = em

        if not result["phone"]:
            ph_patterns = [
                r"\+\d[\d\s\-().]{7,18}\d",
                r"\b0\d[\d\s\-().]{6,14}\d",
                r"\(\d{2,4}\)\s*\d{3,4}[\s\-]\d{3,4}",
            ]
            for pat in ph_patterns:
                ph_m = re.search(pat, full_text)
                if ph_m:
                    candidate = ph_m.group(0).strip()
                    digits    = re.sub(r"\D", "", candidate)
                    if 7 <= len(digits) <= 15:
                        result["phone"] = candidate; break

        if not result["address"]:
            for el in soup.select(
                "[itemprop='address'],[itemprop='streetAddress'],"
                "[itemtype*='PostalAddress']"
            ):
                t = el.get_text(separator=", ", strip=True)
                if len(t) > 10:
                    result["address"] = t[:250]; break
            if not result["address"]:
                addr_m = re.search(
                    r"\d+[\w\s,.\-]{5,120}"
                    r"(?:street|st\b|avenue|ave\b|road|rd\b|boulevard|blvd|"
                    r"lane|ln\b|drive|dr\b|way\b|close|court|building|floor|"
                    r"suite|tower|plaza|district|zone|p\.?\s*o\.?\s*box|"
                    r"tripoli|benghazi|misrata|zawiya|bayda|tobruk|"
                    r"sirte|sebha|zliten|ajdabiya",
                    full_text, re.I)
                if addr_m:
                    result["address"] = addr_m.group(0).strip()[:250]

        if not result["founded"]:
            fy = re.search(
                r"(?:founded|established|incorporated|since|est\.?)\s*[:\-]?\s*((?:19|20)\d{2})",
                full_text, re.I)
            if fy: result["founded"] = fy.group(1)

        if not result["about"]:
            og = (soup.find("meta", property="og:description") or
                  soup.find("meta", attrs={"name": "description"}))
            if og:
                desc = og.get("content", "").strip()
                if len(desc) > 40:
                    result["about"] = desc[:600]
            if not result["about"]:
                for p in soup.find_all("p"):
                    t = p.get_text(strip=True)
                    if len(t) > 80:
                        result["about"] = t[:600]; break

        if not result["industry"]:
            meta_kw = ""
            kw_tag = soup.find("meta", attrs={"name": "keywords"})
            if kw_tag: meta_kw = kw_tag.get("content", "")
            combined_for_industry = (meta_kw + " " + full_text[:1500]).lower()
            for label, keywords in INDUSTRY_KEYWORDS:
                if any(k in combined_for_industry for k in keywords):
                    result["industry"] = label; break

        if not result["social_links"]:
            socials = []
            for a in soup.find_all("a", href=True):
                href = a["href"].lower()
                for platform in ["twitter.com","x.com","facebook.com","instagram.com",
                                  "linkedin.com","youtube.com","tiktok.com"]:
                    if platform in href and a["href"] not in socials:
                        socials.append(a["href"]); break
            if socials:
                result["social_links"] = ", ".join(socials[:5])

    # STEP 1 — Home page
    home_html, home_soup = _get(root)
    if not home_html:
        log.info(f"deep crawl: could not fetch home page {root}")
        return result

    home_links = _same_domain_links(home_soup, root)

    # Harvest enrichment data from homepage + footer immediately
    footer = (home_soup.find("footer") or
              home_soup.select_one("#footer,.footer,[class*='footer']"))
    if footer:
        _harvest_enrichment(BeautifulSoup(str(footer), "html.parser"), str(footer), root)
    _harvest_enrichment(home_soup, home_html, root)

    # STEP 2 — Find the careers page
    CAREER_TEXT_RE = re.compile(
        r"career|job|vacanc|opportunit|recruit|hiring|join\s*us|work\s*with\s*us|"
        r"open\s*positions?|current\s*openings?|we['']?re\s+hiring", re.I)

    careers_url = ""
    for href, txt in home_links:
        if CAREER_TEXT_RE.search(txt) or is_career_url(href):
            careers_url = href
            log.info(f"deep crawl: careers page from home link → {careers_url}")
            break

    if not careers_url:
        for path in ["/careers","/jobs","/job-openings","/vacancies",
                     "/work-with-us","/join-us","/opportunities",
                     "/careers/open-positions","/about/careers",
                     "/career","/recruitment","/join","/we-are-hiring",
                     "/careers.html","/jobs.html"]:
            if time.time() > deadline: break
            candidate = root + path
            try:
                r = requests.head(candidate, headers=_next_headers(),
                                  timeout=6, allow_redirects=True)
                if r.status_code == 200:
                    careers_url = candidate
                    log.info(f"deep crawl: careers page from HEAD probe → {careers_url}")
                    break
            except Exception:
                pass

    result["careers_url"] = careers_url

    # STEP 2b — Find the about page (for enrichment + logo fallback)
    about_url = ""
    for href, txt in home_links:
        if is_about_url(href) or "about" in txt:
            about_url = href; break
    if not about_url:
        for path in ["/about","/about-us","/our-story","/company","/who-we-are",
                      "/about.html","/about-us.html"]:
            if time.time() > deadline: break
            candidate = root + path
            try:
                r = requests.head(candidate, headers=_next_headers(),
                                  timeout=6, allow_redirects=True)
                if r.status_code == 200:
                    about_url = candidate; break
            except Exception:
                pass
    result["about_url"] = about_url

    if about_url and time.time() < deadline:
        about_html, about_soup = _get(about_url)
        if about_soup:
            _harvest_enrichment(about_soup, about_html, about_url)

    # STEP 2c — Contact page (enrichment: email/phone/address)
    contact_url = ""
    for href, txt in home_links:
        if is_contact_url(href) or "contact" in txt:
            contact_url = href; break
    if not contact_url:
        for path in ["/contact","/contact-us","/reach-us","/get-in-touch",
                      "/contact.html","/contact-us.html"]:
            if time.time() > deadline: break
            candidate = root + path
            try:
                r = requests.head(candidate, headers=_next_headers(),
                                  timeout=6, allow_redirects=True)
                if r.status_code == 200:
                    contact_url = candidate; break
            except Exception:
                pass

    if contact_url and time.time() < deadline:
        contact_html, contact_soup = _get(contact_url)
        if contact_soup:
            _harvest_enrichment(contact_soup, contact_html, contact_url)

    # If still no logo, try a dedicated logo hunt (homepage already tried,
    # but re-attempt with about page too)
    if not result["logo"]:
        logo = get_company_website_logo(root, about_url)
        if logo:
            result["logo"] = logo

    # If no careers page, still scan home/about/contact for an email and return
    if not careers_url:
        log.info(f"deep crawl: no careers page found at {root}")
        if result["email"]:
            result["method"] = "deep_site_email"
        return result

    # STEP 3 — Find the specific job listing on the careers page
    careers_html, careers_soup = _get(careers_url)
    if not careers_soup:
        return result

    # Harvest enrichment from careers page too (sometimes about info lives there)
    _harvest_enrichment(careers_soup, careers_html, careers_url)

    career_links          = _same_domain_links(careers_soup, careers_url)
    best_url, best_score  = "", 0.0

    for href, anchor_txt in career_links:
        if any(skip in href.lower() for skip in ["#","login","sign-in","signup",
                                                   "register","privacy","terms"]):
            continue
        score = _title_similarity(job_title, anchor_txt)
        for tag in careers_soup.find_all("a", href=True):
            abs_href = make_absolute(tag.get("href", ""), careers_url)
            if abs_href != href: continue
            parent = tag.parent
            if parent:
                context_txt = parent.get_text(separator=" ", strip=True).lower()
                ctx_score   = _title_similarity(job_title, context_txt)
                score       = max(score, ctx_score * 0.85)
            break
        if score > best_score:
            best_score, best_url = score, href

    email_from_careers = result["email"] or scan_page_for_email(careers_soup, careers_html)
    if email_from_careers and not result["email"]:
        result["email"] = email_from_careers

    if best_score < 0.30:
        log.info(f"deep crawl: no close title match (best={best_score:.2f}) on careers page")
        if email_from_careers:
            result["apply_url"] = ""
            result["method"]    = "deep_careers_email"
            return result
        result["apply_url"] = careers_url
        result["method"]    = "deep_careers_page"
        return result

    log.info(f"deep crawl: matched '{job_title}' → {best_url}  (score={best_score:.2f})")

    # STEP 4 — Job detail page: find the real Apply button
    job_html, job_soup = _get(best_url)
    if not job_soup:
        result["apply_url"] = best_url
        result["method"]    = "deep_job_page_url"
        return result

    # Harvest enrichment from the matched job page too (often has company blurb)
    _harvest_enrichment(job_soup, job_html, best_url)

    APPLY_TEXT_RE  = re.compile(
        r"apply\s*now|apply\s*online|apply\s*for\s+this|apply\s*here|"
        r"submit\s*(your\s*)?(application|cv|resume)|send\s*(cv|resume|application)|"
        r"click\s*to\s*apply|start\s*application", re.I)
    APPLY_CLASS_RE = re.compile(
        r"apply|btn[-_]?apply|cta[-_]?apply|job[-_]?apply|application[-_]?btn", re.I)

    apply_url = ""

    for tag in job_soup.find_all("a", href=True):
        tag_text = tag.get_text(strip=True)
        tag_cls  = " ".join(tag.get("class", []))
        href     = make_absolute(tag.get("href", ""), best_url)
        if not href: continue
        if href.lower().startswith("mailto:"):
            em = clean_email(href.replace("mailto:", "").replace("MAILTO:", ""))
            if em:
                log.info(f"deep crawl: mailto apply found → {em}")
                result["apply_url"] = ""
                result["email"]     = result["email"] or em
                result["method"]    = "deep_apply_email"
                return result
        if APPLY_TEXT_RE.search(tag_text) or APPLY_CLASS_RE.search(tag_cls):
            resolved = _follow_redirect_chain(href) if is_ats_url(href) else href
            if resolved and not is_bad_url(resolved):
                apply_url = resolved
                log.info(f"deep crawl: apply button found → {apply_url}")
                break

    if not apply_url:
        for btn in job_soup.find_all("button"):
            if APPLY_TEXT_RE.search(btn.get_text(strip=True)):
                data_href = btn.get("data-href", "") or btn.get("data-url", "")
                if data_href:
                    href = make_absolute(data_href, best_url)
                    if href and not is_bad_url(href):
                        apply_url = href; break

    if not apply_url:
        for form in job_soup.find_all("form"):
            action = make_absolute(form.get("action", ""), best_url)
            if action and not is_bad_url(action) and "apply" in action.lower():
                apply_url = action; break

    if not apply_url:
        for tag in job_soup.find_all("a", href=True):
            href = make_absolute(tag.get("href", ""), best_url)
            if href and is_ats_url(href) and not is_bad_url(href):
                apply_url = _follow_redirect_chain(href)
                log.info(f"deep crawl: ATS link found on job page → {apply_url}")
                break

    if not apply_url:
        em = scan_page_for_email(job_soup, job_html)
        if em and not result["email"]:
            result["email"] = em
        if em:
            result["apply_url"] = ""
            result["method"]    = "deep_job_email"
            return result

    if apply_url:
        result["apply_url"] = apply_url
        result["method"]    = "deep_apply_button"
        return result

    result["apply_url"] = best_url
    result["method"]    = "deep_job_page_url"
    return result

# =============================================================================
#  v6: ABOUT / CONTACT / FOOTER SCRAPER (standalone, used when deep crawl
#      is skipped, e.g. because an apply URL was already found upstream)
# =============================================================================

def scrape_about_contact_footer(website_url: str) -> dict:
    """
    Visit home, about, contact, and careers pages to harvest missing
    company details: address, phone, email, founded, description,
    industry, social links, and logo.
    """
    empty = {"address": "", "phone": "", "email": "",
             "founded": "", "description": "", "social_links": "",
             "logo": "", "industry": ""}
    if not website_url or should_skip_crawl(website_url): return dict(empty)

    root        = website_url.rstrip("/")
    parsed_root = urlparse(root)
    root_domain = parsed_root.netloc
    deadline    = time.time() + 18

    def _get(url: str):
        if time.time() > deadline: return None, None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=_next_headers(), timeout=12, allow_redirects=True)
            if r.status_code != 200: return None, None
            html = r.text[:FETCH_CHAR_LIMIT]
            return html, BeautifulSoup(html, "html.parser")
        except Exception as e:
            log.debug(f"about/contact fetch error ({url}): {e}")
            return None, None

    result = dict(empty)

    def _harvest(soup: BeautifulSoup, html: str, page_url: str):
        full_text = soup.get_text(" ", strip=True)

        if not result["logo"]:
            logo = _extract_logo_from_soup(soup, page_url)
            if logo and not is_placeholder_logo(logo):
                result["logo"] = logo

        if not result["email"]:
            em = scan_page_for_email(soup, html)
            if em: result["email"] = em

        if not result["phone"]:
            ph_patterns = [
                r"\+\d[\d\s\-().]{7,18}\d",
                r"\b0\d[\d\s\-().]{6,14}\d",
                r"\(\d{2,4}\)\s*\d{3,4}[\s\-]\d{3,4}",
            ]
            for pat in ph_patterns:
                ph_m = re.search(pat, full_text)
                if ph_m:
                    candidate = ph_m.group(0).strip()
                    digits    = re.sub(r"\D", "", candidate)
                    if 7 <= len(digits) <= 15:
                        result["phone"] = candidate; break

        if not result["address"]:
            for el in soup.select(
                "[itemprop='address'],[itemprop='streetAddress'],"
                "[itemtype*='PostalAddress']"
            ):
                t = el.get_text(separator=", ", strip=True)
                if len(t) > 10:
                    result["address"] = t[:250]; break
            if not result["address"]:
                addr_m = re.search(
                    r"\d+[\w\s,.\-]{5,120}"
                    r"(?:street|st\b|avenue|ave\b|road|rd\b|boulevard|blvd|"
                    r"lane|ln\b|drive|dr\b|way\b|close|court|building|floor|"
                    r"suite|tower|plaza|district|zone|p\.?\s*o\.?\s*box|"
                    r"tripoli|benghazi|misrata|zawiya|bayda|tobruk|"
                    r"sirte|sebha|zliten|ajdabiya",
                    full_text, re.I)
                if addr_m:
                    result["address"] = addr_m.group(0).strip()[:250]

        if not result["founded"]:
            fy = re.search(
                r"(?:founded|established|incorporated|since|est\.?)\s*[:\-]?\s*((?:19|20)\d{2})",
                full_text, re.I)
            if fy: result["founded"] = fy.group(1)

        if not result["description"]:
            og = (soup.find("meta", property="og:description") or
                  soup.find("meta", attrs={"name": "description"}))
            if og:
                desc = og.get("content", "").strip()
                if len(desc) > 40:
                    result["description"] = desc[:600]
            if not result["description"]:
                for p in soup.find_all("p"):
                    t = p.get_text(strip=True)
                    if len(t) > 80:
                        result["description"] = t[:600]; break

        if not result["industry"]:
            meta_kw = ""
            kw_tag = soup.find("meta", attrs={"name": "keywords"})
            if kw_tag: meta_kw = kw_tag.get("content", "")
            combined_for_industry = (meta_kw + " " + full_text[:1500]).lower()
            for label, keywords in INDUSTRY_KEYWORDS:
                if any(k in combined_for_industry for k in keywords):
                    result["industry"] = label; break

        if not result["social_links"]:
            socials = []
            for a in soup.find_all("a", href=True):
                href = a["href"].lower()
                for platform in ["twitter.com","x.com","facebook.com","instagram.com",
                                  "linkedin.com","youtube.com","tiktok.com"]:
                    if platform in href and a["href"] not in socials:
                        socials.append(a["href"]); break
            if socials:
                result["social_links"] = ", ".join(socials[:5])

    home_html, home_soup = _get(root)
    if home_soup:
        footer = (home_soup.find("footer") or
                  home_soup.select_one("#footer,.footer,[class*='footer']"))
        if footer:
            _harvest(BeautifulSoup(str(footer), "html.parser"), str(footer), root)
        _harvest(home_soup, home_html, root)

        about_url = contact_url = careers_url = ""
        for tag in home_soup.find_all("a", href=True):
            href = make_absolute(tag.get("href", ""), root)
            if not href or root_domain not in urlparse(href).netloc: continue
            txt = tag.get_text(strip=True).lower()
            if not about_url and (is_about_url(href) or "about" in txt):
                about_url = href
            if not contact_url and (is_contact_url(href) or "contact" in txt):
                contact_url = href
            if not careers_url and (is_career_url(href) or "career" in txt or "job" in txt):
                careers_url = href
            if about_url and contact_url and careers_url: break

        if not about_url and time.time() < deadline:
            for path in ["/about","/about-us","/our-story","/company","/who-we-are"]:
                try:
                    r = requests.head(root + path, headers=_next_headers(),
                                      timeout=5, allow_redirects=True)
                    if r.status_code == 200:
                        about_url = root + path; break
                except Exception:
                    pass

        if not contact_url and time.time() < deadline:
            for path in ["/contact","/contact-us","/reach-us","/get-in-touch"]:
                try:
                    r = requests.head(root + path, headers=_next_headers(),
                                      timeout=5, allow_redirects=True)
                    if r.status_code == 200:
                        contact_url = root + path; break
                except Exception:
                    pass

        if not careers_url and time.time() < deadline:
            for path in ["/careers","/jobs","/vacancies","/work-with-us","/join-us"]:
                try:
                    r = requests.head(root + path, headers=_next_headers(),
                                      timeout=5, allow_redirects=True)
                    if r.status_code == 200:
                        careers_url = root + path; break
                except Exception:
                    pass

        if about_url and time.time() < deadline:
            about_html, about_soup = _get(about_url)
            if about_soup: _harvest(about_soup, about_html, about_url)

        if contact_url and time.time() < deadline:
            contact_html, contact_soup = _get(contact_url)
            if contact_soup: _harvest(contact_soup, contact_html, contact_url)

        if careers_url and time.time() < deadline:
            careers_html, careers_soup = _get(careers_url)
            if careers_soup: _harvest(careers_soup, careers_html, careers_url)

    if not result["logo"]:
        logo = get_company_website_logo(root)
        if logo: result["logo"] = logo

    return result

# =============================================================================
#  COMPANY PAGE SCRAPER  (LinkedIn)
# =============================================================================

def scrape_company_details(company_url: str) -> dict:
    empty = {
        "name":"","industry":"","size":"","headquarters":"","type":"",
        "founded":"","specialties":"","website":"","logo":"","about":"",
        "company_url":"",
    }
    if not company_url: return dict(empty)
    log.info(f"Scraping company page: {company_url}")
    base_url = re.sub(r"\?.*$", "", company_url.rstrip("/"))
    html     = None
    guest_url = base_url.replace(
        "https://www.linkedin.com/company/",
        "https://www.linkedin.com/company-guest/",
    )
    if guest_url != base_url: html = fetch_page(guest_url)
    if not html:
        for attempt in range(3):
            try:
                time.sleep(1.5 + attempt * 2)
                r = requests.get(base_url, headers=_next_headers(),
                                 allow_redirects=True, timeout=20)
                if r.status_code == 429:
                    log.warning("Company page rate-limited — sleeping 60s"); time.sleep(60); continue
                if r.status_code == 200:
                    text = r.text
                    if len(text) > FETCH_CHAR_LIMIT: text = text[:FETCH_CHAR_LIMIT]
                    html = text; break
                log.warning(f"Company page HTTP {r.status_code}: {base_url}"); break
            except Exception as e:
                log.warning(f"Company page fetch error (attempt {attempt+1}): {e}")
                time.sleep(2 + attempt * 2)
    if not html:
        result = dict(empty)
        result["company_url"] = base_url
        return result

    soup = BeautifulSoup(html, "html.parser")
    ld   = _parse_jsonld(html)

    def _sel(*selectors) -> str:
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    def _get_detail(label: str) -> str:
        lower = label.lower()
        for div in soup.select("section.core-section-container dl > div"):
            dt = div.find("dt")
            if dt and lower in dt.get_text().strip().lower():
                dd = div.find("dd")
                if dd: return dd.get_text(strip=True)
        for row in soup.select(".org-page-details__definition-list dt, .about-us__basicInfo dt"):
            if lower in row.get_text().strip().lower():
                dd = row.find_next_sibling("dd")
                if dd: return dd.get_text(strip=True)
        return ""

    og_img_tag = (soup.find("meta", property="og:image") or
                  soup.find("meta", attrs={"name": "og:image"}))
    raw_logo = (og_img_tag.get("content", "") if og_img_tag else "") or ld.get("company_logo", "")
    if not raw_logo:
        for img in soup.select("img.org-top-card-primary-content__logo, img.artdeco-entity-image"):
            src = img.get("src", "") or img.get("data-delayed-url", "")
            if src and "ghost" not in src.lower():
                raw_logo = src; break
    logo = clean_logo_url(raw_logo)

    ws_tag  = soup.select_one("a[data-tracking-control-name='about_website']")
    raw_ws  = (ws_tag.get("href", "") if ws_tag else "") or _get_detail("Website") or ld.get("company_website", "")
    website = decode_linkedin_apply_url(raw_ws) or raw_ws
    # Blank if still LinkedIn
    website = blank_if_linkedin(website)

    name = (ld.get("company_name", "") or _sel("h1.org-top-card-summary__title", "h1", "title") or "")
    if " | LinkedIn" in name: name = name.split(" | ")[0].strip()

    about = (ld.get("company_about", "") or
             _sel("section.about-us p", ".core-section-container__content p",
                  ".org-about-us-organization-description__text",
                  ".org-about-module__description") or "")

    return {
        "name":         name,
        "industry":     _get_detail("Industry") or ld.get("company_industry", ""),
        "size":         _get_detail("Company size"),
        "headquarters": _get_detail("Headquarters") or ld.get("company_address", ""),
        "type":         _get_detail("Type"),
        "founded":      _get_detail("Founded") or ld.get("company_founded", ""),
        "specialties":  _get_detail("Specialties"),
        "website":      website,
        "logo":         logo,            # raw LinkedIn logo — may be placeholder/empty
        "about":        about,
        "company_url":  base_url,        # LinkedIn company page URL — kept!
    }

# =============================================================================
#  WORDPRESS LOGO UPLOAD
# =============================================================================

def upload_logo_to_wordpress(logo_url: str, company_name: str) -> str:
    if not logo_url or not logo_url.startswith("http") or not WP_USER: return ""
    try:
        r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"],
                                             "Referer": "https://www.linkedin.com/"}, timeout=15)
        if r.status_code != 200: return ""
        ct  = r.headers.get("Content-Type", "image/jpeg")
        ext = "png" if "png" in ct else "jpg"
        fn  = re.sub(r"[^a-z0-9]", "-", company_name.lower()) + "-logo." + ext
        creds = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
        wr    = requests.post(WP_URL + "media",
                              headers={"Authorization": "Basic " + creds,
                                       "Content-Disposition": f"attachment; filename={fn}",
                                       "Content-Type": ct},
                              data=r.content, timeout=20)
        if wr.status_code in (200, 201): return wr.json().get("source_url", "")
        log.warning(f"WP upload failed ({wr.status_code})")
    except Exception as e:
        log.warning(f"uploadLogoToWordPress: {e}")
    return ""

# =============================================================================
#  v3 COMPANY WEBSITE CRAWLER  (fallback)
# =============================================================================

def crawl_company_website(website_url: str, job_title: str) -> dict:
    log.info(f"Crawling company site (v3 fallback): {website_url}")
    if should_skip_crawl(website_url):
        return {"url": website_url, "email": "", "method": "fallback_website"}
    deadline = time.time() + 12
    root_url = website_url.rstrip("/")

    def get(url):
        if time.time() > deadline: return None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            if r.status_code != 200: return None
            t = r.text
            return t[:FETCH_CHAR_LIMIT] if len(t) > FETCH_CHAR_LIMIT else t
        except Exception: return None

    home_html = get(root_url)
    if not home_html: return {"url": "", "email": "", "method": ""}
    soup_h     = BeautifulSoup(home_html, "html.parser")
    home_email = scan_page_for_email(soup_h, home_html)
    if home_email: return {"url": "", "email": home_email, "method": "s7_homepage_email"}

    careers_url = contact_url = ""
    for tag in soup_h.find_all("a", href=True):
        href      = make_absolute(tag.get("href", ""), root_url)
        link_text = tag.get_text().lower()
        if not href or is_bad_url(href) or href == root_url: continue
        if root_url not in href: continue
        if not careers_url and is_career_url(href): careers_url = href
        if not contact_url and (is_contact_url(href) or "contact" in link_text): contact_url = href
        if careers_url and contact_url: break

    if careers_url and time.time() < deadline:
        ch = get(careers_url)
        if ch:
            em = scan_page_for_email(BeautifulSoup(ch, "html.parser"), ch)
            if em: return {"url": "", "email": em, "method": "s7_careers_email"}
    if contact_url and time.time() < deadline:
        cth = get(contact_url)
        if cth:
            em = scan_page_for_email(BeautifulSoup(cth, "html.parser"), cth)
            if em: return {"url": "", "email": em, "method": "s7_contact_email"}
    if careers_url: return {"url": careers_url, "email": "", "method": "s7_careers_page"}
    return {"url": root_url, "email": "", "method": "fallback_website"}

# =============================================================================
#  APPLICATION DETAILS EXTRACTOR  (v6 priority chain)
# =============================================================================

def extract_application_details(
    job_url: str,
    soup: BeautifulSoup,
    company_website: str,
    ld: dict,
    job_title: str = "",
    site_info: dict | None = None,
) -> dict:
    """
    Priority chain:
      0  JSON-LD apply URL
      1  LinkedIn "Apply" button on the job page
      2  Script tags on the job page
      3  v6 deep crawl: home → careers → job listing → Apply button
         (also returns enrichment fields used by caller)
      4  Links / URLs in the job description
      5  Email from site_info (about/contact/footer scrape)
      6  v3 company website fallback crawl

    Returns dict with at minimum: url, email, method.
    If step 3 (deep crawl) ran, the result also includes "_deep" with the
    full enrichment dict (logo/about/address/founded/industry/etc.) so the
    caller can use it even if no apply URL/email is ultimately chosen.
    """
    desc_text = ""
    for sel in [".show-more-less-html__markup", ".description__text"]:
        el = soup.select_one(sel)
        if el: desc_text = el.get_text(); break

    # ── 0. JSON-LD ───────────────────────────────────────────────────────────
    if ld.get("apply_url") and not is_bad_url(ld["apply_url"]):
        url = blank_if_linkedin(ld["apply_url"])
        if url:
            log.info(f"apply found via JSON-LD: {url}")
            return {"url": url, "email": "", "method": "s0_jsonld"}

    # ── 1. LinkedIn apply button ──────────────────────────────────────────────
    apply_btn = follow_linkedin_apply_button(soup, job_url)
    if apply_btn:
        apply_btn = blank_if_linkedin(apply_btn)
        if apply_btn:
            log.info(f"apply found via LinkedIn button: {apply_btn}")
            return {"url": apply_btn, "email": "", "method": "s0_apply_button"}

    # ── 2. Script tags on job page ────────────────────────────────────────────
    for script in soup.find_all("script"):
        txt = script.string or ""
        for pat in [r'"applyStartUrl"\s*:\s*"([^"]+)"',
                    r'"applicationUrl"\s*:\s*"([^"]+)"']:
            m = re.search(pat, txt)
            if m:
                cand = decode_html_entities(m.group(1)).replace("\\", "")
                cand = blank_if_linkedin(cand)
                if cand and cand.startswith("http") and not is_bad_url(cand):
                    log.info(f"apply found via script tag: {cand}")
                    return {"url": cand, "email": "", "method": "s1b_script_tag"}

    # ── 3. Deep crawl company website (apply + enrichment) ───────────────────
    deep_info = {}
    if company_website and not should_skip_crawl(company_website):
        log.info(f"v6 deep crawl starting: {company_website} for '{job_title}'")
        deep_info = crawl_company_website_deep(company_website, job_title)
        if deep_info.get("email"):
            log.info(f"apply found via deep crawl (email): {deep_info['email']}")
            return {"url": "", "email": deep_info["email"],
                    "method": deep_info.get("method") or "deep_email", "_deep": deep_info}
        if deep_info.get("apply_url") and not is_bad_url(deep_info["apply_url"]):
            apply_url = blank_if_linkedin(deep_info["apply_url"])
            if apply_url:
                log.info(f"apply found via deep crawl (url): {apply_url}")
                return {"url": apply_url, "email": "",
                        "method": deep_info.get("method") or "deep_url", "_deep": deep_info}

    # ── 4. Links / URLs in job description ────────────────────────────────────
    desc_el = (soup.select_one(".show-more-less-html__markup") or
               soup.select_one(".description__text"))
    if desc_el:
        for a in desc_el.find_all("a", href=True):
            h = blank_if_linkedin(a.get("href", ""))
            if h and not is_bad_url(h):
                log.info(f"apply found via description link: {h}")
                return {"url": h, "email": "", "method": "s3_desc_link", "_deep": deep_info}

    for u in re.findall(r"https?://[^\s\"'<>)(,\]]+", desc_text):
        u = re.sub(r"[.,;:!?)]+$", "", u)
        u = blank_if_linkedin(u)
        if u and not is_bad_url(u):
            log.info(f"apply found via description URL: {u}")
            return {"url": u, "email": "", "method": "s4_desc_url", "_deep": deep_info}

    em = extract_email_from_text(desc_text)
    if em:
        log.info(f"apply found via description email: {em}")
        return {"url": "", "email": em, "method": "s5_desc_email", "_deep": deep_info}

    # ── 5. Email from site_info ────────────────────────────────────────────────
    if site_info and site_info.get("email"):
        log.info(f"apply found via site_info email: {site_info['email']}")
        return {"url": "", "email": site_info["email"], "method": "site_info_email", "_deep": deep_info}

    # ── 6. v3 fallback crawl ──────────────────────────────────────────────────
    resolved = decode_linkedin_apply_url(company_website) or company_website
    resolved = blank_if_linkedin(resolved)
    if resolved and not is_bad_url(resolved):
        if should_skip_crawl(resolved):
            return {"url": resolved, "email": "", "method": "fallback_website", "_deep": deep_info}
        res = crawl_company_website(resolved, job_title)
        if res.get("email") or res.get("url"):
            if res.get("url"): res["url"] = blank_if_linkedin(res["url"])
            res["_deep"] = deep_info
            return res
        return {"url": resolved, "email": "", "method": "fallback_website", "_deep": deep_info}

    return {"url": "", "email": "", "method": "not_found", "_deep": deep_info}

# =============================================================================
#  JOB CRITERIA HELPERS
# =============================================================================

def get_job_criteria(soup: BeautifulSoup, label: str) -> str:
    lower = label.lower()
    for li in soup.select(".description__job-criteria-list > li"):
        h3 = li.find("h3")
        if h3 and lower in h3.get_text().strip().lower():
            spans = li.select(".description__job-criteria-text, span")
            if spans: return spans[-1].get_text(strip=True)
    for chip in soup.select(
        ".job-details-jobs-unified-top-card__job-insight,"
        ".jobs-unified-top-card__job-insight"):
        text = chip.get_text(strip=True).lower()
        if "employment" in lower or "type" in lower:
            if re.search(r"full[\-\s]?time|part[\-\s]?time|contract|temporary|internship|freelance", text, re.I):
                return chip.get_text(strip=True)
        elif "seniority" in lower:
            if re.search(r"entry|associate|mid[\-\s]?senior|senior|director|executive|intern", text, re.I):
                return chip.get_text(strip=True)
    meta_map = {"employment type": soup.find("meta", {"name": "employmentType"}),
                "seniority level": soup.find("meta", {"name": "seniorityLevel"}),
                "industries":      soup.find("meta", {"name": "industry"})}
    tag = meta_map.get(lower)
    if tag: return tag.get("content", "")
    return ""

def get_workplace_type(soup: BeautifulSoup) -> str:
    for s in [".topcard__workplace-type",
              ".job-details-jobs-unified-top-card__workplace-type",
              ".jobs-unified-top-card__workplace-type"]:
        el = soup.select_one(s)
        if el: return el.get_text(strip=True)
    for chip in soup.select(
        ".job-details-jobs-unified-top-card__job-insight,"
        ".jobs-unified-top-card__job-insight"):
        t = chip.get_text(strip=True)
        if re.match(r"^(remote|on[\-\s]?site|hybrid)$", t, re.I): return t
    return ""

# =============================================================================
#  JOB FIELD INFERENCE
# =============================================================================

FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer","developer","devops","frontend","backend","full stack","fullstack",
      "sysadmin","cloud","cybersecurity","data engineer","machine learning","artificial intelligence",
      "ai/ml","it support","network engineer","database","kubernetes","docker","aws","azure",
      "react","node.js","python developer","java developer"],
     ["programming","coding","api","agile","scrum","git","linux","server","infrastructure","software"]),
    ("Finance & Accounting",
     ["accountant","auditor","finance manager","financial analyst","cfo","treasurer","tax",
      "bookkeeper","payroll","budget analyst","credit analyst","investment","portfolio manager",
      "risk analyst","forex","actuary","acca","cfa","cpa"],
     ["financial","accounting","balance sheet","p&l","reconciliation","ifrs","gaap","ledger","invoicing"]),
    ("Sales & Business Development",
     ["sales executive","sales manager","business development","account manager",
      "sales representative","bd manager","regional sales","key account","sales director",
      "commercial manager","sales officer"],
     ["revenue","pipeline","crm","leads","prospects","quota","target","upsell","cross-sell","b2b","b2c"]),
    ("Marketing & Communications",
     ["marketing manager","digital marketing","seo","sem","content marketer","social media manager",
      "brand manager","marketing executive","communications manager","pr manager","copywriter",
      "growth hacker","email marketing","campaign manager"],
     ["marketing","branding","advertising","social media","content","campaign","analytics",
      "google ads","facebook ads","influencer"]),
    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition","hr business partner",
      "hrbp","hr officer","compensation","benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager"],
     ["recruitment","onboarding","performance management","employee relations","hr","workforce"]),
    ("Engineering",
     ["mechanical engineer","civil engineer","electrical engineer","structural engineer",
      "process engineer","project engineer","maintenance engineer","production engineer",
      "quality engineer","safety engineer","site engineer","design engineer"],
     ["engineering","cad","autocad","solidworks","manufacturing","plant","machinery","commissioning"]),
    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical","healthcare manager",
      "occupational therapist","dentist","midwife"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment"]),
    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor","school principal",
      "academic","curriculum","e-learning","instructional designer","teaching assistant"],
     ["school","university","college","classroom","students","pedagogy","curriculum","education"]),
    ("Hospitality & Tourism",
     ["hotel manager","front desk","housekeeping","chef","sous chef","food and beverage",
      "f&b manager","restaurant manager","bartender","waiter","concierge","tour guide",
      "travel agent","events coordinator","catering"],
     ["hospitality","hotel","resort","tourism","guest","accommodation","restaurant","kitchen"]),
    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager","fleet manager",
      "procurement manager","purchasing manager","import export","freight","shipping coordinator",
      "inventory manager","demand planner"],
     ["logistics","supply chain","warehouse","inventory","freight","procurement","sourcing"]),
    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr"]),
    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer","operations manager",
      "pa","personal assistant","receptionist","data entry","office administrator",
      "company secretary","business analyst"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical"]),
    ("Customer Service",
     ["customer service","call centre","customer success","customer support","help desk",
      "service advisor","client relations","customer experience","contact centre"],
     ["customer","support","helpdesk","tickets","escalation","satisfaction","service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor","site supervisor","project manager construction","architect",
      "draughtsman","property manager","estate agent","real estate","building inspector",
      "land surveyor","construction manager"],
     ["construction","building","property","real estate","site","contractor","tender"]),
    ("Manufacturing & Production",
     ["production manager","quality control","quality assurance","qa","qc","factory manager",
      "plant manager","production supervisor","assembly","cnc operator","technician"],
     ["production","manufacturing","factory","assembly","quality","lean","six sigma"]),
    ("Design & Creative",
     ["graphic designer","ui/ux","product designer","art director","creative director",
      "animator","illustrator","photographer","videographer","motion designer","web designer"],
     ["design","creative","adobe","figma","photoshop","illustrator","indesign","sketch","branding"]),
    ("Research & Science",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist","statistician"],
     ["research","analysis","data","laboratory","science","experiment","findings","methodology"]),
    ("Security",
     ["security officer","security guard","security manager","cctv","loss prevention",
      "risk manager","health and safety","hse officer","osh","fire safety"],
     ["security","safety","risk","surveillance","patrol","access control","emergency"]),
    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor","content creator",
      "media manager","radio","television","producer","scriptwriter"],
     ["media","journalism","broadcast","news","editorial","publishing","press"]),
    ("Non-Profit & Social Work",
     ["social worker","ngo","charity","programme coordinator","community development",
      "welfare officer","case manager","development officer","fundraiser","volunteer coordinator"],
     ["social","ngo","community","welfare","beneficiary","donor","impact","charity"]),
]

def infer_job_field(title: str, description: str) -> str:
    if not title and not description: return ""
    combined = ((title or "") + " " + (description or "")).lower()
    best_field, best_score = "", 0
    for label, high_keys, supporting in FIELD_KEYWORD_MAP:
        score  = sum(3 for k in high_keys if k in combined)
        score += sum(1 for k in supporting if k in combined)
        if score > best_score: best_score, best_field = score, label
    if best_score >= 3: return best_field
    return ""

# =============================================================================
#  QUALIFICATION / EXPERIENCE EXTRACTORS
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",          ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",          ["master","msc","m.sc","ma ","m.a ","mba","m.b.a","meng","m.eng","mphil",
                                   "postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",        ["bachelor","bsc","b.sc","ba ","b.a ","beng","b.eng","bcom","b.com","bba",
                                   "llb","degree in","undergraduate degree","honours degree","hons"]),
    ("Higher National Diploma",  ["hnd","hnc","higher national diploma","higher national certificate",
                                   "higher diploma","advanced diploma"]),
    ("Diploma",                  ["diploma","dip ","dip.","associate degree","foundation degree"]),
    ("Professional Certification",["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
                                    "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
                                    "certified public","certified financial","certified project",
                                    "professional certification","professional certificate"]),
    ("A-Levels / HSC",           ["a-level","a level","hsc","higher school certificate","ib diploma",
                                   "international baccalaureate","gce advanced"]),
    ("O-Levels / School Certificate",["o-level","o level","igcse","gcse","school certificate",
                                       "sc ","cpe","certificate of primary"]),
    ("No Formal Qualification Required",["no qualification","no degree","no formal","school leaver",
                                          "entry level","no experience required","training provided","will train"]),
]

def extract_qualification(text: str) -> str:
    if not text: return ""
    if re.search(r"nursery|primary years|ib pyp|aged between|boys and girls", text, re.I): return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(k in lower for k in keywords): return label
    return ""

NO_EXP_KW = ["no experience","no prior experience","fresh graduate","freshers",
              "entry level","entry-level","0 years","zero experience",
              "training provided","will train","no experience required"]
LESS1_KW  = ["less than 1 year","under 1 year","6 months","less than a year",
              "some experience","minimal experience"]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience(text: str) -> str:
    if not text: return ""
    if re.search(r"aged?\s+between|boys\s+and\s+girls|nursery|primary\s+years|IB\s+PYP", text, re.I): return ""
    lower = text.lower()
    if any(k in lower for k in NO_EXP_KW): return "No Experience Required"
    if any(k in lower for k in LESS1_KW):  return "Less than 1 Year"
    patterns = [
        r"(\d+)\s*[-–to]+\s*(\d+)\s*\+?\s*years?",
        r"(\d+)\s*\+\s*years?\s*(?:of\s+)?(?:experience)?",
        r"(?:minimum|at\s+least|over|more\s+than)\s+(\d+)\s*\+?\s*years?",
        r"(\d+)\s*years?\s*(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            raw = int(m.group(1))
            if raw > 20: continue
            return years_to_band(raw)
    return ""

# =============================================================================
#  JOB DETAIL SCRAPER  (v7 + paraphrase)
# =============================================================================

def scrape_job_details(job_url: str, processed_ids: set, processed_urls: set) -> dict | None:
    """
    Scrape a single LinkedIn job (v7 layers) then paraphrase key fields.
    Returns the job dict or None.
    """
    job_id = make_job_id(job_url)
    if job_id in processed_ids or job_url in processed_urls:
        print(C_DIM(f"  ⧳ Already processed — skipped ({job_url})"))
        return None

    log.info(f"Scraping job: {job_url}")
    try:
        resp = requests.get(job_url, headers=_next_headers(), timeout=20)
        if resp.status_code != 200: return None
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")
    except Exception as e:
        log.warning(f"Job fetch failed: {e}"); return None

    def sel_text(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    title = sel_text(".top-card-layout__title", "h1.topcard__title",
                     ".job-details-jobs-unified-top-card__job-title", "h1")
    if not title: return None

    company_name_fallback = sel_text(
        ".topcard__org-name-link",
        ".job-details-jobs-unified-top-card__company-name",
        ".topcard__flavor",
    )
    company_url_el = (soup.select_one(".topcard__org-name-link") or
                      soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    # companyUrl is the LinkedIn company page URL — kept in its own column
    company_url_raw = company_url_el.get("href", "") if company_url_el else ""

    location       = sel_text(".topcard__flavor--bullet",
                               ".job-details-jobs-unified-top-card__bullet")
    workplace_type = get_workplace_type(soup)

    time_el    = soup.find("time")
    raw_posted = (time_el.get("datetime", "") if time_el else "") or \
                 sel_text(".posted-time-ago__text",
                          ".job-details-jobs-unified-top-card__posted-date")
    posted_date = resolve_posted_date(raw_posted)

    raw_desc    = sel_text(".show-more-less-html__markup", ".description__text")
    description = clean_description(raw_desc)

    # v7: detect language of the description so we know whether to run the
    # English-oriented paraphrase prompts.
    desc_lang = detect_text_language(description)

    salary = ""
    for s in [".compensation__salary",".salary","[class*='salary']","[class*='compensation']"]:
        el = soup.select_one(s)
        if el: salary = el.get_text(strip=True); break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|MUR|Rs\.?|SAR|salary|/yr|/hour|per month", t, re.I):
                salary = t; break

    raw_job_type      = get_job_criteria(soup, "Employment type") or workplace_type
    job_type          = raw_job_type or "Full-time"
    linkedin_function = get_job_criteria(soup, "Job function")
    linkedin_industry = get_job_criteria(soup, "Industries")

    real_deadline      = parse_deadline(soup)
    estimated_deadline = estimate_deadline_from_posted(posted_date) if not real_deadline else ""
    effective_deadline = real_deadline or estimated_deadline

    # ── LAYER 1: job-page extraction ─────────────────────────────────────────
    job_page_co = extract_company_from_job_page(html, soup)
    ld          = _parse_jsonld(html)

    # ── LAYER 2: LinkedIn company page ───────────────────────────────────────
    time.sleep(0.5)
    company = scrape_company_details(company_url_raw)

    # ── LAYER 3: merge ────────────────────────────────────────────────────────
    def _first(*vals) -> str:
        for v in vals:
            if v and str(v).strip(): return str(v).strip()
        return ""

    # v7: filter out LinkedIn login-wall placeholders ("Sign in", favicon logo)
    # before merging, so they never win the _first() priority chain.
    company_name_li = company.get("name", "")
    if is_bad_company_name(company_name_li):
        company_name_li = ""

    company_logo_li = company.get("logo", "")
    if is_placeholder_logo(company_logo_li):
        company_logo_li = ""

    company_name_jobpage = job_page_co.get("company_name", "")
    if is_bad_company_name(company_name_jobpage):
        company_name_jobpage = ""

    if is_bad_company_name(company_name_fallback):
        company_name_fallback = ""

    merged_name = _first(
        company_name_li,
        company_name_jobpage,
        company_name_fallback,
        extract_company_from_job_url(job_url),
    )
    merged_industry = _first(company.get("industry"), job_page_co.get("company_industry"), linkedin_industry)
    merged_logo_li  = _first(company_logo_li, job_page_co.get("company_logo"))   # raw LinkedIn logo
    merged_website  = _first(company.get("website"),  job_page_co.get("company_website"))
    merged_hq       = _first(company.get("headquarters"), job_page_co.get("company_address"))
    merged_founded  = _first(company.get("founded"),  job_page_co.get("company_founded"))
    merged_type     = _first(company.get("type"))
    merged_about    = _first(company.get("about"),    job_page_co.get("company_about"))
    company_url_out = _first(company.get("company_url"), company_url_raw)  # LinkedIn company URL, kept as-is

    # Ensure merged_website is not a LinkedIn URL
    merged_website = blank_if_linkedin(merged_website)

    # ── LAYER 4: deep company website crawl (apply + enrichment) ────────────
    site_info: dict = {}
    deep_enrich: dict = {}

    job_field = linkedin_function or merged_industry or infer_job_field(title, description)

    # ── LAYER 5: application details (v6 deep crawl + all fallbacks) ─────────
    time.sleep(0.2)
    apply_data = extract_application_details(
        job_url, soup, merged_website, ld,
        job_title=title,
        site_info=site_info,   # populated below if deep crawl wasn't already done
    )
    deep_enrich = apply_data.get("_deep") or {}

    # If the deep crawl above didn't happen (e.g. no website at all) but we
    # do have a website, run the standalone about/contact/footer scraper so
    # we still try to fill missing company fields + logo.
    if not deep_enrich and merged_website and not should_skip_crawl(merged_website):
        log.info(f"v6 about/contact/footer crawl (standalone): {merged_website}")
        deep_enrich = scrape_about_contact_footer(merged_website)

    # Use deep_enrich to fill in missing company fields
    if deep_enrich:
        if not merged_hq      and deep_enrich.get("address"):
            merged_hq = deep_enrich["address"]
        if not merged_founded and deep_enrich.get("founded"):
            merged_founded = deep_enrich["founded"]
        if not merged_about   and (deep_enrich.get("about") or deep_enrich.get("description")):
            merged_about = deep_enrich.get("about") or deep_enrich.get("description")
        if not merged_industry and deep_enrich.get("industry"):
            merged_industry = deep_enrich["industry"]

        phone = deep_enrich.get("phone")
        if phone:
            phone_note = f"Phone: {phone}"
            if phone_note not in (merged_about or ""):
                merged_about = (merged_about + "\n" + phone_note).strip() if merged_about else phone_note

    # If we still don't have a usable company name (e.g. LinkedIn page was a
    # login wall AND the job page had nothing), try the website domain itself.
    if is_bad_company_name(merged_name) and merged_website:
        try:
            netloc = urlparse(merged_website).netloc
            netloc = re.sub(r"^www\.", "", netloc)
            base   = netloc.split(".")[0]
            if base:
                merged_name = base.replace("-", " ").title()
        except Exception:
            pass

    # ── LOGO: prefer company website logo over LinkedIn's (which is often
    #          missing or a generic placeholder) ─────────────────────────────
    company_site_logo = (deep_enrich.get("logo") if deep_enrich else "") or ""
    if not company_site_logo and merged_website and not should_skip_crawl(merged_website):
        company_site_logo = get_company_website_logo(merged_website)

    if company_site_logo and not is_placeholder_logo(company_site_logo):
        merged_logo = company_site_logo
        logo_source = "company_website"
    elif merged_logo_li and not is_placeholder_logo(merged_logo_li):
        merged_logo = merged_logo_li
        logo_source = "linkedin"
    else:
        merged_logo = company_site_logo or merged_logo_li or ""
        logo_source = "fallback"

    raw_apply = ""
    if apply_data.get("email"):
        raw_apply = clean_email(apply_data["email"])
    elif apply_data.get("url") and apply_data.get("method") != "not_found":
        raw_apply = apply_data["url"]
    apply_link = clean_application_link(raw_apply)

    # Final safety net — blank any remaining LinkedIn URLs in apply/website
    apply_link     = blank_if_linkedin(apply_link)
    merged_website = blank_if_linkedin(merged_website)
    # Company URL (LinkedIn company page) is intentionally NOT blanked

    qualifications = extract_qualification(description)
    experience     = extract_experience(description)

    # ─────────────────────────────────────────────────────────────────────────
    #  ▶▶ PARAPHRASE
    # ─────────────────────────────────────────────────────────────────────────
    mark_scraped(job_id, job_url, title, merged_name)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    paraphrased_title = title
    paraphrased_desc  = description
    paraphrased_about = merged_about

    # v7: skip the English-oriented paraphrase prompts for predominantly
    # Arabic descriptions — they're kept as-is (now correctly preserved by
    # sanitize_text) and flagged via _lang for manual review/translation.
    if ENABLE_PARAPHRASE and MISTRAL_API_KEY and desc_lang != "ar":
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        if merged_about:
            paraphrased_about = paraphrase_company(merged_about)
        mark_paraphrased(job_id)
    elif desc_lang == "ar":
        print(C_DIM("  ⚠️  Paraphrasing skipped (description detected as Arabic)"))
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    return {
        # Paraphrased fields
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    paraphrased_about,
        # Original fields (for audit / duplicate detection)
        "originalTitle":     title,
        "originalDesc":      description,
        # Structured fields — match standardized column order
        "jobType":           job_type,
        "jobQualifications": qualifications,
        "jobExperience":     experience,
        "jobLocation":       location,
        "jobField":          job_field,
        "datePosted":        posted_date,
        "deadline":          effective_deadline,
        "application":       apply_link,          # LinkedIn URLs already blanked
        "companyUrl":        company_url_out,      # LinkedIn company page URL — KEPT
        "companyName":       merged_name,
        "companyLogo":       clean_logo_url(merged_logo),
        "companyIndustry":   merged_industry,
        "companyFounded":    merged_founded,
        "companyType":       merged_type,
        "companyWebsite":    merged_website,       # LinkedIn URLs already blanked
        "companyAddress":    merged_hq,
        "jobUrl":            job_url,
        "estimatedDeadline": estimated_deadline,
        "salaryRange":       salary,
        "_jobId":            job_id,
        "_apply_method":     apply_data.get("method", ""),
        "_logo_source":      logo_source,
        "_lang":             desc_lang,
    }

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str) -> int | None:
    if not name or not name.strip(): return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log.info(f"⏭ Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    # All URL fields already have LinkedIn blanked at scrape time
    # (Company URL is the LinkedIn company page and IS sent — it's a
    #  user-facing LinkedIn profile link, not an apply/application link)
    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "Full-time"))
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", "")) or sanitize_text(job.get("estimatedDeadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    industry    = sanitize_text(job.get("companyIndustry", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    co_founded  = sanitize_text(job.get("companyFounded", ""))
    co_type     = sanitize_text(job.get("companyType", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    # Extra safety check before posting (application + website only —
    # companyUrl is intentionally allowed to be a LinkedIn URL)
    application = blank_if_linkedin(application)
    co_website  = blank_if_linkedin(co_website)

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"],
                                                      "Referer": "https://www.linkedin.com/"}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    tagline = paraphrase_tagline(about[:300]) if about else ""

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_industry":   industry,
            "_company_address":    co_address,
            "_company_founded":    co_founded,
            "_company_type":       co_type,
            "_company_tagline":    tagline,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log.info(f"✅ Job posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2: time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(job: dict, index: int, total: int):
    desc         = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    desc_indented = "\n".join("   " + line for line in desc_preview.splitlines() if line.strip())
    apply        = job.get("application", "")
    logo         = job.get("companyLogo", "")
    logo_source  = job.get("_logo_source", "")
    orig_title   = job.get("originalTitle", "")
    method       = job.get("_apply_method", "")
    lang         = job.get("_lang", "")
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB {index}/{total}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}   : {C_VALUE(orig_title)}")
    print(f"  {C_LABEL('Title (paraphrased)')}: {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Language')}           : {lang or C_DIM('—')}")
    print(f"  {C_LABEL('Job Type')}            : {job.get('jobType','')}")
    print(f"  {C_LABEL('Field')}               : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}            : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Seniority')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualifications')}      : {job.get('jobQualifications','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}              : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Date Posted')}         : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}            : {job.get('deadline','') or C_DIM('—')}")
    print(f"  {C_LABEL('Apply Link')}          : {C_GREEN(apply) if apply else C_DIM('— not found / blanked —')}")
    print(f"  {C_LABEL('Apply Method')}        : {C_DIM(method) if method else C_DIM('—')}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}           : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Industry')}       : {job.get('companyIndustry','') or C_DIM('—')}")
    print(f"  {C_LABEL('Type')}           : {job.get('companyType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Founded')}        : {job.get('companyFounded','') or C_DIM('—')}")
    print(f"  {C_LABEL('Headquarters')}   : {job.get('companyAddress','') or C_DIM('—')}")
    print(f"  {C_LABEL('Website')}        : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Company URL')}    : {job.get('companyUrl','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}           : {logo if logo else C_DIM('— none —')}"
          + (f"  {C_DIM('(' + logo_source + ')')}" if logo else ""))
    about = job.get("companyDetails", "")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}          : {preview}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────────')}")
    print(desc_indented if desc_indented else C_DIM("   — no description —"))
    print(C_DIVIDER())

# =============================================================================
#  URL COLLECTION — GUEST API
# =============================================================================

def _build_guest_api_url(keyword: str, start: int) -> str:
    kw = quote_plus(keyword)
    return (
        "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
        f"?location=Libya&f_TPR=r604800&keywords={kw}&start={start}"
    )

def _collect_job_urls_from_cards(html: str, seen: set) -> list:
    found = []
    for raw_href in re.findall(r'href="(https?://[^"]*?/jobs/view/\d+[^"]*?)"', html):
        c = canonicalise_job_url(raw_href)
        if c and c not in seen: seen.add(c); found.append(c)
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if "/jobs/view/" not in href: continue
        if not href.startswith("http"): href = "https://www.linkedin.com" + href
        c = canonicalise_job_url(href)
        if c and c not in seen: seen.add(c); found.append(c)
    for sel in ["a.base-card__full-link","a.base-main-card__full-link",
                "a[data-tracking-control-name='public_jobs_jserp-name_click']",
                "a.job-card-list__title","a.job-card-container__link"]:
        for tag in soup.select(sel):
            href = tag.get("href", "")
            if "/jobs/view/" not in href: continue
            if not href.startswith("http"): href = "https://www.linkedin.com" + href
            c = canonicalise_job_url(href)
            if c and c not in seen: seen.add(c); found.append(c)
    return found

def _fetch_guest_api_page(keyword: str, start: int, retries: int = 3) -> str | None:
    url = _build_guest_api_url(keyword, start)
    for attempt in range(retries):
        try:
            time.sleep(DELAY_S + attempt * 3)
            r = requests.get(url, headers=_next_headers(), allow_redirects=True, timeout=25)
            if r.status_code == 429:
                wait = 60 + attempt * 60
                print(C_RED(f"  ⏳ Rate limited (429) — waiting {wait}s ..."))
                time.sleep(wait); continue
            if r.status_code in (400, 403, 999):
                log.warning(f"Blocked ({r.status_code}): {url}"); return None
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}"); return None
            text = r.text.strip()
            if not text:
                log.info(f"Empty body (start={start}, kw='{keyword}') — end of results.")
                return None
            return text
        except Exception as e:
            log.warning(f"Guest API error (attempt {attempt+1}, kw='{keyword}'): {e}")
            time.sleep(3 + attempt * 3)
    return None

def _paginate_keyword(keyword: str, seen: set) -> list:
    urls = []; page = 0; empty_streak = 0
    label = keyword if keyword else "(all)"
    while True:
        if MAX_PAGES and page >= MAX_PAGES: break
        start = page * 25
        print(f"  {C_DIM(f'[{label}] page {page+1} (start={start}) ...')}", flush=True)
        html = _fetch_guest_api_page(keyword, start)
        if html is None: break
        new_urls = _collect_job_urls_from_cards(html, seen)
        log.info(f"[{label}] page {page+1}: {len(new_urls)} new (total seen={len(seen)})")
        if new_urls: urls.extend(new_urls); empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= MAX_EMPTY_PAGES: break
        if start >= 975: break
        page += 1
        if page % 10 == 0:
            print(C_DIM("  Pausing 20s (every 10 pages) ..."))
            time.sleep(20)
    return urls

# =============================================================================
#  EXCEL SAVE  (standardized column order)
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details", "Job URL",
    "Estimated Deadline", "Salary Range",
]

def _save_excel(jobs: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"],
            job["jobType"],
            job["jobQualifications"],
            job["jobExperience"],
            job["jobLocation"],
            job["jobField"],
            job["datePosted"],
            job["deadline"],
            job["jobDescription"],
            job["application"],
            job["companyUrl"],
            job["companyName"],
            job["companyLogo"],
            job["companyIndustry"],
            job["companyFounded"],
            job["companyType"],
            job["companyWebsite"],
            job["companyAddress"],
            job["companyDetails"],
            job["jobUrl"],
            job["estimatedDeadline"],
            job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN CRAWL
# =============================================================================

def craw():
    start_time = time.time()
    _init_tracker()

    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  LINKEDIN JOB SCRAPER v7 + MISTRAL PARAPHRASE"))
    print(C_HEADER("=" * 72))
    print(f"  Keywords      : {len(SEARCH_KEYWORDS)}")
    print(f"  Max pages     : {'unlimited' if not MAX_PAGES else MAX_PAGES} per keyword")
    print(f"  Job cap       : {'none' if not JOB_LIMIT else JOB_LIMIT}")
    print(f"  Paraphrase    : {'✅ enabled' if ENABLE_PARAPHRASE else '❌ disabled'} (skipped for Arabic descriptions)")
    print(f"  Apply/Website : ❌ LinkedIn URLs BLOCKED (blanked)")
    print(f"  Company URL   : ✅ LinkedIn company page URL KEPT")
    print(f"  Logo priority : company-website logo > LinkedIn logo")
    print(f"  Company name  : LinkedIn page > job page > job-card > URL slug > website domain")
    print(f"  Email cleanup : known-TLD truncation (fixes '.comtak' style junk)")
    print(f"  NLP available : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started       : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 72))

    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    seen_urls = set(); all_job_urls = []; seen_content = set()

    for qi, keyword in enumerate(SEARCH_KEYWORDS):
        label = keyword if keyword else "(all)"
        print()
        print(C_BLUE(f"┌─ Keyword {qi+1}/{len(SEARCH_KEYWORDS)}: '{label}' ─────────────────"))
        new_urls = _paginate_keyword(keyword, seen_urls)
        all_job_urls.extend(new_urls)
        print(C_BLUE(f"└─ Found {len(new_urls)} new jobs (running total: {len(all_job_urls)})"))
        if JOB_LIMIT and len(all_job_urls) >= JOB_LIMIT: break
        time.sleep(DELAY_S * 2)

    if JOB_LIMIT and len(all_job_urls) > JOB_LIMIT:
        all_job_urls = all_job_urls[:JOB_LIMIT]

    print()
    print(C_HEADER(f"  Total unique URLs collected: {len(all_job_urls)}"))
    print()

    jobs = []; errors = 0
    for j, url in enumerate(all_job_urls):
        print(f"\n{C_HEADER(f'>>> Scraping job {j+1}/{len(all_job_urls)} ...')}")
        log.info(f"URL: {url}")
        try:
            job = scrape_job_details(url, processed_ids, processed_urls)
            if job and job.get("jobTitle"):
                fp = (
                    (job.get("originalTitle") or "").lower().strip(),
                    (job.get("companyName")   or "").lower().strip(),
                    (job.get("jobLocation")   or "").lower().strip(),
                )
                if fp in seen_content:
                    print(C_DIM(f"  ⧳  Duplicate content — skipped"))
                else:
                    seen_content.add(fp)
                    jobs.append(job)
                    print_job_verbose(job, j+1, len(all_job_urls))

                    print(C_BLUE(f"\n  📤 Posting to WordPress …"))
                    wp_id, wp_url = post_job_to_wordpress(job)
                    if wp_id:
                        mark_posted(job["_jobId"], wp_id, wp_url or "")
                        print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
                    else:
                        mark_failed(job["_jobId"], "wp_post_failed")
                        print(C_RED("  ❌ WordPress post failed"))
            else:
                print(C_RED("  ✗  No title found / skipped"))
        except Exception as e:
            errors += 1
            print(C_RED(f"  ✗  ERROR: {e}"))
            log.warning(f"Job error: {e}")

        time.sleep(DELAY_S)
        if len(jobs) % 50 == 0 and len(jobs) > 0:
            _save_excel(jobs)

    _save_excel(jobs)

    mins = round((time.time() - start_time) / 60, 1)
    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 72))
    print(f"  {C_LABEL('Total scraped')}  : {C_GREEN(str(len(jobs)))} jobs")
    print(f"  {C_LABEL('Errors')}         : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}       : ~{mins} min")
    print(f"  {C_LABEL('Output file')}    : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}   : {PROCESSED_IDS_FILE}")

    if jobs:
        from collections import Counter

        fields = Counter(j.get("jobField") or "Unknown" for j in jobs)
        print(f"\n  {C_LABEL('Jobs by field:')}")
        for field, count in fields.most_common():
            print(f"    {field:<35} {'█'*min(count,40)} {count}")

        with_apply = sum(1 for j in jobs if j.get("application"))
        with_email = sum(1 for j in jobs if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(jobs) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

        methods = Counter(j.get("_apply_method", "unknown") for j in jobs)
        print(f"\n  {C_LABEL('Apply method breakdown (v7):')}")
        for method, count in methods.most_common():
            print(f"    {method:<35} {count}")

        logo_sources = Counter(j.get("_logo_source", "unknown") for j in jobs if j.get("companyLogo"))
        print(f"\n  {C_LABEL('Logo source breakdown:')}")
        for src, count in logo_sources.most_common():
            print(f"    {src:<25} {count}")

        langs = Counter(j.get("_lang", "unknown") for j in jobs)
        print(f"\n  {C_LABEL('Description language breakdown:')}")
        for lang, count in langs.most_common():
            print(f"    {lang:<10} {count}")

        fill_fields = [("companyUrl","Company URL"),("companyName","Company Name"),
                       ("companyIndustry","Company Industry"),("companyLogo","Company Logo"),
                       ("companyWebsite","Company Website"),("companyAddress","Company Address"),
                       ("companyFounded","Company Founded"),("companyDetails","Company Details")]
        print(f"\n  {C_LABEL('Company field fill-rate:')}")
        for key, label in fill_fields:
            filled = sum(1 for j in jobs if j.get(key))
            pct    = round(filled / len(jobs) * 100) if jobs else 0
            print(f"    {label:<20} {filled}/{len(jobs)} ({pct}%)")

        para_count = sum(1 for j in jobs if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs)}")

    print(C_HEADER("=" * 72))


if __name__ == "__main__":
    craw()
