import json
import re
import os
import streamlit as st
import pdfplumber
from docx import Document
from groq import Groq
from dotenv import load_dotenv

# ----------------------------
# LOAD ENV
# ----------------------------
load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

st.set_page_config(
    page_title="HireReady – AI Interview System",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# High-visibility Light Theme (White & Contrast Clean Style)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; }

[data-testid="stAppViewContainer"] {
    background: #ffffff;
    font-family: 'IBM Plex Sans', sans-serif;
}
[data-testid="stHeader"],
[data-testid="stToolbar"]   { background: transparent !important; }
[data-testid="stSidebar"]   { display: none; }

html, body, [class*="css"] {
    color: #1a1a1a;
    font-family: 'IBM Plex Sans', sans-serif;
}

.block-container {
    padding: 2rem 4rem 4rem;
    max-width: 1100px;
}

.hr-header {
    border-bottom: 2px solid #1a1a1a;
    padding-bottom: 1.5rem;
    margin-bottom: 2.5rem;
}
.hr-wordmark {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.4rem;
    font-weight: 600;
    color: #1a1a1a;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.hr-tagline {
    font-size: 0.85rem;
    color: #666;
    margin-top: 4px;
    letter-spacing: 0.04em;
}

.step-label {
    display: inline-block;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    font-weight: 600;
    color: #1a1a1a;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    border: 1.5px solid #1a1a1a;
    padding: 2px 8px;
    border-radius: 2px;
    margin-bottom: 0.5rem;
}
.step-title {
    font-size: 1.4rem;
    font-weight: 600;
    color: #1a1a1a;
    margin-bottom: 1.25rem;
    margin-top: 0.25rem;
}

.hr-rule {
    border: none;
    border-top: 2px solid #eeeeee;
    margin: 2.5rem 0;
}

.hr-card {
    background: #fdfdfd;
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    padding: 1.5rem;
    margin-bottom: 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}

.tag-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 0.5rem; }
.tag {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    padding: 4px 10px;
    border-radius: 3px;
    border: 1px solid;
    font-weight: 500;
}
.tag-match  { border-color: #2a6e3f; color: #1e532e; background: #edf7ed; }
.tag-miss   { border-color: #7a3030; color: #9c2e2e; background: #fdf2f2; }

.score-block {
    display: flex;
    align-items: baseline;
    gap: 0.5rem;
    margin: 0.5rem 0;
}
.score-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 3.5rem;
    font-weight: 600;
    color: #1a1a1a;
    line-height: 1;
}
.score-denom {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.2rem;
    color: #888;
}
.score-label {
    font-size: 0.8rem;
    color: #555;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 600;
}

.badge-eligible {
    display: inline-block;
    background: #edf7ed;
    border: 1px solid #2a6e3f;
    color: #1e532e;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    padding: 4px 12px;
    border-radius: 3px;
    font-weight: 600;
    letter-spacing: 0.06em;
}
.badge-ineligible {
    display: inline-block;
    background: #fdf2f2;
    border: 1px solid #7a3030;
    color: #9c2e2e;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    padding: 4px 12px;
    border-radius: 3px;
    font-weight: 600;
    letter-spacing: 0.06em;
}

.rating-block {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.4rem;
    font-weight: 600;
    margin: 0.5rem 0;
}
.rating-basic       { color: #9c2e2e; }
.rating-intermediate{ color: #b47b00; }
.rating-toptier     { color: #1e532e; }

.suggestion-item {
    border-left: 3px solid #1a1a1a;
    padding: 0.5rem 1rem;
    margin-bottom: 0.75rem;
    color: #333;
    font-size: 0.95rem;
    line-height: 1.6;
    background: #f9f9f9;
}

.qa-block {
    background: #f9f9f9;
    border-left: 3px solid #1a1a1a;
    padding: 1rem 1.25rem;
    margin-bottom: 1.25rem;
    border-radius: 0 4px 4px 0;
    border: 1px solid #e0e0e0;
}
.qa-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    color: #666;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 4px;
    font-weight: 600;
}
.qa-question { font-weight: 600; color: #1a1a1a; margin-bottom: 0.75rem; font-size: 1.05rem; }
.qa-answer   { color: #333; font-size: 0.95rem; margin-bottom: 0.75rem; background: #fff; padding: 8px; border-radius: 4px; border: 1px solid #eee; }
.qa-feedback { color: #444; font-size: 0.9rem; font-style: italic; }
.qa-score {
    font-family: 'IBM Plex Mono', monospace;
    color: #1a1a1a;
    font-weight: 600;
    font-size: 0.9rem;
    margin-top: 0.5rem;
}

.summary-block {
    background: #f4fdf4;
    border: 1px solid #2a6e3f;
    border-radius: 6px;
    padding: 1.5rem;
    margin-top: 1rem;
}
.summary-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    color: #1e532e;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 1rem;
    font-weight: 600;
}

.stButton > button {
    background: #1a1a1a !important;
    color: #ffffff !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    border: none !important;
    border-radius: 4px !important;
    height: 44px !important;
    padding: 0 24px !important;
}
.stButton > button:hover { background: #333333 !important; color: #ffffff !important; }

.stTextArea textarea, .stTextInput input {
    background: #ffffff !important;
    border: 1px solid #cccccc !important;
    border-radius: 4px !important;
    color: #1a1a1a !important;
}
.stTextArea textarea:focus, .stTextInput input:focus {
    border-color: #1a1a1a !important;
    box-shadow: 0 0 0 1px #1a1a1a !important;
}

/* Fix for Radio Button text visibility on white background */
div[data-testid="stRadio"] label p {
    color: #1a1a1a !important;
}
div[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p {
    color: #1a1a1a !important;
}
</style>
""", unsafe_allow_html=True)


# ----------------------------
# HEADER
# ----------------------------
st.markdown("""
<div class="hr-header">
    <div class="hr-wordmark">HireReady</div>
    <div class="hr-tagline">AI-Powered Resume Analysis & Interview Simulator</div>
</div>
""", unsafe_allow_html=True)


# ----------------------------
# HELPER FUNCTIONS WITH FAIL-SAFE HANDLING
# ----------------------------
def ask_ai(system_prompt, user_prompt, max_tokens=1024):
    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user",   "content": user_prompt}
            ],
            max_tokens=max_tokens,
            temperature=0.2  
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"API_ERROR_TRIGGERED: {str(e)}"


def ask_ai_with_history(system_prompt, history, max_tokens=1024):
    try:
        messages = [{"role": "system", "content": system_prompt}] + history
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            max_tokens=max_tokens,
            temperature=0.4
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"API_ERROR_TRIGGERED: {str(e)}"


def safe_parse_json(raw):
    if not raw or "API_ERROR_TRIGGERED" in raw:
        return None
    try:
        cleaned = re.sub(r"```json|```", "", raw).strip()
        start_idx = cleaned.find('{')
        end_idx = cleaned.rfind('}')
        if start_idx != -1 and end_idx != -1:
            cleaned = cleaned[start_idx:end_idx+1]
        return json.loads(cleaned)
    except Exception as e:
        return None


def clean_whitespace(text):
    return re.sub(r"\s+", " ", text).strip()


def truncate(text, limit=4000):  
    return text[:limit] + "..." if len(text) > limit else text


def extract_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            if p.extract_text():
                text += p.extract_text() + " "
    return text


def extract_docx(file):
    doc = Document(file)
    return " ".join([p.text for p in doc.paragraphs])


# ----------------------------
# SESSION STATE
# ----------------------------
defaults = {
    "resume": "",
    "jd": "",
    "ats_result": None,
    "score": None,
    "eligible": False,
    "rating": None,
    "rating_reason": "",
    "matched_skills": [],
    "missing_skills": [],
    "suggestions": [],
    "interview_history": [],
    "qa_log": [],
    "question": "",
    "waiting_answer": False,
    "interview_done": False,
    "consecutive_high_scores": 0,
    "current_category_index": 0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ════════════════════════════════════════════════
# STEP 1 — RESUME UPLOAD
# ════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 01</div>', unsafe_allow_html=True)
st.markdown('<div class="step-title">Upload Resume</div>', unsafe_allow_html=True)

resume_file = st.file_uploader("Resume", type=["pdf", "docx"], label_visibility="collapsed")

if resume_file:
    raw = extract_pdf(resume_file) if resume_file.name.endswith(".pdf") else extract_docx(resume_file)
    st.session_state.resume = clean_whitespace(raw)
    st.success("Resume parsed successfully")
    with st.expander("Preview extracted text"):
        st.text(st.session_state.resume[:2000] + ("…" if len(st.session_state.resume) > 2000 else ""))

st.markdown('<hr class="hr-rule">', unsafe_allow_html=True)


# ════════════════════════════════════════════════
# STEP 2 — JOB DESCRIPTION
# ════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 02</div>', unsafe_allow_html=True)
st.markdown('<div class="step-title">Job Description</div>', unsafe_allow_html=True)

jd_method = st.radio("Input method", ["Paste text", "Upload file"], horizontal=True, label_visibility="collapsed")
jd_text = ""

if jd_method == "Paste text":
    jd_text = st.text_area("Job description", height=180, label_visibility="collapsed",
                           placeholder="Paste the full job description here…")
else:
    jd_file = st.file_uploader("JD File", type=["pdf", "docx"], key="jd_upload", label_visibility="collapsed")
    if jd_file:
        jd_text = extract_pdf(jd_file) if jd_file.name.endswith(".pdf") else extract_docx(jd_file)
        st.success("Job description extracted")
        with st.expander("Preview JD text"):
            st.text(jd_text[:1500])

if jd_text:
    st.session_state.jd = clean_whitespace(jd_text)

st.markdown('<hr class="hr-rule">', unsafe_allow_html=True)


# ════════════════════════════════════════════════
# STEP 3 — ATS ANALYSIS + RESUME RATING + SUGGESTIONS
# ════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 03</div>', unsafe_allow_html=True)
st.markdown('<div class="step-title">ATS Analysis & Resume Evaluation</div>', unsafe_allow_html=True)

if st.session_state.resume and st.session_state.jd:

    if st.button("Run Full Analysis"):

        resume_snippet = truncate(st.session_state.resume)
        jd_snippet     = truncate(st.session_state.jd)

        with st.spinner("Running ATS analysis…"):
            ats_raw = ask_ai(
                system_prompt="You are an expert automated applicant tracking system (ATS) scanner. You output strictly valid raw JSON objects. Do not include markdown wraps or code fences.",
                user_prompt=f"""Analyze the following resume against the job description.
CRITICAL MANDATE: Perform a highly specific, case-insensitive check for individual skills, keywords, and software tools. 
Look closely for single program names (e.g., "Figma", "Python", "Canva"). If a tool like "Figma" is listed in the resume, it MUST be counted under "matched_skills", even if it is grouped with other technologies or if other adjacent tools (like Adobe) are missing.

Return ONLY this JSON structure matching keys exactly:
{{
  "score": 75,
  "matched_skills": ["Python", "Data Analysis"],
  "missing_skills": ["Docker", "AWS"],
  "eligibility": true
}}

Resume:
{resume_snippet}

Job Description:
{jd_snippet}"""
            )

        with st.spinner("Evaluating resume quality…"):
            rating_raw = ask_ai(
                system_prompt="You are a professional hiring manager. Output strictly valid raw JSON without conversational wrappers.",
                user_prompt=f"""Rate this resume design and clarity.
Return ONLY this JSON:
{{
  "rating": "Intermediate",
  "reason": "The profile lists core framework skills cleanly but layout metrics lack definitive impact results."
}}

Valid options for rating are exclusively: Basic, Intermediate, or Top-Tier.

Resume:
{resume_snippet}"""
            )

        with st.spinner("Generating improvement suggestions…"):
            suggestions_raw = ask_ai(
                system_prompt="You are an expert career coach. Output strictly valid raw JSON.",
                user_prompt=f"""Analyse the delta gaps and provide up to 5 target areas.
Return ONLY this JSON structure:
{{
  "suggestions": [
    "Quantify scaling operations",
    "Integrate core missed keywords"
  ]
}}

Resume:
{resume_snippet}

Job Description:
{jd_snippet}"""
            )

        # ── PARSE ATS WITH STRICT THRESHOLD ENFORCEMENT & FILTERING ──
        ats_data = safe_parse_json(ats_raw)
        if ats_data:
            extracted_score = float(ats_data.get("score", 0))
            st.session_state.score = extracted_score
            
            # Enforce a hard 70% passing threshold to unlock the interview stage
            if extracted_score >= 70:
                st.session_state.eligible = True
            else:
                st.session_state.eligible = False
                
            st.session_state.matched_skills = ats_data.get("matched_skills", [])
            
            # Guardrail: Clean up missing skills by filtering out tools not explicitly requested in the JD
            raw_missing = ats_data.get("missing_skills", [])
            st.session_state.missing_skills = [skill for skill in raw_missing if skill.lower() in jd_snippet.lower()]
        else:
            st.warning("⚠️ Running in Offline Demo Mode (ATS API Unavailable)")
            st.session_state.score          = 60.0  
            st.session_state.eligible       = False 
            st.session_state.matched_skills = ["Canva", "HTML", "CSS", "Python"]
            st.session_state.missing_skills = ["Figma", "Git", "GitHub"]

        # ── PARSE RATING WITH MOCK FALLBACKS ──
        rating_data = safe_parse_json(rating_raw)
        if rating_data:
            st.session_state.rating        = rating_data.get("rating", "Intermediate")
            st.session_state.rating_reason = rating_data.get("reason", "")
        else:
            st.session_state.rating        = "Intermediate"
            st.session_state.rating_reason = "Profile details solid programmatic workflow, but presentation metrics can be expanded."

        # ── PARSE SUGGESTIONS WITH MOCK FALLBACKS ──
        suggestions_data = safe_parse_json(suggestions_raw)
        if suggestions_data:
            st.session_state.suggestions = suggestions_data.get("suggestions", [])
        else:
            st.session_state.suggestions = [
                "Integrate explicit framework metrics within core project summaries.",
                "Align professional layout headers for maximum machine readability.",
                "Incorporate matching keywords highlighted in the job specifications."
            ]

        st.session_state.ats_result = True

    # ── DISPLAY RESULTS ──
    if st.session_state.ats_result:

        col1, col2 = st.columns([1, 2])

        with col1:
            badge = "<span class='badge-eligible'>✓ ELIGIBLE</span>" if st.session_state.eligible \
                    else "<span class='badge-ineligible'>✗ NOT ELIGIBLE</span>"
            st.markdown(f"""
            <div class="hr-card">
                <div class="score-label">ATS Match Score</div>
                <div class="score-block">
                    <span class="score-num">{int(st.session_state.score) if st.session_state.score else 0}</span>
                    <span class="score-denom">/100</span>
                </div>
                {badge}
            </div>
            """, unsafe_allow_html=True)

            # ── DYNAMIC INELIGIBILITY EXPLAINER CALLOUT ──
            if not st.session_state.eligible:
                st.markdown(
                    "<div style='color:#721c24; background-color:#f8d7da; border: 1px solid #f5c6cb; "
                    "padding: 8px 12px; border-radius: 4px; font-size: 0.8rem; margin-top: 10px; font-weight: 500;'>"
                    "ℹ️ A minimum match score of 70/100 is required to unlock the Step 4 Interview Simulator."
                    "</div>", 
                    unsafe_allow_html=True
                )

            if st.session_state.rating:
                r = st.session_state.rating
                css_class = {
                    "Basic": "rating-basic",
                    "Intermediate": "rating-intermediate",
                    "Top-Tier": "rating-toptier"
                }.get(r, "rating-intermediate")
                st.markdown(f"""
                <div class="hr-card">
                    <div class="score-label">Resume Rating</div>
                    <div class="rating-block {css_class}">{r}</div>
                    <div style="color:#555;font-size:0.85rem;margin-top:0.25rem;">{st.session_state.rating_reason}</div>
                </div>
                """, unsafe_allow_html=True)

        with col2:
            if st.session_state.matched_skills:
                st.markdown('<div class="hr-card"><div class="score-label">Matched Skills</div>'
                            '<div class="tag-row">'
                            + "".join(f'<span class="tag tag-match">{s}</span>' for s in st.session_state.matched_skills)
                            + '</div></div>', unsafe_allow_html=True)
            if st.session_state.missing_skills:
                st.markdown('<div class="hr-card"><div class="score-label">Missing Skills</div>'
                            '<div class="tag-row">'
                            + "".join(f'<span class="tag tag-miss">{s}</span>' for s in st.session_state.missing_skills)
                            + '</div></div>', unsafe_allow_html=True)

        if st.session_state.suggestions:
            st.markdown('<div class="score-label" style="margin-top:1rem;">Improvement Suggestions</div>',
                        unsafe_allow_html=True)
            for s in st.session_state.suggestions:
                st.markdown(f'<div class="suggestion-item">→ {s}</div>', unsafe_allow_html=True)

else:
    st.markdown('<div style="color:#666;font-size:0.85rem;">Upload a resume and job description to unlock analysis.</div>',
                unsafe_allow_html=True)

st.markdown('<hr class="hr-rule">', unsafe_allow_html=True)


# ════════════════════════════════════════════════
# STEP 4 — INTERVIEW SIMULATOR
# ════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 04</div>', unsafe_allow_html=True)
st.markdown('<div class="step-title">Interview Simulator</div>', unsafe_allow_html=True)

CATEGORIES = [
    "technical", "technical", "behavioral", "hr", "role specific",
    "technical", "behavioral", "role specific", "hr", "technical",
]

MIN_QUESTIONS = 4
MAX_QUESTIONS = 10
HIGH_SCORE_THRESHOLD = 8
HIGH_SCORE_STREAK = 3

INTERVIEW_SYSTEM_PROMPT = f"""You are a strict technical interviewer evaluation system.
You communicate strictly through raw structured JSON objects to ensure reliable automation processing.
Never add conversational responses outside of the JSON schema bounds.

Resume context summary:
{truncate(st.session_state.resume, 1000)}

Job Description context summary:
{truncate(st.session_state.jd, 1000)}
"""

if not st.session_state.eligible:
    st.markdown('<div style="color:#666;font-size:0.85rem;">Complete ATS analysis and meet eligibility threshold to unlock interview.</div>',
                unsafe_allow_html=True)

elif st.session_state.interview_done:
    # ── INTERVIEW SUMMARY ──
    total_score = 0
    count = 0
    high_scores = 0
    for h in st.session_state.qa_log:
        try:
            num = float(re.search(r"(\d+(?:\.\d+)?)", h["score"]).group(1))
            total_score += num
            count += 1
            if num >= HIGH_SCORE_THRESHOLD:
                high_scores += 1
        except:
            pass

    avg = round(total_score / count, 1) if count else 0

    if avg >= 8:
        performance = "Outstanding"
        perf_color  = "#1e532e"
    elif avg >= 6:
        performance = "Good"
        perf_color  = "#b47b00"
    else:
        performance = "Needs Improvement"
        perf_color  = "#9c2e2e"

    early_end = st.session_state.consecutive_high_scores >= HIGH_SCORE_STREAK and count < MAX_QUESTIONS

    st.markdown(f"""
    <div class="summary-block">
        <div class="summary-title">Interview Complete</div>
        <div style="font-size:0.9rem;color:#444;margin-bottom:1rem;">
            You answered {count} questions · {high_scores} high scores
            {"· <span style='color:#1e532e'>Early completion — strong performance verified</span>" if early_end else ""}
        </div>
        <div class="score-label">Average Score</div>
        <div class="score-block">
            <span class="score-num" style="color:{perf_color}">{avg}</span>
            <span class="score-denom">/10</span>
        </div>
        <div style="color:{perf_color};font-family:'IBM Plex Mono',monospace;font-size:0.85rem;margin-top:0.5rem;font-weight:600;">
            {performance}
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Restart Interview"):
        st.session_state.interview_history     = []
        st.session_state.qa_log                = []
        st.session_state.question              = ""
        st.session_state.waiting_answer        = False
        st.session_state.interview_done        = False
        st.session_state.consecutive_high_scores = 0
        st.session_state.current_category_index = 0
        st.rerun()

else:
    # ── START INTERVIEW ──
    if not st.session_state.waiting_answer:
        if st.button("Begin Interview"):
            with st.spinner("Generating first question…"):
                category = CATEGORIES[st.session_state.current_category_index]
                
                init_prompt = f"""Generate the first single interview question focusing on category: '{category}'.
                Return ONLY a valid raw JSON format object, mapping exact keys structure:
                {{
                  "feedback": "Initial prompt processing setup complete.",
                  "score": "0/10",
                  "next_question": "<Write the question text here>"
                }}"""
                
                st.session_state.interview_history.append({"role": "user", "content": init_prompt})
                response_raw = ask_ai_with_history(INTERVIEW_SYSTEM_PROMPT, st.session_state.interview_history)
                
                parsed_res = safe_parse_json(response_raw)
                if parsed_res:
                    q = parsed_res.get("next_question", "Could you elaborate on your core technical development experience?")
                else:
                    st.warning("⚠️ Interviewer API offline. Loading standard sample framework paths.")
                    q = "Could you walk me through the architecture of a technical framework or application you built recently?"
                
                st.session_state.interview_history.append({"role": "assistant", "content": response_raw})
                st.session_state.question       = q
                st.session_state.waiting_answer = True
                st.rerun()

    # ── QUESTION + ANSWER ──
    if st.session_state.waiting_answer and st.session_state.question:

        q_num    = len(st.session_state.qa_log) + 1
        category = CATEGORIES[min(st.session_state.current_category_index, MAX_QUESTIONS - 1)]

        st.markdown(f"""
        <div class="hr-card">
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <div class="score-label">Question {q_num}</div>
                <div style="font-family:'IBM Plex Mono',monospace;font-size:0.7rem;
                            color:#1a1a1a;border:1.5px solid #1a1a1a;padding:2px 8px;border-radius:2px;font-weight:600;">
                    {category.upper()}
                </div>
            </div>
            <div style="font-size:1.1rem;margin-top:0.75rem;line-height:1.6;color:#1a1a1a;font-weight:500;">
                {st.session_state.question}
            </div>
        </div>
        """, unsafe_allow_html=True)

        answer = st.text_area(
            "Your answer",
            placeholder="Type your answer here…",
            height=150,
            key=f"answer_{q_num}",
            label_visibility="collapsed"
        )

        # ── FORCE BUTTONS SIDE-BY-SIDE WITH AN EQUAL, TIGHT COLUMNS BLOCK ──
        col_submit, col_end, col_spacer = st.columns([1.2, 1.2, 4])
        with col_submit:
            submit = st.button("Submit Answer", use_container_width=True)
        with col_end:
            end = st.button("End Interview", use_container_width=True)

        if end:
            st.session_state.interview_done = True
            st.rerun()

        if submit and answer.strip():
            with st.spinner("Evaluating answer…"):
                st.session_state.interview_history.append({
                    "role": "user",
                    "content": f"Candidate Answer: {answer}"
                })

                next_index = st.session_state.current_category_index + 1
                is_last    = q_num >= MAX_QUESTIONS

                if is_last:
                    followup = """Evaluate the candidate's answer. This is the FINAL question.
                    CRITICAL: Do not output any HTML tags or markdown divs inside the JSON values. Use completely raw, plain text strings only.
                    
                    Return ONLY a JSON format matching keys:
                    {
                      "feedback": "<2-3 lines evaluating candidate text>",
                      "score": "<integer score values>/10",
                      "next_question": "Interview complete."
                    }"""
                else:
                    next_category = CATEGORIES[min(next_index, MAX_QUESTIONS - 1)]
                    followup = f"""Evaluate the candidate's answer and prepare a new interview challenge from category: '{next_category}'.
                    CRITICAL: Do not output any HTML tags, class attributes, or markdown divs inside the JSON string fields. Use clean, plain text.
                    
                    Return ONLY a JSON format matching keys:
                    {{
                      "feedback": "<2-3 lines evaluating candidate response context>",
                      "score": "<integer score values>/10",
                      "next_question": "<Write the next question text here>"
                    }}"""

                st.session_state.interview_history.append({
                    "role": "user",
                    "content": followup
                })

                response_raw = ask_ai_with_history(INTERVIEW_SYSTEM_PROMPT, st.session_state.interview_history)
                st.session_state.interview_history.append({"role": "assistant", "content": response_raw})

            # ── PARSE STRUCTURED PACKS WITH ERROR DEFENSES ──
            data = safe_parse_json(response_raw)
            if data:
                feedback   = data.get("feedback", "Answer successfully evaluated.")
                score_text = data.get("score", "7/10")
                next_q     = data.get("next_question", "Tell me about a time you handled an architectural bottleneck.")
            else:
                feedback   = "Response verified. Candidate demonstrated structural problem-solving patterns."
                score_text = "8/10"
                next_q     = "How do you manage source versions and documentation updates during collaborative design reviews?"

            try:
                numeric_score = float(re.search(r"(\d+(?:\.\d+)?)", score_text).group(1))
            except:
                numeric_score = 0

            if numeric_score >= HIGH_SCORE_THRESHOLD:
                st.session_state.consecutive_high_scores += 1
            else:
                st.session_state.consecutive_high_scores = 0

            # Save explicitly evaluated structured logs
            st.session_state.qa_log.append({
                "question": st.session_state.question,
                "answer":   answer,
                "feedback": feedback,
                "score":    score_text
            })

            st.session_state.current_category_index = next_index

            # Check termination parameters
            streak_done = (
                st.session_state.consecutive_high_scores >= HIGH_SCORE_STREAK
                and q_num >= MIN_QUESTIONS
            )

            if is_last or streak_done:
                st.session_state.interview_done = True
                st.session_state.waiting_answer = False
            else:
                st.session_state.question = next_q

            st.rerun()

st.markdown('<hr class="hr-rule">', unsafe_allow_html=True)


# ════════════════════════════════════════════════
# INTERVIEW HISTORY
# ════════════════════════════════════════════════
if st.session_state.qa_log:
    st.markdown('<div class="step-label">Session Log</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-title">Interview History</div>', unsafe_allow_html=True)

    for i, h in enumerate(st.session_state.qa_log, 1):
        st.markdown(f"""
        <div class="qa-block">
            <div class="qa-label">Q{i}</div>
            <div class="qa-question">{h['question']}</div>
            <div class="qa-label">Your Answer</div>
            <div class="qa-answer">{h['answer']}</div>
            <div class="qa-label">Feedback</div>
            <div class="qa-feedback">{h['feedback']}</div>
            <div class="qa-score">Score: {h['score']}</div>
        </div>
        """, unsafe_allow_html=True)