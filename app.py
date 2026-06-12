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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

*, *::before, *::after { box-sizing: border-box; }

[data-testid="stAppViewContainer"] {
    background: #0f0f0f;
    font-family: 'IBM Plex Sans', sans-serif;
}
[data-testid="stHeader"],
[data-testid="stToolbar"]   { background: transparent !important; }
[data-testid="stSidebar"]   { display: none; }

html, body, [class*="css"] {
    color: #e8e3d9;
    font-family: 'IBM Plex Sans', sans-serif;
}

.block-container {
    padding: 2rem 4rem 4rem;
    max-width: 1100px;
}

.hr-header {
    border-bottom: 1px solid #2a2a2a;
    padding-bottom: 1.5rem;
    margin-bottom: 2.5rem;
}
.hr-wordmark {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.1rem;
    font-weight: 600;
    color: #f0a500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.hr-tagline {
    font-size: 0.78rem;
    color: #555;
    margin-top: 2px;
    letter-spacing: 0.04em;
}

.step-label {
    display: inline-block;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.68rem;
    font-weight: 600;
    color: #f0a500;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    border: 1px solid #f0a500;
    padding: 2px 8px;
    border-radius: 2px;
    margin-bottom: 0.5rem;
}
.step-title {
    font-size: 1.25rem;
    font-weight: 500;
    color: #e8e3d9;
    margin-bottom: 1.25rem;
    margin-top: 0.25rem;
}

.hr-rule {
    border: none;
    border-top: 1px solid #1e1e1e;
    margin: 2.5rem 0;
}

.hr-card {
    background: #161616;
    border: 1px solid #242424;
    border-radius: 6px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}

.tag-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 0.5rem; }
.tag {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    padding: 3px 10px;
    border-radius: 3px;
    border: 1px solid;
}
.tag-match  { border-color: #2a6e3f; color: #4caf78; background: #0d1f14; }
.tag-miss   { border-color: #7a3030; color: #e05a5a; background: #1a0d0d; }

.score-block {
    display: flex;
    align-items: baseline;
    gap: 0.5rem;
    margin: 1rem 0;
}
.score-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 3.5rem;
    font-weight: 600;
    color: #f0a500;
    line-height: 1;
}
.score-denom {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.2rem;
    color: #444;
}
.score-label {
    font-size: 0.78rem;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-left: 4px;
}

.badge-eligible {
    display: inline-block;
    background: #0d1f14;
    border: 1px solid #2a6e3f;
    color: #4caf78;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    padding: 4px 12px;
    border-radius: 3px;
    letter-spacing: 0.06em;
}
.badge-ineligible {
    display: inline-block;
    background: #1a0d0d;
    border: 1px solid #7a3030;
    color: #e05a5a;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    padding: 4px 12px;
    border-radius: 3px;
    letter-spacing: 0.06em;
}

.rating-block {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.4rem;
    font-weight: 600;
    margin: 0.5rem 0;
}
.rating-basic       { color: #e05a5a; }
.rating-intermediate{ color: #f0a500; }
.rating-toptier     { color: #4caf78; }

.suggestion-item {
    border-left: 2px solid #f0a500;
    padding: 0.5rem 1rem;
    margin-bottom: 0.75rem;
    color: #bbb;
    font-size: 0.9rem;
    line-height: 1.6;
}

.qa-block {
    background: #111;
    border-left: 2px solid #f0a500;
    padding: 1rem 1.25rem;
    margin-bottom: 1.25rem;
    border-radius: 0 4px 4px 0;
}
.qa-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    color: #555;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 4px;
}
.qa-question { font-weight: 500; color: #e8e3d9; margin-bottom: 0.75rem; }
.qa-answer   { color: #aaa; font-size: 0.9rem; margin-bottom: 0.75rem; }
.qa-feedback { color: #ccc; font-size: 0.85rem; font-style: italic; }
.qa-score {
    font-family: 'IBM Plex Mono', monospace;
    color: #f0a500;
    font-size: 0.85rem;
    margin-top: 0.5rem;
}

.summary-block {
    background: #161616;
    border: 1px solid #2a6e3f;
    border-radius: 6px;
    padding: 1.5rem;
    margin-top: 1rem;
}
.summary-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #4caf78;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 1rem;
}

.stButton > button {
    background: #f0a500 !important;
    color: #0f0f0f !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    border: none !important;
    border-radius: 3px !important;
    height: 44px !important;
    padding: 0 24px !important;
    width: auto !important;
}
.stButton > button:hover { background: #ffc13a !important; }

.stTextArea textarea, .stTextInput input {
    background: #161616 !important;
    border: 1px solid #2a2a2a !important;
    border-radius: 4px !important;
    color: #e8e3d9 !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
}
.stTextArea textarea:focus, .stTextInput input:focus {
    border-color: #f0a500 !important;
    box-shadow: 0 0 0 1px #f0a500 !important;
}
[data-testid="stFileUploader"] {
    background: #161616 !important;
    border: 1px dashed #2a2a2a !important;
    border-radius: 4px !important;
    padding: 1rem !important;
}
div[data-testid="stExpander"] {
    background: #161616;
    border: 1px solid #242424;
    border-radius: 4px;
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
# HELPER FUNCTIONS
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
            temperature=0.7
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"API ERROR: {str(e)}"


def ask_ai_with_history(system_prompt, history, max_tokens=1024):
    try:
        messages = [{"role": "system", "content": system_prompt}] + history
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            max_tokens=max_tokens,
            temperature=0.7
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"API ERROR: {str(e)}"


def safe_parse_json(raw):
    cleaned = re.sub(r"```json|```", "", raw).strip()
    return json.loads(cleaned)


def clean_whitespace(text):
    return re.sub(r"\s+", " ", text).strip()


def truncate(text, limit=1500):
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
                system_prompt="You are an ATS system. Return ONLY valid JSON with no markdown, no explanation, no backticks.",
                user_prompt=f"""Analyse this resume against the job description.

Return ONLY this JSON structure:
{{
  "score": <integer 0-100>,
  "matched_skills": [<list of strings>],
  "missing_skills": [<list of strings>],
  "eligibility": <true or false>
}}

Resume:
{resume_snippet}

Job Description:
{jd_snippet}"""
            )

        with st.spinner("Evaluating resume quality…"):
            rating_raw = ask_ai(
                system_prompt="You are a professional resume evaluator. Return ONLY valid JSON with no markdown, no backticks.",
                user_prompt=f"""Rate this resume quality.

Return ONLY this JSON:
{{
  "rating": "<Basic | Intermediate | Top-Tier>",
  "reason": "<one sentence explanation>"
}}

Resume:
{resume_snippet}"""
            )

        with st.spinner("Generating improvement suggestions…"):
            suggestions_raw = ask_ai(
                system_prompt="You are a professional career coach. Return ONLY valid JSON with no markdown, no backticks.",
                user_prompt=f"""Analyse this resume against the job description and give improvement suggestions.

Return ONLY this JSON:
{{
  "suggestions": [
    "<specific suggestion 1>",
    "<specific suggestion 2>",
    "<specific suggestion 3>",
    "<specific suggestion 4>",
    "<specific suggestion 5>"
  ]
}}

Resume:
{resume_snippet}

Job Description:
{jd_snippet}"""
            )

        # ── PARSE ATS ──
        try:
            ats_data = safe_parse_json(ats_raw)
            st.session_state.score          = float(ats_data.get("score", 0))
            st.session_state.eligible       = ats_data.get("eligibility", False)
            st.session_state.matched_skills = ats_data.get("matched_skills", [])
            st.session_state.missing_skills = ats_data.get("missing_skills", [])
        except:
            st.session_state.score          = 0
            st.session_state.eligible       = False
            st.session_state.matched_skills = []
            st.session_state.missing_skills = []
            st.error("ATS parsing failed. Raw output:")
            st.code(ats_raw)

        # ── PARSE RATING ──
        try:
            rating_data = safe_parse_json(rating_raw)
            st.session_state.rating        = rating_data.get("rating", "")
            st.session_state.rating_reason = rating_data.get("reason", "")
        except:
            st.session_state.rating        = ""
            st.session_state.rating_reason = ""

        # ── PARSE SUGGESTIONS ──
        try:
            suggestions_data             = safe_parse_json(suggestions_raw)
            st.session_state.suggestions = suggestions_data.get("suggestions", [])
        except:
            st.session_state.suggestions = []

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

            if st.session_state.rating:
                r = st.session_state.rating
                css_class = {
                    "Basic": "rating-basic",
                    "Intermediate": "rating-intermediate",
                    "Top-Tier": "rating-toptier"
                }.get(r, "")
                st.markdown(f"""
                <div class="hr-card">
                    <div class="score-label">Resume Rating</div>
                    <div class="rating-block {css_class}">{r}</div>
                    <div style="color:#666;font-size:0.82rem;margin-top:0.25rem;">{st.session_state.rating_reason}</div>
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
    st.markdown('<div style="color:#555;font-size:0.85rem;">Upload a resume and job description to unlock analysis.</div>',
                unsafe_allow_html=True)

st.markdown('<hr class="hr-rule">', unsafe_allow_html=True)


# ════════════════════════════════════════════════
# STEP 4 — INTERVIEW SIMULATOR
# ════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 04</div>', unsafe_allow_html=True)
st.markdown('<div class="step-title">Interview Simulator</div>', unsafe_allow_html=True)

INTERVIEW_SYSTEM_PROMPT = f"""You are a strict but fair technical interviewer conducting a real job interview.

You have access to the candidate's resume and the job description.
Ask questions that are specific to their background and the role.
Do not repeat questions already asked.
Keep track of the conversation and build on previous answers.

Resume summary:
{truncate(st.session_state.resume, 800)}

Job Description summary:
{truncate(st.session_state.jd, 800)}
"""

MAX_QUESTIONS = 5

if not st.session_state.eligible:
    st.markdown('<div style="color:#555;font-size:0.85rem;">Complete ATS analysis and meet eligibility threshold to unlock interview.</div>',
                unsafe_allow_html=True)

elif st.session_state.interview_done:
    total_score = 0
    count = 0
    for h in st.session_state.qa_log:
        try:
            num = float(re.search(r"(\d+(?:\.\d+)?)", h["score"]).group(1))
            total_score += num
            count += 1
        except:
            pass

    avg = round(total_score / count, 1) if count else 0

    st.markdown(f"""
    <div class="summary-block">
        <div class="summary-title">Interview Complete</div>
        <div style="font-size:0.9rem;color:#aaa;margin-bottom:1rem;">
            You answered {len(st.session_state.qa_log)} questions.
        </div>
        <div class="score-label">Average Score</div>
        <div class="score-block">
            <span class="score-num">{avg}</span>
            <span class="score-denom">/10</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Restart Interview"):
        st.session_state.interview_history = []
        st.session_state.qa_log            = []
        st.session_state.question          = ""
        st.session_state.waiting_answer    = False
        st.session_state.interview_done    = False
        st.rerun()

else:
    if not st.session_state.waiting_answer:
        if st.button("Begin Interview"):
            with st.spinner("Generating first question…"):
                st.session_state.interview_history.append({
                    "role": "user",
                    "content": "Start the interview. Ask me the first question."
                })
                q = ask_ai_with_history(INTERVIEW_SYSTEM_PROMPT, st.session_state.interview_history)
                st.session_state.interview_history.append({"role": "assistant", "content": q})
                st.session_state.question       = q
                st.session_state.waiting_answer = True
                st.rerun()

    if st.session_state.waiting_answer and st.session_state.question:

        q_num = len(st.session_state.qa_log) + 1

        st.markdown(f"""
        <div class="hr-card">
            <div class="score-label">Question {q_num} of {MAX_QUESTIONS}</div>
            <div style="font-size:1.05rem;margin-top:0.5rem;line-height:1.6;">
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

        col_submit, col_end = st.columns([1, 5])
        with col_submit:
            submit = st.button("Submit Answer")
        with col_end:
            end = st.button("End Interview")

        if end:
            st.session_state.interview_done = True
            st.rerun()

        if submit and answer.strip():
            with st.spinner("Evaluating answer…"):
                st.session_state.interview_history.append({
                    "role": "user",
                    "content": answer
                })

                if q_num >= MAX_QUESTIONS:
                    followup = "Give feedback on this answer. This was the last question so do NOT ask another one."
                else:
                    followup = "Give feedback on this answer in 2-3 lines, give a score out of 10, then ask the next interview question."

                st.session_state.interview_history.append({
                    "role": "user",
                    "content": followup
                })

                response = ask_ai_with_history(INTERVIEW_SYSTEM_PROMPT, st.session_state.interview_history)
                st.session_state.interview_history.append({"role": "assistant", "content": response})

            feedback, score, next_q = "", "", ""

            if "Feedback:" in response:
                feedback = response.split("Feedback:")[1].split("Score:")[0].strip()
            else:
                feedback = response.split("\n")[0] if response else ""

            score_match = re.search(r"[Ss]core[:\s]+(\d+(?:\.\d+)?)\s*/\s*10", response)
            score = score_match.group(0) if score_match else "Score not parsed"

            if q_num < MAX_QUESTIONS:
                lines = response.strip().split("\n")
                next_q = lines[-1].strip() if lines else "Tell me about a challenge you faced in a project."

            st.session_state.qa_log.append({
                "question": st.session_state.question,
                "answer":   answer,
                "feedback": feedback,
                "score":    score
            })

            if q_num >= MAX_QUESTIONS:
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
            <div class="qa-score">{h['score']}</div>
        </div>
        """, unsafe_allow_html=True)