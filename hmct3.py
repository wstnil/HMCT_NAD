import streamlit as st
import csv
from collections import OrderedDict
from datetime import datetime
import re
from openpyxl import Workbook
from io import BytesIO

# =========================
# PAGE CONFIG & CUSTOM CSS
# =========================

st.set_page_config(
    page_title="NAD Converter v2",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

    :root {
        --bg-primary: #020617;
        --bg-secondary: #0f172a;
        --bg-card: #111827;
        --bg-card-hover: #1e293b;
        --accent: #38bdf8;
        --accent-glow: rgba(56, 189, 248, 0.15);
        --success: #22c55e;
        --warning: #facc15;
        --danger: #ef4444;
        --text-primary: #f8fafc;
        --text-secondary: #94a3b8;
        --text-muted: #475569;
        --border: rgba(56, 189, 248, 0.12);
    }

    html, body, [class*="st-"] {
        font-family: 'Inter', sans-serif !important;
    }

    .stApp {
        background: linear-gradient(180deg, #020617 0%, #0a0f1e 40%, #0f172a 100%) !important;
        min-height: 100vh;
    }

    /* Hide default Streamlit elements */
    #MainMenu, footer, header { visibility: hidden !important; }
    .stApp > header { display: none !important; }

    /* Main container */
    .main-container {
        max-width: 780px;
        margin: 0 auto;
        padding: 2rem 1.5rem;
    }

    /* Logo / Brand Section */
    .brand-section {
        text-align: center;
        padding: 3rem 0 1.5rem 0;
        position: relative;
    }

    .brand-icon {
        width: 72px;
        height: 72px;
        background: linear-gradient(135deg, #38bdf8 0%, #818cf8 50%, #c084fc 100%);
        border-radius: 20px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 32px;
        margin-bottom: 1.25rem;
        box-shadow: 0 0 40px rgba(56, 189, 248, 0.25), 0 0 80px rgba(129, 140, 248, 0.1);
        animation: float 4s ease-in-out infinite;
    }

    @keyframes float {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-8px); }
    }

    .brand-title {
        font-size: 2.5rem;
        font-weight: 900;
        letter-spacing: -0.03em;
        background: linear-gradient(135deg, #f8fafc 0%, #38bdf8 60%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        line-height: 1.1;
        margin-bottom: 0.5rem;
    }

    .brand-subtitle {
        font-size: 0.95rem;
        color: var(--text-secondary);
        font-weight: 400;
        letter-spacing: 0.02em;
    }

    .brand-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        margin-top: 1rem;
        padding: 6px 16px;
        background: rgba(56, 189, 248, 0.08);
        border: 1px solid var(--border);
        border-radius: 100px;
        font-size: 0.78rem;
        color: var(--accent);
        font-weight: 500;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }

    .brand-badge .dot {
        width: 6px;
        height: 6px;
        background: var(--success);
        border-radius: 50%;
        animation: pulse-dot 2s ease-in-out infinite;
    }

    @keyframes pulse-dot {
        0%, 100% { opacity: 1; transform: scale(1); }
        50% { opacity: 0.5; transform: scale(0.8); }
    }

    /* Card */
    .card {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 2rem;
        margin-top: 2rem;
        position: relative;
        overflow: hidden;
        transition: border-color 0.3s ease, box-shadow 0.3s ease;
    }

    .card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 1px;
        background: linear-gradient(90deg, transparent, var(--accent), transparent);
        opacity: 0.4;
    }

    .card:hover {
        border-color: rgba(56, 189, 248, 0.25);
        box-shadow: 0 0 40px rgba(56, 189, 248, 0.05);
    }

    .card-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 1.5rem;
    }

    .card-icon {
        width: 40px;
        height: 40px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        flex-shrink: 0;
    }

    .card-icon.upload {
        background: rgba(56, 189, 248, 0.1);
        border: 1px solid rgba(56, 189, 248, 0.15);
    }

    .card-icon.settings {
        background: rgba(129, 140, 248, 0.1);
        border: 1px solid rgba(129, 140, 248, 0.15);
    }

    .card-icon.result {
        background: rgba(34, 197, 94, 0.1);
        border: 1px solid rgba(34, 197, 94, 0.15);
    }

    .card-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: var(--text-primary);
        letter-spacing: -0.01em;
    }

    .card-desc {
        font-size: 0.82rem;
        color: var(--text-muted);
        margin-top: 2px;
    }

    /* Upload Area */
    .upload-zone {
        border: 2px dashed rgba(56, 189, 248, 0.2);
        border-radius: 12px;
        padding: 2.5rem 2rem;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s ease;
        background: rgba(56, 189, 248, 0.02);
        position: relative;
    }

    .upload-zone:hover {
        border-color: rgba(56, 189, 248, 0.4);
        background: rgba(56, 189, 248, 0.04);
    }

    .upload-zone.active {
        border-color: var(--accent);
        background: rgba(56, 189, 248, 0.06);
        box-shadow: 0 0 30px rgba(56, 189, 248, 0.08);
    }

    .upload-icon {
        font-size: 2.5rem;
        margin-bottom: 0.75rem;
        opacity: 0.7;
    }

    .upload-text {
        font-size: 0.95rem;
        color: var(--text-secondary);
        font-weight: 500;
    }

    .upload-hint {
        font-size: 0.78rem;
        color: var(--text-muted);
        margin-top: 0.35rem;
    }

    /* File info bar */
    .file-info {
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 1rem 1.25rem;
        background: rgba(34, 197, 94, 0.06);
        border: 1px solid rgba(34, 197, 94, 0.15);
        border-radius: 10px;
        margin-top: 1rem;
        animation: slideIn 0.3s ease;
    }

    @keyframes slideIn {
        from { opacity: 0; transform: translateY(-8px); }
        to { opacity: 1; transform: translateY(0); }
    }

    .file-info-icon {
        font-size: 1.5rem;
    }

    .file-info-name {
        font-size: 0.88rem;
        color: var(--text-primary);
        font-weight: 600;
        flex: 1;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .file-info-size {
        font-size: 0.75rem;
        color: var(--text-muted);
    }

    .file-info-check {
        color: var(--success);
        font-size: 1.2rem;
    }

    /* Settings Grid */
    .settings-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 1rem;
    }

    .settings-grid .full-width {
        grid-column: 1 / -1;
    }

    .setting-label {
        font-size: 0.78rem;
        color: var(--text-muted);
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 6px;
        display: block;
    }

    /* Convert Button */
    .convert-btn-wrapper {
        margin-top: 2rem;
        text-align: center;
    }

    .convert-btn {
        display: inline-flex;
        align-items: center;
        gap: 10px;
        padding: 14px 48px;
        background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%);
        color: white;
        font-size: 1rem;
        font-weight: 700;
        border: none;
        border-radius: 12px;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 4px 20px rgba(34, 197, 94, 0.3), 0 0 0 0 rgba(34, 197, 94, 0.2);
        letter-spacing: 0.01em;
    }

    .convert-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 30px rgba(34, 197, 94, 0.4), 0 0 0 4px rgba(34, 197, 94, 0.1);
    }

    .convert-btn:active {
        transform: translateY(0);
    }

    .convert-btn:disabled {
        opacity: 0.4;
        cursor: not-allowed;
        transform: none !important;
        box-shadow: none !important;
    }

    .convert-btn .btn-icon {
        font-size: 1.2rem;
    }

    /* Result Card */
    .result-card {
        background: rgba(34, 197, 94, 0.04);
        border: 1px solid rgba(34, 197, 94, 0.15);
        border-radius: 12px;
        padding: 1.5rem;
        margin-top: 1.5rem;
        animation: slideIn 0.4s ease;
    }

    .result-stats {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin-bottom: 1.25rem;
    }

    .stat-box {
        text-align: center;
        padding: 1rem 0.5rem;
        background: rgba(255, 255, 255, 0.02);
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.04);
    }

    .stat-value {
        font-size: 1.75rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38bdf8, #818cf8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        line-height: 1;
        margin-bottom: 4px;
    }

    .stat-label {
        font-size: 0.72rem;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 600;
    }

    .download-btn {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 12px 32px;
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: white;
        font-size: 0.9rem;
        font-weight: 700;
        border: none;
        border-radius: 10px;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 4px 20px rgba(37, 99, 235, 0.3);
        text-decoration: none;
        width: 100%;
        justify-content: center;
        box-sizing: border-box;
    }

    .download-btn:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 25px rgba(37, 99, 235, 0.4);
    }

    /* Error Card */
    .error-card {
        background: rgba(239, 68, 68, 0.06);
        border: 1px solid rgba(239, 68, 68, 0.2);
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        margin-top: 1.5rem;
        animation: slideIn 0.3s ease;
    }

    .error-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #fca5a5;
        margin-bottom: 0.4rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .error-msg {
        font-size: 0.82rem;
        color: #f87171;
        line-height: 1.5;
    }

    /* Progress Bar */
    .progress-container {
        margin: 1.5rem 0;
    }

    .progress-bar-bg {
        height: 6px;
        background: rgba(255, 255, 255, 0.06);
        border-radius: 100px;
        overflow: hidden;
    }

    .progress-bar-fill {
        height: 100%;
        background: linear-gradient(90deg, #38bdf8, #818cf8, #c084fc);
        border-radius: 100px;
        transition: width 0.5s ease;
        position: relative;
    }

    .progress-bar-fill::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
        animation: shimmer 1.5s infinite;
    }

    @keyframes shimmer {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(100%); }
    }

    .progress-text {
        font-size: 0.78rem;
        color: var(--text-muted);
        text-align: center;
        margin-top: 0.5rem;
        font-weight: 500;
    }

    /* Features Strip */
    .features-strip {
        display: flex;
        justify-content: center;
        gap: 2rem;
        margin-top: 2.5rem;
        padding: 1.25rem 0;
        border-top: 1px solid rgba(255, 255, 255, 0.04);
    }

    .feature-item {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.78rem;
        color: var(--text-muted);
        font-weight: 500;
    }

    .feature-item .fi-icon {
        font-size: 0.9rem;
        opacity: 0.7;
    }

    /* Footer */
    .footer-text {
        text-align: center;
        font-size: 0.72rem;
        color: var(--text-muted);
        margin-top: 1.5rem;
        padding-bottom: 2rem;
        opacity: 0.6;
    }

    /* Override Streamlit default styling for inputs */
    div[data-testid="stTextInput"] label,
    div[data-testid="stSelectbox"] label {
        color: var(--text-muted) !important;
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.06em !important;
    }

    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        background: var(--bg-secondary) !important;
        border: 1px solid var(--border) !important;
        border-radius: 8px !important;
        color: var(--text-primary) !important;
        font-size: 0.88rem !important;
        padding: 10px 14px !important;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: var(--accent) !important;
        box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.15) !important;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(17, 24, 39, 0.95) !important;
        border: 1px solid rgba(56, 189, 248, 0.18) !important;
        border-radius: 18px !important;
        padding: 14px !important;
        margin-top: 0.75rem;
    }

    div[data-testid="stFileUploader"] section {
        min-height: 140px;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: rgba(15, 23, 42, 0.95) !important;
        border: 1px dashed rgba(56, 189, 248, 0.25) !important;
        border-radius: 14px !important;
        padding: 20px !important;
        cursor: pointer;
    }

    div[data-testid="stFileUploader"] button {
        display: inline-flex !important;
        margin-top: 12px !important;
    }

    div[data-testid="stFileUploader"] div[data-testid="stFileDropZone"] {
        display: block !important;
    }

    /* Spinner override */
    .stSpinner > div {
        border-top-color: var(--accent) !important;
    }

    /* Scrollbar */
    ::-webkit-scrollbar { width: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.2); }

    /* Processing overlay text */
    .processing-text {
        text-align: center;
        padding: 1.5rem;
    }

    .processing-text .spinner-icon {
        font-size: 2rem;
        display: inline-block;
        animation: spin 1s linear infinite;
    }

    @keyframes spin {
        from { transform: rotate(0deg); }
        to { transform: rotate(360deg); }
    }

    .processing-text .label {
        font-size: 0.9rem;
        color: var(--text-secondary);
        font-weight: 600;
        margin-top: 0.75rem;
    }

    .processing-text .sublabel {
        font-size: 0.78rem;
        color: var(--text-muted);
        margin-top: 0.25rem;
    }
</style>
""", unsafe_allow_html=True)


# =========================
# HELPERS (ALL LOGIC PRESERVED)
# =========================

def normalize_gender(g):
    """Normalize gender to M/F/O."""
    g = (g or "").strip().lower()
    if g in ("male", "m"):
        return "M"
    if g in ("female", "f"):
        return "F"
    return "O"


def clean_number(v):
    """Clean and convert value to number. Returns empty string if invalid."""
    try:
        if v is None:
            return ""
        v = str(v).strip()
        if not v:
            return ""
        if "%" in v:
            return float(v.replace("%", ""))
        return float(v)
    except:
        return v


def safe_text_id(v):
    """Convert scientific notation IDs (E+) to plain text."""
    if v is None:
        return ""
    v = str(v).strip()
    if "E+" in v.upper():
        try:
            return format(float(v), ".0f")
        except:
            return v
    return v


def get_first_present_field(row, *keys):
    """Return the first non-empty value for a list of possible CSV field names."""
    for key in keys:
        value = row.get(key)
        if value not in (None, ""):
            return value
    return ""


def is_last_semester_last_year(row):
    """Determine whether the student is in the final semester of the course."""
    total_years = clean_number(
        get_first_present_field(
            row,
            "Number of Course",
            "Number Of Course",
            "Total Number of Course",
            "Total Course Years",
            "Course Duration",
            "Total Years",
            "Course Duration Years"
        )
    )
    course_year = clean_number(
        get_first_present_field(
            row,
            "Course Year",
            "Current Course Year",
            "Year of Course",
            "Academic Year",
            "Year"
        )
    )
    semester = clean_number(get_first_present_field(row, "Semester Number", "SEM", "Semester"))

    if not isinstance(total_years, (int, float)):
        total_years = course_year

    if isinstance(total_years, (int, float)) and isinstance(semester, (int, float)):
        return semester == total_years * 2

    return False


def format_date(v, out_fmt="%m/%d/%Y", to_upper=False):
    """Parse common date formats and return formatted string or empty string."""
    if not v:
        return ""
    s = str(v).strip()
    fmts = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d %b %Y",
        "%d %B %Y",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y/%m/%d",
        "%d.%m.%Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d %b.%Y",
        "%d %b. %Y",
        "%d-%b-%y",
        "%d-%B-%y",
        "%d/%m/%y",
    ]
    for f in fmts:
        try:
            res = datetime.strptime(s, f).strftime(out_fmt)
            return res.upper() if to_upper else res
        except Exception:
            pass
    try:
        parts = s.replace('-', ' ').replace('/', ' ').replace('.', ' ').replace(',', ' ').split()
        if len(parts) >= 3 and parts[2].isdigit():
            day = parts[0]
            month = parts[1]
            year = parts[2]
            if len(year) == 2:
                year = "20" + year
            try:
                res = datetime.strptime(f"{day} {month} {year}", "%d %b %Y").strftime(out_fmt)
                return res.upper() if to_upper else res
            except Exception:
                try:
                    res = datetime.strptime(f"{day} {month} {year}", "%d %B %Y").strftime(out_fmt)
                    return res.upper() if to_upper else res
                except Exception:
                    pass
    except Exception:
        pass
    return s


def to_roman(n):
    """Convert small integer to Roman numerals (1..20)."""
    if not isinstance(n, int) or n <= 0:
        return ""
    vals = [
        (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
    ]
    res = ''
    x = n
    while x >= 10:
        res += 'X'
        x -= 10
    for val, sym in vals:
        while x >= val:
            res += sym
            x -= val
    return res


def parse_academic_course_id(s):
    """Extract short course code from complex Program string."""
    if not s:
        return ''
    s = str(s).strip()
    token = re.split(r'[\s\-_#$]+', s)[0]
    m = re.match(r'([A-Za-z]+)', token)
    if m:
        return m.group(1)
    return token


# =========================
# NAD TEMPLATE COLUMNS
# =========================

NAD_BASIC_HEADERS = [
    "ORG_NAME", "ORG_NAME_L", "ACADEMIC_COURSE_ID", "COURSE_NAME", "BATCH",
    "COURSE_NAME_L", "STREAM", "STREAM_L", "SESSION", "REGN_NO", "RROLL",
    "CNAME", "GENDER", "DOB", "FNAME", "MNAME", "PHOTO", "MRKS_REC_STATUS",
    "RESULT", "YEAR", "MONTH", "DIVISION", "GRADE", "PERCENT", "DOI", "SEM",
    "EXAM_TYPE", "TOT", "TOT_MRKS", "TOT_CREDIT", "TOT_CREDIT_POINTS",
    "TOT_GRADE_POINTS", "GRAND_TOT_MAX", "GRAND_TOT_MRKS",
    "GRAND_TOT_CREDIT_POINTS", "GRAND_TOT_CREDIT", "CGPA", "REMARKS", "SGPA",
    "ABC_ACCOUNT_ID", "TERM_TYPE", "TOT_GRADE", "GRAND_TOT_GRADE",
]

NAD_TRAILING_HEADERS = [
    "AADHAAR_NAME", "ADMISSION_YEAR", "NCRF_LEVEL", "EXAM_CENTRE",
]

MAX_SUBJECTS = 15


def build_full_headers():
    """Build the complete NAD column headers list in exact template order."""
    headers = list(NAD_BASIC_HEADERS)

    for i in range(1, MAX_SUBJECTS + 1):
        headers.extend([
            f"SUB{i}NM", f"SUB{i}",
            f"SUB{i}_TH_MAX", f"SUB{i}_PR_MAX", f"SUB{i}_CE_MAX",
            f"SUB{i}_TH_MRKS", f"SUB{i}_PR_MRKS", f"SUB{i}_CE_MRKS",
            f"SUB{i}_TOT",
            f"SUB{i}_GRADE", f"SUB{i}_GRADE_POINTS",
            f"SUB{i}_CREDIT", f"SUB{i}_CREDIT_POINTS",
            f"SUB{i}_REMARKS", f"SUB{i}_TOT_WRDS", f"SUB{i}_CREDIT_ELIGIBILITY",
        ])

    headers.extend(NAD_TRAILING_HEADERS)
    return headers


# =========================
# SUBJECT MAP
# =========================

def extract_subject_map(headers):
    """Extract subject column mapping from CSV headers."""
    subjects = OrderedDict()
    for col in headers:
        if " - Subject - Course Code" in col:
            sub = col.split(" - ")[0]
            subjects[sub] = {
                "name":          f"{sub} - Subject - Course Name",
                "code":          f"{sub} - Subject - Course Code",
                "ext_max":       f"{sub} - Subject Info - Ext Total Marks",
                "pr_max":        f"{sub} - Subject Info - Ext PR Total Marks",
                "internal_max":  f"{sub} - Subject Info - Internal Total Marks",
                "ext_marks":     f"{sub} - Result - Ext",
                "pr_marks":      f"{sub} - Result - Ext PR",
                "internal_marks": f"{sub} - Result - Internal",
                "total":         f"{sub} - Result - Total Marks",
                "grade":         f"{sub} - Result - Grade",
                "grade_points":  f"{sub} - Result - Grade Points",
                "credits":       f"{sub} - Result - Credits for Result",
                "result":        f"{sub} - Result - Result",
                "subject_gpa":   f"{sub} - Result - Subject GPA",
            }
    return OrderedDict(reversed(list(subjects.items())))


# =========================
# HARDCODED VALUES
# =========================

ORG_NAME = "ALL INDIA SHRI SHIVAJI MEMORIAL SOCIETY COLLEGE OF HOTEL MANAGEMENT AND CATERING TECHNOLOGY"
EXAM_CENTRE = "AISSMS College of Hotel Management and Catering Technology, Pune"
DEFAULT_SESSION = "2025-26"
DEFAULT_YEAR = "2025"
DEFAULT_MONTH = "December"


# =========================
# PROCESS FUNCTION
# =========================

def process_file(input_file):
    """Convert CSV to NAD format Excel. Returns (workbook_bytes, total_rows, error)."""
    from io import StringIO
    
    # Handle both file paths and file-like objects (Streamlit UploadedFile)
    if isinstance(input_file, str):
        # File path
        infile = None
        for encoding in ["utf-8", "utf-8-sig", "latin-1"]:
            try:
                infile = open(input_file, "r", encoding=encoding)
                infile.readline()
                infile.seek(0)
                break
            except UnicodeDecodeError:
                if infile:
                    infile.close()
                infile = None
        
        if infile is None:
            raise Exception(
                "Could not read file with any supported encoding "
                "(utf-8, utf-8-sig, latin-1)"
            )
    else:
        # File-like object (Streamlit UploadedFile)
        input_file.seek(0)
        infile = None
        for encoding in ["utf-8", "utf-8-sig", "latin-1"]:
            try:
                input_file.seek(0)
                content = input_file.read()
                if isinstance(content, bytes):
                    content = content.decode(encoding)
                infile = StringIO(content)
                break
            except (UnicodeDecodeError, AttributeError):
                pass
        
        if infile is None:
            raise Exception(
                "Could not read file with any supported encoding "
                "(utf-8, utf-8-sig, latin-1)"
            )

    try:
        reader = csv.DictReader(infile)
        subject_map = extract_subject_map(reader.fieldnames)

        wb = Workbook()
        ws = wb.active
        ws.title = "NAD Data"

        headers = build_full_headers()
        ws.append(headers)

        total_rows = 0
        all_processed_rows = []

        for row in reader:
            nad = {}

            # ========== HARDCODED VALUES ==========
            nad["ORG_NAME"] = ORG_NAME
            nad["ORG_NAME_L"] = ""
            nad["BATCH"] = ""
            nad["COURSE_NAME_L"] = ""
            nad["STREAM_L"] = ""
            nad["SESSION"] = DEFAULT_SESSION
            nad["PHOTO"] = ""
            nad["YEAR"] = DEFAULT_YEAR
            nad["MONTH"] = DEFAULT_MONTH
            nad["GRADE"] = ""
            nad["REMARKS"] = ""
            nad["TERM_TYPE"] = ""
            nad["TOT_GRADE"] = ""
            nad["GRAND_TOT_GRADE"] = ""
            nad["NCRF_LEVEL"] = ""
            nad["AADHAAR_NAME"] = ""
            nad["EXAM_CENTRE"] = EXAM_CENTRE

            # ========== DIRECT CSV MAPPINGS ==========
            raw_program = get_first_present_field(row, "Program", "Programs")
            if raw_program:
                course_name = raw_program.split("/")[-1].split("(")[0].strip()
            else:
                course_name = ""
            nad["ACADEMIC_COURSE_ID"] = parse_academic_course_id(raw_program)
            nad["COURSE_NAME"] = course_name

            raw_stream = get_first_present_field(row, "Pattern", "pattern", "Pattern Name", "Pattern Code", "Programs")
            stream_year_match = re.search(r"(\d{4})", str(raw_stream))
            if stream_year_match:
                nad["STREAM"] = f"{stream_year_match.group(1)} PATTERN"
            else:
                nad["STREAM"] = f"{str(raw_stream).strip()} PATTERN" if str(raw_stream).strip() else ""

            nad["REGN_NO"] = safe_text_id(row.get("Univ Roll Number"))
            nad["RROLL"] = safe_text_id(row.get("Seat Number"))
            nad["CNAME"] = row.get("Full Name", "")
            nad["GENDER"] = normalize_gender(row.get("Gender"))
            nad["DOB"] = format_date(row.get("Date of Birth", ""))
            nad["MNAME"] = row.get("Mother's First Name", "")
            nad["FNAME"] = row.get("Father's Full Name", "")
            abc_raw = safe_text_id(row.get("ABC Id"))
            if abc_raw:
                clean_abc = abc_raw.replace("-", "").replace(" ", "")
                if clean_abc.isdigit():
                    nad["ABC_ACCOUNT_ID"] = int(clean_abc)
                else:
                    nad["ABC_ACCOUNT_ID"] = abc_raw
            else:
                nad["ABC_ACCOUNT_ID"] = ""
            nad["ADMISSION_YEAR"] = row.get("Admission Year", "")

            sem_raw = get_first_present_field(row, "Semester Number", "SEM", "Semester")
            sem_num = None
            m = re.search(r"(\d+)", str(sem_raw or ""))
            if m:
                try:
                    sem_num = int(m.group(1))
                except Exception:
                    sem_num = None
            base_sem = to_roman(sem_num) if sem_num else (str(sem_raw).strip() if sem_raw else "")
            nad["SEM"] = base_sem
            nad["TERM_TYPE"] = f"{base_sem} Semester" if base_sem else ""

            if str(row.get("Grade Card Type") or "").strip().lower() == "fresh":
                nad["EXAM_TYPE"] = "REGULAR"
            else:
                nad["EXAM_TYPE"] = "BACKLOG"

            percent_raw = row.get("Semester Percentage")
            if percent_raw:
                percent_raw = str(percent_raw).replace("%", "").strip()
                nad["PERCENT"] = clean_number(percent_raw)
            else:
                nad["PERCENT"] = ""

            nad["DOI"] = format_date(row.get("Verification Date", ""), out_fmt="%d %B %Y", to_upper=True)

            nad["TOT"] = clean_number(row.get("Semester Out Of Marks"))
            nad["TOT_MRKS"] = clean_number(row.get("Semester Obtained Marks"))
            nad["TOT_CREDIT"] = clean_number(row.get("Completed Credits"))
            nad["TOT_CREDIT_POINTS"] = clean_number(
                row.get("Semester Credit Points")
            )
            nad["TOT_GRADE_POINTS"] = clean_number(
                row.get("Semester Credit Points")
            )

            nad["GRAND_TOT_MAX"] = clean_number(
                row.get("Cumulative Out Of Marks")
            )
            nad["GRAND_TOT_MRKS"] = clean_number(
                row.get("Cumulative Obtained Marks")
            )
            nad["GRAND_TOT_CREDIT_POINTS"] = clean_number(
                row.get("Cumulative Credit Points")
            )
            nad["GRAND_TOT_CREDIT"] = clean_number(
                row.get("Total Completed Credits")
            )

            final_semester = is_last_semester_last_year(row)

            nad["CGPA"] = (
                clean_number(row.get("CGPA")) if final_semester else ""
            )
            nad["SGPA"] = clean_number(row.get("SGPA"))

            # ========== RESULT & DIVISION LOGIC ==========
            sem_state = (
                row.get("Semester Passing State") or ""
            ).strip().lower()

            if sem_state == "sem_all_clear":
                nad["RESULT"] = "Pass"
                if final_semester:
                    division_raw = row.get("Semester Passing Class", "")
                    nad["DIVISION"] = division_raw if division_raw else "Pass"
                else:
                    nad["DIVISION"] = ""
            else:
                nad["RESULT"] = "Fail"
                nad["DIVISION"] = "Fail" if final_semester else ""

            # ========== MRKS_REC_STATUS LOGIC ==========
            grade_card = (
                row.get("Grade Card Type") or ""
            ).strip().lower()

            if "backlog" in grade_card:
                nad["MRKS_REC_STATUS"] = "M"
            else:
                nad["MRKS_REC_STATUS"] = "O"

            # ========== SUBJECTS ==========
            for idx, (sub, cols) in enumerate(subject_map.items(), start=1):
                if idx > MAX_SUBJECTS:
                    break

                nad[f"SUB{idx}NM"] = row.get(cols["name"], "")
                nad[f"SUB{idx}"] = row.get(cols["code"], "")

                ext_max = clean_number(row.get(cols["ext_max"]))
                pr_max = clean_number(row.get(cols["pr_max"]))
                internal_max = clean_number(row.get(cols["internal_max"]))

                nad[f"SUB{idx}_TH_MAX"] = ext_max if ext_max != "" else ""
                nad[f"SUB{idx}_PR_MAX"] = pr_max if pr_max != "" else ""
                nad[f"SUB{idx}_CE_MAX"] = (
                    internal_max if internal_max != "" else ""
                )

                nad[f"SUB{idx}_TH_MRKS"] = clean_number(
                    row.get(cols["ext_marks"])
                )
                nad[f"SUB{idx}_PR_MRKS"] = clean_number(
                    row.get(cols["pr_marks"])
                )
                nad[f"SUB{idx}_CE_MRKS"] = clean_number(
                    row.get(cols["internal_marks"])
                )
                nad[f"SUB{idx}_TOT"] = clean_number(
                    row.get(cols["total"])
                )

                nad[f"SUB{idx}_GRADE"] = row.get(cols["grade"], "")
                nad[f"SUB{idx}_GRADE_POINTS"] = clean_number(
                    row.get(cols["grade_points"])
                )
                nad[f"SUB{idx}_CREDIT"] = clean_number(
                    row.get(cols["credits"])
                )
                nad[f"SUB{idx}_CREDIT_POINTS"] = clean_number(
                    row.get(cols["subject_gpa"])
                )
                nad[f"SUB{idx}_REMARKS"] = row.get(cols["result"], "")
                nad[f"SUB{idx}_TOT_WRDS"] = ""
                nad[f"SUB{idx}_CREDIT_ELIGIBILITY"] = "Y"

            # ========== BUILD ROW VALUES ==========
            row_values = []

            for h in NAD_BASIC_HEADERS:
                row_values.append(nad.get(h, ""))

            for i in range(1, MAX_SUBJECTS + 1):
                row_values.extend([
                    nad.get(f"SUB{i}NM", ""),
                    nad.get(f"SUB{i}", ""),
                    nad.get(f"SUB{i}_TH_MAX", ""),
                    nad.get(f"SUB{i}_PR_MAX", ""),
                    nad.get(f"SUB{i}_CE_MAX", ""),
                    nad.get(f"SUB{i}_TH_MRKS", ""),
                    nad.get(f"SUB{i}_PR_MRKS", ""),
                    nad.get(f"SUB{i}_CE_MRKS", ""),
                    nad.get(f"SUB{i}_TOT", ""),
                    nad.get(f"SUB{i}_GRADE", ""),
                    nad.get(f"SUB{i}_GRADE_POINTS", ""),
                    nad.get(f"SUB{i}_CREDIT", ""),
                    nad.get(f"SUB{i}_CREDIT_POINTS", ""),
                    nad.get(f"SUB{i}_REMARKS", ""),
                    nad.get(f"SUB{i}_TOT_WRDS", ""),
                    nad.get(f"SUB{i}_CREDIT_ELIGIBILITY", ""),
                ])

            for h in NAD_TRAILING_HEADERS:
                row_values.append(nad.get(h, ""))

            all_processed_rows.append(row_values)
            total_rows += 1

        # Sort by RROLL
        rroll_idx = NAD_BASIC_HEADERS.index("RROLL")
        all_processed_rows.sort(
            key=lambda x: [
                int(c) if c.isdigit() else c.lower()
                for c in re.split(r'(\d+)', str(x[rroll_idx] or ""))
            ]
        )
        for r_vals in all_processed_rows:
            ws.append(r_vals)

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Save to bytes
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer, total_rows, None

    finally:
        infile.close()


# =========================
# STREAMLIT UI
# =========================

def render_brand():
    st.markdown("""
    <div class="brand-section">
        <div class="brand-icon">⚡</div>
        <div class="brand-title">NAD Converter v2</div>
        <div class="brand-subtitle">CSV → NAD Excel Format · New Template</div>
        <div class="brand-badge">
            <span class="dot"></span>
            Ready to Convert
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_upload_card():
    st.markdown("""
    <div class="card upload-card">
        <div class="card-header">
            <div class="card-icon upload">📂</div>
            <div>
                <div class="card-title">Upload CSV File</div>
                <div class="card-desc">Click anywhere in this box to select your CSV</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "",
        type=["csv"],
        label_visibility="collapsed",
        key="csv_uploader",
    )

    st.markdown("</div>", unsafe_allow_html=True)
    return uploaded_file


def render_settings_card():
    st.markdown("""
    <div class="card">
        <div class="card-header">
            <div class="card-icon settings">⚙️</div>
            <div>
                <div class="card-title">Configuration</div>
                <div class="card-desc">Customize output parameters (defaults pre-filled)</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        session = st.text_input("Session", value=DEFAULT_SESSION, key="session_input")
        year = st.text_input("Year", value=DEFAULT_YEAR, key="year_input")

    with col2:
        month = st.text_input("Month", value=DEFAULT_MONTH, key="month_input")
        exam_centre = st.text_input("Exam Centre", value=EXAM_CENTRE, key="centre_input")

    st.markdown("</div>", unsafe_allow_html=True)

    return {
        "session": session,
        "year": year,
        "month": month,
        "exam_centre": exam_centre,
    }


def render_convert_button(disabled=False):
    st.markdown("""
    <div class="convert-btn-wrapper">
    """, unsafe_allow_html=True)

    clicked = st.button(
        "⚡  Convert to NAD Excel",
        disabled=disabled,
        use_container_width=True,
        type="primary",
        key="convert_btn",
    )

    st.markdown("</div>", unsafe_allow_html=True)
    return clicked


def render_file_info(filename, filesize):
    size_kb = filesize / 1024
    if size_kb > 1024:
        size_str = f"{size_kb/1024:.1f} MB"
    else:
        size_str = f"{size_kb:.1f} KB"

    st.markdown(f"""
    <div class="file-info">
        <span class="file-info-icon">📄</span>
        <span class="file-info-name">{filename}</span>
        <span class="file-info-size">{size_str}</span>
        <span class="file-info-check">✓</span>
    </div>
    """, unsafe_allow_html=True)


def render_progress():
    st.markdown("""
    <div class="progress-container">
        <div class="progress-bar-bg">
            <div class="progress-bar-fill" style="width: 60%;"></div>
        </div>
        <div class="progress-text">Reading CSV · Processing rows · Generating Excel...</div>
    </div>
    """, unsafe_allow_html=True)


def render_success(total_rows, subjects_count, download_url):
    st.markdown(f"""
    <div class="card">
        <div class="card-header">
            <div class="card-icon result">✅</div>
            <div>
                <div class="card-title">Conversion Complete</div>
                <div class="card-desc">Your NAD format Excel file is ready to download</div>
            </div>
        </div>

        <div class="result-card">
            <div class="result-stats">
                <div class="stat-box">
                    <div class="stat-value">{total_rows}</div>
                    <div class="stat-label">Students</div>
                </div>
                <div class="stat-box">
                    <div class="stat-value">{subjects_count}</div>
                    <div class="stat-label">Subjects</div>
                </div>
                <div class="stat-box">
                    <div class="stat-value">{len(build_full_headers())}</div>
                    <div class="stat-label">Columns</div>
                </div>
            </div>

            <a href="{download_url}" download="nad_output.xlsx" class="download-btn">
                ⬇️  Download NAD Excel File
            </a>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_error(error_msg):
    st.markdown(f"""
    <div class="card">
        <div class="error-card">
            <div class="error-title">❌ Conversion Failed</div>
            <div class="error-msg">{error_msg}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_features():
    st.markdown("""
    <div class="features-strip">
        <div class="feature-item">
            <span class="fi-icon">📚</span>
            Up to 15 Subjects
        </div>
        <div class="feature-item">
            <span class="fi-icon">🔐</span>
            ABC ID Support
        </div>
        <div class="feature-item">
            <span class="fi-icon">📐</span>
            Roman Semester
        </div>
        <div class="feature-item">
            <span class="fi-icon">📊</span>
            CGPA / SGPA
        </div>
    </div>

    <div class="footer-text">
        NAD Converter v2 · AISSMS College of Hotel Management · Built with Streamlit
    </div>
    """, unsafe_allow_html=True)


# =========================
# MAIN APP
# =========================

def main():
    render_brand()

    uploaded_file = render_upload_card()

    if uploaded_file:
        render_file_info(uploaded_file.name, uploaded_file.size)

    settings = render_settings_card()
    clicked = render_convert_button(disabled=uploaded_file is None)

    if clicked and uploaded_file is not None:
        render_progress()
        st.markdown("""
        <div class="processing-text">
            <div class="spinner-icon">⚙️</div>
            <div class="label">Processing your file...</div>
            <div class="sublabel">This may take a moment for large files</div>
        </div>
        """, unsafe_allow_html=True)

        try:
            buffer, total_rows, error = process_file(uploaded_file)

            if error:
                render_error(error)
            else:
                # Count subjects from the file
                uploaded_file.seek(0)
                subj_count = 0
                try:
                    for encoding in ["utf-8", "utf-8-sig", "latin-1"]:
                        try:
                            uploaded_file.seek(0)
                            sample = uploaded_file.read(10000)
                            if isinstance(sample, bytes):
                                sample = sample.decode(encoding)
                            lines = sample.split("\n")
                            if len(lines) >= 1:
                                header_line = lines[0]
                                subj_count = header_line.count(" - Subject - Course Code")
                            break
                        except (UnicodeDecodeError, AttributeError):
                            continue
                except Exception:
                    subj_count = 0

                # Provide download
                st.download_button(
                    label="⬇️  Download NAD Excel File",
                    data=buffer,
                    file_name=f"nad_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key="download_btn",
                )

                # Render success stats below download
                st.markdown(f"""
                <div style="margin-top: 1.5rem;">
                    <div class="result-stats">
                        <div class="stat-box">
                            <div class="stat-value">{total_rows}</div>
                            <div class="stat-label">Students</div>
                        </div>
                        <div class="stat-box">
                            <div class="stat-value">{subj_count}</div>
                            <div class="stat-label">Subjects</div>
                        </div>
                        <div class="stat-box">
                            <div class="stat-value">{len(build_full_headers())}</div>
                            <div class="stat-label">Columns</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        except Exception as e:
            render_error(str(e))

    render_features()


if __name__ == "__main__":
    main()