import os
import re
import pandas as pd
from datetime import datetime
import pdfplumber
import fitz
from docx import Document
import json
import spacy

print("Loading spaCy model...")
nlp = spacy.load("en_core_web_sm")

print("ATS_PARSER2 LOADED")

# ----------------------------
# SAFE TEXT EXTRACTION
# ----------------------------
def extract_text(file_path):
    text = ""

    try:
        if file_path.lower().endswith(".pdf"):

            try:
                with pdfplumber.open(file_path) as pdf:
                    text = "\n".join(
                        page.extract_text() or "" for page in pdf.pages
                    )
            except Exception:
                text = ""

            if not text.strip():
                doc = fitz.open(file_path)
                text = "\n".join(page.get_text() for page in doc)

        elif file_path.lower().endswith(".docx"):
            doc = Document(file_path)
            text = "\n".join(p.text for p in doc.paragraphs)

    except Exception as e:
        print(f"[EXTRACTION ERROR] {file_path}: {e}")

    return text


# ----------------------------
# RESUME VALIDATION
# ----------------------------
def is_resume(text, filename=""):
    text_lower = text.lower()
    filename_lower = filename.lower()

    reject_keywords = [
        "certificate", "marks card", "transcript",
        "aadhar", "pan card", "salary slip"
    ]

    if any(word in text_lower for word in reject_keywords):
        return False

    score = 0

    if any(x in filename_lower for x in ["resume", "cv"]):
        score += 3

    if re.search(r"\bexperience\b|\bskills\b|\beducation\b", text_lower):
        score += 2

    if re.search(r"[\w\.-]+@[\w\.-]+", text):
        score += 2

    if re.search(r"\+?\d[\d\s\-]{8,15}", text):
        score += 2

    return score >= 5


# ----------------------------
# FIELD EXTRACTION
# ----------------------------
def extract_email(text):
    match = re.findall(r"[\w\.-]+@[\w\.-]+", text)
    return match[0] if match else ""


def extract_phone(text):
    match = re.findall(r"\+?\d[\d\s\-]{8,15}", text)
    return match[0] if match else ""


def extract_name(filename):
    return os.path.basename(filename).split("_")[0]


def extract_experience(text):
    match = re.findall(r"(\d+)\s*(?:years|yrs|year)", text.lower())
    return max(map(int, match)) if match else ""

def extract_education(text):

    degree_patterns = [
        r"\bB\.?\s?TECH\b",
        r"\bM\.?\s?TECH\b",
        r"\bB\.?\s?E\b",
        r"\bM\.?\s?E\b",
        r"\bB\.?\s?SC\b",
        r"\bM\.?\s?SC\b",
        r"\bMBA\b",
        r"\bPH\.?\s?D\b",
        r"\bDIPLOMA\b"
    ]

    degrees = []

    for pattern in degree_patterns:

        matches = re.findall(
            pattern,
            text,
            re.IGNORECASE
        )

        degrees.extend(matches)

    degrees = list(
        dict.fromkeys(
            [d.upper() for d in degrees]
        )
    )

    return ", ".join(degrees)

def extract_location(text):

    cities = [
        "Bangalore",
        "Bengaluru",
        "Hyderabad",
        "Chennai",
        "Pune",
        "Mumbai",
        "Delhi",
        "Noida",
        "Gurgaon",
        "Mysore",
        "Kolkata",
        "Ahmedabad",
        "Kochi",
        "Coimbatore",
        "Trivandrum"
    ]

    for city in cities:

        if re.search(
                rf"\b{city}\b",
                text,
                re.IGNORECASE):

            return city

    return ""


def extract_companies(text):

    doc = nlp(text)

    companies = []

    blacklist = {
        "rf systems",
        "communication systems",
        "embedded systems",
        "technical skills",
        "education",
        "experience",
        "matlab",
        "python",
        "tensorflow",
        "pytorch"
    }

    for ent in doc.ents:

        if ent.label_ == "ORG":

            org = ent.text.strip()

            if len(org) < 3:
                continue

            if org.lower() in blacklist:
                continue

            # Ignore educational institutions if desired
            if any(x in org.lower() for x in [
                "college",
                "university",
                "institute",
                "school"
            ]):
                continue

            companies.append(org)

    companies = list(dict.fromkeys(companies))

    return (companies + ["", ""])[:2]

def extract_competency(text):
    text_lower = text.lower()

    competency_map = {
        "RF Design": ["rf design", "antenna", "microwave"],
        "LNA": ["lna", "low noise amplifier"],
        "PA": ["power amplifier", "pa design"],
        "Receiver": ["receiver", "rx chain"],
        "Transmitter": ["transmitter", "tx chain"],
        "RF Systems": ["rf system", "radar", "satcom"],
        "RF Testing": ["vna", "spectrum analyzer", "rf testing"]
    }

    result = {}

    for competency, keywords in competency_map.items():

        result[competency] = "No"

        for keyword in keywords:

            if keyword in text_lower:
                result[competency] = "Yes"
                break

    return result


# ----------------------------
# MAIN PIPELINE
# ----------------------------
def process_resumes(folder_path, output_excel=None):

    if output_excel is None:
        output_excel = (
            f"ATS_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

    # -----------------------------------------
    # Load Outlook metadata
    # -----------------------------------------
    metadata = {}

    meta_file = os.path.join(
        folder_path,
        "resume_metadata.json"
    )

    if os.path.exists(meta_file):
        with open(meta_file,"r",encoding="utf-8") as f:
            metadata = json.load(f)

    data = []
    seen_emails = set()
    seen_files = set()

    sl_no = 1

    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if not file.lower().endswith((".pdf", ".docx")):
            continue

        key = (file.lower(), os.path.getsize(file_path))
        if key in seen_files:
            continue
        seen_files.add(key)

        text = extract_text(file_path)
        if len(text.split()) < 80:
            continue

        if not is_resume(text, file):
            continue

        email = extract_email(text)
        if email and email in seen_emails:
            continue
        if email:
            seen_emails.add(email)

        comp = extract_competency(text)

        companies = extract_companies(text)

        email_date = metadata.get(
            file,
            {}
        ).get(
            "email_date",
            ""
        )

        education = extract_education(text)

        location = extract_location(text)

        data.append([
            sl_no,
            extract_name(file),
            email,
            extract_phone(text),
            email_date,
            location,
            education,
            extract_experience(text),
            companies[0],
            companies[1],
            comp["RF Design"],
            comp["LNA"],
            comp["PA"],
            comp["Receiver"],
            comp["Transmitter"],
            comp["RF Systems"],
            comp["RF Testing"]
        ])

        sl_no += 1

    df = pd.DataFrame(data, columns=[
    "Sl.No",
    "Name",
    "Email",
    "Phone",
    "Email Date",
    "Location",
    "Education",
    "Experience",
    "Company 1",
    "Company 2",
    "RF Design",
    "LNA",
    "Power Amplifier / PA",
    "Receiver",
    "Transmitter",
    "RF Systems",
    "RF Testing"
])

    df.to_excel(output_excel, index=False)

    return output_excel