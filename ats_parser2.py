import os
import re
import pandas as pd
import pdfplumber
from docx import Document
from datetime import datetime
import hashlib
import fitz  # PyMuPDF

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

            # fallback if pdfplumber fails or empty
            if not text.strip():
                doc = fitz.open(file_path)
                text = "\n".join(page.get_text() for page in doc)

        elif file_path.lower().endswith(".docx"):
            doc = Document(file_path)
            text = "\n".join([para.text for para in doc.paragraphs])

    except Exception as e:
        print(f"[EXTRACTION ERROR] {file_path}: {e}")

    return text

def is_resume(text, filename=""):
    """
    Detect whether document is likely a resume.
    """

    text_lower = text.lower()
    filename_lower = filename.lower()

    # Immediate reject keywords
    reject_keywords = [
        "certificate of completion",
        "certificate awarded",
        "participation certificate",
        "certificate of appreciation",
        "training certificate",
        "workshop certificate",
        "successfully completed",
        "mark sheet",
        "marks card",
        "transcript",
        "grade sheet",
        "offer letter",
        "appointment letter",
        "relieving letter",
        "experience certificate",
        "salary slip",
        "payslip",
        "aadhaar",
        "aadhar",
        "pan card",
        "passport"
    ]

    for word in reject_keywords:
        if word in text_lower:
            return False

    score = 0

    # Resume filename bonus
    resume_filename_terms = [
        "resume",
        "cv",
        "profile",
        "curriculum vitae"
    ]

    if any(term in filename_lower for term in resume_filename_terms):
        score += 3

    resume_sections = [
        "experience",
        "work experience",
        "professional summary",
        "career objective",
        "education",
        "skills",
        "technical skills",
        "projects",
        "internship",
        "employment history",
        "responsibilities"
    ]

    for section in resume_sections:
        if section in text_lower:
            score += 1

    has_email = bool(
        re.search(
            r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
            text
        )
    )

    has_phone = bool(
        re.search(
            r"\+?\d[\d\s\-]{8,15}",
            text
        )
    )

    if has_email:
        score += 2

    if has_phone:
        score += 2

    return score >= 5

# ----------------------------
# FIELD EXTRACTION HELPERS
# ----------------------------
def extract_email(text):
    match = re.findall(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", text)
    return match[0] if match else ""


def extract_phone(text):
    match = re.findall(r"\+?\d[\d\s\-]{8,15}", text)
    return match[0] if match else ""


def extract_name(filename, text):
    base = os.path.basename(filename)
    name = base.split("_")[0]
    return name.replace(".", " ").strip()


def extract_experience(text):
    match = re.findall(r"(\d+)\+?\s*(years|yrs|year)", text.lower())
    if match:
        return max([int(m[0]) for m in match])
    return ""


def extract_education(text):
    keywords = ["b.tech", "m.tech", "b.e", "m.e", "bsc", "msc", "mba", "phd", "diploma"]
    found = [k for k in keywords if k in text.lower()]
    return ", ".join(found)


def extract_skills(text):
    skill_bank = [
        "python", "java", "c++", "sql", "machine learning",
        "deep learning", "opencv", "tensorflow", "pytorch",
        "embedded", "linux", "aws", "docker"
    ]
    found = [s for s in skill_bank if s in text.lower()]
    return ", ".join(found)

def extract_competency(text):

    text_lower = text.lower()

    competency_map = {
        "RF Design": [
            "rf design",
            "radio frequency design",
            "rf circuit design",
            "microwave design",
            "antenna design"
        ],

        "LNA": [
            "lna",
            "low noise amplifier"
        ],

        "Power Amplifier / PA": [
            "power amplifier",
            "pa design",
            "rf pa",
            "rf power amplifier"
        ],

        "Receiver": [
            "receiver",
            "rx chain",
            "rf receiver",
            "receiver design"
        ],

        "Transmitter": [
            "transmitter",
            "tx chain",
            "rf transmitter",
            "transmitter design"
        ],

        "RF Systems": [
            "rf systems",
            "wireless systems",
            "communication systems",
            "radio systems",
            "satcom",
            "radar system"
        ],

        "RF Testing": [
            "rf testing",
            "rf measurements",
            "network analyzer",
            "vector network analyzer",
            "vna",
            "spectrum analyzer",
            "signal generator",
            "emi",
            "emc",
            "validation testing",
            "verification testing"
        ]
    }

    found = []

    for competency, keywords in competency_map.items():

        for keyword in keywords:

            if keyword in text_lower:
                found.append(competency)
                break

    return ", ".join(found)

def extract_companies(text):
    matches = re.findall(r"(?:at|@)\s([A-Z][A-Za-z0-9& ]{2,})", text)
    unique = list(dict.fromkeys(matches))
    if len(unique) >= 2:
        return unique[:2]
    return unique + [""] * (2 - len(unique))


def extract_location(text):
    match = re.search(
        r"(Bangalore|Bengaluru|Hyderabad|Chennai|Pune|Delhi|Mumbai)",
        text,
        re.IGNORECASE
    )
    return match.group(0) if match else ""


def extract_designation(text):
    match = re.findall(
        r"(software engineer|developer|data scientist|analyst|manager|consultant)",
        text.lower()
    )
    return match[0] if match else ""

def filename_looks_like_resume(filename):

    filename = filename.lower()

    resume_terms = [
        "resume",
        "cv",
        "profile",
        "curriculum vitae"
    ]

    return any(term in filename for term in resume_terms)

seen_files = set()
# ----------------------------
# MAIN ATS PIPELINE
# ----------------------------
def process_resumes(folder_path, output_excel=None):

    date_str = datetime.now().strftime("%Y-%m-%d")
    if output_excel is None:
        output_excel = f"Resume_categorization_{date_str}.xlsx"

    data = []
    sl_no = 1

    seen_hashes = set()  # 🔴 IMPORTANT: prevents duplicates
    seen_emails = set()

    for file in os.listdir(folder_path):

        file_path = os.path.join(folder_path, file)
        file_size = os.path.getsize(file_path)

        duplicate_key = (file.lower(), file_size)

        if duplicate_key in seen_files:
            print(f"[DUPLICATE FILE SKIPPED] {file}")
            continue

        seen_files.add(duplicate_key)

        if not file.lower().endswith((".pdf", ".docx")):
            continue

        # ----------------------------
        # EXTRACT TEXT
        # ----------------------------
        text = extract_text(file_path)

        if not text.strip():
            print(f"[EMPTY FILE] {file}")
            continue

        # ----------------------------
        # MINIMUM CONTENT CHECK
        # ----------------------------
        word_count = len(text.split())

        if word_count < 100:
            print(f"[TOO SMALL TO BE RESUME] {file}")
            continue

        # ----------------------------
        # RESUME VALIDATION
        # ----------------------------
        if not is_resume(text, file):
            print(f"[NOT A RESUME] {file}")
            continue

        # ----------------------------
        # FIELD EXTRACTION
        # ----------------------------
        email = extract_email(text)
        if email and email.lower() in seen_emails:
            print(f"[DUPLICATE EMAIL SKIPPED] {file}")
            continue

        seen_emails.add(email.lower())

        phone = extract_phone(text)
        name = extract_name(file, text)
        experience = extract_experience(text)
        education = extract_education(text)
        skills = extract_skills(text)
        competency = extract_competency(text)
        location = extract_location(text)
        designation = extract_designation(text)

        companies = extract_companies(text)
        company_1 = companies[0] if len(companies) > 0 else ""
        company_2 = companies[1] if len(companies) > 1 else ""

        # ----------------------------
        # EMAIL DATE PARSE
        # ----------------------------
        email_date = ""
        date_match = re.search(r"(\d{8}_\d{6})", file)
        if date_match:
            try:
                email_date = datetime.strptime(
                    date_match.group(), "%Y%m%d_%H%M%S"
                )
            except:
                email_date = ""

        # ----------------------------
        # STORE ROW
        # ----------------------------
        data.append([
            sl_no,
            designation,
            name,
            phone,
            email,
            email_date,
            location,
            education,
            experience,
            company_1,
            company_2,
            designation,
            skills,
            competency
        ])

        sl_no += 1

    # ----------------------------
    # EXPORT TO EXCEL
    # ----------------------------
    df = pd.DataFrame(data, columns=[
        "Sl.No",
        "Position",
        "Name",
        "Contact Number",
        "Email ID",
        "Date of Email",
        "Location",
        "Education",
        "Total Experience",
        "Company 1",
        "Company 2",
        "Designation",
        "Skills",
        "Competency"
    ])

    df.to_excel(output_excel, index=False)

    print(f"\n[DONE] File saved: {output_excel}")
    return output_excel


# ----------------------------
# RUN DIRECTLY
# ----------------------------
if __name__ == "__main__":
    folder = input("Enter resume folder path: ").strip()
    process_resumes(folder)