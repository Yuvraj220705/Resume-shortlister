import pandas as pd
from sklearn.preprocessing import MinMaxScaler
import os
import fitz  # PyMuPDF library for reading PDFs
import docx  # python-docx library for reading DOCX files
import re
from typing import List, Dict
from datetime import datetime # Add this line

# --- Part 1: File Reading and Text Extraction ---

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extracts text from a PDF file using PyMuPDF."""
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return text

def extract_text_from_docx(docx_path: str) -> str:
    """Extracts text from a DOCX file using python-docx."""
    text = ""
    try:
        doc = docx.Document(docx_path)
        for para in doc.paragraphs:
            text += para.text + '\n'
    except Exception as e:
        print(f"Error reading DOCX {docx_path}: {e}")
    return text

def extract_text_from_resume(file_path: str) -> str:
    """Determines file type and extracts text."""
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.endswith('.docx'):
        return extract_text_from_docx(file_path)
    return ""

IMPORTANT_CERTS = [
    "AWS Certified", "Amazon Web Services", "Azure", "Microsoft Certified", 
    "GCP", "Google Cloud", "PMP", "Project Management Professional",
    "Certified ScrumMaster", "CSM", "Oracle Certified"
]
IMPORTANT_SKILLS = [
    "Python", "Java", "C++", "JavaScript", "SQL", "NoSQL", 
    "AI", "Artificial Intelligence", "ML", "Machine Learning", 
    "Data Science", "Data Analysis", "Tableau", "Power BI",
    "Cloud Computing", "AWS", "Azure", "GCP",
    "Communication", "Teamwork", "Collaboration", "Leadership",
    "Problem-solving"
]

def parse_resume(resume_text: str, name: str) -> Dict:
    total_years_exp = 0.0
    date_range_pattern = re.compile(
        r'([A-Z][a-z]{2,}\s+\d{4}|\d{4}|present|current)\s*[-\u2013]\s*([A-Z][a-z]{2,}\s+\d{4}|\d{4}|present|current)',
        re.IGNORECASE
    )
    date_ranges = re.findall(date_range_pattern, resume_text)
    print(f"DEBUG: Found date ranges for {name}: {date_ranges}")
    for start_date_str, end_date_str in date_ranges:
        total_years_exp += calculate_experience_from_dates(start_date_str, end_date_str)
    print(f"DEBUG: Total experience calculated for {name}: {total_years_exp:.2f} years")
    certs_found = [cert for cert in IMPORTANT_CERTS if re.search(r'\b' + re.escape(cert) + r'\b', resume_text, re.IGNORECASE)]
    skills_found = [skill for skill in IMPORTANT_SKILLS if re.search(r'\b' + re.escape(skill) + r'\b', resume_text, re.IGNORECASE)]
    achievements_regex = re.compile(r'(?:[•*-]|\d+\.)\s+(.+)', re.MULTILINE)
    achievements_count = len(re.findall(achievements_regex, resume_text))
    return {
        'Name': name,
        'Experience (Years)': round(total_years_exp, 2),
        'Certifications': ", ".join(sorted(list(set(certs_found)))),
        'Achievements': achievements_count,
        'Skills': ", ".join(sorted(list(set(skills_found))))
    }

def calculate_experience_from_dates(start_date_str: str, end_date_str: str) -> float:
    month_map = {
        'january': 1, 'jan': 1, 'february': 2, 'feb': 2, 'march': 3, 'mar': 3,
        'april': 4, 'apr': 4, 'may': 5, 'june': 6, 'jun': 6, 'july': 7, 'jul': 7,
        'august': 8, 'aug': 8, 'september': 9, 'sep': 9, 'october': 10, 'oct': 10,
        'november': 11, 'nov': 11, 'december': 12, 'dec': 12
    }
    def parse_date(date_str: str):
        date_str = date_str.lower().strip()
        if date_str in ['current', 'present']:
            return datetime.now()
        parts = date_str.split()
        if len(parts) == 2 and parts[0] in month_map:
            month = month_map[parts[0]]
            year = int(parts[1])
            return datetime(year, month, 1)
        return None
    try:
        start_date = parse_date(start_date_str)
        end_date = parse_date(end_date_str)
        if start_date and end_date:
            duration = end_date - start_date
            return duration.days / 365.25
    except Exception as e:
        print(f"DEBUG: Error in calculating duration: {e}")
    return 0.0

def count_certifications(cert_text: str) -> int:
    certs = str(cert_text).split(", ")
    return sum(1 for cert in certs if cert.strip().lower() in [c.lower() for c in IMPORTANT_CERTS])

def count_skills(skill_text: str) -> int:
    skills = str(skill_text).split(", ")
    return sum(1 for skill in skills if skill.strip().lower() in [s.lower() for s in IMPORTANT_SKILLS])

def shortlist_candidates(resume_folder: str, top_n: int = 10) -> pd.DataFrame:
    parsed_candidates_data = []
    if not os.path.exists(resume_folder):
        print(f"Error: The folder '{resume_folder}' was not found. Please create it and add your resume files.")
        return pd.DataFrame()
    for filename in os.listdir(resume_folder):
        file_path = os.path.join(resume_folder, filename)
        if os.path.isdir(file_path):
            continue
        resume_text = extract_text_from_resume(file_path)
        if resume_text:
            candidate_name = os.path.splitext(filename)[0]
            parsed_data = parse_resume(resume_text, candidate_name)
            parsed_candidates_data.append(parsed_data)
            print(f"Successfully parsed: {filename}")
        else:
            print(f"Skipped parsing: {filename} (could not extract text or unsupported file type)")
    if not parsed_candidates_data:
        print("No resumes were parsed. Please check your 'resumes' folder for valid files.")
        return pd.DataFrame()
    data = pd.DataFrame(parsed_candidates_data)
    data['Certifications_Count'] = data['Certifications'].apply(count_certifications)
    data['Skills_Count'] = data['Skills'].apply(count_skills)
    features = ['Experience (Years)', 'Achievements', 'Certifications_Count', 'Skills_Count']
    for feature in features:
        if data[feature].nunique() == 1:
            data[feature] = 0.0
    scaler = MinMaxScaler()
    data[features] = scaler.fit_transform(data[features])
    weights = {
        'Experience (Years)': 0.30,
        'Achievements': 0.30,
        'Certifications_Count': 0.20,
        'Skills_Count': 0.20
    }
    data['Score'] = (
        data['Experience (Years)'] * weights['Experience (Years)'] +
        data['Achievements'] * weights['Achievements'] +
        data['Certifications_Count'] * weights['Certifications_Count'] +
        data['Skills_Count'] * weights['Skills_Count']
    )
    data_sorted = data.sort_values(by='Score', ascending=False)
    print("\n--- Final Shortlisted Candidates ---")
    return data_sorted.head(top_n)

if __name__ == "__main__":
    resume_folder_path = "resumes"
    shortlisted_df = shortlist_candidates(resume_folder_path, top_n=10)
    if not shortlisted_df.empty:
        print(shortlisted_df[['Name', 'Score', 'Certifications', 'Skills', 'Experience (Years)']])
        shortlisted_df.to_csv("shortlisted_candidates_from_resumes.csv", index=False)
        print("\nShortlisted candidates saved to 'shortlisted_candidates_from_resumes.csv'")
# Required packages: pandas, scikit-learn, python-docx, pymupdf
