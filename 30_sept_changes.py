import numpy as np
import pandas as pd
from sklearn.preprocessing import QuantileTransformer
from scipy.stats import rankdata
import os
import fitz
import docx  # python-docx library for reading DOCX files
import re
from fuzzywuzzy import fuzz
from typing import List, Dict
from datetime import datetime

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
    "Cloud Computing", "AWS", "Azure", "GCP", "Infrastructure as Code", "Serverless", "EC2", "Lambda", "Azure Functions",
    "Communication", "Teamwork", "Collaboration", "Leadership",
    "Problem-solving"
] 

def parse_resume(resume_text: str, name: str) -> Dict:
    total_years_exp = 0.0
    date_range_pattern = re.compile(
        r'([A-Za-z]{3,}\s*\d{4}|\d{4}|present|current)\s*[-–—]\s*([A-Za-z]{3,}\s*\d{4}|\d{4}|present|current)',
        re.IGNORECASE
    )
    date_ranges = re.findall(date_range_pattern, resume_text)
    print(f"DEBUG: Found date ranges for {name}: {date_ranges}")

    for start_date_str, end_date_str in date_ranges:
        total_years_exp += calculate_experience_from_dates(start_date_str, end_date_str)

    print(f"DEBUG: Total experience calculated for {name}: {total_years_exp:.2f} years")
    # certs_found = [cert for cert in IMPORTANT_CERTS if re.search(r'\b' + re.escape(cert) + r'\b', resume_text, re.IGNORECASE)]
    certs_found = find_fuzzy_matches(resume_text, IMPORTANT_CERTS)

    # skills_found = [skill for skill in IMPORTANT_SKILLS if re.search(r'\b' + re.escape(skill) + r'\b', resume_text, re.IGNORECASE)]
    skills_found = find_fuzzy_matches(resume_text, IMPORTANT_SKILLS)
    achievements_regex = re.compile(r'(?:[•*-]|\d+\.)\s+(.+)', re.MULTILINE)
    achievements_count = len(re.findall(achievements_regex, resume_text))

    # --- Extract Contact Information ---
    email_regex = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    phone_regex = re.compile(r'(\+?\d{1,2}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}')

    email_match = email_regex.search(resume_text)
    phone_match = phone_regex.search(resume_text)

    email = email_match.group(0) if email_match else "Not Found"
    phone = phone_match.group(0) if phone_match else "Not Found"

    return {
        'Name': name,
        'Email': email,
        'Phone': phone,
        'Experience (Years)': round(total_years_exp, 2),
        'Certifications': ", ".join(sorted(list(set(certs_found)))),
        'Achievements': achievements_count,
        'Skills': ", ".join(sorted(list(set(skills_found))))
    }

def find_fuzzy_matches(resume_text: str, items: List[str], threshold: int = 80) -> List[str]:
    """
    Finds fuzzy matches of items in resume_text with a given threshold.
    """
    found_items = []
    for item in items:
        # Use regex to find whole words or phrases
        pattern = r'\b' + re.escape(item) + r'\b'
        if re.search(pattern, resume_text, re.IGNORECASE):
            found_items.append(item)
        else:
            if fuzz.partial_ratio(resume_text.lower(), item.lower()) >= threshold:
                found_items.append(item)
    return found_items
def calculate_experience_from_dates(start_date_str: str, end_date_str: str) -> float:
    """Calculates the experience in years between two date strings."""
    month_map = {
        'january': 1, 'jan': 1, 'february': 2, 'feb': 2, 'march': 3, 'mar': 3,
        'april': 4, 'apr': 4, 'may': 5, 'june': 6, 'jun': 6, 'july': 7, 'jul': 7,
        'august': 8, 'aug': 8, 'september': 9, 'sep': 9, 'october': 10, 'oct': 10,
        'november': 11, 'nov': 11, 'december': 12, 'dec': 12,
        'sept': 9,  # Corrected September abbreviation
        'sept.': 9  # September with a dot
    }

    def parse_date(date_str: str):
        """Parses a date string and returns a datetime object."""
        date_str = date_str.lower().strip()
        if date_str in ['current', 'present']:
            return datetime.now()

        parts = date_str.split()
        if len(parts) == 2 and parts[0] in month_map:
            month = month_map[parts[0]]
            year = int(parts[1])
            return datetime(year, month, 1)
        if len(parts) == 1 and parts[0].isdigit():
             year = int(parts[0])
             return datetime(year, 1, 1)

        return None

    try:
        start_date = parse_date(start_date_str)
        end_date = parse_date(end_date_str)
        if start_date and end_date:
            duration = end_date - start_date
            years = duration.days / 365.25
            return years
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

    if len(parsed_candidates_data) < 2:
        print("\nError: At least two resumes are required for scoring and comparison.")
        return pd.DataFrame()

    data = pd.DataFrame(parsed_candidates_data)
    data['Certifications_Count'] = data['Certifications'].apply(count_certifications)
    data['Skills_Count'] = data['Skills'].apply(count_skills)

    # --- Scoring System ---
    features_to_score = ['Experience (Years)', 'Achievements', 'Certifications_Count', 'Skills_Count']
    
    # Use rank-based scoring to handle ties and provide a better score distribution.
    for feat in features_to_score:
        score_col = f'{feat} Score'
        # 'dense' ranking ensures candidates with the same value get the same rank
        ranks = rankdata(data[feat], method='dense')
        # Scale ranks to a 1-100 score
        max_rank = ranks.max()
        if max_rank > 1:
            data[score_col] = ((ranks - 1) / (max_rank - 1)) * 99 + 1
        else: # Handle case where all values are the same
            data[score_col] = 100
        # Use percentile-based ranking for a smoother score distribution
        # 'min' method ensures candidates with the same value get the same (lowest) percentile in their group
        data[score_col] = data[feat].rank(method='min', pct=True) * 100

    # Define weights for each score
    weights = {
        'Experience (Years) Score': 0.30,
        'Achievements Score': 0.30,
        'Certifications_Count Score': 0.20,
        'Skills_Count Score': 0.20
    }

    # Calculate the final 'Overall Score'
    data['Overall Score'] = (
        data['Experience (Years) Score'] * weights['Experience (Years) Score'] +
        data['Achievements Score'] * weights['Achievements Score'] +
        data['Certifications_Count Score'] * weights['Certifications_Count Score'] +
        data['Skills_Count Score'] * weights['Skills_Count Score']
    )

    data_sorted = data.sort_values(by='Overall Score', ascending=False)

    print("\n--- Final Shortlisted Candidates ---")
    return data_sorted.head(top_n)

if __name__ == "__main__":
    resume_folder_path = "resumes"
    shortlisted_df = shortlist_candidates(resume_folder_path, top_n=10) #setting top_n =10
    if not shortlisted_df.empty:
        output_columns = [
            'Name', 
            'Email',
            'Phone',
            'Overall Score',
            'Experience (Years)', 
            'Achievements', 
            'Skills_Count', 
            'Certifications_Count',
            'Skills',
            'Certifications'
        ]
        display_columns = [col for col in output_columns if col in shortlisted_df.columns]
        
        print(shortlisted_df[display_columns])
        shortlisted_df.to_csv("shortlisted_candidates_from_resumes.csv", index=False)
        print("\nShortlisted candidates saved to 'shortlisted_candidates_from_resumes.csv'")

# Required packages: pandas, scikit-learn, python-docx, pymupdf