# Google Drive Folder Extractor

## Project Overview
This project is designed to extract the essential applicant information—such as names, phone numbers, and emails—from resumes submitted in **PDF** and **DOCX** formats. These resumes are stored in a centralized folder, typically containing multiple subfolders categorized by group and position.

---

## Goal
- Eliminate manual resume scanning and data entry by HR personnel.
- Automatically collect and clean key candidate details from each resume.
- Organize the extracted information into a structured format and update it in a **Google Sheet** for easy access, review, and follow-up by the recruitment team.

---

## How It Works
1.  This program, written in Python, scans specified folder and all their subfolders for PDF and DOCX files.
2. It extracts text content from each file and uses **Natural Language Processing (NLP)** and **regular expressions** to:
   - Identify candidate names using **spaCy’s Named Entity Recognition (NER)**.
   - Extract **email addresses** and **phone numbers**.
3. The data is then **cleaned and formatted** (e.g., standardizing phone numbers and removing duplicate entries).
4. All results, including files with no extractable data (marked as `"N/A"`), are:
   - **Saved as a CSV backup in Google Drive**.
   - **Automatically written to a Google Sheet** for easy access.

---

## Impact
- **Time-saving:** Reduces hours of manual data entry.
- **Consistency:** Ensures uniform formatting and accurate extraction across all applicants.
- **Accessibility:** Makes applicant data instantly available in a centralized Google Sheet.
- **Transparency:** Includes even files with no readable data for tracking and troubleshooting.
