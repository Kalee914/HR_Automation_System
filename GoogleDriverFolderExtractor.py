# Install necessary libraries
!pip install python-docx
!pip install PyPDF2
!pip install spacy
!python -m spacy download en_core_web_sm
!pip install pdfplumber

from google.colab import drive
drive.mount('/content/drive')

# confirming the path with the JSON file
!ls "/content/drive/MyDrive/PATH_TO_JSON_FILE"
# Confirming Folder Path 
!ls "/content/drive/MyDrive/FOLDER_PATH"

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Path to your JSON credentials file
credentials_file = "/content/drive/MyDrive/PATH_TO_JSON_FILE"

# Define the scope of access
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Authorize using the credentials file
creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
client = gspread.authorize(creds)

# Set the path to the folder dynamically based on the folder name
folder_name = 'FOLDER_NAME'
folder_path = f'/content/drive/MyDrive/{folder_name}/'

# open sheet, set sheetname, and insert sheet
spreadsheet = client.open_by_url('GOOGLESHEET_URL')
sheet_name = 'NAME_SHEET'
worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="50")

import os
import re
#import PyPDF2
import docx
import pandas as pd
import spacy
import logging
logging.getLogger("pdfminer").setLevel(logging.ERROR) #suppress warning 
import pdfplumber #using pdfplumber instead of PyPDF2

# Load spaCy's English NLP model
nlp = spacy.load("en_core_web_sm")

# Function to extract emails and phone numbers from a text string
def extract_emails_and_phone_numbers(text):
    emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    phones = re.findall(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
    return emails, phones

# Function to extract names using spaCy's NER
def extract_names(text):
    doc = nlp(text)
    names = [ent.text for ent in doc.ents if ent.label_ == 'PERSON']
    return names

# Function to extract text from PDF
'''
def extract_text_from_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            text = "".join([page.extract_text() for page in reader.pages if page.extract_text()])
        return text if text else "N/A"
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return "N/A"
'''

# Function to extract text from PDF (with pdfplumber)
def extract_text_from_pdf(file_path):
    try:
        with pdfplumber.open(file_path) as pdf:
            text = "\n".join(
                page.extract_text() for page in pdf.pages if page.extract_text()
            )
        return text if text else "N/A"
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return "N/A"
        
# Function to extract text from DOCX
def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        return text if text else "N/A"
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return "N/A"

# Function to scan a folder and subfolders for PDF and DOCX files
def scan_files_and_extract_data(folder_path):
    data = []
    skipped_files = []

    for subdir, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(subdir, file)
            text = "N/A"

            if file.endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            elif file.endswith('.docx'):
                text = extract_text_from_docx(file_path)

            # Extract names, emails, and phone numbers from the text and store it in the a list
            names = extract_names(text) if text != "N/A" else []
            emails, phones = extract_emails_and_phone_numbers(text) if text != "N/A" else ([], [])

            # Default to "N/A" if no information is found
            name = names[0] if names else "N/A"
            phone = phones[0] if phones else "N/A"
            email_list = emails[:5] if emails else ["N/A"] * 5

            # Append to dataset with folder info
            folder_name = os.path.basename(subdir)  # Get the folder name
            data.append([folder_name, file, name, phone] + email_list)

            # Track skipped files (those with no useful data)
            if name == "N/A" and phone == "N/A" and all(email == "N/A" for email in email_list):
                skipped_files.append(file)

    return data, skipped_files

# Function to clean email (removes unwanted leading/trailing characters)
def clean_email(email):
    if not isinstance(email, str):  # Ensure email is a string
        return "N/A"

    # Remove unwanted leading/trailing characters
    email = re.sub(r'^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$', '', email)

    return email  # Return the cleaned email

# Function to clean entire DataFrame
def clean_data(df):
    # Remove duplicate rows
    df = df.drop_duplicates()

    # Normalize phone numbers (format: (123) 456-7890)
    df["Phone"] = df["Phone"].str.replace(r'\D', '', regex=True)  # Keep only numbers
    df["Phone"] = df["Phone"].apply(lambda x: f"({x[:3]}) {x[3:6]}-{x[6:]}" if len(x) == 10 else "N/A")

    # Apply email cleaning
    df.iloc[:, 4:] = df.iloc[:, 4:].map(clean_email)

    # Capitalize names properly
    df["Name"] = df["Name"].apply(lambda x: " ".join(word.capitalize() for word in x.split()) if x != "N/A" else "N/A")

    # Trim whitespace
    df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

    return df

# Extract data from files
data, skipped_files = scan_files_and_extract_data(folder_path)

# Convert data to DataFrame and include Folder column
df = pd.DataFrame(data, columns=["Folder", "File Name", "Name", "Phone", "Email 1", "Email 2", "Email 3", "Email 4", "Email 5"])

df = clean_data(df)

# Set the folder and filename for saving the CSV
folder_name = 'folder_name'  # Replace with actual folder name
file_name = 'file_name.csv'  # Replace with desired file name

# Construct the full path dynamically
csv_path = f'/content/drive/MyDrive/{folder_name}/{file_name}'

# Save the DataFrame to CSV
df.to_csv(csv_path, index=False)

# Check if the first row contains headers, if not, add them
existing_data = worksheet.get_all_values()
if not existing_data or existing_data[0] != ["Folder", "File Name", "Name", "Phone", "Email 1", "Email 2", "Email 3", "Email 4", "Email 5"]:
    worksheet.insert_row(["Folder", "File Name", "Name", "Phone", "Email 1", "Email 2", "Email 3", "Email 4", "Email 5"], 1)

# Append the data rows
worksheet.append_rows(df.astype(str).values.tolist())  # Convert all data to strings before appending

print("\n✅ Data cleaned (including emails) and successfully written to Google Sheet!")

# Display summary
total_files = sum([len(files) for _, _, files in os.walk(folder_path)])
processed_files = len(df)
skipped_count = total_files - processed_files

print("\n🔹 Summary Report 🔹")
print(f"📂 Total Files in Folder: {total_files}")
print(f"✅ Processed Files: {processed_files}")
print(f"⚠️ Skipped Files: {skipped_count} ")
if skipped_files:
    print("\n🛑 Skipped Files (No Data Extracted):")
    for file in skipped_files:
        print(f"- {file}")

print("\n✅ Data successfully written to Google Sheet!")
