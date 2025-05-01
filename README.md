# HR_Automation_System - Google Sheet Formulas 

## Import Source

#### Pre-interview Questionnaire 

To consolidate applicant pre-interview data from multiple spreadsheet sources into a single, organized view that captures key fields:
- Timestamp
- Full Legal Name
- School Email
- Recipient Email
- School Requirement

#### Calendly Extracted Data
- Email
- Event Date/Time
  
#### Contract Response Data
- Email
- Timestamp

---

## Goal
- Create a unified and reliable applicant tracking system using Google Sheets.
- Automatically import and clean data from up to 33 sheets using `IMPORTRANGE` and `QUERY`.
- Filter only the relevant columns.
- Use `VLOOKUP` to check if each applicant has:
  - Responsed Pre-interview Questions
  - Reserved an interview via Calendly
  - Responsed contract
- Display real-time status for each applicant.

---

## Impact
- Operational Efficiency: Saves time and reduces manual effort by automating data merging and status tracking.
- Accuracy & Consistency: Ensures clean and uniform data by trimming, cleaning, and matching lowercase values.
- Real-Time Visibility: Provides up-to-date insights on each applicant’s progress without navigating multiple sheets.
- Clear Applicant Status: Helps teams stay informed about which applicants are actively progressing, making it easier to follow up, coordinate next steps, or flag incomplete stages.
- Centralized Data Access: Combines responses from multiple sources into a single view, reducing manual tracking and improving overall visibility.

