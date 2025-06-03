# Internship Recruitment Automation Emailing and Workflow System

## Purpose
The purpose of this project is to automate and optimize the internship recruitment workflow while maintaining data security and integrity. The system is designed to streamline the management of intern data by organizing candidates based on their recruitment status (e.g., pre-interview, terminated, international interns, onboarded). It automates the communication process through emails at different stages, ensuring that candidates receive timely and accurate updates based on their current stage in the recruitment pipeline. Additionally, it protects sensitive data by locking specific columns in Google Sheets. A counter is implemented to ensure that the maximum number of emails sent per day is not exceeded.

## Goal
The goals of this project are to:
1. **Data Organization & Scalability**: Automatically move candidates to the appropriate sheets based on their recruitment status:
   - **Terminated**: Candidates who didn't pass the pre-interview or interview or who self-terminate.
   - **International Interns**: Candidates on Optional Practical Training (OPT) or Curricular Practical Training (CPT).
   - **Onboarded**: Candidates who have completed the recruitment process and are onboarded.
> Keep the recruitment system organized and scalable by removing inactive candidates from the system and reducing manual intervention.
2. **Email Automation**: Automatically send pre-interview, Calendly, and contract emails based on the candidate’s recruitment stage, using logic and conditions that trigger the emails at the right time. This includes sending follow-up emails at a set time to ensure that no candidate is overlooked.
3. **Email Sending Limits**: To prevent exceeding the daily email limit, a counter tracks the number of emails sent, ensuring the set daily limit is respected.
4. **Security**: certain columns in the sheet are locked to prevent unauthorized changes and protect sensitive data.

## Impact
By automating the management of intern data and email communication, this system:
- **Improves Data Security**: Locks critical columns to prevent unauthorized edits, ensuring that sensitive information is protected.
- **Enhances Data Organization**: Moves candidates to appropriate sheets based on their status, keeping the active recruitment pool clear and well-organized, improving scalability.
- **Optimizes Email Communication**: Sends timely emails based on the candidate's stage in the recruitment process, enhancing candidate experience and reducing manual oversight.
- **Prevents Over-Sending**: The email counter ensures that no more than the maximum allowed number of emails is sent per day, preventing email overload.
- **Boosts Operational Efficiency**: Minimizes time spent manually sorting and updating data by automating repetitive tasks, allowing the recruitment team to concentrate on high-priority, people-centric activities such as follow-up calls, interviews, and candidate selection.

To set up oAuth2 I refereneced - https://github.com/googleworkspace/apps-script-oauth2
