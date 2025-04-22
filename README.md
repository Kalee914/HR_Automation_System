# HR_Automation_System


### Purpose
This project is all about making the internship recruitment process smoother and less time-consuming. Instead of manually tracking candidates and sending the same emails repeatedly, the goal is to create a system that automatically handles those routine tasks—so the HR team can spend more time having real conversations and less time on administrative work. The system is meant to improve efficiency, reduce manual effort, and support consistent communication while managing a large number of applicants.

### Goals
- Personalize and send onboarding or pre-interview emails based on predefined templates.  
- Track recruitment stages and update statuses in real-time within the same Google Sheet.  
- Use the Gmail API (instead of `GmailApp.sendEmail` and `MailApp.sendEmail`) to get around email sending limits and support larger-scale outreach.  
- Move candidates to corresponding sheets once they’ve reached certain stages—so only active applicants remain in the system. This keeps things clean, scalable, and easy to manage.  
- Keep all information in one central place—eliminating the need to jump between tabs or lose track of past interactions.  
- Use smart logic to only contact the right candidates at the right time.
- Enforce access control security by developing the automation from a main owner account and selectively sharing access with HR using Google’s built-in sharing permissions (Viewer, Commenter, Editor roles). HR team members are granted limited access through protected sheet ranges and role-based sharing, ensuring only authorized users can perform sensitive operations. This helps maintain data integrity, prevent unauthorized changes, and secure confidential information.
  
> **Note:** Logic, code, and formulas for different parts of the system are organized in separate Git branches.  
> Browse the branches to explore specific features, implementations, and decision flows.

### Impact
This automation will significantly reduce the workload of HR team members by eliminating repetitive tasks such as email drafting, sending, and manual status updates. It will help ensure timely communication with candidates, leading to a smoother and more professional recruitment experience. Ultimately, the project will support the organization’s ability to recruit more effectively, especially when handling large internship intakes.
