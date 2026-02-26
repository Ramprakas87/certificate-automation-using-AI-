# certificate-automation-using-AI-
AI-based Certificate Automation system that generates personalized certificates in bulk using participant data from Excel/CSV files.  Automatically creates and exports professional PDF certificates with dynamic data insertion, reducing manual effort and errors.

---

# AI-Assisted Certificate Automation using Excel VBA & Word

![Excel](https://img.shields.io/badge/Excel-VBA-green)
![Word](https://img.shields.io/badge/Word-Automation-blue)
![Automation](https://img.shields.io/badge/Type-Bulk%20PDF%20Generation-orange)
![Status](https://img.shields.io/badge/Status-Completed-success)

---

## 1. Project Overview

This project is an AI-assisted Certificate Automation System built using Microsoft Excel VBA and Microsoft Word.

The system automatically generates personalized PDF certificates in bulk by dynamically replacing placeholders in a Word certificate template using candidate data stored in Excel.

The VBA logic was developed using structured AI prompt engineering, demonstrating practical AI-assisted development and real-world business automation.

---

## 2. Problem Statement

Generating certificates manually for multiple candidates is:

* Time-consuming
* Error-prone
* Repetitive
* Inefficient for bulk operations

This automation system eliminates manual effort and streamlines the entire certificate generation workflow.

---

## 3. Input Structure

### 3.1 Word Certificate Template

The Word template contains the following placeholders inside textboxes:

* `<<CANDIDATE NAME>>`
* `<<DATE>>`
* `<<CERTIFICATE ID>>`

These placeholders are dynamically replaced during automation.

---

### 3.2 Excel Data Sheet

The Excel sheet is structured as:

* Column A — Candidate ID
* Column B — Candidate Name
* Column C — Date
* Column D — Certificate ID

Each row represents one candidate record.

---

## 4. Automation Workflow

When the macro is executed:

1. A file dialog prompts the user to select the Word certificate template.
2. A folder dialog prompts the user to select the destination folder for saving PDFs.
3. The macro reads each candidate record from Excel.
4. It replaces:

   * `<<CANDIDATE NAME>>`
   * `<<DATE>>`
   * `<<CERTIFICATE ID>>`
5. A separate PDF is generated for each candidate.
6. Each file is automatically named using the Candidate ID
   (Example: `1001_Certificate.pdf`).

---

## 5. Key Features

* Excel–Word integration using VBA
* Dynamic placeholder replacement
* Bulk PDF certificate generation
* Automatic file naming
* Folder selection dialog
* One-click macro execution
* AI-generated and optimized VBA logic
* Scalable for large datasets

---

## 6. Technologies Used

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Microsoft Word Object Model
* File Dialog & Folder Picker
* PDF Export Automation
* AI Prompt Engineering

---

## 7. Business Impact

* Reduced certificate generation time by approximately 90%
* Eliminated repetitive manual editing
* Minimized human errors
* Improved operational efficiency
* Suitable for internships, training programs, workshops, and corporate certifications

---

## 8. AI Contribution

The VBA logic was generated and refined using a structured AI prompt that defined:

* File selection requirements
* Data mapping rules
* Placeholder replacement logic
* PDF export process
* File naming conventions

This highlights the practical use of AI in automation and development workflows.

---

## 9. Project Structure

```
AI-Certificate-Automation/
│
├── Certificate_Template.docx
├── Candidate_Data.xlsx
├── VBA_Module.bas
├── Output/
└── README.md
```

---

## 10. Author

Ram Prakash Patel

Intern – Data Analyst

LinkedIn: [https://www.linkedin.com/in/ram-prakash-patel-62863b378/](https://www.linkedin.com/in/ram-prakash-patel-62863b378/)

---

If you found this project useful, please consider giving it a star.

---

