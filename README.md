# Automated Email Dispatch Tool

An automated utility that reads data from an Excel spreadsheet to send customized emails, automatically attaching files and generating rich content summaries from the source data. This project was developed during my internship to streamline business communication processes. The core logic is built with Java and Apache POI, wrapped in a PowerShell script for seamless email automation, and triggered by a simple one-click batch file.

---
## ✨ Key Features

* **Automated Excel Summaries:** For `.xlsx` files, the tool automatically reads the data, generates a pivot-style summary, and embeds it as a clean **HTML table directly within the email body**. This provides an immediate, at-a-glance view of the key data without needing to open the attachment.

* **Flexible Attachments:** Automatically attaches **any file type** specified in the Excel sheet, whether it's a Word document, PDF, text file, or the original Excel sheet itself.

* **Data-Driven Automation:** Reads recipient email addresses, subject lines, body text, and attachment file paths directly from a central `.xlsx` spreadsheet, making it easy to manage bulk email campaigns.

* **Double-Click Execution:** The entire process is initiated by double-clicking a single `.bat` file, making it accessible to non-technical users.

---
## 🚀 Tech Stack

* **Core Logic:** Java, Apache POI
* **Automation/Scripting:** PowerShell, Windows Batch Script
* **Content Generation:** HTML
