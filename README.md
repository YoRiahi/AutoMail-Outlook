# AutoMail-Outlook
This Python script automates the creation of email drafts in Microsoft Outlook using data from an Excel file. It is useful for sending bulk emails with custom messages, allowing easy personalization for each recipient.

## Features
- **Automated Email Drafting**: Automatically creates drafts in Outlook for each recipient.
- **Excel Integration**: Extracts recipient information (email, subject, message, etc.) from an Excel sheet.
- **Customizable**: Modify the script easily to add more fields or adjust email formatting.
- **Safe**: The script only creates drafts in Outlook—emails are not sent automatically, allowing for review before sending.

## Requirements
- **Python 3.x**
- **Libraries**: `pandas`, `pywin32` - Install using: ```bash pip install pandas pywin32 ```
- **Microsoft Outlook** installed and configured on your system.
