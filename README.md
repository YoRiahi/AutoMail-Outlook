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

## Setup
1. Clone or download the repository.
   ```bash
   git clone https://github.com/YoRiahi/AutoMail-Outlook.git
2. Update the file_path variable in the script with the full path to your Excel file containing the email data.
3. Your Excel file should contain the following columns:
- **Name**: Recipient’s name.
- **Email**: Recipient’s email address.
- **Subject**: Subject of the email.
- **Message**: The body of the email.
5. Run the script:
  python main.py

## Usage
1. Install dependencies: Run the following command in your terminal or command prompt:
   pip install pandas pywin32
2. Prepare your Excel file with the necessary columns (Name, Email, Subject, Message).
3. Run the script to create email drafts in Outlook:
python AutoMail-Outlook.py
4. The script will generate email drafts for each recipient in your Outlook inbox. Review the drafts and send them manually.

## Limitations
• Drafts Only: The script does not send emails directly—only drafts are created in Outlook.
• Outlook Specific: This script currently only works with Microsoft Outlook. Future versions may support other email clients.

## Potential Enhancements
• Email Sending: Add functionality to automatically send emails after creating drafts.
• Logging: Implement logging to capture errors and script status.
• Other Email Clients: Extend compatibility to other clients like Gmail or Thunderbird.

## Contributing
Feel free to open issues or create pull requests if you want to contribute improvements or bug fixes.

## License
This project is licensed under the MIT License - see the [LICENSE]LICENSE file for details.

