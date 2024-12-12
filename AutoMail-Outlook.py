import pandas as pd
import win32com.client as win32

# Load the Excel file
file_path = r"C:\Users\yoriahi\Desktop\Cache\To Youssef\Obs\contacts.xlsx"
df = pd.read_excel(file_path)

# Initialize Outlook
outlook = win32.Dispatch('outlook.application')

# Define email template
email_template = """
<html>
    <body>
        <p>Hello,</p>

        <p>This is a sample email template. You can customize the content using HTML.</p>
        <p>Here’s how to structure your email:</p>
        <ul>
            <li><strong>HTML Structure:</strong> Use <code>&lt;html&gt;</code>, <code>&lt;body&gt;</code>, and other HTML tags to format your email.</li>
            <li><strong>Tables:</strong> Use the <code>&lt;table&gt;</code> tag to create tables. For example:</li>
        </ul>
        <pre>
        &lt;table border="1" style="border-collapse:collapse; width:100%;"&gt;
            &lt;tr&gt;
                &lt;th&gt;Header 1&lt;/th&gt;
                &lt;th&gt;Header 2&lt;/th&gt;
            &lt;/tr&gt;
            &lt;tr&gt;
                &lt;td&gt;Data 1&lt;/td&gt;
                &lt;td&gt;Data 2&lt;/td&gt;
            &lt;/tr&gt;
        &lt;/table&gt;
        </pre>
        <p>For your specific use case, please provide the following information:</p>
        <ol>
            <li>Item status (e.g., ACTIVE, END OF LIFE, OBSOLETE).</li>
            <li>Do you offer equivalent items for OBSOLETE or END OF LIFE items?</li>
        </ol>
        <br>
        <table border="1" style="border-collapse:collapse; width:100%;">
            <tr>
                <th style="background-color: darkgreen; color: white; padding: 10px;">Item</th>
                <th style="background-color: darkgreen; color: white; padding: 10px;">Status</th>
                <th style="background-color: darkgreen; color: white; padding: 10px;">Equivalent Item</th>
            </tr>
            <tr>
                <td style="padding: 10px;">{reference}</td>
                <td style="padding: 10px;"></td>
                <td style="padding: 10px;"></td>
            </tr>
        </table>

        <p>If you are not the right contact for this inquiry, please let us know who to reach out to.</p>

        <p>Thank you for your collaboration. We look forward to your response.</p>

        <p>Best Regards,</p>
    </body>
</html>
"""

# Prepare emails
for index, row in df.iterrows():
    try:
        contact = row['Contact'] 
        fabricant = row['Fabricant'] 
        reference = row['Ref']
        
        # Construct the email
        email_subject = f'Request About Status - {fabricant}'
        email_body = email_template.format(reference=reference)

        # Create a new email draft
        mail = outlook.CreateItem(0)  # 0: olMailItem
        mail.Subject = email_subject
        mail.HTMLBody  = email_body
        mail.To = contact
        mail.Save()  # Save the draft
        
        print(f"Draft created for {contact}.")
        
    except Exception as e:
        print(f"Error processing row {index}: {e}")

print("Drafts created successfully in Outlook.")