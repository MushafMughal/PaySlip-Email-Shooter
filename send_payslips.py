import pandas as pd
import os
import win32com.client as win32
import re
from datetime import datetime

df = pd.read_csv('Employee Emails.csv')  # Replace with your file path
pdf_directory = os.getcwd()  # Current working directory
outlook = win32.Dispatch('Outlook.Application')

def find_most_recent_pdf(base_name, pdf_directory):

    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]
    matching_pdfs = [f for f in pdf_files if f.lower().startswith(base_name.lower())]
    
    if not matching_pdfs:
        return None
    
    matching_pdfs.sort(key=lambda f: os.path.getmtime(os.path.join(pdf_directory, f)), reverse=True)
    return os.path.join(pdf_directory, matching_pdfs[0])

for index, row in df.iterrows():
    name = row['Employee Name']
    email = row['Employee Email']
    
    pdf_path = find_most_recent_pdf(name, pdf_directory)
    
    if not pdf_path:
        print(f"No matching PDF found for {name}")
        continue

    pdf_filename = os.path.basename(pdf_path)
    # Extract month and year using regex
    match = re.search(r"([A-Z]+)\s*['-]?\s*(\d{4})", pdf_filename)

    if match:
        month_year = f"{match.group(1)} {match.group(2)}"
        subject = f"PAYSLIP FOR {month_year} - {name}"
    else:
        subject = f"PAYSLIP FOR - {name}"
    
    # Create a new mail item
    mail = outlook.CreateItem(0)  # 0 corresponds to MailItem
    mail.To = email
    mail.Subject = subject
    body = f"""
    <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
            <strong>Dear {name},</strong>

            <p>I hope this email finds you well. Please find your attached payslip for your reference.</p>

            <p>If you have any questions regarding your salary details, deductions, or benefits, 
            or if you require any further clarification, please do not hesitate to reach out. 
            Our HR team is always available to assist you with any concerns you may have.</p>

            <p>We appreciate your dedication and contributions to the company. 
            Thank you for being an integral part of our team!</p>

            <p>Best regards,</p>

            <table style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
                <tr>
                    <td style="font-weight: bold; font-size: 16px;">Syed Mustafa Shah</td>
                </tr>
                <tr>
                    <td style="color: #0056b3; font-weight: bold; font-size: 14px;">HUMAN RESOURCES</td>
                </tr>
                <tr>
                    <td>322 Julie Rivers Drive | Sugar Land, TX 77478</td>
                </tr>
                <tr>
                    <td>
                        <strong>Email:</strong> 
                        <a href="mailto:MustafaS@xclusivetradinginc.com" style="color: #0056b3; text-decoration: none;">
                            MustafaS@xclusivetradinginc.com
                        </a>
                    </td>
                </tr>
                <tr>
                    <td>
                        <strong>Contact:</strong> <span style="color: #0056b3;">(EXT 207)</span>
                    </td>
                </tr>
                <tr>
                    <td>
                        <strong>WhatsApp:</strong> 
                        <a href="https://wa.me/923113859635" style="color: #0056b3; text-decoration: none;">
                            +92-311-3859635
                        </a>
                    </td>
                </tr>
                <tr>
                    <td>
                        <img src="cid:xti" alt="XTI Logo" width="250" height="250" style="display:block;">
                    </td>
                    <td>
                        <img src="cid:rss" alt="RSS Logo" width="350" height="150" style="display:block;">
                    </td>
                </tr>
            </table>
        </body>
    </html>
    """

    mail.HTMLBody = body  # Set the email body as HTML
    mail.Attachments.Add(pdf_path)
    xti_path = os.path.abspath("xti.png")
    rss_path = os.path.abspath("rss.png")
    # Attach images and set CID for embedding in email
    attachment1 = mail.Attachments.Add(xti_path)
    attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "xti")
    attachment2 = mail.Attachments.Add(rss_path)
    attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "rss")
    mail.Send()
    print(f"Email sent to {email} with attachment {pdf_path}")

print("All emails have been sent.")