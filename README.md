# Automated Email Sender for Employee Payslips

This Python script automates the process of sending payslips to employees via email. It reads employee data from a CSV file, finds the most recent PDF payslip for each employee, and sends an email with the payslip attached. The email body is formatted in HTML and includes embedded images.

## Features

- **CSV Data Handling**: Reads employee names and email addresses from a CSV file.
- **PDF Matching**: Finds the most recent PDF file matching the employee's name.
- **Email Automation**: Sends personalized emails with attached payslips.
- **HTML Email Body**: Uses HTML to format the email body, including embedded images.
- **Logging**: Prints status messages to the console for each email sent.
 
---
## Prerequisites

Before running the script, ensure you have the following:

- **Python 3.x**: The script is written in Python 3.12.5
- **Required Libraries**: Install the necessary Python libraries using pip:
  ```bash
  pip install pandas==2.2.2 pywin32==306
  ```                  
- **CSV File**: A CSV file named Employee Emails.csv with columns Employee Name and Employee Email.

- **PDF Files**: PDF payslips named in a way that they can be matched with employee names (e.g., JohnDoe.pdf).

- **Images**: Two images (xti.png and rss.png) for embedding in the email.
---
## Getting Started
 - Clone the Repository
 - To get started, clone this repository to your local machine using the following command:
   ```bash
      git clone https://github.com/MushafMughal/PaySlip-Email-Shooter.git
   ```
- After cloning, navigate to the project directory:
  ```bash
    cd PaySlip-Email-Shooter
  ```
---

## Usage
 - Prepare the CSV File: Ensure your CSV file (Employee Emails.csv) is formatted correctly with the columns Employee Name and Employee Email.

 - Place PDFs and Images: Place all PDF payslips and the images (xti.png and rss.png) in the same directory as the script.

 - Run the Script: Execute the script using Python:
   ```bash
      python send_payslips.py
   ```
 - Check Output: The script will print status messages to the console, indicating whether each email was sent successfully or if any errors occurred.
---

## Contributing
 - Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.
