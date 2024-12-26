# MLSA Certificate Automation 🚀  

This project automates the process of generating and sending certificates for events conducted under the Microsoft Learn Student Ambassadors (MLSA) program. With this tool, you can create personalized certificates in both `.docx` and `.pdf` formats and send them directly via email through Outlook.  

---

## Features  
- **Certificate Automation**: Quickly generate certificates for all participants.  
- **PDF & DOCX Output**: Organizes output files into `.docx` and `.pdf` folders.  
- **Email Automation**: Send certificates directly via email using Excel macros.  

---

## Project Structure  

📂 Data/
├── Event Certificate Template.docx # Certificate design template
├── Mail.xlsm # Excel file with email macros
├── ParticipantList.csv # Participant names and emails
├── mailtemplate.html # HTML email template
├── temp.csv # Temporary file for data handling
📄 .gitignore # Ignored files for Git
📄 README.md # Project documentation
📄 certificate.py # Certificate generation script
📄 main_certificate.py # Main script for execution
📄 requirements.txt # Python dependencies

yaml
Copy code

---

## Prerequisites  

1. **Python**: Install Python 3.6 or above.  
2. **Microsoft Office**: Ensure you have Excel and Outlook configured.  
3. **Macro Settings**: Enable macros in Excel (instructions provided below).  

---

## Installation  

1. Clone the repository:  
   ```bash
   git clone <repository-url>
   cd <repository-folder>
Install the required Python packages:
bash
Copy code
pip3 install -r requirements.txt
How to Use
Step 1: Add Participant Details
Open the ParticipantList.csv file located in the Data/ folder.
Add participant names and email addresses in the provided columns.
Step 2: Generate Certificates
Navigate to the root directory of the project.
Run the main_certificate.py script:
bash
Copy code
python main_certificate.py
This will:
Create an Output/ folder in the main directory.
Inside Output/, you will find:
doc/ folder containing .docx certificates.
pdf/ folder containing .pdf certificates.
Step 3: Send Emails
Open the Mail.xlsm file from the Data/ folder.
Enable Macros:
If macros are disabled, a bar will appear at the top of the Excel sheet prompting you to enable them.
If no prompt appears, go to Excel settings and manually enable macros.
Search for "Macros" using the search bar in Excel (top center).
Run the send_mail macro:
Open the "Macros" window.
Select send_mail and click "Run."
Emails with attached certificates will now be sent automatically through Outlook.
Requirements
Below are the Python dependencies used in this project. Install them using the command:

bash
Copy code
pip3 install -r requirements.txt
Dependencies:
python-docx
docx2pdf
openpyxl
uvicorn
aiofiles
fastapi
aiosmtplib
python-dotenv
python-multipart
Enabling Macros
To enable macros in Excel:

Go to File > Options.
Select Trust Center > Trust Center Settings.
Go to Macro Settings and choose Enable all macros.
Notes
Ensure Outlook is configured with an active account before running the send_mail macro.
Customize the certificate template by editing the Event Certificate Template.docx file in the Data/ folder.
Modify the email body by editing the mailtemplate.html file.
Troubleshooting
Macros Disabled: Ensure macros are enabled in Excel as described above.
Emails Not Sending: Verify that Outlook is open and properly configured.
Contributions
Contributions are welcome! Feel free to fork this repository, create new features, or fix bugs. Submit a pull request, and I’ll be happy to review it.

License
This project is licensed under the MIT License. See the LICENSE file for more information.

Contact
For queries or assistance, feel free to reach out:

Email: your-email@example.com
GitHub: Your GitHub Profile
🎉 Automate. Generate. Simplify!
vbnet
Copy code

### How to Use It  
1. Copy the above markdown code into a `README.md` file in your project directory.  
2. Replace `<repository-url>` and `[your-email@example.com]` with your actual details.  
3. Customize further as needed!  

Let me know if you'd like any tweaks or enhancements. 😊
