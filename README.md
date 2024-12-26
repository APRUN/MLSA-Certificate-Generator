# MLSA Certificate Generator and Mailer

Automate the generation and distribution of Microsoft Learn Student Ambassador (MLSA) certificates using Python and Microsoft Office automation.

## Features

- Automated certificate generation in both DOCX and PDF formats
- Bulk email distribution through Outlook
- Support for custom certificate templates
- Automated participant data processing

## Prerequisites

- Python 3.x
- Microsoft Office (Outlook and Word)
- Git

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/mlsa-certificate-automation.git
cd mlsa-certificate-automation
```

2. Install required dependencies:
```bash
pip3 install -r requirements.txt
```

## Project Structure

```
.
├── Data/
│   ├── Event Certificate Template.docx
│   ├── Mail.xlsm
│   ├── ParticipantList.csv
│   ├── mailtemplate.html
│   └── temp.csv
├── Output/
│   ├── doc/
│   └── pdf/
├── certificate.py
├── main_certificate.py
├── requirements.txt
└── README.md
```

## Usage

### 1. Preparing Participant Data

1. Navigate to the `Data` folder
2. Open `ParticipantList.csv`
3. Add participant names and email addresses in the required format

### 2. Generating Certificates

1. Run the main certificate generator:
```bash
python main_certificate.py
```
2. Check the `Output` folder for generated certificates:
   - `doc/`: Contains DOCX format certificates
   - `pdf/`: Contains PDF format certificates

### 3. Sending Certificates via Email

1. Open `Mail.xlsm` in the Data folder
2. Enable macros when prompted:
   - Click "Enable Content" if shown
   - Or enable via File > Options > Trust Center > Trust Center Settings > Macro Settings
3. Access macros:
   - Press Alt + F8
   - Or search for "Macros" in Excel's search bar
4. Run the `send_mail` macro to begin email distribution through Outlook

## Dependencies

```
python-docx
docx2pdf
openpyxl
uvicorn
aiofiles
fastapi
aiosmtplib
python-dotenv
python-multipart
```

## Troubleshooting

- If macros are disabled, ensure you've enabled them in Excel's Trust Center
- Check Outlook configuration if emails fail to send
- Verify all paths in configuration files match your setup

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
