# Email Toolkit

A Streamlit web application for extracting information from Outlook `.msg` and standard `.eml` email files.

## Features

- Extract TO, CC, and BCC recipients into Excel
- Extract attachments, grouped by file extension
- Download attachments individually, by extension as ZIP, or as one ZIP containing all groups
- Extract readable email body into TXT
- Export email metadata into Excel
- Generate structured Markdown (`.md`) as AI-friendly email context

## Run locally

Requires Python 3.10 or newer.

```bash
git clone https://github.com/<your-username>/email-toolkit.git
cd email-toolkit

python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Open the local URL shown by Streamlit (usually http://localhost:8501).

## Deploy to Streamlit Community Cloud

1. Push this repository to GitHub.
2. Open https://share.streamlit.io/
3. Select **Create app**.
4. Choose the repository, branch, and `app.py` as the main file.
5. Deploy.

No secrets are required for the current version.

## Output naming

Recipient workbook:

`extracted_receipient-<original-email-filename-without-extension>.xlsx`

The spelling `receipient` is retained to match the requested output naming convention.

## Supported formats and limitations

- `.msg`: parsed using `extract-msg`.
- `.eml`: parsed using Python's standard email package.
- BCC is only available if it is actually present in the source file. Many received emails do not retain BCC recipient information.
- Some proprietary Outlook properties may not be exposed by all MSG files.
- HTML email bodies are converted to readable text for the TXT and Markdown outputs. Original HTML is not exported as a separate file in this version.
- Password-protected, encrypted, malformed, or unusually large email files may fail to parse.

## Privacy and deployment considerations

The app processes uploaded files in the Streamlit server session. Do not deploy this publicly for confidential corporate email processing unless your organization's information-security and data-handling policies explicitly permit it.

For an internal production deployment, consider authentication, access restrictions, upload-size limits, audit logging, and a defined retention policy. Avoid logging email content or attachment contents.

## Repository structure

```text
email-toolkit/
├── .streamlit/
│   └── config.toml
├── .gitignore
├── README.md
├── app.py
└── requirements.txt
```

## License

MIT
