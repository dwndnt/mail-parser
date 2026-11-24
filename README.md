
# ✉️ Mail Parser

**A lightweight email extraction and parsing tool built with Streamlit.**

Mail Parser is a web-based application designed to extract, organize, and export information from Microsoft Outlook `.msg` and standard `.eml` email files.

It transforms email content into structured, reusable outputs — making it easier to inspect recipients, manage attachments, review metadata, and prepare email context for AI-assisted workflows.

<p align="center">

  ![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)
  ![Streamlit](https://img.shields.io/badge/Framework-Streamlit-FF4B4B?style=flat-square&logo=streamlit&logoColor=white)
  ![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
  ![Status](https://img.shields.io/badge/Status-Active-success?style=flat-square)

</p>

---

## 📌 Overview

Mail Parser simplifies the process of extracting useful information from email files without requiring manual inspection.

Upload an email file, process its contents, and export the extracted information into practical formats such as Excel, TXT, ZIP, and Markdown.

### Key Highlights

- **Email Parsing** — Supports Outlook MSG and standard EML formats.
- **Recipient Extraction** — Identifies TO, CC, and BCC recipients.
- **Attachment Management** — Organizes attachments by file extension.
- **Metadata Extraction** — Retrieves important email properties.
- **AI-Ready Export** — Generates structured Markdown for downstream AI workflows.

---

## 🚀 Features

| Feature | Description | Output |
|---|---|---|
| Recipient Extraction | Extracts TO, CC, and BCC recipients | Excel `.xlsx` |
| Attachment Extraction | Extracts and groups attachments by extension | Individual files |
| Attachment Archiving | Downloads attachments individually, by extension, or together | ZIP `.zip` |
| Email Body Extraction | Converts readable email content into plain text | TXT `.txt` |
| Metadata Extraction | Exports email headers and properties | Excel `.xlsx` |
| AI-Friendly Context | Generates structured email information for AI processing | Markdown `.md` |

### 1. 👥 Recipient Extraction

Extract recipient information from email files, including:

- TO recipients
- CC recipients
- BCC recipients, when available
- Recipient names and email addresses

Export the extracted recipient list into a structured Excel workbook.

**Output example:**

`extracted_receipient-email-name.xlsx`

> The filename spelling `receipient` is intentionally retained for compatibility with the requested naming convention.

### 2. 📎 Attachment Management

Extract email attachments and automatically group them according to their file extensions.

Available download options:

- Download each attachment individually
- Download all attachments of a specific extension as ZIP
- Download all attachment groups in one ZIP archive

### 3. 📝 Email Body Extraction

Extract readable email body content and convert it into plain text.

- Supports plain-text email content
- Converts HTML email content into readable text
- Provides an in-app text preview
- Exports the extracted content as TXT

### 4. ℹ️ Metadata Extraction

Retrieve and export important email information, including:

- Subject
- Sender
- Sender email address, when available
- TO, CC, and BCC headers
- Date
- Message-ID
- Importance
- Sensitivity
- Attachment count
- Source format

Metadata can be exported into an Excel workbook for further analysis.

### 5. ✨ AI-Friendly Markdown Export

Generate a structured Markdown document containing:

- Email information
- Sender and recipient details
- Attachment inventory
- Readable email body

The generated Markdown can be used as contextual input for AI assistants, summarization workflows, and further text analysis.

---

## 🛠️ Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application logic |
| Streamlit | Web interface and interaction |
| extract-msg | Outlook MSG parsing |
| Python Email Package | Standard EML parsing |
| OpenPyXL | Excel workbook generation |
| ZIP / io | Attachment archiving and in-memory processing |

---

## 📂 Project Structure

```text
mail-parser/
│
├── .streamlit/
│   └── config.toml
│
├── .gitignore
│
├── README.md
│
├── app.py
│
└── requirements.txt
```

---

## ⚙️ Installation & Setup

### Prerequisites

Make sure the following are installed:

- Python 3.10 or newer
- Git
- pip

### 1. Clone the Repository

```bash
git clone https://github.com/dwndnt/mail-parser.git

cd mail-parser
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
```

### 3. Activate the Environment

**Windows PowerShell**

```powershell
.venv\Scripts\Activate.ps1
```

**macOS / Linux**

```bash
source .venv/bin/activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Run the Application

```bash
streamlit run app.py
```

Once the application starts, open the local URL displayed in your terminal.

Default address:

```text
http://localhost:8501
```

---

## ☁️ Deployment

Mail Parser can be deployed using [Streamlit Community Cloud](https://share.streamlit.io/).

### Deployment Steps

1. Push the project repository to GitHub.
2. Open [Streamlit Community Cloud](https://share.streamlit.io/).
3. Sign in and select **Create app**.
4. Select your GitHub repository.
5. Choose the deployment branch.
6. Set `app.py` as the main application file.
7. Click **Deploy**.

No API keys or application secrets are required for the current version.

---

## 📤 Output Naming Convention

Mail Parser uses descriptive filenames to make exported files easier to identify.

| Output | Naming Convention |
|---|---|
| Recipient Excel | `extracted_receipient-<email-name>.xlsx` |
| Metadata Excel | `email-metadata-<email-name>.xlsx` |
| Email Body | `email-body-<email-name>.txt` |
| AI Markdown | `email-context-<email-name>.md` |
| Attachment ZIP | `attachments-<email-name>-<extension>.zip` |
| All Attachments ZIP | `all-attachments-<email-name>.zip` |

`<email-name>` represents the original email filename without its extension.

---

## 📋 Supported Formats & Limitations

### Supported Email Formats

| Format | Parser | Support |
|---|---|---|
| `.msg` | `extract-msg` | Outlook email files |
| `.eml` | Python standard email package | RFC 822 email files |

### Known Limitations

- **BCC availability:** BCC information can only be extracted when it exists in the source file. Many received emails do not retain BCC recipient information.
- **Outlook properties:** Some proprietary Outlook properties may not be exposed by every MSG file.
- **HTML content:** HTML email bodies are converted into readable text. Original HTML is not exported separately.
- **Encrypted files:** Password-protected or encrypted email files may fail to parse.
- **Malformed files:** Corrupted or non-standard email files may produce extraction errors.
- **Large files:** Processing performance depends on file size and available server resources.

---

## 🔒 Privacy & Security

Mail Parser processes uploaded files within the active Streamlit application session.

Email files may contain confidential information, personal data, and sensitive business communications.

**Important considerations:**

- Avoid deploying the application publicly for confidential corporate email processing unless explicitly permitted by your organization's policies.
- Restrict application access when processing internal business communications.
- Apply appropriate upload-size limits.
- Define a clear data retention policy.
- Avoid logging email body content or attachment contents.
- Use authentication and access controls for internal production deployments.

For enterprise or corporate environments, deployment should follow the organization's information-security and data-handling requirements.

---

## 🗺️ Roadmap

Potential areas for future enhancement:

- [ ] Batch processing for multiple email files
- [ ] Export original HTML email body
- [ ] Enhanced Outlook metadata extraction
- [ ] Email conversation and thread analysis
- [ ] Advanced recipient analytics
- [ ] Search and filtering across extracted emails
- [ ] AI-powered email summarization
- [ ] Additional export formats

---

## 📄 License

This project is licensed under the **MIT License**.

You are free to use, modify, and distribute this software in accordance with the license terms.

---

<p align="center">

  **Mail Parser**

  *Extract. Organize. Understand.*

  Built with Python and Streamlit

</p>
