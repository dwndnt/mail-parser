
from __future__ import annotations

import email
import hashlib
import html
import io
import re
import textwrap
import zipfile

from collections import Counter, defaultdict
from email import policy
from email.header import decode_header, make_header
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Any

import extract_msg
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


# ---------- Page configuration ----------

st.set_page_config(
    page_title="Mail Parser",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------- Theme / CSS ----------

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1350px;
        }

        .hero {
            padding: 2rem 2.2rem;
            border-radius: 18px;
            background: linear-gradient(
                120deg,
                #10243a 0%,
                #1b4965 65%,
                #287094 100%
            );
            color: white;
            margin-bottom: 1.5rem;
            box-shadow: 0 8px 24px rgba(16, 36, 58, 0.15);
        }

        .hero h1 {
            color: white;
            margin: 0 0 .5rem 0;
            font-size: 2.2rem;
            font-weight: 750;
        }

        .hero p {
            color: #dceaf2;
            margin: 0;
            font-size: 1rem;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 16px;
            margin: 12px 0 24px 0;
        }

        .metric-card {
            position: relative;
            overflow: hidden;
            padding: 22px;
            border-radius: 16px;
            color: white;
            min-height: 125px;
            box-shadow: 0 5px 16px rgba(0, 0, 0, 0.10);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .metric-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.16);
        }

        .metric-card::after {
            content: "";
            position: absolute;
            width: 110px;
            height: 110px;
            right: -25px;
            bottom: -40px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.12);
        }

        .metric-card.recipients {
            background: linear-gradient(135deg, #173b57, #287094);
        }

        .metric-card.attachments {
            background: linear-gradient(135deg, #087f8c, #20a4a8);
        }

        .metric-card.body {
            background: linear-gradient(135deg, #5146a5, #8174d8);
        }

        .metric-card.format {
            background: linear-gradient(135deg, #b36b18, #e3a13c);
        }

        .metric-label {
            font-size: 0.92rem;
            font-weight: 600;
            color: rgba(255, 255, 255, 0.90);
            margin-bottom: 12px;
        }

        .metric-value {
            font-size: 2rem;
            font-weight: 750;
            line-height: 1.1;
            color: #ffffff;
            overflow-wrap: anywhere;
        }

        .metric-description {
            font-size: 0.78rem;
            color: rgba(255, 255, 255, 0.82);
            margin-top: 8px;
        }

        .app-footer {
            margin-top: 35px;
            padding: 20px 0 5px 0;
            border-top: 1px solid rgba(128, 128, 128, 0.25);
            text-align: center;
            color: #84909d;
            font-size: 0.85rem;
        }

        .app-footer strong {
            color: #287094;
        }

        .stButton button,
        .stDownloadButton button {
            border-radius: 9px;
            min-height: 2.55rem;
        }

        [data-testid="stFileUploader"] {
            border: 1px dashed #9aabb9;
            border-radius: 14px;
            padding: 12px;
        }

        @media (max-width: 900px) {
            .metric-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }

        @media (max-width: 520px) {
            .metric-grid {
                grid-template-columns: 1fr;
            }

            .hero {
                padding: 1.5rem;
            }

            .hero h1 {
                font-size: 1.8rem;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------- Constants ----------

RECIPIENT_TYPES = {
    "to": "TO",
    "cc": "CC",
    "bcc": "BCC",
    "1": "TO",
    "2": "CC",
    "3": "BCC",
}


# ---------- Helper functions ----------

def decode_text(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, bytes):
        try:
            decoded = value.decode("utf-8", errors="replace")
            return str(make_header(decode_header(decoded)))
        except Exception:
            return value.decode("utf-8", errors="replace")

    try:
        return str(make_header(decode_header(str(value))))
    except Exception:
        return str(value)


def clean_excel_text(value: Any) -> str:
    value = "" if value is None else str(value)
    value = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", value)
    return value[:32767]


def safe_filename(
    name: str,
    fallback: str = "attachment",
) -> str:
    name = Path(name or fallback).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or fallback


def normalize_type(value: Any) -> str:
    if value is None:
        return ""

    value = getattr(value, "value", value)
    text = str(value).strip().lower()

    return RECIPIENT_TYPES.get(
        text,
        text.upper() if text else "",
    )


def parse_address_header(
    value: str,
    kind: str,
) -> list[dict]:
    results = []

    for name, address in getaddresses([value or ""]):
        address = address.strip()
        name = decode_text(name).strip()

        if address or name:
            results.append(
                {
                    "Nama": name,
                    "Email": address,
                    "CC/BCC": kind,
                }
            )

    return results


def msg_attr(
    obj: Any,
    *names: str,
) -> str:
    for name in names:
        try:
            value = getattr(obj, name, None)

            if value is not None and str(value).strip():
                return decode_text(value).strip()

        except Exception:
            pass

    return ""


# ---------- MSG extraction ----------

def extract_msg_data(raw: bytes) -> dict:
    msg = extract_msg.Message(io.BytesIO(raw))

    try:
        recipients = []

        for recipient in (
            getattr(msg, "recipients", None) or []
        ):
            kind = normalize_type(
                getattr(recipient, "type", None)
            )

            if kind not in ("TO", "CC", "BCC"):
                continue

            name = msg_attr(
                recipient,
                "name",
                "displayName",
                "transmittableDisplayName",
            )

            address = msg_attr(
                recipient,
                "smtpAddress",
                "email",
                "address",
            )

            formatted = msg_attr(
                recipient,
                "formatted",
            )

            if not address and formatted:
                parsed = getaddresses([formatted])

                if parsed:
                    name = name or decode_text(
                        parsed[0][0]
                    ).strip()

                    address = parsed[0][1].strip()

            if name or address:
                recipients.append(
                    {
                        "Nama": name,
                        "Email": address,
                        "CC/BCC": kind,
                    }
                )

        # Fallback for MSG parsers that expose header-like fields.
        if not recipients:
            for attr, kind in (
                ("to", "TO"),
                ("cc", "CC"),
                ("bcc", "BCC"),
            ):
                recipients.extend(
                    parse_address_header(
                        msg_attr(msg, attr),
                        kind,
                    )
                )

        attachments = []

        for index, attachment in enumerate(
            getattr(msg, "attachments", []) or [],
            start=1,
        ):
            try:
                data = attachment.data

                if data is None:
                    continue

                if not isinstance(data, bytes):
                    data = bytes(data)

                original = msg_attr(
                    attachment,
                    "longFilename",
                    "shortFilename",
                    "filename",
                )

                original = original or f"attachment_{index}"

                attachments.append(
                    {
                        "filename": safe_filename(original),
                        "data": data,
                        "content_type": "",
                    }
                )

            except Exception:
                continue

        body = msg_attr(msg, "body")
        html_body = msg_attr(msg, "htmlBody")

        if isinstance(body, bytes):
            body = body.decode(
                "utf-8",
                errors="replace",
            )

        if isinstance(html_body, bytes):
            html_body = html_body.decode(
                "utf-8",
                errors="replace",
            )

        metadata = {
            "Format": "MSG",
            "Subject": msg_attr(msg, "subject"),
            "From": msg_attr(msg, "sender", "senderName"),
            "Sender email": msg_attr(msg, "senderEmail"),
            "To": msg_attr(msg, "to"),
            "CC": msg_attr(msg, "cc"),
            "BCC": msg_attr(msg, "bcc"),
            "Date": msg_attr(msg, "date"),
            "Message-ID": msg_attr(msg, "messageId"),
            "Importance": msg_attr(msg, "importance"),
            "Sensitivity": msg_attr(msg, "sensitivity"),
            "Attachment count": len(attachments),
        }

        return {
            "metadata": metadata,
            "recipients": recipients,
            "attachments": attachments,
            "body_text": body or "",
            "body_html": html_body or "",
        }

    finally:
        try:
            msg.close()
        except Exception:
            pass


# ---------- EML extraction ----------

def part_filename(
    part,
    index: int,
) -> str:
    filename = part.get_filename()

    if filename:
        filename = decode_text(filename)
    else:
        ext = part.get_content_subtype() or "bin"
        filename = f"attachment_{index}.{ext}"

    return safe_filename(filename)


def extract_eml_data(raw: bytes) -> dict:
    message = email.message_from_bytes(
        raw,
        policy=policy.default,
    )

    recipients = []

    for header, kind in (
        ("To", "TO"),
        ("Cc", "CC"),
        ("Bcc", "BCC"),
    ):
        recipients.extend(
            parse_address_header(
                str(message.get(header, "")),
                kind,
            )
        )

    attachments = []
    body_text_parts = []
    body_html_parts = []
    attachment_index = 0

    parts = (
        message.walk()
        if message.is_multipart()
        else [message]
    )

    for part in parts:
        if part.is_multipart():
            continue

        disposition = (
            part.get_content_disposition() or ""
        ).lower()

        filename = part.get_filename()

        if disposition == "attachment" or filename:
            attachment_index += 1

            try:
                payload = part.get_payload(decode=True) or b""

                attachments.append(
                    {
                        "filename": part_filename(
                            part,
                            attachment_index,
                        ),
                        "data": payload,
                        "content_type": part.get_content_type(),
                    }
                )

            except Exception:
                continue

            continue

        content_type = part.get_content_type()

        try:
            content = part.get_content()

            if isinstance(content, bytes):
                charset = (
                    part.get_content_charset() or "utf-8"
                )

                content = content.decode(
                    charset,
                    errors="replace",
                )

            if content_type == "text/plain" and content:
                body_text_parts.append(str(content))

            elif content_type == "text/html" and content:
                body_html_parts.append(str(content))

        except Exception:
            payload = part.get_payload(decode=True)

            if payload:
                charset = (
                    part.get_content_charset() or "utf-8"
                )

                decoded = payload.decode(
                    charset,
                    errors="replace",
                )

                if content_type == "text/html":
                    body_html_parts.append(decoded)
                else:
                    body_text_parts.append(decoded)

    date_value = str(
        message.get("Date", "") or ""
    )

    try:
        parsed_date = parsedate_to_datetime(date_value)

        date_iso = (
            parsed_date.isoformat()
            if parsed_date
            else date_value
        )

    except Exception:
        date_iso = date_value

    metadata = {
        "Format": "EML",
        "Subject": decode_text(
            message.get("Subject", "")
        ),
        "From": decode_text(
            message.get("From", "")
        ),
        "Sender email": "",
        "To": decode_text(
            message.get("To", "")
        ),
        "CC": decode_text(
            message.get("Cc", "")
        ),
        "BCC": decode_text(
            message.get("Bcc", "")
        ),
        "Date": date_iso,
        "Message-ID": str(
            message.get("Message-ID", "") or ""
        ),
        "Importance": str(
            message.get("Importance", "")
            or message.get("X-Priority", "")
            or ""
        ),
        "Sensitivity": str(
            message.get("Sensitivity", "") or ""
        ),
        "Attachment count": len(attachments),
    }

    return {
        "metadata": metadata,
        "recipients": recipients,
        "attachments": attachments,
        "body_text": "\n\n".join(
            body_text_parts
        ).strip(),
        "body_html": "\n\n".join(
            body_html_parts
        ).strip(),
    }


# ---------- Body processing ----------

def html_to_text(value: str) -> str:
    if not value:
        return ""

    value = re.sub(
        r"(?is)<(script|style).*?>.*?</\1>",
        " ",
        value,
    )

    value = re.sub(
        r"(?i)<br\s*/?>|</p>|</div>|</li>|</tr>",
        "\n",
        value,
    )

    value = re.sub(
        r"(?s)<[^>]+>",
        " ",
        value,
    )

    value = html.unescape(value)
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n\s*\n\s*\n+", "\n\n", value)

    return value.strip()


def best_body(data: dict) -> str:
    return (
        data["body_text"].strip()
        or html_to_text(data["body_html"])
    )


# ---------- Excel generation ----------

def build_recipients_xlsx(
    recipients: list[dict],
    original_name: str,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Recipients"

    ws.append([
        "Nama",
        "Email",
        "CC/BCC",
    ])

    for item in recipients:
        ws.append(
            [
                clean_excel_text(item.get("Nama", "")),
                clean_excel_text(item.get("Email", "")),
                clean_excel_text(item.get("CC/BCC", "")),
            ]
        )

    header_fill = PatternFill(
        "solid",
        fgColor="173B57",
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(
            color="FFFFFF",
            bold=True,
        )
        cell.alignment = Alignment(
            vertical="center",
        )

    ws.row_dimensions[1].height = 25
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 14
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    if ws.max_row > 1:
        table = Table(
            displayName="EmailRecipients",
            ref=ws.dimensions,
        )

        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        ws.add_table(table)

    output = io.BytesIO()
    wb.save(output)

    return output.getvalue()


def build_metadata_xlsx(
    metadata: dict,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Metadata"

    ws.append([
        "Field",
        "Value",
    ])

    for key, value in metadata.items():
        ws.append(
            [
                clean_excel_text(key),
                clean_excel_text(value),
            ]
        )

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid",
            fgColor="173B57",
        )
        cell.font = Font(
            color="FFFFFF",
            bold=True,
        )

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90
    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)

    return output.getvalue()


# ---------- Markdown generation ----------

def build_markdown(
    data: dict,
    source_name: str,
) -> str:
    meta = data["metadata"]
    recipients = data["recipients"]
    body = best_body(data)

    lines = [
        f"# Email: {meta.get('Subject') or '(No subject)'}",
        "",
        f"- **Source file:** {source_name}",
        f"- **Format:** {meta.get('Format', '')}",
        f"- **From:** {meta.get('From', '')}",
        f"- **Date:** {meta.get('Date', '')}",
        f"- **Message-ID:** {meta.get('Message-ID', '')}",
        "",
        "## Recipients",
    ]

    if recipients:
        lines += [
            "| Type | Name | Email |",
            "|---|---|---|",
        ]

        for item in recipients:
            vals = [
                item.get("CC/BCC", ""),
                item.get("Nama", ""),
                item.get("Email", ""),
            ]

            vals = [
                str(value)
                .replace("|", "\\|")
                .replace("\n", " ")
                for value in vals
            ]

            lines.append(
                "| " + " | ".join(vals) + " |"
            )

    else:
        lines.append(
            "_No recipients were detected._"
        )

    lines += [
        "",
        "## Attachments",
    ]

    if data["attachments"]:
        for item in data["attachments"]:
            lines.append(
                f"- {item['filename']} "
                f"({len(item['data']):,} bytes)"
            )

    else:
        lines.append("_No attachments._")

    lines += [
        "",
        "## Email body",
        "",
        body or "_No readable body content._",
        "",
    ]

    return "\n".join(lines)


# ---------- Attachment ZIP generation ----------

def make_attachment_archives(
    attachments: list[dict],
) -> dict[str, bytes]:
    grouped = defaultdict(list)

    for item in attachments:
        ext = (
            Path(item["filename"])
            .suffix.lower()
            .lstrip(".")
            or "no_extension"
        )

        grouped[ext].append(item)

    archives = {}

    for ext, items in grouped.items():
        buffer = io.BytesIO()

        with zipfile.ZipFile(
            buffer,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as zf:
            used = set()

            for idx, item in enumerate(
                items,
                start=1,
            ):
                filename = item["filename"]
                candidate = filename

                if candidate in used:
                    stem = Path(filename).stem
                    suffix = Path(filename).suffix

                    candidate = (
                        f"{stem}_{idx}{suffix}"
                    )

                used.add(candidate)

                zf.writestr(
                    candidate,
                    item["data"],
                )

        archives[ext] = buffer.getvalue()

    return archives


# ---------- Upload processing ----------

def process_upload(uploaded_file) -> dict:
    raw = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".msg":
        result = extract_msg_data(raw)

    elif suffix == ".eml":
        result = extract_eml_data(raw)

    else:
        raise ValueError(
            "Unsupported file type. "
            "Please upload a .msg or .eml file."
        )

    result["source_name"] = uploaded_file.name

    result["source_sha256"] = hashlib.sha256(
        raw
    ).hexdigest()

    return result


# ---------- Session state ----------

if "email_data" not in st.session_state:
    st.session_state.email_data = None

if "source_name" not in st.session_state:
    st.session_state.source_name = None


# ---------- Main UI ----------

# Hero header
st.html("""
<div class="hero">
    <h1>✉️ Mail Parser</h1>
    <p>
        A simple workspace to extract email recipients, attachments,
        body content, metadata, and AI-friendly Markdown.
    </p>
</div>
""")

# Upload section
st.subheader("1. Upload email file")

uploaded = st.file_uploader(
    "Choose an email file",
    type=["msg", "eml"],
    help=(
        "Upload one .msg or .eml file. Maximum size depends on "
        "your Streamlit deployment configuration."
    ),
)

left, right = st.columns([1, 2])

with left:
    process_clicked = st.button(
        "⚙️ Process email",
        type="primary",
        use_container_width=True,
        disabled=uploaded is None,
    )

with right:
    if uploaded:
        st.info(
            f"Selected: **{uploaded.name}** · "
            f"{uploaded.size:,} bytes"
        )
    else:
        st.caption(
            "Upload a file to enable processing."
        )

if process_clicked and uploaded is not None:
    with st.spinner(
        "Reading email and extracting its contents..."
    ):
        try:
            st.session_state.email_data = process_upload(
                uploaded
            )

            st.session_state.source_name = uploaded.name

            st.success(
                "Email processed successfully."
            )

        except Exception as exc:
            st.session_state.email_data = None
            st.session_state.source_name = None

            st.error(
                f"Could not process this email: {exc}"
            )


# ---------- Extraction workspace ----------

data = st.session_state.email_data

if data:
    metadata = data["metadata"]
    recipients = data["recipients"]
    attachments = data["attachments"]
    body = best_body(data)

    st.divider()
    st.subheader("2. Extraction workspace")

    st.caption(
        f"Processed file: **{data['source_name']}**"
    )

    counts = Counter(
        item["CC/BCC"]
        for item in recipients
    )

    email_format = html.escape(
        str(metadata.get("Format", "—"))
    )

    # ---------- Metric cards ----------

    metric_html = f"""
    <div class="metric-grid">
        <div class="metric-card recipients">
            <div class="metric-label">👥 Recipients</div>
            <div class="metric-value">{len(recipients):,}</div>
            <div class="metric-description">
                Total detected recipients
            </div>
        </div>

        <div class="metric-card attachments">
            <div class="metric-label">📎 Attachments</div>
            <div class="metric-value">{len(attachments):,}</div>
            <div class="metric-description">
                Extracted email attachments
            </div>
        </div>

        <div class="metric-card body">
            <div class="metric-label">📝 Body Characters</div>
            <div class="metric-value">{len(body):,}</div>
            <div class="metric-description">
                Total readable text characters
            </div>
        </div>

        <div class="metric-card format">
            <div class="metric-label">✉️ Email Format</div>
            <div class="metric-value">{email_format}</div>
            <div class="metric-description">
                Detected source file format
            </div>
        </div>
    </div>
    """

    st.html(
        textwrap.dedent(metric_html)
    )

    # ---------- Extraction tabs ----------

    (
        tab_recipients,
        tab_attachments,
        tab_body,
        tab_metadata,
        tab_markdown,
    ) = st.tabs(
        [
            "👥 Recipients",
            "📎 Attachments",
            "📝 Body",
            "ℹ️ Metadata",
            "✨ AI Markdown",
        ]
    )

    # ---------- Recipients ----------

    with tab_recipients:
        st.markdown("#### Email recipients")

        col_to, col_cc, col_bcc = st.columns(3)

        col_to.metric(
            "TO",
            counts.get("TO", 0),
        )

        col_cc.metric(
            "CC",
            counts.get("CC", 0),
        )

        col_bcc.metric(
            "BCC",
            counts.get("BCC", 0),
        )

        if recipients:
            st.dataframe(
                recipients,
                use_container_width=True,
                hide_index=True,
            )

            xlsx = build_recipients_xlsx(
                recipients,
                data["source_name"],
            )

            stem = Path(
                data["source_name"]
            ).stem

            st.download_button(
                "⬇️ Download recipient Excel",
                data=xlsx,
                file_name=(
                    f"extracted_receipient-{stem}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
            )

        else:
            st.warning(
                "No recipients were detected in this file. "
                "Some received messages do not retain BCC information."
            )

    # ---------- Attachments ----------

    with tab_attachments:
        st.markdown("#### Extracted attachments")

        if not attachments:
            st.info(
                "This email does not contain extractable attachments."
            )

        else:
            groups = defaultdict(list)

            for item in attachments:
                ext = (
                    Path(item["filename"])
                    .suffix.lower()
                    .lstrip(".")
                    or "no_extension"
                )

                groups[ext].append(item)

            st.write(
                f"Found **{len(attachments)} attachment(s)** "
                f"across **{len(groups)} format group(s)**."
            )

            archives = make_attachment_archives(
                attachments
            )

            for ext, items in sorted(groups.items()):
                with st.expander(
                    f".{ext} · {len(items)} file(s)",
                    expanded=True,
                ):
                    for idx, item in enumerate(
                        items,
                        start=1,
                    ):
                        col_file, col_download = st.columns(
                            [3, 1]
                        )

                        with col_file:
                            st.write(
                                f"**{item['filename']}**"
                            )

                            st.caption(
                                f"{len(item['data']):,} bytes · "
                                f"{item.get('content_type') or 'Unknown content type'}"
                            )

                        with col_download:
                            st.download_button(
                                "Download",
                                data=item["data"],
                                file_name=item["filename"],
                                key=(
                                    f"att_{ext}_{idx}_"
                                    f"{hashlib.md5(item['data']).hexdigest()[:8]}"
                                ),
                                use_container_width=True,
                            )

                    st.download_button(
                        f"Download all .{ext} as ZIP",
                        data=archives[ext],
                        file_name=(
                            f"attachments-"
                            f"{Path(data['source_name']).stem}-"
                            f"{ext}.zip"
                        ),
                        mime="application/zip",
                        key=f"zip_{ext}",
                    )

            all_zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                all_zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as zf:
                for ext, archive_data in archives.items():
                    zf.writestr(
                        f"{ext}/attachments-{ext}.zip",
                        archive_data,
                    )

            st.download_button(
                "⬇️ Download all attachment groups (ZIP)",
                data=all_zip_buffer.getvalue(),
                file_name=(
                    f"all-attachments-"
                    f"{Path(data['source_name']).stem}.zip"
                ),
                mime="application/zip",
                type="primary",
            )

    # ---------- Body ----------

    with tab_body:
        st.markdown("#### Email body")

        if body:
            st.text_area(
                "Extracted body text",
                value=body,
                height=350,
            )

            st.download_button(
                "⬇️ Download body as TXT",
                data=body.encode("utf-8"),
                file_name=(
                    f"email-body-"
                    f"{Path(data['source_name']).stem}.txt"
                ),
                mime="text/plain",
            )

        else:
            st.info(
                "No readable text body was found."
            )

    # ---------- Metadata ----------

    with tab_metadata:
        st.markdown("#### Email metadata")

        st.dataframe(
            [
                {
                    "Field": key,
                    "Value": value,
                }
                for key, value in metadata.items()
            ],
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Download metadata as Excel",
            data=build_metadata_xlsx(metadata),
            file_name=(
                f"email-metadata-"
                f"{Path(data['source_name']).stem}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        st.caption(
            f"SHA-256: `{data['source_sha256']}`"
        )

    # ---------- AI Markdown ----------

    with tab_markdown:
        st.markdown("#### AI-friendly Markdown")

        markdown_text = build_markdown(
            data,
            data["source_name"],
        )

        st.caption(
            "Combines key headers, recipient list, attachment inventory, "
            "and readable body text into one structured document."
        )

        st.text_area(
            "Markdown preview",
            value=markdown_text,
            height=430,
        )

        st.download_button(
            "⬇️ Download Markdown (.md)",
            data=markdown_text.encode("utf-8"),
            file_name=(
                f"email-context-"
                f"{Path(data['source_name']).stem}.md"
            ),
            mime="text/markdown",
            type="primary",
        )

    st.divider()

    if st.button(
        "Clear current email and start over"
    ):
        st.session_state.email_data = None
        st.session_state.source_name = None
        st.rerun()


# ---------- Empty state ----------

else:
    st.markdown("### What you can do")

    col_recipients, col_attachments, col_ai = st.columns(3)

    with col_recipients:
        st.markdown("#### 👥 Recipients")

        st.write(
            "Extract TO, CC, and BCC recipients into a clean Excel workbook."
        )

    with col_attachments:
        st.markdown("#### 📎 Attachments")

        st.write(
            "Group files by extension and download individually "
            "or in ZIP packages."
        )

    with col_ai:
        st.markdown("#### ✨ AI-ready output")

        st.write(
            "Convert email context into a structured Markdown file "
            "for efficient AI analysis."
        )

    st.divider()
    st.markdown("**Supported formats**")

    format_col_1, format_col_2 = st.columns(2)

    with format_col_1:
        st.markdown("- Outlook `.msg`")

    with format_col_2:
        st.markdown("- RFC 822 `.eml`")

    st.caption(
        "Files are processed in the active app session. Avoid uploading "
        "confidential emails to a public deployment unless your "
        "organization's data policy permits it."
    )


# ---------- Footer ----------

st.html("""
<div class="app-footer">
    Mail Parser · Extract • Organize • Convert<br>
    Developed by <strong>Dewan Dinata</strong>
</div>
""")
