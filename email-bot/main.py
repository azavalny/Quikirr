from __future__ import annotations

import base64
import logging
import os
import re
from pathlib import Path

from fastapi import FastAPI, Request
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    FileContent,
    FileName,
    FileType,
    Mail,
)
from dotenv import load_dotenv
load_dotenv()

from analysis import run_waterfall_analysis

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("email-bot")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

EXAMPLE_FILE = Path(__file__).resolve().parent / "Source.xlsx"
EXAMPLE_FILENAME = "Source.xlsx"

USAGE_SUBJECT = "How to use the Waterfall Analysis Bot"
USAGE_BODY = """Hi,

Send this address an email with ONE Excel file (.xlsx or .xls) attached. It must contain a sheet named "Source Data" with a "Customer" column and at least 3 monthly date-headed columns of MRR values.

I've attached Source.xlsx as an example of the expected format. Reply with a file like it and you'll get back the full waterfall analysis.
"""

RESULT_SUBJECT = "Waterfall Analysis — Results"
RESULT_BODY = """Hi,

Your waterfall analysis is complete. Please find the results attached.

If the output looks incorrect, double-check that your input file matches the expected format described in the usage guide.

To re-run the analysis, simply reply with an updated Excel file.
"""

app = FastAPI()


def _parse_email(raw: str) -> str:
    match = re.search(r"<([^>]+)>", raw or "")
    if match:
        return match.group(1).strip()
    return (raw or "").strip()


def send_email(
    to: str,
    subject: str,
    body: str,
    attachment_bytes: bytes | None = None,
    attachment_filename: str | None = None,
) -> None:
    message = Mail(
        from_email=os.environ["FROM_EMAIL"],
        to_emails=to,
        subject=subject,
        plain_text_content=body,
    )

    if attachment_bytes is not None and attachment_filename is not None:
        encoded = base64.b64encode(attachment_bytes).decode()
        message.attachment = Attachment(
            FileContent(encoded),
            FileName(attachment_filename),
            FileType(XLSX_MIME),
            Disposition("attachment"),
        )

    client = SendGridAPIClient(os.environ["SENDGRID_API_KEY"])
    client.send(message)


def send_usage(to: str) -> None:
    try:
        example = EXAMPLE_FILE.read_bytes()
    except OSError:
        example = None
    if example is not None:
        send_email(to, USAGE_SUBJECT, USAGE_BODY, example, EXAMPLE_FILENAME)
    else:
        send_email(to, USAGE_SUBJECT, USAGE_BODY)


@app.post("/inbound")
async def inbound(request: Request):
    try:
        form = await request.form()

        sender = _parse_email(form.get("from", ""))
        try:
            attachments = int(form.get("attachments", 0) or 0)
        except (TypeError, ValueError):
            attachments = 0

        if attachments == 0 or attachments > 1:
            send_usage(sender)
            return {"status": "invalid"}

        upload = form.get("attachment1")
        filename = (getattr(upload, "filename", "") or "").lower()
        if not filename.endswith((".xlsx", ".xls")):
            send_usage(sender)
            return {"status": "invalid"}

        file_bytes = await upload.read()

        try:
            output_bytes = run_waterfall_analysis(file_bytes)
        except (KeyError, ValueError) as exc:
            logger.info("Invalid input file from %s: %s", sender, exc)
            del file_bytes
            send_usage(sender)
            return {"status": "invalid"}

        send_email(
            sender,
            RESULT_SUBJECT,
            RESULT_BODY,
            attachment_bytes=output_bytes,
            attachment_filename="waterfall_output.xlsx",
        )

        del file_bytes
        del output_bytes
        return {"status": "ok"}
    except Exception:
        logger.exception("Unexpected error handling inbound email")
        return {"status": "error"}
