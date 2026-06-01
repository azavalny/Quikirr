# Waterfall Analysis Email Bot

A FastAPI service that acts as an email bot. It receives inbound emails via a
SendGrid Inbound Parse webhook, validates the Excel attachment, runs the
`quikirr` waterfall engine on it, and replies with the output workbook. Invalid
emails get a usage guide instead.

## Expected input

The attached Excel file (`.xlsx` or `.xls`) must contain a sheet named
`Source Data` with:

- a `Customer` column, and
- monthly date-headed columns of MRR values (at least 3 month columns).

The bot replies with the full multi-tab waterfall workbook
(`Annual Waterfall`, `Quarterly Waterfall`, etc.).

## Local development

```bash
pip install -r email-bot/requirements.txt
cp email-bot/.env.example email-bot/.env   # then fill in the values
```

Set the environment variables (or load the `.env`):

```
SENDGRID_API_KEY=SG.xxxxx
FROM_EMAIL=analyze@bot.alexzavalny.com
```

Run the server from the repository root (so the `quikirr` package is
importable):

```bash
uvicorn main:app --app-dir email-bot --reload --host 0.0.0.0 --port 8000
```

The endpoint is then available at `http://localhost:8000/inbound`.

## Pointing SendGrid at the webhook

1. Deploy the service somewhere publicly reachable (e.g. Railway uses the
   included `Procfile`).
2. In SendGrid, go to **Settings -> Inbound Parse** and add a host/URL.
3. Set the destination URL to `https://<your-host>/inbound`.
4. Configure the MX records for the receiving domain to point at SendGrid as
   instructed there.

SendGrid will then POST inbound emails as `multipart/form-data` to `/inbound`.
The handler always returns HTTP 200 so SendGrid does not retry.
