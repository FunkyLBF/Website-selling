# LeadScope Sales Pipeline

A small, provider-independent enrichment pipeline for Giantic Studios.

## What it does

`162 active Berlin tattoo leads -> public website/contact discovery -> website audit -> sales-opportunity score -> personalized draft`

It reads the existing Excel/CSV export and adds `ls_*` fields. It does **not** send email.

### Outputs

- public emails found on the business website/contact/impressum pages
- HTTPS
- mobile viewport signal
- title/meta description
- booking/appointment signal
- CTA signal
- contact page signal
- social links
- portfolio/image signal
- concise audit notes
- 0–100 sales opportunity score
- A/B/C/SKIP tier
- recommended pitch angle
- personalized subject + email draft

## Run

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python pipeline.py /path/to/berlin_tattoo_studios_2026_active_leads.xlsx \
  -o berlin_tattoo_studios_leadscope_enriched.xlsx \
  --workers 8
```

For a safe test first:

```bash
python pipeline.py input.xlsx -o test.xlsx --limit 10 --workers 4
```

## Important design choice

The pipeline is intentionally **research + draft only**. There is no SMTP/SendGrid/Mailgun/Resend send step. Outreach should remain human-approved, especially for unsolicited commercial email.

## Next integration into the existing LeadScope app

Move these functions into the current Next.js app as server-side jobs:

- `analyze_site()` -> WebsiteAnalysis worker
- `score()` -> new OpportunityScore service
- email discovery -> ContactDiscovery service
- `draft()` -> OutreachDraft generator
- add `email`, `emailSource`, `opportunityScore`, `opportunityTier`, `pitchAngle`, `auditNotes`, and `draftStatus` to Prisma
- UI: Pipeline > Enrich Leads > Review > Approve Draft > Copy/Send manually

Keep the existing LeadScope architecture. Do not replace the current scoring model; this creates a separate **sales opportunity score** so website deficiency and commercial opportunity remain distinct.
