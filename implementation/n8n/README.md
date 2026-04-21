# n8n Workflow Skeletons

These JSON files are starter workflow skeletons for self-hosted n8n.

They are organized by the actual Phase 1 lifecycle:

1. `01-inbound-enquiry-capture.json`
2. `02-create-quotation-draft.json`
3. `03-approve-generate-pdf-send.json`
4. `04-reminders-and-conversion.json`

## Intended sequence

Until WhatsApp is live, the recommended test mode is:

- create enquiries manually in Airtable
- move them to `Ready for Draft`
- test quotation generation, Drive save, and email sending first
- keep WhatsApp nodes disabled or mocked

### 1. Inbound enquiry capture

- Triggered by Meta WhatsApp webhook
- Normalizes forwarded messages
- Extracts lead details
- Writes enquiry and customer data to Airtable
- Syncs contact/deal basics to Zoho Bigin
- Creates lead-wise folder in Google Drive

### 2. Create quotation draft

- Runs once an enquiry is marked `Ready for Draft`
- Chooses domestic or Myanmar template
- Creates quotation header and line-item records in Airtable
- Starts draft document generation job

### 3. Approve, generate PDF, and send

- Watches Airtable for `Send Quotation = true`
- Creates final XLSX or DOCX draft if missing
- Generates approved PDF
- Saves all files in Google Drive
- Sends WhatsApp first when live, otherwise email-only during testing
- Updates Airtable and Zoho

### 4. Reminders and conversion

- Daily scheduled reminder scan
- Manual reminder trigger support
- Accepted quotation creates order
- Rejected quotation updates pipeline and logs reason

## Notes

- Replace all placeholder node credentials before import/use.
- The draft document endpoint now produces a real `.xlsx` workbook plus preview HTML. The final PDF endpoint is still a placeholder HTML/PDF-source artifact until the PDF renderer is implemented.
- The `Document Jobs` Airtable table is used to keep file-generation status visible to the team.
