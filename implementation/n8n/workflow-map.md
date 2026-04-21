# n8n Workflow Map

## Workflow 1: Inbound enquiry capture

File: `workflows/01-inbound-enquiry-capture.json`

### Trigger

- Meta WhatsApp webhook receives forwarded message

### Main outputs

- create/update `Enquiries`
- create/update `Customers`
- create lead folder in Google Drive
- upsert Zoho Bigin contact/deal

### Notes

- Sender in webhook is the internal employee, not the actual lead
- parser must extract lead phone/name/company from message body
- if phone is missing, mark Airtable `Parser Status = Needs Review`

### Manual test mode

Until Meta WhatsApp is available, create enquiry rows manually in Airtable using:

- `Source Channel = Manual Entry`
- `Parser Status = Ready for Draft`

This lets us test the rest of the pipeline without waiting on WhatsApp onboarding.

## Workflow 2: Create quotation draft

File: `workflows/02-create-quotation-draft.json`

### Trigger

- Airtable `Enquiries` record updated to `Ready for Draft`

### Main outputs

- create `Quotations` row
- create `Quotation Line Items`
- create `Document Jobs` row for draft generation

### Template rule

- choose `domestic-standard` by default
- choose `myanmar-proforma` when customer type or destination requires export/proforma flow

### Draft rule

- quotation should first exist as draft `XLSX` or `DOCX`
- no PDF send before manual approval
- current API implementation supports real `XLSX` draft generation
- final PDF generation remains a separate post-approval step

## Workflow 3: Approve, generate PDF, and send

File: `workflows/03-approve-generate-pdf-send.json`

### Trigger

- Airtable `Quotations.Send Quotation = true`

### Main outputs

- build payload from quotation header + line items
- generate draft file if missing
- generate final PDF
- upload to Drive
- send on WhatsApp when the Meta API is live
- send on email during test mode using `sales@reshamsutra.com` as the business sender identity
- update status to `Sent`

### Field updates expected

- `Status` -> `Sent`
- `Draft File URL`
- `Final PDF URL`
- `Sent Date`
- `Send Quotation` reset to false

## Workflow 4: Reminders and conversion

File: `workflows/04-reminders-and-conversion.json`

### Trigger A

- daily schedule at 10:00

### Trigger B

- Airtable `Quotations.Send Reminder = true`

### Main outputs

- send reminder on WhatsApp
- optionally send reminder email
- update reminder counters

### Conversion path

- Airtable `Quotations.Mark Accepted = true`
- create `Orders` record
- update Zoho Bigin deal stage

## Recommended implementation order

1. connect Airtable credentials
2. connect Google Drive
3. wire SMTP and verify outbound email from `sales@reshamsutra.com`
4. test manual Airtable enquiry entry end to end
5. test workflow 2 with domestic template
6. connect draft document generation
7. connect final PDF generation
8. enable workflow 3 in email-only mode
9. wire Meta webhook and send endpoint later
10. wire Zoho Bigin upsert/update endpoints
11. enable reminders and conversion
