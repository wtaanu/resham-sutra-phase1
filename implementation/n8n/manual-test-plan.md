# Manual Entry Test Plan

Use this plan until WhatsApp onboarding is complete.

## Goal

Test the full operational flow starting from a manually entered Airtable enquiry:

1. create enquiry
2. create quotation draft
3. generate draft document
4. approve quotation
5. generate final PDF
6. save files in Google Drive
7. send email quotation

## Preconditions

- Airtable base is ready
- Google Drive credentials are working
- SMTP test credentials are working
- WhatsApp nodes remain disabled or skipped

## Test record setup

### 1. Customer

Create one `Customers` row with:

- `Client ID`: `CUST-001`
- `Customer Name`: `Test Buyer`
- `Company`: `Test Silk Works`
- `Email`: your test recipient email
- `Customer Type`: `Domestic`

Expected folder name:

- `CUST-001-Test Buyer`

### 2. Enquiry

Create one `Enquiries` row with:

- `Source Channel`: `Manual Entry`
- `Parser Status`: `Ready for Draft`
- `Lead Name`: `Test Buyer`
- `Company`: `Test Silk Works`
- `Phone`: a test phone number
- `Email`: same test recipient email
- `Location`: `Nagpur`
- `Requirement Summary`: `Mulberry reeling machine quotation required`
- `Requested Asset`: `Quotation`
- link to the customer row

### 3. Quotation

After workflow 2 runs, confirm:

- `Quotations` row is created
- `Template` is set to domestic standard
- `Draft Format` is set
- `Document Jobs` row is queued

### 4. Line items

Add 1-2 `Quotation Line Items` manually if workflow 2 does not yet auto-populate them.

Recommended:

- `MRM-010`
- `BRM-01`

## Approval test

Update the quotation:

- `Status`: `Approved`
- `Send Quotation`: checked
- `Preferred Send Channel`: `Email`

Expected result:

- draft document created if missing
- final PDF generated
- files uploaded to Drive under:
  - `<Client ID>-<Client Name>/quotation.xlsx`
  - `<Client ID>-<Client Name>/quotation.pdf`
- email sent from the configured SMTP sender

## Reminder test

After send:

- set `Status = Sent`
- set `Next Reminder Date = today`
- check `Send Reminder`

Expected result:

- reminder email sent
- reminder counters updated

## Notes

- During test mode, `sales@reshamsutra.com` should be the business sender identity, but temporary SMTP credentials may still be used underneath until client credentials are swapped in.
- Once WhatsApp goes live, only workflow 1 input changes. The rest of the Airtable -> document -> Drive -> send pipeline remains the same.
