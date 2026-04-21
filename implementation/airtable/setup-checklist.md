# Airtable Setup Checklist

## Base creation

1. Create a new base named `Resham Sutra Phase 1`
2. Create tables from `base-schema.json`
3. Create views:
   - New Enquiries
   - Needs Review
   - Ready for Draft
   - Draft Quotations
   - Approved to Send
   - Sent Quotations
   - Follow-Ups Due
   - Converted Orders

## Import seed data

1. Import [quotation-templates.csv](C:\Users\pawan\Documents\Codex\ReshmaSutra\Code\implementation\airtable\seeds\quotation-templates.csv) into `Quotation Templates`
2. Import [products-initial.csv](C:\Users\pawan\Documents\Codex\ReshmaSutra\Code\implementation\airtable\seeds\products-initial.csv) into `Products`
3. Keep [product-aliases.csv](C:\Users\pawan\Documents\Codex\ReshmaSutra\Code\implementation\airtable\seeds\product-aliases.csv) as parser reference or import it into a separate alias helper table if needed

## Action fields

Ensure the following checkbox fields exist and are visible in the interface:

- `Send Quotation`
- `Send Reminder`
- `Mark Accepted`
- `Mark Rejected`

## Interface pages

Recommended first interface pages:

- Enquiry Inbox
- Quotation Review
- Send Queue
- Reminder Queue
- Order Conversion

## Formula / automation notes

- `Quotation Number` can be generated in n8n using the chosen template prefix
- `Document Jobs` should remain visible to the team so they can see draft/PDF/send failures
- `Draft Format` should default to `XLSX` unless a DOCX-first template is later preferred
