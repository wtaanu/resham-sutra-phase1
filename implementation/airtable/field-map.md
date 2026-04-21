# Airtable Field Map

## Workbook-driven design rules

The schema is designed around the shared client workbooks:

- `Products List.xlsx` has multiple relevant sheets:
  - `pricelist`
  - `Sheet1`
  - `Sheet2`
  - `RSPL`
- `Quotation format.xlsx` has two output templates:
  - `jha (2)` for domestic quotations
  - `MYANMAR` for export/proforma invoices

This means the Airtable base must preserve:

- source sheet on each product row
- template selection on each quotation
- draft document workflow before final PDF generation

## Product workbook mapping

### `pricelist`, `Sheet1`, `RSPL`

Mapped into the `Products` table:

- `Category` -> `Category`
- `Model` -> `Model`
- `Narration` -> `Narration`
- `Variant` -> `Variant`
- `Purchase Cost` -> `Purchase Cost`
- `Ex-factory` / `Ex-factory Price` -> `Ex-factory`
- `Freight` -> `Freight`
- `GST` -> `GST Amount`
- `Bulk Sale Price` -> `Bulk Sale Price`
- `MRP` -> `MRP`
- `Supplier` -> `Supplier`
- sheet name -> `Source Sheet`

### `Sheet2`

Used as alias/reference support:

- model family names and shorthand names -> `Alias Keywords`

This helps the parser match messages like:

- `Mulberry reeling machine`
- `Buniyaad reeling machine`
- `Sonalika`

to normalized products in Airtable.

## Quotation workbook mapping

### Domestic template: `jha (2)`

Mapped via `Quotation Templates` row:

- `Template Code` -> `domestic-standard`
- `Document Kind` -> `Domestic quotation`
- `Number Prefix` -> `QO/25-26/`

Header fields move into `Quotations`:

- Buyer block -> `Buyer Block`
- No.# -> `Quotation Number` / `Reference Number`
- Date -> document date / sent date
- Ref -> `Reference Number`

Line columns move into `Quotation Line Items`:

- `S.No.` -> `Line No.`
- `Description` -> `Description Override`
- `Qty` -> `Qty`
- `Rate Per Unit` -> `Rate Per Unit`
- `Pkg & Trnsprt` -> `Pkg & Transport`
- `GST %` -> `GST %`
- `GST Amt` -> `GST Amount`
- `Total Amount` -> `Total Amount`

### Export template: `MYANMAR`

Mapped via `Quotation Templates` row:

- `Template Code` -> `myanmar-proforma`
- `Document Kind` -> `Export / Myanmar proforma`
- `Number Prefix` -> `QI/23-24/`

Extra fields needed:

- consignee block -> `Consignee Block`
- `Unit Value` column -> `Unit Value` on line items

## Document lifecycle

The flow must remain:

1. enquiry captured
2. quotation draft created
3. team edits draft in Airtable
4. draft XLSX or DOCX generated
5. user manually reviews and approves
6. final PDF generated from approved draft
7. PDF saved in Google Drive
8. PDF sent on WhatsApp and/or email

To support this, the `Quotations` table stores:

- `Draft Format`
- `Draft File URL`
- `Final PDF URL`
- `Send Quotation`
- `Send Reminder`
- `Mark Accepted`
- `Mark Rejected`

The `Document Jobs` table stores the actual automation job state for:

- draft XLSX
- draft DOCX
- final PDF
- WhatsApp send
- email send

## Immediate setup order

1. Create Airtable tables from `base-schema.json`
2. Load `Quotation Templates` with the 2 template rows
3. Import products from all 4 workbook sheets
4. Add alias keywords from `Sheet2`
5. Connect n8n triggers to the action fields in `Quotations`
