# Cook County Property Due Diligence v1.0

## What this is

A self-contained Windows Excel/VBA property-research tool.

**Workflow:** enter a Cook County PIN → run the research macro → review source status and issues → export one PDF.

There is **no website, server, or hosted API dependency**.

## Production files

- `Cook_Property_Due_Diligence_v1_0_Template.xlsx`
- `CookPropertyDueDiligence_v1_0.bas`

The template is intentionally `.xlsx` because the build environment used here cannot embed a VBA project. The one-time Windows setup below converts it into your production `.xlsm`.

## One-time Windows setup

1. Open `Cook_Property_Due_Diligence_v1_0_Template.xlsx`.
2. Use **Save As** and save a working copy as **Excel Macro-Enabled Workbook (*.xlsm)**.
3. Press **Alt+F11** to open the VBA Editor.
4. Choose **File > Import File** and import `CookPropertyDueDiligence_v1_0.bas`.
5. Choose **Debug > Compile VBAProject**.
6. If Excel reports no compile error, save the `.xlsm`.
7. If your firm blocks unsigned macros, use the firm's approved Trusted Location / macro-security process.

## Running a property

1. Go to **Start**.
2. Enter a Cook County PIN in the yellow cell. Hyphens and spaces are acceptable.
3. Press **Alt+F8**.
4. Run `GenerateCookPropertyReport`.
5. Watch **Source Status**. One failed source should not stop the remaining research.
6. Review **Issues / Items to Verify** before relying on the report.
7. The macro exports one combined PDF and, by default, opens it automatically.

You can later assign `GenerateCookPropertyReport` to a worksheet button.

## What the PDF includes

- Executive property summary
- Core parcel/property facts
- Assessment history
- Tax history, tax rate, payment status and exemption history
- Tax-sale/delinquency status surfaced by the Property Tax Portal
- Sales/deeds
- Recent recorded-document/deed/lien index
- Assessor and Board of Review appeal history
- TIF / Enterprise Zone checks
- Permit history
- Issues / items to verify
- Official source appendix

## Configuration

On the **Config** sheet:

- HTTP Timeout: 30 seconds per source by default
- PDF Output Folder: optional; blank uses the workbook folder, then Documents
- Open PDF After Export: YES / NO
- Source URLs are isolated so future endpoint changes can be patched without redesigning the workbook

## Required first production test

Before relying on v1.0 in client work:

1. In Windows Excel, import the module and run **Debug > Compile VBAProject**.
2. Run `16-30-204-020-0000`.
   - Useful smoke test for billed-tax history, first-installment treatment, exemptions and recent documents.
3. Run `14-33-323-023-0000`.
   - Useful smoke test for a current payment/balance condition and adverse recorded-document language.
4. Compare the workbook/PDF against the live Cook County Property Tax Portal and CookViewer.
5. Run at least one known commercial/industrial parcel and one known TIF parcel from your own work files.

## Important limitations

- **Enterprise Zone is deliberately PARTIAL.** The workbook uses the Assessor's Enterprise Zone signal when the expected field is available, but the official Illinois DCEO map remains the boundary-verification source until its authoritative statewide FeatureServer is positively resolved and tested.
- Assessor owner/mailing information can lag recorded title activity.
- Assessor permit data is not a complete universe of municipal permits.
- The Cook County Property Tax Portal does not include municipal special assessments or omitted taxes; verify those separately when relevant.
- Public systems can disagree. The macro is designed to flag conflicts rather than silently choose a value.
- v1.0 has been source/schema checked and statically reviewed here, but **has not been compiled or executed inside Windows Excel in this environment**. The Windows compile + first end-to-end run is the final production gate.
