# Changelog — v1.0

## Architecture

- Removed the website/server concept entirely.
- Production target is a self-contained Windows Excel/VBA workbook.
- Each public source is isolated behind an adapter with non-fatal error handling and a 30-second default timeout.

## GIS

- Replaced the older historical parcel layer as the primary GIS source.
- Added Cook County's direct current CookViewer parcel feature layer.
- Added current spatial municipality point-in-polygon verification.
- Added GIS class description, major class, land SF, building SF, construction, building age and current value fields.
- Added explicit Assessor-vs-GIS class conflict checking.

## Property Tax Portal

- Corrected the live PIN results route.
- Added PIN-specific results-page validation so the generic search page is not mistaken for property data.
- Added current composite tax rate and tax code.
- Added five-year billed-tax history.
- Added current payment/status parsing.
- Added first-installment labeling.
- Added exemption history.
- Added refund-section review flagging.
- Added tax-sale/delinquency status by year.
- Added recent documents/deeds/liens.
- Added issue escalation for balance-due conditions, lis pendens/foreclosure language and liens.
- Added lot/building SF cross-checks.

## TIF / Enterprise Zone

- Current Cook County TIF boundary is now the primary TIF signal.
- 2024 TIF detail/revenue layer is retained only as a labeled secondary cross-check.
- TIF layer disagreement generates an Issues entry.
- Enterprise Zone is intentionally displayed as PARTIAL until the official DCEO statewide boundary service is positively resolved/tested.
- Absence of an Assessor Enterprise Zone signal is not treated as a conclusive “No.”

## PDF / workbook

- Added a recent recorded-documents sheet.
- Added tax-sale/delinquency column to Tax History.
- Added a Sources appendix to the combined PDF.
- Added page headers/footers, PIN/address identification and page numbering.
- Added professional number formats.
- Added source-status conditional formatting.
- Added a self-contained Instructions sheet.
- Added explicit public-record limitations, including the Property Tax Portal's exclusion of municipal special assessments and omitted taxes.

## Reliability

- Added ArcGIS error-response validation.
- Reset logic clears old subject/report values between PINs.
- Source failures continue to the next adapter.
- Missing or conflicting information is flagged instead of guessed.
