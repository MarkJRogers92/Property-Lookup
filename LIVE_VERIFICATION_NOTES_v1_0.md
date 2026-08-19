# Live Verification Notes — Cook County Property Due Diligence v1.0

## Verification status

| Adapter | Status | Notes |
|---|---|---|
| Assessor Parcel Addresses | LIVE SCHEMA VERIFIED | Situs, owner/taxpayer and mailing-address fields reviewed against the current official dataset. |
| Assessor Parcel Universe | LIVE SCHEMA VERIFIED — CORE | Current class, township, neighborhood, coordinates and tax-code fields reviewed. Tax-code value remains secondary because Cook County warns it may not be current. |
| Enterprise Zone Assessor signal | PARTIAL | Code is schema-safe: if the expected signal field is absent, it logs a verification issue rather than inferring “No.” DCEO boundary verification remains manual. |
| Assessed Values | LIVE SCHEMA VERIFIED | Mailed, certified and Board values reviewed against current official field names. |
| Parcel Sales | LIVE SCHEMA VERIFIED | Sale/deed/document/buyer/seller/multi-PIN fields reviewed. |
| Assessor Appeals | LIVE SCHEMA VERIFIED | Current appeal case/type/status and pre/post assessment fields reviewed. |
| Board of Review | LIVE SCHEMA VERIFIED | BOR complaint identifiers, appeal type and Assessor/BOR total-value fields reviewed. |
| Permits | LIVE SCHEMA VERIFIED | Core permit fields reviewed; source itself is not a complete municipal permit universe. |
| CookViewer current parcel | LIVE ENDPOINT + SCHEMA VERIFIED | Direct `CookViewer3Parcels` feature layer is the primary GIS source. It exposes PIN, class, class descriptions, land/building SF, construction, neighborhood and current value fields. |
| Current municipality boundary | LIVE ENDPOINT VERIFIED | Point-in-polygon adapter uses Cook County's current municipality boundary layer and parcel centroid coordinates. |
| Property Tax Portal | LIVE CONTENT/PARSER PATTERNS VERIFIED | Current route and live page content were checked for tax rate/code, five-year billed history, current payment/status, exemptions, tax-sale/delinquency sections and recent documents/deeds/liens. |
| Current TIF boundary | LIVE ENDPOINT + SCHEMA VERIFIED | Current Cook County TIF polygon layer is the primary TIF signal. |
| 2024 TIF detail/revenue layer | VERIFIED AS SECONDARY | Used only as a second cross-check/name/detail source and clearly labeled 2024. |
| Illinois DCEO Enterprise Zone boundary | MANUAL / PARTIAL | Official DCEO interactive map is retained. Underlying authoritative statewide FeatureServer was not wired into v1.0 without positive identification/testing. |
| Windows VBA compile/runtime | NOT EXECUTED HERE | Must be compiled and run once in actual Windows Excel before production reliance. |

## Source-level regression PINs

### 16-30-204-020-0000

Useful for validating:

- PIN-specific Property Tax Portal results
- five-year billed history
- first-installment indicator handling
- exemption parsing
- recent documents/deeds/liens
- ordinary tax-sale status text

### 14-33-323-023-0000

Useful for validating:

- property characteristics
- a current payment/balance condition
- recorded-document index containing `LIS PENDENS FORECLOSURE`
- issue escalation for potentially adverse recorded-document language

## Recommended Windows regression set before declaring “production locked”

Run and record at least:

1. ordinary residential PIN
2. commercial/industrial PIN
3. multi-PIN commercial transaction
4. known TIF parcel
5. known non-TIF parcel
6. known Enterprise Zone parcel
7. parcel with an Assessor appeal
8. parcel with a Board of Review appeal
9. parcel with permit history
10. parcel with a known Assessor/GIS class mismatch, if available

The tool is designed so any adapter that fails or times out is marked in **Source Status** and **Issues**, while the rest of the research continues.

## Technical QA completed here

- VBA structure statically checked
- `Option Explicit` present
- 35 Subs matched by 35 `End Sub` statements
- 28 Functions matched by 28 `End Function` statements
- VBA module is ASCII-only for safer VBE import
- 30-second WinHTTP timeout present
- ArcGIS error-response guard present
- old website/API dependency removed
- old historical GIS layer removed as the primary GIS source
- current Property Tax Portal results route used
- current CookViewer direct parcel feature layer used
- current TIF boundary used as primary
- combined PDF includes the Sources appendix
- workbook formula-error scan returned no `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` cells before packaging

## Production gate

The only material step that cannot be completed in this environment is:

**Windows Excel → import module → Debug > Compile VBAProject → execute end-to-end against live PINs.**

If that compile/test surfaces any issue, preserve the v1.0 package and patch the failure rather than changing the overall adapter architecture.
