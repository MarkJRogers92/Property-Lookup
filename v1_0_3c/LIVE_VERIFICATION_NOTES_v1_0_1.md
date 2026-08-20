# Live Verification Notes — v1.0.1 Review

**Verification date:** August 19, 2026

This file supersedes the unsupported `LIVE VERIFIED` labels in the earlier v1.0 notes.  
The evidence below records what was actually observed from official Cook County / Illinois sources in this review round.

## Bottom line

- The **current CookViewer parcel layer exists and its schema supports every field used by the GIS adapter**.
- The **Municipality layer is ID 2** and the **current TIF boundary layer is ID 24** under Cook County's `politicalBoundary` service.
- The **2024 TIF detail layer is ID 3** under `tifSrvc` and exposes the TIF name / municipality / first-year fields used by the secondary cross-check.
- The **Property Tax Portal section markers and two regression PINs were observed live/indexed from the official portal**.
- The **Board of Review field names used by the macro are confirmed by the current official dataset**.
- The core Assessor Socrata schemas used for addresses, assessed values, sales, appeals, and permits were checked against the current official dataset pages.
- **Enterprise Zone remains PARTIAL/manual**. No unverified Assessor field or random ArcGIS Enterprise Zone layer is treated as authoritative.
- **Windows Excel compile/runtime remains the final production gate.**
- A limitation of this browser session: it could inspect the live ArcGIS layer/schema and confirm Query support, but it could not directly submit an arbitrary parameterized parcel query and show the returned feature row. Therefore the **GIS schema and query capability are verified; specific-PIN ArcGIS execution still needs the Windows regression run**.

---

## 1. CookViewer current parcel layer

Official endpoint:

`https://gis.cookcountyil.gov/traditional/rest/services/CookViewer3Parcels/MapServer/0`

Observed raw service evidence:

```text
"id": 0
"name": "Parcels"
"type": "Feature Layer"
"geometryType": "esriGeometryPolygon"
"latestWkid": 3435
"capabilities": "Map,Query,Data"
```

The raw `f=pjson` response is pretty-printed with whitespace, for example:

```text
"currentVersion": 11.5,
"id": 0,
"name": "Parcels",
```

That confirms why the old exact-literal `"features":[]` test was brittle. v1.0.1 retains Claude's whitespace-tolerant `ArcGisFeatureCount()` approach.

Fields observed in the live schema and used by `FetchGISParcel`:

```text
PIN14
PIN14_dash
TAXYR
TAXDIST
XCOORD
YCOORD
street_address
township_name
latitude
longitude
LANDSF
CURRENTVALUE_TOTAL
CURRENTVALUE_LAND
CURRENTVALUE_BLDG
BLDGSQFT
bldg_const_desc
BCLASS
major_class_description
class_description
NBHD
BLDGAGE
tax_municipality_name
```

The layer also has indexes on `PIN14` and `PIN14_dash` and advertises Query support.

**Status:** LIVE ENDPOINT + SCHEMA VERIFIED.  
**Still to test in Windows:** a real PIN query executed by WinHTTP from the macro.

---

## 2. Spatial municipality boundary

Official endpoint:

`https://gis.cookcountyil.gov/traditional/rest/services/politicalBoundary/MapServer/2`

Observed:

```text
Name: Municipality
ID: 2
Geometry Type: esriGeometryPolygon
Spatial Reference: latestWkid 3435
Fields:
AGENCY
AGENCY_DESC
MUNICIPALITY
```

**Status:** LIVE LAYER + FIELD SCHEMA VERIFIED.  
**Still to test in Windows:** parcel-centroid point-in-polygon execution.

---

## 3. Current TIF boundary

Official endpoint:

`https://gis.cookcountyil.gov/traditional/rest/services/politicalBoundary/MapServer/24`

Observed:

```text
Name: Tax Increment Financing (TIF) District
ID: 24
Geometry Type: esriGeometryPolygon
Spatial Reference: latestWkid 3435
Fields:
AGENCY
AGENCY_DESCRIPTION
AGENCYNUM
```

The parent `politicalBoundary` service lists Municipality as layer 2 and TIF District as layer 24.

**Status:** LIVE LAYER + FIELD SCHEMA VERIFIED.  
**Still to test in Windows:** parcel-centroid point-in-polygon execution.

---

## 4. 2024 TIF detail cross-check

Official endpoint:

`https://gis.cookcountyil.gov/traditional/rest/services/tifSrvc/MapServer/3`

Observed:

```text
Name: Tax Increment Finance Dist. (2024)
Display Field: TIF_NAME
Geometry Type: esriGeometryPolygon
Fields include:
AGENCY
AGENCY_DES
TIF_NAME
Municipality
First_Year
rev2024
```

This is correctly retained as a **secondary, explicitly 2024** data source rather than the primary current-boundary test.

**Status:** LIVE LAYER + FIELD SCHEMA VERIFIED.

---

## 5. Property Tax Portal

Primary route in Config:

`https://www.cookcountypropertyinfo.com/PINResults.aspx`

v1.0.1 also has a fallback to:

`https://www.cookcountypropertyinfo.com/cookviewerpinresults.aspx`

The Portal landing/results content confirms it provides:

```text
Billed Amounts & Tax History
Property Description
Tax Exemptions
Refund Search
Documents, Deeds & Liens
```

### Regression PIN A — 16-30-204-020-0000

Observed official Portal content includes:

```text
PROPERTY CHARACTERISTICS
2235 EAST AVE
BERWYN
Lot Size (SqFt): 3,660
Building (SqFt): 910
Property Class: 2-02
Tax Rate: 12.242
Tax Code: 11002

TAX BILLED AMOUNTS & TAX HISTORY
2025: $4,232.14*
2024: $7,694.86 — Paid in Full
*=(1st Install Only)

EXEMPTIONS
2024: 1 Exemptions Received — Homeowner Exemption
2020: 0 Exemptions Received, Certificate of Error Applied

REFUNDS AVAILABLE
No Refund Available

TAX SALE (DELINQUENCIES)
2025: Tax Sale Has Not Occureed
2024: Tax Sale Has Not Occurred
2022: No Tax Sale

DOCUMENTS, DEEDS & LIENS
2421420042 - MORTGAGE - 08/01/2024
2421420041 - WARRANTY DEED - 08/01/2024
```

The live/indexed page therefore confirms the parser markers and the first-installment asterisk path.

### Regression PIN B — 14-33-323-023-0000

Observed official Portal content includes:

```text
PROPERTY CHARACTERISTICS
418 W EUGENIE ST
CHICAGO
Lot Size (SqFt): 2,262
Building (SqFt): 2,289
Property Class: 2-06
Tax Rate: 6.618
Tax Code: 74026

TAX BILLED AMOUNTS & TAX HISTORY
2024: $28,328.04
Pay Online: $108.29
```

The Portal explains that this amount represents a **balance due**.

Observed recorded-document index:

```text
2334545070 - LIS PENDENS FORECLOSURE - 12/11/2023
```

That supports the v1.0.1 issue escalation for both current balance-due conditions and adverse recorded-document language.

Portal disclaimer observed on both property-result pages:

```text
The displayed information does not include municipal special assessments or omitted taxes.
```

**Status:** LIVE/INDEXED OFFICIAL CONTENT VERIFIED FOR BOTH REGRESSION PINS.  
**Windows test:** confirm WinHTTP receives the same page body and parser output.

---

## 6. Board of Review dataset

Official dataset:

`https://datacatalog.cookcountyil.gov/Property-Taxation/Board-of-Review-Appeal-Decision-History/7pny-nedm`

The current official schema confirms:

```text
tax_year
appealtrk
appealseq
appealtype
appealtypedescription
assessor_totalvalue
bor_totalvalue
```

The dataset is current through June 9, 2026. v1.0.1 also preserves the returned `result` value in the workbook's appeal-status column when present.

**Status:** LIVE OFFICIAL SCHEMA VERIFIED.

---

## 7. Other Cook County Assessor Socrata adapters

### Parcel Addresses — `3723-97qp`

Observed current official fields include:

```text
pin
year
prop_address_full
prop_address_city_name
prop_address_zipcode_1
mail_address_name
mail_address_full
mail_address_city_name
mail_address_state
mail_address_zipcode_1
owner_address_name
```

The Assessor warns owner/mailing data is only intermittently updated.

### Assessed Values — `uzyt-m557`

Observed fields used by the macro:

```text
year
class
mailed_land / mailed_bldg / mailed_tot
certified_land / certified_bldg / certified_tot
board_land / board_bldg / board_tot
```

### Parcel Sales — `wvhk-k5uv`

Observed fields used by the macro:

```text
sale_date
sale_price
doc_no
deed_type
mydec_deed_type
seller_name
buyer_name
is_multisale
num_parcels_sale
```

### Assessor Appeals — `y282-6ig3`

Observed fields used by the macro:

```text
year
case_no
hearing_type
appeal_type
status
mailed_tot
certified_tot
```

### Assessor Permits — `6yjf-dfxs`

Observed fields used by the macro include:

```text
date_issued
year
permit_number
local_permit_number
status
assessable
amount
municipality
applicant_name
work_description
```

Cook County expressly warns this permit dataset is not a complete universe of municipal permits.

**Status:** LIVE OFFICIAL SCHEMA VERIFIED FOR THE MACRO FIELD SETS ABOVE.

---

## 8. Enterprise Zone

Official DCEO page:

`https://dceo.illinois.gov/expandrelocate/incentives/taxassistance/enterprisezone.html`

The earlier v1.0 code attempted to use `econ_enterprise_zone_num` from Parcel Universe as an Assessor signal. This review did **not** substantiate that field in the current Parcel Universe material reviewed.

v1.0.1 therefore removes the pseudo-signal entirely.

Current behavior:

```text
Enterprise Zone (DCEO): Manual boundary verification required
Source status: PARTIAL
```

The macro preserves the official DCEO map link and does not return a fabricated Yes/No boundary result.

**Status:** MANUAL / PARTIAL BY DESIGN.

---

## 9. Claude regression fixes reconciled into v1.0.1

The v1.0 file available in this session did not actually contain the four fixes described in Claude's handoff. They have now been reapplied:

1. Socrata `extraQuery` spaces are encoded as `%20`.
2. Excel print-header values escape literal `&` as `&&`.
3. `ConfigurePrintAreas` guards against `Cells.Find` returning `Nothing`.
4. ArcGIS empty-result detection uses whitespace-tolerant `ArcGisFeatureCount()` rather than exact `"features":[]`.

Additional v1.0.1 changes:

- removed the unverified Enterprise Zone Assessor pseudo-signal;
- added Property Tax Portal fallback route;
- surfaced BOR `result` when returned;
- synchronized the workbook's Enterprise Zone field and source notes with the reviewed VBA.

---

## 10. Remaining production gate

This environment cannot run Windows Excel/VBA.

Before client reliance:

1. Open the reviewed template on the Windows work computer.
2. Save as `.xlsm`.
3. Import `CookPropertyDueDiligence_v1_0_1_REVIEWED.bas`.
4. Run **Debug > Compile VBAProject**.
5. Run PIN `16-30-204-020-0000`.
6. Compare Portal fields / PDF output against the live Portal.
7. Run PIN `14-33-323-023-0000`.
8. Confirm the macro flags the `$108.29` balance and the `LIS PENDENS FORECLOSURE` document.
9. Run at least one known commercial/industrial PIN, one known TIF parcel, and one known non-TIF parcel.
10. Confirm CookViewer, municipality, and TIF ArcGIS point-in-polygon calls succeed from WinHTTP on the work network.

If any adapter fails, preserve this v1.0.1 baseline and patch only that adapter.
