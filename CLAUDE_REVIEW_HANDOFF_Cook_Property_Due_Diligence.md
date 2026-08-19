# Cook County Property Due Diligence — Standalone Excel / VBA Review Handoff

## Goal

Convert the property due-diligence concept into a **self-contained Windows Excel tool**, modeled after the existing Cook/Lake appeal macros.

Workflow:

1. User enters a Cook County 14-digit PIN.
2. Excel queries each official source directly.
3. Each source adapter has its own timeout/error handling.
4. Results populate the workbook.
5. Conflicts and missing data are listed in **Issues / Items to Verify**.
6. Excel exports a single attorney-facing PDF.
7. A broken source must **not** kill the whole run.

There is no website/API dependency in this version.

## Files in this handoff

- `Cook_Property_Due_Diligence_Standalone_Template.xlsx`
- `CookPropertyDueDiligence_Standalone.bas`
- this handoff

The workbook is `.xlsx` because the build environment cannot embed VBA. For Windows production, save/import into an `.xlsm` copy and import the `.bas` module.

## Source architecture already selected

### 1. Cook County Assessor — Parcel Addresses
Dataset ID: `3723-97qp`

Use for:
- situs address
- owner name
- taxpayer/mailing name
- mailing address

Important: Cook County states owner/mailing data is only intermittently updated. Do not treat it as dispositive title evidence.

### 2. Cook County Assessor — Parcel Universe (Current Year Only)
Dataset ID: `pabr-t5kh`

Use for:
- current Assessor class
- township
- neighborhood
- tax/spatial geography
- parcel centroid (`lon`, `lat`, `x_3435`, `y_3435`)
- tax code only as a secondary value because the dataset warns that its tax district code may not be current

Important: the Assessor explicitly says there are two municipality concepts (tax-record and spatial) and they may disagree. The final tool should preserve both and flag conflicts.

### 3. Cook County Assessor — Assessed Values
Dataset ID: `uzyt-m557`

Use for historical:
- mailed land/building/total
- certified land/building/total
- Board of Review certified land/building/total

### 4. Cook County Assessor — Parcel Sales
Dataset ID: `wvhk-k5uv`

Use for:
- recorded sale date
- price
- Clerk document number
- deed type / MyDec deed type
- buyer/seller
- multi-PIN sale indicator

Important: sales can lag the actual recording date.

### 5. Cook County Assessor — Appeals
Dataset ID: `y282-6ig3`

Use for Assessor appeal history.

Important: the dataset has warned of historical sparsity/gaps. Preserve that limitation in the report.

### 6. Cook County Board of Review — Appeal Decision History
Dataset ID: `7pny-nedm`

Use for BOR appeal decisions and certified values, generally 2010-present.

### 7. Cook County Assessor — Permits
Dataset ID: `6yjf-dfxs`

Use for permit number/date/status/amount/applicant/work description.

Important: this is municipality-submitted data known to the Assessor, **not a complete universe of all municipal permits**.

### 8. Cook County GIS — Parcel layer
Current discovered layer:
`https://gis.cookcountyil.gov/traditional/rest/services/parcelHistorical/MapServer/2025`

Useful fields include:
- `Name` / PIN14
- `TAXCODE`
- `Latitude`
- `Longitude`
- `AssessorBLDGclass`
- `AssessorNBHD`
- `MUNICIPALITY`

Important: current Assessor open-data tables are updated in 2026 while this GIS historical parcel service currently exposes a 2025 parcel layer. Never compare them without showing source year/layer.

### 9. Cook County Property Tax Portal
`https://www.cookcountypropertyinfo.com/`

Use for:
- current composite tax rate
- tax code
- approximately five years of billed-tax history
- exemptions/payment context where practical
- property class description as a useful cross-check

This is an HTML adapter, so isolate the parser. If page markup changes, the adapter should fail cleanly and the rest of the workbook should continue.

### 10. Cook County TIF GIS
Current discovered TIF feature layer:
`https://gis.cookcountyil.gov/traditional/rest/services/tifSrvc/MapServer/3`

The layer is labeled **Tax Increment Finance Dist. (2024)**.

Current MVP approach:
- take parcel centroid `x_3435`, `y_3435`
- query the polygon layer with point geometry / `inSR=3435`
- return `TIF_NAME` when intersecting
- label output with the data year/layer
- never call "no intersection in 2024 layer" a permanent legal conclusion without verification

Recommended improvement:
- inspect the map service root dynamically and choose the latest TIF layer by layer name/year instead of hard-coding ID 3.

### 11. Illinois DCEO Enterprise Zone — unresolved automatic adapter
Official DCEO page points to this interactive map:
`https://idor.maps.arcgis.com/apps/webappviewer/index.html?id=f82fc6b62fde435abb41f5f72db2db48`

**Do not substitute a random ArcGIS Enterprise Zone layer from search results.**

Review task:
1. Resolve the ArcGIS application item to its webmap item.
2. Inspect the webmap operational layers.
3. Identify the authoritative current statewide Enterprise Zone polygon FeatureServer actually used by the DCEO/IDOR map.
4. Confirm its coordinate system and identifying zone-name field.
5. Test point-in-polygon against multiple known Cook County zones.
6. Only then enable automatic Yes/No + zone name.

Until that is done, the workbook should say `Manual verification required` and preserve the official map link.

## Workbook output

The current template includes:

- Start
- Property
- Assessment
- Tax History
- Sales-Deeds
- Appeals
- Incentives
- Permits
- Issues
- Sources
- Report
- Config

The intended final PDF is a multi-sheet PDF containing:
- Report
- Property
- Assessment
- Tax History
- Sales-Deeds
- Appeals
- Incentives
- Permits
- Issues

## Review priorities for Claude

Please **review and correct the supplied VBA instead of rewriting the whole concept from scratch**.

### Priority 1 — compile/runtime correctness
Check:
- all VBA compiles under 64-bit Windows Excel
- late-bound WinHTTP usage
- `Scripting.Dictionary` usage
- `VBScript.RegExp`
- collection/array handling in the CSV parser
- selected-sheet PDF export
- URLs and encoding
- error handling leaves Excel application state restored

### Priority 2 — Socrata fields
The adapters use official API field names where verified. For Parcel Universe municipality fields, the module intentionally tries several aliases because the exact municipality field names need to be confirmed against the current 124-column schema.

Please inspect the current dataset schema and replace alias guessing with the actual:
- tax-record municipality field
- spatial municipality field

Also confirm all fields used in:
- assessed values
- sales
- appeals
- BOR
- permits

### Priority 3 — Property Tax Portal parser
The first parser is deliberately lightweight.

Please test against:
- residential PIN
- commercial/industrial PIN
- PIN with exemptions
- PIN with unpaid/current balance
- PIN with unusual/missing characteristics

Improve parsing so the workbook reliably captures:
- current composite tax rate
- tax-rate history
- current tax code
- billed tax history
- payment/status if exposed
- exemption history if exposed

Do not make portal parser failure fatal.

### Priority 4 — GIS PIN query
The GIS layer uses `Name` as PIN14. Confirm whether the layer stores:
- 14 digits only
- formatted/hyphenated PIN
- both

The MVP queries both forms.

Confirm the parcel class field. The goal is specifically the GIS/Clerk/Assessor building class value the user currently checks manually in Cook County GIS.

### Priority 5 — TIF
Replace hard-coded annual layer when practical with dynamic latest-layer discovery.

### Priority 6 — Enterprise Zone
This is the main unresolved source. Resolve the official ArcGIS app to the real authoritative statewide polygon layer. Do not turn this on until tested.

### Priority 7 — PDF presentation
Keep the report legal/internal-professional:
- date only, no time
- compact tables
- PIN/address in header/footer if practical
- page numbers
- no AI/GPT references
- clear source-year labels
- `Could not verify automatically` rather than invented values
- public-record verification disclaimer

## Failure behavior

Every adapter must follow this contract:

- 30-second hard timeout by default
- mark source as failed
- add an Issues row
- continue to the next source
- do not leave Excel calculation/events/screen updating in a broken state
- never replace missing facts with assumptions

## Regression test set to build

Please create a small regression list with at least:
1. ordinary residential PIN
2. commercial/industrial PIN
3. multi-PIN commercial property
4. known TIF parcel
5. known non-TIF parcel
6. known Enterprise Zone parcel once the EZ layer is resolved
7. parcel with Assessor appeal
8. parcel with BOR appeal
9. parcel with permit history
10. parcel where Assessor/GIS class differs, if one can be found

The current smoke-test PIN in the module is `16-30-204-020-0000`; it was selected only as a public example and is not enough for production regression testing.

## Desired return handoff

Return:
- corrected `.bas` module(s)
- any workbook changes required
- a short changelog
- a list of source adapters marked `verified`, `partial`, or `manual`
- exact PINs used for regression testing and what each test proves
- any source that cannot be automated reliably, with the reason

Do not deploy a website. The production target is a self-contained Windows Excel macro workbook.
