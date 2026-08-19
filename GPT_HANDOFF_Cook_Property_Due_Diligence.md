# Handoff to GPT — Cook County Property Due Diligence (finish the live-verification pass)

## Why this handoff exists

This project is a self-contained Windows Excel/VBA tool (no server, no website) that takes a
Cook County PIN and pulls property due-diligence facts from ~11 public sources into one workbook
and PDF. It went through one review/fix pass already (by Claude), documented in
`CLAUDE_REVIEW_RESPONSE_Cook_Property_Due_Diligence.md`. That pass fixed real bugs and confirmed
most Socrata field names — but it ran in a sandboxed environment with **no network access to any
of the actual data sources** (`datacatalog.cookcountyil.gov`, `gis.cookcountyil.gov`,
`cookcountypropertyinfo.com`, `arcgis.com` were all blocked). Everything in that pass was verified
indirectly, against Cook County Assessor's own public source code on GitHub, never by making a
live request. If you have live internet access, please pick up exactly where that gap is and
finish the verification this project actually needs before anyone relies on it.

**Read `CLAUDE_REVIEW_RESPONSE_Cook_Property_Due_Diligence.md` first** — it has the full account
of what was checked, what changed, and why. Don't repeat that analysis; build on it.

## Files

- `CookPropertyDueDiligence_Standalone.bas` — the corrected VBA module. Import into a `.xlsm` copy
  of the template to run it (the build environment that produced the `.xlsx` can't embed VBA).
- `Cook_Property_Due_Diligence_Standalone_Template.xlsx` — the workbook. Sheet layout and every
  cell the macro references have already been cross-checked against the code; you shouldn't need
  to change the template unless live testing surfaces a real mismatch.
- `CLAUDE_REVIEW_RESPONSE_Cook_Property_Due_Diligence.md` — prior review notes and rationale.
- `CLAUDE_REVIEW_HANDOFF_Cook_Property_Due_Diligence.md` — the original design handoff (source
  architecture, per-dataset intent). Still the best reference for *why* each adapter exists.

## What to do, in priority order

### 1. Live smoke test (do this first, before anything else)

In Windows Excel, import the `.bas` module into a `.xlsm`, then run `TestAdapters` — it prints
the constructed Socrata URLs to the Immediate window without touching the workbook. Paste each URL
into a browser and confirm it returns real CSV rows with the expected columns. Then run
`GenerateCookPropertyReport` end-to-end against a real PIN and check every sheet populates
sensibly. This alone will catch anything the source-code-only review missed.

### 2. Confirm the Board of Review dataset (`7pny-nedm`)

This is the one dataset that isn't published from CCAO's own open-source pipeline, so its field
names (`tax_year`, `appealtrk`, `appealseq`, `assessor_totalvalue`, `bor_totalvalue`, `appealtype`)
were inferred, not confirmed. Query `https://datacatalog.cookcountyil.gov/resource/7pny-nedm.json?$limit=1`
directly and diff the returned keys against what `FetchBORAppeals` in the `.bas` expects.

### 3. Validate the Property Tax Portal HTML parser

`FetchTaxPortal` scrapes `https://www.cookcountypropertyinfo.com/Pages/Pin-Results.aspx?pin=<PIN>`
with regex against collapsed page text. This was never tested against live markup. Pull the page
for a few real PINs (residential, commercial, one with exemptions, one with a current balance) and
verify:
- `Tax Rate` / `Tax Code` regexes still match
- the `Tax Rate History` and `TAX BILLED AMOUNTS` section boundaries (`TextBetween` calls) still
  find the right text
- whether payment/status and exemption history are exposed in a scrapeable form — the `Tax
  History` sheet has a "Payment / Status" column that nothing currently populates; only add
  scraping for it if you can see and test the real markup, don't guess at a regex

If the portal's markup has changed meaningfully, the parser needs rework — but keep the contract
that a portal failure must stay non-fatal (caught by `SafeAdapter`, logged to Issues, rest of the
run continues).

### 4. Confirm the GIS parcel layer fields and query approach

`FetchGISParcel` queries `https://gis.cookcountyil.gov/traditional/rest/services/parcelHistorical/MapServer/<year>/query`
filtering on `Name='<PIN>' OR Name='<formatted-PIN>'`, requesting `Name,PIN10,TAXCODE,Latitude,
Longitude,AssessorBLDGclass,AssessorNBHD,MUNICIPALITY`. Hit `.../MapServer/<year>?f=json` live and
check the actual field list — confirm `Name` (not `PIN14` or something else) is really the queryable
PIN field, and that the other five field names are exact. This was corroborated only through
search-engine-cached metadata, not a live schema fetch.

### 5. Resolve the DCEO Enterprise Zone layer (if you have time for it)

The macro now reports an automatic Enterprise Zone signal sourced from the Assessor's own
`econ_enterprise_zone_num` field in Parcel Universe (see the review response for why). That's a
real improvement but it's still not the live state boundary. If you can reach
`https://idor.maps.arcgis.com/apps/webappviewer/index.html?id=f82fc6b62fde435abb41f5f72db2db48`,
trace it to its webmap item, find the actual operational Enterprise Zone FeatureServer layer,
confirm its coordinate system and zone-name field, and test point-in-polygon against a few known
zones before wiring up a second, authoritative signal alongside the existing one. Don't replace
the Parcel Universe signal — keep both, and flag disagreement between them the same way TIF and
municipality conflicts are already flagged.

### 6. Run the regression test plan

`CLAUDE_REVIEW_RESPONSE_Cook_Property_Due_Diligence.md` has a 10-item test plan (residential,
commercial, multi-PIN, TIF, non-TIF, Enterprise Zone, Assessor appeal, BOR appeal, permit history,
class-mismatch) with guidance on how to find a qualifying PIN for each. Actually run it, record
the PINs used and what each run proved, and fold that list back into the project docs — that's
the one thing this review still owes and could never produce without live access.

## Constraints that still apply

- No website/API dependency — this stays a self-contained Windows Excel macro.
- A failed/slow source must never kill the whole run (the `SafeAdapter` wrapper pattern already
  enforces this — keep it that way for anything new).
- Never invent a value to fill a gap; "Could not verify automatically" plus an Issues-sheet entry
  is always correct where "verified" isn't.
- Don't guess at ArcGIS/Socrata field names or endpoints the way earlier drafts of this file did —
  confirm them, live, before wiring them in.
