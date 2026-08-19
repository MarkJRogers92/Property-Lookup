# Cook County Property Due Diligence — Review Response

This responds to `CLAUDE_REVIEW_HANDOFF_Cook_Property_Due_Diligence.md`. It covers what was
checked, what changed in `CookPropertyDueDiligence_Standalone.bas`, and — importantly — what
could **not** be verified live and why, so the next person picking this up knows exactly where
it stands.

## How this review was done

This session's network egress is restricted to a small allowlist (GitHub only, in practice —
`datacatalog.cookcountyil.gov`, `gis.cookcountyil.gov`, `www.cookcountypropertyinfo.com`,
`arcgis.com`, and even `dev.socrata.com`/`en.wikipedia.org` all returned `EGRESS_BLOCKED`). So
this was **not** a live test against Cook County's systems. Instead, field-name correctness was
verified against the Cook County Assessor's own public source code — the `ccao-data/data-architecture`
GitHub repository, specifically `dbt/models/open_data/`, which contains the exact SQL views and
`exposures.yml` that CCAO uses to publish these Socrata datasets. That's about as authoritative a
source as exists short of querying the live API, since it's the code that *generates* the data
you're pulling — but it is still not a substitute for a live smoke test in Windows Excel, which
this environment cannot run. **Before production use, run `TestAdapters` and a full
`GenerateCookPropertyReport` against real PINs on a Windows machine with internet access.**

## Corrections made

### Priority 2 — Socrata field names (the main open question in the handoff)

Confirmed via `ccao-data/data-architecture/dbt/models/open_data/open_data.vw_parcel_universe_historical.sql`
(which `pabr-t5kh` / "Parcel Universe (Current Year)" is a filtered copy of):

- **Municipality - Tax Record** → field is `tax_municipality_name` (assigned via the parcel's tax
  code, `tax.agency_info` where `minor_type = 'MUNI'`).
- **Municipality - Spatial** → field is `cook_municipality_name` (point-in-polygon spatial join).

The macro previously guessed at several plausible aliases (`municipality_name`,
`municipality_name_spatial`, etc.) — none of which are the real field names. That guessing has
been replaced with the confirmed fields directly. This was a real bug: the old code would have
silently written blank values to both Property-sheet municipality rows on every run.

All other field names already in the module were cross-checked the same way and are correct as
written — no changes needed:
- Parcel Addresses (`3723-97qp`) ↔ `open_data.vw_parcel_address.sql`
- Assessed Values (`uzyt-m557`) ↔ `open_data.vw_assessed_value.sql`
- Parcel Sales (`wvhk-k5uv`) ↔ `open_data.vw_parcel_sale.sql`
- Assessor Appeals (`y282-6ig3`) ↔ `open_data.vw_appeal.sql`
- Permits (`6yjf-dfxs`) ↔ `open_data.vw_permit.sql`

Also confirmed: all seven Socrata dataset IDs used in the module (including `pabr-t5kh`) exactly
match `exposures.yml` in that repo — the dataset IDs are not stale.

The Board of Review dataset (`7pny-nedm`) is published by a separate agency (the Board of Review,
not the Assessor) and isn't in CCAO's own repo, so its field names (`tax_year`, `appealtrk`,
`appealseq`, `assessor_totalvalue`, `bor_totalvalue`, etc.) could only be checked indirectly via
search-indexed display names, which strongly suggest the existing lowercase/underscore field
names are already correct (Socrata API field names are always lower-snake-case regardless of
display name). Left as-is, but this is the one dataset where I'd prioritize a live smoke test.

### Priority 6 — Enterprise Zone (better outcome than originally scoped)

The handoff asked for the DCEO ArcGIS web app to be reverse-engineered to its authoritative
FeatureServer. I did not do that (still can't verify a guessed layer without live testing, and
the review guidance is explicit not to substitute one). Instead, I found something better: the
same Parcel Universe row already being fetched for every run carries `econ_enterprise_zone_num`
and `econ_enterprise_zone_data_year` — the Assessor's own spatial join against an Enterprise Zone
boundary layer, confirmed present via the same dbt source. `FetchEnterpriseZone` now reports this
automatically (zone number + data year) at **zero extra network cost**, clearly labeled as a
county-maintained signal rather than the live state boundary, with the official DCEO map link
still preserved and an Issues-sheet reminder to confirm before transaction-sensitive reliance.
Falls back to the old "Manual verification required" wording if Parcel Universe didn't return
data this run.

**Status: upgraded from "manual" to "partial/automatic-with-caveat."** Full "verified" status
still requires resolving DCEO's live layer, which needs a live test session.

### Priority 5 — TIF dynamic layer discovery

Implemented: `ResolveTifLayerUrl()` queries the TIF MapServer root (`.../tifSrvc/MapServer?f=json`)
at runtime, finds the layer whose name matches "Tax Increment"/"TIF" with the highest year, and
uses that. If discovery fails for any reason (network, parsing, unexpected format), it falls back
to the static `Config!TIF Layer URL` exactly as before — so this is additive, not a
behavior-risking change. Config sheet note updated to reflect the fallback role.

Also added: a cross-check against the Parcel Universe's `tax_tif_district_name` (same "tax-code
vs. spatial" duality already flagged for municipality). If the spatial query and the tax-code
field disagree, a HIGH issue is raised — this mirrors the existing municipality conflict-flagging
pattern and cost nothing extra to add since the data was already being fetched.

### Priority 1 — Correctness bugs found and fixed

1. **Socrata query strings could contain a literal space** (`$order=year DESC`). A raw space in
   the middle of a URL is not valid inside an HTTP request line; `SocrataCsv` now encodes it after
   assembly. This affected every adapter that passes `$order=...DESC` (Address, Assessment, Sales,
   Appeals, BOR, Permits).
2. **ArcGIS empty-result detection was fragile and had a worse silent-failure mode.** The code
   checked for the literal substring `"features":[]`, but Esri's `f=pjson` pretty-printer spaces
   it as `"features" : []`, so the check would never match — and, worse, if the ArcGIS service
   returned an *error* payload (no `features` key at all), the code fell through to the "found a
   result" branch and would have reported a false positive (e.g. "TIF intersection found" with a
   blank name) instead of failing loudly. Replaced with `ArcGisFeatureCount()`, which counts
   `"attributes":` occurrences (robust to formatting) and explicitly raises on an `"error"`
   payload so a real service failure surfaces as a failed source, not a wrong answer. Used in both
   `FetchGISParcel` and `FetchTIF`. Also switched both to `f=json` (compact) since pretty-printing
   was never needed for machine consumption.
3. **`ConfigurePrintAreas` could raise an uncaught runtime error** if a run sheet ended up
   completely empty (`.Cells.Find` returns `Nothing` on a blank sheet, and `.Row` on `Nothing` is
   error 91). Unlikely given the static headers in every sheet, but added a guard so a degraded
   run (e.g. every adapter failing) still produces a PDF instead of a fatal error.
4. **Header/footer `&` escaping.** Excel print header/footer text uses `&` as a formatting escape
   character (`&B` = bold). A property address containing a literal `&` (e.g. a corner lot, "MAIN
   ST & 5TH AVE") would have been mangled in the PDF header. Added `EscapeHeaderFooterAmp`.
5. Minor: reordered the BOR appeal-type fallback to try the confirmed real field (`appealtype`)
   first, keeping the alternate as a defensive fallback only.

I also did a full structural pass (balanced `Sub`/`Function`/`End` pairs — 28/28 and 25/25 — and
verified every `SetProperty`/`Range` reference in the module against the actual template cells,
described next).

### Workbook

**One template cell changed** (`Config!D12`, a note, updated to describe the new fallback
behavior — see above). Everything else required **no workbook changes**: I dumped every sheet's
cell contents with openpyxl and cross-checked it against every `Range(...)`/`Cells(...)` reference
and every `SetProperty`/field-name string in the macro (Start rows 12–22 source names, Property
rows 5–23 field labels, all data-sheet header rows, all Report-sheet cell references, all Config
setting names). Everything lines up exactly — this template and module were clearly built
together carefully. That's a meaningful finding on its own: Priority 1's "does everything the
code references actually exist in the workbook" check came back clean.

## Adapter status

| Source | Dataset/URL | Status | Notes |
|---|---|---|---|
| Assessor Addresses | `3723-97qp` | **Verified** (field names, via CCAO source) | Owner/mailing data caveat already flagged |
| Parcel Universe | `pabr-t5kh` | **Verified** (field names, via CCAO source) | Municipality field bug fixed |
| Assessed Values | `uzyt-m557` | **Verified** (field names) | |
| Parcel Sales | `wvhk-k5uv` | **Verified** (field names) | |
| Assessor Appeals | `y282-6ig3` | **Verified** (field names) | |
| Board of Review | `7pny-nedm` | **Partial** | Field names plausible but not confirmed against agency source code; recommend a live smoke test |
| Permits | `6yjf-dfxs` | **Verified** (field names) | |
| Cook County GIS (parcel layer) | `parcelHistorical/MapServer/<year>` | **Partial** | Fields (`Name`, `AssessorBLDGclass`, `AssessorNBHD`, `MUNICIPALITY`, `TAXCODE`) corroborated via search-indexed metadata, not a live query; empty-result bug fixed regardless |
| Property Tax Portal | `cookcountypropertyinfo.com` | **Partial/manual** | HTML parser untouched — could not fetch live markup to validate or improve it in this session; failure is non-fatal by design |
| TIF GIS | `tifSrvc/MapServer` | **Partial, improved** | Dynamic layer discovery added; spatial result cross-checked against tax-code TIF field; underlying layer fields (`TIF_NAME`, `AGENCY_DES`) not live-verified |
| Enterprise Zone | Parcel Universe + DCEO map | **Partial, improved** | Automatic county-signal added (see Priority 6 above); live DCEO boundary still not resolved |

"Verified" here means: field names confirmed against CCAO's own source code that generates the
dataset. It does **not** mean this session executed a live HTTP request — none were possible.

## Regression testing — could not be executed here

I want to be direct about this rather than paper over it: **no live PINs were tested against any
of these endpoints in this session**, because the sandbox has no route to
`datacatalog.cookcountyil.gov`, `gis.cookcountyil.gov`, or `cookcountypropertyinfo.com`. Any list
of "PINs that were tested and what they proved" would be fabricated if I produced one now — so
I'm not going to invent one. What I can offer is the test **plan**, matching what the original
handoff asked for, for someone (or a future session) with real access to execute:

1. An ordinary residential PIN (single-family, class 2xx) — exercises the baseline path.
2. A commercial/industrial PIN (class 5xx/6xx) — exercises the class-mismatch check if GIS class
   differs.
3. A multi-PIN commercial property (`num_parcels_sale` > 1 in a sale record).
4. A parcel inside a currently active TIF district — look one up via the Cook County TIF GIS
   layer or the Clerk's TIF report, then find its PIN via the Assessor's address search.
5. A parcel clearly outside any TIF district (most residential PINs qualify).
6. A parcel with a non-blank `econ_enterprise_zone_num` in Parcel Universe — cross-reference
   against DCEO's published Enterprise Zone list to also validate the new automatic signal.
7. A PIN with a recent Assessor appeal (check the Assessor's own appeal-status lookup for a
   PIN with a filed complaint in the current reassessment cycle).
8. A PIN with a Board of Review decision.
9. A PIN with permit history (search the Assessor's permits dataset directly, unfiltered by PIN,
   sorted by date, and pick any recent `pin`).
10. A PIN where Assessor class and GIS building class differ, if one can be found — likely rare;
    may require sampling several commercial parcels.

Recommended first step in a live session: run `TestAdapters` (prints the constructed URLs to the
Immediate window without touching the workbook) against a known PIN, then paste one URL into a
browser to eyeball the raw CSV/JSON before running the full macro.

## Sources that remain manual / cannot be fully automated

- **Illinois DCEO Enterprise Zone (live boundary).** Still not resolved to a FeatureServer.
  Reason: doing so requires live access to `idor.maps.arcgis.com` to trace the web app → webmap →
  operational layer chain and then test point-in-polygon against known zones — none of which was
  possible here. The Parcel Universe proxy value narrows this from "fully manual" to "manual
  confirmation of an automatic county-maintained hint," which is a real improvement but not the
  same as a live state check.
- **Property Tax Portal payment/exemption history.** The handoff asked whether payment status and
  exemption history could be captured; the `Tax History` sheet already has a "Payment / Status"
  column that the parser has never populated. I did not add scraping for this because I have no
  way to see the live markup and validate a regex against it in this session — guessing here risks
  adding a brittle, untested parser rather than a working one. Left as a known gap.

## Files in this response

- `CookPropertyDueDiligence_Standalone.bas` — corrected module (see diff/commit for the full change set)
- `Cook_Property_Due_Diligence_Standalone_Template.xlsx` — one cell changed (`Config!D12`)
- this document
