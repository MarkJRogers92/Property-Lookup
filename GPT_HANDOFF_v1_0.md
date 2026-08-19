# Handoff to GPT — v1.0 review round 2 (the live-verification gate is still open)

## Where things stand

You (or a prior session under this name) produced `v1.0`: a substantial expansion of the tool —
a new primary GIS source (`CookViewer3Parcels`), a spatial municipality-boundary layer, a
current-boundary-first TIF design with the old 2024 layer kept as a secondary cross-check, a much
richer Property Tax Portal parser (payment/status, exemptions, tax-sale/delinquency, recorded
documents/deeds/liens), and a new Documents sheet. `CHANGELOG_v1_0.md` and
`LIVE_VERIFICATION_NOTES_v1_0.md` describe it as schema/endpoint-verified against live Cook County
sources.

Claude reviewed it (`CLAUDE_REVIEW_OF_v1_0.md` — read that first, don't repeat the analysis) and:

- Confirmed the template/workbook wiring is fully correct — every `SetProperty`/`Range`/`Cells`
  reference in the `.bas` matches an actual cell in the `.xlsx`. No workbook changes needed.
- Confirmed the structural claims in the verification notes (Sub/Function/End balance, ASCII-only).
- Found and fixed **four regressions** — bugs that were already found and fixed in the prior
  round (v0) but came back in v1.0: a URL space-encoding bug in `SocrataCsv`, missing `&`
  escaping in print headers, a missing `Cells.Find`-returns-`Nothing` guard, and a fragile
  ArcGIS empty-result check (exact literal `"features":[]` match with no whitespace tolerance,
  used in all four spatial queries) — replaced with a regex-based `ArcGisFeatureCount()` helper.
- **Could not corroborate any of the "LIVE VERIFIED" claims**, because that review session had no
  network route to any Cook County or ArcGIS host either.
- Asked the user directly whether the session that produced v1.0 actually had live access. **The
  user doesn't know.** So right now, nobody in this chain has confirmed the new endpoints/fields
  against a live response. That's the situation you're picking up.

**Use `CookPropertyDueDiligence_v1_0.bas` as your base going forward — it already has Claude's
four fixes applied.** Don't reintroduce the patterns described above.

## What actually needs to happen now

This is the second time around this loop, and the live-verification gate still hasn't closed. If
you have real internet access in this session, use it — and **leave evidence**, not just a
verified/unverified label, so the next review round (whoever runs it) doesn't have to take it on
faith the way this round had to. Concretely: when you check an endpoint, paste the actual raw
JSON/HTML snippet you got back (or at least the field names/section headers you observed) into
`LIVE_VERIFICATION_NOTES_v1_0.md` next to the claim. "LIVE SCHEMA VERIFIED" with no supporting
snippet is exactly what got questioned this round.

Specifically, confirm or correct:

1. **`https://gis.cookcountyil.gov/traditional/rest/services/CookViewer3Parcels/MapServer/0`** —
   does this service exist, and does it expose all of: `PIN14`, `PIN14_dash`, `TAXYR`, `TAXDIST`,
   `XCOORD`, `YCOORD`, `street_address`, `township_name`, `latitude`, `longitude`, `LANDSF`,
   `CURRENTVALUE_TOTAL`, `CURRENTVALUE_LAND`, `CURRENTVALUE_BLDG`, `BLDGSQFT`, `bldg_const_desc`,
   `BCLASS`, `major_class_description`, `class_description`, `NBHD`, `BLDGAGE`,
   `tax_municipality_name`? Query it live: `.../MapServer/0?f=json` for the field list, then a
   real PIN query for actual values.
2. **`https://gis.cookcountyil.gov/traditional/rest/services/politicalBoundary/MapServer/2`**
   (spatial municipality) and **`/24`** (current TIF boundary) — do these layer IDs under
   `politicalBoundary` actually correspond to what the Config sheet's labels claim? Confirm the
   output fields (`MUNICIPALITY`, `AGENCY_DESC` for municipality; `AGENCY_DESCRIPTION`/
   `AGENCY_DESC`/`AGENCY_DES`/`TIF_NAME`, `AGENCY`/`AGENCYNUM` for TIF).
3. **The Property Tax Portal route** — `https://www.cookcountypropertyinfo.com/PINResults.aspx?pin=<PIN>`.
   Does this route exist and return PIN-specific content (not a generic search page)? Confirm the
   section markers the parser depends on: `PROPERTY CHARACTERISTICS`, `Property Class Description`,
   `Tax Rate History`, `TAX BILLED AMOUNTS & TAX HISTORY` (and the two fallback headings the code
   also tries), `EXEMPTIONS`, `REFUNDS AVAILABLE` / `No Refund Available`, `TAX SALE
   (DELINQUENCIES)`, `DOCUMENTS, DEEDS & LIENS`. Pull at least the two regression PINs already in
   `LIVE_VERIFICATION_NOTES_v1_0.md` (`16-30-204-020-0000`, `14-33-323-023-0000`) and paste what
   you actually see for the sections that PIN's page has.
4. Also worth a second look: **the ArcGIS `f=pjson` whitespace question itself.** Claude's fix
   assumes it doesn't matter (whitespace-tolerant regex either way), but if you can see a real raw
   response, note whether Esri is spacing colons as `"key" : value` or `"key":value` on this
   specific service — useful context even though the code no longer depends on getting it right.
5. **Board of Review dataset (`7pny-nedm`)** — still never confirmed against source code the way
   the CCAO-published datasets were (see the round-1 response doc). If you have live access, hit
   `https://datacatalog.cookcountyil.gov/resource/7pny-nedm.json?$limit=1` and diff the keys
   against what `FetchBORAppeals` expects (`tax_year`, `appealtrk`, `appealseq`, `appealtype`,
   `appealtypedescription`, `assessor_totalvalue`, `bor_totalvalue`).

## If you don't have live access either

Say so plainly in your response and in the notes file — don't reformat "unverified" as "verified"
under a different heading. It's fine for this loop to bottom out at "needs a human with a Windows
machine and a browser" — that's honest progress, not a failure. If that's the situation, the most
useful thing you can do instead is a pure code-quality pass on `CookPropertyDueDiligence_v1_0.bas`
(it's ~1340 lines now) — logic bugs, edge cases in the new HTML-parsing regexes, anything that
would misbehave even if every endpoint/field name above turns out to be correct.

## Constraints (unchanged)

- Self-contained Windows Excel macro — no website/server dependency.
- A failed/slow source must never kill the whole run (`SafeAdapter` wrapper).
- Never invent a value to fill a gap. "Could not verify automatically" + an Issues-sheet entry
  beats a guess, every time.
- Don't claim "LIVE VERIFIED" for anything you didn't actually observe live this session.
