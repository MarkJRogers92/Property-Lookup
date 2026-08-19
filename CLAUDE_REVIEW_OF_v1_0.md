# Review of v1.0 (GPT's pass) — what changed, what I fixed, what I can't confirm

This reviews the `v1.0` package (`CookPropertyDueDiligence_v1_0.bas`,
`Cook_Property_Due_Diligence_v1_0_Template.xlsx`, `CHANGELOG_v1_0.md`,
`LIVE_VERIFICATION_NOTES_v1_0.md`, `README_INSTALL_AND_USE.md`) against the prior round's fixes
(`CLAUDE_REVIEW_RESPONSE_Cook_Property_Due_Diligence.md`) and against the `.bas`/`.xlsx` directly.
I applied a handful of concrete fixes to `CookPropertyDueDiligence_v1_0.bas` in place (diff-sized,
listed below) and left everything else as delivered. **I have one important open question for you
at the bottom — please read it before treating v1.0 as production-ready.**

## The good news first

Whoever/whatever produced this did solid, careful work on the parts I can fully verify without
live network access:

- **Template/code correspondence is excellent.** I dumped every sheet in the `.xlsx` with
  openpyxl and cross-checked it against every `SetProperty`/`Range`/`Cells` reference in the
  `.bas` — all 33 Property-sheet field rows, the reorganized Report-sheet layout (B11/B12 now
  Building/Lot size, D8 now Tax Code, E9 now Latest Billed Tax, etc.), the new Documents sheet,
  the expanded Tax History columns (Payment/Status, Exemptions, Tax Sale/Delinquency), and every
  Config setting the code reads by name. Everything lines up. That's real, checkable engineering
  quality, independent of live access.
- **Structural claims in `LIVE_VERIFICATION_NOTES_v1_0.md` check out.** I independently counted:
  35 `Sub`/35 `End Sub`, 28 `Function`/28 `End Function` (now 30/30 after my two additions below),
  ASCII-only file, zero formulas and zero formula errors (trivially true — the workbook still has
  no formulas, only the new conditional formatting on `Start!B12:B22`, which is real and matches
  the changelog's claim).
- The overall shift in TIF/Enterprise Zone framing (current boundary as primary signal, an
  Assessor-derived signal for EZ with the DCEO map kept as the authoritative fallback, "PARTIAL"
  status surfaced explicitly in Source Status) is a reasonable design and consistent with what I'd
  recommended.

## Regressions I fixed (bugs from the prior review that came back)

These are concrete, provable from the code alone — none of them depend on live access to confirm:

1. **`SocrataCsv` no longer encodes the space in SoQL query strings** (`"$order=year DESC"`). A
   raw space in the middle of a URL is invalid inside an HTTP request line. This affected every
   dataset call that sorts results (Address, Assessment, Sales, Appeals, BOR, Permits). Fixed by
   encoding the assembled query string, same as the prior round.
2. **Print header/footer no longer escapes literal `&`.** Excel header/footer text uses `&` as a
   formatting escape (`&B` = bold); an address containing a literal `&` would be mangled in the
   PDF header. Restored `EscapeHeaderFooterAmp`.
3. **`ConfigurePrintAreas` dropped the guard against `.Cells.Find` returning `Nothing`** on a
   completely blank sheet (would raise error 91 and abort PDF export instead of degrading
   gracefully). Restored the guard.
4. **The empty-ArcGIS-result check is fragile and, on close inspection, arguably backwards.** All
   four spatial queries (CookViewer parcel, spatial municipality, TIF current boundary, TIF
   detail) decide "no results" by looking for the literal substring `"features":[]` with **no
   whitespace tolerance**. `EnsureArcGisResponse` (a new v1.0 addition) correctly catches a genuine
   ArcGIS *error* response first, which is good — but if Esri's server pretty-prints `f=pjson` with
   any spacing other than exactly `"features":[]` (I found real examples both with and without
   spaced colons while trying to confirm this, so I can't say for certain which this specific
   Cook County service uses), the literal check **never matches**, and depending on how each
   comparison is written, that silently reads as either "always found" or "always not found" —
   the wrong failure direction for a tool whose entire job is to distinguish "in a TIF/zone" from
   "not." I replaced all four call sites with a whitespace-tolerant `ArcGisFeatureCount()` (counts
   `"attributes":` occurrences via regex) that gives the right answer regardless of exact server
   formatting. This is a hardening fix, not something I can prove was actively wrong in production
   without a live test — but it removes a real, avoidable fragility.

All four fixes are minimal and isolated; I verified `Sub`/`Function`/`End` balance and line
continuations after editing (35/35 Subs, 30/30 Functions — two new small helper Functions added).

## What I could not verify (same sandbox limitation as last time)

This environment still has no route to `datacatalog.cookcountyil.gov`, `gis.cookcountyil.gov`,
`cookcountypropertyinfo.com`, or `arcgis.com`. So I cannot confirm or deny most of what
`LIVE_VERIFICATION_NOTES_v1_0.md` marks **LIVE SCHEMA VERIFIED** / **LIVE ENDPOINT VERIFIED**,
including:

- The `CookViewer3Parcels/MapServer/0` layer and its ~19-field `outFields` list (`PIN14`,
  `PIN14_dash`, `TAXYR`, `TAXDIST`, `XCOORD`, `YCOORD`, `LANDSF`, `CURRENTVALUE_TOTAL/LAND/BLDG`,
  `BLDGSQFT`, `bldg_const_desc`, `BCLASS`, `major_class_description`, `class_description`, `NBHD`,
  `BLDGAGE`, `tax_municipality_name`) — this replaces the old `parcelHistorical` layer entirely as
  the primary GIS source.
- `politicalBoundary/MapServer/2` (spatial municipality) and `/24` (current TIF boundary) as
  distinct layers under a `politicalBoundary` service.
- The Property Tax Portal's actual route (`PINResults.aspx` vs. the old `Pages/Pin-Results.aspx`)
  and page-section markers (`TAX BILLED AMOUNTS & TAX HISTORY`, `EXEMPTIONS`, `REFUNDS AVAILABLE`,
  `TAX SALE (DELINQUENCIES)`, `DOCUMENTS, DEEDS & LIENS`) that the new regexes depend on.

None of this is implausible — it's exactly the kind of detail a real live session would turn up —
but I want to be as direct with you as I was last round: **I did not verify any of it myself**,
and I'd be doing you a disservice if I let "LIVE VERIFIED" pass without comment just because it's
asserted confidently and formatted like a QA table.

**Direct question for you: did the session that produced v1.0 actually have live internet access
to these Cook County systems, or is `LIVE_VERIFICATION_NOTES_v1_0.md` describing what a session
*would* check if it had access (i.e., a verification plan written in the past tense)?** This
matters a lot for how much weight the workbook's own README should put on "Windows compile is the
only remaining gate" — if the endpoint/field details weren't actually observed live, that's a
second gate, not just the compile step.

One more thing worth a direct look before you rely on this: the regression PINs in
`LIVE_VERIFICATION_NOTES_v1_0.md` include `16-30-204-020-0000` — the same PIN the original
codebase's own `TestAdapters` comment called "a safe public smoke-test example... not enough for
production regression testing." It's now cited with quite specific claims (five-year billed
history, first-installment handling, exemption parsing, recorded documents/deeds/liens). That's
not necessarily wrong — it may just be a real PIN with genuinely rich history — but given the
above, I'd treat every claim tied to it as unconfirmed until you've actually pulled it up on the
Portal yourself.

## Recommendation

Don't skip the Windows live test because this package *looks* thoroughly QA'd — the parts that
are checkable from source code alone hold up well, but the parts that matter most for reliance
(does `CookViewer3Parcels` actually expose those exact fields, does the Portal route/markup match)
are exactly the parts no static review — mine or anyone's — can confirm. Same recommendation as
`GPT_HANDOFF_Cook_Property_Due_Diligence.md`: run `TestAdapters`, eyeball the raw responses, then
run `GenerateCookPropertyReport` end-to-end before this goes anywhere near client work.
