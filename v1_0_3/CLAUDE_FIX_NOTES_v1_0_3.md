# v1.0.3 — xlsx repair fix

## What was broken

Opening `Cook_Property_Due_Diligence_v1_0_3_MAC_WINDOWS.xlsx` triggered Excel's repair dialog:

```
Repair Result to Cook_Property_Due_Diligence_v1_0_3_MAC_WINDOWS1.xml
Removed Records: Merge cells from /xl/worksheets/sheet14.xml part
```

`sheet14.xml` is the **Instructions** sheet (confirmed via `xl/_rels/workbook.xml.rels`).

## Root cause

Two pairs of merged-cell ranges on the Instructions sheet overlapped the same cells — invalid
under the OOXML spec, so Excel silently strips the merge records on open (that's what "repair"
means here; it's not random corruption, it's Excel refusing to render an inconsistent file):

- `A3:F3` and `A3:B3` both claimed cells A3/B3 (row 3 = "MAC — ONE-TIME SETUP")
- `A24:F24` and `A24:B24` both claimed cells A24/B24 (row 24 = "MAC HELPER / FIRST TEST")

Both rows look like they got merged twice by two different code paths — once as a two-column
section-header (matching the working pattern on rows 11 and 18, which only merge `A:B`), and once
as a full-width spacer band (matching rows 10 and 17, which correctly merge `A:F` as blank
dividers). Rows 3 and 24 are section headers, not spacer rows, so they should only have followed
the first pattern.

I checked all 14 sheets programmatically for this exact defect class (overlapping merge ranges) —
**only the Instructions sheet was affected.** Nothing else in the workbook has this problem.

## A second, related bug found in the same sweep

`A25:F25` (row 25: "Required folder" / `~/Library/Application Scripts/com.microsoft.Excel/`) was
also merged across all six columns. That merge itself isn't invalid (no overlap), so Excel
wouldn't have flagged it — but a merged range only displays its **top-left cell's** value. B25's
actual folder path would have been silently invisible even after Excel's automatic repair, which
only removes the genuinely broken merges and leaves this one in place. Every other row in that
section (26–29) correctly leaves label/value in separate, unmerged A/B cells.

## Fix applied

Removed exactly three `<mergeCell>` tags from `xl/worksheets/sheet14.xml` — nothing else in the
file was touched (verified byte-for-byte identical elsewhere by rewriting only that one zip
member):

- Removed `A3:F3` (kept `A3:B3`, matching the working section-header pattern)
- Removed `A24:F24` (kept `A24:B24`, same reasoning)
- Removed `A25:F25` entirely (so both "Required folder" and the actual path are visible again)

Verified: `openpyxl` loads the fixed file cleanly, and the same 14-sheet overlap/hidden-value
scan now reports clean across the whole workbook.

## The .bas module needed no changes

While in here, I checked `CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas` against the three
concrete bugs flagged in earlier review rounds (Socrata query-string space encoding, print
header `&` escaping, the `Cells.Find`-returns-`Nothing` guard, and the fragile ArcGIS
empty-result check). **All four are correctly fixed and present** — the ArcGIS check in
particular now strips whitespace from the response before comparing, which solves the same
problem a different way than my earlier regex-based fix, and works correctly. Structural balance
checked out too (42/42 `Sub`, 44/44 `Function`, ASCII-only). No `.bas` changes were made.

## Not re-verified here

Same standing limitation as the last two rounds: this review environment has no network route to
Cook County or ArcGIS hosts, so none of the live endpoint/field claims for v1.0.3 could be
confirmed or denied. This fix is scoped strictly to the reported symptom (the file-corruption
repair dialog), which is a static file-format defect fully diagnosable without live access.
