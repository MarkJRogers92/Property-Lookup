# Review response — v1.2.8b

Reviewed both files against the exact hashes cited in the handoff (both matched — this really is
the audited pair, not a stale copy): workbook `1715943e...ba39`, module `933f2532...2081`.

## Bottom line

**No BLOCKER or HIGH found. Safe to test in Windows Excel as-is — no patch required.**

The specific reported failure — `Resetting workbook — Error 1004: We can't do that to a merged
cell` — is fixed correctly. I didn't just check that `ClearContentsMergeSafe_v128b` exists; I
independently mapped every merged range in the actual `.xlsx` (Report has 7, Tax Detail has 7,
Executive Summary has 67, Appeal Report has 65, Tax Intelligence Report has 50 — dumped and
overlap-checked all of them programmatically, zero overlaps found, matching the audit) and traced
every single `.Clear` / `.ClearContents` / `.Delete` / `.Merge` / `.UnMerge` call in the module (49
occurrences) against that actual geometry to confirm none of them can hit the partial-merge case.

## 1. Merged-cell/runtime safety — checked in full

- `ClearContentsMergeSafe_v128b` (line 2422): logic is correct — loops `target.Areas`, then cells,
  clears `cell.MergeArea.ClearContents` once per unique merge (tracked via a Collection) or
  `cell.ClearContents` for a plain cell. This can never hit "partial merge" because it always
  either clears a whole `MergeArea` or a single unmerged cell.
- `ResetRunSheets` (2386): every content-clearing call now routes through the helper — verified
  every line, not just the ones the audit already listed. Confirmed fixed.
- `SimplifyTaxDetail_v128` (3834), `BuildTaxEstimatePresentation_v128` (3941),
  `UpdateExecutiveSummaryTaxEstimate_v128` (4068): these use a different but equally valid pattern
  — `UnMerge` the whole target block, `Clear` it, then rebuild merges from scratch. I checked that
  each `UnMerge` target range **fully contains** every merge that exists in that region in the
  actual template (not just assumed it) — e.g. `BuildTaxEstimatePresentation_v128`'s
  `A20:L35` unmerge covers all 27 of Tax Intelligence Report's merges in that band, and the
  rebuild recreates the identical 27-merge geometry, so it's idempotent across repeated runs, not
  just correct on a fresh template.
- `EnsureV122MockupWorkbookStructure` (188) does an unqualified `.Cells.Clear` on the whole Tax
  Detail sheet — that's always safe regardless of merges (a whole-sheet clear can never partially
  overlap anything on that same sheet) — but see the dead-code note below; it's not actually in
  the live run's danger path anyway.
- `InstallRunButton_v128Windows` (508) and `InstallTaxEstimateConfig_v128` (3772): both clear/unmerge
  small ranges (`E6:H7`, `A26:D31`) that don't intersect any existing merge on those sheets in the
  template. Safe.

## 2. Compile correctness

No blockers found. `HttpGet`'s Windows branch uses late-bound `CreateObject("WinHttp.WinHttpRequest.5.1")`
correctly; every Office constant used (`xlCalculationManual`, `msoShapeRoundedRectangle`,
`xlHAlignCenter`, etc.) is a standard built-in needing no extra reference; `Address(External:=True)`
is a real, valid named parameter. Independently re-verified the audit's Sub/Function balance
(73/73 each) and duplicate-name check (zero duplicates) rather than trusting it blind — both hold.

## 3. Reset → run sequence

Traced the full call chain in `GenerateCookPropertyReport_v128Windows`. One thing worth knowing,
not fixing: `SimplifyTaxDetail_v128` / `BuildTaxEstimatePresentation_v128` /
`UpdateExecutiveSummaryTaxEstimate_v128` run **twice** — once right after `ResetRunSheets` (before
any adapter has populated `Property`, so those formulas reference empty cells at that point), and
again via `RefreshTaxEstimateEnhancement_v128` after all adapters finish. This isn't a bug: the run
sets `Application.Calculation = xlCalculationManual` at the very start, so nothing actually
recalculates until the explicit `Application.Calculate` / `Application.CalculateFullRebuild` calls
— and those all happen after the *second* rebuild, so the values baked into the PDF are correct.
The first rebuild is simply redundant work (rebuilding a formula skeleton that gets rebuilt
identically moments later) — worth trimming someday, not a defect today.

## 4. PDF generation

- `BuildPdfSheetList` sheet names all cross-checked against the actual workbook's 18 sheet names —
  every one exists (Executive Summary, Property, Assessment, Tax History, Tax Intelligence Report,
  Sales-Deeds, Documents, Appeal Report, Incentives, Permits, Issues, Sources). No missing-sheet
  regression.
- `ExportDueDiligencePDF` calls `Application.CalculateFullRebuild` twice (once before copying the
  sheets out, once after, "for Mac compatibility" per its own comment) — that's a genuinely
  thorough guarantee that formula values are fresh before export.
- Downloads path: `Environ$("USERPROFILE") & "\Downloads"` — correct for Windows.
- PDF stays closed: `OpenGeneratedFile` (which would call `FollowHyperlink`) is defined but **never
  called** anywhere in the run — confirmed by grep, not assumption. PDF export ends at
  `SetRunStatus "Complete - PDF saved to " & fullPath` and returns.
- Print-area `Find`-returns-`Nothing` guard (from an earlier round) is still correctly in place at
  line 2755 — no regression there.

## 5. Source independence

`SafeAdapter` (704) is unchanged from the established pattern — every adapter call is wrapped, a
failure logs a HIGH issue and continues. Confirmed intact.

## 6. Tax Detail / tax estimate — verified formula-by-formula

- `B20` (Estimated Gross Taxes) = `B18*F19*D19/100` — Assessed Value × equalizer × rate ÷ 100,
  exactly the stated formula, and it references `B18` (AV) directly, **not** `H18` (market value) —
  so the Config assessment-level override (which only feeds `F18`/`H18`) genuinely cannot leak into
  the gross-tax number. This was explicitly asked about in the handoff; confirmed correct.
- `B19` (Estimate Basis Year) checks for 2025 first, falls back to 2024 exactly as specified, via
  `Config!B26`/`Config!B28`; `F19` (equalizer) maps 2025→`Config!B27` (3.03) and 2024→`Config!B29`
  (3.0355) — matches the stated factors exactly.
- Traced the data path one level further than the formulas alone: `UpsertTaxDetailRow` (3278)
  writes the Portal's year/rate rows into Tax Detail rows 9–15, which is what `B19`/`D19`'s
  `COUNTIF`/`MATCH` against `A9:A14`/`B9:B14` depend on. Minor note: the writer allows 7 rows
  (9–15) but the lookup formulas only cover 6 (9–14) — if a 7th distinct year ever landed in row
  15, it would be invisible to the estimate. Very low risk in practice (the Portal returns ~5–6
  years) and not related to the reported error — flagging as LOW, not fixing without being asked.

## 7. Stale/dead references (informational, not blocking)

None of these affect the current Windows run — they're unreachable — but you asked me to flag
anything with behavior/error-message risk:

- `EnsureV122MockupWorkbookStructure`, `RunTreasurerBrowserLookup_v122Mockup`, and
  `InstallTreasurerLookupButton_v122Mockup` are Mac-era Treasurer-browsing leftovers. The button
  they'd wire up is never installed in the Windows flow (`InstallRunButton_v128Windows` only
  installs the one blue "RUN PROPERTY RESEARCH" button), so they're dead code — but if an *older*
  workbook still has a leftover Treasurer button shape from a prior version whose `OnAction` still
  points at `RunTreasurerBrowserLookup_v122Mockup`, clicking it would run
  `EnsureV122MockupWorkbookStructure`'s full Tax Detail rebuild (harmless on its own, but pointless
  and confusing given the manual-only design). Not worth touching now; worth deleting eventually.
- `FetchTreasurerOverview` (1343) is fully wired into `SafeAdapter`'s dispatch table but never
  called from the main run — `ConfigureWindowsTreasurerManual_v128` is what actually runs for
  Treasurer. Consistent with "Treasurer is manual only," just more retained-but-unreachable code.
- `DefaultDocumentsFolder` (2364) has a typo — `Environ$("USERPROFILE") & "\\Documents"` (doubled
  backslash) — but the function is never called anywhere (`ExportDueDiligencePDF` uses
  `DefaultDownloadsFolder`, not this one), so it's cosmetic dead code, not a live bug.

None of section 7 requires action before testing. I did not patch or rewrite any of it, per the
stabilization-mode instruction — flagging only.

## Confirmation

Nothing here rises to BLOCKER or HIGH. **This package should be safe to test in Windows Excel as
delivered** — no patch needed from me before you run `Debug > Compile VBAProject`, the one live PIN
run, and the PDF export the handoff already planned as the final gate.
