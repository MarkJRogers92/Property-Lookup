# Claude Review Handoff — Cook County Property Due Diligence v1.2.8b

## Review request
Please review this **Windows Excel/VBA production package** before further changes. The user has spent substantial time testing regressions, so prioritize **runtime correctness and minimal targeted fixes** over refactoring.

Please inspect both files together:

- `Cook_Property_Due_Diligence_Windows_v1_2_8b_COMPLETE_TEMPLATE.xlsx`
- `CookPropertyDueDiligence_v1_2_8b_WINDOWS_COMPLETE_MERGE_SAFE.bas`

Current hashes:

- Workbook SHA256: `1715943e93d3d8a51a407aeb8699ef58e01231c6080991299f46f818725ba39a`
- VBA SHA256: `933f25322e5050a77352f6ffa76ed901dfabe0ebdbaac9524920427b527f2081`

## Production target

- Windows desktop Excel at the user's work computer.
- Self-contained workbook after one `.bas` import.
- No website/server dependency.
- No Mac helper dependency.
- Cook County Treasurer is deliberately **manual only** because its session-sensitive workflow was unreliable on the work computer.
- Property Tax Portal remains the automatic tax-detail source.
- PDF should save to Downloads and remain closed.

## Normal install/run

1. Open the fresh `.xlsx`.
2. Save As `.xlsm`.
3. Import the one full `.bas` module.
4. Debug > Compile VBAProject.
5. Run `SetupCookDueDiligence_v128Windows` once.
6. Enter PIN on Start.
7. Click the blue button, which runs `GenerateCookPropertyReport_v128Windows`.

## Current source architecture

The full module includes the working research adapters for:

- Assessor Addresses — Cook County Socrata
- Parcel Universe — Cook County Socrata
- Assessed Values — Cook County Socrata
- Parcel Sales — Cook County Socrata
- Assessor Appeals — Cook County Socrata
- Board of Review history — official Cook County BOR dataset (no CAPTCHA website dependency)
- PTAB status check
- Permits
- Cook County GIS / current parcel and municipality
- Property Tax Portal / 5-year tax detail
- TIF GIS
- Enterprise Zone county signal + manual DCEO verification
- PDF packet generation

## Tax estimate enhancement

The executive report now shows:

- Current Assessed Value
- Estimated Market Value
- Latest available composite tax rate
- Estimated Gross Taxes

Formula:

`Estimated Gross Taxes = Current Assessed Value × Cook County equalizer × composite tax rate ÷ 100`

The model uses the most recent year for which both inputs exist:

- 2025 Cook County equalizer = 3.0300
- 2024 fallback equalizer = 3.0355

Gross estimate is before exemptions.

Estimated market value uses standard class assessment levels:

- Major Classes 1/2/3 = 10%
- Major Class 4 = 20%
- Major Class 5 = 25%
- Config includes an assessment-level override for incentive/special classes.

## Treasurer behavior

Automatic Treasurer browsing was removed from the Windows production run.

The workbook retains only a compact manual cross-check and official Treasurer search link.

Do not reintroduce an automatic Treasurer session/browser workflow unless specifically requested.

## Important failure history

### Earlier working baseline
A v1.2.6a Windows workbook had the core production research engine working.

### Fresh-template macro problem
A later `.xlsx` template was saved as `.xlsm` and only an add-on was imported. The resulting file lacked the production research macro. v1.2.8 solved this by consolidating the **entire production module** into one `.bas`.

### v1.2.8 reset failure
Windows Excel stopped at:

`Resetting workbook — Error 1004: We can't do that to a merged cell.`

The first identified cause was stale Tax Detail reset ranges after its layout was redesigned.

### v1.2.8a reset failure
The same error still occurred. The remaining cause was:

- `Report!A14:H21` is one merged issue-summary block.
- The reset routine attempted `Range("A14").ClearContents`.
- Clearing only the top-left cell of a merge is still treated by Excel as modifying part of a merged cell.

## v1.2.8b fix

v1.2.8b does **not** merely patch `A14`.

It introduces:

`ClearContentsMergeSafe_v128b(ByVal target As Range)`

Behavior:

- Iterates `target.Areas` explicitly so multi-area ranges such as `B4,D4,F4,H4` are handled.
- For each cell:
  - if merged, clears `cell.MergeArea.ClearContents`;
  - otherwise clears `cell.ClearContents`.
- Tracks already-cleared merge addresses so a merge area is cleared once.

`ResetRunSheets` now routes every content clear through this helper, including:

- raw result sheets
- Property
- Tax Detail raw Portal cells
- Report B/E values
- full `Report!A14:H21`
- Start progress/note cells

The run-button setup also uses the helper when clearing its merged target area.

## Static audit already completed

`AUDIT_v1_2_8b.txt` records the checks.

Current results:

- Workbook has **zero overlapping merge definitions**.
- VBA: 73 Subs / 73 End Subs.
- VBA: 73 Functions / 73 End Functions.
- No duplicate procedure names.
- `ResetRunSheets` uses merge-safe clearing throughout.
- Old partial `Range("A14").ClearContents` is absent.
- Full `Report!A14:H21` reset is present.
- Main run and setup entrypoints are present.

## Specific review tasks for Claude

Please review for the following, in this order:

### 1. BLOCKER — merged-cell/runtime safety
Search every VBA use of:

- `.Clear`
- `.ClearContents`
- `.Delete`
- `.Merge`
- `.UnMerge`
- range writes that may target only part of a merged range

Confirm every runtime path is safe in **Windows Excel**.

Pay special attention to:

- `ResetRunSheets`
- `EnsureV122MockupWorkbookStructure`
- `InstallRunButton_v128Windows`
- `SimplifyTaxDetail_v128`
- `BuildTaxEstimatePresentation_v128`
- `UpdateExecutiveSummaryTaxEstimate_v128`
- PDF preparation routines

### 2. BLOCKER — compile correctness
Check for anything that could fail `Debug > Compile VBAProject` in Windows Excel:

- declarations
- argument types
- conditional compilation
- Office constants / references
- `Shape` / `TextFrame` calls
- collection iteration
- named arguments such as `Address(External:=True)`

Prefer late-bound/reference-free forms where appropriate.

### 3. BLOCKER/HIGH — reset → run sequence
Trace the exact normal run sequence:

1. Read PIN
2. Reset workbook
3. Rebuild Tax Detail presentation
4. Rebuild Tax Intelligence presentation
5. Rebuild Executive Summary formulas
6. Run adapters
7. Build final report/issues
8. Calculate presentation sheets
9. Export PDF

Identify any location where the reset clears formulas or structural cells that are **not restored before they are needed**.

### 4. HIGH — PDF generation
Check:

- temporary workbook sheet list
- page order
- print areas
- no missing `Report`/mockup-sheet reference regression
- Downloads path on Windows
- leave PDF closed
- no Mac-only or unsupported workbook properties

### 5. HIGH — source independence
Confirm one failed optional adapter cannot terminate the entire report. Preserve current `SafeAdapter` behavior.

### 6. HIGH — Tax Detail / tax estimate
Confirm:

- Treasurer is manual only
- Portal rate history drives the tax-rate selection
- 2025 factor 3.0300 / 2024 fallback 3.0355 are wired correctly
- gross tax estimate uses AV directly
- market-value assessment-level override does not accidentally change the gross-tax formula
- blank/missing 2025 rate correctly falls back to 2024

### 7. MEDIUM — stale version references
There may still be internal procedure names/comments inherited from v1.2.2/v1.2.6. Do not rename them solely for cosmetics if renaming risks regression, but flag any stale reference that could change behavior or produce a misleading user error.

## Requested output from Claude

Please return:

1. **BLOCKER / HIGH / MEDIUM / LOW findings**, with procedure + approximate line/location.
2. A minimal patch/diff for any BLOCKER or HIGH issues.
3. Explicit confirmation whether the package should be safe to test in Windows Excel after the patch.
4. Do not redesign the workbook or rewrite working adapters unless a concrete defect requires it.
5. If no additional blocking defect is found, say so explicitly rather than inventing changes.

## Key principle

This project is now in stabilization mode. Prefer one boring, reliable Windows workbook over clever refactors.
