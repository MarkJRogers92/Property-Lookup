COOK COUNTY PROPERTY DUE DILIGENCE v1.2.8b — WINDOWS COMPLETE

FRESH PACKAGE
Do not layer this on any older workbook.

INSTALL
1. Open Cook_Property_Due_Diligence_Windows_v1_2_8b_COMPLETE_TEMPLATE.xlsx.
2. Save As an Excel Macro-Enabled Workbook (*.xlsm).
3. Press Alt+F11.
4. File > Import File.
5. Import CookPropertyDueDiligence_v1_2_8b_WINDOWS_COMPLETE_MERGE_SAFE.bas.
6. Debug > Compile VBAProject.
7. Close the VBA editor.
8. Press Alt+F8 and run SetupCookDueDiligence_v128Windows once.
9. Enter a Cook County PIN and click RUN PROPERTY RESEARCH.

WHAT v1.2.8b FIXES
The reset routine is now merge-safe. It no longer assumes that clearing the top-left
cell of a merged area is valid. Every reset content-clear goes through a helper that
clears the entire MergeArea when required.

This specifically fixes the current Report!A14:H21 merged issue-summary block and
also protects future reset ranges from the same class of error.

TREASURER
Manual cross-check only in the Windows build.

TAX ESTIMATE
Uses the most recent common year for Portal tax rate + final Cook County equalizer:
- 2025 equalizer 3.0300
- 2024 fallback 3.0355

Estimated Gross Taxes are before exemptions.

VALIDATED HERE
- No formula error matches in the fresh workbook scan.
- No overlapping merged-cell definitions in the XLSX.
- VBA procedure counts balanced.
- No duplicate procedure names.
- Reset targets route through merge-safe clearing.

FINAL REQUIRED TEST
Windows Excel must still perform:
- Debug > Compile VBAProject
- one live PIN run
- PDF export
