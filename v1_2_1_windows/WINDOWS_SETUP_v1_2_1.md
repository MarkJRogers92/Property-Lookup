# Getting `new_property_search.xlsm` working on your Windows work computer

## What was wrong

This workbook's buttons ("RUN PROPERTY RESEARCH" and "RE-RUN TREASURER LOOKUP") are wired to the
code in **Module3** (internally called v1.2.1) — I confirmed this directly from the saved button
shapes in the file, not just from version text in a report, since those can get out of sync.

Everything in this workbook is already cross-platform **except one feature**: the Cook County
Treasurer integration (current rate, 20-year tax bill history, installments, exemption grid — the
"COOK COUNTY TAX INTELLIGENCE" section of the PDF). That was built using Safari browser automation
via AppleScript, which only exists on a Mac — there's no Safari and no `AppleScriptTask` on
Windows. On Windows as delivered, that one feature would have simply produced nothing (the rest of
the report — Assessor data, GIS, TIF, Property Tax Portal, Sales, Appeals, Permits, PDF export —
was already fully cross-platform and unaffected).

I checked every other Mac-specific code path in the workbook (there are several, including a newer
PTAB/Illinois Property Tax Appeal Board check) and confirmed they already degrade gracefully on
Windows with a clear status message — Treasurer was the only one with no Windows path at all.

## What I built

A genuine Windows equivalent of the Treasurer browser session, using **Internet Explorer COM
automation** (`CreateObject("InternetExplorer.Application")`) — the standard, long-established way
VBA drives a real browser session on Windows. This ships with Windows itself; nothing extra to
install. It works the same way the Mac path does: load the Treasurer's Overview page and the
20-Year History page in the same browser session, read the rendered page text, and feed it through
the exact same parsing/validation logic already used and already proven against real data (I didn't
touch that part — only how the page text gets fetched).

I also carried over a design decision already made on the Mac side: real-world testing found that a
plain HTTP request to the Treasurer site essentially never works on its own (the site needs an
interactive session) — so, like the Mac path, Windows now goes straight to the browser-based lookup
by default instead of wasting time on a probe that's unlikely to succeed.

**Important, please read before tomorrow:** Internet Explorer automation is a real risk area on
modern Windows. Microsoft has been retiring the standalone IE application for years, and some
locked-down or fully updated corporate machines may not have the COM automation components
available or enabled anymore. I built this to fail safely — if IE automation isn't available or
errors out, the Treasurer section is marked PARTIAL with a clear message, and **the rest of the
report generates completely normally**. Worst case tomorrow: everything works except the Treasurer
section, and you fall back to the "Open official Treasurer page" links already in the workbook for
that one property. I can't test this against your actual work machine from here, so please budget a
few minutes to try it before you need it live.

## Setup steps

1. Open `new_property_search.xlsm` in Windows Excel.
2. Press **Alt+F11** to open the VBA editor.
3. In the Project pane, right-click **Module3** → **Remove Module3...** → when prompted to export
   first, choose **No** (you don't need to keep the old one).
4. **File > Import File...** and select `Module3_v1_2_1_WINDOWS_TREASURER_FIX.bas` from this
   delivery. It will re-import as `Module3`, so the existing buttons (which point to
   `Module3`-defined macro names) keep working without any other changes.
5. **Debug > Compile VBAProject**. Fix nothing manually — if this shows an error, stop and send me
   the exact message and highlighted line before proceeding.
6. Save the workbook (keep it as `.xlsm`).
7. Close and reopen with macros enabled, then run a test PIN via the **RUN PROPERTY RESEARCH**
   button.

**You do not need to touch `Module1` or `Module2`.** They're older, unused duplicates left over
from earlier iterations — harmless, not wired to anything, safe to ignore for now. (Worth cleaning
out eventually, just not blocking you today.)

**No workbook/Config-sheet changes are required.** The new code reads two new optional settings —
`Windows Treasurer Strategy` and `Show Treasurer Browser Window` — but since those rows don't exist
in your Config sheet, it silently uses sensible defaults (go straight to the browser-based lookup;
run it hidden). If you want to see the IE window while it works (useful for a first test, or if
hidden mode has trouble), add a row to Config: column A = `Show Treasurer Browser Window`, column
B = `YES`.

## If the Treasurer section doesn't populate on your machine

Check the status message the macro writes into **Tax Detail!H74** — it will say exactly what
happened (IE unavailable, page didn't validate, timed out, etc.). That's your fastest diagnostic.
Send me that message and I can adjust from there. Everything else in the report is unaffected
either way.
