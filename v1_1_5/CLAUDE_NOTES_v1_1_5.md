# Quick notes on v1.1.5 (this checked in as working, real progress)

Not a full review pass — the user confirmed this is working well and is now having GPT focus on
visual/formatting polish in a separate branch. Recording what a second real live run confirms and
two small things worth a look, so the trail stays consistent for whoever picks this up next.

## Confirmed fixed (visible directly in the new PDF)

- `GIS Building Class` now shows `202` (matches Assessor Class), not `BCLASS`.
- `PIN` displays as `16-30-204-021-0000`, not scientific notation.
- `Municipality - Tax Record` / `Municipality - Spatial` both show real values (`CITY OF BERWYN` /
  `Berwyn`) and — notably — **no longer trigger the HIGH "municipality mismatch" issue** that
  showed up in the previous PIN's run. That confirms what was suspected last round: that mismatch
  was an artifact of the `JsonScalar` bug, not a real data conflict.
- Property Tax Portal now returns real 5-year history, exemptions, and tax-sale/delinquency
  status — this was failing live in the previous round's PDF. Whatever changed to fix it, it's
  working now.

## Real new capability since the last round

PTAB (Illinois Property Tax Appeal Board) status check, and a substantial new Cook County
Treasurer integration (current rate, 20-year tax bill history, current/recent installments,
exemption-by-year table) — the Source Notes describe this as using "the same Safari property
session," i.e. driving an actual browser via AppleScript rather than a raw HTTP GET, which would
explain why the old direct-HTTP Portal fetch was unreliable if the real page needs session state
or JS to render.

## Two things worth a look (not yet fixed, flagging for whenever this gets picked up)

1. **`Percent Change` and `20-Year Change` show a bare `+` with no number** (Treasurer section,
   page 11 of this PDF). Mechanically diagnosable from the code: `ValueAfterLabel(historyBlock,
   "Percent Change", "+-0123456789.%")` stops capturing at the first character not in
   `allowedChars`. Getting only `"+"` back means whatever immediately follows the sign in the real
   page text isn't a digit — most likely a space, newline, or icon/markup artifact separating a
   colored sign indicator from the number in the source page. Can't pin the exact fix without
   seeing the real page text around "Percent Change" / "20-Year Change," but the mechanism is
   clear enough to know it's a real gap, not a display/formatting issue — the actual number is
   simply not being captured.
2. **`PIN` row's "As Of / Layer" column shows `7/18/05`** instead of `2026` (page 2), while every
   other field sourced from "Assessor Addresses" in the same run correctly shows `2026`. All of
   these fields pull `Nz(DictGet(d, "year"))` from the *same* dictionary in `FetchAddress`, so they
   should be identical — worth checking directly in the workbook rather than trusting this PDF
   text extraction, since `7/18/05` matches the M/D/YY date style used elsewhere in this same PDF
   (Sales/Deeds dates), which raises the possibility this is a text-layout artifact from reading
   the PDF rather than a real bug in the workbook. Flagging rather than asserting.

## Housekeeping, not urgent

The `.xlsm` now bundles **three full generations** of the whole application as separate modules
(`Module1` = v1.1.5, `Module2` = v1.2.0, `Module3` = v1.2.1), each with its own button-install
routine. The PDF's own embedded text confirms `Module1` (v1.1.5) is what actually ran — so v1.2.0
and v1.2.1 are present but dormant, same pattern as the orphaned `v104c` module a few rounds back.
Worth consolidating down to one active module before this grows further, so "what actually ran"
stays unambiguous — same reasoning as last time, just now three generations deep instead of two.
