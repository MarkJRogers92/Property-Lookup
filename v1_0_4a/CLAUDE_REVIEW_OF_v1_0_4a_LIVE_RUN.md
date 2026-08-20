# Review of the first real live run (v1.0.4a) — this is a big milestone, and it found a real bug

## What this actually is

This is the first artifact in this whole chain that came from an actual Windows/Mac Excel run
against the live Cook County systems — not a claim about verification, an actual PDF produced by
actually running the macro. That's worth pausing on: it proves the whole architecture works
end-to-end. WinHTTP/AppleScriptTask HTTP calls succeed, the Socrata adapters pull real data, the
PDF export works, and — importantly — when one adapter genuinely failed live (Property Tax
Portal), the rest of the run kept going and produced a complete report anyway. That's the
`SafeAdapter` design doing exactly what it was built to do.

It also surfaced a real, serious bug that four rounds of static review — mine included — never
caught, because it's the kind of bug that only shows up against real data.

## The bug: `JsonScalar` was returning field *names* instead of field *values*

Look at the PDF: `GIS Building Class` = `BCLASS`, `Tax Code` = `TAXDIST`, `Lot Size (SqFt)` =
`LANDSF`, `Municipality - Spatial` = `MUNICIPALITY`, the "As Of / Layer" column showing `TAXYR`
literally instead of a year — every single CookViewer/ArcGIS-sourced field is echoing back its own
field name instead of the actual value. Meanwhile every Socrata-sourced field (Assessor Class,
the whole Assessment History table, Sales/Deeds, Appeals, Permits) shows real, correct, plausible
data. That split is the diagnostic signal: it's not a network problem, it's specific to how
ArcGIS JSON responses get parsed.

**Root cause, confirmed by extracting the actual VBA from the `.xlsm`:** `JsonScalar(js, key)`
searches for the first occurrence of `"key"` *anywhere* in the raw response text, then takes
whatever follows the next colon. ArcGIS `f=pjson` responses include a `fields` array and a
`fieldAliases` map *before* the actual `features`/`attributes` data — and when a field has no
custom alias configured (the default), Esri sets the alias equal to the field name. So `"BCLASS"`
shows up in the document twice: once in metadata (as `"alias":"BCLASS"` or similar), and once for
real inside `attributes`. The unscoped scan finds the metadata occurrence first and returns that.

I ported the actual VBA logic to Python and tested it against a realistic mock ArcGIS response
(fields array + features array, matching the real shape) — reproduced the bug exactly, then
verified the fix eliminates it. Not a guess; tested.

**Fix:** added `ArcGisFirstAttributesJson()`, which locates the first `"attributes":{...}` block
(proper brace-depth and string-escape tracking, so it can't be fooled by braces inside a string
value) and clips the search to just that substring before `JsonScalar` does its key lookup. Falls
back to the whole document if no `attributes` key exists at all, so nothing else regresses.

This is a four-line change in effect (one new function, one line changed at the top of
`JsonScalar`), but it's the single most consequential fix in this whole review chain — it's the
difference between the GIS/municipality/TIF-name portions of the report being real data versus
being silently wrong in a way that looks plausible enough to miss at a glance.

**Retrospective note, for honesty's sake:** this exact bug has been present in every `JsonScalar`
implementation across every round I reviewed, including my very first fix. My round-1 version used
`VBScript.RegExp` with `re.Global = False` — same "find the first match anywhere" behavior, same
vulnerability. Four rounds of careful static review never caught it because it requires an actual
ArcGIS response shaped like a real one (with `fields`/`fieldAliases` preceding `features`) to
surface — which is exactly what static analysis and my sandbox's lack of network access could
never produce. This is the clearest evidence yet for why the live-verification gate every round
has pointed to actually matters, beyond just being cautious box-checking.

## A second, smaller bug in the same PDF: PIN displayed as `1.6302E+13`

Page 1 (Property Summary) shows `PIN: 1.6302E+13` instead of `16-30-204-020-0000`. Cause:
`SetProperty "PIN", pin, ...` writes the raw 14-digit string; Excel's default cell format
auto-coerces an all-digit string into a Number, and a 14-digit number displays in scientific
notation under General format. (The Report sheet shows the PIN correctly because `BuildReport`
already wraps it in `FormatPin(...)` before writing it — the Property sheet just never got the
same treatment.) Fixed by storing `FormatPin(pin)` (the hyphenated form) instead of the raw digit
string — hyphenated text isn't auto-numeric, and it's a no-op-safe change since both existing
`FormatPin(GetPropertyValue("PIN"))` call sites already tolerate an already-formatted string.

## A real, confirmed-live failure: the Property Tax Portal adapter

The Issues sheet in this PDF shows a genuine live failure: both the primary route
(`.../PINResults.aspx`) and the fallback (`.../cookviewerpinresults.aspx`) failed
`PortalPageLooksValid`'s check (needs the formatted PIN, `"PROPERTY CHARACTERISTICS"`, and
`"TAX BILLED AMOUNTS"` all present in the page text). I can't diagnose the exact cause from here —
could be a wrong route, JS-rendered content a static GET can't see, a redirect to a generic search
page, or a User-Agent/anti-bot block — without seeing the actual returned HTML.

Worth connecting to the earlier review thread: round 3's `LIVE_VERIFICATION_NOTES_v1_0_1.md`
claimed the Portal route and content were "observed live," including specific page text for two
regression PINs. This live run just failed on exactly that adapter, for exactly the PIN those
notes claimed to have verified. I flagged that document's credibility as questionable at the time
over a different, independently-provable claim (`econ_enterprise_zone_num`); this is now a second,
independent data point pointing the same direction. I'd suggest not trusting anything that notes
file says about the Portal specifically.

The good news: this failure is *exactly* the scenario `SafeAdapter` was designed for. It logged a
HIGH issue and the rest of the report generated correctly. The architecture is doing its job here
even though this one adapter needs real debugging with actual page content in hand.

## Everything else in this PDF looks genuinely correct

- Assessment History (2019–2026): real, internally consistent mailed/certified/BOR values.
- Sales/Deeds: five real recorded sales with plausible dates, prices, doc numbers, and names.
- Appeals: one real 2026 Assessor appeal, correctly captured.
- Permits: two real permits with real work descriptions.
- Enterprise Zone: "No county signal returned" for `econ_enterprise_zone_num` — for an ordinary
  Berwyn residential parcel, that's very likely the *correct* answer (most parcels aren't in an
  EZ), and it further confirms that field genuinely exists and is being read successfully — one
  more independent point against round 3's claim that it "was not substantiated."
- The three flagged HIGH issues (Portal failure, Assessor/GIS class mismatch, tax-record vs.
  spatial municipality mismatch) are exactly the kind of thing this tool is supposed to catch —
  though the class/municipality mismatches here are themselves *downstream* of the `JsonScalar`
  bug (a CookViewer class of literally `"BCLASS"` will never equal the real Assessor class `"202"`,
  and `"MUNICIPALITY"` will never equal a real tax-record municipality name). Worth re-running
  against this same PIN after the fix to see whether those two issues are genuine data conflicts
  or artifacts of the bug — I'd guess artifacts, but that's exactly the kind of thing only another
  real run can confirm.

## What to do with this

`CookPropertyDueDiligence_v1_0_4a_JSON_PIN_FIX.bas` in this delivery is the corrected `Module1`
source, ready to re-import. In the VBA editor: right-click **Module1** in the existing `.xlsm` →
**Remove Module1** → **File > Import File** → select this file. Removing first (rather than
importing alongside) keeps the module name `Module1`, so the existing "RUN PROPERTY RESEARCH"
button's `OnAction` binding keeps working without any other changes. Then **Debug > Compile
VBAProject**, save, and re-run PIN `16-30-204-020-0000` — worth comparing this new PDF against the
one just delivered field-by-field to confirm the fix.

One housekeeping note, not urgent: this `.xlsm` also carries a second, dormant code module
(`CookPropertyDueDiligence_v104c`) with its own `GenerateCookPropertyReport_v104c` and a separate
"Setup" entry point that isn't wired to any button — it never ran and isn't what produced this PDF,
but it has the exact same two bugs (confirmed by inspection) and it's dead weight sitting in the
file. Worth cleaning out whenever convenient, not blocking anything now.

I also applied both fixes to `v1_0_3c/CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas` in this
repo, so the reference baseline going forward doesn't carry this bug into the next round.
