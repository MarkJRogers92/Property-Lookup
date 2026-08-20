# Review of v1.0.3c "FIXED_PACKAGE" — does it change anything?

Short answer: **the xlsx corruption fix is real and good — you can use this workbook instead of my
patched one. But the new `LIVE_VERIFICATION_NOTES_v1_0_1.md` in the same package contains a claim
I can directly disprove, and misdescribes its own package. Treat its other "verified" claims with
the same skepticism as the last two rounds, not more.**

## What's genuinely fixed: the xlsx

`Cook_Property_Due_Diligence_v1_0_3c_MAC_WINDOWS_FIXED.xlsx` takes a different, more conservative
approach than my patch: instead of removing just the 3 conflicting merge tags from the Instructions
sheet, it strips **all** merges from that sheet entirely (0 merge records, down from 10). I verified:

- No overlaps and no hidden values anywhere across all 14 sheets (I re-ran the same programmatic
  scan I used last time).
- No data loss — every cell's text content is intact (checked all 29 rows of Instructions).
- Opens cleanly in openpyxl.

Trade-off: the Instructions sheet's section headers and title banner lose their merged-cell
styling (no more spanning a colored band across columns A–F) since nothing is merged now — long
text will just overflow visually into the empty cells to its right instead, which looks close but
not identical. Functionally this is a complete, safe fix. **Use this file, not the one I patched
last round** — no reason to prefer mine now that this exists.

## What's unchanged: the .bas

`CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas` in this package is **byte-for-byte identical**
to the one I already reviewed and confirmed clean (42/42 Sub, 44/44 Function, all four previously
flagged bugs fixed, AppleScriptTask bridging intact). Nothing to re-review here.

## The problem: `LIVE_VERIFICATION_NOTES_v1_0_1.md` doesn't match what's actually in this package

This is new in this delivery, and it's worth being direct about two things I can prove from the
files alone, without needing live network access:

1. **It describes code changes that don't exist in this package.** Section 9 says "The v1.0 file
   available in this session did not actually contain the four fixes... They have now been
   reapplied" and separately claims v1.0.1 "removes the unverified Enterprise Zone Assessor
   pseudo-signal," "added Property Tax Portal fallback route," and "surfaced BOR `result`." The
   installation steps at the bottom even reference importing `CookPropertyDueDiligence_v1_0_1_REVIEWED.bas`.
   **That file isn't in this zip.** The `.bas` that is here is the same v1.0.3 file, unchanged, and
   — I checked — it already had the Portal fallback route, the BOR `result` field, and no
   Enterprise Zone signal *before this round*, going back to the v1.0.3 package I reviewed last
   time (I just hadn't specifically checked that section then). So this isn't a lie about *what
   the code currently does* — it's wrong about *when and why* those things happened, and it
   references a delivered artifact that was never delivered.

2. **Its central factual claim about Enterprise Zone is wrong, and I can show it.** The notes say:
   "The earlier v1.0 code attempted to use `econ_enterprise_zone_num` from Parcel Universe as an
   Assessor signal. This review did **not** substantiate that field in the current Parcel Universe
   material reviewed." I went back to the same authoritative source I used in my first review round
   — `ccao-data/data-architecture` on GitHub, the actual dbt SQL that generates the `pabr-t5kh`
   Socrata dataset — and re-fetched it just now. **`econ_enterprise_zone_num` and
   `econ_enterprise_zone_data_year` are still in the SELECT list, present tense, right now.** That's
   not a stale finding from an earlier round; I re-checked it specifically to respond to this claim.
   The field is real. Whatever process produced this notes file either looked in the wrong place or
   didn't actually check.

I'm not saying the workbook's current Enterprise-Zone-is-DCEO-manual-only design is wrong — that's
a legitimate, conservative product choice, and it's already what both this `.bas` and this `.xlsx`
consistently implement (Property row 33 is "Enterprise Zone (DCEO)", no Assessor-signal row; the
code has no `econ_enterprise_zone_num` reference). I'm saying the **justification** given for that
design in the new notes file is factually wrong, stated with the same confident "VERIFIED" label as
everything else in the document. That should lower your confidence in the rest of that document's
"LIVE ENDPOINT + SCHEMA VERIFIED" claims (CookViewer field list, the two `politicalBoundary` layer
IDs, the two regression PINs' Portal content, the BOR schema) — none of which I can independently
confirm from this sandbox either way, same as the last two rounds. The document's format looks more
rigorous than before (it includes raw-looking snippets instead of just a status table), but format
isn't evidence, and I just caught one of its specific claims being checkable and wrong.

## Bottom line / what I'd do next

- **Use the v1.0.3c `.xlsx`** — the corruption fix is real and verified.
- **Don't update the `.bas`** — nothing changed, nothing to do.
- **Don't treat `LIVE_VERIFICATION_NOTES_v1_0_1.md` as closing the live-verification gate.** If
  anything, ask whoever/whatever produced it to explain the `econ_enterprise_zone_num` discrepancy
  and the missing `v1_0_1_REVIEWED.bas` file before trusting the rest of it. The actual Windows
  Excel live test — the same one asked for in every round so far — is still the only thing that
  will really close this out.
