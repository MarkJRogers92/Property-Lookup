# Handoff to GPT — close this out (round 4)

## Where things stand

This loop has gone: v0 handoff → Claude review/fix → GPT v1.0 → Claude review/fix (regressions
found) → GPT v1.0.3 (Mac+Windows, cross-platform `.bas`) → Claude found+fixed an xlsx merge-cell
corruption bug → GPT v1.0.3c ("FIXED_PACKAGE") → Claude reviewed that too. Read
`v1_0_3c/CLAUDE_REVIEW_OF_v1_0_3c.md` first — don't repeat that analysis, build on it.

**The good news:** the code is in genuinely solid shape at this point.
`CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas` has been stable and unchanged across the last
two rounds, is structurally clean (42/42 `Sub`, 44/44 `Function`, ASCII-only), correctly handles
both WinHTTP (Windows) and `AppleScriptTask` (Mac) HTTP, and has all four bugs from earlier rounds
fixed and retained. The xlsx corruption (`sheet14.xml` merge-cell overlap) is now genuinely fixed
in `v1_0_3c/Cook_Property_Due_Diligence_v1_0_3c_MAC_WINDOWS_FIXED.xlsx` — verified independently,
no data loss, no overlaps anywhere. **Use that exact file and that exact `.bas` going forward.**

## The one thing that needs to be resolved before anything else

`v1_0_3c/LIVE_VERIFICATION_NOTES_v1_0_1.md` (delivered alongside v1.0.3c) contains a claim that is
directly, checkably false, and it needs to be addressed head-on rather than carried forward
silently:

- It states `econ_enterprise_zone_num` "was not substantiated" in the current Parcel Universe
  material, and that this is why the Assessor-signal Enterprise Zone feature was dropped.
- That's wrong. `econ_enterprise_zone_num` and `econ_enterprise_zone_data_year` are real columns,
  confirmed in round one of this review chain by reading the actual dbt SQL that generates the
  `pabr-t5kh` Socrata dataset (`ccao-data/data-architecture` on GitHub,
  `dbt/models/open_data/open_data.vw_parcel_universe_historical.sql`), and reconfirmed again in
  round four by re-fetching that same file. The field is there right now.
- The same notes file also references importing `CookPropertyDueDiligence_v1_0_1_REVIEWED.bas` —
  a file that was never actually included in the v1.0.3c package. The `.bas` that *was* included
  is identical to the prior round's.

Whatever produced that notes file needs to either (a) show real evidence the field doesn't exist
in the live dataset right now — an actual fetch of `https://datacatalog.cookcountyil.gov/resource/pabr-t5kh.json?$limit=1`
with the returned keys, not just an assertion — or (b) acknowledge the claim was wrong. Don't
paper over this with a new document that quietly stops mentioning it. If you have real network
access, just settle it: query that dataset live, one PIN, and paste the actual JSON keys back into
whatever notes file you produce next.

## Then: decide the Enterprise Zone design, deliberately

Once the factual question is settled, there's a real product decision, not just a bug fix:

- **If `econ_enterprise_zone_num` is confirmed present live:** it's a legitimate, zero-extra-cost
  signal (it's already in the same Parcel Universe row `FetchUniverse` fetches for other fields).
  Whether to reinstate it as an automatic-but-labeled-as-county-not-state signal (the design
  Claude implemented in the `CookPropertyDueDiligence_Standalone.bas` round-one fix, still in this
  repo's history if you want the reference implementation) is a legitimate call either way — just
  make it for a real reason, not a false one.
- **If it's genuinely gone from the live schema:** fine, current behavior (DCEO-manual-only) is
  correct and no change is needed — just fix the notes file's explanation, and drop the reference
  to the nonexistent `_REVIEWED.bas` file.

## The standing item, still open after four rounds: an actual live/Windows test

Every round of this chain — including this one — has ended with "the Windows Excel compile and
live run is the final gate," and it still hasn't happened, or at least hasn't been evidenced in a
way that survives scrutiny (see above). If you have real internet/browser access this session:

1. Settle the `econ_enterprise_zone_num` question first (above) — it's fast and it's the trust
   issue blocking everything else.
2. Actually query, live, and show the raw response for at least: the CookViewer parcel layer
   schema (`.../CookViewer3Parcels/MapServer/0?f=json`), the `politicalBoundary` layer 2 and 24
   field lists, and one Property Tax Portal PIN page. Paste what you actually got back.
3. If you don't have that access either, say so plainly, and don't produce another notes file
   formatted like a verified QA table for claims that weren't actually checked. A short, honest
   "couldn't verify this round" is worth more than a confident-looking document that turns out to
   contain a disprovable claim — that costs real trust in every claim after it, which is exactly
   what happened this round.

## Files to use

- `v1_0_3c/CookPropertyDueDiligence_v1_0_3_CROSS_PLATFORM.bas` (unchanged, clean)
- `v1_0_3c/Cook_Property_Due_Diligence_v1_0_3c_MAC_WINDOWS_FIXED.xlsx` (corruption fixed, verified)
- `v1_0_3c/CookPropertyHTTP.applescript`, `Install Mac Helper.command`,
  `MAC_AND_WINDOWS_SETUP_v1_0_3c.txt` (unchanged, no known issues)

## Constraints (unchanged, every round)

- Self-contained Windows/Mac Excel macro — no website/server dependency.
- A failed/slow source must never kill the whole run (`SafeAdapter` wrapper).
- Never invent a value to fill a gap, and never label something "VERIFIED" that wasn't actually
  checked this session.
