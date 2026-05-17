# Content/Chronology Review: Current vs Outdated Artifacts

This review classifies files using **content differences** (line/byte/hash deltas) and **git chronology** (commit dates/messages), not just filenames.

## Method used

- Compared VBA module candidates by line count, byte size, and content hash.
- Compared pairwise diffs by insert/delete volume.
- Reviewed `git log --follow` timestamps and commit messages for each candidate.

## VBA module lineage (content + chronology)

### 1) `legacy/SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed.bas`
- Size: 1,618 lines / 61,223 bytes
- Hash prefix: `ffbf40e96ab3a8a980fd42ca`
- Earliest observed commit in history: `13511af` (2026-03-19, "uploaded newest")
- Status: **outdated/interim baseline**.

### 2) `legacy/SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed_2.bas`
- Size: 1,751 lines / 66,282 bytes
- Hash prefix: `83c11e118b16d9cbb43b6892`
- Observed chronology: appears after `_fixed.bas` (`5606f83`, 2026-03-21)
- Content delta vs `_fixed.bas`: ~136 insertions, 3 deletions.
- Status: **newer than `_fixed.bas`, but still interim/outdated**.

### 3) `SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed_2_documented.bas`
- Size: 2,804 lines / 104,867 bytes
- Hash prefix: `e5e3a62296b7b06a18784280`
- Observed chronology: appears in same lineage window as `_fixed_2` and represented as documented evolution.
- Content delta vs `_fixed_2.bas`: ~1,053 insertions (substantial growth/documentation/refinement).
- Status: **current recommended macro in active tree**.

### 4) `SCLX-specification-package/SCLX_Ledger_IO_v13_with_supplemental_dualrefs.bas`
- Size: 3,061 lines / 116,780 bytes
- Hash prefix: `d0818447a038ecfd4dc38869`
- Chronology includes later packaging commit (`0131eff`, 2026-03-31, "add all .bas and new json").
- Content delta vs documented module: ~317 insertions, 60 deletions.
- Interpretation: this is a **distinct branch/variant** (dualrefs + supplemental behavior), not just a strict superseding drop-in replacement.
- Status: **active specialized variant** (not legacy by age alone).

## Current vs outdated determination

### Keep as current (active tree)
- `SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed_2_documented.bas`
- `SCLX-specification-package/SCLX_Ledger_IO_v13_with_supplemental_dualrefs.bas` (specialized variant)
- SCLX 1.3 schemas/manuals/rules in `SCLX-specification-package/`

### Keep as legacy
- `legacy/SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed.bas`
- `legacy/SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed_2.bas`
- SCLX 1.2 schema/manual/rules set in `legacy/SCLX-specification-package/`
- `legacy/SCLX-specification-package/problems.txt` (captured run output)

## Important nuance

The `...dualrefs.bas` file is **not** necessarily "newer and therefore replacement" for the documented macro. Content and naming indicate it is a feature variant. Recommend documenting both explicitly as:
- default/general macro,
- dualrefs/supplemental macro for advanced compatibility paths.

## Suggested follow-up cleanup

1. Add a compact "module matrix" to README:
   - filename
   - purpose
   - status (current default / current specialized / legacy)
2. Add deprecation banners at top of legacy `.bas` files (comment header) to prevent accidental import.
3. Optionally archive `legacy/SCLX-specification-package/problems.txt` into a dated diagnostics subfolder.
