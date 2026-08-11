# Failures

## 2026-07-27: `git add <deleted-path>` aborted staging; deletion swept into wrong commit

- **What happened:** `base.py` was removed with `git rm` (staging the deletion), then later commits staged specific files with `git add fileA fileB ...`. When one `git add` included the already-deleted `base.py` path, the whole `git add` aborted with a pathspec error — but the previously-staged deletion stayed in the index and got committed with the *cardinality* fix instead of the intended tidiness commit.
- **Why it matters:** Broke one-concern-per-commit; the dead-file removal is now mixed into a bug-fix commit.
- **Fix / avoidance:** Stage a deletion in the same `git add` group as its intended commit, and run `git status` between commits to see what's actually staged before committing. Don't reference an already-deleted path in a later `git add`.
- **Tags:** git, workflow

## 2026-07-27: CI silently red for ~2 weeks because releases don't gate on it

- **What happened:** `pip-audit` in `ci.yml` had failed on every `main` push since v2.3.1 (2026-07-15) — the runner's `setuptools 79.0.1` is flagged by PYSEC-2026-3447. It went unnoticed because `publish.yml` (tag-triggered) doesn't run pip-audit, so v2.3.2 published green while CI was red.
- **Why it matters:** A red CI that doesn't block releases trains you to ignore CI. setuptools here is a build-env tool, not a datascope runtime dep, so the package was fine — but the next red could be real.
- **Fix / avoidance:** Fixed by upgrading setuptools in the CI install step. Check `gh run list --workflow=ci.yml` after pushes; consider having the publish workflow depend on CI passing.
- **Tags:** ci, release, dependencies

## 2026-07-31: v2.3.3 released with stale sample artifacts

- **What happened:** A session cut v2.3.3 (changed `html.py` print-block colors `#000`→`#0d0d0d`/`#ccc`→`#d9d9d9` and the `_palette.health_assessment_text` wording) but did not regenerate `samples/output/`. The committed HTML/PDF/Excel portfolio artifacts still carried pre-v2.3.3 output. Same failure mode as the earlier v2.2.0 freeze — samples silently drift because nothing regenerates and compares them.
- **Why it matters:** Samples are shown to prospects; a release shipping stale samples misrepresents the current tool. This is the second occurrence.
- **Fix / avoidance:** Added `tests/test_samples_fidelity.py` — regenerates HTML samples via `cli.main()` and diffs committed copies (fails on drift). PDF/Excel are not covered (binary/nondeterministic) — regenerate by hand on any report/palette/version change. Consider a release-checklist step: regenerate all samples before tagging.
- **Tags:** release, samples, process

## 2026-07-31: built on a stale base — no fetch at session start

- **What happened:** Started editing local `main` without `git fetch`. A concurrent session had already pushed v2.3.3 to origin. First `git push` was rejected; had to fetch + rebase mid-session.
- **Why it matters:** Building on a stale base risks conflicts and duplicated work — my palette-token cleanup could have collided with the concurrent "palette tokens" sweep (it happened to be disjoint this time).
- **Fix / avoidance:** `git fetch` at session start, especially on repos where multiple agent sessions run concurrently. Rebase unpushed local commits onto the updated remote — never force.
- **Tags:** git, workflow, process
