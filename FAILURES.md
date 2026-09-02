# FAILURES

What didn't work and why, so we don't repeat it. Newest on top.

## 2026-09-02 — Git-Bash `/c/...` path passed to Windows Python produced no output

**Tags:** #windows #paths #tooling

**What:** Regenerated a sample by passing an output dir as a Git-Bash-style
`/c/Users/.../scratchpad/regen` path to the datascope CLI under Windows Python.
The command exited without error but wrote no file; the follow-up read then
`FileNotFoundError`'d.

**Why:** Windows CPython does not resolve MSYS/Git-Bash `/c/...` mount paths. The
CLI treated it as a nonexistent relative dir and (with stderr suppressed) produced
nothing visible.

**Fix used:** Drove the CLI entry point directly in-process
(`from datascope.cli import main; main([...])`) with `tempfile.mkdtemp()`, which
yields a native Windows path. For shell invocations, use native `C:\...` paths or
`$(cygpath -w ...)`, and don't suppress stderr while debugging.
