#!/usr/bin/env python3
"""Lailara engagement deploy guard (Python, stdlib-only).

Exit 2 if an ACTIVE (non-demo) client engagement.yml is present in the current
directory. No-op otherwise, so demo builds and clean CI checkouts are unaffected.
Self-contained (no dependency on the installed lailara_engagement package) so it can
run in any repo's deploy/build environment.
"""
import os
import re
import sys

for _f in ("engagement.yml", "engagement.yaml"):
    if os.path.isfile(_f):
        with open(_f, encoding="utf-8-sig") as _fh:
            _txt = _fh.read()
        if re.search(r"^\s*demo:\s*true\s*$", _txt, re.M):
            continue  # demo config -> safe
        sys.stderr.write(
            f"ENGAGEMENT GUARD: active client engagement config present ({_f}). "
            "Client mode is runtime-only and must never deploy. Deactivate it "
            "(set 'demo: true', or use engagement.demo.yml) before deploying.\n"
        )
        raise SystemExit(2)
raise SystemExit(0)
