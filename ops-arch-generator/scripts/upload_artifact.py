#!/usr/bin/env python3
"""Compatibility wrapper for the canonical ops-arch-generator uploader."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path


def main() -> int:
    current_file = Path(__file__).resolve()
    repo_root = current_file.parents[2]
    canonical_script = repo_root / "skills" / "ops-arch-generator" / "scripts" / "upload_artifact.py"
    if not canonical_script.exists():
        raise FileNotFoundError(f"Canonical uploader not found: {canonical_script}")

    sys.argv[0] = str(canonical_script)
    runpy.run_path(str(canonical_script), run_name="__main__")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
