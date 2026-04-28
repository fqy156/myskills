---
name: ops-arch-generator
description: Compatibility entry for the canonical ops-arch-generator skill now maintained under `skills/ops-arch-generator`. Use the canonical skill for workbook parsing, PowerPoint generation, and layout iteration. Keep this copy only to preserve old repo-local entrypoints.
---

# Ops Arch Generator

## Status

This directory is no longer the source of truth. The canonical, publishable skill now lives at [skills/ops-arch-generator](/home/indigo/myprj/skills/ops-arch-generator).

The scripts in `ops-arch-generator/scripts/` are compatibility wrappers that forward to the canonical skill so older commands keep working.

## Canonical Location

- Skill: [skills/ops-arch-generator](/home/indigo/myprj/skills/ops-arch-generator)
- Canonical instructions: [skills/ops-arch-generator/SKILL.md](/home/indigo/myprj/skills/ops-arch-generator/SKILL.md)
- Canonical generator: [skills/ops-arch-generator/scripts/generate_architecture_ppt.py](/home/indigo/myprj/skills/ops-arch-generator/scripts/generate_architecture_ppt.py)
- Canonical uploader: [skills/ops-arch-generator/scripts/upload_artifact.py](/home/indigo/myprj/skills/ops-arch-generator/scripts/upload_artifact.py)

## Compatibility Commands

These old commands still work, but they now forward to the canonical skill:

```bash
python3 /home/indigo/myprj/ops-arch-generator/scripts/generate_architecture_ppt.py \
  --workbook /absolute/path/to/customer.xlsx
```

```bash
python3 /home/indigo/myprj/ops-arch-generator/scripts/upload_artifact.py \
  --file /absolute/path/to/output.pptx
```

Prefer invoking the canonical scripts directly for any new automation or release packaging.
