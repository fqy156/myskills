---
name: ops-arch-generator
description: Generate deployment architecture diagrams in PowerPoint from operations workbooks and iteratively refine the output against user feedback. Use when Codex needs to read `.xlsx` infrastructure inventories, service sheets, and container/pod sheets, produce a `.pptx` architecture deck plus summary JSON, reject unsupported workbook shapes such as planning-only resource summaries, or repeatedly adjust K8s/pod layout, grouping, labels, ports, and cross-zone connections.
---

# Ops Arch Generator

## Overview

Use this skill to turn a customer workbook into a deliverable `.pptx` deployment architecture diagram and a machine-readable `.json` summary. Prefer this skill when the user is iterating on the same deck and wants parsing rules, layout rules, and regression checks applied consistently.

## Workflow

1. Verify the workbook shape against [references/workbook-format.md](references/workbook-format.md).
2. Reject unsupported resource inputs before changing code.
3. Run the generator.
4. Inspect the summary JSON for grouping, ports, pods, and unmatched resources.
5. When the user requests layout changes, encode the rule in code or references so the next run is stable.
6. Re-run on the same workbook and at least one sample workbook to catch regressions.

## Accepted Inputs

- Standard resource inventory workbooks with machine-level facts.
- Optional service sheets for topology and port labeling.
- Optional container sheets whose tab name or content indicates container or pod data.

Do not accept planning-only or resource-summary sheets as substitutes for machine inventory. If the workbook only contains planned totals without machine names or IP-like endpoints, stop and say the workbook is unsupported rather than trying to infer deployment topology.

## Run

Use the packaged generator:

```bash
python3 scripts/generate_architecture_ppt.py \
  --workbook /absolute/path/to/customer.xlsx
```

Use `--title`, `--deck-name`, `--output-dir`, and `--upload` only when the user asks for overrides.

## Iteration Rules

- Treat every user correction as a reusable rule, not a one-off patch.
- Prefer workbook facts over visual symmetry.
- Keep unsupported inputs rejected unless the user explicitly changes the product scope.
- When pod sheets vary, match them by content and headers, not by fixed tab index.
- Keep K8s pod blocks inside the K8s region and validate that pod rows land in the same environment page as their cluster resources.
- After layout changes, re-run at least one known-good workbook and compare counts in the summary JSON.

## Outputs

- `outputs/<deck-name>.pptx`
- `outputs/<deck-name>.json`

The JSON is the primary debugging artifact. Read it before adjusting layout logic.

## References

- Workbook rules: [references/workbook-format.md](references/workbook-format.md)
- Upload modes: [references/cloud-upload.md](references/cloud-upload.md)
- Iteration checklist: [references/layout-iteration.md](references/layout-iteration.md)

## Bundled Files

- Template: `references/arch-model.pptx`
- Manual example deck: `references/arch123.pptx`
- Generator: `scripts/generate_architecture_ppt.py`
- Uploader: `scripts/upload_artifact.py`
