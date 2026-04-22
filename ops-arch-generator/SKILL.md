---
name: ops-arch-generator
description: Generate customer deployment architecture diagrams in PowerPoint from uploaded operations inventory workbooks that follow the sheets 1-资源清单, 2-服务清单, and 3-容器应用. Use when a user wants a PPT architecture diagram matched to real servers, IPs, ports, pod layout, and resource sizing, optionally uploaded to cloud storage.
---

# Ops Arch Generator

## Overview

Use this skill when a customer provides an operations workbook and expects a deliverable architecture diagram in `.pptx` format instead of a text summary.

The workflow reads the workbook, merges infrastructure, service, and pod metadata, generates a customer-specific PowerPoint based on `references/arch-model.pptx`, emits a summary JSON, and can upload the generated deck to a cloud target.

## Required Inputs

The workbook should contain these sheets:

- `1-资源清单`
- `2-服务清单`
- `3-容器应用`

The parser expects the structure documented in [workbook-format.md](/home/indigo/myprj/ops-arch-generator/references/workbook-format.md).

## Workflow

1. Ask the user to upload the customer workbook if it is not already in the workspace.
2. Verify the workbook has the required sheets and roughly matches the sample layout.
3. Run the generator:

```bash
python3 /home/indigo/myprj/ops-arch-generator/scripts/generate_architecture_ppt.py \
  --workbook /absolute/path/to/customer.xlsx
```

4. Review the summary JSON in `outputs/` if there are anomalies or unmapped resources.
5. If cloud upload is configured, either pass `--upload` to the generator or call the uploader explicitly:

```bash
python3 /home/indigo/myprj/ops-arch-generator/scripts/upload_artifact.py \
  --file /absolute/path/to/output.pptx
```

## Output Contract

The workflow produces:

- A `.pptx` architecture diagram derived from `references/arch-model.pptx`
- A `.json` summary with parsed resources, services, pods, grouping decisions, and unmapped items

The generated diagram should include:

- Deployment zones
- Access path arrows
- Server and service labels
- IP addresses
- Access ports
- Resource sizing such as CPU, memory, and disk
- K8s node and pod placement
- Storage and middleware dependencies

## Generation Rules

- Prefer workbook facts over assumptions.
- Use `1-资源清单` for machine inventory and sizing.
- Use `2-服务清单` for service category, access ports, and service-to-service access hints.
- Use `3-容器应用` to enrich the K8s area with pod names, replica counts, pod sizing, exposed ports, and mount paths.
- If a service sheet port conflicts with a pod sheet port, keep both in the summary and prefer the pod sheet for pod-level labels.
- Never include passwords from the workbook in the output.

## Upload Rules

Cloud upload is generic by design. Supported modes are documented in [cloud-upload.md](/home/indigo/myprj/ops-arch-generator/references/cloud-upload.md):

- `copy`
- `webdav`
- `http-put`

If no upload target is configured, return the local output path instead of blocking.

## Files

- Template package: `/home/indigo/myprj/ops-arch-generator/references/arch-model.pptx`
- Manual sample output: `/home/indigo/myprj/ops-arch-generator/references/arch123.pptx`
- Sample workbook: `/home/indigo/myprj/ops-arch-generator/zyqd-test.xlsx`
- Generator: `/home/indigo/myprj/ops-arch-generator/scripts/generate_architecture_ppt.py`
- Uploader: `/home/indigo/myprj/ops-arch-generator/scripts/upload_artifact.py`
