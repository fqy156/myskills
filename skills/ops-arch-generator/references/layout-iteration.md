# Layout Iteration

Use this checklist when the user keeps correcting the generated deck.

## First Check

- Open the summary JSON before touching layout code.
- Confirm resource count, service count, pod count, and family count match expectations.
- Check whether the problem is parsing, grouping, or only presentation.

## Parse Problems

- If the workbook shape is unsupported, reject it instead of silently guessing.
- If a pod tab was missed, expand recognition by header semantics, not by fixed tab name or index.
- If a pod landed on a separate environment page, normalize the pod environment string before grouping.

## Layout Problems

- Prefer fixed anchors for major zones: access, application, data, platform.
- Keep pods inside the K8s cluster rectangle.
- When space gets tight, reduce per-card detail before changing zone order.
- Encode every accepted layout correction as a deterministic rule.

## Regression Checks

- Re-run the workbook that triggered the change.
- Re-run at least one previously good workbook.
- Confirm the new run does not drop pods, ports, or families.
- Confirm the generated `.pptx` passes ZIP and XML validation.

## Output Review

- The `.json` file is the canonical debug artifact.
- The `.pptx` is the customer deliverable.
- If both parsing and layout changed, state that explicitly in the changelog or release notes outside the skill.
