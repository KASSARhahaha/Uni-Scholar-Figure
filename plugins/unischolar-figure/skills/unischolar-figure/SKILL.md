---
name: unischolar-figure
description: Research a scientific topic, define the evidence and content of a paper's Figure 1 workflow or method schematic, derive an original journal-appropriate visual system from a cited published figure, draw the figure, and use FigEdit to deliver editable SVG and native PowerPoint. Use for requests about Figure 1 workflows, research-to-figure pipelines, journal-style scientific schematics, or editable PPT versions of scientific diagrams. Do not use for data-only charts without a workflow or schematic component.
metadata:
  short-description: Research and build editable Figure 1 workflows
---

# UniScholar Figure

Create Figure 1 as a defensible visual argument. Research and content logic come before styling. The final PowerPoint must expose ordinary labels, shapes, and connectors as editable objects.

## Intake gate

Require a research topic, manuscript section, proposal, dataset description, or source bundle. If none is available, ask one concise question requesting the topic or files and stop. Do not invent a scientific claim, method, result, citation, or quantitative value.

Record the intended journal when the user gives one. Otherwise use a provisional broad high-impact-journal format and say that journal-specific compliance remains pending.

Choose the smallest mode that satisfies the request:

- `full`: research through editable PPTX
- `research-only`: evidence ledger and Figure 1 content recommendation
- `blueprint-only`: node, arrow, panel, and style blueprint without drawing
- `convert-only`: FigEdit conversion of an approved source figure
- `audit-only`: scientific, originality, visual, and editability review

Do not force FigEdit or drawing work into a request that only asks for research or a blueprint.

## Workflow

1. Read [research and figure contract](references/research-and-contract.md). Build the evidence ledger and a one-sentence core conclusion before drawing.
2. Read [journal reference method](references/journal-reference.md). Select one published reference figure from the target journal or closest scientific field. Cite the source and extract only abstract design rules. Never trace or copy its scientific content, icons, composition, or data.
3. Write a compact content blueprint. Every node, arrow, branch, inset, and output must have a defined role in the argument. Mark hypotheses, planned work, and unsupported transitions explicitly.
4. Draw an original SVG or other high-resolution source figure. Keep a white background, one dominant reading direction, direct labels, restrained color roles, consistent connectors, and enough whitespace to remain legible at final journal size. Use lowercase panel labels for Nature-family conventions unless the target journal requires another style.
5. Render the source figure to a high-resolution PNG for visual review and FigEdit measurement. Preserve the original SVG and any source script separately.
6. Read the bundled `../figedit-v2/SKILL.md` completely, then use its measure, manifest, compose, and audit workflow. For a clean vector workflow diagram, choose the deterministic SVG route and avoid unnecessary OCR-driven asset extraction or paid image generation.
7. Read [FigEdit handoff](references/figedit-handoff.md) before running conversion. Keep all task files in the user's project, never in the plugin directory.
8. Read [QA and delivery](references/qa-and-delivery.md) before release. Render the final slide, compare it with the approved source figure, and verify native editability rather than relying on file extension alone.

## Non-negotiable boundaries

- Separate literature evidence, author-supplied facts, hypotheses, and visual analogy.
- A cited journal figure provides design vocabulary only. Its protected expression and scientific message stay out of the new artwork.
- `scipilot-figure-skill` contributes data-figure QA ideas but explicitly excludes flowcharts and schematics. Do not claim that it supplies the diagram generator.
- `figures4papers` provides examples and general visual principles under CC BY-NC 4.0. Do not copy its images, manuscript-specific labels, raw data, or substantial code into a submission figure.
- FigEdit reconstructs structure; it cannot verify the science. Scientific review must pass before conversion.
- Do not call a PPTX editable when it contains only one full-slide image or when ordinary labels and connectors cannot be selected independently.

## Deliverables

Use `scripts/init_figure1_workspace.py` to create a consistent project bundle when one does not already exist. Deliver:

- `research/research_brief.md`
- `research/evidence_ledger.csv`
- `design/figure_contract.md`
- `design/style_reference.md`
- source SVG plus a high-resolution preview
- the FigEdit output package, including `editable.pptx`, `editable.svg`, `manifest.json`, and quality reports
- `delivery/release_notes.md` with limitations and unresolved reviewer risks

Label the final verification state as `PPTX_STRUCTURAL_PASS`, `LIBREOFFICE_RENDER_PASS_ADVISORY`, or `POWERPOINT_NATIVE_PASS`. On macOS and Linux, FigEdit's automatic render uses LibreOffice and cannot establish native PowerPoint acceptance.

Run `scripts/validate_figure1_bundle.py <task-directory> --strict` before calling the bundle complete.

For provenance and upstream limitations, read [upstream research notes](references/upstream-notes.md).
