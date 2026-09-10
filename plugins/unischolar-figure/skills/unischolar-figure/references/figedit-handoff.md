# FigEdit handoff

The plugin contains a pinned FigEdit snapshot at `../figedit-v2/`. Read that skill and its routed references before execution.

## Environment check

From the `unischolar-figure` skill directory, run:

```bash
python3 scripts/check_figedit_environment.py
```

The default checks the deterministic SVG route used for an original workflow diagram. For OCR, arbitrary raster reconstruction, or generated assets, run `python3 scripts/check_figedit_environment.py --mode full`. Add `--require-math` when editable equations are present. The checker reports the bundled snapshot path and missing Python modules. It does not install packages. Ask before installing dependencies in the user's environment.

## Project location

Run FigEdit commands from the user's project directory. A normal task bundle uses:

```text
figure1-task/
  research/
  design/
  source/
  figedit/
    work/
    out/
  delivery/
```

Do not write measurements or generated output inside the plugin.

## Stable command entry point

Resolve this skill's directory, then use its wrapper while keeping the user's project as the working directory:

```bash
python3 <RESEARCH_FIGURE1_SKILL>/scripts/figedit.py prepare source/figure1.png --init figedit
python3 <RESEARCH_FIGURE1_SKILL>/scripts/figedit.py compose figedit/manifest.json --out figedit/out --stage svg
python3 <RESEARCH_FIGURE1_SKILL>/scripts/figedit.py compose figedit/manifest.json --out figedit/out --stage pptx
```

The wrapper resolves the vendored FigEdit scripts relative to itself and forwards the remaining arguments unchanged. Use `measure`, `draft-elements`, `manifest-edit`, `validate-manifest`, `quality-audit`, `editability-audit`, or `render-pptx` for the other documented entry points. Do not use a bare `python scripts/...` command from the user's project.

## Deterministic workflow-diagram route

For an original flowchart built from basic shapes and editable text:

1. Freeze the scientific content and review the high-resolution PNG.
2. Create the FigEdit task with the PNG as the source image.
3. Record `reconstruction_plan` as a deterministic SVG route with no generated assets unless the artwork genuinely contains non-geometric imagery.
4. Build or adopt the manifest, then compose the SVG stage.
5. Review the fix worklist and preview. Correct the manifest in batches.
6. Export the PPTX stage only after the SVG matches the approved source.
7. Run editability and quality audits. Render PowerPoint only when the FigEdit validation tier calls for it and user authorization permits attaching to an active PowerPoint session.

The authoritative arguments and safety rules live in the bundled FigEdit `SKILL.md` and `references/scripts.md`. The wrapper above only fixes path resolution; it does not change FigEdit behavior.

## Editable-output gate

The output passes only when:

- ordinary labels are editable text
- workflow boxes and simple symbols are editable shapes
- connectors remain separately selectable and point to the intended nodes
- the slide is not a single flattened screenshot
- grouped objects use meaningful, limited grouping
- the visual comparison shows no missing scientific content or reversed arrows

Preserve the original source SVG, preview PNG, FigEdit manifest, and reports next to the PPTX. FigEdit changes representation, not scientific authority.
