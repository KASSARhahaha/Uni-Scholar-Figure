# QA and delivery

Review the scientific argument, the visual source, and the editable package as separate gates.

## Scientific gate

- Every factual node maps to author-supplied material or an evidence-ledger row.
- Hypotheses and planned steps remain visibly distinct from validated results.
- Arrow direction and branch logic match the real method.
- Figure 1 does not imply a result that belongs only to later figures.
- Terminology and sample states remain consistent with the manuscript.

## Visual gate

- The core method dominates the reading hierarchy.
- The figure remains legible at the declared final size.
- Panel labels, gutters, line weights, and color roles stay consistent.
- Red and green are not the only distinction.
- The journal reference influenced abstract design tokens only.
- The source SVG/PDF retains editable text and the preview contains no clipping or unresolved placeholder text.

## PowerPoint gate

- Render the final PPTX and compare it with the approved source preview.
- Select representative labels, boxes, and connectors independently.
- Check text wrapping, font substitution, arrowheads, connector endpoints, and group depth.
- Inspect slide XML or use `validate_figure1_bundle.py` to detect a flattened-image-only export.
- Do not claim PowerPoint-native rendering was checked unless it was actually opened and inspected or rendered through the approved FigEdit path.

Use one explicit status:

- `PPTX_STRUCTURAL_PASS`: package and object-level checks passed, but no office renderer established visual fidelity
- `LIBREOFFICE_RENDER_PASS_ADVISORY`: LibreOffice rendered the slide without a structural defect; font and Office Math behavior in PowerPoint remain unverified
- `POWERPOINT_NATIVE_PASS`: the file passed a native Microsoft PowerPoint render and the required edit operations

FigEdit's bundled native PowerPoint renderer supports Windows. Its non-Windows path uses LibreOffice and states that the result is advisory. Never upgrade an advisory render to native verification.

## Release notes

Record:

- core conclusion and target journal
- reference figure citation and the abstract style tokens used
- research cutoff date
- source files and conversion path
- editability summary
- unresolved evidence, font, rendering, or journal-compliance risks
- whether native PowerPoint rendering was completed or deferred

Keep failed gates and warnings in the release notes. A visually polished draft remains a draft when its scientific or editability checks have not passed.
