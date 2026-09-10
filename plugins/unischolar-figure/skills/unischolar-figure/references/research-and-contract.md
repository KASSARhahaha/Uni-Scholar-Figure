# Research and figure contract

Use this reference before drawing.

## Research bundle

Create a brief that states:

- the research question and intended audience
- the scientific gap or decision that motivates Figure 1
- supplied facts and their provenance
- literature-supported claims with source links or DOIs
- unresolved disagreements, missing evidence, and assumptions
- the conclusion Figure 1 can defend at the current evidence level

For current or niche claims, search the web and prefer primary literature, official datasets, standards, and the target journal's author guide. Record each source in `evidence_ledger.csv` with these columns:

```text
claim_id,claim_text,evidence_type,source_title,doi_or_url,publication_year,accessed_date,support_level,figure_element,notes
```

Use `evidence_type` values `author-supplied`, `primary-literature`, `official-guidance`, `hypothesis`, or `visual-reference`. Use `support_level` values `supports`, `partly-supports`, `contradicts`, `context-only`, or `pending`.

Each `figure_element` entry must name one or more real FigEdit panel or element IDs, separated by semicolons when needed. Each semantic manifest element of type `text`, `line`, `path`, or `polyline` must carry `evidence_ids: ["C1", ...]` or a specific `evidence_exempt_reason`. When its cited rows do not support the element, add `evidence_state: hypothesis`, `planned`, or `context` so the visual cannot silently promote unsupported content into a result.

## Figure contract

Complete this contract before choosing a layout:

```text
Core conclusion:
Figure role: motivation | workflow | mechanism | validation overview
Figure archetype: schematic-led composite | asymmetric mixed-modality
Target journal and article type:
Final width and height:
Reading direction:
Panel map:
  a:
  b:
  c:
Hero evidence:
Supporting evidence:
Controls or validation:
Source data needed:
Statistics needed:
Image-integrity notes:
Main reviewer risk:
```

Figure 1 usually establishes the paper's visual vocabulary. Use the same term, symbol, and color for the same concept throughout the figure. Remove any panel that does not carry unique information.

## Workflow grammar

Prefer a small number of semantic layers:

1. Research gap or system boundary
2. Inputs and provenance
3. Core method or experiment
4. Validation gate
5. Outputs and the decision enabled

This is a content grammar, not a fixed five-box template. Merge layers when the science is simpler. Add feedback loops or branches only when the method truly contains them. Every arrow must name a real transformation, measurement, decision, or dependency.

Mark planned or hypothetical steps with a distinct line style and explain that style in the legend. Do not let a visual arrow imply causality that the evidence does not establish.
