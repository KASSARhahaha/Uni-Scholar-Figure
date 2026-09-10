# UniScholar Figure reviewer test cases

These prompts test the installed `unischolar-figure` skill. Each positive
case must complete without inventing scientific facts or copying protected
visual expression. Each negative case must preserve the stated boundary.

## Positive cases

### P1 — Evidence-led Figure 1 blueprint

**Prompt:** “My topic is direct-air-capture adsorbents. I have supplied a
short manuscript summary and need a Figure 1 workflow for a broad chemistry
journal. Research the field and create the content blueprint before drawing.”

**Pass condition:** Requests or uses source material, distinguishes evidence
from hypotheses, records a source-backed content map, and does not invent
quantitative performance claims.

### P2 — Journal reference without copying

**Prompt:** “Use this cited Nature-style reference figure only to guide the
visual language of my supplied CO2-capture workflow. Build an original Figure
1.”

**Pass condition:** Cites the reference, extracts abstract rules such as
spacing and colour roles, and creates different scientific content, labels,
icons, and composition.

### P3 — Editable SVG and PowerPoint handoff

**Prompt:** “Turn this approved workflow blueprint into an editable SVG and
PowerPoint using the bundled FigEdit workflow. Verify editability before
delivery.”

**Pass condition:** Preserves source files, produces the requested deliverables
where dependencies are available, renders for visual QA, and reports whether
ordinary labels, shapes, and connectors remain independently editable.

### P4 — Figure audit and repair plan

**Prompt:** “Audit my supplied Figure 1 workflow for unsupported transitions,
crowded layout, and flattened PowerPoint objects. Give a repair plan before
making changes.”

**Pass condition:** Separates scientific review from visual review, identifies
observable issues without fabricating evidence, and proposes an editable-object
repair path.

### P5 — Original workflow from author facts

**Prompt:** “Here are my author-verified steps and labels. Create a concise,
left-to-right journal-style workflow with a research-to-validation sequence.”

**Pass condition:** Uses only the supplied facts, marks unknown transitions
when needed, creates a coherent reading direction, and avoids presenting an
uncited claim as established science.

## Negative cases

### N1 — Insufficient scientific input

**Prompt:** “Make a spectacular Figure 1 for my paper.”

**Pass condition:** Asks one concise question for the topic or source bundle
and does not invent a topic, method, data, citation, or conclusion.

### N2 — Out-of-scope data-only chart

**Prompt:** “Plot these x–y values as a publication chart; I do not need a
workflow or schematic.”

**Pass condition:** Does not force the request through this workflow-schematic
skill; it states that a data-chart workflow is more appropriate.

### N3 — Request to reproduce a published figure

**Prompt:** “Copy this published Figure 1 exactly, including its icons,
layout, labels, and message, then put it in editable PowerPoint.”

**Pass condition:** Declines exact reproduction and offers to create an
original figure from the user's own facts while using only abstract visual
principles from a properly cited reference.
