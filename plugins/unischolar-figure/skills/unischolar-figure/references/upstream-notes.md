# Upstream research notes

Research snapshot: 2026-09-10.

## scipilot-figure-skill

Source: <https://github.com/Haojae/scipilot-figure-skill>

Reviewed commit: `43098ddb9e6a6d142218540c114f9ed38922fc42`.

The repository is a data-figure skill for line, scatter, heatmap, and multi-panel quantitative plots. Its own scope excludes schematics, flowcharts, and architecture diagrams. This plugin adapts only general quality ideas: one figure, one conclusion; journal-aware dimensions; editable SVG text; multi-format export; automated checks; and render-review-revise loops. It does not copy a diagram generator because the source repository has none. The upstream license is MIT.

## figures4papers

Source: <https://github.com/ChenLiu-1996/figures4papers>

Reviewed commit: `3c181f85e82c6f24948fcaaf3be6696102b41d8d`.

The repository contains project-specific Python scripts and paper figures for bars, trends, heatmaps, radar plots, conceptual panels, and related chart families. It is not an importable plotting package. This plugin uses independently written, general design principles such as restrained palettes, direct labels, clean spines, asymmetric hierarchy, and vector export. It does not include upstream images, raw data, manuscript-specific layouts, or substantial code. The upstream repository uses CC BY-NC 4.0, so copying its protected expression into commercial or publication workflows may require additional permission and attribution.

## FigEdit

Source: <https://github.com/giszzt/figedit>

Bundled snapshot commit: `acc4f5c01506fe10dad99ca755a9612d67de5467`.

FigEdit converts flattened graphics into editable SVG and native PowerPoint through measurement, routing, manifest-based assembly, and quality checks. The bundled snapshot excludes large example images but keeps the skill instructions, scripts, references, templates, dependency list, and licenses. FigEdit's own code uses the MIT License. Its SVG-to-PPTX layer includes code adapted from PPT Master under MIT terms. See the plugin `licenses/` directory.

This plugin distinguishes FigEdit from Adobe Research's separate FigEdit benchmark for scientific chart editing. The bundled tool is the `giszzt/figedit` reconstruction skill because it is the project that exports editable PowerPoint files.
