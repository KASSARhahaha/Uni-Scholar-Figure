# Uni-Scholar Figure — Provenance & Trademark Notice

## Clean-room reimplementation

Uni-Scholar Figure is a **behavior-level clean-room reimplementation** of the
commercially released **DakeSCI 1.2** PowerPoint add-in (sold at
`dakekyht.top/plugin`). The implementation is based solely on:

- the **public release-notes feature list** (7 bullet points describing
  user-visible behavior),
- independent application of `python-pptx`, `Pillow`, PowerPoint VBA,
  and GDI+ flat APIs.

## What was NOT done

The following were explicitly avoided to keep this work clean-room:

- ❌ No decompilation, disassembly, or reverse engineering of the original
  `.exe` / `.dll` / `.ppam` binaries.
- ❌ No extraction or reuse of any proprietary asset from the original
  product — including but not limited to: icons, ribbon images, VBA source,
  installer scripts, UI strings beyond the 7 generic feature names.
- ❌ No network probing of the original license server / 兑换码 / 机器码 system.
- ❌ No circumvention of any technical protection measure (TPM) — the original
  product's licensing flow is out of scope and not reproduced.

## Trademarks

"DakeSCI" and "dakekyht" are trademarks of their respective owner.
Uni-Scholar Figure is **not affiliated with, endorsed by, or derived from**
the original product or its author. All product names, logos, and brands
referenced in the documentation are property of their respective owners and
are used here only for factual, comparative identification of behavior.

## Icon catalog

All icons shipped under `src/unisfigure/icons.py` are hand-traced primitive
geometric paths (SVG `d` strings authored from scratch). No third-party
icon font or image library is bundled.

## External dependencies

- `python-pptx` — MIT License
- `Pillow` — HPND License
- `typer` — MIT License

These are standard open-source libraries; their inclusion does not
constitute reuse of proprietary code.

## Authorship & contact

Maintained at `https://github.com/KASSARhahaha/Uni-Scholar-Figure`.
