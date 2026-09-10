# UniScholar Figure public-plugin submission package

This folder is the release-facing package for a **skills-only** public Codex
plugin submission. It records what can be supplied to the submission portal;
it does not submit the plugin or claim OpenAI approval.

## Listing record

| Field | Submission value |
| --- | --- |
| Listing name | UniScholar Figure |
| Publisher | KASSARhahaha, pending verification in the OpenAI Platform portal |
| Category | Productivity |
| Plugin type | Skills only |
| Short description | Research, design, draw, and export editable Figure 1 workflows. |
| Website | https://github.com/KASSARhahaha/Uni-Scholar-Figure |
| Support | https://github.com/KASSARhahaha/Uni-Scholar-Figure/blob/main/docs/support.md |
| Privacy policy | https://github.com/KASSARhahaha/Uni-Scholar-Figure/blob/main/docs/privacy.md |
| Terms of use | https://github.com/KASSARhahaha/Uni-Scholar-Figure/blob/main/docs/terms.md |
| Logo source | `docs/assets/unischolar-figure-logo.svg` |

### Long description

UniScholar Figure turns a scientific topic or manuscript into an original,
evidence-backed Figure 1 workflow. It creates a claim and evidence ledger,
uses a cited journal figure only for abstract design principles, and produces
a journal-appropriate workflow diagram. The bundled FigEdit workflow exports
editable SVG and PowerPoint deliverables, with checks that labels, shapes, and
connectors remain independently editable.

## Submission gates

- [ ] The OpenAI Platform organization has a verified developer or business
  identity and the submitter has Apps Management write access.
- [ ] The portal listing uses the values above and uploads the logo.
- [ ] The final bundle is created from this reviewed commit and passes plugin
  validation.
- [ ] The reviewer exercises all cases in
  [`../plugin-review-test-cases.md`](../plugin-review-test-cases.md).
- [ ] The release notes and availability choices have been confirmed by the
  publisher.
- [ ] The publisher, not this repository, completes every platform policy
  attestation.

## Included files

- [`../privacy.md`](../privacy.md): public privacy policy.
- [`../terms.md`](../terms.md): use terms and licensing boundary.
- [`../support.md`](../support.md): support and issue-reporting route.
- [`../plugin-review-test-cases.md`](../plugin-review-test-cases.md): five
  positive and three negative reviewer cases.
- [`../release-notes/codex-plugin-v0.1.0.md`](../release-notes/codex-plugin-v0.1.0.md):
  initial plugin release notes.

## Scope boundary

This submission package covers the Codex plugin in
`plugins/unischolar-figure/`. The repository also contains a separate Python
CLI and PowerPoint add-in. Their optional online functions are explained in
the shared privacy policy but are not required to use the Codex plugin.
