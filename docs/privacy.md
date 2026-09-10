# UniScholar Figure Privacy Policy

**Effective date: 2026-09-10**

UniScholar Figure is maintained through the
[KASSARhahaha/Uni-Scholar-Figure](https://github.com/KASSARhahaha/Uni-Scholar-Figure)
repository. This policy explains the data boundary of the Codex plugin and
the separate local CLI and PowerPoint add-in.

## What this project operates

The project does not operate a UniScholar Figure telemetry, analytics,
advertising, or crash-reporting service. The Codex plugin is a bundle of
local skill instructions, references, and scripts; it has no dedicated
UniScholar Figure web backend.

## Codex plugin

The plugin can work from material supplied in the current Codex task, such as
a research topic, manuscript text, reference figure, or project files. Its
deterministic drawing and PowerPoint-export scripts operate on files in the
user's chosen project location.

When a user requests literature research, the Codex environment may access
public web sources. When a user expressly approves an image-generation or
clean-plate route, the selected source image and prompt may be sent to the
image service selected by the user or available in that Codex environment.
The plugin does not automatically send figures to a third-party image
service. Do not use optional online routes for sensitive, unpublished,
regulated, or confidential material unless the user has evaluated the
applicable service's terms and data controls.

Processing by Codex or a selected third-party provider is governed by that
provider's terms and privacy documentation. This project does not control or
make representations about those providers' retention, training, or security
practices.

## Local CLI and PowerPoint add-in

Local drawing, layout, and PowerPoint transformations run on the user's
device. Two optional local-tool features make outbound requests only after a
user action:

- **Check for Updates** requests the latest GitHub release from
  `api.github.com` after the user selects that ribbon action.
- **Research Records** sends a user-provided bearer token and selected query
  parameters to `uni-scholar.asia` after the user chooses that feature. The
  CLI stores that token in the user's local configuration directory with
  best-effort owner-only permissions; the PowerPoint add-in stores it in its
  documented local application-support path.

These optional services are not needed for the normal local layout commands.

## Data sharing and retention

The project maintainer does not receive a copy of local project files through
the plugin or local tools. Data may nevertheless be received by a service the
user intentionally accesses, including public websites, Codex, an image
provider, GitHub, or `uni-scholar.asia`. Review those services' policies
before using them. This project does not sell user data.

## Contact and changes

For questions or requests about this policy, open an issue at
[the project issue tracker](https://github.com/KASSARhahaha/Uni-Scholar-Figure/issues)
without posting access tokens, unpublished research, or other confidential
material. Material policy changes will be recorded in the repository history
and will carry a new effective date.
