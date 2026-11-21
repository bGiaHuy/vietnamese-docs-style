# Reviewing changes

Keep changes small enough to review as a single concern. Separate document rules,
DOCX tooling, template assets, and CI changes when they can be reviewed independently;
keep a behavior change with the tests and documentation it requires.

For each change, describe the problem, the affected entry points, why the approach
fits the existing code, the checks actually run, and any unchecked behavior.
AI assistance does not replace understanding or review; contributors are responsible
for the submitted result. Do not claim checks that were not performed.

Use focused commits going forward. Preserve published history rather than splitting
old commits to imply a development process that did not happen.

## Where to review

- `SKILL.md`, `GEMINI.md`, and `references/`: document selection, rules, and workflow
  guidance (`GEMINI.md` remains at repository root for Gemini CLI compatibility).
- `scripts/`: executable DOCX creation, editing, validation, rendering, and packaging.
- `assets/`: rules and starter templates; preserve provenance and verification metadata.
- `tests/`: automated tooling checks; run `python -m pytest -q` after installing
  `requirements-dev.txt`.
- `evals/`: manual behavior evaluations, separate from automated tooling tests.
- `.github/workflows/`: CI checks. `docs/internal/VERIFICATION.md` records verification
  evidence; `docs/internal/IMPLEMENTATION-PLAN.md` records implementation planning,
  not proof of correctness.

Update user-facing documentation with behavior or path changes. Keep local runtimes,
generated output, and research data out of distributable packages. Do not assume
third-party source material inherits a future project license.
