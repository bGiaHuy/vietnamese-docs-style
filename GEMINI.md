# Repository Instructions: Vietnamese Academic Document Generator

## Project Goal
Generate standardized Vietnamese academic and administrative documents (.docx) following Decree 30/2020/NĐ-CP and FPT University styles.

## Key Standards
- **Decree 30/2020/NĐ-CP**: Found in `references/administrative-format-nd30.md`. This is the primary source of truth for administrative formatting.
- **Visual Style**: Black text only (#000000), Times New Roman font.

## Skill Profile System
- Use `references/document-profiles.md` to map user requests to appropriate profiles (e.g. `administrative`, `academic`, `proposal`, `minutes-administrative`, `minutes-general`).
- The canonical entry point is `SKILL.md` at the root of the repository.

## Python-Docx Helper
- Use `scripts/build_docx.py` as the implementation reference when writing code to generate Word files.
- Refer to `references/style-spec.md` for style configurations.

## Workflow for AI Agents
1. **Research**: Consult `SKILL.md` and then read the relevant reference documents based on the document profile.
2. **Implementation**: Use the `python-docx` library along with `scripts/build_docx.py`.
3. **Verification**: 
   - Run the structural validation script: `python scripts/validate_docx.py <file.docx> --profile <profile>`.
   - Cross check against `references/validation-checklist.md`.

## Prohibited Actions
- DO NOT use colors other than black.
- DO NOT use standard Word bullet points (`•`) for administrative documents; use manual prefixes (`-`, `+`, `*`) instead.
- DO NOT ignore the margin requirements for administrative docs (Left: 3.0-3.5cm, Right: 1.5-2.0cm, Top/Bottom: 2.0-2.5cm).
- DO NOT enforce rigid NĐ30 rules (like Quốc hiệu or manual bullets) onto `academic` or `custom` profiles unless the user explicitly requests it.
