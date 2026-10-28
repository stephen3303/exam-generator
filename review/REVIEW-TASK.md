# Review task for the orchestrator

Requested by Stephen, 2026-10-26. Review only: make no code changes and push nothing except the reports below.

Run on the Windows review VMs:

1. **Code review (VM 1).** Pull this repo and review all code with `review/code-review-prompt.md`.
   Save the result as `review/reports/code-review-YYYY-MM-DD.md`.
2. **Visual QA (VM 2).** Pull this repo. In the `access` folder run:
   ```
   .\backend\new-backend.ps1; .\build.ps1 -BackEnd .\build\ExamGen_BE.accdb
   ```
   Open `access\build\ExamGenerator.accdb` and test it with `review/qa-visual-review-prompt.md`.
   Save the result as `review/reports/qa-YYYY-MM-DD.md`, with screenshots in `review/reports/screenshots/`.

Note: all 10 build phases are built, so no page should show a placeholder; one that does is a failure.

When both reports are in, commit them on a branch named `review/YYYY-MM-DD` and tell Stephen the verdicts.
