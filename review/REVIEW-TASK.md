# Review task for the orchestrator

Requested by Stephen, 2026-10-26. Review only: make no code changes and push nothing except the reports below.

Review the branch `phase-10-logs-release`: it holds all 10 build phases (only Phase 6 is on `main` so far). Check it out on each VM before you start.

Run on the Windows review VMs:

1. **Code review (VM 1).** Pull this repo and review all code with `review/code-review-prompt.md`.
   Save the result as `review/reports/code-review-YYYY-MM-DD.md`.
2. **Visual QA (VM 2).** Pull this repo. In the `access` folder run:
   ```
   Remove-Item .\build\ExamGen_BE.accdb -ErrorAction SilentlyContinue
   .\backend\new-backend.ps1
   if ($?) { .\build.ps1 -BackEnd .\build\ExamGen_BE.accdb }
   ```
   The data file is made new each time: a reused one keeps the help text from the phase it was made in.
   Open `access\build\ExamGenerator.accdb` and test it with `review/qa-visual-review-prompt.md`.
   Save the result as `review/reports/qa-YYYY-MM-DD.md`, with screenshots in `review/reports/screenshots/`.

3. **Functional checks (VM 3).** Pull this repo and build as in step 2, on a VM with Word and Outlook. Run `review/functional-checks.md`.
   Save the result as `review/reports/checks-YYYY-MM-DD.md`.

Note: all 10 build phases are built, so no page should show a placeholder; one that does is a failure.

When all three reports are in, commit them on a branch named `review/YYYY-MM-DD` and tell Stephen the verdicts.
