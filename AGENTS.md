# Global Codex instructions

Operate as a disciplined coding team. Prefer clear, minimal, tested changes. Never assume permission to modify unrelated repositories.

- Treat repository `AGENTS.md` as the source of project-specific requirements.
- Keep tasks, branches, worktrees, artifacts, and secrets isolated by repository.
- Never expose secrets, tokens, private keys, or sensitive source code in status reports.
- Do not modify project priority; only the owner may set it.
- For nontrivial work, write a brief plan, acceptance criteria, and test strategy first.
- Developer may automatically fix issues raised by code review, security, and QA and resubmit, for at most three correction cycles. Escalate unresolved findings.
- Never self-approve. Require independent review before requesting merge.
- Never merge to protected branches, release, deploy, or perform destructive operations without explicit owner approval.
- Run available tests and report results accurately; never claim tests passed without execution.
- Prefer least privilege, dependency pinning, and local-first development.
- Do not initiate continuous background activity unless a persistent runner has been explicitly configured.

## Linear (task tracking)

Every piece of work is tracked in Linear, team "Stephen Smitherman" (key STE). Stephen set these rules; follow them without waiting to be asked.

**Before you start**
- Search Linear for an issue that covers the work. If none exists, create it yourself, with every field below. Don't wait for approval to file it; Stephen changes scope or priority afterwards.
- If your session has no Linear tool, put the fields below in the PR body under a heading "Linear issue to file", so the next session or the weekly sweep files it.

**Every issue gets**
1. **Title:** `[Location] APP: outcome in plain words`, for example `[Cloud] SPM: requirements get a recurring periodicity`. Stories under a parent add a wave number after the app: `[Cloud] SPM 2a: ...`.
   - Location is where the work is performed: `[Powerspec]` (local work on the PowerSpec or the review VMs), `[Cloud]` (a cloud agent session), `[GitHub, you]` (a step only Stephen takes on GitHub, such as a merge), `[Hands-on, you]` (Stephen's physical or in-person step).
   - App short names: EG, ANCHOR, SABER, SPM, MG, HELM, BERTH, Web, Kit, Review servers, Compliance.
2. **Labels**, one from each group:
   - Location: Powerspec, Cloud, "GitHub, you", "Hands-on, you" (the same as the title tag).
   - App: Exam Generator, ANCHOR, SABER, SPM, Message Generator, HELM, BERTH, Web platform, Kit, Review servers, Compliance. If the app has no label, propose one to Stephen and use the closest until he answers.
   - Goal: Security, Modern UI, Fewer clicks, Quick help or Multi-user (the main goal the work serves).
   - Type: Feature, Improvement or Bug.
3. **Project:** Navy Web Apps for anything in the navy-web-apps repo; Navy Access Solutions for the Access apps and the kit; Anchor for the ANCHOR prototype repos. Cross-project process work (Compliance) has no project.
4. **Parent:** group a batch of related stories under one parent issue.
5. **Priority:** Stephen's order is EG, ANCHOR, SABER, SPM. Use what he said, or the priority of the parent or the app's other open issues. Never change a priority that is already set.
6. **Estimate:** story points 1, 2, 3, 5 or 8. Split anything that is 8 or more.
7. **Waves:** stories under a parent get wave numbers (1, 2a, 2b, 3). Mirror every dependency as a Linear "blocked by" relation.
8. **Description:** "As a [role], I want [capability], so that [benefit]", then acceptance criteria written as Given / When / Then.

**While you work**
- Name the STE key in every PR title or body, for example `STE-201`. A PR that fixes several issues names each one.
- Move the issue as the work moves: In Progress when you start, In Review when the PR is open, and add the PR link to the issue. Only Stephen moves an issue to Done after he merges or accepts it.
- Work you find along the way that is out of scope gets its own issue, not a line in the PR.
- Work threads for a story are titled `[app] [assignee] [STE-key] wave short title`, for example `[SPM] [Claude] [STE-159] 2a Recurring periodicity`.
