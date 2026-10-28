# Functional checks (VM 3)

Run these on a Windows VM with Access, Word and Outlook (desktop, with a mail profile). Build first, as in REVIEW-TASK.md step 2. Report PASS or FAIL for each item, with what you saw, and screenshots where it helps. Make no code changes.

1. **Upgrade.** Build a data file from the `main` branch's `backend\new-backend.ps1` (data version 1), then open it with this build. It should upgrade to version 4 without errors.
2. **Upgrade while another copy is open.** Open the data file in a second copy of the app as an admin, then start the upgrade from the first copy. The upgrade should wait and say why, not fail with a raw error.
3. **Word page markings.** Approve an exam long enough to fill three pages, with questions of different classification levels. Download it. Each page should carry its own highest marking, and the cover and body should keep their formatting.
4. **Outlook.**
   - When Outlook asks to allow the send, try it once allowing the send and once refusing it. Each time the app should say what happened.
   - With Outlook closed, take a step that sends notices. The emails should leave the Outbox.
   - Give one person in a batch a bad email address. Only that email should fail, and the rest should still go.
5. **Toasts.** Any warning toast with a long note should not be cut off.
6. **Page timers.**
   - Open My reviews and Exams while a dialog is open. Their opening emails should go after the dialog closes.
   - Leave either page straight away. The emails should still go.
7. **Release build.** Make the locked .accde and click through every page.
8. **Replaced questions.**
   - Generate an exam, then approve a newer version of one of its questions. Sign off on that topic should refuse until an SME swaps the question.
   - If that topic was already signed off, it should go back to its SMEs (with an email) when the exam's last topic is signed off. The exam must not reach the DH.
9. **Reviews for new SMEs.**
   - Add an SME to a topic of an exam in SME review, then open My reviews as that SME. They should have a review.
   - Repeat with two copies at the same moment: one signs off the topic while the other opens My reviews. No open review should be left on a topic that is already signed off.
10. **Speed.** With three or more exams in SME review, time how long My reviews and Exams take to open.
11. **Sign out** (My settings; the feature-sign-out branch).
   - Sign out and sign in as another person. The app should reopen on their Dashboard with only their pages.
   - Sign out and close the sign-in screen. The app should close.
   - Sign out and sign in as someone with no Exam Generator role. The app should say why and close.
   - As a Windows-only user with no app password, Sign out should ask you to choose one first. Signed in by password on someone else's Windows account whose owner has no app password, Sign out should refuse.
