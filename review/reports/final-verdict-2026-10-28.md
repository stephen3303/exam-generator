# Final verdict, exam-generator, 2026-10-28 (VM #1, VREVIEW-01)

Final confirmation build: PR 3 head **74e976a** (earlier builds noted per item). Fresh data file, build and second compile clean; error_log empty; no VBA or Access dialogs.

## Demo path at 74e976a: PASS
Scripted walk (calls the app's own functions as each person; Short Answer questions): 3 questions authored by SME Demo and approved DH, XO, SM; Generate exam; Swap; Sign off; exam approved DH, XO, SM; Word document made; Download a copy; History. The one-page demo exam opens in Word as 2 pages (cover, then the questions): UNCLASSIFIED header and footer on both, "(U)" on each question, no answers ("Reference answer" absent), no blank page.

## Results of the remaining checks
| Check | Commit | Result |
|---|---|---|
| Locked .accde: all 13 pages open, no dialogs | de98b7e | PASS (opened in code, not every control clicked) |
| Data upgrade v1 to v4 incl. backup (v1 file from commit d545d46) | de98b7e | PASS: backup made first and complete; version 4; no checks skipped |
| Upgrade with another copy holding the tables | de98b7e | PASS: friendly "close it on every other computer" message, version stays 1. Minor: if the file is held exclusively, a raw "Run-time error 3045" appears |
| Per-page markings, 12-question mixed-marking exam | de98b7e, f076809 | **FAIL (Major)**: every page header is SECRET (the overall marking), toast "Word couldn't split the pages cleanly"; the next-page-break change (f076809) did not fix it and was reverted. Over-marks, so on the safe side. Demo with a one-page exam |
| Replaced question blocks Sign off until swapped | de98b7e | PASS |
| Already signed-off topic with a replaced question goes back to its SMEs when the last topic is signed off; exam does not reach the DH | 74e976a | PASS (topic's sign-off cleared, new open review task, exam stays in SME Review). The email to the SMEs was not tested (no Outlook) |
| New SME added to a topic gets a review | de98b7e | PASS (0 before opening My reviews, 1 after) |
| Who can reach which page, per role | de98b7e | PASS |
| Validation (empty, bad points, long text, no reference, quotes/HTML, duplicate topic) | de98b7e | PASS |
| Overlap slider label | de98b7e | PASS (30%, follows the slider) |
| Resize | de98b7e | Minor: 1100 px Question bank detail panel overlaps its card edge; 900 px Approvals and Rules are cut off on the right |

## Not tested
125% and 150% display scaling (needs a display-setting change on the VM), page timers (My reviews and Exams emails after a dialog closes or after leaving the page), two copies at once (sign-off racing My reviews), speed with 3+ exams, keyboard order and Enter/Esc, double-click Save, delete confirmations, a click through the .accde with real mouse input, and everything that needs Outlook (allow/refuse send, Outlook closed, bad address, toasts with long notes).

## Open issues to know about for the demo
1. Per-page markings fall back to the overall marking on multi-page exams (Major, STE-180 open): use a one-page exam in the demo.
2. Narrow windows (about 1100 px and under) clip the right-hand panels (Minor): demo maximized.
3. Raw error if the data file is held exclusively during an upgrade (Minor, unlikely).

## Addendum: 125% display scaling (de98b7e build, run on 74e976a data)
Method: set the display scale to 125% through the Windows API (no sign-out was needed; signing out would have ended this review session), launched Access fresh, opened Dashboard, Exams, Approvals, Questions, My settings, Rules and People, then restored 100% (screen back to 1920 x 1200, AppliedDPI 96). No dialogs, no errors.
- The app scales with the display: text and controls are 25% bigger and nothing is unreadable.
- At 125% a 1920 px screen is about 1536 logical px wide, wider than the 1500 px design size, so a maximized window should fit; I could only capture the left 1536 physical px of the screen, so I did not see the right edge of a maximized window (Approvals and Rules looked correct as far as visible).
- **Minor/Major (same cause as the narrow-window issue):** in a window about 1180 px wide at 125% (about 944 logical px), **Question bank breaks**: the detail panel's text starts left of its card and runs under the list ("pproved by Demo..." clipped), "Q4 v1" is drawn over the list's last column header, and labels such as MARKING are cut ("ARKING", "FIED"). Screenshot: eg-dpi-shots/dpi125c-pgQuestions.png. Demo maximized, or at 100% scaling.
- 150% not tested.
