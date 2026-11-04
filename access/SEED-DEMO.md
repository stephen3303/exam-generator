# Demo data (PowerSpec)

`tools\seed-demo.ps1` fills a new, empty data file with made-up people, topics, questions and exams, so every page has rows to show and two exams already have their Word documents. Use it to click through the app without typing a bank of questions first.

Everything it adds is invented: the people, the references (each marked "(demo)") and the questions. It never touches a data file that already has questions or exams.

It drives the app's own code, signed in as each made-up person in turn, so the same rules the pages enforce apply. If a rule stops it, it prints the app's own message and leaves the data file as it was.

Nothing here has been run in Access or Word yet. If it stops, send the PowerShell output.

## Run it

Close every Access window first. Then, from the `access` folder:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
Remove-Item .\build\ExamGen_BE.accdb -ErrorAction SilentlyContinue
.\backend\new-backend.ps1
.\tools\seed-demo.ps1
.\build.ps1 -BackEnd .\build\ExamGen_BE.accdb
```

`seed-demo.ps1` takes a few minutes (Word runs hidden for the two documents) and ends with `Seeded:` and a short report: who was added, the questions by status, the exams by status and where the Word documents went.

### You, the owner

The seed makes you the Exam Generator admin with the DH, XO and SM roles.

- **No Personnel file yet** (first time on this PC): it makes one in `access\Personnel`, with you as its admin, signed in by your Windows account. Pass `-OwnerFirst`, `-OwnerLast` and `-OwnerRank` to put your name on it (otherwise "Site Owner"). Your app password is the demo password.
- **Personnel file already there**: it finds you by your Windows account. If your Windows account isn't saved on your Personnel record, pass your sign-in name: `.\tools\seed-demo.ps1 -OwnerUser <your sign-in name>`. You must be a Personnel admin. Your password is left alone (set to the demo password only if you have none).

Other switches:

- `-BackEnd <file>` seeds another data file (default `build\ExamGen_BE.accdb`).
- `-DemoPassword <text>` changes the made-up people's password (default `ExamDemo-2026`, 8 characters or more).
- `-NoDocs` skips Word. **Download a copy** on the Exams page makes the documents later.

## Who's who

Every made-up person signs in with the demo password (`ExamDemo-2026` unless you changed it). Use **My settings**, then **Sign out**, to switch people.

| Sign-in name | Person | Exam Generator role | Topics |
|---|---|---|---|
| `dreyes` | CDR Dana Reyes | DH | |
| `mhale` | CAPT Morgan Hale | XO | |
| `tbrooks` | CMDCM Taylor Brooks | SM | |
| `achen` | LT Alex Chen | SME | Damage Control, Engineering Fundamentals |
| `srivera` | CWO2 Sam Rivera | SME | Damage Control, Engineering Fundamentals |
| `jpatel` | LTJG Jordan Patel | SME | Navigation, Communications and Security |
| `ckim` | ENS Casey Kim | SME | Communications and Security |
| `rford` | ENS Riley Ford | SME | Navigation (access to the app turned off) |

Emails are off for everyone, and each has a made-up `@example.com` address, so nothing goes to Outlook.

## What each page should show

**Signed in as you (admin, DH, XO, SM):**

- **Dashboard**: one question waiting at each of your three steps, and exam E3 waiting at the XO step. No topic is short of questions.
- **Approvals**: three questions (one at the DH step, one at the XO step, one at the SM step) and exam E3 at the XO step. Approve E3 at the XO step and then the SM step to watch the app make its Word document live, the same way the SM's approval does on the real chain.
- **Question bank**: 36 approved questions across four topics, multiple choice, true/false and short answer, some with parts A and B, three parts marked (CUI). Use the status chips to find the drafts, the waiting ones, the rejected one and the retired version 1. Open **History** on "Water from a fire hose..." to see version 1 retired by version 2.
- **Exams**: five exams.
  - E1, USS Harborview: approved, with its Word document.
  - E2, USS Ridgeline: approved, with its Word document; one damage control question was swapped first.
  - E3, Damage Control School: approved by the DH, waiting on the XO.
  - E4, USS Cedar Point: sent back by the DH to its damage control SMEs, with a reason.
  - E5, Engineering Training Unit: still with its SMEs; engineering signed off, damage control not.
- **Statistics**: on 6 months, the questions each topic gave on the two approved exams (E1 and E2).
- **Admin: People and routing**: the eight made-up people with their roles; Riley Ford shown as turned off.
- **Admin: Topics**: four topics with their SMEs, and "Legacy Seamanship (turned off)".
- **Admin: Templates**: Basic Qualification Board (9 questions: DC 3, NAV 2, ENG 2, COM 2) and Damage Control Refresher (4: DC 3, ENG 1).
- **Admin: Logs**: the audit trail, with dates spread over the last 80 days. The error log should be empty.

**Signed in as a made-up person:**

- `achen` **Dashboard**: one question sent back by the XO (the Class D fire question), with the XO's reason. **My reviews**: E4's damage control topic, sent back by the DH, and E5's damage control topic, both open (Alex Chen and Sam Rivera share them).
- `jpatel` **Question bank**: a draft (the Mercator question) and a version 2 in draft of "How many minutes of latitude...". Version 1 stays in the bank until version 2 is approved.
- `ckim`: a draft (the garbled message question).
- `dreyes`, `mhale`, `tbrooks`: the same Approvals list as yours, each at their own step.

## The Word documents

They go to `build\Exam documents`, next to the data file. Open E1's and E2's and check the cover page, the topics, the questions with their portion marks, and the answer key with the parts.

If a document fails, the approval still stands and the report prints a `FAILED:` line with Word's message. Open Word once by hand and close it before the seed, so a first-run prompt doesn't stall it.

## Starting over

The seed only fills an empty data file. To run it again, delete `build\ExamGen_BE.accdb` and run the five commands above. The made-up people stay in the Personnel file (they're reused, with their passwords as they are); delete them from Personnel if you don't want them.
