# Demo data (PowerSpec)

`tools\seed-demo.ps1` fills a new, empty data file with made-up people, topics, questions and exams, so every page has rows to show: about 100 questions in every status and at every approval step, and 13 exams in every status and at every step, four of them approved with their Word documents. Use it to click through the app without typing a bank of questions first.

Everything it adds is invented: the people, the references (each marked "(demo)") and the questions. It never touches a data file that already has questions or exams.

It drives the app's own code, signed in as each made-up person in turn, so the same rules the pages enforce apply. If a rule stops it, it prints the app's own message and leaves the data file as it was. The Personnel file isn't copied, so people it already added there stay; a second run reuses them.

The first, smaller seed passed on review VM #1 on 2026-11-06. The larger seed (six topics, 13 exams) hasn't been run yet. If it stops, send the PowerShell output.

## Run it

Close every Access window first. Then, from the `access` folder:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
Remove-Item .\build\ExamGen_BE.accdb -ErrorAction SilentlyContinue
.\backend\new-backend.ps1
.\tools\seed-demo.ps1
.\build.ps1 -BackEnd .\build\ExamGen_BE.accdb
```

`seed-demo.ps1` takes several minutes (Word runs hidden for the four documents) and ends with `Seeded:` and a short report: who was added, the questions by status, the exams by status, what waits at each step and where the Word documents went.

### For the web app (Database Files)

The web app reads the same data file, so seeding it fills the Exam Generator pages there too. Stop the web server (and the demo server) first, since they hold the Personnel file open, then seed the web's empty data file:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\tools\seed-demo.ps1 -BackEnd "C:\Users\steph\Projects\Database Files\ExamGen_BE.accdb" -NoDocs
```

It finds the Personnel file in the `Personnel` folder beside the data file (pass `-Personnel <file>` for another one) and adds the made-up people there, so the demo sign-in page lists them with their Exam Generator roles. `-NoDocs` skips Word; leave it off to make the four Word documents as well. The data file must still have no questions and no exams.

### You, the owner

The seed makes you the Exam Generator admin with the DH, XO and SM roles.

- **No Personnel file yet** (first time on this PC): it makes one in `access\Personnel`, with you as its admin, signed in by your Windows account. Pass `-OwnerFirst`, `-OwnerLast` and `-OwnerRank` to put your name on it (otherwise "Site Owner"). Your app password is the demo password.
- **Personnel file already there**: it finds you by your Windows account. If your Windows account isn't saved on your Personnel record, pass your sign-in name: `.\tools\seed-demo.ps1 -OwnerUser <your sign-in name>`. You must be a Personnel admin. Your password is left alone (set to the demo password only if you have none).

Other switches:

- `-BackEnd <file>` seeds another data file (default `build\ExamGen_BE.accdb`).
- `-Personnel <file>` uses that Personnel file. Default: `Personnel\Personnel_be.accdb` beside the data file when there is one, else the usual place (a `Personnel` folder beside the data file's folder).
- `-DemoPassword <text>` changes the made-up people's password (default `ExamDemo-2026`, 8 characters or more).
- `-NoDocs` skips Word. **Download a copy** on the Exams page makes the documents later. Use it if Word stalls the run.
- `-NoLineNumbers` loads the code without line numbers, as in `build.ps1`.

## Who's who

Every made-up person signs in with the demo password (`ExamDemo-2026` unless you changed it). Use **My settings**, then **Sign out**, to switch people.

| Sign-in name | Person | Exam Generator role | Topics |
|---|---|---|---|
| `dreyes` | CDR Dana Reyes | DH | |
| `mhale` | CAPT Morgan Hale | XO | |
| `tbrooks` | CMDCM Taylor Brooks | SM | |
| `achen` | LT Alex Chen | SME | Damage Control, Engineering Fundamentals, Seamanship |
| `srivera` | CWO2 Sam Rivera | SME | Damage Control, Engineering Fundamentals, Hazardous Materials |
| `jpatel` | LTJG Jordan Patel | SME | Navigation, Communications and Security, Hazardous Materials |
| `ckim` | ENS Casey Kim | SME | Communications and Security, Seamanship |
| `rford` | ENS Riley Ford | SME | Navigation (access to the app turned off) |

Emails are off for everyone, and each has a made-up `@example.com` address, so nothing goes to Outlook.

## What each page should show

The numbers below assume a new data file, so the exams are E1 to E13.

**Questions** (about 100):

| Status | How many | Which |
|---|---|---|
| Approved | 80 | the bank: Damage Control 18, Navigation 14, Engineering 14, Communications 14, Seamanship 10, Hazardous Materials 10; multiple choice, true/false and short answer, easy to hard, some with parts A and B, three parts marked (CUI) |
| In Review, at the DH | 6 | the soundings, Class A fire, snap-back zone and storeroom issuer questions; the safety data sheet question (sent back by the DH, fixed and sent again); version 2 of "What does a gauge reading of zero psig..." |
| In Review, at the XO | 3 | the password, portable equipment check and "avast" questions |
| In Review, at the SM | 3 | the relief valve, gyro repeater and aerosol can questions |
| Rejected | 3 | by the DH (the anchor question), by the XO (the Class D fire question) and by the SM (the message numbers question), each with its reason |
| Draft | 7 | one in every topic but Navigation, which has the Mercator question and version 2 of "How many minutes of latitude..." |
| Retired | 3 | version 1 of the water, light list and oily rags questions, each replaced by an approved version 2 |

**Exams** (13):

| Exam | Template | Where it stands |
|---|---|---|
| E1 USS Harborview | Basic Qualification Board | Approved, with its Word document |
| E2 USS Ridgeline | Basic Qualification Board | Approved, with its Word document; one damage control question swapped first |
| E3 USS Bayfield | Deck and Hazmat Qualification | Approved, with its Word document |
| E4 Surface Warfare School | Comprehensive Board | Approved, with its Word document |
| E5 USS Kestrel | Comprehensive Board | In Review, at the DH |
| E6 Damage Control School | Damage Control Refresher | In Review, at the XO |
| E7 USS Glenmore | Deck and Hazmat Qualification | In Review, at the SM |
| E8 Fleet Training Center | Damage Control Refresher | In Review, at the XO, version 2: the DH sent its engineering topic back, an SME swapped the question and signed off again |
| E9 USS Cedar Point | Basic Qualification Board | Rejected by the DH: damage control sent back to its SMEs |
| E10 Deck Department Training | Deck and Hazmat Qualification | Rejected by the XO: every topic sent back |
| E11 USS Northpoint | Basic Qualification Board | Rejected by the SM: navigation sent back |
| E12 Engineering Training Unit | Damage Control Refresher | SME Review: engineering signed off, damage control not |
| E13 Engineering Duty School | Comprehensive Board | SME Review: no topic signed off yet |

**Signed in as you (admin, DH, XO, SM):**

- **Dashboard** and **Approvals**: the 12 questions waiting at your three steps, and E5 (DH), E6 and E8 (XO) and E7 (SM). Approve E7 at the SM step to watch the app make its Word document live.
- **Question bank**: use the status chips to find each group in the table above. Open **History** on "Water from a fire hose..." to see version 1 retired by version 2, and on the safety data sheet question to see it sent back, changed and sent again.
- **Exams**: the 13 exams above. The Topics short tile may still show topics as short, because the overlap rule (10 exams back, 30% repeats) wants a bigger bank than the demo has.
- **Statistics**: on 6 months, the questions each topic gave on the four approved exams.
- **Admin: People and routing**: the eight made-up people with their roles; Riley Ford shown as turned off.
- **Admin: Topics**: six topics with their SMEs, and "Legacy Seamanship (turned off)".
- **Admin: Templates**: Basic Qualification Board (9 questions: DC 3, NAV 2, ENG 2, COM 2), Damage Control Refresher (4: DC 3, ENG 1), Deck and Hazmat Qualification (8: SEA 3, HAZ 3, DC 2) and Comprehensive Board (12: two from each topic).
- **Admin: Logs**: the audit trail, with dates spread over the last 80 days. The error log should be empty.

**Signed in as a made-up person:**

- `achen` **Dashboard**: the Class D fire question sent back by the XO and the anchor question sent back by the DH, with their reasons. **My reviews**: open reviews on E9 (damage control), E10 (every topic Alex Chen reviews), E12 (damage control) and E13.
- `jpatel` **Question bank**: the Mercator draft, the hazardous materials draft and version 2 of "How many minutes of latitude..." in draft (version 1 stays in the bank until version 2 is approved). **My reviews**: E11's navigation topic, sent back by the SM.
- `srivera`: a draft, and open reviews on E9, E10, E12 and E13. `ckim`: a draft, and open reviews on E10 and E13.
- `dreyes` **Approvals**: the six questions and E5 at the DH step. `mhale`: three questions, E6 and E8. `tbrooks`: three questions and E7. You see all of these because you hold all three steps.
- `dreyes` **Exams**, **New exam**: topics show as repeats, since the demo exams used most of the bank lately. Generate asks before it goes over the rule.

## The Word documents

They go to `Exam documents`, next to the data file. Open E1's to E4's and check the cover page, the topics, and the questions and their parts with their portion marks. The documents carry no answer key: reference answers stay in the app.

If a document fails, the approval still stands and the report prints a `FAILED:` line with Word's message. Open Word once by hand and close it before the seed, so a first-run prompt doesn't stall it.

## Starting over

The seed only fills an empty data file. To run it again, delete `build\ExamGen_BE.accdb` and run the five commands above. The made-up people stay in the Personnel file (they're reused, with their passwords as they are); delete them from Personnel if you don't want them.
