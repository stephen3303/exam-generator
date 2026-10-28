# Prototype quickstart (PowerSpec)

Build the prototype from `main` (all ten phases and Sign out), open it, and walk one question and one exam from start to the Word document. Setup takes about 10 minutes and the smoke path about 10 more.

Nothing has been run in Access, Word or Outlook yet: every check so far was a code review. If anything stops you, paste the PowerShell output or the message's exact text into the thread and it gets fixed on the branch.

## 1. Before you start (one time)

- **Close every Access window.** The build deletes and remakes `ExamGenerator.accdb` and can't while it's open.
- **Open Word once by hand and close it.** The app drives Word hidden with no timeout, so a first-run prompt (sign in, privacy, file types) would hang the SM's final approval behind its busy pop-up.
- **Use git, not a downloaded ZIP.** A ZIP's files carry the internet mark, and Windows then blocks their macros.

## 2. Get the branch and build

In PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
git clone -b main https://github.com/stephen3303/exam-generator.git $HOME\eg-proto
cd $HOME\eg-proto\access
```

Already have a clone? From its folder instead:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
git fetch origin
git switch main
git pull
cd access
```

Then build a fresh data file and the front end linked to it:

```powershell
Remove-Item .\build\ExamGen_BE.accdb -ErrorAction SilentlyContinue
.\backend\new-backend.ps1
if ($?) { .\build.ps1 -BackEnd .\build\ExamGen_BE.accdb }
```

Both scripts end with `Created:` and `Built:`. A compile error or an Access dialog stops the build and is printed instead; send that output.

## 3. Trust the build folder (one time)

Without this, Access opens the app with its code turned off. This adds `access\build` as a trusted location for Access 2016 and later, Microsoft 365 included:

```powershell
$key = 'HKCU:\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\ExamGenProto'
New-Item $key -Force | Out-Null
Set-ItemProperty $key -Name Path -Value ((Resolve-Path .\build).Path + '\')
Set-ItemProperty $key -Name AllowSubfolders -Value 1 -Type DWord
Set-ItemProperty $key -Name Description -Value 'Exam Generator prototype'
```

If a policy blocks trusted locations, open the app and press **Enable Content** on the yellow bar instead.

## 4. Open it and set up

```powershell
Start-Process .\build\ExamGenerator.accdb
```

1. **Personnel.** The first open finds no Personnel file and offers to make one. Press **Create**, keep the folder it shows (`access\Personnel`, which git ignores) and the key, type your name and an app password, and create it. You become the Personnel admin, linked to your Windows account.
2. **A second person.** Add person opens next. Add `Test SME` with the sign-in name `testsme`, no Windows account, a password of 8 or more characters, and **Can open Exam Generator** ticked. Save, then Close. The app opens with you as its admin.
3. **People and routing.**
   - **Add person**, pick Test SME, and turn on **SME**.
   - Open yourself and turn on **DH**, **XO** and **SM**.
4. **Topics.** **New topic** and name it `Smoke test`. Set questions per exam to **1** and **Save**. Under SMEs, pick Test SME and **Add**.
5. **Templates.** **New template** and name it `Smoke`. **Add topic**, pick Smoke test, set the slider to **1**, and **Add**.
6. **Switching people.** **My settings**, then **Sign out**, brings up the sign-in screen, where you sign in as `testsme` or as yourself.
   - Closing Access and opening it again signs you straight back in by your Windows account, so use Sign out to switch.

## 5. Smoke path

Switch people with **My settings**, then **Sign out**.

| As | Page | Do |
|---|---|---|
| testsme | Question bank | **New question**: topic Smoke test, Short Answer, Easy, 1 point, marking (U), any text. **Cite a reference**, add a new one with a paragraph. **Save**, then **Submit for approval**. |
| you | Approvals, Questions tab | **Approve** at the DH step, then the XO step, then the SM step. The question is approved. |
| you | Exams | **New exam**: template Smoke, type who it's for, **Generate**. |
| testsme | My reviews | Open the topic and **Sign off**. The exam goes to the DH. |
| you | Approvals, Exams tab | **Approve** three times (DH, XO, SM). The SM's approval makes the Word document. Expect a busy pop-up while Word works. |
| you | Exams | Pick the exam, then **Download a copy**. Open the .docx: a cover page, then the question with no answer, marked (U). |

The documents are kept in `access\build\Exam documents\`.

Emails are off for everyone by default, so this path never touches Outlook. To try email, turn on Emails in **My settings** and send the test email; Outlook may ask to allow the send.

## 6. If something goes wrong

| What you see | What to do |
|---|---|
| The build prints `Compile error: ...` or a dialog | Send the output. Nothing to change on your side. |
| `There is already a file at ...` | The data file from the last run is still there. Run the `Remove-Item` line again first. |
| The build can't delete or open `ExamGenerator.accdb` | An Access window is still open, or a stuck `MSACCESS.EXE` from a failed build. Close it in Task Manager. |
| The app opens with a yellow security bar, or nothing happens | Step 3 didn't take. Press **Enable Content**, or check the path in the trusted location. |
| "You don't have an Exam Generator role yet" | That person needs a role in People and routing. |
| The busy pop-up never ends at the SM's approval | Word is stuck on a prompt. End `WINWORD.EXE` in Task Manager, open Word by hand to clear the prompt, then use **Download a copy**, which makes the document if it's missing. |
| Anything else | Copy the message's exact text, or the newest rows on **Logs**, Error log tab, into the thread. |

## Also on the branch

- `review/REVIEW-TASK.md` sets up the three-VM review, which is optional and not needed for tonight.
- `review/functional-checks.md` is the longer list (upgrades, two copies, page markings across pages, Outlook, the locked `.accde`).
