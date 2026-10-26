# OSA Pup Pause

A one-minute phone check-in for HGSE students, guided by Howie the Shiba pup and three mice.
Student prototype for HGSE T522 Impact by Design (Kimber Cooper). Not an official OSA service.

## What's in this folder

| Path | What it is |
|---|---|
| `index.html` | The whole app in one file: Howie, the mice, 12 languages, seasonal themes, OSA staff routing. |
| `apps-script/Code.gs` | Google Apps Script that receives anonymous check-ins and saves them to a private Google Sheet, one tab per month, plus a Summary tab and a Test report tab (Test 1 vs Test 2 with go-forward verdicts), refreshed hourly. |
| `qr/` | QR codes for the everyday link and the Test 1 / Test 2 links (generated after Pages is live). |

## How the data stays private

- **Public (GitHub):** only the app code and design. No student answers are ever stored here.
- **Private (Google Sheet):** anonymous check-ins, no names. The script can only add rows; it never sends data back.
- **Check on me emails** (optional, student-entered) go to a separate "Check on me" tab.
- Before real students use it, confirm with HGSE whether wellbeing data must stay in a Harvard-approved account (Harvard Google Workspace or Qualtrics).

## Setup (about 10 minutes)

1. **Results Sheet:** create a Google Sheet called "Pup Pause results". Extensions → Apps Script → paste `apps-script/Code.gs` → Save.
2. Run **setup** once (approve the permissions). This creates the Settings, Summary and Test report tabs and an hourly refresh. Reopen the Sheet to see the **Pup Pause → Refresh reports now** menu.
3. **Deploy → New deployment → Web app.** Execute as: *Me*. Who has access: *Anyone*. Copy the web-app URL.
4. In `index.html`, find `var ENDPOINT='';` and paste the URL between the quotes. Commit.
5. **GitHub Pages:** Settings → Pages → Deploy from branch → `main` / root. The app goes live at `https://supersher0.github.io/pup-pause/`.
6. In the Sheet's **Settings** tab, enter the number of HGSE students enrolled (for the "% of HGSE" newsletter line), and after each test the date and your headcount of students in the room.
7. Share the Sheet as **Viewer** with OSA staff: Rilda Kissel (rilda_kissel@gse.harvard.edu) and Kevin Boehm (kevin_boehm@gse.harvard.edu).

## Links

- Everyday link: `https://supersher0.github.io/pup-pause/`
- Test rounds: add `#test1` or `#test2` to the end, so each round's results stay separate.
- Preview another month (decorations and greetings): add `#preview=2026-12-24`.
