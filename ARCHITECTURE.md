# EG-MMS: How it is built

Meeting Management System for the Educate Girls Government Relations team, Uttar Pradesh.

This document is the build record: what the system is made of, how each module works, why each choice was made, and what was deliberately not done. It is written so that someone who has never seen the code, including a future version of the people who wrote it, can pick it up.

Written 19 September 2026, against deployment @126.

---

## 1. What the system does

Field officers meet government officials: BSAs, ABSAs, District Magistrates, DIET principals, block education officers. Before this system those meetings lived in WhatsApp messages and personal notebooks, and the monthly report was assembled by hand.

The system now covers the whole loop:

| Stage | What happens |
|---|---|
| Plan | Officer records who they intend to meet, when, and what they want from it |
| Conduct | After the meeting: what was discussed, what the official said, what comes next, the outcome, photos |
| Record | Minutes document written automatically, photos filed, colleague notified |
| Report | District, zone and state analytics; a monthly report emailed on the 1st |
| Follow | AI tags what needs attention, escalations go to seniors, prep briefs before the next meeting |

**Scale as of writing:** about 400 meetings, 75 officers filing them, 297 people in the employee master, 24 districts across three zones.

---

## 2. At a glance

```
Browser (phone mostly)
   |
   |-- dataimpact.in  (GitHub Pages, static)
   |      index.html        sign-in
   |      dashboard.html    the officer app
   |      report.html + 3   the open analytics portal
   |      data/*.json       a snapshot of public data, rebuilt periodically
   |
   +-- Apps Script web app  (one URL, doGet and doPost)
          |
          +-- Google Sheets   the database
          +-- Google Drive    photos, documents, government PDFs
          +-- Google Docs     minutes of meeting
          +-- Gmail           codes, notifications, escalations, monthly report
          +-- Google Calendar meeting events
          +-- Gemini, Mistral tagging, prep briefs, report narrative
```

| | |
|---|---|
| Backend | Google Apps Script, one file, `Code.gs`, about 5,900 lines |
| Frontend | Three hand written HTML pages, no framework, no build step for the app |
| Database | One Google Spreadsheet, eight tabs |
| Hosting | GitHub Pages on a custom domain, `dataimpact.in` |
| Deployment | `clasp` for the script, `git push` for the site |
| Runs as | `gr@educategirls.ngo`, web app access `ANYONE_ANONYMOUS` |
| Cost | Zero. Everything is inside the existing Workspace account. |

---

## 3. Why this stack

This is the part most worth reading, because the constraints drove everything else.

**The organisation already had Google Workspace and nothing else.** No server, no database, no budget line for either, and no one to administer them. Apps Script and Sheets were not chosen over Postgres and a Node server; they were the only things available.

That decision has three consequences that shape the whole system:

1. **There is no server you control.** Apps Script sits behind a front door that takes three to seventeen seconds before your code runs, and sometimes answers with an error page instead. Nothing in your code can change that. The only lever is crossing it less often.
2. **Every execution has a six minute ceiling and the account has six hours a day.** A dependency that hangs does not just fail; it eats the allowance for everyone.
3. **The database is a spreadsheet.** People can open it, edit it, and break it. It also means anyone in the organisation can look at the raw data without being given a login, which is genuinely useful.

**GitHub Pages for the frontend** rather than serving HTML from Apps Script. Apps Script can serve pages, but every page load would then cross that slow front door. Static files from a CDN arrive in half a second. This is the single biggest speed decision in the project.

**No framework.** Three HTML files with plain JavaScript. A build step would add a barrier between a change and seeing it work, and the whole app is about 6,000 lines of markup and script. React would have cost more than it returned here.

---

## 4. Data model

One spreadsheet, `1a7068K07gE40PLkxIs39A6OJvalCK7IgDJTZB5NQH40`, eight tabs.

A meeting starts life as a row in **Plan Meetings** and stays there for its whole life; its `Status` column is what changes. When it is conducted, a second row is written to **Conducted Meetings** holding everything the officer wrote. Postponed and Cancelled work the same way. So Plan Meetings is the register of what exists, and the other three are the record of what happened.

### Plan Meetings

| Col | Field | Notes |
|---|---|---|
| A | Meeting ID | `MTG-YYYYMMDD-HHMMSS`, generated at save, never reused |
| B | District | Where the meeting is, not where the officer is posted |
| C | Employee Name | |
| D | Post | |
| E | Email | The identity key everywhere |
| F | Meeting Date | |
| G | Meeting Time | Forced to text so Sheets does not reparse it |
| H | Duration | |
| I | Meeting Type | One-on-One, Joint Visit, Group Meeting, Dept. Review, Field Visit |
| J | Stakeholder Name | |
| K | Stakeholder Post | The one that matters, see Relationships below |
| L | Meeting Purpose | |
| M | Meeting Agenda | |
| N | **Status** | Planned, Follow-up, Conducted, Postponed, Cancelled |
| O-Q | Start, End, Reason | Filled on update |
| R-S | Colleague Name, Post | |
| T | Submitted At | |
| U | Parent Meeting ID | Set when this meeting was created as a follow-up |
| V | Documents folder URL | |
| W | Calendar Event ID | `COL_CAL_EVENT` |
| X | Stakeholder Block | `COL_PLAN_SKBLOCK` |

### Conducted Meetings

A-M mirror the plan row, then:

| Col | Field |
|---|---|
| N | Conduct Date |
| O | Conduct Time |
| P | **Key Points** — the three answers, joined |
| Q | Photos Folder |
| R | MoM Doc |
| S-T | Colleague Name, Post |
| U | Conducted At — the system timestamp, distinct from Conduct Date |
| V | Govt MoM — official PDFs, comma separated |
| W-AC | AI tags: priority, flag, next action, escalate, category, tagged at |
| AD | Escalation Sent At (`COL_ESC_SENT`) |
| AE | Stakeholder Block (`COL_CON_SKBLOCK`) |
| AF | Outcome (`COL_CON_OUTCOME`) |

Column U against column N is worth noting: it makes back-dating visible. A meeting held on the 19th and filed on the 25th shows both.

### Employee_DB

District, Block, Employee Name, Designation, Email, Role, Zone, **Additional Districts**.

`Role` is one of State, Zone, District, Field and decides what a person can see. `Additional Districts` (column H) is how one person holds charge of more than one district; it is comma or semicolon separated and the names must match exactly, including the space in `LAKHIMPUR KHERI`.

### Postponed, Cancelled, Officials, Stakeholder Type, Meeting Purpose

Postponed keeps Original Date, New Date, Reason, Postponed At. Cancelled keeps Reason and Cancelled At. The last three are small reference lists.

---

## 5. Modules

### 5.1 Sign-in

**Where:** `docs/index.html`, and `sendOTP`, `verifyOTP`, `loginPassword`, `setPassword` in `Code.gs`.

The first version was a code by email, every time. That meant two round trips through a slow front door plus however long the mail took, and each new code invalidated the one before it, so pressing the button again while waiting left you holding a dead code. On one bad morning it took nine codes to get in.

**How it works now:**

1. Email and password, one request.
2. A first-time user, or one who has forgotten, clicks a link to get a code. After the code is accepted the page requires them to set a password before continuing.
3. Forgotten password takes the same path.

**How the password is stored:** it is not. Each person gets a random salt, and Script Properties holds `salt$SHA-256(salt + password)` under `PW_<email>`. Five attempts per email per fifteen minutes. A wrong password, an unknown email and an account with no password set all return **the same message**, so the endpoint cannot be used to discover who is registered. The password travels in the request body, never the query string, because query strings are written into server logs.

**The session** is a UUID in CacheService for one hour, renewed on every authenticated call. It is held in the browser's `localStorage` for up to twelve hours, so reopening the app does not cost a sign-in. The server remains the authority: a token it no longer knows sends the person back to sign in. Logging out clears it.

**The part that matters most:** sign-in does not read the spreadsheet. The employee master is mirrored into Script Properties and `getEmployeeByEmail` reads cache, then that mirror, and only then the sheet. This exists because the sheet died one night and nobody could get in to see that everything else was fine.

### 5.2 Plan a meeting

**Where:** `docs/dashboard.html`, `saveMeeting` in `Code.gs`.

Captures district, date, approximate time, duration, meeting type, the official by **post** rather than name, their block if they are block-level, purpose, agenda, an optional colleague, and optional documents.

Decisions that took a while to get right:

- **The district list is in the app, not fetched.** The same zone mapping as `ZONE_DISTRICTS` in `Code.gs` is duplicated in `dashboard.html` so the dropdown fills on first paint. The server is still asked in the background and its answer wins, so a district added on the server but not mirrored still appears, one round trip later. When a district is added it must be changed in both places.
- **Time is mandatory**, because the calendar event was otherwise created at the wrong hour. Labelled "(Approx)".
- **Validation names the missing field** and outlines it red. "Please fill all required fields" sent people hunting, usually because an unset AM/PM leaves the time blank while the hour and minute beside it look filled.
- **The same plan twice is refused.** Same officer, official, date and purpose within ten minutes returns the existing meeting id instead of filing a second. This exists because one slow save became three meetings 105 and 43 seconds apart.

### 5.3 Conduct a meeting

**Where:** the modal in `dashboard.html`, `conductMeeting` in `Code.gs`.

Three questions rather than one free text box:

1. What was discussed?
2. What did the official say or agree to?
3. What needs to happen next? (optional)

Then a fourth, a dropdown: what came out of this meeting — Commitment, Information, Permission, Courtesy, Awaiting, Nothing concrete. The three answers are joined into one `Discussed: / Official said: / Next step:` block and stored as Key Points.

The three-question form replaced a single box because a `NOTE_quality()` audit found the average note was 147 characters and 53% were under 80. Splitting the question changed what people wrote more than any instruction did.

Also here:

- **Conduct date** defaults to today, is editable backwards for a meeting filed late, and cannot be set in the future.
- **Photos** are resized to 1600px and re-encoded as JPEG in the browser before upload. They were previously sent exactly as the camera took them: three or four megabytes each, a third larger again as base64, up to five of them. Anything that cannot be decoded is sent untouched, because losing a photo is worse than sending a large one.
- **The MoM document is optional and off by default.** Writing it is four to six seconds of Google Docs calls and most meetings do not need one.
- **One conducted record per meeting.** A second attempt is refused before the photos are uploaded, so a repeat leaves nothing behind in Drive either.
- **A note repeated word for word** from one of that officer's own earlier meetings is refused.
- **A draft** of everything typed is kept in `localStorage` and restored if the save fails. Photos are not kept: five files at 3MB would not fit.

### 5.4 Minutes, photos and documents

`createMoMDoc` writes a formatted Google Doc with the meeting details, the notes, follow-up information and links. Photos go to Drive under `EG-GR-Meetings / <district> / <meeting id> / `. Both are shared with anyone holding the link, which is how the portal can link to them.

**This is worth being explicit about:** because the portal is open and carries those links, the minutes of any meeting are reachable by anyone who opens the portal. The notes themselves are not in the published data, but they are inside the document that is linked. This was raised with the owner and left as it is deliberately.

### 5.5 Viewing meetings

Four screens, all in `dashboard.html`, differing only in scope:

| Screen | Scope | Server function |
|---|---|---|
| Manage Meetings | Your own planned meetings, overdue grouped first | `getMyMeetings` |
| My Meetings | Your own, by month, with a "My Month" summary card | `getAllMyMeetings` |
| District Meetings | Everything in your district | `getDistrictAllMeetings` |
| Zone / State Meetings | Your zone, or everything | `getZoneAllMeetings`, `getStateAllMeetings` |

Each keeps its last list in `sessionStorage` and draws it immediately, then corrects it from the server. That is per tab deliberately: the lists die with the tab, which matters when a phone is passed around.

District filters accept several districts at once, as tap-on tap-off chips behind a summary button. A native `<select multiple>` is close to unusable on a phone, which is where most of this is read.

### 5.6 The analytics portal

**Where:** `docs/report.html` is the master. `build-pages.js` generates three more from it by swapping one marker line, `var PORTAL_VIEW = '...'; /* BUILD:VIEW */`.

| Page | Shows |
|---|---|
| report.html | Overview: totals, outcomes, month trend, types, top districts, recent meetings |
| districtreports.html | One or several districts, by block and month |
| teamperformance.html | Who did how many, and coverage gaps |
| stakeholders.html | Relationships by office, transfers, coverage |

**After any edit to `report.html`, run `node build-pages.js` and commit all four.** Never hand edit the generated three.

**No login.** These pages read `docs/data/*.json`, a snapshot committed by a GitHub Action, and fall back to a live Apps Script call if the file is missing. The snapshot serves the same 242KB in half a second where Apps Script takes four to six. A chip shows the snapshot time and forces a live fetch if pressed.

The Action is scheduled every thirty minutes but GitHub does not honour that on free runners; gaps of two to five hours are normal. The timestamp on the page is what keeps this honest.

### 5.7 Relationships and stakeholders

Keyed on the **post, not the person**, because officials transfer and the office carries the relationship: the BSA of a district stays the BSA. Block-level posts (ABSA, ARP, and others) are one office per block, so the block joins the key.

The page surfaces offices met once, offices met three or more times, offices gone quiet, apparent transfers, and coverage gaps — posts that other districts meet and this one never has.

**Transfer detection** compares names within one office over time, and had to learn that the same person is written many ways. `K.K Singh` and `Shri Krishna Kant Singh` are one man. A transliteration normaliser handles `Pandey` against `Panday` and `Akhhilanand` against `Akkhilanand`. An edit-distance rule was written and then deliberately removed: no threshold separates `Dileep` from `Dilip`, who are the same person, from `Amit` and `Ajit`, who are not. `Annupurna` against `Annapurna` is still read as a transfer, and that is an accepted imperfection, because the rule that would merge them would also merge `Nandana` with `Nandan` and hide a real one.

### 5.8 The AI parts

Three places, and it is worth being precise because people assume more.

| Where | When | What it does |
|---|---|---|
| `tagOneMeeting_` | hourly | Reads the note, sets priority, flag, next action, escalate, category |
| `getMeetingPrep` | on demand | Reads every past meeting with that post and writes a brief |
| `aiReportNarrative` | monthly | Writes the prose in the monthly report |

**`callLLM` tries Gemini first, then Mistral**, two rounds with a 2.5 second wait. Gemini first because Mistral's free tier runs out. Keys live in Script Properties only, never in the code, because the repository is public.

`gemini-3.6-flash` is a thinking model and charges its reasoning to `maxOutputTokens`, so thinking is disabled and the budget raised; without that it returned 200 with empty content and `finishReason: MAX_TOKENS`, which looked like a working fallback that had never once worked.

**The prep brief answers in the officer's own language.** If their notes are Roman-script Hinglish, so is the brief. Devanagari gets Devanagari.

**Escalation is not the AI's decision.** The AI only marks the meeting. A three line rule decides the email:

```javascript
if (flag === 'Resolved') continue;
var esc = prio === 'High' || flag === 'Blocked' || escY === 'Yes';
```

That first line was missing once and a senior received an escalation about a meeting that had already been resolved. The fix was a rule, not a model.

**Prompt lessons.** The tagging prompt carries a glossary, because "block" was read as obstruction when it means an administrative block, and a COMPLETED-vs-PENDING section, because "report submit ki" was read as a pending request when it means the work is done. Past-tense Hinglish is listed explicitly.

**Not switched on:** commitment extraction. `extractCommitments_` exists and runs only from `COMMIT_dryRun()`. An early version over-tightened into finding zero, then rebalanced; it waits for notes written under the new form to accumulate.

### 5.9 Scheduled jobs

| Job | When | Does |
|---|---|---|
| `taggingJob` | hourly | Tags up to 12 untagged meetings; exits in a second when there is nothing |
| `escalationJob` | hourly | Sends escalations for meetings the rule selects |
| `calendarJob` | hourly | Creates calendar events, and refreshes the employee and purpose mirrors |
| `monthlyReportJob` | 1st, 7am | Emails the monthly report, scoped per recipient |
| `nudgeJob` | Mondays | Built, not installed. Revisit around mid-October. |

`TRIGGER_status()` lists what is actually installed. `TRIGGERS_pauseHourly()` and `TRIGGERS_restoreHourly()` take the hourly three off and put them back exactly as they were.

### 5.10 The monthly report

Emailed on the 1st, scoped to the recipient: state level gets the whole state, a district lead their district, a zone lead their zone. It carries performance by zone, a district leaderboard, participation, outcomes, relationships, and an AI-written narrative.

A district with no zone is labelled **State** rather than "Unzoned"; Lucknow is a state-level district with no team of its own, selectable by everyone through `STATE_EXTRA_DISTRICTS`.

---

## 6. Caching

Three layers, and getting the middle one wrong was the most expensive mistake in the project.

**1. CacheService, per sheet, shared.** `sheetRows_(name)` holds an entire tab, split across numbered keys because a cache value is capped near 100KB, and in small character chunks because a Devanagari note is three bytes a character while the cap is on bytes. Dates survive: `getValues` returns real `Date` objects, JSON turns them into ISO text, and a reviver turns them back, matched strictly enough that a note or a name cannot be caught by it.

**2. CacheService, per answer.** `reportData`, `stateMtg_all`, `distMtg_<district>`, `planmtg_<email>`, and so on. Ten minutes, safe because every write invalidates.

**3. The browser.** `sessionStorage` for lists, `localStorage` for the session.

### The mistake worth remembering

Caching was originally only layer 2: each person's finished answer, keyed by their email. Fifty people opening My Meetings meant the same three sheets read fifty times inside the cache window, because the key was the person and not the data. It is invisible when you test alone.

```
Per-user answers : ~200 full sheet reads per ten minutes
Per-sheet shared :     4
```

**Every write must clear the shared copy.** A related bug survived a long time: `invalidateUser` cleared the writer's own keys and the state list but not `distMtg_<district>` or `zoneMtg_<zone>`, which are shared, so a meeting just filed stayed invisible to that district's lead.

---

## 7. Resilience

Everything here exists because something failed badly once.

**Circuit breaker.** Two sheet failures inside two minutes leaves the document alone for five. Requests then return in milliseconds instead of six minutes. Tripping on a single failure was the first design and was wrong: services throw one-offs that clear themselves, and stopping everyone over one is a cure worse than the illness. `BREAKER_status()`, `BREAKER_reset()`.

**A request budget.** A web request that has already spent ninety seconds will not start another sheet read. `getReportData` opens five tabs; there is no sense discovering the same failure five times. Timed jobs are exempt, because tagging sleeps between LLM calls.

**Reads throw, they do not return empty.** A failed read must never look like "there are no meetings". That distinction is the difference between an error and apparent data loss.

**Retry only what could not have happened.** An error page means the script never ran, so sending it again is free. Silence means it may have run, and a second send could act twice. Reads retry both; writes retry only the error page.

**Dropped POST bodies.** The front door delivers a POST without its body often enough to have been the commonest complaint about saving: the script saw no date, no name, no purpose, and refused. Roughly one browser probe in seven arrived empty. The server now answers `BODY_MISSING` and the browser quietly sends the same save again. This is safe precisely because the server refuses to write a half meeting.

**Idempotency.** The same save within ten minutes returns the first meeting's id. A verified code offered again within five minutes returns the same session. A conducted meeting cannot be conducted twice.

**`fetch` has no timeout.** Requests are abandoned after 45 seconds. Without this the sign-in button sat on "Please wait" indefinitely with nothing to read and nothing to press.

**Maintenance switch.** `MAINT_on()` turns sheet-backed actions away in milliseconds so the execution queue can drain, while sign-in stays allowed.

---

## 8. Security

- **Identity comes from the session token, never from client-supplied parameters.** `apiResponse` looks up the session and uses `session.email`; a browser claiming to be someone else gets nowhere.
- **Only `@educategirls.ngo` addresses** can request a code, and only people in the employee master.
- **Public actions are enumerated**, not assumed: `sendOTP`, `verifyOTP`, `loginPassword`, and the four that feed the open portal. Everything else needs a token.
- **Admin actions** are restricted to an explicit allowlist of two addresses.
- **Passwords are salted and hashed**, rate limited, and travel in the body.
- **API keys live in Script Properties**, never in the repository, which is public.
- **Access is re-verified on every page load**: a person removed from the employee master is logged out.
- **A shared common password was proposed during an outage and declined**, because the app is anonymous-access on a public URL and the code is the only thing keeping strangers out.

**Known and accepted:** the portal is open by design, and the minutes and photos it links to open for anyone holding the link.

---

## 9. Deployment

**Frontend**

```bash
git push educategirls main     # GitHub Pages rebuilds dataimpact.in
```

Remote `educategirls` is production. `origin` is a stale personal fork; do not use it. Pages caches HTML for ten minutes, so a browser can hold the previous version that long; `Ctrl+Shift+R` after a deploy.

**Backend**

```bash
npx clasp push -f
npx clasp redeploy AKfycbw2JJ5xmZ-zLUolbZJb7eApczsZjsXwzVY6uXpAYO-7h8j9CyNF9y5Upgxji8rD2oJb -d "what changed"
```

The deployment id must stay the same or the frontend's `GAS_URL` breaks. Editor functions run at HEAD, so `clasp push` alone is enough to try a helper; the web app only changes on redeploy.

**The version number**

Every page shows `Version 1.nn` in its footer. It counts **releases of these
pages**, not of the Apps Script: the pages are what browsers cache, so they
are what the number has to track. It started at `1.26` to match Apps Script
deployment 126 and has moved on its own since. The dot is cosmetic, so it
reads like a version rather than a serial number. It is written into the HTML rather than fetched, which is the whole
point: a page served from a stale cache shows the stale number, and that is the
only way anyone can tell that is what happened.

**Bump it in all three files on every deploy**, `docs/index.html`,
`docs/dashboard.html` and `docs/report.html`, then run `node build-pages.js`. A
version number that stops moving is worse than none, because people start
trusting it.

**Traps that have cost time**

- **Never run `clasp pull`.** `.clasp.json` has an absolute `rootDir`, so it writes `Code.js` and `Setup.js` next to the real `.gs` files and those get pushed as duplicate definitions.
- **`clasp push` uploads every `.js` and `.html` in rootDir** unless `.claspignore` excludes it. A plain Node script at the root took the entire web app down once with `require is not defined`. `.claspignore` currently holds `docs/**`, `.git/**`, `*.md`, `.claspignore`, `build-pages.js`, `*.sh`, `scripts/**`, `.github/**`.
- **Editing a form field: change `docs/dashboard.html`.** The root `MeetingForm.html` and `Index.html` are the legacy Apps Script frontend and are not what anyone sees.
- **`appsscript.json` must keep `ANYONE_ANONYMOUS`.** A local copy once had `ANYONE` and pushing it would have changed who can reach the app.
- **clasp's login reverts.** It stores one Google account; any `clasp login` on this machine overwrites it with the default browser account, which is not an editor on the gr-owned script.

---

## 10. Operator's reference

Run these from the Apps Script editor.

| Function | Use |
|---|---|
| `TRIGGER_status()` | Which jobs are installed |
| `BREAKER_status()` / `BREAKER_reset()` | Is the sheet being left alone; let requests back in |
| `ROWS_status()` | Is the shared sheet copy filled |
| `SHEET_sizes()` | Rows and columns per tab, without reading a cell |
| `SHEET_trimTail("DELETE")` | Remove empty rows below the data |
| `SHEET_probe()` | Does a new sheet open, does ours |
| `EMP_refreshMirror()` | Refresh the sign-in copy after changing the employee sheet |
| `EMP_mirrorStatus()` | How many people are in the copy, and when it was taken |
| `PURPOSE_refresh()` | Refresh the purpose list |
| `PW_status()` / `PW_clear("email")` | Who has set a password; clear one |
| `LOGIN_debug("email")` | Mail quota left, employee lookup, triggers, waiting code |
| `NOTE_quality()` | Note length distribution and repeats |
| `CONDUCT_findDupes()` / `PLAN_findDupes()` | Rows that look like one meeting saved twice |
| `MTG_trace("MTG-...")` | Every row any tab holds for one meeting |
| `ESC_preview()` | What escalation would send, sending nothing |
| `LLM_probe()` | Each AI provider's real HTTP code and body |
| `MAINT_on()` / `MAINT_off()` | Turn sheet-backed actions away while something is wrong |

**Row numbers are per tab.** Row 127 of Conducted Meetings is a different meeting from row 127 of Plan Meetings, and any bulk delete works bottom-up.

---

## 11. Decisions taken and not taken

| Considered | Outcome |
|---|---|
| Merge endpoints into one bootstrap call | **No.** Measured: four parallel requests cost the same as one. It would have bought nothing. |
| Split the data into several spreadsheets | **No.** `openById` is the expensive part, and `getReportData` reads five tabs; splitting would make it five document opens. Revisit when meetings reach a few thousand, then archive by year. |
| Move the MoM document to a background job | **No, for now.** It would need a new trigger, a changed success screen, a delayed colleague email, and it fails silently if the trigger stops. Making the document optional removed most of the cost with none of that. |
| A shared password during an outage | **No.** The app is reachable by anyone; the code is what keeps strangers out. |
| Seed the sign-in copy from a four month old spreadsheet during an outage | **No.** Someone who had left would have regained access. |
| One meeting spanning several districts | **No.** No evidence in the data; two officers work across districts and `Additional Districts` already covers them. It would ripple through every report. |
| A note quality gate | **Built, then removed.** It refused genuine notes, and a conduct that will not save never produces its document either. |
| Edit-distance name matching for transfers | **Written, then removed.** No threshold separates `Dileep`/`Dilip` from `Amit`/`Ajit`. |
| Weekly nudges | **Built, not installed.** Revisit around mid-October. |
| Commitment tracking | **Built, not installed.** Waiting for notes under the new form. |

---

## 12. Limits, and what comes next

**What this architecture cannot do.** Apps Script's front door costs three to seventeen seconds a request and a POST costs three to four times a GET. No amount of work inside the script changes that. The system is built to cross it rarely rather than to cross it quickly, and that is close to exhausted as a strategy.

**Where it breaks.** Somewhere in the low thousands of meetings, full-tab reads stop being cheap and the data needs archiving by year. If the team grows several times over, the daily runtime allowance becomes the ceiling.

**The twelve state plan.** One app, one spreadsheet per state, routing on the email domain, with Firestore or BigQuery only if and when it is genuinely needed. That is recorded separately; the important point is that the per-state split is a data decision, not a code rewrite, and the current shape supports it.

**The near list:** weekly nudges, commitment tracking, a second `NOTE_quality()` reading against the 2026-09-16 baseline of 147 characters average and 53% under 80.

---

## 13. How the work actually went

Worth recording, because the method mattered more than any single decision.

Nothing in this system was designed up front. Every part of it came from somebody using it and saying what was wrong, usually in the middle of a working day, usually about something the people building it had not thought about. The duplicate meetings, the locked date field, the note check that saved the bad version and rejected the corrected one, the emptied dropdowns, the default nobody would change: all reported from ordinary use, none caught in testing.

The second habit that paid was measuring before believing. Several confident diagnoses turned out wrong — that Google was down, that the document was corrupt, that the mail quota had run out, that the script had too many functions. Each was disproved by a measurement that took two minutes, and each would have cost hours if it had been acted on instead.

The third was capturing output before a risky change and comparing after. Rewriting how every sheet is read was shipped the same hour it was written, because 401 records could be compared either side and not one date had moved.
