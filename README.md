# Work Hours

A phone-first web app for logging work hours that writes each day
straight into a real Excel workbook in OneDrive, so John and Ashley always see
current numbers without you sending anything.

- Log a day with Date, Time In, Time Out, or tap Clock In / Clock Out
- Hours are calculated in decimal (9:00 AM to 1:30 PM = 4.50)
- Running total for the current pay period, plus any month you pick
- Pay periods are the 1st through the 15th and the 16th through end of month
- Entries save on the device first, then upload, so a dead signal never loses a day
- Logging from more than one device works: every sync pulls the workbook down
  too, so a phone and a laptop both stay caught up with what the other logged

## The workbook

Already created here:

```
LV Intern Work - Documents / 2026 / For Ashley /
Jackson Darr - Work Hours 2026.xlsx
```

That library is already shared with the team, so John and Ashley can open it
without you sharing anything separately. Move or rename it if you prefer another
spot, then update the Workbook URL in the app's Settings.

It has two sheets:

**Summary** opens on the current pay period (start date, end date, hours, days
worked), then a month-by-month grid with the 1st to 15th subtotal, the 16th to
end subtotal, and the month total, finishing with a year total. Every number is
a live formula reading the log, so totals update the moment a row lands.

**Hours Log** is the daily log itself: Date, Day, Time In, Time Out, Hours,
Period, Notes. This is the table the app appends to.

Log your hours in the app, not by typing into the workbook. The app is the
source of truth and a hand-typed row will not have a matching entry on your
phone.

## One-time setup

### 1. Register the app in Azure

This is what lets the app write to your OneDrive as you. It takes about ten
minutes and only has to be done once.

1. Go to <https://entra.microsoft.com> and sign in with your Bakers Creek account
2. **Applications** > **App registrations** > **New registration**
3. Name it `Jackson-Hours` (the existing registration uses this name)
4. Supported account types: **Accounts in this organizational directory only**
5. Under **Redirect URI**, change the dropdown to **Single-page application (SPA)**
   and enter the URL where the app will live:
   `https://jd182-jpg.github.io/work-hours/`
6. Click **Register**
7. On the Overview page, copy the **Application (client) ID** and the
   **Directory (tenant) ID**. You need both in step 3.
8. Go to **API permissions** > **Add a permission** > **Microsoft Graph** >
   **Delegated permissions**, search for `Files.ReadWrite.All`, tick it, and
   click **Add permissions**

You do not need a client secret. A browser app cannot keep one, which is why
this uses PKCE instead.

The platform must be **Single-page application**, not **Web**. Registered as
Web, Microsoft blocks the browser's final token request and sign-in fails on the
last step with a CORS error.

### 2. Put the app online

Microsoft only accepts `https://` redirect URIs, with `http://localhost` as the
single exception. That rules out serving from your Mac over Wi-Fi at an address
like `http://192.168.1.50:8000`, because Azure will not accept that as a
redirect URI. So the phone needs a real HTTPS host.

GitHub Pages is free, gives you HTTPS, and works from anywhere including
cellular. From this folder:

```bash
cd ~/Documents/Projects/work-hours && git add -A && git commit -m "Work hours tracker"
```

```bash
gh repo create work-hours --public --source=. --push
```

```bash
gh api -X POST repos/jd182-jpg/work-hours/pages -f "source[branch]=main" -f "source[path]=/"
```

Give it a minute, then open <https://jd182-jpg.github.io/work-hours/>.

Nothing about your tenant or your workbook is in this code, which is why the
repo can be public. Those three values live only in your browser's storage on
each device you set up.

To push later changes:

```bash
cd ~/Documents/Projects/work-hours && git add -A && git commit -m "Update" && git push
```

### 3. Enter your settings on the phone

Fastest way: open the one-time setup link. It carries the three values in the
URL fragment, which browsers never send to a server, so nothing about the tenant
is published anywhere. The app reads it, saves it, and strips it from the address
bar. Do that on each device you want to log from.

Otherwise, open the site, scroll to **Settings**, and paste in:

- **Application (client) ID** from step 1
- **Directory (tenant) ID** from step 1
- **Workbook URL**: currently
  `https://onpointcustomhomes.sharepoint.com/sites/LVInternWork/Shared%20Documents/2026/For%20Ashley/Jackson%20Darr%20-%20Work%20Hours%202026.xlsx`
  If that ever stops resolving, open the workbook in a browser and copy the
  address bar instead, or use Share > Copy link. Either form works.

Tap **Save Settings**, then **Connect Excel** and sign in with your Bakers Creek
account. Approve the permission prompt once.

The Settings card also prints the exact **Redirect URI** the app is using. It has
to match what you registered in Azure character for character. If sign-in fails,
compare those two first.

Repeat this step on your laptop if you want to log there too. Settings are per
device, but every sync also pulls the workbook down and reconciles it against
what's on the device, so logging on your phone one day and your laptop the
next keeps both showing the same complete picture rather than each only
knowing about what it personally typed in.

### 4. Add it to your home screen

In Safari on the phone, tap Share, then **Add to Home Screen**. It then opens
full screen like an app.

## Running it on your Mac for testing

You do not need this for normal use, but it is handy for trying changes:

```bash
cd ~/Documents/Projects/work-hours && python3 -m http.server 8000
```

Then open <http://localhost:8000>. Sign-in works here too, as long as you also
add `http://localhost:8000/` as a second SPA redirect URI in the Azure app
registration. Stop the server with Ctrl-C.

## Daily use

Tap **Clock In** when you start. The timer keeps running if you close the app or
your phone restarts. Tap **Clock Out** when you finish and the times drop into
the form, where you can add a note and tap **Save Entry**.

Or skip the timer and fill in Date, Time In, Time Out yourself.

An amber dot next to an entry means it has not reached the workbook yet. It
retries on its own the next time the app opens, or you can tap the badge in the
top right to sync immediately.

Deleting an entry in the app also removes that row from the workbook.

## If something goes wrong

**Sign-in fails right after the Microsoft page.** The redirect URI in Azure does
not match. Compare it to the one printed in the app's Settings card. Also confirm
the platform is Single-page application, not Web.

**"Need admin approval" on the permission prompt.** This is what happened on the
first attempt. The leggettventures.com tenant has user consent for apps switched
off, so no amount of fiddling with the registration will fix it from your side.
An admin has to approve the app once, either by opening

```
https://login.microsoftonline.com/ebfd2362-5605-4f59-b603-aa68d309e7d1/adminconsent?client_id=7401adb8-8ee7-4b4f-bac6-5b49c145742a
```

and signing in with an admin account, or in the Entra admin center under
**App registrations > Jackson-Hours > API permissions > Grant admin consent**.

It is a one-time action for the whole tenant. Until it happens the app still logs
hours locally and uploads the backlog the first time you connect, so keep using
it and nothing is lost.

The permission is **delegated**, not application. That means the app acts as you
and can never touch a file you could not already open yourself, and it has no
background or app-only access. Worth saying explicitly if IT asks, because
"Files.ReadWrite.All" reads broader than it behaves.

**"Cannot open the workbook".** The Workbook URL in Settings is wrong, or the
file was moved or renamed. Paste the current URL and save again.

**Totals in Excel look stale.** Excel Online recalculates on open. Refresh the
browser tab.

**Hours are right in the app but missing in Excel.** Tap the badge in the top
right to force a sync and watch the message under Excel workbook.

## Rebuilding the workbook

If the file is ever lost, regenerate an empty one:

```bash
python3 ~/Documents/Projects/work-hours/tools/make_workbook.py --year 2026 --out "Jackson Darr - Work Hours 2026.xlsx"
```

Use **Download CSV backup** in the app first if you need to re-import past
entries. The app keeps its own copy of every entry on the device regardless.

For a new calendar year, run it again with `--year 2027`.

## Files

| File | What it does |
|---|---|
| `index.html` | The page |
| `styles.css` | Phone-first styling |
| `app.js` | Entries, pay-period maths, totals, clock, sync queue |
| `excel.js` | Microsoft sign-in (PKCE) and the Graph calls that write rows |
| `config.js` | Optional defaults; normally left blank and set in Settings |
| `sw.js` | Service worker, network-first so the app is never stale |
| `manifest.json`, `icons/` | Home-screen install |
| `tools/make_workbook.py` | Generates a fresh workbook with the summary formulas |

`sw.js` is deliberately network-first. A cached stale timesheet is worse than a
slow one. If you change any shell file, bump `CACHE` in `sw.js` and the matching
`?v=` numbers in `index.html` together.
