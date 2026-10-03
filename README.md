# MICA timetable → live calendar feed (S2, S4)

Every hour, GitHub Actions downloads the timetable Excel from SharePoint, picks out
S2 and S4 classes, and publishes `docs/calendar.ics`. Subscribe to it once and
reschedules, cancellations and new sessions flow into your calendar automatically.

## Setup (about 10 minutes)

1. **Create a repo** on GitHub (e.g. `mica-calendar`) and upload all these files,
   keeping the folder structure (`.github/workflows/update.yml`, `docs/`).
   Note: GitHub Pages on a free account needs a **public** repo, which means the
   class schedule will be publicly viewable. The SharePoint link itself stays hidden.

2. **Add the SharePoint link as a secret**: Settings → Secrets and variables → Actions →
   New repository secret. Name `SHEET_URL`, value = the full SharePoint link.

3. **Run it once**: Actions tab → "Update calendar" → Run workflow. Open the run log:
   it shows each sheet, the layout detected and how many classes it found.

4. **Turn on Pages**: Settings → Pages → Source "Deploy from a branch" →
   Branch `main`, folder `/docs` → Save.

5. **Your feed URL** is:
   `https://<your-username>.github.io/<repo-name>/calendar.ics`

## Subscribing

- **Google Calendar** (web): Other calendars → + → From URL → paste the link.
  Google refreshes subscribed calendars on its own schedule, often every 8–24 hours;
  this can't be forced.
- **Apple Calendar**: File → New Calendar Subscription (Mac) or Settings → Calendar →
  Accounts → Add Subscribed Calendar (iPhone). Set auto-refresh to every hour.
- **Outlook**: Add calendar → Subscribe from web.

Use **subscribe**, not import. Importing a file makes a one-time copy that never updates.

## Troubleshooting

**"could not download the Excel file"**: the link only works for signed-in MICA
accounts. Ask the sender for an "Anyone with the link can view" link, or see the
alternative below.

**0 classes found / wrong classes**: download a copy of the Excel and run locally:

    pip install -r requirements.txt
    python build_ics.py --file timetable.xlsx --inspect

That prints the first rows of each sheet. Then run without `--inspect` to see what gets
picked up. The parser expects a header row containing a "Date" column, and either a
"Section" column (list layout) or time slots as column headers (grid layout).

**Changing sections**: edit `SECTIONS` in `.github/workflows/update.yml`.

**Workflow stopped running**: GitHub pauses scheduled workflows after 60 days with no
repo activity. Timetable updates count as activity, but over a long break re-enable it
from the Actions tab.
