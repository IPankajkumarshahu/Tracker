# Gmail Starred Mails → Excel Tracker

Exports every **starred** Gmail message to Excel with these columns:

| Date | Time | Subject | Link | Reg No |
|------|------|---------|------|--------|
| 24-09-2026 | 13:30:00 | Claim for MH 12 AB 1234 | https://mail.google.com/mail/u/0/#all/... | MH12AB1234 |

**Reg No** is the vehicle registration number found in the subject line, written without spaces
(`MH 12 AB 1234`, `mh-12-ab-1234` → `MH12AB1234`). It is blank if the subject has none. If there are several, they are separated by commas.
Supported formats: the standard state format (`MH12AB1234`, `DL3CAB1234`, `KA05MN123`) and the Bharat series (`22BH1234AA`).
Only real RTO state codes are accepted, so ordinary words are not picked up by mistake.

There are two ways to run it. Pick one.

---

## Option A — Google Apps Script (no install, easiest)

1. Open <https://script.google.com> → **New project**.
2. Delete the sample code and paste in everything from [`apps_script/StarredMailsExport.gs`](apps_script/StarredMailsExport.gs).
3. Select `exportStarredMails` in the function dropdown → **Run** → approve Gmail/Sheets access.
4. Open **Execution log**. It shows a link to the new Google Sheet and a direct **.xlsx download** link.
   (You can also use File → Download → Microsoft Excel.)

Date and time use the script's time zone (Project Settings → Time zone).

## Option B — Python script

### One-time setup
1. Go to Google Cloud Console → create a project → **enable the Gmail API**.
2. Go to **APIs & Services → Credentials → Create credentials → OAuth client ID → Desktop app**.
   Download the JSON file and save it as `credentials.json` in this folder.
   (If the consent screen is in "Testing" mode, add your Gmail address as a test user.)
3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Run
```bash
python gmail_starred_export.py                      # writes starred_mails.xlsx
python gmail_starred_export.py -o my_list.xlsx      # custom file name
python gmail_starred_export.py --tz Asia/Kolkata    # time zone for Date/Time (default)
```
The first run opens a browser for Google sign-in (read-only Gmail access). After that, the sign-in is saved in `token.json`.
If you use several Google accounts in one browser and the links open the wrong one, pass `--account-index N`
(the `N` in `mail.google.com/mail/u/N`).

### Tests
```bash
python -m unittest test_regno
```
