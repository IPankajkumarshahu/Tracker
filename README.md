# Gmail Starred Mails → Excel Tracker

Exports every **starred** Gmail message to Excel with these columns:

| Date | Time | Subject | Link | Reg No |
|------|------|---------|------|--------|
| 24-09-2026 | 13:30:00 | Claim for MH 12 AB 1234 | https://mail.google.com/mail/?authuser=you@example.com#all/... | MH12AB1234 |

**Reg No** is the vehicle registration number found in the subject line, written without spaces
(`MH 12 AB 1234`, `mh-12-ab-1234` → `MH12AB1234`). If there are several, they are separated by commas.
Supported formats: the standard state format (`MH12AB1234`, `DL3CAB1234`, `KA05MN123`) and the Bharat series (`22BH1234AA`).
If the subject labels the number (`Regn. No. HR890648`), it is picked up even without series letters.
Only real RTO state codes are accepted, so ordinary words are not picked up by mistake.

**No reg number? → claim number.** If the subject has no reg number, the claim number goes in the Reg No
column instead, and the cell is **highlighted light yellow** so you can tell it apart
(e.g. `<CL26217676>`, `Claim no. 10110425750`, `C1274101122507`). The cell stays blank only if neither is found.

Links use `?authuser=<your email>`, so they open the right account even when several are signed in.

There are two ways to run it. Pick one.

---

## Option A — One-click Google Sheet (recommended, no install)

**Set up once (about 2 minutes):**
1. Go to <https://sheets.new> to create a blank Google Sheet. Name it, for example, *Starred Mails Tracker*.
2. In the sheet, open **Extensions → Apps Script**.
3. Delete the sample code, paste in everything from
   [`apps_script/StarredMailsExport.gs`](apps_script/StarredMailsExport.gs), and click **Save** (💾).
4. Go back to the sheet and **reload the page**. A new **Starred Mails** menu appears next to *Help*.
5. The first time you use the menu, Google asks for permission: **Continue → choose your account → Allow**.
   (The script reads your Gmail and writes to this sheet only.)

**Every time after that — one click:**
- **Starred Mails → Export to Excel**: refreshes the sheet with all currently starred mails
  and downloads it as an `.xlsx` file. (If your browser blocks the automatic download, click the
  *Download Excel file* link in the pop-up.)
- **Starred Mails → Refresh sheet only**: updates the sheet without downloading.

Date and time use the spreadsheet's time zone (File → Settings → Time zone).

**Already created the script at script.google.com instead?** That works too:
select **exportToExcel** in the dropdown next to *Debug* and click **Run**. The **Execution log**
shows a *Download Excel* link. The first run creates a spreadsheet named *Starred Mails Tracker* in your
Drive, and later runs update that same spreadsheet. (The one-click menu only appears when the
script is added from inside a Google Sheet.)

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

### Tests
```bash
python -m unittest test_regno
```
