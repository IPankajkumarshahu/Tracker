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
node --test 'tests/*.test.js'
```

---

# Uploading Sheet Filler (starred mails → "my work")

[`apps_script/UploadingSheetFiller.gs`](apps_script/UploadingSheetFiller.gs) fills the **my work** tab of the
uploading sheet from your **starred** Gmail mails. For each mail it reads the mail and the attached **RC** and
**vehicle photos** with Claude (Anthropic's AI model), then fills the other columns from the lookup tabs in
the same spreadsheet. It also runs QC checks and highlights anything a person needs to verify.

## What it fills

| Source | Columns |
|--------|---------|
| Mail | Contact Person Email / Name / No, ACR CC Email, ACR subject line, Claim No, Intimation Date, Yard Name, Yard Location |
| RC + photos (Claude) | Reg No, Engine No, Chassis No, Vehicle Type, Make Year, Make Month, Fuel Type, Ownership, Reg Date, Reg Type, Tax Validity / Permit / Fitness dates (commercial only), Transmission, Luxury, Flood / Superdari / Theft / Transit |
| **Seller id** tab | Seller ID (from the sender's email domain, via the **Seller Rules** tab), Customer Name, CD contact number and email |
| **Saller Master Sheet** | Contact Person Name and No (by the sender's email); blacklisted contacts are flagged |
| **MMV** tab | Make and Model as written in MMV; the variant and any other RC text go to **Specification** |
| **State List** | State, City, Zone from where the vehicle is lying |
| **CD Matrix** tab (from *Metrix*) | CD Contact Person Name, chosen by luxury / CV / state and seller |
| **BD Team** tab | Other Manager and BD Regional Head, filled only for surveyor-seller cases |
| **Seller Rules** tab (from *Seller T & C*) | End Date, ACR date & time, Bid Sheet Send Date (seller closing times, e.g. Go Digit 2:45 PM, ACR 3:00 PM) |
| Fixed values | Enable Vahan Yes, Auto Extend No / 0 / 0 / 0, Bid Limit 20, RC Available Yes, Owner Type Individual, Loan Paid Off No, Vehicle Condition Normal, Increment Type Fixed, Increment Amount 100 (2W) / 1000 (others), Start and Acknowledgement Date = today |

**Vehicle Type and Reg Type:** goods carriers, trucks and buses are *Commercial Vehicle*. A yellow-plate car is
*Passenger Carrying Vehicle*. Reg Type is *Private* only for 2-wheelers and 4-wheelers without a yellow plate;
everything else, including a yellow-plate 2-wheeler, is *Commercial*.

**Summary:** starts with any special conditions, then the tax status, then the default summary, e.g.
`Flood vehicle // OTT Paid // Buyer Needs to Prepare Affidavit…`, `Burnt Vehicle // Lifetime Tax Paid // …` or
`Tax details not mentioned please check on vahan // …`. Superdari, theft and TATA AIG Delhi-NCR notes from
*Seller T & C* are added the same way.

## QC checks

Cells that need checking are **highlighted yellow with a note**. Every mail also gets a row in the
**QC Report** tab (OK / CHECK / ERROR, with the reasons and a link to the mail). Checks include:

- number plate in the photos ≠ RC / mail registration number;
- vehicle in the photos ≠ RC model (e.g. RC says Swift, photos show Alto);
- transmission not visible (the cell is left blank; it's filled only when the photos or variant name show it);
- chassis / engine / model / year differ between subject, mail and RC; RC blurred or missing;
- fitness, permit or tax expired; hypothecation on the RC;
- seller not identified, contact blacklisted, city not in State List, model not in MMV, no CD person found;
- no quotation date in the mail (End Date assumed from the seller's TAT);
- reg number already in *my work* (possible duplicate);
- Excel / zip / HEIC attachments the tool couldn't read.

## Set up (once)

1. **Get a Claude API key** at <https://console.anthropic.com> (Settings → API keys). Usage is billed per
   mail: roughly a few US cents to about 25 cents, depending on how many photos it reads.
2. Upload the uploading sheet (`Uploading_sheet.xls`) to Google Drive and open it with **Google Sheets**
   (File → Save as Google Sheets). Keep the tab names: *my work*, *Seller id*, *Saller Master Sheet*, *MMV*,
   *State List*, *BD Team*.
3. In that sheet: **Extensions → Apps Script**, delete the sample code, paste in
   [`apps_script/UploadingSheetFiller.gs`](apps_script/UploadingSheetFiller.gs) and click **Save**.
   Use a separate script project from the Starred Mails Tracker (each has its own `onOpen` menu).
4. Reload the sheet. A new **Uploading Sheet** menu appears.
5. **Uploading Sheet → Set Claude API key…** and paste the key. It's saved for your Google account only.
6. **Uploading Sheet → Create helper tabs** adds **CD Matrix**, **Seller Rules** and **QC Report**.
   **Review the CD Matrix and Seller Rules tabs once** (see *Check these assumptions* below). You can edit them
   any time; the script reads them on every run.
7. The first run asks for permission (Gmail read, this spreadsheet, external request to Claude):
   **Continue → choose your account → Allow**.

## Daily use

- **Uploading Sheet → Fill from starred mails**: adds a row to *my work* for every starred mail that isn't
  in the QC Report yet. Each run stops starting new mails after about 4.5 minutes (Google's limit is 6); if
  some are left, run it again.
- **Uploading Sheet → Fill from one mail…**: to fill one mail, starred or not, paste its Gmail link or type
  a search (e.g. the reg number or claim number). You can type the Seller ID if you already know it.
- Check the yellow cells and the QC Report, then **File → Download → Microsoft Excel (.xlsx)** to upload.
  To redo a mail, delete its row in *my work* and in *QC Report* and run again.

## Check these assumptions

The *Metrix* tab is free-form, so the **CD Matrix** tab is my reading of it:

- **Rows are checked top to bottom, and the first match wins.** Luxury comes first (TN → Yanoke/G Dinesh,
  AP-TS → Kashish, rest → Ata), then CV pan India by seller (Sophia / Ankit / Vishal, and Ravi Shankar Tripathi
  for Surveyor/PSU), then the regions.
- **Surveyor / PSU cases outside AP-TS, TN and KL** go by state using the *Name / State* list under
  "Surveyor And Psu Except AP TS and TN KL" (Shubham, Pappu, Anuradha, Manish, Kamaldeep). Anything left over
  goes to Bhuwan (East/North) or Ravi Shankar Tripathi (South/West).
- **Karnataka:** SBI General goes to Rahul Kumar Rai, as in the Metrix.
- **Rajasthan:** one person (Shalini Kumari) handles all sellers.
- **East region:** Chhattisgarh and the North-East states are in the East block (Pankaj / Rachna).
  "Pankaj" is mapped to **Pankaj Kumar** from the contact list.
- **Delhi NCR:** Delhi, Gurgaon/Gurugram, Faridabad, Noida, Greater Noida and Ghaziabad.
- **CV bucket:** Commercial Vehicle, Commercial Equipment and Farm Equipment. Yellow-plate cars use the
  regional rows. You can change this in `US_CV_TYPES` at the top of the script.
- **Seller Rules → Email Domains:** these are the insurers' usual mail domains. A mail from a domain that isn't
  listed, sent from Gmail (or another personal mail id such as Yahoo), counts as *Surveyor Seller*. Any other
  unknown domain is flagged so you can add it to the right row.
- **Theft / Superdari:** Cardekho Region Id is set to 10 / 13 as in *Seller T & C*.

## Settings

At the top of the script (`US_CONFIG`): the model (`claude-opus-5`), effort, the number of photos sent per
mail (8, spread across the attachments), and the time budget per run.
