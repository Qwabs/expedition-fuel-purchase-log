# Google Sheets setup — year logs, sidebar, mobile Form and audit

`Fuel_Purchase_Log_Master.xlsx` is the template. The live file is the Google Sheet owned
by the designated owner Google account, because the Add Purchase sidebar, mobile Form and
automatic email/timestamp capture require Apps Script.

Each year has:

- **`2026 Log`** — the only entry sheet; one row per purchase, 5,000 rows.
- **`2026 Summary`** — monthly totals, purchaser totals and the year check.
- **`2026 Jan` … `2026 Dec`** — read-only weekly/monthly summaries.

The All Years sheet combines the annual summaries.

---

## 1. Owner and sharing

The Sheet, Form, script and triggers must be owned by the designated owner Google account.

1. Sign in to that account before importing or installing anything.
2. If the workbook currently belongs to another account, use **Share → Make owner**
   to transfer it first.
3. After setup, share the Google Sheet:
   - **Office editors:** Editor, invited individually.
   - **Field officers:** Editor, invited individually. They can use the desktop
     sidebar and see all purchases, amounts and audit emails.
   - **Auditors:** Viewer.
4. Do not use “Anyone with the link” for the Sheet. Named accounts are required for
   audit emails.
5. Personal Gmail accounts cannot use a Workspace Shared Drive. Keep periodic
   backups and plan an ownership transfer if the admin account changes. The Forms
   and installable triggers stop working if the owner account loses access.

## 2. Import the workbook

1. Upload `Fuel_Purchase_Log_Master.xlsx` to the owner’s Google Drive.
2. Right-click → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**.
4. Keep the `.xlsx` as the template/archive; enter data in the Google Sheet.

## 3. Install the scripts

1. In the Google Sheet: **Extensions → Apps Script**.
2. Delete the default `myFunction` code.
3. Paste `Code.gs` into `Code.gs`.
4. Create an HTML file: **+ → HTML**, name it exactly **Sidebar**, paste
   `Sidebar.html`.
5. Save the project as “Fuel Log”.
6. Run **setupAudit** once and approve the authorization prompts. It installs both the edit and change triggers.
7. Run **setupMobileForm** once and approve any additional Form/Drive prompts.

If you are upgrading an existing Sheet, replace `Code.gs` and run `setupAudit` again.
That installs the new `auditOnChange` trigger without changing any log rows.

Reload the Sheet. The **Fuel Log** menu should offer:

- **Add Purchase** — desktop sidebar.
- **Open current year log**
- **Mobile form link** — shows the responder URL and refreshes the Settings link.
- **Create / refresh mobile form**
- **Sync mobile form purchasers**
- **Audit logs** — read-only health scan of every `YYYY Log`.
- **Open Audit Log**
- **Install audit trigger**

The edit trigger is `auditOnEdit`, not `onEdit`, so it cannot run twice as a simple
and installable trigger. `setupAudit` also installs the `auditOnChange` trigger for
row/column insert and delete events. The Form trigger is `processFormResponse`.

## 4. The mobile Form

`setupMobileForm` creates a Google Form. Responses stay in the Form’s own
**Responses** tab; the form-submit trigger writes valid submissions into the year log.
The Form does not need to be linked to the 73-sheet workbook.

The form has:

- Purchase date (year included)
- Purchaser dropdown from Settings
- **Car Number (required)**
- Amount (GH₵)

It collects respondent email and requires sign-in. On submit, `processFormResponse`:

1. Checks the respondent is an owner/editor of the Sheet.
2. Validates date, purchaser, amount and car number.
3. Calls the same audit-first save as the sidebar.
4. Rejects invalid/unauthorised submissions into the Audit Log, without creating a
   purchase row.

If Google Forms does not automatically enable **Collect email addresses**, open the
Form → Settings and enable it; submissions without an email are rejected. The form
must be created while signed in as the designated owner account.

Only current Sheet editors/owner can successfully submit. Removing a field officer
from Sheet sharing therefore also stops their Form submissions. The raw responses
remain in the Form’s Responses tab; the year log is the source for reporting.

**After changing purchaser names in Settings**, run **Fuel Log → Sync mobile form
purchasers**. The form dropdown is a snapshot until you do.

**Form link:** the published responder URL is in **Settings → B67** after setup, or
via **Fuel Log → Mobile form link**. Share that responder link with field officers.
Never share the Form’s edit URL. Field staff need a Google account because respondent
email is collected.

The Form works from any desktop or mobile browser, anywhere with internet; no office
network or VPN is required. Bookmark the responder link if needed.

## 5. Car Number

- Required on every sidebar and Form entry.
- Free text, trimmed and uppercased: `gt 1234-24` → `GT 1234-24`.
- One car per row. If one payment covers several cars, enter one row per car.
- Direct log rows with a date and amount but no car number are highlighted amber.
- The Audit Log records the car number in column H.

## 6. How the save works

The sidebar and Form use one audit-first function:

1. Acquire the script lock.
2. Find the next log row.
3. Append the Audit Log entry.
4. Write the purchase row and F–G stamps.
5. If the purchase write fails, remove the audit row and return an error.
6. Flush only when both succeed.

A successful “Purchase added” message therefore means both records exist. Direct
corrections to log rows A–E are audited by `auditOnEdit`; the original Entered by/at
is not overwritten. Pasting several cells logs one row per changed field.

## 7. Weekly report and email

The **Fuel Log** menu includes:

- **Weekly report** — prompts for a week start and creates/refreshes the
  `Weekly Report` tab.
- **Email weekly report** — generates the report and emails it.
- **Set weekly report recipients** — saves approved addresses.

The report covers Monday–Sunday, all purchasers, with detail rows, total, purchaser
subtotals and car subtotals. A week crossing New Year is read from both year logs.
No workbook re-import is required; the script creates the report tab and the Settings
section on first use.

Recipients are stored in **Settings → B71** (comma-separated). Manual sends can
override the saved list. **Settings → B72** records the last successful send.

The email contains the formatted HTML breakdown and a PDF attachment when Google can
generate it. Recipients are BCC’d; the sender receives a copy. MailApp requires a
one-time authorization and is subject to the sender’s daily email quota. Sending is
manual; no scheduled trigger is installed.

The report reads the live `YYYY Log` at run time, so it includes every purchase
currently in the selected week, including entries added after the last export.

## 8. Protection

Google Sheets does not import the Excel protection. Set it there:

1. **Data → Protect sheets and ranges**.
2. Protect all **summaries** and **Settings** (leave Settings `A15:B64` and the
   weekly-recipient cell `B71` editable for the administrator).
3. Protect the **Audit Log** using **Show a warning**, not **Restrict who can edit**.
   The sidebar runs as the signed-in user and must be able to append its audit row.
4. Leave the year **Logs** editable for sidebar writes and corrections.
5. Test one sidebar entry after protecting. If it fails, loosen the protected range.

If you need a truly locked Audit Log, use a Form with an owner-run web app rather
than the user-run sidebar.

## 9. Log health and structural audit

Run **Fuel Log → Audit logs** at any time. This is a read-only scan of every
`YYYY Log`; it does not edit cells, rows or history.

It reports:

- **Row gaps** — blank rows between the first and last purchase.
- **Hidden rows** — rows hidden by a filter, a collapsed row group or the
  row-hiding setting, with the number of purchase rows not visible. This is the
  first thing to check when a log appears to have lost entries.
- **Possible duplicates** — rows with the same date, purchaser, amount and car number.
- **Missing date**, **invalid date**, **missing purchaser**, **missing amount** and
  **missing car number**.
- **Dates outside the log year**.

Results appear in a modal with row numbers. Correct the findings in the log, then run
the check again. The scan never repairs data automatically.

Filter views are personal; a normal filter hides rows for everyone. Use **Filter
views** when browsing a log, or turn the filter off before handing the Sheet to
another editor.

### Structural changes

`setupAudit` installs two installable triggers:

- `auditOnEdit` records changes to the purchase fields in the year logs.
- `auditOnChange` records `INSERT_ROW`, `REMOVE_ROW`, `INSERT_COLUMN` and
  `REMOVE_COLUMN` events in the Audit Log.

Google's change event identifies the event type, time and editor, but not the sheet
or row that changed. A structure-change audit row therefore uses `Workbook` as the
sheet and leaves the purchase fields blank. It proves that a row was inserted or
deleted; it cannot reconstruct the deleted values. Use **File → Version history** to
recover those rows, and copy only the missing rows into the current log rather than
restoring the entire workbook.

Keep the year logs editable for the sidebar and Form. Do not rely on sheet
protection alone to prevent row deletion; the edit and change audit plus version
history are the recovery path.

## 10. Test before real data

1. Add two purchaser names in **Settings → A15:A16**.
2. Run **setupMobileForm**, then open **Settings → B67** and confirm the Form link.
3. As an office editor, use **Add Purchase**: enter date, purchaser, car number and
   amount. Confirm the row appears in the correct year log, F–G are stamped, and the
   Audit Log shows **Added via sidebar**.
4. Add a second purchase on the same date with a different purchaser. Confirm both
   appear and the weekly summary includes both.
5. Submit the mobile Form from a second Google account that has Editor access.
   Confirm the row, car number, email, timestamp and **Added via form** audit row.
6. Submit the Form from an account without Sheet access; confirm it is rejected and
   no purchase row appears.
7. Edit a log amount; confirm the Audit Log records old/new and F–G stay unchanged.
8. Clear purchaser and amount; confirm F–G clear and **Cleared** is logged.
9. Run **Fuel Log → Weekly report** for the week containing the test entries.
   Confirm the detail, total, purchaser and car subtotals.
10. Set **Settings → B71** to your own email, run **Email weekly report**, and
    confirm the HTML body and PDF attachment arrive.
11. Run **Fuel Log → Audit logs**; confirm it reports no gaps, hidden rows,
    duplicates or missing fields.
12. Insert and delete a test row; confirm the Audit Log now contains a
    **Structure change** row for `INSERT_ROW` and `REMOVE_ROW`, then delete the
    test audit rows.
13. Delete the remaining test log rows and audit rows before real use.

## 11. Remote access

- **Owner/admin:** sign in to the designated owner account; open the Sheet or Form from
  Google Drive / `forms.google.com`; use the Settings hyperlink or Fuel Log menu.
- **Office/field editors:** open the shared Sheet link; use Fuel Log → Add Purchase
  on a desktop browser.
- **Phone field staff:** open the Form responder link; no Sheet access is needed for
  submission, though field officers also have Editor access as requested.
- Custom sidebars are desktop-web only; the Form is the mobile route.

## 12. Troubleshooting

**No Fuel Log menu** — reload; confirm the script is bound and `onOpen` exists.

**Sidebar says “Could not load”** — check the HTML file name is exactly `Sidebar`.

**No Audit Log rows** — check the installable triggers for `auditOnEdit` and
`auditOnChange`; rerun `setupAudit`.

**Audit logs is not in the Fuel Log menu** — replace `Code.gs` with the current
version, save, and reload the Sheet.

**Audit logs reports hidden rows** — a filter or filter view is hiding the rows.
Use **Data → Turn off filter**, close any active filter view, or select the row numbers
and choose **Unhide rows**. `COUNTA` including the hidden rows is the quick proof
that the purchases still exist; `SUBTOTAL(103, …)` counts only visible rows.

**Audit logs reports a row gap** — do not delete or re-import the workbook. Check
**File → Version history** and copy the missing rows into the current log. A gap is
blank row space, not evidence that the purchases were deleted.

**No Structure change audit rows after inserting or deleting a row** — confirm
`setupAudit` was run after the current `Code.gs` was installed. Only row/column
inserts and deletes are logged; filters, hiding rows and formatting are not
structural changes. Google does not provide the affected sheet or row in the change
event, so the audit row cannot include the deleted purchase values.

**Mobile Form has no purchaser choices** — add names in Settings, then run
**Sync mobile form purchasers**.

**`setupMobileForm` says the Purchaser question is not a dropdown** — open the
Google Form, delete the existing “Purchaser” question, and rerun
`setupMobileForm`. The script will recreate it as a dropdown.

**`setupMobileForm` reports `asListItem is not a function`** — you are running an
older `Code.gs`. Replace the entire file with the current `Code.gs`, save, and run
`setupMobileForm` again.

**`setupMobileForm` reports `dateItem.setRequired is not a function`** — same cause:
an older `Code.gs` is still loaded. Replace `Code.gs` completely, save, and rerun
`setupMobileForm`; the partially created Form is reused.

**`setupMobileForm` exceeds maximum execution time** — an older version tried to link
the Form to the 73-sheet workbook, which can hang. Replace `Code.gs` with the current
version and rerun. The Form is not linked to the Sheet; responses stay in the Form and
the trigger writes valid submissions into the year log.

**Form submissions are rejected** — the respondent must be an owner/editor of the
Sheet; check Sheet sharing. Check the Audit Log for the rejection reason.

**Form link missing** — run **setupMobileForm**, then use Fuel Log → Mobile form link.

**Weekly report menu is missing** — replace `Code.gs` with the current version,
save, and reload the Sheet.

**Weekly report reports a timeZone error** — replace `Code.gs` with the current
version; it falls back from the spreadsheet timezone to the script timezone. You can
also set **File → Settings → Time zone** to **GMT / Africa/Accra**.

**Weekly report has no entries** — confirm the selected week starts on Monday and the
purchases are in the matching `YYYY Log` dates. The report reads the live log.

**Weekly email has no recipients** — set **Settings → B71**, or use
**Fuel Log → Set weekly report recipients**. A manual send can still type addresses.

**Weekly email fails** — approve MailApp authorization, check the sender’s daily
email quota, and confirm addresses are valid. The report is still generated on the
Sheet even if email fails.

**No PDF attachment** — Google refused HTML-to-PDF conversion; the HTML report is
still sent. Check the Apps Script execution log.

**Purchase saved but audit error** — audit-first save should prevent this; check
protection and the Apps Script execution log.

**Entered by says `unknown`** — the user was not signed in or Google returned no email.

**Log is full** — 5,000 purchases for that year; extend the template or start the next
year block.

**Copied the Sheet** — triggers do not follow a copy. Run `setupAudit` and
`setupMobileForm` again in the copy.

## Privacy

The Audit Log stores email addresses. Form responses also contain respondent email
addresses in the Form. All Sheet editors can see all purchases and emails. Share only
with people who should see them. The Form records editors, not viewers; Drive sharing
remains the access control.

## Adding 2031 later

`add_year.py` extends the **`.xlsx` template**, not the live Google Sheet. Before 2031:

- ask for a Sheets-side year extension that updates All Years in the live file;
- import the new `.xlsx`, copy existing data, then reinstall triggers; or
- start a new 2031–2035 workbook and keep 2026–2030 as an archive.

Do not re-import the master over the live Sheet: it would replace the Sheet and its
Apps Script project.
