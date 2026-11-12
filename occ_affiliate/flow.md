# Affiliate OCC Daily: Power Automate build guide

This flow runs in the cloud under your own M365 identity, so your PC can be off. It uses only standard connectors: Office 365 Outlook and Excel Online (Business).

## One-time setup
1. Upload `Daily_Hanweck_Logs.xlsx` to OneDrive or SharePoint. From then on, that copy is the master. Office Scripts can't reach local or network-drive files.
2. Open the workbook in Excel for the web and go to **Automate → New Script**.
   - Paste in `office_scripts/ingestHanweck.ts` and save it as **ingestHanweck**.
   - Do the same with `applyApex.ts`, saving it as **applyApex**.
3. **Test the scripts on a copy of the workbook first** (see Verification below).

## Power Automate basics for this flow
- **Everything runs in sequence.** Add each step with **+ New step** (or the **+** under the last action). Don't use **Add a parallel branch**. Each step needs the previous step's output, and both scripts edit the same workbook.
- **Rename every action** as you add it: click **…**, then **Rename**. Use the names in **bold** below, because the expressions refer to them. A space in a name becomes `_` inside an expression, so the action *Get Hanweck* is written `Get_Hanweck`.
- **Enter expressions on the Expression tab.** Click into a field, open the **Expression** (fx) tab, paste the expression, and click **OK/Add**. Don't paste expressions in as plain text.
- **Watch out for "Apply to each".** If you pick an email field from the **Dynamic content** list, Power Automate quietly wraps your step in an **Apply to each** loop. The `first(...)` expressions below avoid that. If an Apply to each shows up anyway, delete it and use the expression.

## Build the flow: *Scheduled cloud flow*
**Trigger: Recurrence.** Set Interval `1` and Frequency `Week`. Under advanced options, set the time zone to Eastern, tick Mon–Fri, set the hour to `7` and the minute to `0`. On market holidays the flow won't find any emails and will email you "inputs missing". That's expected.

**1. Initialize variable** **Init hanweckId**: Name `hanweckId`, Type String, Value left blank.

**2. Initialize variable** **Init apexId**: Name `apexId`, Type String, Value left blank.

**3. Do until** **Wait for emails**. Its condition is an expression: `and(not(empty(variables('hanweckId'))), not(empty(variables('apexId'))))` *is equal to* `true`. Under **Change limits**, set Count to `13` and Timeout to `PT3H`, which means it re-checks every 15 minutes until about 10:00. Inside the loop, add these actions in order:

   **3a. Get emails (V3)** (Office 365 Outlook), named **Get Hanweck**:
   - Folder: Inbox
   - Include Attachments: **Yes**
   - Top: `1`
   - Search Query: `from:report@email.webull.com subject:"Webull US Accurate OCC Requirement" received:today`

   **3b. Get emails (V3)**, named **Get Apex**:
   - Folder: Inbox
   - Top: `1`
   - Search Query: `from:Tidal-NoReply@apexclearing.com subject:"OCC 157 Req" received:today`

   **3c. Set variable** `hanweckId` to the expression `if(empty(body('Get_Hanweck')?['value']), '', first(body('Get_Hanweck')?['value'])?['id'])`

   **3d. Set variable** `apexId` to the same expression, with `Get_Apex` in place of `Get_Hanweck`.

   **3e. Condition:** `or(empty(variables('hanweckId')), empty(variables('apexId')))` *is equal to* `true`. Under **If yes**, add a **Delay** of 15 minutes. Leave **If no** empty.

**4. Condition** **Inputs missing?**: the same `or(...)` expression as 3e.
- **If yes:**
  - **Send an email (V2)** to yourself, with subject `Affiliate OCC: inputs missing`.
  - Then **Terminate** with status Failed.
- **If no:** leave it empty.

**5. Scope** **Main**. Steps 5a–5f all go *inside* this Scope.

   **5a. Filter array** **Find CSV**:
   - From: `first(body('Get_Hanweck')?['value'])?['attachments']`
   - Condition (switch to **Edit in advanced mode**): `@startsWith(item()?['name'], 'apex-STANSATTR-BULL')`

   **5b. Run script** (Excel Online (Business)), named **Run ingestHanweck**:
   - Location: your OneDrive or SharePoint site
   - Document Library: as appropriate
   - File: `Daily_Hanweck_Logs.xlsx`
   - Script: **ingestHanweck**
   - `csv`: `base64ToString(first(body('Find_CSV'))?['contentBytes'])`

   **5c. Run script** **Run applyApex**:
   - Same file as 5b
   - Script: **applyApex**
   - `emailBody`: `first(body('Get_Apex')?['value'])?['body']`
   - `hanweckDate`: `outputs('Run_ingestHanweck')?['body/result/date']`

   **5d. Condition** **Diff over 3%?**: `outputs('Run_applyApex')?['body/result/alert']` *is equal to* `true`. Under **If yes**, add **Draft an email message (V2)**:
   - To: the risk stakeholders
   - Subject: `ALERT: Hanweck vs Apex OCC diff @{formatNumber(mul(outputs('Run_applyApex')?['body/result/diffPct'],100),'0.00')}% on @{outputs('Run_applyApex')?['body/result/date']}`
   - Body: `Hanweck @{outputs('Run_applyApex')?['body/result/hanweck']} vs Apex @{outputs('Run_applyApex')?['body/result/apex']}`

   **5e. Draft an email message (V2)** **Draft affiliate email**:
   - To / CC: the current distribution list. Copy it from your last sent email; it's kept out of the repo on purpose.
   - Subject: `Affiliate OCC 157 Requirement`
   - Body: switch the editor to code view (`</>`) and enter `@{outputs('Run_applyApex')?['body/result/html']}<br><br>Best,<br>Simon`

   The draft lands in your Outlook **Drafts** folder for you to review and send.

**6. Send an email (V2)** **Notify failure**. This goes *after* the Scope, not inside it.
- Open **…**, then **Configure run after**, untick *is successful*, and tick *has failed* and *has timed out*.
- To: you
- Subject: `Affiliate OCC flow failed`
- Body: `@{result('Main')}`

This is the standard Power Automate "try/catch": if any step inside **Main** fails, this email arrives with the error. The scripts' errors are written to be readable, for example "Date mismatch: Hanweck … vs Apex …" or "Unexpected STANSATTR header …".

## Test flow (the end-to-end test)
Use **Save As** to copy the flow, then change the copy as follows:
1. Replace the Recurrence trigger with **Manually trigger a flow**.
2. Delete **Wait for emails** (step 3) and **Inputs missing?** (step 4). Recreate the two **Get emails** actions (3a, 3b) at the top level, *before* **Main**, with the same names.
3. In both search queries, replace `received:today` with a fixed past date, for example `received:2026-11-12`. Use the day the emails *arrived*, which is one day after the data date.
4. Point both **Run script** actions at the **workbook copy**.
5. **Prepare the copy:** delete that day's rows from `data`, and clear that day's `pivot!L:S` and `email!C7`.
6. Click **Test**, then **Manually**, then **Run**. Check the following:
   - Every step is green.
   - The new rows are in `data`.
   - `pivot!L` and `email!C7` are filled in.
   - The draft matches the email you sent that day.
7. Click **Test** again. **Run ingestHanweck** should now show `skipped: true` in its outputs, and there should be no duplicate rows.

## Verification (before going live)
1. **Python reference.** From the repo root, install the dependencies once:

   ```bash
   pip install pytest openpyxl
   ```

   Then run the tests. In **PowerShell**:

   ```powershell
   $env:OCC_WORKBOOK = "C:\path\to\Daily_Hanweck_Logs.xlsx"; python -m pytest occ_affiliate\tests
   ```

   In **cmd**:

   ```bat
   set OCC_WORKBOOK=C:\path\to\Daily_Hanweck_Logs.xlsx && python -m pytest occ_affiliate\tests
   ```

   In **bash/macOS**:

   ```bash
   OCC_WORKBOOK=/path/to/Daily_Hanweck_Logs.xlsx python -m pytest occ_affiliate/tests
   ```

   This backtests the calculation against every day in the pivot history and the current `email` tab. Without `OCC_WORKBOOK`, the two workbook tests are skipped.
2. **Scripts on a copy.**
   1. In a copy of the workbook, delete the latest day's rows from `data`, and clear that day's `pivot!L:S` and `email!C7`.
   2. Add a sheet named `csv_input`. Open that day's `apex-STANSATTR-BULL_*.csv` in **Notepad**, select all, copy, and paste it into `csv_input!A1`.
   3. Go to **Automate → ingestHanweck → Run**. When it asks for `csv`, leave it **blank**, and the script reads from `csv_input`. Note that `csv` means the file's *text*, not a path, because Office Scripts can't open local files. In the flow, the text comes from the email attachment automatically.
   4. Go to **Automate → applyApex → Run**. Paste the Apex email body into `emailBody`; it doesn't matter if line breaks get lost. Set `hanweckDate` to that day as `YYYY-MM-DD`.
   5. The `email` table should match the one you sent manually. You can delete `csv_input` afterwards.
3. **Flow on the copy.** Run the **Test flow** above.
4. **Parallel run.** Run the flow alongside the manual process for about 2 weeks, then switch it to the master workbook.

## Notes
- If the `apex-STANSATTR` CSV header or date format changes, `ingestHanweck` fails loudly instead of appending bad data.
- Manual columns L–S on the `pivot` tab are aligned to pivot rows by position. Back-filling an *older* date would shift them, which is true of the manual process today too.
- `email!C10:J10` uses `MAX(prop × Apex total, 0)` without renormalizing. When any affiliate's share is negative, the allocations add up to more than the Apex total. On 2026-11-11 they came to $326.97M against $317.12M. The scripts keep this formula as-is.
