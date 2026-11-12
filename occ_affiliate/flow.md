# Affiliate OCC Daily: Power Automate build guide

This flow runs in the cloud under your own M365 identity, so your PC can be off. It uses only standard connectors: Office 365 Outlook and Excel Online (Business).

## One-time setup
1. Upload `Daily_Hanweck_Logs.xlsx` to OneDrive or SharePoint. From then on, that copy is the master. Office Scripts can't reach local or network-drive files.
2. Open the workbook in Excel for the web and go to **Automate → New Script**.
   - Paste in `office_scripts/ingestHanweck.ts` and save it as **ingestHanweck**.
   - Do the same with `applyApex.ts`, saving it as **applyApex**.
3. **Test the scripts on a copy of the workbook first** (see Verification below).

## Flow: *Scheduled cloud flow*
**Trigger: Recurrence.** Set Interval 1, Frequency Week, days Mon–Fri, at 07:00, with time zone Eastern. On market holidays it finds no emails and emails you "inputs missing". That's the intended behaviour.

**Initialize variables:**
- `hanweckId` (String) = `""`
- `apexId` (String) = `""`
- `tries` (Integer) = `0`

**Do until** `and(not(empty(variables('hanweckId'))), not(empty(variables('apexId'))))`, with limits Count 13 and Timeout `PT3H`. That is about 07:00 to 10:00 in 15-minute steps. Inside the loop:

1. **Get emails (V3)**, Office 365 Outlook, to find the Hanweck email:
   - Folder: Inbox
   - From: `report@email.webull.com`
   - Subject filter: `Webull US Accurate OCC Requirement`
   - Include attachments: **Yes**
   - Top: 1
   - Search query: `received:today`
2. **Get emails (V3)** to find the Apex email:
   - From: `Tidal-NoReply@apexclearing.com`
   - Subject filter: `OCC 157 Req`
   - Top: 1
   - Search query: `received:today`
3. **Set variable** `hanweckId` = `if(empty(body('Get_Hanweck')?['value']), '', first(body('Get_Hanweck')?['value'])?['id'])`
4. **Set variable** `apexId`, the same way using `Get_Apex`.
5. **Condition:** if either ID is empty, add a **Delay** of 15 minutes.

**Condition after the loop:** if either ID is still empty, use **Send an email (V2)** to send yourself *"Affiliate OCC: inputs missing (Hanweck: …, Apex: …)"*, then **Terminate**.

**Filter array**, to find the CSV:
- From: `first(body('Get_Hanweck')?['value'])?['attachments']`
- Condition: `startsWith(item()?['name'], 'apex-STANSATTR-BULL')`

**Run script** (Excel Online (Business)):
- Location: the OneDrive or SharePoint site
- File: `Daily_Hanweck_Logs.xlsx`
- Script: **ingestHanweck**
- `csv` = `base64ToString(first(body('Filter_array'))?['contentBytes'])`

**Run script**, renamed `Run_applyApex`:
- Script: **applyApex**
- `emailBody` = `first(body('Get_Apex')?['value'])?['body']`
- `hanweckDate` = `outputs('Run_script')?['body/result/date']`

**Condition:** `outputs('Run_applyApex')?['body/result/alert']` is equal to `true`. If it is, use **Draft an email message (V2)** to create the alert:
- To: the risk stakeholders
- Subject: `ALERT: Hanweck vs Apex OCC diff @{formatNumber(mul(outputs('Run_applyApex')?['body/result/diffPct'],100),'0.00')}% on @{outputs('Run_applyApex')?['body/result/date']}`
- Body: the Hanweck and Apex figures from the script result

**Draft an email message (V2)**, for the main email:
- To / CC: the current distribution list (copy it from your last sent email; kept out of the repo on purpose)
- Subject: `Affiliate OCC 157 Requirement`
- Body (HTML): `@{outputs('Run_applyApex')?['body/result/html']}<br><br>Best,<br>Simon`

The draft appears in your Outlook **Drafts** folder for you to review and send.

**Error handling:** add a parallel branch that runs after any step *has failed / has timed out*. It should **Send an email (V2)** to you with the expression `result('Run_applyApex')` or the failing action's error message. Both scripts throw clear errors, for example on a date mismatch, a CSV header change, or a missing email field.

## Verification (before going live)
1. **Python reference.** Run:

   ```bash
   OCC_WORKBOOK=/path/to/Daily_Hanweck_Logs.xlsx pytest occ_affiliate/tests
   ```

   This backtests the calculation against every day in the pivot history and the current `email` tab.
2. **Scripts on a copy.** In a copy of the workbook, delete the latest day's rows from `data` and clear that day's `pivot!L:S` and `email!C7`. Then, from the Automate pane, run `ingestHanweck` with that day's CSV text and `applyApex` with the Apex email text. The resulting `email` table should match the one you sent manually.
3. **Flow on the copy.** Point the flow at the copy and run it manually. Check the draft, then run it again and confirm `skipped: true`, meaning no duplicate rows were added.
4. **Parallel run.** Run the flow alongside the manual process for about 2 weeks, then switch it to the master workbook.

## Notes
- If the `apex-STANSATTR` CSV header or date format changes, `ingestHanweck` fails loudly instead of appending bad data.
- Manual columns L–S on the `pivot` tab are aligned to pivot rows by position. Back-filling an *older* date would shift them, which is true of the manual process today too.
- `email!C10:J10` uses `MAX(prop × Apex total, 0)` without renormalizing. When any affiliate's share is negative, the allocations add up to more than the Apex total. On 2026-11-11 they came to $326.97M against $317.12M. The scripts keep this formula as-is.
