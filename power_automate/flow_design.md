# Power Automate Flow: BCECN PDF → Master File

## Overview

This flow triggers automatically when a BCECN pricing email arrives in the
Bell treasury Outlook mailbox, extracts the PDF attachment, calls the Azure
Function parser, and writes the results to the Master File Excel workbook
stored on SharePoint.

---

## Prerequisites

| Resource | Details |
|---|---|
| Azure Function App | Deployed from `azure_function/` in this repo, Python 3.11, Consumption plan |
| Azure Function URL | `https://<app>.azurewebsites.net/api/parse` |
| Azure Function Key | Stored in Power Automate as a named credential or connection reference |
| SharePoint site | e.g. `https://bell.sharepoint.com/sites/Treasury` |
| Master File path | `/sites/Treasury/Shared Documents/BCECN/Master File.xlsx` |
| Office Script name | `BCECN Master File` (paste content of `office_script/bcecn_master.ts`) |

---

## Flow Steps

### Step 1 — Trigger: When a new email arrives (V3)

**Connector:** Office 365 Outlook
**Action:** When a new email arrives (V3)

```
Folder:           Inbox
Include Attachments: Yes
Only with Attachments: Yes
```

Add a filter to avoid triggering on unrelated emails:

```
Subject Filter:   BCECN
```

> At Bell's scale, consider using a dedicated shared mailbox
> (e.g. `treasury-bcecn@bell.ca`) and pointing the trigger at that mailbox.
> This avoids processing every email in a personal inbox.

---

### Step 2 — Apply to each (attachment loop)

**Action:** Apply to each
**Input:** `triggerOutputs()?['body/attachments']`

Wrap steps 3–7 inside this loop so each PDF attachment in the email is
processed independently.

**Condition inside loop (filter non-PDFs):**

```
Condition: endsWith(items('Apply_to_each')?['name'], '.pdf')
           is equal to true
```

Only proceed with the YES branch for PDF files.

---

### Step 3 — Call Azure Function (parse PDF)

**Connector:** HTTP
**Action:** HTTP

```
Method:       POST
URI:          https://<app>.azurewebsites.net/api/parse?code=<function-key>
Headers:
  Content-Type: application/json
Body (Expression):
  {
    "pdf_base64": "@{items('Apply_to_each')?['contentBytes']}",
    "filename":   "@{items('Apply_to_each')?['name']}"
  }
```

> `contentBytes` is already base64-encoded by Power Automate — pass it
> directly. No additional encoding step is needed.

**Settings → Retry Policy:**
```
Type:     Exponential interval
Count:    4
Interval: PT5S
```

---

### Step 4 — Check for parse errors

**Action:** Condition

```
Expression: outputs('HTTP')?['statusCode']
Operator:   is equal to
Value:      200
```

**NO branch:** Send an email alert (Step 4a) and terminate this iteration.

#### Step 4a — Send failure notification (NO branch only)

**Connector:** Office 365 Outlook
**Action:** Send an email (V2)

```
To:      treasury-alerts@bell.ca
Subject: BCECN Parser Error – @{items('Apply_to_each')?['name']}
Body:
  The BCECN PDF parser failed for attachment:
  @{items('Apply_to_each')?['name']}

  Error:
  @{body('HTTP')?['error']}

  Email received: @{triggerOutputs()?['body/receivedDateTime']}
```

---

### Step 5 — Parse the JSON response

**Action:** Parse JSON
**Content:** `body('HTTP')`

**Schema** (paste as-is into the Parse JSON action):

```json
{
  "type": "object",
  "properties": {
    "date":            { "type": "string" },
    "bank":            { "type": "string" },
    "cad_spread_3y":   { "type": ["number", "null"] },
    "cad_yield_3y":    { "type": ["number", "null"] },
    "cad_spread_5y":   { "type": ["number", "null"] },
    "cad_yield_5y":    { "type": ["number", "null"] },
    "cad_spread_7y":   { "type": ["number", "null"] },
    "cad_yield_7y":    { "type": ["number", "null"] },
    "cad_spread_10y":  { "type": ["number", "null"] },
    "cad_yield_10y":   { "type": ["number", "null"] },
    "cad_spread_30y":  { "type": ["number", "null"] },
    "cad_yield_30y":   { "type": ["number", "null"] },
    "usd_spread_3y":   { "type": ["number", "null"] },
    "usd_yield_3y":    { "type": ["number", "null"] },
    "usd_spread_5y":   { "type": ["number", "null"] },
    "usd_yield_5y":    { "type": ["number", "null"] },
    "usd_spread_7y":   { "type": ["number", "null"] },
    "usd_yield_7y":    { "type": ["number", "null"] },
    "usd_spread_10y":  { "type": ["number", "null"] },
    "usd_yield_10y":   { "type": ["number", "null"] },
    "usd_spread_30y":  { "type": ["number", "null"] },
    "usd_yield_30y":   { "type": ["number", "null"] },
    "cad_nc5_spread":  { "type": ["number", "null"] },
    "cad_nc5_coupon":  { "type": ["number", "null"] },
    "cad_nc10_spread": { "type": ["number", "null"] },
    "cad_nc10_coupon": { "type": ["number", "null"] },
    "usd_nc5_spread":  { "type": ["number", "null"] },
    "usd_nc5_coupon":  { "type": ["number", "null"] },
    "usd_nc10_spread": { "type": ["number", "null"] },
    "usd_nc10_coupon": { "type": ["number", "null"] }
  }
}
```

---

### Step 6 — Run Office Script (append row + update charts)

**Connector:** Excel Online (Business)
**Action:** Run script

```
Location:    SharePoint
Document Library: Shared Documents
File:        /BCECN/Master File.xlsx
Script:      BCECN Master File
```

**Script parameters:**

```
dataJson (Expression):
  string(body('Parse_JSON'))
```

> The Office Script receives the full parsed JSON as a single string, appends
> the row to "Pricing", and rebuilds the four yield-curve charts in
> "Summary Charts". Both operations happen in one Excel session, which avoids
> the file-lock issues you would get with two separate "Run script" calls.

---

### Step 7 — Check Office Script result

**Action:** Condition

```
Expression: outputs('Run_script')?['body/result']
Operator:   starts with
Value:      OK
```

**NO branch:** Send a failure notification (same pattern as Step 4a).

---

### Step 8 — Archive the processed email (optional, recommended)

**Connector:** Office 365 Outlook
**Action:** Move email (V2)

```
Message Id:         triggerOutputs()?['body/id']
Destination Folder: BCECN/Processed
```

Moving the processed email prevents it from being reprocessed if the flow
is re-run manually or on retry.

---

## Flow Variables (recommended)

Set these at the top of the flow with **Initialize variable** actions so
they are easy to change without editing every step:

| Variable name     | Type   | Value                                       |
|---|---|---|
| `FunctionUrl`     | String | `https://<app>.azurewebsites.net/api/parse` |
| `FunctionKey`     | String | `<function-key>` (use a secure input)       |
| `MasterFilePath`  | String | `/BCECN/Master File.xlsx`                   |
| `SharePointSite`  | String | `https://bell.sharepoint.com/sites/Treasury`|
| `AlertEmail`      | String | `treasury-alerts@bell.ca`                   |

---

## Deployment Checklist

- [ ] Deploy Azure Function App (Python 3.11, zip deploy from `azure_function/`
      + root `parsers/` directory)
- [ ] Copy the function key from Azure Portal → Function App → Functions →
      `parse` → Function Keys
- [ ] Upload `Master File.xlsx` to the SharePoint document library
- [ ] Open Excel Online → Automate tab → New script → paste
      `office_script/bcecn_master.ts` → save as "BCECN Master File"
- [ ] Create the Power Automate flow in
      `https://make.powerautomate.com` under the Bell tenant
- [ ] Set all flow variables with correct values
- [ ] Run a manual test with a sample PDF as the email attachment
- [ ] Verify the row appears in the Pricing sheet and charts update
