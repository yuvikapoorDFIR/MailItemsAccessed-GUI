<img width="606" height="362" alt="Designer (1)" src="https://github.com/user-attachments/assets/187ffcc7-51d2-44a7-aed8-cf715a15d9b9" />

<img width="1275" height="850" alt="image" src="https://github.com/user-attachments/assets/d930f151-d0fc-4a26-b01f-5a606185f40b" />

# Mail Items Accessed Correlator

**DFIR GUI Tool — UAL MailItemsAccessed × Microsoft Purview Content Search Correlation**

Developed by [Yuvi Kapoor](https://www.linkedin.com/in/yuvi-kapoor-5a38521a5/) | v1

---

## Overview

The **Mail Items Accessed Correlator** is a Windows GUI tool built for Digital Forensics and Incident Response (DFIR) analysts. It automates the correlation of `MailItemsAccessed` events from a Unified Audit Log (UAL) export against email metadata from a Microsoft Purview Content Search export — allowing you to identify and document exactly which emails a Threat Actor (TA) accessed during a mailbox compromise.

All processing runs **entirely locally**. No data is transmitted to any external service.

---

## Use Case

During a Business Email Compromise (BEC) or mailbox intrusion investigation, the Microsoft 365 UAL records `MailItemsAccessed` events that include `InternetMessageId` values for emails accessed via sync-protocol clients (MAPI/EWS/Outlook). By cross-referencing these IDs against a Purview Content Search export of the compromised mailbox, you can produce a definitive list of emails the TA read — suitable for inclusion in a TA Activity Booklet or evidence package.

---

## Requirements

| Requirement | Detail |
|---|---|
| Python | 3.9 or later |
| Operating System | Windows 10 / 11 |
| M365 Permissions | Audit Log access (Compliance Portal or PowerShell) |
| Purview Permissions | eDiscovery Manager or equivalent role |

> **Note:** `tkinter` ships with the standard Python Windows installer. If it is missing, reinstall Python and ensure the *tcl/tk* component is selected during setup.

---

## Installation

**1. Install Python dependencies:**

```cmd
pip install -r requirements.txt
```

Dependencies: `pandas`, `openpyxl`, `Pillow`

The tool will also attempt to auto-install missing packages on first launch.

**2. Launch the tool:**

```cmd
python mail_items_correlator.py
```

---

## Workflow

The tool follows a four-step workflow split across two sections.

---

### Section A — UAL Processing

#### Step 1: Prepare Your UAL Export

Before loading, your UAL export must be **pre-filtered to the TA's activity only**. Do not upload a full unfiltered UAL.

**Via Microsoft 365 Compliance Portal:**

1. Navigate to [compliance.microsoft.com](https://compliance.microsoft.com) → **Audit** → **Search**
2. Set your date range to cover the period of suspected compromise
3. Under **Activities**, select `MailItemsAccessed`
4. Under **Users**, enter the compromised account's UPN
5. Run the search and export results as CSV
6. Open in Excel and filter the `AuditData` column to rows containing the TA's confirmed session token(s) or client IP address(es)
7. Delete all other rows — the remaining rows are the TA's confirmed malicious activity
8. Save as CSV or XLSX — this is your input file

**Via PowerShell (Exchange Online):**

```powershell
Connect-ExchangeOnline

$results = Search-UnifiedAuditLog `
    -StartDate "2024-01-01" `
    -EndDate "2024-01-31" `
    -UserIds "compromised.user@domain.com" `
    -Operations "MailItemsAccessed" `
    -ResultSize 5000

$results | Export-Csv -Path "TA_UAL_Export.csv" -NoTypeInformation
```

> If the result set exceeds 5,000 rows, paginate using `-SessionId` and `-SessionCommand ReturnLargeSet`.

#### Step 2: Load UAL Export

1. Click **Load UAL Export** in the sidebar
2. Click **Browse UAL File** and select your pre-filtered CSV or XLSX
3. The tool auto-detects column names and runs an immediate file inspection — a green or red banner confirms whether the file is valid and how many `MailItemsAccessed` rows were found
4. Adjust the column mapping fields if your export uses non-standard column names
5. Click **Continue: Extract Message IDs**

#### Step 3: Extract Message IDs

1. Optionally enter a **Session Token** substring to filter to a specific TA session
2. Optionally enter a **UserId** substring to restrict to a specific account
3. Click **Run Extraction**
4. The tool parses every `AuditData` JSON payload, walks the `Folders → FolderItems → InternetMessageId` structure, and deduplicates all extracted IDs
5. A diagnostics panel shows columns detected, row counts, REST-only events, parse errors, and unique IDs extracted

> **REST-only events:** When a TA accesses mail via OWA, a mobile app, or the Microsoft Graph API, M365 logs `MailItemsAccessed` but the `Folders` array is empty — no `InternetMessageId` is recorded. Only sync-protocol access (MAPI, EWS, Outlook desktop) populates `FolderItems`. If extraction returns 0 IDs with many REST-only events, the TA used a client that does not log message-level detail.

---

### Section B — Content Search Correlation

#### Step 1: Run a Purview Content Search

1. Navigate to [compliance.microsoft.com](https://compliance.microsoft.com) → **Content search** → **New search**
2. Name the search (e.g. `MailboxExport_InternetMessageIDs_Username`)
3. Under **Locations**, choose **Specific users or groups** and add:
   - The TA's primary mailbox
   - Any shared or delegated mailboxes the TA had access to
4. Under **Conditions**, **remove** the default Keyword condition entirely — leave it blank to capture all message metadata
5. Click **Submit** and wait for status to show **Completed**

#### Step 2: Export the Items Report

1. Click the search name → **Actions** → **Export results**
2. Select **Export report only (Items report)** — do **not** include message content or PST
3. Click **Export**, then switch to the **Export** tab and download the ZIP when ready
4. Extract the ZIP and locate `Items.csv` inside

#### Step 3: Load Items.csv

1. Click **Load Items.csv** in the sidebar
2. Select the extracted `Items.csv`
3. The tool auto-detects Internet Message ID, Subject, Date Sent, Recipients, and Original Path columns
4. Adjust column mappings via the dropdowns if needed
5. Click **Continue: Correlate & Export**

---

### Correlate & Export

The tool matches every `InternetMessageId` in Items.csv against the UAL-extracted set. Matching is normalised — case-insensitive, angle brackets stripped.

- **Green rows** = confirmed TA-accessed emails
- Filter: All / Matched / Unmatched
- Search bar filters across Subject, Recipients, and Message ID in real time
- Click **Export Results (.xlsx)** to download the output workbook

---

## Output XLSX Structure

| Sheet | Contents |
|---|---|
| **All Items** | Full Items.csv with a `TA Accessed` column (YES/NO). Matched rows highlighted green. Frozen headers, auto-width columns. |
| **TA Accessed (Matched)** | Confirmed TA-accessed emails only — ready for direct inclusion in a TA Activity Booklet. |
| **UAL Message IDs** | All unique `InternetMessageId` values extracted from the UAL. |

**Recommended post-export steps:**
- Sort the **TA Accessed** sheet by `Email date sent` to reconstruct the chronological access sequence
- Cross-reference `Email recipients` with known internal targets to assess data exposure scope
- Retain the **UAL Message IDs** sheet as an evidentiary artefact

---

## Column Mapping Reference

| Field | Common Column Names |
|---|---|
| AuditData | `AuditData`, `Audit Data`, `auditdata`, `audit_data` |
| Operations | `Operations`, `Operation` |
| UserId | `UserId`, `UserID` |
| CreationDate | `CreationDate`, `Date` |

Purview Items.csv columns are auto-detected via keyword matching and can be overridden via dropdown on the Load Items.csv screen.

---

## Troubleshooting

| Symptom | Likely Cause | Resolution |
|---|---|---|
| 0 IDs extracted, many REST-only events | TA used OWA / Graph API — no `FolderItems` logged | Cannot extract IDs from REST events; consider MessageTrace logs |
| 0 IDs extracted, 0 MailItemsAccessed rows | Wrong Operation column, or file not pre-filtered | Check column mapping; confirm file contains `MailItemsAccessed` rows |
| JSON parse errors in diagnostics | AuditData column contains extra quote wrapping | Tool handles this automatically; if errors persist, re-export from M365 |
| XLSX loads slowly or hangs | Large file; openpyxl memory usage | Save as CSV from Excel (File → Save As → CSV UTF-8) for faster loading |
| 0 correlation matches | ID format mismatch or wrong column selected | Verify correct Message ID column is selected in Items.csv mapping |
| Favicon not showing | PIL ICO conversion failed | Non-critical — app still functions. Ensure Pillow is installed. |

---

## Data Privacy

- **No network calls are made** — all processing is local and offline
- No data is written to disk other than the XLSX export you explicitly request
- The tool does not log, cache, or retain any file contents between sessions

---

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | DataFrame operations for CSV/XLSX loading and manipulation |
| `openpyxl` | XLSX reading and styled XLSX export |
| `Pillow` | Logo image processing and ICO favicon generation |

---

*Mail Items Accessed Correlator — v1 — Created by [Yuvi Kapoor](https://www.linkedin.com/in/yuvi-kapoor-5a38521a5/)*
