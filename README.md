# Salesforce Constituent Data Cleanup Tools v4

Tools for extracting missing contact information from Salesforce case text fields, matching unlinked cases to existing contacts, and identifying Outlook emails that never made it into the system.

## The Problem

Our Salesforce has **6,669 cases** but only 38% have a Contact linked. Many constituent emails, phone numbers, addresses, and names are buried in case Description, Subject, and Notes fields, typed in by staff during phone intake, pasted from email threads, or submitted through web forms. This toolkit extracts that data, matches it to existing contacts, and tells our team exactly what to do with each case.

## Tools

| Script | Purpose |
|--------|---------|
| `constituent_extractor_v4.py` | Extract contact info from case text, match to contacts, recommend actions |
| `outlook_matcher_v2.py` | Compare Outlook inbox export against Salesforce to find missed emails |

## Quick Start

```bash
# Install dependencies
pip install pandas openpyxl rapidfuzz anthropic python-dotenv

# Run the extractor (regex-only — free, fast, ~10 seconds)
python scripts/constituent_extractor_v4.py data/cases.xlsx data/contacts.xlsx -o output/results.xlsx

# Run the extractor with LLM disambiguation for ambiguous cases (~$0.30-0.50)
# PowerShell:
$env:ANTHROPIC_API_KEY="your_key_here"
python scripts/constituent_extractor_v4.py data/cases.xlsx data/contacts.xlsx -o output/results.xlsx --llm
# Or pass it directly:
python scripts/constituent_extractor_v4.py data/cases.xlsx data/contacts.xlsx -o output/results.xlsx --llm --api-key sk-ant-...

# Run the Outlook matcher
python scripts/outlook_matcher_v2.py data/cases.xlsx data/contacts.xlsx data/outlook_export.csv -o output/outlook_results.xlsx

# Cross-reference Outlook matcher with extractor output
python scripts/outlook_matcher_v2.py data/cases.xlsx data/contacts.xlsx data/outlook_export.csv \
    -o output/outlook_results.xlsx --extractor-output output/results.xlsx --after 2024-01-01
```

### Project Layout

```
salesforce-cleanup/
├── data/
│   ├── cases.xlsx          # Salesforce cases export
│   └── contacts.xlsx       # Salesforce contacts export
├── output/                 # Generated results go here
├── scripts/
│   ├── constituent_extractor_v4.py
│   └── outlook_matcher_v2.py
├── venv/                   # Python virtual environment
├── README.md
└── requirements.txt
```

### Input Formats

Both scripts auto-detect CSV and XLSX — even if the extension is wrong, they'll try both formats. No conversion needed.

**Cases export** — Standard Salesforce case export with columns: `Case Number`, `Subject`, `Description`, `Case Notes`, `Case Comments`, `Contact Name`, `Contact: Email`, `Web Email`, `Contact ID`, `Contact: Phone`.

**Contacts export** — Standard Salesforce contacts export with columns: `Contact ID`, `First Name`, `Last Name`, `Email`, `Phone`, `Mobile`, `Mailing Street`.

**Outlook export** — CSV/XLSX from Access, VBA macro, or Power Automate with sender email, subject, and date columns. Column names are auto-detected.

---

## Constituent Extractor v4

### What It Does

1. **Extracts contact info** from all text fields using regex patterns for emails, phones, addresses, postal codes, and names
2. **Detects person blocks** — groups of contact info belonging to one person (signature blocks, contact cards, call log entries)
3. **Identifies multi-person cases** — call logs or notes mentioning several constituents with separate contact info
4. **Parses structured data** — `[Parsed Address]` fields in Case Notes are extracted directly, not regex'd
5. **Matches unlinked cases** to existing Salesforce contacts by exact email, fuzzy email (Levenshtein distance), or phone number
6. **Disambiguates with LLM** (optional) — for cases with multiple emails or unclear constituent identity, Claude Haiku determines which info belongs to the constituent

### Actions

The extractor assigns each case an action. Sheets are ordered by staff priority — work top to bottom.

| Action | What To Do |
|--------|-----------|
| `LINK_TO_EXISTING` | Email or phone matches an existing contact. Just link the case to that contact in Salesforce. |
| `FIX_CONTAMINATION_RECOVERABLE` | Contact email is a staff email, but the real constituent email was found in the case text. Replace it. |
| `CREATE_CONTACT` | No contact linked, but a clear email was extracted. Create a new contact and link. |
| `CAN_ENRICH` | Contact is linked but we found phone, address, or name in the text that the contact doesn't have. Add it. |
| `CREATE_CONTACT_REVIEW` | Email was found but the case is ambiguous (multi-person, email chain). Review before creating. |
| `FIX_CONTAMINATION_MANUAL` | Staff email contamination with no clear replacement. Manual investigation needed. |
| `HAS_PHONE_ONLY` | Phone number found but no email. Can create contact if you have a name, or match manually. |
| `HAS_PARTIAL_INFO` | Address or name found but no email or phone. Limited actionability. |
| `NO_CONTACT_INFO` | Has text but nothing extractable. |
| `NO_TEXT` | No text fields populated. |
| `COMPLETE` | Contact linked, no additional info found. Nothing to do. |

### Output Sheets

| Sheet | Contents |
|-------|----------|
| **Dashboard** | Summary stats, action counts with descriptions, extractable field counts |
| **Link** | `LINK_TO_EXISTING` cases — includes matched Contact ID and match method |
| **Fix-Recover** | Staff contamination with recoverable email |
| **Create** | Cases where a new contact should be created |
| **Enrich** | Linked contacts that could have fields added |
| **Create-Review** | Ambiguous cases needing human judgment |
| **Fix-Manual** | Staff contamination needing investigation |
| **Phone-Only** | Cases with phone but no email |
| **Partial** | Cases with address/name but no email or phone |
| **Multi-Person** | Cases where multiple people were detected in the text |
| **Duplicates** | Email addresses appearing in multiple unlinked cases |
| **All Cases** | Full dataset with all extraction columns |

### Key Output Columns

| Column | Description |
|--------|-------------|
| `Best_Email` | Best-guess constituent email (from Web Email, person block, or LLM) |
| `Best_Phone` | Best-guess phone (prefers valid 10-digit, constituent-flagged) |
| `Best_Address` | Best-guess address (prefers parsed, then constituent-context, then single) |
| `Best_Name` | Best-guess name (from person blocks, signatures, title-prefix patterns) |
| `Person_Count` | Number of distinct people detected in the case text |
| `Persons` | Structured breakdown of each person's contact info and source pattern |
| `Matched_Contact_ID` | For unlinked cases, the existing Contact ID that matched |
| `Match_Method` | How the match was made: `exact_email`, `fuzzy_email(92%)`, `phone` |
| `Flags` | Diagnostic flags: `EMAIL_CHAIN`, `MULTI_PERSON(3)`, `HAS_PARSED_ADDR`, etc. |

### Extraction Patterns

**Emails** — Standard regex, staff emails filtered out (`@toronto.ca`, `@diannesaxe.ca`, `@envirolaw.com`).

**Phones** — 10-digit (with area code validation) → `valid`. 7-digit → `partial`. Handles `(416) 555-1234`, `416.555.1234`, `416-555-1234`, and `Cell: 416-555-1234` label formats.

**Addresses** — Five regex patterns covering: `123 Main St`, `690 MANNING AVE`, `120 macpherson ave`, `Unit 5, 123 Main St`, and `St. George St` / `St. Clair Ave` names. False positive filtering rejects `20 minute walk`, `4 dogs`, etc. Constituent context detection flags addresses near "I live at" / "my property at" phrases with `[CONST]`.

**Postal codes** — Toronto format `M#L #L#` validated. Invalid formats still extracted but marked `✗`.

**Names** — Extracted from "my name is...", "I'm...", sign-offs (`Regards,\nJohn Smith`), signature blocks (name on line before email/phone), `Mr./Mrs./Dr.` prefixed lines, and end-of-text patterns. Validated as 2-5 capitalized words.

**Person blocks** — Contact info grouped by proximity and structure:
- **Contact block at start**: `Name\nAddress\nPhone` (common in phone intake notes)
- **Signature block at end**: `Regards,\nName\nCell: 416-...\nEmail: x@y.com`
- **Call log entries**: `Name 5:00pm\nAddress\nPhone\nnotes`
- **Inline entries**: Name on its own line followed by phone/address/email on next lines

### Matching

The extractor builds a contact index from both your Contacts export and linked cases. Three match methods:

| Method | Description |
|--------|-------------|
| **Exact email** | Direct email match against 3,146 known emails |
| **Fuzzy email** | Same domain + ≥85% prefix similarity (catches `annestevenz` ↔ `anne.stevenz`) |
| **Phone** | 10-digit phone match against 806 known numbers |

### LLM Disambiguation

When `--llm` is enabled, Claude Haiku is called for cases where regex alone can't determine the constituent:
- Multiple emails extracted (which one is the constituent's?)
- Multiple addresses without constituent context clues
- Email chains where forwarded content may confuse extraction

Estimated cost: ~$0.30–0.50 for a full run. Use `--llm-limit N` to cap calls.

---

## Outlook Matcher v2

### What It Does

Compares every sender in your Outlook export against all known Salesforce emails to find people who emailed the office but never got a case or contact created.

### Improvements Over v1

- **Contacts file as input** — matches against 3,000+ contacts, not just case emails
- **Fuzzy matching** — catches email typos and alias variations
- **Frequency analysis** — surfaces repeat senders first (someone who emailed 15 times and got no case is a bigger gap than a one-off)
- **Sender classification** — constituent / bulk / org based on domain patterns
- **Cross-references extractor output** — if v4 already found the email in a case, it's flagged separately to avoid duplicate work
- **Date range filtering** — `--after` and `--before` flags

### Output Sheets

| Sheet | Contents |
|-------|----------|
| **Dashboard** | Status counts with message volume |
| **Not In Salesforce** | Unknown senders sorted by frequency — the main action list |
| **Fuzzy Matches** | Probable matches that need human verification |
| **In Extractor** | Emails found by v4 but not yet imported to Salesforce |
| **Top Senders** | 50 most frequent senders regardless of status |
| **All Senders** | Complete results |

---

## Recommended Workflow

```
1.  Export cases and contacts from Salesforce (CSV or XLSX, either works)
    Place them in the data/ folder.

2.  Run the extractor:
      python scripts/constituent_extractor_v4.py data/cases.xlsx data/contacts.xlsx -o output/results.xlsx

3.  Work through the output sheets in order:
      a. Link       — Connect unlinked cases to existing contacts (709 cases)
      b. Fix-Recover — Replace staff emails with real ones (9 cases)
      c. Create     — Make new contacts and link (1,042 cases)
      d. Enrich     — Add phone/address to existing contacts (839 cases)
      e. Create-Review — Cases needing judgment (38 cases)

4.  (Optional) Run with LLM for the ambiguous cases:
      python scripts/constituent_extractor_v4.py data/cases.xlsx data/contacts.xlsx -o output/results_llm.xlsx --llm

5.  (Optional) Run the Outlook matcher to find missed emails:
      python scripts/outlook_matcher_v2.py data/cases.xlsx data/contacts.xlsx data/outlook.csv \
          -o output/outlook.xlsx --extractor-output output/results.xlsx

6.  Import changes to Salesforce
```

## Dependencies

```
pandas>=1.5.0         # Data processing
openpyxl>=3.0.0       # Excel output
rapidfuzz>=3.0.0      # Fuzzy email matching (optional but recommended)
anthropic>=0.18.0     # LLM disambiguation (only needed with --llm)
python-dotenv         # .env file support (optional)
```

Install all: `pip install pandas openpyxl rapidfuzz anthropic python-dotenv`

## Files

```
scripts/constituent_extractor_v4.py   # Main extraction and matching script
scripts/outlook_matcher_v2.py         # Outlook-to-Salesforce email matcher
README.md                             # This file
requirements.txt                      # Python dependencies
```
