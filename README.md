# Salesforce Constituent Data Cleanup Tools v3

Extracts missing contact info from Salesforce case text fields with LLM disambiguation for ambiguous cases.

## Quick Start

```bash
# Install dependencies
pip install pandas openpyxl anthropic

# Run extraction (regex only - free, fast)
python constituent_extractor_v3.py cases.csv output.xlsx --no-llm

# Run extraction with LLM disambiguation (~$0.30-0.50)
export ANTHROPIC_API_KEY=your_key_here
python constituent_extractor_v3.py cases.csv output.xlsx

# Optional: limit LLM calls
python constituent_extractor_v3.py cases.csv output.xlsx --llm-limit 200
```

## What It Does

### 1. Extracts Contact Info
- **From all text fields**: Subject, Description, Case Notes, Case Comments
- **Email addresses**: Filters out staff emails (@toronto.ca, @diannesaxe.ca)
- **Phone numbers**: Validates 10-digit (full) vs 7-digit (partial)
- **Street addresses**: Multiple patterns for "123 Main St", "690 MANNING AVE", "St. George St", lowercase, etc.
- **Postal codes**: Validates Toronto format (M#L #L#)
- **Names**: From signatures, "my name is...", sign-offs

### 2. LLM Disambiguation (Optional)
For ambiguous cases, Claude Haiku determines:
- Which email belongs to the constituent vs forwarded senders
- Which address is the constituent's home vs the complaint location
- Confidence level and reasoning

**Triggers LLM when:**
- Multiple email addresses found
- Multiple addresses found without clear constituent context
- Email chains with extracted content

**Estimated cost:** ~$0.30-0.50 for ~400-500 LLM calls

### 3. Smart Matching
- **Matches unlinked cases to existing Contacts** by email
- **Groups duplicate unlinked cases** that likely belong to same person
- **Detects email chains** that may contain forwarded content
- **Flags constituent vs complaint addresses** based on context ("I live at" vs "the property at")

## Actions

| Action | What It Means |
|--------|---------------|
| COMPLETE | Has Contact ID, no fields to add |
| CAN_ENRICH | Has Contact ID but we found phone/address/name in text |
| LINK_TO_EXISTING | No Contact ID but email matches existing Contact |
| CREATE_CONTACT | No Contact ID, good email found |
| CREATE_CONTACT_REVIEW | No Contact ID, has email but needs review (email chain/multiple) |
| FIX_CONTAMINATION_CAN_RECOVER | Staff email contamination but real email found |
| FIX_CONTAMINATION_MANUAL | Staff email contamination, manual fix needed |
| HAS_PHONE_ONLY | Phone found but no email |
| NO_CONTACT_INFO | Has text but no extractable contact info |
| NO_TEXT | No text fields to extract from |

## Output Columns

| Column | Description |
|--------|-------------|
| Action | Recommended action for this case |
| Confidence | high/medium/low/none |
| Missing_Fields | What the Contact doesn't have |
| Extractable_Fields | What we found that could be added |
| SF_Has_Name/Email/Phone | Yes/No - does linked Contact have this? |
| SF_Is_Contaminated | Yes if Contact email is staff email |
| Extracted_Emails | Emails found (staff filtered out) |
| Extracted_Phones | Format: "valid:(416) 555-1234" or "partial:555-1234" |
| Extracted_Addresses | Prefixed [CONST] if context suggests constituent's home |
| Extracted_Postals | ✓ = valid Toronto format, ✗ = invalid |
| Extracted_Names | Format: "Name (pattern_type)" |
| LLM_Email/Phone/Address | LLM's determination (if used) |
| LLM_Complaint_Addr | Complaint location distinct from home (if found) |
| LLM_Confidence | LLM's confidence in its determination |
| LLM_Reasoning | Brief explanation of LLM's logic |
| Matched_Contact_ID | For unlinked cases, existing Contact ID that matches |
| Flags | EMAIL_CHAIN, MULTIPLE_EMAILS, MULTIPLE_ADDRESSES, DUP_GROUP, etc. |

## Excel Sheets

- **Summary** - Statistics and action counts
- **All Cases** - Full data, color-coded by action
- **Enrich** - Cases with Contact ID but missing extractable fields
- **Link** - Unlinked cases matching existing Contacts
- **Create** - Unlinked cases with good extracted data
- **Create-Review** - Unlinked cases needing review
- **Fix-Recover** - Staff contaminated with recoverable email
- **Fix-Manual** - Staff contaminated needing manual fix
- **Duplicates** - Emails appearing in multiple unlinked cases

## Validation Rules

### Postal Codes
- Must be Toronto format: M#L #L# (e.g., M5V 1A1)
- Invalid formats still shown but marked with ✗

### Phone Numbers
- 10-digit with valid area code → "valid"
- 7-digit (missing area code) → "partial"
- Invalid patterns → not extracted

### Addresses
- Multiple patterns for various formats
- False positive filtering (rejects "20 minute walk", "4 dogs", etc.)
- Constituent context detection ("I live at" → [CONST] prefix)

### Names
- From signatures, "my name is...", sign-offs
- Cleaned of trailing junk ("On Behalf Of", etc.)
- Validated as 2-5 capitalized words

## Files

```
constituent_extractor_v3.py   # Main extraction script
outlook_matcher.py            # Outlook-to-Salesforce email matcher
README.md                     # This file
sample_output.xlsx           # Sample output for reference
```

## Dependencies

```
pandas>=1.5.0
openpyxl>=3.0.0
anthropic>=0.18.0  # Only needed for LLM features
```

## Example Workflow

```bash
# Step 1: Run without LLM to review
python constituent_extractor_v3.py cases.csv review.xlsx --no-llm

# Step 2: Check the output, especially CAN_ENRICH and CREATE_CONTACT sheets

# Step 3: Run with LLM for ambiguous cases
python constituent_extractor_v3.py cases.csv final.xlsx

# Step 4: Work through sheets in priority order:
#   1. Fix-Recover (staff contamination with easy fix)
#   2. Link (just need to connect case to existing Contact)
#   3. Create (make new Contacts)
#   4. Enrich (add phone/address to existing Contacts)
#   5. Create-Review (needs human judgment)
#   6. Fix-Manual (staff contamination, manual work)
```
