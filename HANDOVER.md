# Premium Summary Tool -- Technical Handover Document

**Prepared for:** IT Development Team
**Date:** 14 April 2026
**Repo:** `https://github.com/jasperh0917/premiumsummary.git`
**Live:** Deployed on Vercel (auto-deploys from `main` branch)

---

## 1. WHAT THIS TOOL DOES

The Premium Summary Tool automates the end-to-end group health insurance policy onboarding workflow for UAE-based brokers. It takes two inputs:

1. **Census file** (XLSX/CSV) -- the member list with names, DOBs, genders, categories, etc.
2. **Rates PDF** -- the insurer's quote document containing premium rate tables per age bracket

It then:
- Extracts rates from the PDF using Claude Vision AI
- Calculates per-member premiums (base + maternity surcharge)
- Compares the calculated total against the quoted premium (reconciliation)
- Generates a branded Excel workbook (Premium Summary + Premium per Member + Reconciliation sheets)
- Persists everything to a Supabase PostgreSQL database for audit, search, and post-issuance member management

---

## 2. WHY YOU'RE TAKING THIS OVER

The **PDF upload step is being replaced**. Your quote tool already has the rate tables, so there's no need to extract them from a PDF. The integration point is:

> **Instead of uploading a PDF and using Claude Vision to extract rates, the tool should receive rates directly from the quote tool's API.**

This eliminates:
- The Anthropic API dependency (and its cost: ~$0.05-0.15 per policy)
- The PDF parsing pipeline (`parse_rates_pdf`, `parse_rates_xlsx`, Claude Vision calls)
- The rate verification step in the UI (since rates come from a trusted source)

Everything downstream of rate extraction stays the same: premium calculation, Excel generation, database persistence, reconciliation, dashboard.

---

## 3. ARCHITECTURE OVERVIEW

```
                  CURRENT FLOW
                  ============
                  
User uploads       Claude Vision        Premium         Excel          Supabase
Census + PDF  -->  extracts rates  -->  Calculation --> Generation --> Database
                                             |
                                        Reconciliation
                                        (calc vs quote)


                  FUTURE FLOW (with Quote Tool)
                  ==============================
                  
Quote Tool API     Census upload        Premium         Excel          Supabase
sends rates   -->  (same as now)  -->  Calculation --> Generation --> Database
                                             |
                                        Reconciliation
                                        (calc vs quote)
```

---

## 4. TECH STACK

| Layer | Technology | Notes |
|-------|-----------|-------|
| Backend | Flask (Python 3.9+) | Single file: `app.py` (4,342 lines) |
| Frontend | Vanilla HTML/CSS/JS | 5 templates in `templates/` |
| Database | Supabase PostgreSQL | 6 tables (see schema below) |
| AI (current) | Anthropic Claude 3.5 Sonnet | PDF rate extraction only -- will be removed |
| Excel | openpyxl | Generates branded .xlsx workbooks |
| PDF parsing | PyMuPDF + pdfplumber | Will be removed with quote tool integration |
| Hosting | Vercel (serverless Python) | Auto-deploys from `main` branch |

### Dependencies (`requirements.txt`)
```
Flask
PyMuPDF          # PDF rendering -- can be removed after integration
pandas           # Census parsing (XLSX/CSV)
openpyxl         # Excel generation
anthropic        # Claude API -- can be removed after integration
xlrd             # Legacy .xls support
msoffcrypto-tool # Password-protected Excel files
supabase         # Database client
Pillow           # Image processing for PDF
```

### Environment Variables (set in Vercel dashboard)
```
ANTHROPIC_API_KEY=sk-ant-...     # Will be removed
SUPABASE_URL=https://xxx.supabase.co
SUPABASE_ANON_KEY=eyJh...
SUPABASE_SERVICE_ROLE_KEY=eyJh... # Preferred -- bypasses RLS
```

---

## 5. DATABASE SCHEMA

### `policies` -- main policy record
| Column | Type | Description |
|--------|------|-------------|
| id | BIGINT PK | Auto-generated |
| company_name | TEXT | Policyholder name |
| broker | TEXT | Broker company |
| underwriter | TEXT | Underwriter name |
| rm_person | TEXT | Relationship manager |
| plan | TEXT | "Healthx", "Healthxclusive", "OpenX", "DOH" |
| plan_type | TEXT | "New Business", "Renewal" |
| start_date | DATE | Policy inception date |
| confirmation_date | DATE | When confirmed |
| inception_payment | TEXT | "Annual", "Semi-Annual", "Quarterly", "Monthly" |
| endorsement_freq | TEXT | "Monthly", "Quarterly" |
| has_lsb | BOOLEAN | Low Salary Band flag |
| rm_broker | NUMERIC | Broker commission % |
| rm_insurer | NUMERIC | Insurer (QIC/DNI) % |
| rm_wellx | NUMERIC | Admin (Healthx/Openx) % |
| rm_tpa | NUMERIC | TPA (NAS) % |
| rm_insurance_tax | NUMERIC | Insurance levy % |
| confirmed_quote | NUMERIC | Quoted premium from PDF |
| quoted_members | INT | Member count from quote |
| quote_grand_total | NUMERIC | Grand total from quote |
| member_count | INT | Confirmed census count |
| total_net | NUMERIC | Sum of base premiums |
| total_maternity | NUMERIC | Sum of maternity surcharges |
| total_basmah | NUMERIC | members x AED 37 |
| subtotal | NUMERIC | net + maternity + basmah |
| vat | NUMERIC | subtotal x 5% |
| grand_total | NUMERIC | subtotal + VAT |
| recon_match | BOOLEAN | calc vs quote match (<AED 20) |
| recon_difference | NUMERIC | calc - quote |
| recon_note | TEXT | Free-text note |

### `policy_members` -- individual insured members
| Column | Type | Description |
|--------|------|-------------|
| id | BIGINT PK | Auto-generated |
| policy_id | BIGINT FK | References policies(id) ON DELETE CASCADE |
| member_no | INT | 1-based sequence |
| name | TEXT | Full name |
| dob | DATE | Date of birth |
| gender | CHAR(1) | M or F |
| marital_status | TEXT | Single, Married, Divorced, Widowed |
| relation | TEXT | Employee, Spouse, Child |
| category | CHAR(1) | A, B, C, etc. |
| age_alb | INT | Age Last Birthday at start_date |
| emirate | TEXT | Dubai, Abu Dhabi, Ajman, etc. |
| age_bracket | TEXT | "26-30" |
| base_premium | NUMERIC | Rate from bracket lookup |
| maternity_premium | NUMERIC | Surcharge if eligible |
| gross_loading | NUMERIC | Endorsement surcharge |
| final_premium | NUMERIC | base + maternity + loading |

### `policy_categories` -- rate metadata per category
| Column | Type | Description |
|--------|------|-------------|
| id | BIGINT PK | Auto-generated |
| policy_id | BIGINT FK | References policies(id) |
| category | CHAR(1) | A, B, C |
| maternity_rate | NUMERIC | Per-female surcharge |

### `policy_brackets` -- age bracket rates
| Column | Type | Description |
|--------|------|-------------|
| id | BIGINT PK | Auto-generated |
| category_id | BIGINT FK | References policy_categories(id) |
| label | TEXT | "18-25" |
| age_lo | INT | 18 |
| age_hi | INT | 25 |
| male_rate | NUMERIC | Male premium |
| female_rate | NUMERIC | Female premium |

### `sessions` -- ephemeral workflow state
| Column | Type | Description |
|--------|------|-------------|
| token | TEXT PK | UUID |
| data | JSONB | Serialized session state |
| created_at | TIMESTAMP | Auto-set, 6-hour expiry |

### `rm_targets` -- dashboard targets
| Column | Type | Description |
|--------|------|-------------|
| id | BIGINT PK | Auto-generated |
| rm_name | TEXT UNIQUE | Relationship manager name |
| target_amount | NUMERIC | Monthly revenue target |

---

## 6. API ENDPOINTS

### Upload & Calculation Flow (the main pipeline)

| # | Endpoint | Method | What it does |
|---|----------|--------|-------------|
| 1 | `/api/upload` | POST | Accepts census file + rates PDF, parses both, returns `token` + extracted data |
| 2 | `/api/update_census/<token>` | POST | Accept user-corrected census (editable table) |
| 3 | `/api/compare_census/<token>` | POST | Compare confirmed census vs a second "quoted" census -- find discrepancies |
| 4 | `/api/calculate` | POST | Apply rates to census, calculate all premiums, return `out_token` + reconciliation |
| 5 | `/api/finalize_summary` | POST | Generate Excel workbook, persist policy to DB, return download `prd_token` |
| 6 | `/download/<token>/prd` | GET | Stream Excel file to browser |

### Policy Management (post-issuance)

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/policies` | GET | List policies (paginated, filterable) |
| `/api/policies/<pid>` | GET/PUT/DELETE | Read/update/delete a policy |
| `/api/policies/bulk_delete` | POST | Delete multiple policies |
| `/api/policies/<pid>/members` | GET/POST | List or add members |
| `/api/policies/<pid>/members/<mid>` | PUT/DELETE | Edit or remove a member |
| `/api/policies/<pid>/rates` | PUT | Edit rate table for a policy |
| `/api/policies/<pid>/census` | PUT | Replace entire census (re-upload) |
| `/api/policies/<pid>/export` | GET | Re-generate Excel for existing policy |

### Dashboard

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/dashboard/monthly` | GET | Monthly premium & member trends |
| `/api/dashboard/brokers` | GET | Broker performance ranking |
| `/api/dashboard/rm` | GET | RM revenue & take rate |
| `/api/upsert_rm_target` | POST | Set RM revenue targets |

---

## 7. PREMIUM CALCULATION LOGIC

This is the core business logic that **stays the same** regardless of where rates come from.

### Per-Member Formula
```
1. age = floor((start_date - dob) / 365.25)         # Age Last Birthday
2. bracket = find_bracket(age, category_brackets)     # e.g., "26-30"
3. base_premium = bracket.female if female else bracket.male
4. maternity = maternity_rate if (female AND married AND 18 <= age <= mat_age_max) else 0
5. total = base_premium + maternity
```

### Maternity Eligibility
| Condition | Value |
|-----------|-------|
| Gender | Female |
| Marital status | Married |
| Minimum age | 18 |
| Maximum age (Dubai) | **45** |
| Maximum age (Abu Dhabi) | 50 |

**Known issue:** The display in the PDF output (PRD template) sometimes shows 46 as the maternity limit -- this is inherited from the old MGH reference template, not from the calculation logic. The backend correctly uses 45 for Dubai.

### Totals
```
total_net       = sum(base_premium for all members)
total_maternity = sum(maternity for all members)
total_basmah    = member_count x 37.0 AED
subtotal        = total_net + total_maternity + total_basmah
vat             = subtotal x 0.05
grand_total     = subtotal + vat
```

### Commission Structure
```
Gross Premium = base + maternity (per member)
Net Premium   = Gross x (1 - total_fees%)

Where total_fees% = broker% + insurer% + admin% + tpa% + levy%

LSB Mode: broker capped at 5%, excess split 50/50 between insurer and admin
```

### Hardcoded Constants
```python
BASMAH_FEE       = 37.0    # AED per member (DHA fee)
VAT_RATE         = 0.05    # 5%
MAT_AGE_MIN      = 18
MAT_AGE_MAX_DXB  = 45      # Dubai
MAT_AGE_MAX_AUH  = 50      # Abu Dhabi
```

---

## 8. THE INTEGRATION POINT -- WHERE THE QUOTE TOOL CONNECTS

### What to Replace

The entire PDF rate extraction pipeline:

| Function | Lines | What it does | Replace with |
|----------|-------|-------------|-------------|
| `parse_rates_pdf()` | 303-460 | Sends PDF pages to Claude Vision, parses JSON response | API call to quote tool |
| `parse_rates_xlsx()` | 462-560 | Fallback: extracts rates from Excel rate file | Not needed |
| PDF page rendering | 265-300 | Renders PDF pages to images for Vision API | Not needed |
| Rate verification UI | index.html Step 2 | Manual rate correction panel | Optional: read-only display |

### What the Quote Tool Must Provide

The downstream calculation expects this exact data structure:

```json
{
  "categories": {
    "A": {
      "brackets": [
        {"label": "0-10",  "age_lo": 0,  "age_hi": 10, "male": 8368, "female": 6844},
        {"label": "11-17", "age_lo": 11, "age_hi": 17, "male": 6169, "female": 6260},
        {"label": "18-25", "age_lo": 18, "age_hi": 25, "male": 7059, "female": 7763},
        {"label": "26-30", "age_lo": 26, "age_hi": 30, "male": 7565, "female": 9337},
        {"label": "31-35", "age_lo": 31, "age_hi": 35, "male": 8869, "female": 10906},
        {"label": "36-40", "age_lo": 36, "age_hi": 40, "male": 10655, "female": 13296},
        {"label": "41-45", "age_lo": 41, "age_hi": 45, "male": 11123, "female": 15078},
        {"label": "46-50", "age_lo": 46, "age_hi": 50, "male": 14158, "female": 18552},
        {"label": "51-55", "age_lo": 51, "age_hi": 55, "male": 18545, "female": 24041},
        {"label": "56-60", "age_lo": 56, "age_hi": 60, "male": 24789, "female": 29128},
        {"label": "61-65", "age_lo": 61, "age_hi": 65, "male": 30211, "female": 34359},
        {"label": "66-99", "age_lo": 66, "age_hi": 99, "male": 53294, "female": 60763}
      ],
      "maternity_rate": 5697
    },
    "B": {
      "brackets": [
        {"label": "0-10", "age_lo": 0, "age_hi": 10, "male": 6329, "female": 5121},
        ...
      ],
      "maternity_rate": 4793
    }
  },
  "quote_totals": {
    "total_premium": 340102,
    "members": 30,
    "grand_total": 387673
  }
}
```

**Key fields:**
- `brackets[].male` / `brackets[].female` -- GROSS annual premium in AED (not net)
- `maternity_rate` -- per-eligible-female surcharge in AED (not total across all females)
- `quote_totals.total_premium` -- the indicative total premium (net, before basmah/VAT)
- `quote_totals.members` -- expected member count
- `quote_totals.grand_total` -- total including basmah + VAT + loading

### Optional: Dependent Rates (4-column format)

Some plans have separate Employee vs Dependent rates. If the quote tool provides them:

```json
{
  "label": "26-30",
  "age_lo": 26, "age_hi": 30,
  "male": 7565,           // Employee Male
  "female": 9337,         // Employee Female
  "dep_male": 7200,       // Dependent Male (Spouse/Child)
  "dep_female": 8900      // Dependent Female (Spouse/Child)
}
```

If `dep_male`/`dep_female` are absent, the tool uses `male`/`female` for all members regardless of relation.

---

## 9. SUGGESTED INTEGRATION APPROACH

### Option A: New API Endpoint (Recommended)

Create a new endpoint that accepts rates from the quote tool directly:

```
POST /api/upload_from_quote
Body: {
  "plan": "Healthx",
  "start_date": "2026-04-08",
  "company_name": "Latin Textiles & Trade FZCO",
  "broker": "Lifecare International",
  "categories": { ... },        // rate data from quote tool
  "quote_totals": { ... },      // quoted premiums
  "census_file": <multipart>    // census still uploaded separately
}
```

This skips Steps 1-2 of the current UI flow (PDF upload + rate verification) and jumps straight to calculation.

### Option B: Modify `/api/upload` to Accept Pre-Extracted Rates

Add a flag to the existing upload:

```
POST /api/upload
Body: {
  ...existing fields...,
  "pre_extracted_rates": { ... }  // if present, skip PDF parsing
}
```

### Option C: Direct Database Write

If the quote tool manages its own policies, it could write directly to the Supabase tables and only call the Premium Summary tool for Excel generation:

```
POST /api/policies/<pid>/export
```

---

## 10. WHAT TO KEEP, WHAT TO REMOVE

### KEEP (core value)
- Census parsing (`parse_census`) -- robust multi-format parser
- Premium calculation (`calculate_premiums`) -- the business logic
- Excel generation (`make_combined_excel`) -- branded output
- Census comparison (`compare_censuses`, `api_compare_census`) -- reconciliation
- Database persistence (`save_policy`, member CRUD)
- Dashboard & analytics
- Commission structure (HSB/LSB)
- Reconciliation sheet (newly added)

### REMOVE (replaced by quote tool)
- `parse_rates_pdf()` -- Claude Vision extraction
- `parse_rates_xlsx()` -- Excel rate extraction fallback
- PDF page rendering functions
- Rate verification UI panel (Step 2 in index.html)
- `ANTHROPIC_API_KEY` environment variable
- `anthropic` and `PyMuPDF` dependencies

### MODIFY
- `/api/upload` -- accept rates as JSON instead of PDF file
- Step 2 UI -- show rates as read-only (confirmation only) or skip entirely
- `vercel.json` -- can reduce `maxDuration` from 300s to 60s (no more PDF processing)

---

## 11. FILE MAP -- WHERE TO FIND THINGS

```
app.py (4,342 lines)
├── Lines 1-100:      Constants, imports, Flask app setup
├── Lines 100-260:    Utility functions (age calc, rate lookup, member matching)
├── Lines 260-625:    PDF/XLSX rate extraction (REMOVE for integration)
├── Lines 625-825:    PDF page rendering, vision API (REMOVE)
├── Lines 825-975:    Census parsing (KEEP)
├── Lines 975-1130:   Rate lookup, bracket matching (KEEP)
├── Lines 1130-1175:  calculate_premiums() (KEEP -- core logic)
├── Lines 1175-1240:  Excel styles, themes, LSB calculator (KEEP)
├── Lines 1240-2000:  make_combined_excel() (KEEP -- Excel generation)
├── Lines 2000-2250:  Database: save_policy, member CRUD (KEEP)
├── Lines 2250-2460:  Census comparison & reconciliation (KEEP)
├── Lines 2460-2670:  API routes: calculate, finalize, download (KEEP)
├── Lines 2670-3100:  Policy management routes (KEEP)
├── Lines 3100-3400:  Member CRUD routes, rate recalculation (KEEP)
├── Lines 3400-4342:  Dashboard routes, settings, brand assets (KEEP)
```

---

## 12. LOCAL DEVELOPMENT

### Setup
```bash
git clone https://github.com/jasperh0917/premiumsummary.git
cd premiumsummary
pip install -r requirements.txt

# Create .env with:
# SUPABASE_URL=...
# SUPABASE_ANON_KEY=...
# SUPABASE_SERVICE_ROLE_KEY=...
# ANTHROPIC_API_KEY=...  (only needed until quote tool integration)

python app.py
# Runs on http://localhost:5000
```

### Deployment
```bash
git push origin main
# Vercel auto-deploys within ~60 seconds
```

### Testing with Sample Data
- Sample census: `Sample files/Input/UPDATED QIC MEMBER REGISTER-EMVE trading.xlsx`
- Sample rates PDF: `Sample files/Input/QHX-2603542-EMVE Trading-25.03.26.pdf`
- Expected output: `Sample files/Output/` directory

---

## 13. KNOWN ISSUES & EDGE CASES

| Issue | Status | Notes |
|-------|--------|-------|
| Maternity age limit display shows 46 in some Excel outputs | Cosmetic | Backend uses 45 correctly; the "46" comes from the PRD template's "Maternity Age Limit" column header which shows the cutoff+1 |
| Session expiry (6 hours) | By design | Users must complete workflow within 6 hours |
| Census date parsing | Fragile | Supports DD-MMM-YYYY, YYYY-MM-DD, DD/MM/YYYY; other formats may fail silently |
| Name matching in census comparison | Case-insensitive + stripped | May fail if names have different Unicode characters or transliterations |
| Dependent rates (4-column) | Partial | Only used if `dep_male`/`dep_female` keys exist in bracket data |
| LSB commission cap | Working | Broker capped at 5%, excess split 50/50 |
| Large census (500+ members) | Slow | Excel generation takes 10-15 seconds |

---

## 14. CONTACTS & ACCESS

| Resource | Location |
|----------|----------|
| GitHub repo | `https://github.com/jasperh0917/premiumsummary.git` |
| Vercel project | Dashboard > premiumsummary (linked to GitHub) |
| Supabase project | `https://supabase.com/dashboard/project/bktbqaxpcozpaasbkpud` |
| Production URL | Set in Vercel dashboard |
