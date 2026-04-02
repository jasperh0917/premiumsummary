import os, io, uuid, re, base64, json
from datetime import datetime, date
from flask import Flask, request, jsonify, render_template, send_file
import fitz  # PyMuPDF
import pandas as pd
import openpyxl
import anthropic

# ── Load .env file if present ──────────────────────────────────────────────────
_ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.isfile(_ENV_PATH):
    with open(_ENV_PATH) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith('#') and '=' in _line:
                _k, _v = _line.split('=', 1)
                _k = _k.strip(); _v = _v.strip().strip('"').strip("'")
                # Set if the env var is missing or empty, and the file value is non-empty / non-placeholder
                if _v and _v != 'your-api-key-here' and not os.environ.get(_k):
                    os.environ[_k] = _v
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_ROOT = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__,
            root_path=_ROOT,
            template_folder=os.path.join(_ROOT, 'templates'),
            instance_path=os.path.join(_ROOT, 'instance'))
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

_store = {}  # In-memory: token -> data

# ── Constants ────────────────────────────────────────────────────────────────
BASMAH_FEE       = 37.0
VAT_RATE         = 0.05
MAT_AGE_MIN      = 18
MAT_AGE_MAX_DXB  = 45
MAT_AGE_MAX_AUH  = 50

DEFAULT_BRACKETS = [
    (0,10,"0-10"), (11,17,"11-17"), (18,25,"18-25"),
    (26,30,"26-30"), (31,35,"31-35"), (36,40,"36-40"),
    (41,45,"41-45"), (46,50,"46-50"), (51,55,"51-55"),
    (56,59,"56-59"), (60,64,"60-64"), (65,99,"65-99"),
]

DOH_BRACKETS = [
    (0, 17, "0-17"), (18, 40, "18-40"),
    (41, 60, "41-60"), (61, 99, "61-99"),
]

# ── PDF Utilities ─────────────────────────────────────────────────────────────
def pdf_page_image(pdf_bytes, page_idx, scale=2.0):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if page_idx < 0 or page_idx >= len(doc):
        page_idx = len(doc) - 1
    pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(scale, scale))
    return base64.b64encode(pix.tobytes("png")).decode(), len(doc)

def try_extract_text(pdf_bytes):
    try:
        import pdfplumber
        pages = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                pages.append(pg.extract_text() or "")
        return pages
    except Exception:
        return []

def find_rate_page_idx(page_texts, pdf_bytes):
    keywords = ["PREMIUM SUMMARY", "Age Range", "Category A", "PREMIUM OVERVIEW"]
    for i, text in enumerate(page_texts):
        if text and any(kw.lower() in text.lower() for kw in keywords):
            return i
    # Try pytesseract OCR fallback
    try:
        import pytesseract
        from PIL import Image
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i in range(len(doc)):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img)
            if any(kw.lower() in text.lower() for kw in keywords):
                return i
    except Exception:
        pass
    # Heuristic: rate table is often in the last quarter of the document
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = len(doc)
    return max(0, n - 4)  # Try 4th-from-last page

def try_parse_rates_from_text(text):
    rates = []
    pattern = re.compile(
        r'\b(\d{1,2})\s*[-–]\s*(\d{2,3})\b\s+([\d,]{4,})\s*(?:\d+\s+)?([\d,]{4,})?',
        re.MULTILINE
    )
    seen = set()
    for m in pattern.finditer(text):
        lo, hi = int(m.group(1)), int(m.group(2))
        if lo > hi or hi > 120:
            continue
        male_rate = int(m.group(3).replace(',', ''))
        if male_rate < 100:
            continue
        fem_str = m.group(4) or m.group(3)
        fem_rate = int(fem_str.replace(',', ''))
        if fem_rate < 100:
            fem_rate = male_rate
        key = (lo, hi)
        if key not in seen:
            seen.add(key)
            rates.append({'age_lo': lo, 'age_hi': hi, 'label': f"{lo}-{hi}",
                          'male': male_rate, 'female': fem_rate})
    return sorted(rates, key=lambda x: x['age_lo'])

def try_parse_maternity_rate(text):
    m = re.search(r'Additional\s+Maternity\s+Premium\s+(?:AED\s+)?([\d,]+)', text, re.I)
    if m:
        v = int(m.group(1).replace(',', ''))
        return v if v > 100 else 0
    return 0

def extract_rates_with_claude_vision(pdf_bytes, page_idx):
    """Use Claude Vision API to extract premium rates from an image-based PDF page."""
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        print("[Vision] ANTHROPIC_API_KEY not set — skipping Claude Vision extraction")
        return None, {}

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_idx < 0 or page_idx >= len(doc):
            page_idx = len(doc) - 1
        pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(3, 3))
        img_b64 = base64.standard_b64encode(pix.tobytes("png")).decode()
        doc.close()
    except Exception as e:
        print(f"[Vision] Failed to render PDF page: {e}")
        return None, {}

    prompt = """You are extracting premium rate tables from an insurance quote PDF page.

Return ONLY a valid JSON object — no markdown, no explanation — with this exact structure:
{
  "categories": {
    "<letter>": {
      "brackets": [
        {"label": "<lo>-<hi>", "age_lo": <int>, "age_hi": <int>, "male": <float>, "female": <float>}
      ],
      "maternity_rate": <float or 0>
    }
  },
  "quote_totals": {
    "total_premium": <float or 0>,
    "members": <int or 0>,
    "grand_total": <float or 0>
  }
}

Rules:
- Extract ALL categories visible (A, B, C, etc.). If only one category, use "A".
- "male" and "female" are annual premiums in AED. If the table shows one rate column, use the same value for both.
- "maternity_rate" is the additional annual maternity premium per eligible female (look for "Additional Maternity Premium" or similar). Use 0 if not shown.
- "total_premium" is the net premium before fees (look for "TOTAL PREMIUM"). Use 0 if not shown.
- "members" is the total member count from the summary. Use 0 if not shown.
- "grand_total" includes BASMAH + VAT. Use 0 if not shown.
- Age bracket labels must match the lo-hi format exactly, e.g. "0-10", "11-17", "18-25"."""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=2048,
            system="You are a precise structured data extractor for insurance documents. Return only valid JSON.",
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_b64,
                        },
                    },
                    {"type": "text", "text": prompt}
                ],
            }]
        )
        raw = response.content[0].text.strip()
        # Strip markdown fences if Claude wraps the JSON
        raw = re.sub(r'^```(?:json)?\s*', '', raw)
        raw = re.sub(r'\s*```$', '', raw)
        data = json.loads(raw)
        cats = data.get('categories', {})
        totals = data.get('quote_totals', {})
        print(f"[Vision] Extracted {len(cats)} categorie(s): {list(cats.keys())}")
        return cats, totals
    except Exception as e:
        print(f"[Vision] Claude Vision extraction failed: {e}")
        return None, {}

def try_parse_quote_totals(text):
    totals = {}
    m = re.search(r'TOTAL\s+PREMIUM.*?(\d+)\s+members?.*?AED\s+([\d,]+)', text, re.I | re.DOTALL)
    if m:
        totals['members']       = int(m.group(1))
        totals['total_premium'] = int(m.group(2).replace(',', ''))
    m = re.search(r'BASMAH.*?AED\s+([\d,]+)', text, re.I | re.DOTALL)
    if m:
        totals['basmah'] = int(m.group(1).replace(',', ''))
    m = re.search(r'VALUE\s+ADDED\s+TAX.*?AED\s+([\d,]+)', text, re.I | re.DOTALL)
    if m:
        totals['vat'] = int(m.group(1).replace(',', ''))
    all_aed = re.findall(r'AED\s+([\d,]+)', text)
    if all_aed:
        amounts = sorted([int(a.replace(',', '')) for a in all_aed], reverse=True)
        totals['grand_total'] = amounts[0]
    return totals

# ── HealthXclusive Tool Parser ────────────────────────────────────────────────
def _safe_num(v, default=0.0):
    """Return float from a cell value, or default if None / error string."""
    try:
        if v is None:
            return default
        if isinstance(v, str) and ('#' in v or not v.strip() or v.strip().upper() == 'N/A'):
            return default
        return float(v)
    except Exception:
        return default

def parse_healthxclusive_tool(excel_bytes):
    """
    Parse the HealthXclusive Tool Excel's 'Premium Summary' sheet.
    Returns a structured dict with policy info, fees and per-category rate brackets.
    """
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        xl = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)

    if 'Premium Summary' not in xl.sheetnames:
        raise ValueError("'Premium Summary' sheet not found. Please upload the HealthXclusive Tool Excel.")

    ws = xl['Premium Summary']
    rows = [[cell.value for cell in row] for row in ws.iter_rows()]

    result = {
        'company_name': '', 'product': '', 'start_date': '', 'end_date': '',
        'underwriter': '', 'validated_by': '', 'policy_type': 'New Business',
        'payment_frequency': 'Annual', 'endorsement_frequency': 'Quarterly',
        'confirmed_quote': 0.0, 'members': 0,
        'flat_fees': {'basmah': 19, 'hcv': 18, 'trudoc': 12, 'slash': 0},
        'fees': {
            'broker':  {'hsb': 0.0, 'lsb': 0.0},
            'nas':     {'hsb': 0.0, 'lsb': 0.0},
            'qic':     {'hsb': 0.0, 'lsb': 0.0},
            'healthx': {'hsb': 0.0, 'lsb': 0.0},
            'levy':    {'hsb': 0.0, 'lsb': 0.0},
            'total':   {'hsb': 0.0, 'lsb': 0.0},
        },
        'categories': {}
    }

    for row in rows:
        if not any(v is not None for v in row):
            continue

        def g(i):
            return row[i] if i < len(row) else None

        r0  = str(g(0) or '').strip()
        r0l = r0.lower()

        # ── Policy info (col A = label, col B = value) ──
        if r0l == 'product':
            result['product'] = str(g(1) or '').strip()
        elif 'policy holder' in r0l:
            result['company_name'] = str(g(1) or '').strip()
        elif 'policy start date' in r0l:
            v = g(1)
            result['start_date'] = v.strftime('%Y-%m-%d') if isinstance(v, datetime) else str(v or '')
        elif 'policy end date' in r0l:
            v = g(1)
            result['end_date'] = v.strftime('%Y-%m-%d') if isinstance(v, datetime) else str(v or '')
        elif 'policy type' in r0l:
            result['policy_type'] = str(g(1) or '').strip()
        elif 'inception premium' in r0l:
            result['payment_frequency'] = str(g(1) or '').strip()
        elif 'endorsement frequency' in r0l:
            result['endorsement_frequency'] = str(g(1) or '').strip()
        elif r0l == 'underwriter':
            result['underwriter'] = str(g(1) or '').strip()
        elif 'validated by' in r0l:
            result['validated_by'] = str(g(1) or '').strip()

        # ── Fees section (label in col N=13, HSB in col O=14, LSB in col Q=16) ──
        r13 = str(g(13) or '').strip().lower()

        if 'dha basmah' in r13:
            result['flat_fees']['basmah'] = int(_safe_num(g(14), 19))
        elif 'dha hcv' in r13:
            result['flat_fees']['hcv'] = int(_safe_num(g(14), 18))
        elif 'trudoc' in r13:
            result['flat_fees']['trudoc'] = int(_safe_num(g(14), 12))
        elif 'slash' in r13:
            result['flat_fees']['slash'] = int(_safe_num(g(14), 0))
        elif r13 == 'broker':
            result['fees']['broker'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}
            cq = g(20)
            if cq is not None and _safe_num(cq) > 0:
                result['confirmed_quote'] = _safe_num(cq)
        elif 'nas' in r13 or 'tpa' in r13:
            result['fees']['nas'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}
            mb = g(20)
            if mb is not None and _safe_num(mb) > 0:
                result['members'] = int(_safe_num(mb))
        elif 'qic' in r13:
            result['fees']['qic'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}
        elif 'healthx' in r13 or 'health x' in r13:
            result['fees']['healthx'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}
        elif 'insurance levy' in r13:
            result['fees']['levy'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}
        elif 'total fees' in r13:
            result['fees']['total'] = {'hsb': _safe_num(g(14)), 'lsb': _safe_num(g(16))}

        # ── Rate rows: 'Cat A (EMPDEP)', 'Cat B (EMPDEP)', maternity rows ──
        if r0.startswith('Cat ') and len(r0) >= 5:
            cat_letter = r0[4].upper()
            is_mat = 'MAT' in r0.upper() and 'EMPDEPMAT' in r0.upper() or \
                     (str(g(15) or '').lower().strip() == 'additional maternity premium')
            is_dep = 'EMPDEP' in r0 and not is_mat

            if cat_letter not in result['categories']:
                result['categories'][cat_letter] = {'brackets': [], 'maternity_rate': 0}

            age_band     = str(g(1) or '').strip()
            male_gross   = g(6)
            female_gross = g(7)

            if is_dep and age_band:
                m_ok = isinstance(male_gross, (int, float))
                f_ok = isinstance(female_gross, (int, float))
                if m_ok or f_ok:
                    parts = age_band.split('-')
                    try:
                        lo = int(parts[0])
                        hi = int(parts[1]) if len(parts) > 1 else 99
                    except Exception:
                        continue
                    result['categories'][cat_letter]['brackets'].append({
                        'label':  age_band,
                        'age_lo': lo,
                        'age_hi': hi,
                        'male':   round(float(male_gross), 2) if m_ok else 0.0,
                        'female': round(float(female_gross), 2) if f_ok else 0.0,
                    })
            elif is_mat:
                if isinstance(female_gross, (int, float)):
                    result['categories'][cat_letter]['maternity_rate'] = round(float(female_gross), 2)

    # Remove categories with no valid (non-zero) rate brackets
    result['categories'] = {
        k: v for k, v in result['categories'].items()
        if v['brackets'] and any(b['male'] > 0 or b['female'] > 0 for b in v['brackets'])
    }

    return result

# ── Census Parsing ────────────────────────────────────────────────────────────
def detect_header_row(rows):
    search = {'dob', 'date of birth', 'gender', 'relation', 'category', 'marital'}
    for i, row in enumerate(rows):
        text = ' '.join(str(v).lower() for v in row if v)
        if sum(1 for t in search if t in text) >= 2:
            return i
    return 0

def detect_col_map(headers):
    h = [str(v).lower().strip() if v else '' for v in headers]
    col_map = {}

    for i, c in enumerate(h):
        if 'full name' in c or 'full_name' in c:
            col_map['full_name'] = i; break
    if 'full_name' not in col_map:
        for i, c in enumerate(h):
            if 'first' in c and 'name' in c:
                col_map['first_name'] = i; break
        for i, c in enumerate(h):
            if ('last' in c or 'family' in c or 'sur' in c) and 'name' in c:
                col_map['last_name'] = i; break
        if 'first_name' not in col_map:
            for i, c in enumerate(h):
                if 'name' in c and i not in col_map.values():
                    col_map['full_name'] = i; break

    for i, c in enumerate(h):
        if 'dob' in c or 'birth' in c:
            col_map['dob'] = i; break
    for i, c in enumerate(h):
        if 'gender' in c or c == 'sex':
            col_map['gender'] = i; break
    for i, c in enumerate(h):
        if 'marital' in c:
            col_map['marital'] = i; break
    for i, c in enumerate(h):
        if 'relation' in c:
            col_map['relation'] = i; break
    for i, c in enumerate(h):
        if c == 'category' or (c.startswith('cat') and len(c) <= 8):
            col_map['category'] = i; break
    for i, c in enumerate(h):
        if 'visa' in c and 'issuance' in c:
            col_map['emirate'] = i; break
        elif 'emirate' in c and ('visa' in c or 'issu' in c):
            col_map['emirate'] = i; break
    return col_map

def calculate_alb(dob, start_date):
    if isinstance(dob, datetime):
        dob = dob.date()
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
    elif isinstance(start_date, datetime):
        start_date = start_date.date()
    age = start_date.year - dob.year
    if (start_date.month, start_date.day) < (dob.month, dob.day):
        age -= 1
    return max(0, age)

def calculate_anb(dob, start_date):
    """Age Next Birthday = Age Last Birthday + 1 (used for Healthx plans)."""
    return calculate_alb(dob, start_date) + 1

def parse_census(file_bytes, filename, start_date_str, age_method='alb'):
    if filename.lower().endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
        all_rows  = [list(df.columns)] + df.values.tolist()
        header_idx = 0
        data_start = 1
    else:
        xl = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_name = next((s for s in xl.sheetnames if 'INCEP' in s.upper()), xl.sheetnames[0])
        ws = xl[sheet_name]
        all_rows  = [[cell.value for cell in row] for row in ws.iter_rows()]
        header_idx = detect_header_row(all_rows)
        data_start = header_idx + 2  # skip notes row

    col_map = detect_col_map(all_rows[header_idx])
    if 'dob' not in col_map:
        raise ValueError("Cannot find Date of Birth column in census file")
    if 'emirate' not in col_map:
        raise ValueError("Cannot find 'Emirates of Visa Issuance' column in census file")

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    members = []
    blank_emirate_rows = []

    for row_vals in all_rows[data_start:]:
        dob_raw = row_vals[col_map['dob']] if col_map['dob'] < len(row_vals) else None
        if dob_raw is None:
            continue
        if isinstance(dob_raw, str):
            if not dob_raw.strip() or 'passport' in dob_raw.lower() or dob_raw.startswith('('):
                continue

        try:
            if isinstance(dob_raw, (datetime, date)):
                dob = dob_raw.date() if isinstance(dob_raw, datetime) else dob_raw
            else:
                dob = pd.to_datetime(str(dob_raw)).date()
        except Exception:
            continue

        def get_val(key, default=''):
            idx = col_map.get(key)
            if idx is None or idx >= len(row_vals):
                return default
            v = row_vals[idx]
            return str(v).strip() if v is not None else default

        # Build name
        name = ''
        if 'full_name' in col_map:
            name = get_val('full_name')
            if name.startswith('=') or not name or name == 'nan':
                name = ''

        if not name:
            # QIC format: columns B(1), C(2), D(3) = first, second, last name
            parts = []
            for idx in [1, 2, 3]:
                if idx < len(row_vals) and row_vals[idx]:
                    p = str(row_vals[idx]).strip()
                    if p and p != 'nan' and not p.startswith('=') and not p.startswith('('):
                        parts.append(p)
            name = ' '.join(parts)

        if not name or name == 'nan':
            # Try first_name + last_name col_map keys
            fn = get_val('first_name')
            ln = get_val('last_name')
            parts = [p for p in [fn, ln] if p and p != 'nan']
            name = ' '.join(parts)

        if not name:
            continue

        gender        = get_val('gender', 'Unknown')
        if gender in ('nan', ''):
            gender = 'Unknown'

        marital       = get_val('marital', 'Unknown')
        if marital in ('nan', ''):
            marital = 'Unknown'

        relation      = get_val('relation', 'Employee')
        if relation in ('nan', ''):
            relation = 'Employee'

        category      = get_val('category', 'A').upper()
        if category in ('NAN', ''):
            category = 'A'

        emirate_raw = get_val('emirate', '')
        if not emirate_raw or emirate_raw.lower() in ('nan', 'none', ''):
            blank_emirate_rows.append(name or f"Row {len(members) + data_start + 1}")
            continue

        emirate = emirate_raw.strip()

        age_fn = calculate_anb if age_method == 'anb' else calculate_alb
        members.append({
            'name':           name,
            'dob':            dob.strftime('%d-%b-%Y'),
            'gender':         gender,
            'marital_status': marital,
            'relation':       relation,
            'category':       category,
            'age_alb':        age_fn(dob, start_date),
            'emirate':        emirate,
        })

    if blank_emirate_rows:
        names_list = ', '.join(blank_emirate_rows[:5])
        extra = f' and {len(blank_emirate_rows) - 5} more' if len(blank_emirate_rows) > 5 else ''
        raise ValueError(
            f"Emirates of Visa Issuance is blank for {len(blank_emirate_rows)} member(s): "
            f"{names_list}{extra}. Please complete the census before proceeding."
        )

    return members

# ── Rate Lookup ───────────────────────────────────────────────────────────────
def find_bracket(age, brackets):
    for b in brackets:
        if b['age_lo'] <= age <= b['age_hi']:
            return b
    return None

def get_member_rate(member, categories_data):
    cat = member['category'].upper()
    if cat not in categories_data and categories_data:
        cat = sorted(categories_data.keys())[0]
    if cat not in categories_data:
        return 0, None, f"No rate data for category {member['category']}"
    brackets = categories_data[cat]['brackets']
    bracket  = find_bracket(member['age_alb'], brackets)
    if not bracket:
        return 0, None, f"No bracket for age {member['age_alb']}"
    rate = bracket.get('female' if member['gender'].lower().startswith('f') else 'male', 0)
    return rate, bracket['label'], None

# ── Calculation Engine ────────────────────────────────────────────────────────
def calculate_premiums(members, categories_data):
    results = []
    for i, m in enumerate(members):
        rate, bracket_label, error = get_member_rate(m, categories_data)
        cat      = m['category'].upper()
        mat_rate = categories_data.get(cat, {}).get('maternity_rate', 0)
        maternity = 0
        emirate   = m.get('emirate', 'Dubai')
        mat_age_max = MAT_AGE_MAX_AUH if 'abu dhabi' in emirate.lower() else MAT_AGE_MAX_DXB
        if (m['gender'].lower().startswith('f')
                and m['marital_status'].lower().startswith('m')
                and MAT_AGE_MIN <= m['age_alb'] <= mat_age_max):
            maternity = float(mat_rate or 0)

        results.append({
            **m,
            'no':               i + 1,
            'base_premium':     float(rate),
            'age_bracket':      bracket_label or 'N/A',
            'maternity_premium': maternity,
            'mat_age_max':      mat_age_max,
            'basmah_fee':       BASMAH_FEE,
            'total_excl_vat':   float(rate) + maternity,
            'error':            error or '',
        })

    net       = sum(r['base_premium']     for r in results)
    mat_total = sum(r['maternity_premium'] for r in results)
    bas_total = sum(r['basmah_fee']        for r in results)
    subtotal  = net + mat_total + bas_total
    vat       = subtotal * VAT_RATE

    return results, {
        'total_net':       net,
        'total_maternity': mat_total,
        'total_basmah':    bas_total,
        'subtotal':        subtotal,
        'vat':             vat,
        'grand_total':     subtotal + vat,
        'member_count':    len(results),
    }

# ── Excel Styles ─────────────────────────────────────────────────────────────
NAVY   = "003780"; ORANGE = "fb9b35"; SKY    = "35c5fc"
VIOLET = "8431cb"; WHITE  = "FFFFFF"; DARK   = "1a2332"
LGRAY  = "EEF2FF"; MGRAY  = "F0F4FF"; PINK   = "f1517b"
GREEN_BG = "D4EDDA"; GREEN_FG = "155724"
RED_BG   = "F8D7DA"; RED_FG   = "721c24"
ORANGE_BG = "FFF3E0"

def _f(name="Inter", bold=False, size=10, color=DARK):
    return Font(name=name, bold=bold, size=size, color=color)
def _fill(c):
    return PatternFill("solid", fgColor=c)
def _al(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _border(c="D1D9E6"):
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)

# ── LSB Rate Calculator ───────────────────────────────────────────────────────
def calc_lsb_rates(form_data, has_lsb):
    broker  = float(form_data.get('rm_broker',        0) or 0)
    insurer = float(form_data.get('rm_insurer',       0) or 0)
    admin   = float(form_data.get('rm_wellx',         0) or 0)
    nas     = float(form_data.get('rm_tpa',           0) or 0)
    levy    = float(form_data.get('rm_insurance_tax', 0) or 0)

    if not has_lsb:
        return broker, insurer, admin, nas, levy, broker, insurer, admin, nas, levy

    broker_lsb  = min(broker, 5.0)
    excess      = max(0.0, broker - 5.0)
    insurer_lsb = insurer + excess / 2.0
    admin_lsb   = admin   + excess / 2.0
    return broker, insurer, admin, nas, levy, broker_lsb, insurer_lsb, admin_lsb, nas, levy


# ── Plan Themes ───────────────────────────────────────────────────────────────
THEMES = {
    'healthx':        {'primary': '1A5E20', 'light': 'E8F5E9', 'mid': '2E7D32'},
    'healthxclusive': {'primary': '7B1818', 'light': 'FDECEA', 'mid': 'B71C1C'},
    'openx':          {'primary': '003780', 'light': 'EEF2FF', 'mid': '1565C0'},
}

def get_theme(plan: str) -> dict:
    return THEMES.get((plan or '').lower().replace(' ', ''), THEMES['openx'])


# ── Mismatch Analysis ─────────────────────────────────────────────────────────
def build_mismatch_analysis(members_data, q_premium, c_premium, q_members, net_loading_total=0.0):
    causes = []
    if q_members and abs(len(members_data) - q_members) > 0:
        causes.append(f"Member count: census {len(members_data)} vs quote {q_members}")
    error_members = [m for m in members_data if m.get('error', '')]
    if error_members:
        causes.append(f"{len(error_members)} member(s) have no matching rate bracket")
    if net_loading_total > 0:
        causes.append(f"Net loading AED {net_loading_total:,.2f} included in final premium")
    if not causes:
        causes.append("Review age brackets, maternity eligibility and category assignments")
    return causes


# ── Combined Excel Generator ─────────────────────────────────────────────────
def make_combined_excel(form_data, members_data, verified_rates, maternity_rates,
                        loading_members, has_lsb, totals, quote_totals,
                        quoted_totals_calc=None, quoted_member_count=None):
    from collections import defaultdict as _dd

    plan        = form_data.get('plan', '')
    plan_type   = form_data.get('plan_type', '')
    company     = form_data.get('company_name', '')
    broker_name = form_data.get('broker', '')
    underwriter = form_data.get('underwriter', '')
    start_date  = form_data.get('start_date', '')
    inception   = form_data.get('inception_payment', '')
    endorse     = form_data.get('endorsement_freq', '')

    # Insurer / admin label based on plan
    is_openx     = plan.lower() == 'openx'
    insurer_name = 'DNI' if is_openx else 'QIC'
    admin_name   = 'Openx' if is_openx else 'Healthx'

    (broker_h, insurer_h, admin_h, nas_h, levy_h,
     broker_l, insurer_l, admin_l, nas_l, levy_l) = calc_lsb_rates(form_data, has_lsb)

    theme = get_theme(plan)
    PRI   = theme['primary']
    LGT   = theme['light']
    MID   = theme['mid']

    # Loading totals
    total_fees_frac  = (broker_h + insurer_h + admin_h + nas_h + levy_h) / 100.0
    total_gross_load = sum(float(lm.get('gross_loading', 0) or 0) for lm in loading_members)
    total_net_load   = total_gross_load * (1.0 - total_fees_frac)
    loading_vat      = total_net_load * VAT_RATE
    loading_grand    = total_net_load + loading_vat
    final_grand_total = totals['grand_total'] + loading_grand

    # Reconciliation
    is_healthx = plan.lower() == 'healthx'
    q_premium  = float(quote_totals.get('total_premium', 0) or 0)
    # Confirmed census premium (used for Sheet 2 totals)
    conf_premium = totals['total_net'] + totals['total_maternity']
    # For reconciliation: Healthx uses quoted census vs PDF quote; others use confirmed census
    if is_healthx and quoted_totals_calc is not None:
        c_premium = quoted_totals_calc['total_net'] + quoted_totals_calc['total_maternity']
    else:
        c_premium = conf_premium
    diff      = c_premium + total_net_load - q_premium
    is_match  = abs(diff) < 20.0
    q_members = int(quote_totals.get('members', 0) or 0)

    # ── Workbook ───────────────────────────────────────────────────────────────
    wb  = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = 'Premium Summary'
    ws2 = wb.create_sheet('Premium per Member')

    def thin_border(c='D1D9E6'):
        s = Side(style='thin', color=c)
        return Border(left=s, right=s, top=s, bottom=s)

    def apply_borders(ws):
        tb = thin_border()
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                if cell.value is not None:
                    cell.border = tb

    def pct_val(v):
        return round(v / 100.0, 6)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 1: PREMIUM SUMMARY (PRD FORMAT)
    # ══════════════════════════════════════════════════════════════════════════
    ws = ws1

    def cv(r, c, val, bold=False, color=DARK, size=10, name='Inter',
           halign='left', fill_color=None, num_fmt=None, italic=False):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font      = Font(name=name, bold=bold, size=size, color=color, italic=italic)
        cell.alignment = Alignment(horizontal=halign, vertical='center')
        if fill_color:
            cell.fill = PatternFill('solid', fgColor=fill_color)
        if num_fmt:
            cell.number_format = num_fmt
        return cell

    # Column widths for Sheet 1
    col_widths = [20, 13, 22, 22, 20, 12, 18, 3,
                  20, 14,  9, 10, 11, 16, 16, 20, 26]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Title (rows 1-2) ───────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 38
    ws.merge_cells('A1:G1')
    t1 = ws.cell(row=1, column=1, value='WELLX  |  PREMIUM SUMMARY')
    t1.font      = Font(name='Raleway', bold=True, size=16, color=WHITE)
    t1.fill      = PatternFill('solid', fgColor=PRI)
    t1.alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[2].height = 22
    ws.merge_cells('A2:G2')
    t2 = ws.cell(row=2, column=1, value=company.upper())
    t2.font      = Font(name='Raleway', bold=True, size=11, color=ORANGE)
    t2.fill      = PatternFill('solid', fgColor=PRI)
    t2.alignment = Alignment(horizontal='center', vertical='center')

    # ── Policy Info Block (rows 3–16) ──────────────────────────────────────────
    LABEL_FILL = LGT
    LABEL_CLR  = '374151'

    def info_label(r, c, text):
        cell = ws.cell(row=r, column=c, value=text)
        cell.font      = Font(name='Inter', bold=True, size=9.5, color=LABEL_CLR)
        cell.fill      = PatternFill('solid', fgColor=LABEL_FILL)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    def info_val(r, c, text, bold=False):
        cell = ws.cell(row=r, column=c, value=text)
        cell.font      = Font(name='Inter', bold=bold, size=9.5, color=DARK)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Row 3
    info_label(3, 1, 'Plan');      info_val(3, 2, plan, bold=True)
    ws.merge_cells('E3:F3')
    cv(3, 5, 'FLAT FIXED FEES', bold=True, color=PRI, halign='center', name='Raleway')

    # Row 4
    info_label(4, 1, 'Policy Holder'); info_val(4, 2, company, bold=True)
    cv(4, 5, 'DHA Basmah Fee', bold=True, color=DARK, size=9.5)
    cv(4, 6, 19, halign='center', size=9.5)
    cv(4, 7, 'Paid by client', size=9)

    # Row 5
    info_label(5, 1, 'Broker'); info_val(5, 2, broker_name)
    cv(5, 5, 'DHA HCV Fee', bold=True, color=DARK, size=9.5)
    cv(5, 6, 18, halign='center', size=9.5)
    cv(5, 7, 'Paid by client', size=9)

    # Row 6 — Start Date (DD MMM YYYY)
    info_label(6, 1, 'Policy Start Date')
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d').date()
    except Exception:
        start_dt = start_date
    start_cell = ws.cell(row=6, column=2, value=start_dt)
    start_cell.font          = Font(name='Inter', size=9.5, color=DARK)
    start_cell.alignment     = Alignment(horizontal='center', vertical='center')
    start_cell.number_format = 'DD MMM YYYY'
    cv(6, 5, 'TruDoc Fee', bold=True, color=DARK, size=9.5)
    cv(6, 6, 12, halign='center', size=9.5)
    cv(6, 7, 'Paid by Healthx', size=9)

    # Row 7 — End Date formula (DD MMM YYYY)
    info_label(7, 1, 'Policy End Date')
    end_cell = ws.cell(row=7, column=2)
    end_cell.value          = '=DATE(YEAR(B6)+1,MONTH(B6),DAY(B6))-1'
    end_cell.font           = Font(name='Inter', size=9.5, color=DARK)
    end_cell.alignment      = Alignment(horizontal='center', vertical='center')
    end_cell.number_format  = 'DD MMM YYYY'
    cv(7, 5, 'Slash Data Fee', bold=True, color=DARK, size=9.5)
    cv(7, 6, 25, halign='center', size=9.5)

    # Row 8
    info_label(8, 1, 'Policy Type'); info_val(8, 2, plan_type)

    # Row 9
    cv(9, 6, 'FEES IN %', bold=True, color=PRI, halign='center', name='Raleway', size=9.5)

    # Row 10
    info_label(10, 1, 'Inception Premium Payment'); info_val(10, 2, inception)
    cv(10, 6, 'HSB', bold=True, color=PRI, halign='center', name='Raleway', size=9.5)
    if has_lsb:
        cv(10, 7, 'LSB', bold=True, color=PRI, halign='center', name='Raleway', size=9.5)

    # Row 11
    info_label(11, 1, 'Endorsement Frequency'); info_val(11, 2, endorse)
    cv(11, 5, 'Broker', bold=True, color=DARK, size=9.5)
    cv(11, 6, pct_val(broker_h), halign='center', size=9.5, num_fmt='0.00%')
    if has_lsb:
        cv(11, 7, pct_val(broker_l), halign='center', size=9.5, num_fmt='0.00%')

    # Row 12
    cv(12, 5, f'{insurer_name} (Insurer)', bold=True, color=DARK, size=9.5)
    cv(12, 6, pct_val(insurer_h), halign='center', size=9.5, num_fmt='0.00%')
    if has_lsb:
        cv(12, 7, pct_val(insurer_l), halign='center', size=9.5, num_fmt='0.00%')

    # Row 13
    info_label(13, 1, 'Underwriter'); info_val(13, 2, underwriter)
    cv(13, 5, f'{admin_name} (Admin)', bold=True, color=DARK, size=9.5)
    cv(13, 6, pct_val(admin_h), halign='center', size=9.5, num_fmt='0.00%')
    if has_lsb:
        cv(13, 7, pct_val(admin_l), halign='center', size=9.5, num_fmt='0.00%')

    # Row 14
    cv(14, 5, 'NAS (TPA)', bold=True, color=DARK, size=9.5)
    cv(14, 6, pct_val(nas_h), halign='center', size=9.5, num_fmt='0.00%')
    if has_lsb:
        cv(14, 7, pct_val(nas_l), halign='center', size=9.5, num_fmt='0.00%')

    # Row 15
    cv(15, 5, 'Insurance Levy', bold=True, color=DARK, size=9.5)
    cv(15, 6, pct_val(levy_h), halign='center', size=9.5, num_fmt='0.00%')
    if has_lsb:
        cv(15, 7, pct_val(levy_l), halign='center', size=9.5, num_fmt='0.00%')

    # Row 16 — TOTAL FEES with live SUM formulas
    cv(16, 5, 'TOTAL FEES', bold=True, color=PRI, name='Raleway', size=9.5, halign='center')
    sum_f = ws.cell(row=16, column=6)
    sum_f.value         = '=SUM(F11:F15)'
    sum_f.font          = Font(name='Raleway', bold=True, size=9.5, color=PRI)
    sum_f.alignment     = Alignment(horizontal='center', vertical='center')
    sum_f.number_format = '0.00%'
    if has_lsb:
        sum_g = ws.cell(row=16, column=7)
        sum_g.value         = '=SUM(G11:G15)'
        sum_g.font          = Font(name='Raleway', bold=True, size=9.5, color=PRI)
        sum_g.alignment     = Alignment(horizontal='center', vertical='center')
        sum_g.number_format = '0.00%'

    # Row heights for info block
    for r in range(3, 17):
        ws.row_dimensions[r].height = 18

    # ── Premium Table Header (row 19) ──────────────────────────────────────────
    hdr_cols = ['Category', 'Age Band', 'GROSS PREMIUM\n(MALE)', 'GROSS PREMIUM\n(FEMALE)',
                'NET PREMIUM\n(MALE)', 'NET PREMIUM\n(FEMALE)']
    for ci, h in enumerate(hdr_cols, 1):
        cell = ws.cell(row=19, column=ci, value=h)
        cell.font      = Font(name='Raleway', bold=True, size=9, color=PRI)
        cell.fill      = PatternFill('solid', fgColor=LGT)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = thin_border('AABDD6')
    ws.row_dimensions[19].height = 30

    # ── Premium Data Rows (starting row 20) ────────────────────────────────────
    DATA_CLR = '1a2332'
    EMP_FILL = LGT
    DEP_FILL = 'F5F5F5'
    MAT_FILL = 'FFF9F0'

    row = 20
    MAT_BANDS = {'A': '18 - 45'}
    DEFAULT_MAT_BAND = '18 - 50'

    for cat in sorted(verified_rates.keys()):
        brackets = verified_rates[cat]
        mat_rate = float(maternity_rates.get(cat, 0) or 0)

        # Category header
        ws.merge_cells(f'A{row}:F{row}')
        cell = ws.cell(row=row, column=1, value=f'Category {cat}')
        cell.font      = Font(name='Raleway', bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill('solid', fgColor=PRI)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[row].height = 20
        row += 1

        # EMP rows
        for bi, b in enumerate(brackets):
            age_lo = b.get('age_lo', 0)
            if age_lo < 18:
                continue
            age_label = f"{age_lo} - {b.get('age_hi', 99)}"
            male_g    = b.get('male', 0)
            female_g  = b.get('female', 0)
            fill_c    = EMP_FILL if bi % 2 == 0 else 'EEF2FF'

            for ci, val in enumerate([f'Cat {cat} (EMP)', age_label, male_g, female_g, None, None], 1):
                cell = ws.cell(row=row, column=ci, value=val)
                cell.font      = Font(name='Inter', size=9.5, color=DATA_CLR)
                cell.fill      = PatternFill('solid', fgColor=fill_c)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border    = thin_border()
                if ci in (3, 4):
                    cell.number_format = '#,##0.00'
            net_m = ws.cell(row=row, column=5)
            net_m.value         = f'=C{row}*(1-$F$16)'
            net_m.font          = Font(name='Inter', size=9.5, color=DATA_CLR)
            net_m.fill          = PatternFill('solid', fgColor=fill_c)
            net_m.alignment     = Alignment(horizontal='center', vertical='center')
            net_m.border        = thin_border()
            net_m.number_format = '#,##0.00'
            net_f = ws.cell(row=row, column=6)
            net_f.value         = f'=D{row}*(1-$F$16)'
            net_f.font          = Font(name='Inter', size=9.5, color=DATA_CLR)
            net_f.fill          = PatternFill('solid', fgColor=fill_c)
            net_f.alignment     = Alignment(horizontal='center', vertical='center')
            net_f.border        = thin_border()
            net_f.number_format = '#,##0.00'
            ws.row_dimensions[row].height = 16
            row += 1

        # DEP rows (all brackets)
        for bi, b in enumerate(brackets):
            age_lo    = b.get('age_lo', 0)
            age_label = f"{age_lo} - {b.get('age_hi', 99)}"
            male_g    = b.get('male', 0)
            female_g  = b.get('female', 0)
            fill_c    = DEP_FILL if bi % 2 == 0 else 'EBEBEB'

            for ci, val in enumerate([f'Cat {cat} (DEP)', age_label, male_g, female_g, None, None], 1):
                cell = ws.cell(row=row, column=ci, value=val)
                cell.font      = Font(name='Inter', size=9.5, color=DATA_CLR)
                cell.fill      = PatternFill('solid', fgColor=fill_c)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border    = thin_border()
                if ci in (3, 4):
                    cell.number_format = '#,##0.00'
            net_m = ws.cell(row=row, column=5)
            net_m.value         = f'=C{row}*(1-$F$16)'
            net_m.font          = Font(name='Inter', size=9.5, color=DATA_CLR)
            net_m.fill          = PatternFill('solid', fgColor=fill_c)
            net_m.alignment     = Alignment(horizontal='center', vertical='center')
            net_m.border        = thin_border()
            net_m.number_format = '#,##0.00'
            net_f = ws.cell(row=row, column=6)
            net_f.value         = f'=D{row}*(1-$F$16)'
            net_f.font          = Font(name='Inter', size=9.5, color=DATA_CLR)
            net_f.fill          = PatternFill('solid', fgColor=fill_c)
            net_f.alignment     = Alignment(horizontal='center', vertical='center')
            net_f.border        = thin_border()
            net_f.number_format = '#,##0.00'
            ws.row_dimensions[row].height = 16
            row += 1

        # MAT row
        if mat_rate > 0:
            mat_band = MAT_BANDS.get(cat, DEFAULT_MAT_BAND)
            ws.cell(row=row, column=1).value = f'Cat {cat} (EMPDEPMAT)'
            ws.cell(row=row, column=2).value = mat_band
            ws.cell(row=row, column=4).value = mat_rate

            for ci in range(1, 8):
                cell = ws.cell(row=row, column=ci)
                cell.font      = Font(name='Inter', size=9.5, color=DATA_CLR, italic=True)
                cell.fill      = PatternFill('solid', fgColor=MAT_FILL)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border    = thin_border()

            net_mat = ws.cell(row=row, column=6)
            net_mat.value         = f'=D{row}*(1-$F$16)'
            net_mat.font          = Font(name='Inter', size=9.5, color=DATA_CLR, italic=True)
            net_mat.fill          = PatternFill('solid', fgColor=MAT_FILL)
            net_mat.alignment     = Alignment(horizontal='center', vertical='center')
            net_mat.border        = thin_border()
            net_mat.number_format = '#,##0.00'
            ws.cell(row=row, column=4).number_format = '#,##0.00'
            ws.cell(row=row, column=7).value = 'additional maternity premium'
            ws.cell(row=row, column=7).font  = Font(name='Inter', size=8.5, color='888888', italic=True)
            ws.row_dimensions[row].height = 16
            row += 1

    # ── Additional Loading Section (cols I–Q) ──────────────────────────────────
    cv(21, 9, 'Additional Loading', bold=True, color=PRI, name='Raleway', size=10)

    cv(23, 9, 'Net Loading',   bold=True, color=DARK, size=9.5)
    net_load_cell = ws.cell(row=23, column=10)
    net_load_cell.value         = '=J24*(1-F16)'
    net_load_cell.font          = Font(name='Inter', size=9.5, color=DARK)
    net_load_cell.alignment     = Alignment(horizontal='center', vertical='center')
    net_load_cell.number_format = '#,##0.00'

    cv(24, 9, 'Gross Loading', bold=True, color=DARK, size=9.5)

    # Loading member table headers at row 27
    loading_hdrs = ['Member', 'DOB', 'Gender', 'Category', 'Relation',
                    'Net Loading', 'Gross Loading', 'Declaration', 'Comments']
    for ci, h in enumerate(loading_hdrs, 9):
        cell = ws.cell(row=27, column=ci, value=h)
        cell.font      = Font(name='Raleway', bold=True, size=9, color=WHITE)
        cell.fill      = PatternFill('solid', fgColor=PRI)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = thin_border(WHITE)
    ws.row_dimensions[27].height = 22

    MEMBER_ROW_START = 28
    last_loading_row = MEMBER_ROW_START - 1

    for li, lm in enumerate(loading_members):
        r      = MEMBER_ROW_START + li
        fill_c = LGT if li % 2 == 0 else 'FFFFFF'
        dob_str = str(lm.get('dob', ''))

        vals = [lm.get('name', ''), dob_str, lm.get('gender', ''),
                lm.get('category', ''), lm.get('relation', ''),
                None,
                lm.get('gross_loading', 0),
                lm.get('diagnosis', ''),
                lm.get('notes', '')]

        for ci, val in enumerate(vals, 9):
            cell = ws.cell(row=r, column=ci, value=val)
            cell.font      = Font(name='Inter', size=9.5, color=DATA_CLR)
            cell.fill      = PatternFill('solid', fgColor=fill_c)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border    = thin_border()
            if ci == 15:
                cell.number_format = '#,##0.00'

        # Net Loading formula: =O{row}*(1-$F$16)
        net_cell = ws.cell(row=r, column=14)
        net_cell.value         = f'=O{r}*(1-$F$16)'
        net_cell.font          = Font(name='Inter', size=9.5, color=DATA_CLR)
        net_cell.fill          = PatternFill('solid', fgColor=fill_c)
        net_cell.alignment     = Alignment(horizontal='center', vertical='center')
        net_cell.border        = thin_border()
        net_cell.number_format = '#,##0.00'

        ws.row_dimensions[r].height = 18
        last_loading_row = r

    if loading_members:
        total_r = last_loading_row + 1
        for ci in range(9, 18):
            cell = ws.cell(row=total_r, column=ci)
            cell.font   = Font(name='Raleway', bold=True, size=9.5, color=PRI)
            cell.border = thin_border()
        ws.cell(row=total_r, column=14).value         = f'=SUM(N{MEMBER_ROW_START}:N{last_loading_row})'
        ws.cell(row=total_r, column=14).number_format = '#,##0.00'
        ws.cell(row=total_r, column=15).value         = f'=SUM(O{MEMBER_ROW_START}:O{last_loading_row})'
        ws.cell(row=total_r, column=15).number_format = '#,##0.00'

        j24 = ws.cell(row=24, column=10)
        j24.value         = f'=O{total_r}'
        j24.font          = Font(name='Inter', size=9.5, color=DARK)
        j24.alignment     = Alignment(horizontal='center', vertical='center')
        j24.number_format = '#,##0.00'
    else:
        j24 = ws.cell(row=24, column=10)
        j24.value         = 0
        j24.font          = Font(name='Inter', size=9.5, color='888888', italic=True)
        j24.alignment     = Alignment(horizontal='center', vertical='center')
        j24.number_format = '#,##0.00'

    # ── Reconciliation block ────────────────────────────────────────────────────
    rec_row = row + 2
    ws.merge_cells(f'A{rec_row}:G{rec_row}')
    rec_hdr = ws.cell(row=rec_row, column=1, value='RECONCILIATION')
    rec_hdr.font      = Font(name='Raleway', bold=True, size=10, color=WHITE)
    rec_hdr.fill      = PatternFill('solid', fgColor=MID)
    rec_hdr.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[rec_row].height = 22
    rec_row += 1

    if is_healthx and quoted_totals_calc is not None:
        recon_data = [
            ('Confirmed Quote Premium (PDF)',        q_premium,         '#,##0.00'),
            ('Calculated — Quoted Census',           c_premium,         '#,##0.00'),
            ('Calculated — Confirmed Census',        conf_premium,      '#,##0.00'),
            ('Net Loading Applied',                  total_net_load,    '#,##0.00'),
            ('Final Premium (incl. loading)',         final_grand_total, '#,##0.00'),
            ('Variance (quoted census vs PDF quote)', diff,              '#,##0.00'),
        ]
        if q_members or quoted_member_count:
            recon_data += [
                ('Members – Quote (PDF)',    q_members if q_members else '—',    '0' if q_members else '@'),
                ('Members – Quoted Census',  quoted_member_count if quoted_member_count else '—', '0' if quoted_member_count else '@'),
                ('Members – Confirmed Census', len(members_data), '0'),
            ]
    else:
        recon_data = [
            ('Confirmed Quote Premium (PDF)', q_premium,         '#,##0.00'),
            ('Calculated Premium (Census)',   c_premium,         '#,##0.00'),
            ('Net Loading Applied',           total_net_load,    '#,##0.00'),
            ('Final Premium (incl. loading)', final_grand_total, '#,##0.00'),
            ('Variance (vs quote)',           diff,              '#,##0.00'),
        ]
        if q_members:
            recon_data += [
                ('Members – Quote',   q_members,               '0'),
                ('Members – Census',  len(members_data),       '0'),
                ('Member Variance',   len(members_data)-q_members, '0'),
            ]

    for label, value, fmt in recon_data:
        ws.merge_cells(f'A{rec_row}:B{rec_row}')
        lbl = ws.cell(row=rec_row, column=1, value=label)
        lbl.font      = Font(name='Inter', bold=True, size=9.5, color='374151')
        lbl.fill      = PatternFill('solid', fgColor=LGT)
        val = ws.cell(row=rec_row, column=3, value=value)
        val.font           = Font(name='Inter', size=9.5, color=DARK)
        val.number_format  = fmt
        val.alignment      = Alignment(horizontal='right', vertical='center')
        rec_row += 1

    status_txt = '✅  MATCH (within AED 20)' if is_match else '❌  MISMATCH'
    banner = ws.cell(row=rec_row, column=1, value=status_txt)
    banner.font      = Font(name='Raleway', bold=True, size=11,
                            color=GREEN_FG if is_match else RED_FG)
    banner.fill      = PatternFill('solid', fgColor=GREEN_BG if is_match else RED_BG)
    banner.alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'A{rec_row}:G{rec_row}')
    ws.row_dimensions[rec_row].height = 22
    rec_row += 1

    if not is_match:
        causes = build_mismatch_analysis(members_data, q_premium, c_premium, q_members, total_net_load)
        for cause in causes:
            cc = ws.cell(row=rec_row, column=1, value=f'• {cause}')
            cc.font      = Font(name='Inter', size=9, color=RED_FG, italic=True)
            cc.fill      = PatternFill('solid', fgColor=RED_BG)
            ws.merge_cells(f'A{rec_row}:G{rec_row}')
            rec_row += 1

    # Footer
    ws.cell(row=rec_row+1, column=1).value = f'Created: {datetime.now().strftime("%d %b %Y  %H:%M")}'
    ws.cell(row=rec_row+1, column=1).font  = Font(name='Inter', size=8, color='9aa5b4', italic=True)

    apply_borders(ws1)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 2: PREMIUM PER MEMBER
    # ══════════════════════════════════════════════════════════════════════════
    ws = ws2

    # Column widths (A–S)
    mem_widths = [5, 34, 14, 9, 9, 14, 14, 14, 9, 14, 14, 12, 12, 18, 22, 16, 16, 22, 22]
    for i, w in enumerate(mem_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.row_dimensions[1].height = 42
    ws.merge_cells('A1:S1')
    t1m = ws.cell(row=1, column=1, value='WELLX  |  PREMIUM PER MEMBER')
    t1m.font      = Font(name='Raleway', bold=True, size=16, color=WHITE)
    t1m.fill      = PatternFill('solid', fgColor=PRI)
    t1m.alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[2].height = 22
    ws.merge_cells('A2:S2')
    t2m = ws.cell(row=2, column=1, value=company.upper())
    t2m.font      = Font(name='Raleway', bold=True, size=11, color=ORANGE)
    t2m.fill      = PatternFill('solid', fgColor=PRI)
    t2m.alignment = Alignment(horizontal='center', vertical='center')

    # Summary bar (row 3)
    mem_row = 3
    sub_with_load   = totals['subtotal'] + total_net_load
    vat_with_load   = sub_with_load * VAT_RATE
    grand_with_load = sub_with_load + vat_with_load

    summary_items = [
        ('Total Members',      len(members_data),         False, PRI),
        ('Net Premium',        totals['total_net'],        True,  PRI),
        ('Total Maternity',    totals['total_maternity'],  True,  MID),
        ('BASMAH',             totals['total_basmah'],     True,  '005580'),
        ('Subtotal excl. VAT', sub_with_load,              True,  PRI),
        ('VAT (5%)',           vat_with_load,              True,  PRI),
        ('Grand Total',        grand_with_load,            True,  ORANGE),
    ]
    for i, (label, value, is_aed, color) in enumerate(summary_items):
        col = (i * 2) + 1
        if col + 1 <= 14:
            lc = ws.cell(row=mem_row, column=col, value=label)
            lc.font      = Font(name='Inter', bold=True, size=8, color=WHITE)
            lc.fill      = PatternFill('solid', fgColor=color)
            lc.alignment = Alignment(horizontal='center', vertical='center')
            vc = ws.cell(row=mem_row, column=col+1, value=value)
            vc.font      = Font(name='Inter', bold=True, size=9, color=WHITE)
            vc.fill      = PatternFill('solid', fgColor=color)
            vc.alignment = Alignment(horizontal='center', vertical='center')
            if is_aed:
                vc.number_format = '"AED "#,##0'
    ws.row_dimensions[mem_row].height = 22
    mem_row += 1

    # Column headers (row 4)
    age_col_label = 'Age (ANB)' if is_healthx else 'Age (ALB)'
    hdrs2 = ['No.', 'Full Name', 'Date of Birth', 'Gender', 'Category',
             'Relationship', 'Marital Status', 'Emirate (Visa)', age_col_label, 'Age Bracket',
             'Base Premium', 'Maternity', 'BASMAH', 'Total (excl VAT)', 'Notes',
             'Gross Loading', 'Net Loading', 'Diagnosis', 'Loading Notes']
    for c, h in enumerate(hdrs2, 1):
        cell = ws.cell(row=mem_row, column=c, value=h)
        cell.font      = Font(name='Inter', bold=True, size=9, color=WHITE)
        cell.fill      = PatternFill('solid', fgColor=PRI)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = thin_border(WHITE)
    ws.row_dimensions[mem_row].height = 30
    mem_row += 1

    # Build loading lookup
    loading_map = {lm.get('name', '').strip().lower(): lm for lm in loading_members}

    for m in members_data:
        alt     = m['no'] % 2 == 0
        has_mat = m['maternity_premium'] > 0
        lm_data = loading_map.get(m['name'].strip().lower())
        has_load = lm_data is not None

        if has_load:
            bg = LGT
        elif has_mat:
            bg = ORANGE_BG
        elif alt:
            bg = 'F5F5F5'
        else:
            bg = WHITE

        total = m['base_premium'] + m['maternity_premium'] + m['basmah_fee']
        notes = 'Maternity surcharge applied' if has_mat else (m.get('error', '') or '')

        # Parse DOB to date object for DD MMM YYYY format
        dob_val = m['dob']
        try:
            dob_val = datetime.strptime(str(m['dob']), '%d-%b-%Y').date()
        except Exception:
            try:
                dob_val = datetime.strptime(str(m['dob']), '%Y-%m-%d').date()
            except Exception:
                pass

        vals = [m['no'], m['name'], dob_val, m['gender'], m['category'],
                m['relation'], m.get('marital_status', ''), m.get('emirate', ''),
                m['age_alb'], m['age_bracket'],
                m['base_premium'], m['maternity_premium'], m['basmah_fee'], total, notes,
                float(lm_data.get('gross_loading', 0) or 0) if has_load else None,
                None,
                lm_data.get('diagnosis', '') if has_load else None,
                lm_data.get('notes', '') if has_load else None]

        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=mem_row, column=c, value=v)
            cell.font   = Font(name='Inter', size=9)
            cell.fill   = PatternFill('solid', fgColor=bg)
            cell.border = thin_border()
            if c == 3:
                cell.number_format = 'DD MMM YYYY'
                cell.alignment     = Alignment(horizontal='center', vertical='center')
            elif c in (11, 12, 13, 14, 16):
                cell.number_format = '#,##0.00'
                cell.alignment     = Alignment(horizontal='right', vertical='center')
            elif c == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Net Loading formula col 17 (Q)
        if has_load:
            nl = ws.cell(row=mem_row, column=17)
            nl.value         = f"=P{mem_row}*(1-'Premium Summary'!$F$16)"
            nl.font          = Font(name='Inter', size=9)
            nl.fill          = PatternFill('solid', fgColor=bg)
            nl.border        = thin_border()
            nl.number_format = '#,##0.00'
            nl.alignment     = Alignment(horizontal='right', vertical='center')

        mem_row += 1

    # Totals row
    for c, v in enumerate(['', 'TOTALS', '', '', '', '', '', '', '', '',
                            totals['total_net'], totals['total_maternity'],
                            totals['total_basmah'], totals['subtotal'], '',
                            total_gross_load, total_net_load, '', ''], 1):
        cell = ws.cell(row=mem_row, column=c, value=v)
        cell.font      = Font(name='Raleway', bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill('solid', fgColor=PRI)
        cell.border    = thin_border(WHITE)
        cell.alignment = Alignment(horizontal='right' if c >= 11 else 'left', vertical='center')
        if c in (11, 12, 13, 14, 16, 17):
            cell.number_format = '#,##0.00'
    mem_row += 1

    # VAT row
    ws.cell(row=mem_row, column=13, value='VAT (5%)').font = Font(name='Raleway', bold=True, size=9.5, color=PRI)
    vc2 = ws.cell(row=mem_row, column=14, value=vat_with_load)
    vc2.font          = Font(name='Raleway', bold=True, size=9.5, color=PRI)
    vc2.number_format = '#,##0.00'
    vc2.alignment     = Alignment(horizontal='right', vertical='center')
    mem_row += 1

    # Grand Total row
    ws.merge_cells(f'M{mem_row}:N{mem_row}')
    gt_lbl = ws.cell(row=mem_row, column=13, value='GRAND TOTAL')
    gt_lbl.font      = Font(name='Raleway', bold=True, size=11, color=WHITE)
    gt_lbl.fill      = PatternFill('solid', fgColor=ORANGE)
    gt_lbl.alignment = Alignment(horizontal='right', vertical='center')
    gt_val = ws.cell(row=mem_row, column=14, value=grand_with_load)
    gt_val.font          = Font(name='Raleway', bold=True, size=11, color=WHITE)
    gt_val.fill          = PatternFill('solid', fgColor=ORANGE)
    gt_val.number_format = '#,##0.00'
    gt_val.alignment     = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[mem_row].height = 22
    mem_row += 1

    # ── Summary block (Fix 10) ─────────────────────────────────────────────────
    mem_row += 2

    total_fees_pct = broker_h + insurer_h + admin_h + nas_h + levy_h
    comm_str = (f'Broker {broker_h}% | {insurer_name} {insurer_h}% | '
                f'{admin_name} {admin_h}% | NAS {nas_h}% | Levy {levy_h}% | '
                f'Total {total_fees_pct}%')

    if is_match:
        analysis_txt = '✅ Match (within AED 20)'
    else:
        causes = build_mismatch_analysis(members_data, q_premium, c_premium, q_members, total_net_load)
        analysis_txt = '❌ ' + ' | '.join(causes)

    summary_block = [
        ('Company Name',                  company),
        ('Broker',                        broker_name),
        ('Commission Structure',          comm_str),
        ('Confirmed Premium',             f'AED {q_premium:,.2f}'),
        ('Final Premium (incl. loading)', f'AED {final_grand_total:,.2f}'),
        ('Premium Analysis',              analysis_txt),
    ]

    for label, value in summary_block:
        lbl_cell = ws.cell(row=mem_row, column=1, value=label)
        lbl_cell.font      = Font(name='Inter', bold=True, size=9.5, color='374151')
        lbl_cell.fill      = PatternFill('solid', fgColor=LGT)
        lbl_cell.alignment = Alignment(horizontal='left', vertical='center')

        ws.merge_cells(f'B{mem_row}:E{mem_row}')
        val_cell = ws.cell(row=mem_row, column=2, value=value)
        val_cell.font      = Font(name='Inter', size=9.5, color=DARK)
        val_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        if label == 'Premium Analysis':
            fg_c = GREEN_FG if is_match else RED_FG
            bg_c = GREEN_BG if is_match else RED_BG
            lbl_cell.font = Font(name='Inter', bold=True, size=9.5, color=fg_c)
            lbl_cell.fill = PatternFill('solid', fgColor=bg_c)
            val_cell.font = Font(name='Inter', bold=True, size=9.5, color=fg_c)
            val_cell.fill = PatternFill('solid', fgColor=bg_c)

        ws.row_dimensions[mem_row].height = 18
        mem_row += 1

    apply_borders(ws2)

    return wb


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def api_upload():
    plan = request.form.get('plan', '')

    # ── Healthx path: no tool file, two census files ───────────────────────────
    if plan.lower() == 'healthx':
        quoted_file    = request.files.get('quoted_census')
        confirmed_file = request.files.get('confirmed_census')
        if not quoted_file or not confirmed_file:
            return jsonify({'error': 'Both Quoted Census and Confirmed Census are required for Healthx'}), 400

        token = str(uuid.uuid4())
        _store[token] = {
            'plan':                      'Healthx',
            'quoted_census_bytes':       quoted_file.read(),
            'quoted_census_filename':    quoted_file.filename,
            'confirmed_census_bytes':    confirmed_file.read(),
            'confirmed_census_filename': confirmed_file.filename,
        }
        return jsonify({
            'token':             token,
            'tool_data':         {},
            'vision_categories': {},   # empty → manual entry in UI
        })

    # ── Non-Healthx path (Healthxclusive / Openx): tool + census ──────────────
    tool_file   = request.files.get('tool')
    census_file = request.files.get('census')
    if not tool_file or not census_file:
        return jsonify({'error': 'Both the Quote Tool and Member Census files are required'}), 400

    tool_bytes   = tool_file.read()
    census_bytes = census_file.read()

    try:
        tool_data = parse_healthxclusive_tool(tool_bytes)
    except Exception as e:
        return jsonify({'error': f'Tool parsing error: {str(e)}'}), 400

    token = str(uuid.uuid4())
    _store[token] = {
        'plan':            plan,
        'census_bytes':    census_bytes,
        'census_filename': census_file.filename,
        'tool_data':       tool_data,
    }

    return jsonify({
        'token':             token,
        'tool_data':         tool_data,
        'vision_categories': tool_data['categories'],
    })

@app.route('/api/page/<token>/<int:page_idx>')
def api_page(token, page_idx):
    stored = _store.get(token)
    if not stored:
        return jsonify({'error': 'Session not found'}), 404
    img, total = pdf_page_image(stored['pdf_bytes'], page_idx, scale=2.0)
    return jsonify({'image': img, 'page': page_idx, 'total': total})

@app.route('/api/calculate', methods=['POST'])
def api_calculate():
    body           = request.json
    token          = body.get('token')
    form_data      = body.get('form_data', {})
    verified_rates = body.get('verified_rates', {})   # {cat: [brackets...]}
    maternity_rates= body.get('maternity_rates', {})  # {cat: amount}
    quote_totals   = body.get('quote_totals', {})

    stored = _store.get(token)
    if not stored:
        return jsonify({'error': 'Session expired. Please re-upload files.'}), 400

    start_date  = form_data.get('start_date')
    age_method  = form_data.get('age_method', 'alb')   # 'anb' for Healthx
    plan_stored = stored.get('plan', '')

    if not start_date:
        return jsonify({'error': 'Start date is required'}), 400

    # ── Healthx: parse confirmed census (final premium) + quoted census (reconciliation)
    if plan_stored.lower() == 'healthx':
        try:
            members = parse_census(stored['confirmed_census_bytes'],
                                   stored.get('confirmed_census_filename', 'confirmed_census.xlsx'),
                                   start_date, age_method)
        except Exception as e:
            return jsonify({'error': f'Confirmed census parsing error: {str(e)}'}), 400

        try:
            quoted_members_raw = parse_census(stored['quoted_census_bytes'],
                                              stored.get('quoted_census_filename', 'quoted_census.xlsx'),
                                              start_date, age_method)
        except Exception as e:
            return jsonify({'error': f'Quoted census parsing error: {str(e)}'}), 400
    else:
        # ── Non-Healthx: single census ──────────────────────────────────────────
        try:
            members = parse_census(stored['census_bytes'],
                                   stored.get('census_filename', 'census.xlsx'),
                                   start_date, age_method)
        except Exception as e:
            return jsonify({'error': f'Census parsing error: {str(e)}'}), 400
        quoted_members_raw = None

    if not members:
        return jsonify({'error': 'No valid members found in census file'}), 400

    # Build categories_data
    categories_data = {}
    for cat, brackets in verified_rates.items():
        cat_upper = cat.upper()
        mat_rate  = float(maternity_rates.get(cat, maternity_rates.get('A', 0)) or 0)
        categories_data[cat_upper] = {
            'brackets':      brackets,
            'maternity_rate': mat_rate,
        }

    if not categories_data:
        return jsonify({'error': 'No rate data provided'}), 400

    # Ensure all census categories have a rate (fallback)
    first_cat = sorted(categories_data.keys())[0]
    for m in members:
        c = m['category'].upper()
        if c not in categories_data:
            categories_data[c] = categories_data[first_cat]

    members_data, totals = calculate_premiums(members, categories_data)
    company_name         = form_data.get('company_name', 'Company')

    # For Healthx: also calculate premiums on the quoted census
    quoted_totals_calc = None
    quoted_member_count = None
    if quoted_members_raw is not None:
        for m in quoted_members_raw:
            c = m['category'].upper()
            if c not in categories_data:
                categories_data[c] = categories_data[first_cat]
        _, quoted_totals_calc = calculate_premiums(quoted_members_raw, categories_data)
        quoted_member_count   = len(quoted_members_raw)

    out_token = str(uuid.uuid4())
    _store[out_token] = {
        'company_name':       company_name,
        'members_data':       members_data,
        'verified_rates':     verified_rates,
        'maternity_rates':    maternity_rates,
        'form_data':          form_data,
        'quote_totals':       quote_totals,
        'totals':             totals,
        'quoted_totals_calc': quoted_totals_calc,   # None for non-Healthx
        'quoted_member_count': quoted_member_count, # None for non-Healthx
    }

    q_premium = float(quote_totals.get('total_premium', 0) or 0)
    # For Healthx reconciliation: compare quoted census calc to PDF quote
    if quoted_totals_calc is not None:
        c_premium = quoted_totals_calc['total_net'] + quoted_totals_calc['total_maternity']
    else:
        c_premium = totals['total_net'] + totals['total_maternity']
    diff      = c_premium - q_premium
    is_match  = abs(diff) < 20.0
    q_members = int(quote_totals.get('members', 0) or 0)

    diff_items = []
    if not is_match and q_premium > 0:
        diff_items.append({'label': 'Net Premium', 'quote': q_premium, 'calc': c_premium, 'diff': diff})
    if q_members and q_members != len(members_data):
        diff_items.append({'label': 'Member Count', 'quote': q_members, 'calc': len(members_data),
                           'diff': len(members_data) - q_members})

    return jsonify({
        'success':      True,
        'out_token':    out_token,
        'totals':       totals,
        'member_count': len(members_data),
        'members': [
            {'name': m['name'], 'dob': str(m['dob']),
             'gender': m['gender'], 'category': m['category'],
             'relation': m['relation']}
            for m in members_data
        ],
        'reconciliation': {
            'quote_premium': q_premium,
            'calc_premium':  c_premium,
            'difference':    diff,
            'match':         is_match,
            'items':         diff_items,
        }
    })

@app.route('/api/finalize_summary', methods=['POST'])
def api_finalize_summary():
    body            = request.json or {}
    out_token       = body.get('out_token', '')
    loading_members = body.get('loading_members', [])
    has_lsb         = bool(body.get('has_lsb', False))

    stored = _store.get(out_token)
    if not stored or 'members_data' not in stored:
        return jsonify({'error': 'Session expired. Please recalculate.'}), 400

    # Merge any overrides from the finalize call into form_data
    fd = dict(stored['form_data'])
    if body.get('inception_payment'):
        fd['inception_payment'] = body['inception_payment']
    if body.get('endorsement_freq'):
        fd['endorsement_freq'] = body['endorsement_freq']
    fd['has_lsb'] = has_lsb

    quote_totals        = stored.get('quote_totals', {})
    totals              = stored.get('totals', {})
    quoted_totals_calc  = stored.get('quoted_totals_calc')   # None for non-Healthx
    quoted_member_count = stored.get('quoted_member_count')  # None for non-Healthx

    try:
        wb = make_combined_excel(
            fd,
            stored['members_data'],
            stored['verified_rates'],
            stored['maternity_rates'],
            loading_members,
            has_lsb,
            totals,
            quote_totals,
            quoted_totals_calc=quoted_totals_calc,
            quoted_member_count=quoted_member_count,
        )
    except Exception as e:
        return jsonify({'error': f'Excel generation error: {str(e)}'}), 500

    buf = io.BytesIO()
    wb.save(buf)

    prd_token = str(uuid.uuid4())
    _store[prd_token] = {
        'prd_bytes':    buf.getvalue(),
        'company_name': stored.get('company_name', 'Company'),
    }
    return jsonify({'prd_token': prd_token})


@app.route('/download/<token>/<file_type>')
def download(token, file_type):
    stored = _store.get(token)
    if not stored:
        return "File not found or expired", 404
    company = re.sub(r'[^\w\s-]', '', stored.get('company_name', 'Company')).strip()
    if file_type == 'prd':
        data, filename = stored['prd_bytes'], f"Premium Summary - {company}.xlsx"
    else:
        return "Unknown file type", 400
    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(debug=False, port=port, host='0.0.0.0')
