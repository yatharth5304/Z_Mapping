"""Local-only Zuventus pharmaceutical row classification.

All brands and OCR-recovery descriptors come from zuventus_product.xlsx.
No remote service, online processing mode, or company-specific exception is used.
"""

from pathlib import Path
import argparse
import re

from openpyxl import load_workbook
from pro_diof import (parse_description as parse_pro_diof_description,
                      match_pro_brand, match_diof_brand)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "test1.xlsx"
PRODUCT_MASTER_FILE = BASE_DIR / "zuventus_product.xlsx"
ROW_LIMIT = None
CLEAR_EXISTING = True

EXACT_GARBAGE_ROWS = {
    '',
    '--',
    '++',
    '++--',
    'TOTAL',
    '0',
    '0 0',
    '-- --',
    '-- ++',
    'TOTAL --',
    'EMPTY',
    '++ ++',
    'TOTAL (VALUE IN RS.) --',
    'VALUE IN RS.',
    'END OF REPORT TOTAL',
    'TOTAL QUANTITY',
    'TOTAL VALUE',
    'TOTALS',
    'TOTALS --',
    '-',
    'QUANTITY',
    'TOTAL VALUE --',
    'LAST MONTH SALE',
    'STOCK & SALES ANALYSIS',
    'GRAND TOTAL --',
    'CONTINUED',
    'SUPPLIER NAME',
    'QUANTITY --',
    'PAGE NO.',
    'COMPANY TOTAL --',
    'TOTAL SALES IN THIS PERIOD --',
    'TOTAL CLOSING STOCK (SALES VALUE)',
    'DIVISION : 00 --',
    'TOTAL VALUE (00) --',
    'LAST MONTH SALE FEBRUARY',
    'GRAND TOTAL',
    'PARTICULARS PKG.',
    'TOTAL 0',
    'INDIA LIMITED *',
    'PHARMACUTICALS LTD. *',
    'PHARMACUTICALS LTD.',
    'PURCHASE VALUE --',
    'PAGE',
    'INDIA LIMITED',
    'AMOUNT TOTAL',
    'QTY TOTAL',
    'PURCHASE VALUE 0.00 CLOSING',
    'OUR SOFTWARE MARG ERP 9880074116080233041179880427548 --',
    'GRANDTOTAL',
    'GRANDTOTAL --',
    'SALE VALUE --',
    'OPENING VALUE --',
    '** LIQUDATION IS BASED ON LAST THREE MONTHS SALES --',
    'THIS PDF REPORT IS CREATED FROM MEDICA ULTIMATE. FOR SOFTWARE ENQUIRY CONTACT +91-022-47474747 9750000648/658 9702074265 --',
    'THIS PDF REPORT IS CREATED FROM MEDICA ULTIMATE. FOR SOFTWARE ENQUIRY CONTACT 91-022-47474747 9750000648/658 9702074265 --',
    'TOTAL QUANTITY --',
    'POWERED BY SWILERP FOR RETAIL DISTRIBUTION & CHAIN STORES --',
    'PURCHASE DETAIL :-',
    'SUPPLIER NAME INVOICE',
    '***** --',
    'TOTAL:',
    '#NAME?',
    'RATE .01',
    'VALUE',
    'LAST MONTH',
    'BILL NOS. --',
    'GROUP WISE',
    'PRODUCT NAME STRENGTH',
    'RS.',
    'VALUE 0.00',
    'MANUFACTURER GROUP: --',
    'VALUE IN',
    '* *',
    '** - > NOT SOLD FOR 180 DAYS * -> NOT SOLD FOR 90 DAYS #-> 90 DAYS NEAR EXPIRY STOCK. & -> UNSUPPLIED --',
    'DIVISION TOTAL --',
    'GST --',
    'SUB TOTAL --',
    'DSTK : DAYS STOCK ON 3 MONTH SALE : -2 NO SALE NO STOCK -1 NO SALE 0 NO STOCK REST DAYS STOCK# --',
    'NEW PAGE STARTS HERE ++',
    'NEW PAGE STARTS HERE 0',
    'PRODUCT DETAILS',
    'NEW PAGE STARTS HERE',
    'NEW PAGE STARTS HERE --',
    'ITEM DESCRIPTION BLANK_HEADER1',
    'ITEM DESCRIPTION',
    'ITEM NAME UNIT',
    'ITEM NAME ++',
    'ITEM PACK',
    'ITEM NAME PACK',
    'ITEM NAME',
    'ITEM DESCRIPTION COL2',
    'ITEM DESCRIPTION RATE',
    'ITEM',
    'ITEMS PACKING',
    'ITEM NAME UOM',
    'ITEM NAME PACKG',
    'ITEM DESCRIPTION OPENING STOCK (1)',
    'ITEM NAME PACK SCM',
    'NAME OF ITEM',
    'NO/PRODUCT NAME UNIT',
    'PRODUCT NAME PACKING',
    'PRODUCT & PACK',
    'PRODUCT PACK',
    'PRODUCT PKG',
    'PRODUCT DESCRIPTION PACKING',
    'PRODUCT NAME PACK',
    'NON MOVING PRODUCTS ABOVE 30 DAYS --',
    'PRODUCT NAME UNIT',
    'PRODUCT NAME',
    'SRNO. PRODUCT',
    'PRODUCTNAME PACK',
    'PRODUCT NAME BLANK_HEADER1',
    'NON PRODUCTS ABOVE 30 BLANK_HEADER1',
    'PRODUCT NAME PACKG',
    'PRODUCTNAME PACK OP',
    'PRODUCT OPENING',
    'PRODUCT NAME AND PACK',
    'PRODUCT PACKING',
    'PRODUCT DESRIPTION',
    'PRODUCT NAME CI VAL',
    'NON MOVING PRODUCTS ABOVE 90 DAYS',
    'PARTICULARS PACK',
    'COL1 COL2',
    'ID',
    'ID --',
    'PRODUCT NAME SHELF PACKING',
    'CATEGORY TOTAL --',
    'DISTRIBUTION CHAIN STORES',
    'DEFAULT',
    'All Marketing Groups',
}

# Keep the inherited regex.match boundary behavior.
SAFE_GARBAGE_PATTERNS = (
    re.compile('^MANUFACTURER\\s+\\d+$', re.IGNORECASE),
    re.compile('^QUANTITY\\s+\\d+(?:\\.\\d+)?$', re.IGNORECASE),
    re.compile('^VALUE\\s+\\d+(?:\\.\\d+)?$', re.IGNORECASE),
    re.compile('^NO\\s*/\\s*PRODUCT\\s+NAME\\s+UNIT$', re.IGNORECASE),
    re.compile('^PAGE\\s+NO\\.?\\s*\\d+.*PRODUCT\\s+NAME.*PACKING.*ID$', re.IGNORECASE),
    re.compile('^Bill(s?|[\\s:].*|Nos\\.?|--)\\b', re.IGNORECASE),
    re.compile('^[A-Z]{4}\\d{6,}\\s*/\\s*\\d{2}/\\d{2}/\\d{4}'),
    re.compile('^Division\\s*:', re.IGNORECASE),
    re.compile('^Division\\s+\\w', re.IGNORECASE),
    re.compile('^Total\\s*[:(]', re.IGNORECASE),
    re.compile('^Total\\s+Value\\s*\\(', re.IGNORECASE),
    re.compile('^Company\\s*(Group)?\\s*:', re.IGNORECASE),
    re.compile('^Com\\s*:\\s*', re.IGNORECASE),
    re.compile('^Company\\s+\\w', re.IGNORECASE),
    re.compile('^Primary Company\\s+', re.IGNORECASE),
    re.compile('^Company Name\\s+', re.IGNORECASE),
    re.compile('^Supplier\\s+\\w', re.IGNORECASE),
    re.compile('^Store\\s+\\w', re.IGNORECASE),
    re.compile('^Manufacturer\\s+(Group|Name)\\s*:', re.IGNORECASE),
    re.compile('^For Manufacturer Group\\s*:', re.IGNORECASE),
    re.compile('^(Firm|Group|Com)[\\s\\w]*Total', re.IGNORECASE),
    re.compile('^COMNAME\\s+TOTAL', re.IGNORECASE),
    re.compile('^ANNEXURE\\s+\\d', re.IGNORECASE),
    re.compile('^Generated at\\s+\\d{4}-\\d{2}-\\d{2}', re.IGNORECASE),
    re.compile('^Prepared by\\s+\\w+\\s+on\\s+\\d', re.IGNORECASE),
    re.compile('^Print Date\\s+\\d', re.IGNORECASE),
    re.compile('^Print\\s+HEALTH\\s+QRCODE', re.IGNORECASE),
    re.compile('^Powered\\s+By:', re.IGNORECASE),
    re.compile('^Medica Ultimate', re.IGNORECASE),
    re.compile('^Report End\\)', re.IGNORECASE),
    re.compile('^Printed on', re.IGNORECASE),
    re.compile('^hr/>'),
    re.compile('^u>'),
    re.compile('^span\\s+style', re.IGNORECASE),
    re.compile('^PURC\\s*:', re.IGNORECASE),
    re.compile('^EXP\\d+M\\s*:', re.IGNORECASE),
    re.compile('^(Sales|Opening|Closing|Open|Close)\\s+(Stock|Val|Value)(s)?\\s*[:\\-]?', re.IGNORECASE),
    re.compile('^(Opening|Closing)\\s+Val\\.', re.IGNORECASE),
    re.compile('^Values?\\s*:\\s*[\\d.]', re.IGNORECASE),
    re.compile('^(TAX AMOUNT|COMPANY NAME:)', re.IGNORECASE),
    re.compile('^(UC SALE|CL\\.STK\\.|MR\\.Balance|Op\\.Val\\.|Pur\\.Val\\.)', re.IGNORECASE),
    re.compile('^(Prev Month Tot|Value Age)', re.IGNORECASE),
    re.compile('^STOCK\\s*&\\s*SALES STATEMENT', re.IGNORECASE),
    re.compile('^Purchase Return', re.IGNORECASE),
    re.compile('^\\w+\\)\\s+(Sale|Purchase|Stock|Return)\\s+Value', re.IGNORECASE),
    re.compile('^Non Moving Since\\s+\\d+\\s+Days?', re.IGNORECASE),
    re.compile('^Sales and Last Month\\s+sales based on', re.IGNORECASE),
    re.compile('^Last Month\\s+sales between', re.IGNORECASE),
    re.compile('^Purchase bill included values', re.IGNORECASE),
    re.compile('^(PENDING DEBIT NOTES|NO SUGGESTION|SHEDULED DRUG)$', re.IGNORECASE),
    re.compile('\\b(AGENCIES?|DISTRIBUTORS?|MEDICO[SE]?|CORPORATION|ENTERPRISES?|TRADERS?|TRADING)\\b', re.IGNORECASE),
    re.compile('\\b(PVT\\.?\\s+)?LTD\\.?\\s*$', re.IGNORECASE),
    re.compile('\\bLIMITED\\.?\\s*$', re.IGNORECASE),
    re.compile('(?:.*\\b(PHARMACEUTICALS?|PHARMACEUTICAL|PHARMA|BIOPHARMA|BIOPHARMACEUTICALS?|PHARMACEUTI|PHARMACEUTICA)\\b.*\\b(PVT|LTD|LIMITED|DIV|INC|CORP|TOTAL|MFG|MANUFACTURER|COMPANY|DIVISION|SUPPLIER|STORE)\\b|.*\\b(PHARMACEUTICALS?|PHARMACEUTICAL|PHARMA|BIOPHARMA|BIOPHARMACEUTICALS?|PHARMACEUTI|PHARMACEUTICA)\\b.*(?:--|__|//|/u|/w|/d)|^(?:.*\\b)?(?:PHARMACEUTICALS?|PHARMACEUTICAL|BIOPHARMA|BIOPHARMACEUTICALS?)\\b(?:.*\\b(?:PVT|LTD|LIMITED|DIV|INC|CORP|TOTAL|MFG|COMPANY|DIVISION|SUPPLIER|STORE)\\b|.*(?:--|__|//|/u|/w|/d)|\\s*$)|.*\\bPHARMACEUTICA\\d+\\b.*|.*\\bPharmAssist\\b.*|.*\\bPHARMACEUTICALS\\w+\\b.*)', re.IGNORECASE),
    re.compile('^GSTIN\\s*:', re.IGNORECASE),
    re.compile('^GST\\s*NO\\.?\\s*:', re.IGNORECASE),
    re.compile('^Op\\.\\s+Amt', re.IGNORECASE),
    re.compile('^O\\.B\\.\\s+PUR\\.VAL', re.IGNORECASE),
    re.compile('^Items\\s+\\d+\\s*$', re.IGNORECASE),
    re.compile('^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\\s+(Sal|Sale|Stock|Pur)\\b', re.IGNORECASE),
    re.compile('^Copyright\\s+\\d{4}', re.IGNORECASE),
    re.compile('^Powered by\\s', re.IGNORECASE),
    re.compile('\\bcolumn_\\d+\\b', re.IGNORECASE),
    re.compile('^TEST\\s+ITEM\\b', re.IGNORECASE),
    re.compile('^[A-Z]{4}\\d{4,}\\s+Dt\\.\\d{2}/\\d{2}/\\d{4}', re.IGNORECASE),
    re.compile('^Manufacturer\\s+\\d+\\s*~', re.IGNORECASE),
    re.compile('^MARG\\s+ERP\\b', re.IGNORECASE),
    re.compile('^\\s*(TAB(LE)?\\s+)*TAB(LE)?\\s*$', re.IGNORECASE),
    re.compile('\\b(BEDSHEET|BEDSEET|HOTPOT|TOWEL|DUFFAL|ELETRAL|WATER\\s+JAR|CHOCOLATE)\\b', re.IGNORECASE),
    re.compile('\\bSTELL?\\s+BOW[AL]+\\b', re.IGNORECASE),
    re.compile('\\bSTEEL\\s+GLASS\\b', re.IGNORECASE),
    re.compile('^Item\\s+Name\\b', re.IGNORECASE),
    re.compile('\\bSOFTWARE\\s+LICEN', re.IGNORECASE),
    re.compile('^Receipt\\s+Value\\b', re.IGNORECASE),
    re.compile('^Near\\s+Expiry\\s+Product\\b', re.IGNORECASE),
    re.compile('\\bMEDICAL\\s+STOR\\b', re.IGNORECASE),
    re.compile('^[^A-Za-z0-9]*[A-Za-z0-9]{0,3}[^A-Za-z0-9]*$'),
    re.compile('^Totals?\\s*[\\d:,+.\\s\\-]+$', re.IGNORECASE),
    re.compile('^Totals?\\s+[\\d,]+(?:\\.\\d+)?', re.IGNORECASE),
    re.compile('^Purchase\\s+[\\d,]+(?:\\.\\d+)?\\s+Purchase\\s+Return', re.IGNORECASE),
    re.compile('^Total\\s+Of\\s+\\w', re.IGNORECASE),
    re.compile('^Total\\s+(cost\\s+of|G/R)\\s+', re.IGNORECASE),
    re.compile('^Sale\\s+of\\s+period\\s+[\\d.]', re.IGNORECASE),
    re.compile('^Z{4,}\\b', re.IGNORECASE),
    re.compile('\\bDRUG\\s+HOUSE\\b', re.IGNORECASE),
    re.compile('^Digital\\s+Purchase\\s+ERP\\b', re.IGNORECASE),
)

METADATA_GARBAGE_VOCAB = {
    # aggregations / totals
    "TOTAL", "SUBTOTAL", "GRAND", "GRANDTOTAL", "VALUE", "QUANTITY", "QTY",
    "AMOUNT", "NET", "GROSS", "SUM",
    # financial
    "MRP", "RATE", "DISCOUNT", "CREDIT", "DEBIT", "BALANCE",
    "PURCHASE", "SALE", "SALES", "INVOICE", "RS",
    # document / report structure
    "PAGE", "REPORT", "SUMMARY", "ANALYSIS", "GROUP", "CATEGORY",
    "DIVISION", "COMPANY", "SUPPLIER", "MANUFACTURER", "DISTRIBUTOR",
    "DATE", "MONTH", "YEAR", "PERIOD", "CONTINUED",
    # stock
    "STOCK", "CLOSING", "OPENING", "CARRY", "FORWARD",
    # column / header labels
    "PRODUCT", "ITEM", "NAME", "PACKING", "PACK", "UNIT", "UOM",
    "PACKG", "DETAIL", "DETAILS", "DESCRIPTION", "DESRIPTION",
    "CODE", "TYPE", "SRNO", "SR", "ID", "NO", "COL", "BLANK", "HEADER",
    "PARTICULARS", "PKG", "STRENGTH", "SCM",
    # movement / ageing
    "LAST", "CURRENT", "NEW", "MOVING", "NON", "ABOVE", "DAYS",
    # month names — needed for rows like "LAST MONTH SALE MAY"
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
    # common prepositions / conjunctions safe to include
    "IN", "OF", "BY", "THE", "AND", "FOR", "TO", "AT", "FROM",
    "THIS", "SUB", "WISE", "BRAND",
    # misc garbage-header words
    "STARTS", "HERE", "SECTION", "SCHEDULE", "SHELF",
}

MAYBE_PRODUCT_FORM_PATTERN = re.compile(
    r"\b(TABS?|TABLETS?|CAPS?|CAPSULES?|SYP|SYRUP|INJ|INJN|OINT|OINTMENT|CREAM|GEL|"
    r"DROPS?|LOTION|SUSP|SUSPENSION|SACHET?|SACH|POWDER|SPRAY|INHALER|PATCH|SYRINGE|"
    r"VIAL|AMP|AMPULE|AMPOULE|SOLUTION|SOLN|SYR|SUS|SOL|PFS|LOZENGE|"
    r"SUPPOSITORIES?|SUPPOS?|SUPPO|SUPP|ENEMA|EYE|EAR|NASAL|CHEWABLE|EFFERVESCENT)\b",
    re.IGNORECASE,
)

MAYBE_PRODUCT_QTY_PATTERN = re.compile(
    r"(\b\d+\.?\d*\s*(MG|MCG|IU|ML|GM|G)\b"      # strength: 500MG, 1.5ML, 100IU
    r"|\b\d+\s*[xX*]\s*\d+\b"                      # pack: 1X10, 10X10, 1*10
    r"|\b\d+['']?\s*[sS]\b"                         # count: 10S, 15s, 10'S, 15's
    r"|\b\d+\s*(TAB|TABS|CAP|CAPS|VIAL|AMP)\b"     # qty+form: 10TAB, 1VIAL
    r"|\b\d+\s*(ML|GM|G)\b"                         # volume/weight: 100ML, 30GM
    r"|\b\d+\s*MD\b"                                # dose unit: 500MD, 1MD
    r"|\bPCS\b)",                                   # pieces: PCS anywhere in row
    re.IGNORECASE,
)

MAYBE_PRODUCT_GUARD_PATTERN = re.compile(
    r"\b(TOTAL|SALE|PURCHASE|INVOICE|STOCK|VALUE|REPORT|SUMMARY|IMPORT|CALL)\b",
    re.IGNORECASE,
)

def _master_values(master_file=None):
    """Read one validated master sheet without keeping workbook handles open."""
    path = Path(master_file) if master_file is not None else Path(PRODUCT_MASTER_FILE)
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        headers = next(rows, None)
        if headers is None:
            raise ValueError(f"{path} is empty")
        columns = {str(value).strip(): index for index, value in enumerate(headers) if value is not None}
        required = {"BRAND_NAME", "VARIANT", "SUB_VARIANT", "product_name"}
        missing = required - set(columns)
        if missing:
            raise ValueError(f"{path} is missing required columns: {sorted(missing)}")
        return columns, list(rows)
    finally:
        wb.close()


def load_brands(master_file=None):
    """Return distinct master BRAND_NAME values in their worksheet order."""
    columns, rows = _master_values(master_file)
    result = []
    seen = set()
    for row in rows:
        value = row[columns["BRAND_NAME"]]
        if value is None or not str(value).strip():
            continue
        brand = str(value).strip()
        if brand not in seen:
            result.append(brand)
            seen.add(brand)
    return result


def normalize_text(value):
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def row_tokens(value):
    return re.findall(r"[A-Z0-9]+", str(value).upper())


def edit_distance_leq_one(a, b):
    """Return the edit type when Levenshtein distance is <= 1, else None."""
    if a == b:
        return "exact"
    if abs(len(a) - len(b)) > 1:
        return None

    if len(a) > len(b):
        a, b = b, a
        shorter_is_first = False
    else:
        shorter_is_first = True

    i = 0
    j = 0
    edits = 0
    mismatch_type = None
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1
            j += 1
            continue
        edits += 1
        if edits > 1:
            return None
        if len(a) == len(b):
            mismatch_type = "1-substitution"
            i += 1
            j += 1
        else:
            mismatch_type = "1-insertion" if shorter_is_first else "1-deletion"
            j += 1

    if j < len(b) or i < len(a):
        edits += 1
        if len(a) == len(b):
            mismatch_type = "1-substitution"
        else:
            mismatch_type = "1-insertion" if shorter_is_first else "1-deletion"

    return mismatch_type if edits <= 1 else None


def damerau_distance(a, b, max_dist=2):
    """Damerau-Levenshtein distance between a and b, capped at max_dist.
    Handles substitution, insertion, deletion, and adjacent transpositions.
    Returns the integer distance if <= max_dist, else None.
    Uses per-row early exit: if the minimum value in a completed DP row already
    exceeds max_dist, the final distance provably exceeds max_dist too.
    """
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if abs(la - lb) > max_dist:
        return None

    # DP table; initialised to a value larger than any real distance we'd accept
    inf = la + lb + 1
    dp = [[inf] * (lb + 1) for _ in range(la + 1)]
    for i in range(la + 1):
        dp[i][0] = i
    for j in range(lb + 1):
        dp[0][j] = j

    for i in range(1, la + 1):
        row_min = inf
        for j in range(1, lb + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            dp[i][j] = min(
                dp[i - 1][j] + 1,        # deletion
                dp[i][j - 1] + 1,        # insertion
                dp[i - 1][j - 1] + cost, # substitution
            )
            # adjacent transposition
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                dp[i][j] = min(dp[i][j], dp[i - 2][j - 2] + 1)
            if dp[i][j] < row_min:
                row_min = dp[i][j]
        # Early exit: if no cell in this row is <= max_dist, no path can succeed
        if row_min > max_dist:
            return None

    dist = dp[la][lb]
    return dist if dist <= max_dist else None


def token_spans(name):
    """Generate normalized consecutive token spans across the full row."""
    tokens = row_tokens(name)
    spans = []
    for start in range(len(tokens)):
        combined = ""
        span_tokens = []
        for end in range(start, len(tokens)):
            span_tokens.append(tokens[end])
            combined += tokens[end]
            spans.append({
                "tokens": tuple(span_tokens),
                "text": " ".join(span_tokens),
                "norm": combined,
            })
    return spans


_BRAND_CHAR_SETS = {}


def build_brand_index(brands):
    index = [(brand, normalize_text(brand)) for brand in brands]
    for _, brand_norm in index:
        if brand_norm not in _BRAND_CHAR_SETS:
            _BRAND_CHAR_SETS[brand_norm] = frozenset(brand_norm)
    return index


def _fused_identity_pattern(brand_norm, variant, strength, product_name):
    """Compile a complete, literal master identity for otherwise unresolved rows."""
    if len(brand_norm) < 4 or not re.fullmatch(r"\d+(?:\.\d+)?", strength):
        return None
    dose = re.escape(strength)
    units = set(re.findall(
        rf"(?<![\d.]){dose}\s*(MCG|MG|GM|ML|IU|UNIT|G)\b",
        str(product_name).upper(),
    ))
    if len(units) != 1:
        return None
    unit = re.escape(next(iter(units)))
    brand = re.escape(brand_norm)
    if variant:
        identity = rf"{brand}{re.escape(variant)}\s+{dose}(?:\s*{unit})?"
    else:
        # A bare brand+number is insufficient: an attached dose needs its unit.
        identity = rf"{brand}{dose}{unit}"
    # Count packs carry no dose; a unit-bearing pack must repeat the same dose/unit.
    pack = rf"(?:[1-9]\d*[X*][1-9]\d*|[1-9]\d*[X*]{dose}{unit})"
    return re.compile(rf"{identity}(?:\s+{pack})?")


def build_master_product_evidence(brand_index, master_file=None):
    """Return the master-backed evidence needed for conservative OCR recovery.

    The normal brand matcher intentionally works on complete token spans.  OCR
    sometimes glues a real brand to a master variant, making the ordinary
    token-span brand match unreachable.
    This index keeps only verified non-PLAIN variants and numeric strengths
    from the existing product master; it does not add any hand-written product
    aliases.
    """
    brand_norm_to_brand = {brand_norm: brand for brand, brand_norm in brand_index}
    evidence = {}

    columns, rows = _master_values(master_file)
    for row in rows:
        brand_norm = normalize_text(row[columns["BRAND_NAME"]] or "")
        canonical_brand = brand_norm_to_brand.get(brand_norm)
        if not canonical_brand:
            continue

        entry = evidence.setdefault(
            brand_norm,
            {
                "brand": canonical_brand,
                "variants": set(),
                "strengths": set(),
                "products": [],
            },
        )
        variant = str(row[columns["VARIANT"]] or "").strip()
        sub_variant = str(row[columns["SUB_VARIANT"]] or "").strip()
        variant_norm = normalize_text(variant)
        sub_variant_norm = normalize_text(sub_variant)
        if variant and variant.upper() != "PLAIN":
            entry["variants"].add(variant_norm)
        if sub_variant and re.search(r"\d", sub_variant):
            entry["strengths"].add(sub_variant_norm)
        entry["products"].append(
            {
                "variant": "" if variant.upper() == "PLAIN" else variant_norm,
                "strength": "" if sub_variant.upper() == "PLAIN" else sub_variant_norm,
                "forms": master_form_groups(row[columns["product_name"]] or ""),
                "fused_pattern": _fused_identity_pattern(
                    brand_norm,
                    "" if variant.upper() == "PLAIN" else variant_norm,
                    sub_variant,
                    row[columns["product_name"]] or "",
                ),
            }
        )

    return evidence


def master_form_groups(value):
    """Return coarse dosage-form groups from a canonical master product name."""
    text = normalize_text(value)
    groups = set()
    for aliases, group in (
        (("TABLETS", "TABLET", "TABS", "TAB"), "TAB"),
        (("CAPSULES", "CAPSULE", "CAPS", "CAP"), "CAP"),
        (("INJECTION", "INJ", "VIAL"), "INJ"),
        (("DROPS", "DROP"), "DROP"),
        (("SYRUP", "SYP", "SUSPENSION", "SUSP"), "SYP"),
        (("CREAM",), "CREAM"),
        (("OINTMENT", "OINT"), "OINT"),
        (("GEL",), "GEL"),
        (("SOLUTION", "SOLN", "SOL"), "SOL"),
    ):
        if any(alias in text for alias in aliases):
            groups.add(group)
    return groups


def input_form_groups(value):
    """Return the same coarse dosage-form groups for an OCR input row."""
    return master_form_groups(value)


def _fused_master_identity_match(name, master_evidence):
    """Recover one complete master identity, never an arbitrary brand substring."""
    text = str(name).strip().upper()
    first_token = text.split(maxsplit=1)[0] if text else ""
    matched_brand = None
    for brand_norm, entry in master_evidence.items():
        if not first_token.startswith(brand_norm) or len(first_token) <= len(brand_norm):
            continue
        for product in entry["products"]:
            pattern = product.get("fused_pattern")
            if pattern is not None and pattern.fullmatch(text):
                if matched_brand is not None:
                    return None
                matched_brand = entry["brand"]
    if matched_brand is None:
        return None
    return {
        "brand": matched_brand,
        "matched_text": first_token,
        "match_type": "master-fused-identity",
    }


def master_product_match(name, master_evidence, precomputed_spans=None):
    """Return verified master evidence for fused variants or a narrow OCR typo.

    This path is deliberately available only after normal exact/fuzzy brand
    matching fails. Without a form/quantity signal, only a complete, unique
    fused master identity is accepted. The existing form/quantity-backed
    recovery paths otherwise retain their original behavior.
    """
    if not master_evidence:
        return None

    has_form = bool(MAYBE_PRODUCT_FORM_PATTERN.search(name))
    has_qty = bool(MAYBE_PRODUCT_QTY_PATTERN.search(name))
    if not (has_form or has_qty):
        return _fused_master_identity_match(name, master_evidence)

    spans = precomputed_spans if precomputed_spans is not None else token_spans(name)
    single_token_spans = [span for span in spans if len(span["tokens"]) == 1]
    row_norm = normalize_text(name)
    row_tokens_set = set(row_tokens(name))
    input_forms = input_form_groups(name)

    # Product-level evidence also covers a base brand fused to a strength or
    # dosage form. The brand must be a
    # real master brand and at least one master-backed descriptor must agree;
    # this is not a generic "looks like a product" promotion.
    for brand_norm, entry in master_evidence.items():
        if len(brand_norm) < 4 or brand_norm not in row_norm:
            continue
        for product in entry["products"]:
            variant = product["variant"]
            strength = product["strength"]
            product_forms = product["forms"]
            variant_present = not variant
            if variant:
                variant_present = (
                    brand_norm + variant in row_norm
                    or variant in row_tokens_set
                    or (len(variant) > 1 and variant in row_norm)
                )

            strength_present = not strength
            if strength:
                strength_present = bool(
                    re.search(
                        rf"{re.escape(brand_norm)}{re.escape(strength)}(?!\d)",
                        row_norm,
                    )
                    or re.search(
                        rf"(?<!\d){re.escape(strength)}\s*(?:MG|MCG|GM|ML|IU|G|K)\b",
                        str(name).upper(),
                    )
                    or strength in row_tokens_set
                    or (any(ch.isalpha() for ch in strength) and strength in row_norm)
                )

            forms_agree = not product_forms or bool(product_forms & input_forms)
            descriptor_present = (
                (variant and variant_present)
                or (strength and strength_present)
                or (product_forms and forms_agree and input_forms)
            )
            if not (variant_present and strength_present and forms_agree and descriptor_present):
                continue

            matched_text = entry["brand"]
            for span in spans:
                if brand_norm in span["norm"]:
                    matched_text = span["text"]
                    break
            return {
                "brand": entry["brand"],
                "matched_text": matched_text,
                "match_type": "master-product",
            }

    # Exact master brand + non-PLAIN variant evidence handles OCR fusion while
    # retaining the existing token-boundary protection for ordinary brand hits.
    for span in single_token_spans:
        for brand_norm in sorted(master_evidence, key=len, reverse=True):
            suffix = span["norm"][len(brand_norm):]
            if (
                span["norm"].startswith(brand_norm)
                and suffix
                and suffix in master_evidence[brand_norm]["variants"]
            ):
                return {
                    "brand": master_evidence[brand_norm]["brand"],
                    "matched_text": span["text"],
                    "match_type": "master-variant",
                }

    # A two-edit brand spelling is more permissive than the normal matcher, so
    # require both form and quantity evidence plus an exact numeric master
    # strength.  Excluding prefix/suffix truncations prevents a partial brand
    # from being promoted to a longer master brand.
    if not (has_form and has_qty):
        return _fused_master_identity_match(name, master_evidence)
    candidates = []
    for span in single_token_spans:
        for brand_norm, entry in master_evidence.items():
            if len(brand_norm) < 8 or abs(len(span["norm"]) - len(brand_norm)) > 2:
                continue
            if brand_norm.startswith(span["norm"]) or span["norm"].startswith(brand_norm):
                continue
            if not any(
                re.search(rf"(?<!\d){re.escape(strength)}(?!\d)", row_norm)
                for strength in entry["strengths"]
            ):
                continue
            distance = damerau_distance(span["norm"], brand_norm, max_dist=2)
            if distance is not None and distance > 0:
                candidates.append((distance, -len(brand_norm), entry["brand"], span["text"]))

    if not candidates:
        return _fused_master_identity_match(name, master_evidence)
    candidates.sort()
    if len(candidates) > 1 and candidates[0][:2] == candidates[1][:2]:
        return None
    return {
        "brand": candidates[0][2],
        "matched_text": candidates[0][3],
        "match_type": "master-2-edit",
    }


def exact_product_match(name, brand_index, precomputed_spans=None):
    spans = precomputed_spans if precomputed_spans is not None else token_spans(name)
    if not spans:
        return None

    for brand, brand_norm in brand_index:
        if len(brand_norm) <= 3:
            master_tokens = tuple(row_tokens(brand))
            for span in spans:
                if span["norm"] == brand_norm and (
                    len(span["tokens"]) == 1
                    or span["tokens"] == master_tokens
                ):
                    return {
                        "brand": brand,
                        "matched_text": span["text"],
                        "match_type": "exact",
                    }
            continue
        for span in spans:
            if span["norm"] == brand_norm:
                return {
                    "brand": brand,
                    "matched_text": span["text"],
                    "match_type": "exact",
                }
    return None


def fuzzy_product_match(name, brand_index, precomputed_spans=None):
    spans = precomputed_spans if precomputed_spans is not None else token_spans(name)
    if not spans:
        return None

    # Lower rank = better quality match; 2-edit is least preferred
    rank = {
        "1-substitution": 0,
        "1-insertion": 1,
        "1-deletion": 2,
        "1-transposition": 3,
        "2-edit": 4,
    }
    best_match = None

    # One character set per span instead of one per (brand, span) pair. Parallel
    # to `spans`, so the span visitation order inside the brand loop is unchanged.
    span_sets = [frozenset(span["norm"]) for span in spans]

    for brand, brand_norm in brand_index:
        brand_len = len(brand_norm)
        if brand_len <= 3:
            continue  # 3-char brands handled exclusively by exact_product_match

        # Adaptive threshold: 2 edits only for long brands (>= 10 chars) where
        # accidental collisions with common English words are extremely unlikely.
        # Brands 4-9 chars stay at <= 1 edit to prevent false positives.
        max_edits = 2 if brand_len >= 10 else 1
        brand_set = _BRAND_CHAR_SETS[brand_norm]

        for span, span_set in zip(spans, span_sets):
            span_norm = span["norm"]
            span_len = len(span_norm)
            if abs(span_len - brand_len) > max_edits:
                continue

            # --- Character-set pre-filter (set ops, runs before the O(m*n) DP) ---
            # A character present in the span but absent from the brand has to be
            # removed by a deletion or a substitution -- transpositions only reorder
            # characters, they never remove one -- and distinct characters need
            # distinct edits. So len(span_set - brand_set) is a lower bound on the
            # Damerau-Levenshtein distance: a necessary condition, which means no
            # pair the exact distance checks below would accept is rejected here.
            if len(span_set - brand_set) > max_edits:
                continue

            # --- Try standard single-edit first for precise match type ---
            match_type = edit_distance_leq_one(span_norm, brand_norm)
            if match_type == "exact":
                continue  # exact hits are handled by exact_product_match

            if match_type is None:
                # Could still be a transposition (DL dist == 1) or 2-edit (for long brands)
                dist = damerau_distance(span_norm, brand_norm, max_dist=max_edits)
                if dist is None:
                    continue
                if dist == 1:
                    match_type = "1-transposition"
                elif dist == 2 and brand_len >= 10:
                    match_type = "2-edit"
                else:
                    continue

            candidate = {
                "brand": brand,
                "matched_text": span["text"],
                "match_type": match_type,
            }
            if best_match is None:
                best_match = candidate
                continue
            current_rank = rank.get(candidate["match_type"], 99)
            best_rank = rank.get(best_match["match_type"], 99)
            if current_rank < best_rank:
                best_match = candidate
                continue
            if current_rank == best_rank and len(span["norm"]) > len(normalize_text(best_match["matched_text"])):
                best_match = candidate

    return best_match


def build_garbage_set():
    garbage = {normalize_text(value) for value in EXACT_GARBAGE_ROWS}
    garbage.add("")
    return garbage


def is_exact_garbage_row(name, garbage_rows):
    return normalize_text(name) in garbage_rows


def is_pattern_garbage_row(name):
    text = str(name).strip()
    return any(pattern.match(text) for pattern in SAFE_GARBAGE_PATTERNS)


def is_metadata_garbage_row(name):
    """Return True when every alphabetic token in the row is a known
    document-structure / metadata word, indicating the row is a header,
    summary, or total line rather than a real product entry."""
    tokens = re.findall(r"[A-Z]+", str(name).upper())
    if not tokens:
        return False  # blank / numeric-only rows are handled elsewhere
    return all(t in METADATA_GARBAGE_VOCAB for t in tokens)


def fuzzy_garbage_match(name, garbage_rows):
    row_norm = normalize_text(name)
    if not row_norm:
        return None

    best_match = None
    rank = {"exact": 0, "1-substitution": 1, "1-insertion": 2, "1-deletion": 3}

    for garbage_phrase in EXACT_GARBAGE_ROWS:
        garbage_norm = normalize_text(garbage_phrase)
        if not garbage_norm:
            continue
        if abs(len(row_norm) - len(garbage_norm)) > 1:
            continue
        match_type = edit_distance_leq_one(row_norm, garbage_norm)
        if not match_type or match_type == "exact":
            continue
        candidate = {
            "garbage_phrase": garbage_phrase,
            "match_type": match_type,
        }
        if best_match is None:
            best_match = candidate
            continue
        current_rank = rank[candidate["match_type"]]
        best_rank = rank[best_match["match_type"]]
        if current_rank < best_rank:
            best_match = candidate
            continue
        if current_rank == best_rank and len(garbage_norm) > len(normalize_text(best_match["garbage_phrase"])):
            best_match = candidate

    return best_match


def has_result(value):
    """Treat None/blank strings as not-yet-classified."""
    return value is not None and str(value).strip() != ""


def write_local_match(ws, row, match):
    ws.cell(row=row, column=3, value=match["brand"])
    ws.cell(row=row, column=4, value=match["matched_text"])
    ws.cell(row=row, column=5, value=match["match_type"])


def clear_match_details(ws, row):
    for column in (3, 4, 5, 6, 7):
        ws.cell(row=row, column=column).value = None


def write_garbage_match(ws, row, match):
    ws.cell(row=row, column=6, value=match["garbage_phrase"])
    ws.cell(row=row, column=7, value=match["match_type"])


def clear_garbage_match(ws, row):
    for column in (6, 7):
        ws.cell(row=row, column=column).value = None


def is_maybe_product_row(name):
    """Preserve the general form-OR-quantity signal, guarded by structural words."""
    if MAYBE_PRODUCT_GUARD_PATTERN.search(str(name)):
        return False
    return bool(MAYBE_PRODUCT_FORM_PATTERN.search(str(name)) or MAYBE_PRODUCT_QTY_PATTERN.search(str(name)))


def classify_single_product(name, brand_index, master_evidence, garbage_rows):
    """Return (label, brand, matched text, match type, garbage phrase, garbage type).

    A confirmed master brand wins before structural garbage rules. Every
    unresolved row remains MAYBE_PRODUCT; no inferred product is promoted to 0.
    """
    name = "" if name is None else str(name).strip()
    spans = token_spans(name)
    match = exact_product_match(name, brand_index, precomputed_spans=spans)
    if match is None:
        match = fuzzy_product_match(name, brand_index, precomputed_spans=spans)
    if match is None and 'CORLION' in name.upper():
        from source_brand_aliases import resolve_source_brand_alias
        alias = resolve_source_brand_alias(name, (brand for brand, normalized in brand_index))
        if alias:
            match = {'brand': alias, 'matched_text': 'CORLION', 'match_type': 'source-alias'}
    if match is None and 'DIOF' in name.upper():
        description = parse_pro_diof_description(name)
        pro_brands = [brand for brand, normalized in brand_index if normalized == 'PRODIOF']
        if description is not None and not description.error and len(pro_brands) == 1:
            match = {'brand': pro_brands[0], 'matched_text': description.brand_text,
                     'match_type': 'master-pro-diof-form'}
        elif not match_pro_brand(name):
            diof = match_diof_brand(name)
            brands = [brand for brand, normalized in brand_index if normalized == 'DIOF']
            if diof and len(brands) == 1 and (master_evidence or {}).get('DIOF'):
                match = {'brand': brands[0], 'matched_text': diof.group(),
                         'match_type': 'master-diof-descriptor'}
    if match is None and 'MAXTRA' in name.upper():
        from maxtra import fused_brand_evidence
        entry = (master_evidence or {}).get('MAXTRA')
        if entry and fused_brand_evidence(name, entry['variants'], entry['products']):
            match = {'brand': entry['brand'], 'matched_text': 'MAXTRA',
                     'match_type': 'master-maxtra-descriptor'}
    if match is None:
        # TA is an existing tablet alias in the mapper. Preserve exact/fuzzy
        # precedence and use it only for otherwise unresolved FORTE evidence.
        recovery_name = re.sub(r'\bTA\b', 'TABLET', name, flags=re.I) if re.search(r'\bFORTE\b', name, re.I) else name
        match = master_product_match(recovery_name, master_evidence, precomputed_spans=spans)
    if match:
        # Brand-specific post-matching corrections.
        if match["brand"] == "ESLO":
            if "MET" in name.upper():
                match["brand"] = "ESLOMET"
        if match["brand"] == "DIOF":
            if match_pro_brand(name):
                # Only a complete PRO component (or confirmed PEO source alias)
                # establishes this family. PRODUCT/PROBIOTIC are not evidence.
                match["brand"] = "PRO DIOF"
                if "DUCT" not in name.upper():
                    pro_brands = [brand for brand, normalized in brand_index
                                  if normalized == "PRODIOF"]
                    if len(pro_brands) == 1:
                        match["brand"] = pro_brands[0]
        if match["brand"] == "NEW":
            if not re.search(r"\bNEW[\W_]*NORMET\b", name, re.IGNORECASE):
                # Retry the normal flow without the rejected brand for this row.
                return classify_single_product(
                    name,
                    [(brand, norm) for brand, norm in brand_index if brand != "NEW"],
                    {norm: entry for norm, entry in (master_evidence or {}).items()
                     if norm != "NEW"},
                    garbage_rows,
                )
        if match:
            return ("0", match["brand"], match["matched_text"], match["match_type"], None, None)
    if (is_exact_garbage_row(name, garbage_rows)
            or is_pattern_garbage_row(name)
            or is_metadata_garbage_row(name)):
        return ("1", None, None, None, None, None)
    garbage_match = fuzzy_garbage_match(name, garbage_rows)
    if garbage_match:
        return ("1", None, None, None, garbage_match["garbage_phrase"], garbage_match["match_type"])
    matched_text = "form+qty match" if is_maybe_product_row(name) else "no-brand-fallback"
    return ("MAYBE_PRODUCT", "pattern-inferred", matched_text, "MAYBE_PRODUCT", None, None)


def main(excel_file=None, master_file=None, sheet_name="garbage_check", clear_existing=None, row_limit=None):
    """Classify column A locally and write the inherited B:G result columns."""
    input_path = Path(excel_file) if excel_file is not None else Path(EXCEL_FILE)
    clear_existing = CLEAR_EXISTING if clear_existing is None else clear_existing
    row_limit = ROW_LIMIT if row_limit is None else row_limit
    brands = load_brands(master_file)
    brand_index = build_brand_index(brands)
    master_evidence = build_master_product_evidence(brand_index, master_file)
    garbage_rows = build_garbage_set()
    wb = load_workbook(input_path)
    counts = {"0": 0, "1": 0, "MAYBE_PRODUCT": 0}
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"{input_path} has no {sheet_name!r} worksheet")
        ws = wb[sheet_name]
        headers = ("Garbage check", "Matched brand", "Matched text", "Match type", "Matched garbage phrase", "Garbage match type")
        for column, header in enumerate(headers, start=2):
            ws.cell(row=1, column=column, value=header)
        if clear_existing:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
        processed = 0
        for row in range(2, ws.max_row + 1):
            value = ws.cell(row=row, column=1).value
            if value is None or str(value).strip() == "":
                continue
            if has_result(ws.cell(row=row, column=2).value):
                continue
            if row_limit is not None and processed >= row_limit:
                break
            result = classify_single_product(value, brand_index, master_evidence, garbage_rows)
            for column, output in enumerate(result, start=2):
                ws.cell(row=row, column=column).value = output
            counts[result[0]] += 1
            processed += 1
            print(f"row {row}: {str(value).strip()!r} -> {result[0]}")
        wb.save(input_path)
    finally:
        wb.close()
    print(f"{sum(counts.values())} rows checked locally: {counts['0']} product, {counts['1']} garbage, {counts['MAYBE_PRODUCT']} MAYBE_PRODUCT")
    return counts


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Classify Zuventus workbook rows using only local master evidence.")
    parser.add_argument("--input", type=Path, default=EXCEL_FILE)
    parser.add_argument("--master", type=Path, default=PRODUCT_MASTER_FILE)
    parser.add_argument("--sheet", default="garbage_check")
    parser.add_argument("--resume", action="store_true", help="Retain existing classification results.")
    parser.add_argument("--row-limit", type=int, default=None)
    args = parser.parse_args()
    if args.row_limit is not None and args.row_limit < 0:
        parser.error("--row-limit must be nonnegative")
    main(args.input, args.master, args.sheet, not args.resume, args.row_limit)
