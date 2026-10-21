"""Local Zuventus mapping with master-derived brands, variants, and diagnostics."""
import re
import sys
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from rapidfuzz import fuzz, process
from rapidfuzz.distance import Levenshtein, JaroWinkler
MASTER_XLSX_PATH = Path(__file__).resolve().with_name('zuventus_product.xlsx')
INPUT_OUTPUT_XLSX_PATH = Path(__file__).resolve().with_name('mapping.xlsx')
ROW_LIMIT = None
PROCESS_MAYBE_PRODUCT_ONLY = False
PROCESS_CONFIRMED_ZERO_ONLY = True
if PROCESS_CONFIRMED_ZERO_ONLY and PROCESS_MAYBE_PRODUCT_ONLY:
    raise ValueError('PROCESS_CONFIRMED_ZERO_ONLY and PROCESS_MAYBE_PRODUCT_ONLY are mutually exclusive; set at most one of them to True.')
PROCESSING_MODE = 'local'
INPUT_COLUMN = 'Input_Column'
OUTPUT_COLUMN = 'Output_Column'
PRODUCT_CODE_COLUMN = 'Product_Code'
STATUS_COLUMN = 'Match_Status'
CONFIDENCE_COLUMN = 'Confidence'
CANDIDATES_COLUMN = 'Candidates_To_Model'
SUGGESTIONS_COLUMN = 'Suggestions'
SOURCE_COLUMN = 'Source'
PRODUCT_COL = 'product_name'
BRAND_COL = 'BRAND_NAME'
PACK_SIZE_COL = 'PACK_SIZE'
VARIANT_COL = 'VARIANT'
SUB_VARIANT_COL = 'SUB_VARIANT'
PRODUCT_CODE_COL = 'PRODUCT_CODE'
FUZZY_MIN_SCORE = 80
BRAND_NAME_NEXTSTEP_SCORE = 99
JW_MIN_SCORE = 0.85
BRAND_FUZZY_MAX_EDITS = 2
DOSAGE_UNIT_SUFFIXES = {'MG', 'GM', 'ML', 'MCG', 'IU', 'UNIT'}
FORM_HINT_TOKENS = {'BARS', 'SYR', 'TABLETS', 'INJ', 'CAP', 'GRANULES', 'CAPSULE', 'PFS', 'INJECTION', 'CREAM', 'AMPOULE', 'CAPSULES', 'LOTION', 'POWDER', 'TABLET', 'CAPS', 'PEN', 'VIAL', 'AMPS', 'IN', 'GEL', 'GRANUL', 'TAB', 'TONIC', 'AMPOULES', 'BAR', 'SYP', 'OINTMENT', 'SOLUTION', 'SYRUP', 'TABS', 'SACHET', 'SUSP', 'DROPS', 'TA', 'AMP', 'DROP', 'SUSPENSION'}
VARIANT_TOKEN_EQUIVALENTS = {'TA': 'TABLET', 'TAB': 'TABLET', 'TABS': 'TABLET', 'TABLETS': 'TABLET', 'CAP': 'CAPSULE', 'CAPS': 'CAPSULE', 'CAPSULES': 'CAPSULE', 'INJ': 'INJECTION', 'INJECTIONS': 'INJECTION', 'DROP': 'DROPS', 'SYP': 'SYRUP', 'BARS': 'BAR', 'GRANUL': 'GRANULES', 'LOT': 'LOTION', 'LO': 'LOTION', 'CRM': 'CREAM', 'CR': 'CREAM'}
INPUT_TOKEN_EQUIVALENTS = {'ING': 'INJECTION', 'INJ': 'INJECTION', 'INJN': 'INJECTION', 'INJN.': 'INJECTION', 'INJ.': 'INJECTION', 'SYR': 'SYRUP', 'SYP': 'SYRUP', 'SUSP': 'SUSPENSION', 'SUSPN': 'SUSPENSION', 'AMP': 'INJECTION', 'AMPS': 'INJECTION', 'AMPUL': 'INJECTION', 'AMPULE': 'INJECTION', 'AMPULES': 'INJECTION', 'AMPOULE': 'INJECTION', 'AMPOULES': 'INJECTION', 'VIAL': 'VIAL', 'VL': 'VIAL', 'PFS': 'PFS', 'PEN': 'PEN', 'TAB': 'TABLET', 'TABS': 'TABLET', 'TA': 'TABLET', 'TABLETS': 'TABLET', 'CAP': 'CAPSULE', 'CAPS': 'CAPSULE', 'CAPSULE': 'CAPSULE', 'CAPSULES': 'CAPSULE', 'DROP': 'DROPS', 'DROPS': 'DROPS', 'INJECTION': 'INJECTION', 'SYRUP': 'SYRUP', 'SUSPENSION': 'SUSPENSION', 'TABLETSNU': 'TABLETS', 'AMPOULEP': 'INJECTION', 'TABLETP': 'TABLET', 'AMPOUL': 'INJECTION', 'AMPOULS': 'INJECTION', 'INJG': 'INJECTION', 'INJE': 'INJECTION', 'NEDDLES': 'NEEDLES', 'NEEDALES': 'NEEDLES', 'PAKET': 'PACK', 'PACKET': 'PACK', 'VAIL': 'VIAL', 'VALE': 'VIAL', 'INJD': 'INJECTION', 'INGD': 'INJECTION', 'SY': 'SYRUP', 'TABD': 'TABLET', 'CAPD': 'CAPSULE', 'CAPSUL': 'CAPSULE', 'CAPSULS': 'CAPSULE', 'CREA': 'CREAM', 'OINT': 'OINTMENT', 'OINTB': 'OINTMENT', 'POW': 'POWDER', 'GRANUL': 'GRANULES', 'GRANU': 'GRANULES', 'GRANULS': 'GRANULES', 'LOT': 'LOTION', 'LOTION': 'LOTION', 'CRM': 'CREAM', 'CR': 'CREAM', 'CRM.': 'CREAM', 'CR.': 'CREAM', 'OINTM': 'OINTMENT', 'OINTMENT': 'OINTMENT', 'BAR': 'BAR', 'BARS': 'BAR', 'TABLETNU': 'TABLET', 'TABSP': 'TABLET', 'CAPSULP': 'CAPSULE', 'AMPOULP': 'INJECTION', 'DROPSP': 'DROPS', 'IN': 'INJECTION', 'GRANULIS': 'GRANULES', 'CHEB': 'CHEWABLE', 'F/WASH': 'FACE WASH', 'FACEWASH': 'FACE WASH', 'SYNDEN': 'SYNDET', 'MOISTURIS': 'MOISTURIZING', 'HYD': 'HYDRATING', 'CLEN': 'CLEAN', 'FLV': 'FLAVOR', 'FLAV': 'FLAVOR', 'FLAVOUR': 'FLAVOR'}
GENERIC_VARIANT_MATCH_TOKENS = {'TABLET', 'CAPSULE', 'INJECTION', 'SYRUP', 'DROPS', 'CREAM', 'LOTION', 'GEL', 'OINTMENT', 'SOLUTION', 'SUSPENSION', 'POWDER', 'SACHET', 'BAR', 'GRANULES'}

def norm(s: str) -> str:
    s = str(s or '').upper().strip()
    s = s.replace('+', ' PLUS ')
    s = s.replace('-', ' ')
    s = s.replace('/', ' ')
    s = s.replace('*', ' ')
    s = re.sub('\\bX\\b', ' ', s)
    s = s.replace('&', ' AND ')
    s = re.sub('\\b([A-Z])\\.([A-Z])\\b', '\\1\\2', s)
    s = re.sub('\\b([A-Z])\\.([A-Z])\\.([A-Z])\\b', '\\1\\2\\3', s)
    s = re.sub('[^A-Z0-9]+', ' ', s)
    s = re.sub('\\s+', ' ', s).strip()
    return s

def normalize_variant_token(s: str) -> str:
    s = str(s or '').upper().replace('+', ' PLUS ').strip()
    s = re.sub('[^A-Z0-9]+', ' ', s)
    return re.sub('\\s+', ' ', s).strip()

def compact_variant_token(s: str) -> str:
    return re.sub('[^A-Z0-9]+', '', str(s or '').upper().replace('+', ' PLUS '))

def compact_brand_token(s: str) -> str:
    return re.sub('[^A-Z0-9]+', '', str(s or '').upper())

def _expand_strength_token(tok: str) -> str:
    m = re.fullmatch('(\\d+(?:\\.\\d+)?)([KM])', str(tok or '').upper())
    if not m:
        return tok
    mult = 1000 if m.group(2) == 'K' else 1000000
    return str(int(float(m.group(1)) * mult))

def _normalize_strength_words(words) -> set:
    return {_expand_strength_token(w) for w in words}

def _input_compact_joins(input_text) -> set:
    toks = re.findall('[A-Z0-9]+', str(input_text or '').upper())
    joins = set()
    for i in range(len(toks) - 1):
        joins.add(toks[i] + toks[i + 1])
    for i in range(len(toks) - 2):
        joins.add(toks[i] + toks[i + 1] + toks[i + 2])
    return joins

def _expand_strength_in_compact(s: str) -> str:
    return re.sub('(\\d+(?:\\.\\d+)?)([KM])(?=\\d|$)', lambda m: str(int(float(m.group(1)) * (1000 if m.group(2) == 'K' else 1000000))), str(s or '').upper())

def canonical_variant_match_token(token: str) -> str:
    token = str(token or '').upper().strip()
    return VARIANT_TOKEN_EQUIVALENTS.get(token, token)

def extract_variant_match_tokens(text: str) -> set:
    normalized = normalize_ocr_numeric_noise(str(text or '').upper())
    tokens = set()
    for token in re.findall('[A-Z0-9]+', normalized):
        if token.isdigit():
            continue
        if re.fullmatch('\\d+(?:\\.\\d+)?', token):
            continue
        if re.fullmatch('\\d+[X*]\\d+(?:MG|GM|ML|MCG|IU|UNIT|G|L|KG)?', token):
            continue
        if re.fullmatch('\\d+(?:MG|GM|ML|MCG|IU|UNIT|G|L|KG)', token):
            continue
        if re.fullmatch('\\d+[STCP]', token):
            continue
        if token in {'X', 'PACK', 'OF'}:
            continue
        token = canonical_variant_match_token(token)
        if token and token not in DOSAGE_UNIT_SUFFIXES:
            tokens.add(token)
    return tokens

def build_flexible_brand_pattern(brand_name: str) -> str:
    brand_compact = compact_brand_token(brand_name)
    if not brand_compact:
        return ''
    separator = '[\\s\\-]*'
    char_pattern = separator.join((re.escape(ch) for ch in brand_compact))
    return '(?<![A-Z0-9])' + char_pattern + '(?![A-Z0-9])'

def get_brand_removal_name(brand_name: str, available_variants: set=None) -> str:
    tokens = re.findall('[A-Z0-9]+', str(brand_name or '').upper().replace('+', ' PLUS '))
    if len(tokens) <= 1:
        return str(brand_name or '').upper().strip()
    variant_compacts = set()
    if available_variants:
        variant_compacts = {compact_variant_token(variant) for variant in available_variants if variant and str(variant).upper() != 'PLAIN'}
    removable_tokens = [tokens[0]]
    suffix_tokens = tokens[1:]
    i = 0
    while i < len(suffix_tokens):
        token = suffix_tokens[i]
        two_token_compact = ''
        if i + 1 < len(suffix_tokens):
            two_token_compact = compact_variant_token(token + ' ' + suffix_tokens[i + 1])
        preserve_suffix = token in FORM_HINT_TOKENS or compact_variant_token(token) in variant_compacts or (two_token_compact and two_token_compact in variant_compacts)
        if not preserve_suffix:
            removable_tokens.append(token)
        i += 1
    return ' '.join(removable_tokens).strip()

def strip_brand_from_input(input_text: str, brand_name: str, available_variants: set=None) -> str:
    if not brand_name:
        return input_text
    input_upper = str(input_text or '').upper().strip()
    removal_brand = get_brand_removal_name(brand_name, available_variants=available_variants)
    brand_normalized = removal_brand.upper().replace('-', ' ').strip()
    brand_pattern = '\\b' + re.escape(brand_normalized).replace('\\ ', '[-\\s]+') + '\\b'
    stripped = re.sub(brand_pattern, '', input_upper, count=1, flags=re.IGNORECASE).strip()
    stripped = re.sub('\\b' + re.escape(brand_normalized) + '\\b', '', stripped, flags=re.IGNORECASE).strip()
    if stripped != input_upper:
        return re.sub('\\s+', ' ', stripped).strip()
    flexible_pattern = build_flexible_brand_pattern(removal_brand)
    if flexible_pattern:
        stripped = re.sub(flexible_pattern, '', input_upper, count=1, flags=re.IGNORECASE).strip()
        if stripped != input_upper:
            return re.sub('\\s+', ' ', stripped).strip()
    return re.sub('\\s+', ' ', input_upper).strip()

def normalize_master_text(value, default: str='') -> str:
    if pd.isna(value):
        return default
    value = str(value).strip()
    if value in ('nan', 'NaN', 'None'):
        return default
    return value

def normalize_master_code(value) -> str:
    if pd.isna(value):
        return ''
    if isinstance(value, (int, float)) and (not isinstance(value, bool)):
        if float(value).is_integer():
            return str(int(value))
    return normalize_master_text(value, default='')

def normalize_master_attr(value, default: str='PLAIN') -> str:
    normalized = normalize_master_text(value, default='').upper()
    return normalized if normalized else default

def alpha_signature(value: str) -> str:
    letters = re.sub('[^A-Z]', '', str(value or '').upper())
    consonants = re.sub('[AEIOU]', '', letters)
    return consonants or letters

def collapse_spaced_thousands(s: str) -> str:
    prev = None
    out = str(s or '')
    while prev != out:
        prev = out
        out = re.sub('\\b(\\d{1,3})\\s+(\\d{3})\\b', '\\1\\2', out)
    return out

def numeric_strength_variant_value(variant: str) -> str:
    text = collapse_spaced_thousands(str(variant or '').upper().strip())
    text = re.sub('[\\s\\-]*(?:PFS|IU|UNIT|MG|MCG|ML|GM)\\.?$', '', text).strip()
    return text if re.fullmatch('\\d+(?:\\.\\d+)?', text) else ''

def normalize_ocr_numeric_noise(value: str) -> str:
    text = str(value or '').upper().replace('+', ' PLUS ')
    for source_token, target_token in INPUT_TOKEN_EQUIVALENTS.items():
        text = re.sub('\\b' + re.escape(source_token) + '\\b', target_token, text)
    text = re.sub('\\b(\\d+(?:\\.\\d+)?)\\s*K\\b', lambda m: str(int(float(m.group(1)) * 1000)) if float(m.group(1)) * 1000 == int(float(m.group(1)) * 1000) else str(float(m.group(1)) * 1000), text)
    text = re.sub('(?<=\\d)[/\\\\\\-](?=\\d)', '/', text)
    text = re.sub('\\b(\\d+)(10|15|20|30|14|28)S\\b', lambda m: f'{m.group(1)} {m.group(2)}S' if int(m.group(1) + m.group(2)) > 100 else m.group(0), text)
    return re.sub('\\s+', ' ', text).strip()

def extract_numeric_strength(value: str) -> str:
    text = normalize_master_attr(value, default='')
    if not text:
        return ''
    cleaned = re.sub('\\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT|TAB|TABS|TABLET|TABLETS|INJ|INJECTION)\\s*$', '', text).strip()
    if re.fullmatch('\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)*', cleaned):
        return cleaned
    return ''

def is_numeric_subvariant_like(value: str) -> bool:
    return bool(re.fullmatch('\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)*', str(value or '').strip().upper()))
_ORAL_FORM_TOKENS = {'TABLET', 'TABLETS', 'TAB', 'TABS', 'CAPSULE', 'CAPSULES', 'CAP', 'CAPS', 'SYRUP', 'SYP', 'DROPS', 'DROP', 'SUSPENSION', 'SUSP', 'CREAM', 'GEL', 'OINTMENT', 'LOTION', 'SACHET', 'POWDER', 'POWD', 'GRANULES', 'GRANULE', 'SPRAY', 'NASALSPRAY', 'SOLUTION', 'SOL', 'TONIC', 'BAR', 'KIT'}

def repair_master_variant_fields(product_name: str, variant: str, sub_variant: str) -> tuple:
    variant = normalize_master_attr(variant)
    sub_variant = normalize_master_attr(sub_variant)
    _vt = [t for t in variant.split() if t.upper() not in _ORAL_FORM_TOKENS]
    variant = ' '.join(_vt).strip() or 'PLAIN'
    numeric_pattern = '\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)*'
    numeric_with_unit_pattern = numeric_pattern + '(?:\\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT))?'
    noisy_sub_variant_tokens = {'MG', 'GM', 'ML', 'MCG', 'G', 'M', 'IU', 'UNIT', 'TAB', 'TABS', 'TABLET', 'TABLETS', 'INJ', 'INJECTION', 'PFS', 'VIAL', 'AMP', 'AMPS', 'SYR', 'SYRINGE', 'PEN'}
    if sub_variant == 'PLAIN':
        collapsed_variant = collapse_spaced_thousands(variant)
        form_dose = re.fullmatch('(\\d+(?:\\.\\d+)?)\\s*-?\\s*(?:PFS|VIAL|AMP|AMPS|SYR|SYRINGE|PEN|IU|INJ|INJECTION)', collapsed_variant)
        if form_dose:
            return ('PLAIN', form_dose.group(1))
    variant_numeric = extract_numeric_strength(variant)
    if sub_variant in noisy_sub_variant_tokens and variant_numeric:
        return ('PLAIN', variant_numeric)
    if re.fullmatch(numeric_pattern, variant) and re.search('[A-Z]', sub_variant) and (not is_numeric_subvariant_like(sub_variant)) and (sub_variant not in noisy_sub_variant_tokens):
        return (normalize_master_attr(sub_variant), variant)
    if sub_variant != 'PLAIN':
        return (variant, sub_variant)
    if re.fullmatch(numeric_pattern, variant):
        return ('PLAIN', variant)
    pure_numeric_with_unit_match = re.fullmatch('(' + numeric_pattern + ')\\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT)', variant)
    if pure_numeric_with_unit_match:
        return ('PLAIN', pure_numeric_with_unit_match.group(1))
    split_match = re.fullmatch('(.+?)[\\s\\-]+(' + numeric_pattern + ')', variant)
    if split_match:
        repaired_variant = split_match.group(1).strip()
        repaired_sub_variant = split_match.group(2).strip()
        if repaired_variant:
            return (repaired_variant, repaired_sub_variant)
    compact_combo_match = re.fullmatch('([A-Z]+(?:\\s+[A-Z]+)*?)(' + numeric_pattern + ')', variant)
    if compact_combo_match:
        repaired_variant = compact_combo_match.group(1).strip()
        repaired_sub_variant = compact_combo_match.group(2).strip()
        if repaired_variant and repaired_sub_variant and ('/' in repaired_sub_variant or len(repaired_sub_variant.replace('.', '')) >= 2):
            return (repaired_variant, repaired_sub_variant)
    reverse_split_match = re.fullmatch('(' + numeric_pattern + ')\\s+([A-Z]+(?:\\s+[A-Z0-9]+)*)', variant)
    if reverse_split_match:
        repaired_sub_variant = reverse_split_match.group(1).strip()
        repaired_variant = reverse_split_match.group(2).strip()
        if repaired_variant and repaired_variant not in noisy_sub_variant_tokens:
            return (repaired_variant, repaired_sub_variant)
    unit_split_match = re.fullmatch('(.+?)[\\s\\-]+(' + numeric_with_unit_pattern + ')', variant)
    if unit_split_match:
        repaired_variant = unit_split_match.group(1).strip()
        repaired_sub_variant = unit_split_match.group(2).strip()
        repaired_sub_variant = re.sub('\\s*(?:MG|GM|ML|MCG|G|M|IU|UNIT)\\s*$', '', repaired_sub_variant).strip()
        if repaired_variant and repaired_sub_variant:
            return (repaired_variant, repaired_sub_variant)
    return (variant, sub_variant)

def is_reasonable_pack_count(value) -> bool:
    try:
        pack_num = float(str(value).strip())
    except (TypeError, ValueError):
        return False
    return pack_num.is_integer() and 0 < pack_num <= 100

def find_tablike_pack_matches(input_upper: str) -> list:
    matches = []
    pattern = '(?<![\\d.])(\\d+)\\s*(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\\b'
    for m in re.finditer(pattern, input_upper):
        if is_reasonable_pack_count(m.group(1)):
            matches.append(m)
    return matches

def parse_pack_size_value(pack: str):
    pack_upper = str(pack or '').upper().strip()
    if not pack_upper or pack_upper in ('NAN', 'NONE'):
        return None
    combo_match = re.search('(\\d+)\\s*[X*]\\s*(\\d+)', pack_upper)
    if combo_match:
        return float(int(combo_match.group(1)) * int(combo_match.group(2)))
    numeric_match = re.search('(\\d+(?:\\.\\d+)?)', pack_upper)
    if numeric_match:
        return float(numeric_match.group(1))
    return None

def extract_combination_dosage(input_name: str) -> str:
    input_name = normalize_ocr_numeric_noise(input_name)
    input_name = re.sub('\\s*/\\s*', '/', input_name)
    pattern = '(?<![\\d./])(\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)+)(?![\\d./])'
    for match in re.finditer(pattern, input_name):
        candidate = match.group(1)
        if re.search('(?<!\\d)\\d+[X*×]', candidate):
            continue
        after = input_name[match.end():match.end() + 2]
        if re.match('[X*×]', after):
            continue
        components = candidate.split('/')
        for comp in components:
            comp_pos = input_name.find(comp, match.start())
            if comp_pos >= 0:
                after_comp = input_name[comp_pos + len(comp):comp_pos + len(comp) + 2]
                if re.match('[X*×]', after_comp):
                    continue
        return candidate
    return ''

def clean_duplicate_words(input_name: str) -> str:
    words = input_name.split()
    cleaned = []
    for i, word in enumerate(words):
        if i == 0 or word.upper() != words[i - 1].upper():
            cleaned.append(word)
    result = cleaned
    for n in (2, 3):
        i = 0
        deduped = []
        while i < len(result):
            chunk = result[i:i + n]
            if len(chunk) == n and result[i + n:i + 2 * n] == chunk:
                deduped.extend(chunk)
                i += 2 * n
            else:
                deduped.append(result[i])
                i += 1
        result = deduped
    return ' '.join(result)

def extract_brand_like_query(input_name: str) -> str:
    text = normalize_ocr_numeric_noise(input_name)
    tokens = re.findall('[A-Z0-9.]+', text)
    if not tokens:
        return ''
    skip_tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    skip_tokens |= {'PACK', 'OF', 'PCS', 'PC', 'NOS', 'NO', 'VIAL', 'AMP', 'AMPS', 'PFS', 'SYR'}
    brand_tokens = []
    for token in tokens:
        if token in skip_tokens:
            break
        if re.search('\\d', token):
            glued = re.match('^([A-Z]{2,})\\d', token)
            if glued and (not brand_tokens):
                brand_tokens.append(glued.group(1))
            break
        brand_tokens.append(token)
    return ' '.join(brand_tokens).strip()

def get_brand_variant_tags(brand_map: dict, brand_hint: str) -> set:
    if not brand_hint or not brand_map:
        return set()
    hint = str(brand_hint).strip().upper()
    key = hint if hint in brand_map else next((k for k in brand_map if k.upper() == hint), None)
    if key is None:
        return set()
    tags = set()
    for item in brand_map[key]:
        variant = str(item.get('variant', '') or '').strip().upper()
        if variant and variant != 'PLAIN' and re.fullmatch('[A-Z]{1,4}', variant):
            tags.add(variant)
    return tags

def strip_parenthetical_noise(s: str, brand_variant_tags: set=None) -> str:
    if not s:
        return s
    text = str(s)
    tag_pattern = '\\(\\s*([A-Z]{1,4})\\s*\\)'
    outside_tokens = None
    if brand_variant_tags:
        outside_tokens = set(re.findall('[A-Z0-9]+', re.sub(tag_pattern, ' ', text, flags=re.IGNORECASE).upper()))

    def _unwrap_or_drop(m):
        tag = m.group(1).strip().upper()
        if brand_variant_tags and tag in brand_variant_tags and (tag not in outside_tokens):
            return ' ' + tag + ' '
        return ' '
    cleaned = re.sub(tag_pattern, _unwrap_or_drop, text, flags=re.IGNORECASE)
    return re.sub('\\s+', ' ', cleaned).strip()

def extract_brand_suffix_tokens(input_name: str, brand_name: str) -> list:
    if not brand_name:
        return []
    input_upper = normalize_ocr_numeric_noise(input_name)
    brand_upper = brand_name.upper().strip()
    tokens = []
    skip_tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    skip_tokens |= {'PACK', 'OF', 'PCS', 'PC', 'NOS', 'NO', 'VIAL', 'AMP', 'AMPS', 'PFS', 'SYR', 'POWDE', 'POWDER', 'GRANUL', 'GRANULE', 'GRANULES', 'GRANULS'}

    def should_keep_token(token: str) -> bool:
        token_letters = re.sub('[^A-Z]', '', token.upper())
        return bool(token_letters) and token_letters not in skip_tokens

    def is_pack_notation(token: str) -> bool:
        return bool(re.match('^\\d+[Xx*×]\\d+', token) or re.match('^\\d+[Xx]\\d+', token))
    pattern1 = re.escape(brand_upper) + '-([A-Z]+\\d+|\\d+[A-Z]+)\\b'
    for m in re.findall(pattern1, input_upper):
        if should_keep_token(m) and (not is_pack_notation(m)):
            tokens.append(m)
    pattern2 = re.escape(brand_upper) + '\\s+([A-Z]+\\d+|\\d+[A-Z]+)\\b'
    for m in re.findall(pattern2, input_upper):
        if should_keep_token(m) and (not is_pack_notation(m)) and (m not in tokens):
            tokens.append(m)
    bm = re.search('\\b' + re.escape(brand_upper) + '\\b', input_upper)
    if bm:
        tail_match = re.match('[\\s\\-]*([A-Z]{2,})\\b', input_upper[bm.end():])
        if tail_match:
            alpha_token = tail_match.group(1)
            if alpha_token not in skip_tokens and alpha_token not in tokens:
                tokens.append(alpha_token)
    return tokens

def filter_items_by_compound_tokens(items: list, compound_tokens: list, verbose=False) -> list:
    if not compound_tokens:
        return items
    if verbose:
        print(f'\n  [COMPOUND TOKEN] Tokens to match: {compound_tokens}')
        print(f'  [COMPOUND TOKEN] Products before filter: {len(items)}')
    normalized_tokens = {compact_variant_token(token) for token in compound_tokens if token}
    matched = []
    for item in items:
        prod_upper = item['product'].upper().replace('-', ' ').replace('_', ' ')
        prod_compact = compact_variant_token(prod_upper)
        product_tokens = re.findall('[A-Z0-9]+', prod_upper)
        product_compact_tokens = {compact_variant_token(tok) for tok in product_tokens if tok}
        for i in range(len(product_tokens) - 1):
            product_compact_tokens.add(compact_variant_token(product_tokens[i] + product_tokens[i + 1]))
        for i in range(len(product_tokens) - 2):
            product_compact_tokens.add(compact_variant_token(product_tokens[i] + product_tokens[i + 1] + product_tokens[i + 2]))
        has_match = False
        for token in compound_tokens:
            normalized = compact_variant_token(token)
            if not token:
                continue
            if re.fullmatch('(?:[A-Z]+\\d+|\\d+[A-Z]+)', token):
                if token in product_tokens or normalized in product_compact_tokens:
                    has_match = True
                    break
                continue
            if token in prod_upper or (normalized and normalized in prod_compact):
                has_match = True
                break
        if has_match:
            matched.append(item)
            if verbose:
                print(f"    [OK] Matched '{item['product']}' on token(s) {compound_tokens}")
    if matched:
        if verbose:
            print(f'  [COMPOUND TOKEN] Products after filter: {len(matched)}')
        return matched
    if verbose:
        print(f'  [COMPOUND TOKEN] No products matched tokens -- keeping all {len(items)} products')
    return items

def get_compound_token_digits(compound_tokens: list) -> set:
    digits = set()
    for token in compound_tokens:
        for d in re.findall('\\d+', token):
            digits.add(d)
    return digits

def get_all_pack_positions(input_upper: str) -> list:
    positions = []
    for m in re.finditer('\\b(\\d+)\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)(?:\\s*(?:ML|GM|MG|MCG|IU|G|L|KG|TA|TABS?|TABLETS?|CAPS?|CAPSULES?|STRIPS?))?', input_upper):
        positions.append((m.start(), m.end()))
    for m in re.finditer('\\bX\\s*(\\d+)\\b', input_upper):
        positions.append((m.start(), m.end()))
    for m in re.finditer('\\b(\\d+)([STCP])\\b', input_upper):
        if m.group(2) in ['S', 'T', 'C', 'P']:
            positions.append((m.start(), m.end()))
    for m in re.finditer('\\b(\\d+)\\s+(?:S|T|PCS)\\b', input_upper):
        positions.append((m.start(), m.end()))
    tablike_matches = find_tablike_pack_matches(input_upper)
    if tablike_matches and (not positions):
        last_match = tablike_matches[-1]
        positions.append((last_match.start(), last_match.end()))
    return positions

def is_position_in_pack_pattern(match_start: int, match_end: int, pack_positions: list) -> bool:
    for pack_start, pack_end in pack_positions:
        if pack_start <= match_start < pack_end or pack_start < match_end <= pack_end:
            return True
    return False

def split_strength_and_variant(input_name: str, available_variants: set, verbose=False) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    input_upper = input_name.upper()
    pack_positions = get_all_pack_positions(input_upper)
    pattern = '\\b(\\d+(?:\\.\\d+)?)([A-Z]+)\\b'
    for match in re.finditer(pattern, input_upper):
        number_part = match.group(1)
        letter_part = match.group(2)
        found_letter_variant = False
        found_full_letter_variant = False
        is_dosage_suffix = letter_part in DOSAGE_UNIT_SUFFIXES
        if is_position_in_pack_pattern(match.start(), match.end(), pack_positions):
            continue
        if not is_dosage_suffix:
            if all((letter in available_variants for letter in letter_part)):
                for i, letter in enumerate(letter_part):
                    if letter in available_variants:
                        found_letter_variant = True
                        detected_variants.add(letter)
                        if i == 0:
                            detected_sub_variants.add(number_part)
            if letter_part in available_variants:
                found_full_letter_variant = True
                detected_variants.add(letter_part)
                detected_sub_variants.add(number_part)
        if number_part in available_variants and (not found_letter_variant) and (not found_full_letter_variant):
            detected_variants.add(number_part)
    return (detected_variants, detected_sub_variants)

def detect_compact_variant_strength_combo(input_name: str, available_variants: set) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    tokens = re.findall('[A-Z0-9]+', normalize_ocr_numeric_noise(input_name))
    skip_tokens = _pack_form_skip_tokens() | {'PACK', 'OF'}
    for i in range(len(tokens) - 1):
        compact_token = tokens[i]
        next_token = tokens[i + 1]
        m = re.fullmatch('([A-Z]+)(\\d+(?:\\.\\d+)?)', compact_token)
        if not m:
            continue
        letter_part, numeric_part = m.groups()
        if letter_part in available_variants:
            detected_variants.add(letter_part)
            detected_sub_variants.add(numeric_part)
        if next_token in skip_tokens:
            continue
        candidate_variant = f'{letter_part} {next_token}'
        if candidate_variant in available_variants:
            detected_variants.add(candidate_variant)
            detected_sub_variants.add(numeric_part)
    return (detected_variants, detected_sub_variants)

def detect_variant_numeric_pairs(input_name: str, available_variants: set) -> tuple:
    detected_variants = set()
    detected_sub_variants = set()
    normalized = normalize_ocr_numeric_noise(input_name)
    seen = set()
    patterns = ['\\b([A-Z]+)(\\d+(?:\\.\\d+)?)\\b', '\\b([A-Z]+)[\\-\\s]+(\\d+(?:\\.\\d+)?)\\b']
    for pattern in patterns:
        for m in re.finditer(pattern, normalized):
            letter_part, numeric_part = m.groups()
            if letter_part in DOSAGE_UNIT_SUFFIXES:
                continue
            key = (letter_part, numeric_part)
            if key in seen:
                continue
            seen.add(key)
            if letter_part in available_variants:
                detected_variants.add(letter_part)
                detected_sub_variants.add(numeric_part)
    return (detected_variants, detected_sub_variants)

def product_contains_combination(product_name: str, combination: str) -> bool:
    product_norm = product_name.upper().replace(' ', '').replace('-', '')
    combo_norm = combination.upper().replace(' ', '').replace('-', '')
    if not combo_norm:
        return False
    pat = '(?<![\\d.])' + re.escape(combo_norm) + '(?![\\d.])'
    return bool(re.search(pat, product_norm))

def extract_available_attributes_from_items(items: list) -> tuple:
    available_variants = {'PLAIN'}
    available_sub_variants = {'PLAIN'}
    available_pack_sizes = set()
    for item in items:
        variant = str(item.get('variant', 'PLAIN')).strip().upper()
        sub_variant = str(item.get('sub_variant', 'PLAIN')).strip().upper()
        pack_size = str(item.get('pack_size', '')).strip()
        if variant and variant not in ('', 'NAN', 'NONE'):
            available_variants.add(variant)
        if sub_variant and sub_variant not in ('', 'NAN', 'NONE'):
            available_sub_variants.add(sub_variant)
        if pack_size and pack_size not in ('', 'nan', 'NaN', 'None'):
            available_pack_sizes.add(pack_size)
    return (available_variants, available_sub_variants, available_pack_sizes)

def extract_pack_size_from_input(s: str) -> str:
    s_upper = normalize_ocr_numeric_noise(s)
    triple_nxmxp = re.search('(?<![\\.\\d])(\\d+(?:\\.\\d+)?)\\s*[X*×]\\s*(\\d+)\\s*[X*×]\\s*(15)\\b', s_upper)
    if triple_nxmxp:
        return '15'
    nxm_matches = list(re.finditer('\\b(\\d+)\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)\\s*([STCP])?\\b', s_upper))
    for m in nxm_matches:
        n1 = float(m.group(1))
        n2 = float(m.group(2))
        total = n1 * n2
        if n1 <= 10 and n2 <= 100 and (total <= 500):
            return str(int(total) if total.is_integer() else total)
    full_nx1xm = re.search('\\b(\\d+)\\s*[X*×]\\s*1\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)\\s*([STCP])?\\b', s_upper)
    if full_nx1xm:
        n1 = float(full_nx1xm.group(1))
        n2 = float(full_nx1xm.group(2))
        total = 1 * n2
        return str(int(total) if total.is_integer() else total)
    if len(nxm_matches) > 1:
        pass
    m = re.search('\\b(\\d+)\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)(?:\\s*(?:ML|GM|MG|MCG|IU|G|L|KG|TA|TABS?|TABLETS?|CAPS?|CAPSULES?|T|C))?\\b', s_upper)
    if m:
        n1 = float(m.group(1))
        n2 = float(m.group(2))
        total = n1 * n2
        if n1 <= 10 and n2 <= 100 and (total <= 500):
            return str(int(total) if total.is_integer() else total)
    m = re.search('\\bPACK\\s*(?:OF\\s*)?(\\d+)\\b', s_upper)
    if m:
        return m.group(1)
    m = re.search('(?<!\\d)(\\d+)(?=[STCP])', s_upper)
    if m:
        return m.group(1)
    m = re.search('\\b(\\d+)\\s+(?:S|T|PCS)\\b', s_upper)
    if m:
        return m.group(1)
    adjacent_form_matches = []
    for m in re.finditer('(?<![\\d.])(\\d+)\\s*(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\\b', s_upper):
        if is_reasonable_pack_count(m.group(1)):
            adjacent_form_matches.append((m.start(), m.group(1)))
    for m in re.finditer('\\b(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\\s*(\\d+)\\b', s_upper):
        if is_reasonable_pack_count(m.group(1)):
            adjacent_form_matches.append((m.start(), m.group(1)))
    if adjacent_form_matches:
        adjacent_form_matches.sort(key=lambda x: x[0], reverse=True)
        return adjacent_form_matches[0][1]
    m = re.search('\\b(\\d+(?:\\.\\d+)?)\\s*(?:ML|GM|G|L|KG)\\b', s_upper)
    if m:
        return m.group(1)
    return ''

def load_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, engine='openpyxl', dtype=object)
    for required in (PRODUCT_COL, BRAND_COL):
        if required not in df.columns:
            raise ValueError(f"Missing column '{required}'. Found: {list(df.columns)}")
    cols = [PRODUCT_COL, BRAND_COL]
    if PACK_SIZE_COL in df.columns:
        cols.append(PACK_SIZE_COL)
        print(f"[OK] Found '{PACK_SIZE_COL}' column in master - pack-size matching ENABLED.")
    else:
        print(f"[WARN] '{PACK_SIZE_COL}' column NOT found in master - pack sizes will be ignored.")
    if VARIANT_COL in df.columns:
        cols.append(VARIANT_COL)
        print(f"[OK] Found '{VARIANT_COL}' column in master - variant-based filtering ENABLED.")
    else:
        print(f"[WARN] '{VARIANT_COL}' column NOT found in master - all products will be shown.")
    if SUB_VARIANT_COL in df.columns:
        cols.append(SUB_VARIANT_COL)
        print(f"[OK] Found '{SUB_VARIANT_COL}' column in master - sub-variant filtering ENABLED.")
    else:
        print(f"[WARN] '{SUB_VARIANT_COL}' column NOT found in master - sub-variant filtering DISABLED.")
    if PRODUCT_CODE_COL in df.columns:
        cols.append(PRODUCT_CODE_COL)
        print(f"[OK] Found '{PRODUCT_CODE_COL}' column in master - product codes will be returned.")
    else:
        print(f"[WARN] '{PRODUCT_CODE_COL}' column NOT found in master - product codes will not be available.")
    df = df[cols].copy()
    df[PRODUCT_COL] = df[PRODUCT_COL].map(lambda value: '' if pd.isna(value) else str(value))
    df[BRAND_COL] = df[BRAND_COL].map(normalize_master_text)
    if PACK_SIZE_COL in df.columns:
        df[PACK_SIZE_COL] = df[PACK_SIZE_COL].map(normalize_master_text)
    else:
        df[PACK_SIZE_COL] = ''
    if VARIANT_COL in df.columns:
        df[VARIANT_COL] = df[VARIANT_COL].map(normalize_master_attr)
    else:
        df[VARIANT_COL] = 'PLAIN'
    if SUB_VARIANT_COL in df.columns:
        df[SUB_VARIANT_COL] = df[SUB_VARIANT_COL].map(normalize_master_attr)
    else:
        df[SUB_VARIANT_COL] = 'PLAIN'
    if PRODUCT_CODE_COL in df.columns:
        df[PRODUCT_CODE_COL] = df[PRODUCT_CODE_COL].map(normalize_master_code)
    else:
        df[PRODUCT_CODE_COL] = ''
    df = df[(df[PRODUCT_COL] != '') & (df[BRAND_COL] != '')]
    df['product_norm'] = df[PRODUCT_COL].map(norm)
    df['brand_norm'] = df[BRAND_COL].map(norm)
    return df

def build_brand_product_map(df: pd.DataFrame) -> dict:
    brand_map = {}
    for brand_norm, group in df.groupby('brand_norm'):
        original_brand = group.iloc[0][BRAND_COL]
        seen = set()
        items = []
        for _, row in group.iterrows():
            prod = str(row[PRODUCT_COL])
            pack = str(row.get(PACK_SIZE_COL, '')).strip()
            pack = '' if pack in ('nan', 'NaN', 'None') else pack
            variant = str(row.get(VARIANT_COL, 'PLAIN')).strip().upper()
            variant = 'PLAIN' if variant in ('', 'NAN', 'NONE') else variant
            sub_variant = str(row.get(SUB_VARIANT_COL, 'PLAIN')).strip().upper()
            sub_variant = 'PLAIN' if sub_variant in ('', 'NAN', 'NONE') else sub_variant
            variant, sub_variant = repair_master_variant_fields(prod, variant, sub_variant)
            product_code = str(row.get(PRODUCT_CODE_COL, '')).strip()
            product_code = '' if product_code in ('nan', 'NaN', 'None') else product_code
            key = (prod, pack, product_code, variant, sub_variant)
            if prod and key not in seen:
                seen.add(key)
                items.append({'product': prod, 'pack_size': pack, 'variant': variant, 'sub_variant': sub_variant, 'product_code': product_code, 'product_norm': norm(prod)})
        brand_map[original_brand] = items
    return brand_map

def extract_all_variants_from_data(brand_map: dict) -> set:
    all_variants = set()
    for brand, items in brand_map.items():
        for item in items:
            variant = item.get('variant', 'PLAIN')
            if variant != 'PLAIN':
                all_variants.add(variant.upper())
    all_variants.add('PLAIN')
    return all_variants

def extract_all_sub_variants_from_data(brand_map: dict) -> set:
    all_sub_variants = set()
    for brand, items in brand_map.items():
        for item in items:
            sub_variant = item.get('sub_variant', 'PLAIN')
            if sub_variant != 'PLAIN':
                all_sub_variants.add(sub_variant.upper())
    all_sub_variants.add('PLAIN')
    return all_sub_variants

def find_potential_brands(input_name: str, brand_map: dict) -> list:
    input_upper = input_name.upper()
    input_words = _normalize_strength_words(re.findall('\\b[A-Z0-9]+\\b', input_upper))
    input_compact = _expand_strength_in_compact(compact_brand_token(input_upper))
    potential = []
    for brand, items in brand_map.items():
        brand_upper = brand.upper().replace('-', ' ').strip()
        brand_words = _normalize_strength_words(brand_upper.split())
        brand_compact = _expand_strength_in_compact(compact_brand_token(brand))
        brand_pattern = '\\b' + re.escape(brand_upper) + '\\b'
        if re.search(brand_pattern, input_upper):
            extra_words = brand_words - input_words
            if not extra_words:
                potential.append((brand, items, 3))
            else:
                potential.append((brand, items, 1))
            continue
        flexible_brand_pattern = build_flexible_brand_pattern(brand)
        if flexible_brand_pattern and re.search(flexible_brand_pattern, input_upper):
            potential.append((brand, items, 4))
            continue
        if brand_compact and input_compact.startswith(brand_compact):
            remainder = input_compact[len(brand_compact):]
            if not remainder or remainder[0].isdigit():
                potential.append((brand, items, 3))
                continue
        if brand_words.issubset(input_words):
            potential.append((brand, items, 2))
            continue
        brand_first_word = brand_upper.split()[0] if brand_upper.split() else ''
        if brand_first_word and brand_first_word in input_words:
            potential.append((brand, items, 0))
    potential.sort(key=lambda x: (-x[2], len(x[0])))
    return [(brand, items) for brand, items, _ in potential]

def group_related_brands(primary_brand: str, brand_map: dict, input_name: str='', verbose=False) -> list:
    primary_upper = primary_brand.upper().replace('-', ' ').strip()
    input_words = _normalize_strength_words(re.findall('\\b[A-Z0-9]+\\b', input_name.upper())) if input_name else set()
    grouped_items = []
    related_brands = []
    primary_parts = primary_upper.split()
    primary_root = primary_parts[0] if primary_parts else ''
    for brand, items in brand_map.items():
        brand_upper = brand.upper().replace('-', ' ').strip()
        if brand_upper == primary_upper:
            grouped_items.extend(items)
            related_brands.append(brand)
        elif brand_upper.startswith(primary_upper + ' '):
            primary_words = _normalize_strength_words(primary_upper.split())
            brand_words = _normalize_strength_words(brand_upper.split())
            extra_words = {word for word in brand_words - primary_words if word not in _pack_form_skip_tokens()}
            if not extra_words or (input_words and extra_words.issubset(input_words)):
                grouped_items.extend(items)
                related_brands.append(brand)
            elif verbose:
                print(f"    [WARN] Skipping sub-brand '{brand}' -- extra words {extra_words} not in input")
        elif primary_root:
            brand_parts = brand_upper.split()
            brand_root = brand_parts[0] if brand_parts else ''
            if brand_root == primary_root:
                extra_words = {word for word in _normalize_strength_words(brand_parts[1:]) if word not in _pack_form_skip_tokens()}
                if extra_words and input_words and extra_words.issubset(input_words):
                    grouped_items.extend(items)
                    related_brands.append(brand)
    if verbose and len(related_brands) > 1:
        print(f'\n  Found {len(related_brands)} related brands:')
        for rb in sorted(related_brands):
            print(f"    - '{rb}' ({len(brand_map[rb])} products)")
    return grouped_items

def _pack_form_skip_tokens():
    tokens = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES)
    tokens |= {'PCS', 'PC', 'NOS', 'NO', 'VIAL', 'AMP', 'AMPS', 'PFS', 'SYR', 'POWDE', 'POWDER', 'GRANUL', 'GRANULE', 'GRANULES', 'GRANULS'}
    return tokens

def _brand_letters(s: str) -> str:
    return re.sub('[^A-Z]', '', str(s or '').upper())

def _is_strong_brand_token(token: str, brand_compacts: set) -> bool:
    c = compact_brand_token(token)
    if not c:
        return False
    if c in brand_compacts:
        return True
    return any((c.startswith(bc) and len(c) > len(bc) and c[len(bc)].isdigit() for bc in brand_compacts))
_DESEG_UNITS = 'MG|MCG|ML|GM|IU'
_DESEG_FORMS = 'TABLETS|TABLET|TABS|TAB|CAPSULES|CAPSULE|CAPS|CAP|INJECTION|INJ|SYRUP|SYP|DROPS|GRANULES|SUSPENSION'

def _split_glued_brand_token(tok: str, brand_letters: set) -> str:
    m = re.match('^[A-Za-z]+', tok)
    if not m:
        return tok
    lead = m.group(0).upper()
    best = ''
    for b in brand_letters:
        if len(best) < len(b) <= len(lead) and lead.startswith(b):
            best = b
    if best and best != lead:
        return tok[:len(best)] + ' ' + tok[len(best):].lstrip('-')
    if best == lead and len(tok) > len(lead):
        return lead + ' ' + tok[len(lead):].lstrip('-')
    return tok
_VST_CACHE_MAP = None
_VST_CACHE_VAL = None

def _variant_subvariant_tokens(brand_map: dict) -> tuple:
    global _VST_CACHE_MAP, _VST_CACHE_VAL
    if _VST_CACHE_MAP is brand_map:
        return _VST_CACHE_VAL
    var, subvar = (set(), set())
    for items in brand_map.values():
        for it in items:
            for t in re.findall('[A-Z]{2,}', str(it.get('variant', '')).upper()):
                var.add(t)
            for t in re.findall('[A-Z]{2,}', str(it.get('sub_variant', '')).upper()):
                subvar.add(t)
    _VST_CACHE_MAP, _VST_CACHE_VAL = (brand_map, (var, subvar))
    return (var, subvar)

def _split_glued_variant(tok: str, var_tokens: set, subvar_tokens: set) -> str:
    u = tok.upper()
    if not tok.isalpha() or len(u) < 4 or u in var_tokens or (u in subvar_tokens):
        return tok
    for i in range(2, len(u) - 1):
        if u[:i] in var_tokens and u[i:] in subvar_tokens:
            return tok[:i] + ' ' + tok[i:]
    return tok

def desegment_ocr_input(input_name: str, brand_map: dict) -> str:
    s = str(input_name or '')
    s = re.sub('[Xx]{2,}(?=\\s*-{2,})', '', s)
    s = re.sub('(?<!\\S)[Xx]{2,}(?!\\S)', ' ', s)
    s = re.sub('\\s*-{1,}\\s*$', '', s)
    s = re.sub('(\\d)\\.\\s+(\\d)', '\\1.\\2', s)
    s = re.sub('\\s+', ' ', s).strip()
    tokens = s.split()
    if not tokens:
        return s
    brand_compacts = {compact_brand_token(b) for b in brand_map}
    brand_letters = {_brand_letters(b) for b in brand_map if len(_brand_letters(b)) >= 4}
    if len(tokens) >= 2 and len(tokens[0]) == 1 and tokens[0].isalpha() and _is_strong_brand_token(tokens[0] + tokens[1], brand_compacts):
        tokens = [tokens[0] + tokens[1]] + tokens[2:]
    tokens[0] = _split_glued_brand_token(tokens[0], brand_letters)
    s = ' '.join(tokens)
    var_tokens, subvar_tokens = _variant_subvariant_tokens(brand_map)
    known_subbrands = var_tokens | subvar_tokens
    parts = s.split()
    new_parts = []
    for p in parts:
        split_done = False
        for b in brand_compacts:
            if p.startswith(b) and len(p) > len(b):
                suffix = p[len(b):]
                if suffix in known_subbrands:
                    new_parts.append(b)
                    new_parts.append(suffix)
                    split_done = True
                    break
        if not split_done:
            new_parts.append(p)
    s = ' '.join(new_parts)
    s = re.sub('\\bIN[- ](\\d+)', 'INJ \\1', s)
    s = re.sub('\\bSY[- ](\\d+)', 'SYP \\1', s)
    s = re.sub('\\bSUSP[- ](\\d+)', 'SUSPENSION \\1', s)
    var_tokens, subvar_tokens = _variant_subvariant_tokens(brand_map)
    parts = s.split()
    s = ' '.join(parts[:1] + [_split_glued_variant(p, var_tokens, subvar_tokens) for p in parts[1:]])
    s = re.sub('\\b(AMPOULEP)(\\d)', '\\1 \\2', s)
    s = re.sub('\\b(AMPOULE)(\\d)', '\\1 \\2', s)
    s = re.sub('\\b(AMP)(\\d)', '\\1 \\2', s)
    s = re.sub('\\b(AMPS)(\\d)', '\\1 \\2', s)
    s = re.sub('(?i)(\\d)\\s*(' + _DESEG_UNITS + ')(' + _DESEG_FORMS + ')', '\\1 \\2 \\3', s)
    s = re.sub('(?i)(?<=[A-WYZa-wyz])(?=\\d)', ' ', s)
    s = re.sub('(?i)(\\d)(' + _DESEG_FORMS + ')\\b', '\\1 \\2', s)
    return re.sub('\\s+', ' ', s).strip()

def is_substring_brand_reject(query_brand: str, candidate_brand: str) -> bool:
    q = _brand_letters(query_brand)
    c = _brand_letters(candidate_brand)
    if not q or not c:
        return False
    return c != q and c in q

def find_unmatched_qualifier_tokens(input_name: str, brand: str, available_variants, available_sub_variants, brand_items=None) -> list:
    text = str(input_name or '').upper().replace('-', ' ')
    brand_up = str(brand or '').upper().replace('-', ' ').strip()

    def _find_boundary(hay, needle):
        if not needle:
            return -1
        start = 0
        while True:
            i = hay.find(needle, start)
            if i < 0:
                return -1
            end = i + len(needle)
            before_ok = i == 0 or not hay[i - 1].isalnum()
            after_ok = end >= len(hay) or not hay[end].isalnum()
            if before_ok and after_ok:
                return i
            start = i + 1
    idx = _find_boundary(text, brand_up)
    if idx >= 0:
        tail = text[idx + len(brand_up):]
    else:
        first = brand_up.split()[0] if brand_up.split() else ''
        pos = _find_boundary(text, first)
        if pos >= 0:
            tail = text[pos + len(first):]
        else:
            bq = extract_brand_like_query(input_name).upper()
            bq_first = bq.split()[0] if bq.split() else bq
            bpos = _find_boundary(text, bq_first)
            tail = text[bpos + len(bq_first):] if bpos >= 0 else ''
    known = set(FORM_HINT_TOKENS) | set(DOSAGE_UNIT_SUFFIXES) | _pack_form_skip_tokens()
    known |= {'PLAIN'}
    known |= {str(v).upper() for v in available_variants or set()}
    known |= {str(v).upper() for v in available_sub_variants or set()}
    brand_vocab = set()
    for it in brand_items or []:
        for tok in re.findall('[A-Z]+', str(it.get('product', '')).upper()):
            brand_vocab.add(tok)
    prefix_targets = (brand_vocab | known) - ({'PLAIN'} | set(FORM_HINT_TOKENS))
    brand_compact = _brand_letters(brand_up)
    orphans = []
    tail_tokens = re.findall('[A-Z0-9]+', tail)
    for ti, tok in enumerate(tail_tokens):
        if any((ch.isdigit() for ch in tok)):
            break
        if tok in FORM_HINT_TOKENS:
            break
        canon = canonical_variant_match_token(tok)
        if tok in known or canon in known or tok in brand_vocab:
            continue
        if brand_compact + tok in brand_vocab:
            continue
        prev_join = tail_tokens[ti - 1] + tok if ti > 0 else ''
        next_join = tok + tail_tokens[ti + 1] if ti + 1 < len(tail_tokens) else ''
        if prev_join and (prev_join in known or prev_join in brand_vocab) or (next_join and (next_join in known or next_join in brand_vocab)):
            continue
        if len(tok) >= 3 and any((t != tok and t.startswith(tok) for t in prefix_targets)):
            continue
        orphans.append(tok)
    if idx > 0:
        for tok in re.findall('[A-Z]+', text[:idx]):
            if tok in {'SEMI', 'HALF'} and tok not in brand_vocab and (tok not in orphans):
                orphans.append(tok)
    return orphans

def find_abbrev_promoted_brand(input_name: str, matched_brand: str, brand_map: dict) -> str:
    if not matched_brand or not brand_map:
        return None
    input_upper = str(input_name or '').upper()
    matched_upper = str(matched_brand).upper().replace('-', ' ').strip()
    if not matched_upper:
        return None
    bm = re.search('\\b' + re.escape(matched_upper) + '\\b', input_upper)
    if not bm:
        return None
    tail_match = re.match('[\\s\\-]*([A-Z]{2,})\\b', input_upper[bm.end():])
    if not tail_match:
        return None
    abbrev_token = tail_match.group(1)
    if abbrev_token in _pack_form_skip_tokens():
        return None
    prefix = matched_upper + ' '
    candidates = []
    for brand in brand_map.keys():
        b_upper = str(brand).upper().replace('-', ' ').strip()
        if b_upper == matched_upper or not b_upper.startswith(prefix):
            continue
        next_word = b_upper[len(prefix):].split()[0] if b_upper[len(prefix):] else ''
        if len(next_word) < 2:
            continue
        if next_word.startswith(abbrev_token) or abbrev_token.startswith(next_word):
            candidates.append(brand)
    return candidates[0] if len(candidates) == 1 else None

def filter_items_by_product_name_abbrev(input_name: str, brand_name: str, items: list, verbose=False) -> list:
    if not brand_name or not items:
        return items
    input_upper = str(input_name or '').upper()
    matched_upper = str(brand_name).upper().replace('-', ' ').strip()
    if not matched_upper:
        return items
    bm = re.search('\\b' + re.escape(matched_upper) + '\\b', input_upper)
    if not bm:
        return items
    tail_tokens = re.findall('[A-Z]{2,}', input_upper[bm.end():])
    if not tail_tokens:
        return items
    brand_tokens = set(matched_upper.split())
    filtered_items = list(items)
    for abbrev in tail_tokens:
        if abbrev in _pack_form_skip_tokens():
            continue
        expansions = set()
        abbrev_as_exact_match = False
        for item in filtered_items:
            prod = str(item['product']).upper().replace('-', ' ')
            for tok in re.findall('\\b[A-Z]{2,}\\b', prod):
                if tok in brand_tokens or tok in _pack_form_skip_tokens():
                    continue
                if tok == abbrev:
                    abbrev_as_exact_match = True
                elif len(tok) > len(abbrev) and tok.startswith(abbrev):
                    expansions.add(tok)
        if abbrev_as_exact_match:
            continue
        if len(expansions) != 1:
            continue
        expansion = next(iter(expansions))
        narrowed = [i for i in filtered_items if re.search('\\b' + re.escape(expansion) + '\\b', str(i['product']).upper().replace('-', ' '))]
        if narrowed and len(narrowed) < len(filtered_items):
            if verbose:
                print(f"  [PRODUCT-NAME ABBREV] '{abbrev}' -> '{expansion}': filtered {len(filtered_items)} -> {len(narrowed)} items")
            filtered_items = narrowed
    return filtered_items

def find_best_brand_for_input(input_name: str, brand_map: dict, verbose=False) -> tuple:
    if verbose:
        print('\n' + '=' * 80)
        print('STEP 1: BRAND DETECTION')
        print('=' * 80)
        print(f"Input name: '{input_name}'")
    potential = find_potential_brands(input_name, brand_map)
    if verbose:
        print(f'\nPattern matching found {len(potential)} potential brand(s):')
        for i, (brand, items) in enumerate(potential[:5], 1):
            print(f"  {i}. '{brand}' ({len(items)} products)")
    input_upper = input_name.upper().replace('-', ' ')
    input_words = _normalize_strength_words(re.findall('\\b[A-Z0-9]+\\b', input_upper))
    input_compact = _expand_strength_in_compact(compact_brand_token(input_name))

    def _brand_position(brand_name: str) -> int:
        brand_upper = str(brand_name or '').upper().replace('-', ' ').strip()
        if not brand_upper:
            return 10 ** 6
        m = re.search('\\b' + re.escape(brand_upper) + '\\b', input_upper)
        if m:
            return m.start()
        brand_compact = _expand_strength_in_compact(compact_brand_token(brand_name))
        if brand_compact:
            compact_pos = input_compact.find(brand_compact)
            if compact_pos >= 0:
                return compact_pos
        return 10 ** 6
    input_joins = _input_compact_joins(input_name)
    filtered_potential = []
    for brand, items in potential:
        brand_upper = brand.upper().replace('-', ' ')
        brand_words = _normalize_strength_words(brand_upper.split())
        brand_compact = _expand_strength_in_compact(compact_brand_token(brand))
        extra_words = {w for w in brand_words - input_words if compact_brand_token(w) not in input_joins}
        compact_match = False
        if brand_compact and input_compact.startswith(brand_compact):
            remainder = input_compact[len(brand_compact):]
            if not remainder or remainder[0].isdigit():
                compact_match = True
        if extra_words and (not compact_match):
            if verbose:
                print(f"  [FAIL] Rejecting '{brand}' - extra words: {extra_words}")
            continue
        filtered_potential.append((brand, items))
    if filtered_potential:
        filtered_potential.sort(key=lambda x: (_brand_position(x[0]), -len(x[0])))
        selected_brand, _ = filtered_potential[0]
        promoted = find_abbrev_promoted_brand(input_name, selected_brand, brand_map)
        if promoted and promoted != selected_brand:
            if verbose:
                print(f"\n[OK] Promoted '{selected_brand}' -> '{promoted}' (input contains abbreviation of sub-brand)")
            selected_brand = promoted
        if verbose:
            print(f"\n[OK] Selected brand: '{selected_brand}'")
        grouped_items = group_related_brands(selected_brand, brand_map, input_name=input_name, verbose=verbose)
        return (selected_brand, grouped_items)
    if not potential:
        input_norm = norm(input_name)
        brand_query_norm = norm(extract_brand_like_query(input_name))
        if verbose:
            print(f'\nFuzzy matching against {len(brand_map)} brands...')
        all_brands = list(brand_map.keys())
        candidate_brands = []
        for brand in all_brands:
            brand_upper = brand.upper().replace('-', ' ')
            brand_first_word = brand_upper.split()[0] if brand_upper.split() else ''
            if not brand_first_word or brand_first_word in input_upper or len(brand.split()) == 1:
                candidate_brands.append(brand)
        if not candidate_brands:
            candidate_brands = all_brands
        best = None
        best_source = 'fuzzy'
        brand_root_norm = brand_query_norm.split()[0] if brand_query_norm else ''
        brand_tokens_norm = [t for t in brand_query_norm.split() if len(t) >= 4]
        if brand_query_norm:
            best = process.extractOne(brand_query_norm, candidate_brands, scorer=fuzz.ratio, score_cutoff=BRAND_NAME_NEXTSTEP_SCORE)
            if best:
                best_source = 'brand-only'
        for fuzzy_query in [brand_query_norm, brand_root_norm, *brand_tokens_norm, input_norm]:
            if best:
                break
            if not fuzzy_query:
                continue
            best = process.extractOne(fuzzy_query, candidate_brands, scorer=fuzz.token_sort_ratio, score_cutoff=FUZZY_MIN_SCORE)
            if best:
                break
        if not best:
            for fuzzy_query in [brand_query_norm, brand_root_norm, *brand_tokens_norm]:
                if not fuzzy_query:
                    continue
                best = process.extractOne(fuzzy_query, candidate_brands, scorer=JaroWinkler.normalized_similarity, score_cutoff=JW_MIN_SCORE)
                if best:
                    best_source = 'jaro-winkler'
                    break
        if best:
            best_brand, score, _ = best
            if is_substring_brand_reject(extract_brand_like_query(input_name), best_brand):
                if verbose:
                    print(f"\n[FAIL] Rejecting fuzzy brand '{best_brand}' -- it is a sub-token of the input brand (different brand, not a typo)")
                return (None, None)
            cand_compact = compact_brand_token(best_brand)
            edit_dists = [Levenshtein.distance(compact_brand_token(q), cand_compact) for q in (extract_brand_like_query(input_name), brand_root_norm, *brand_tokens_norm) if compact_brand_token(q)]
            if edit_dists and min(edit_dists) > BRAND_FUZZY_MAX_EDITS:
                if verbose:
                    print(f"\n[FAIL] Rejecting fuzzy brand '{best_brand}' -- {min(edit_dists)} edits from input brand (> {BRAND_FUZZY_MAX_EDITS}); different brand, not a typo")
                return (None, None)
            promoted = find_abbrev_promoted_brand(input_name, best_brand, brand_map)
            if promoted and promoted != best_brand:
                if verbose:
                    print(f"\n[OK] Promoted fuzzy match '{best_brand}' -> '{promoted}'")
                best_brand = promoted
            if verbose:
                if best_source == 'brand-only':
                    print(f"\n[OK] Brand-only high-confidence match: '{best_brand}' (score: {score})")
                else:
                    print(f"\n[OK] Fuzzy match: '{best_brand}' (score: {score})")
            grouped_items = group_related_brands(best_brand, brand_map, input_name=input_name, verbose=verbose)
            return (best_brand, grouped_items)
    return (None, None)

def detect_variants_in_input(input_name: str, brand_name: str, available_variants: set, compound_token_letters: set=None, verbose=False, brand_items: list=None) -> set:
    input_upper = normalize_ocr_numeric_noise(input_name)
    pack_positions = get_all_pack_positions(input_upper)

    def _single_letter_variant_only_in_pack(token: str) -> bool:
        token = str(token or '').strip().upper()
        if token not in {'S', 'T', 'C', 'P'}:
            return False
        token_matches = list(re.finditer('\\b' + re.escape(token) + '\\b', input_upper))
        if not token_matches:
            return False
        return all((is_position_in_pack_pattern(m.start(), m.end(), pack_positions) for m in token_matches))
    if brand_name:
        input_without_brand = strip_brand_from_input(input_upper, brand_name, available_variants=available_variants)
    else:
        input_without_brand = input_upper
    input_without_brand = re.sub('\\s+', ' ', input_without_brand).strip()
    normalized_input = normalize_variant_token(input_without_brand)
    compact_input_token_list = [compact_variant_token(token) for token in re.findall('[A-Z0-9.\\-]+', input_without_brand) if compact_variant_token(token)]
    compact_input_tokens = set(compact_input_token_list)
    for i in range(len(compact_input_token_list) - 1):
        compact_input_tokens.add(compact_input_token_list[i] + compact_input_token_list[i + 1])
    for i in range(len(compact_input_token_list) - 2):
        compact_input_tokens.add(compact_input_token_list[i] + compact_input_token_list[i + 1] + compact_input_token_list[i + 2])
    detected_variants = set()
    split_variants, _ = split_strength_and_variant(input_without_brand, available_variants, verbose=verbose)
    if compound_token_letters:
        split_variants -= compound_token_letters
    detected_variants.update(split_variants)
    combo_variants, combo_sub_variants = detect_compact_variant_strength_combo(input_without_brand, available_variants)
    if compound_token_letters:
        combo_variants -= compound_token_letters
    detected_variants.update(combo_variants)
    pair_variants, _ = detect_variant_numeric_pairs(input_without_brand, available_variants)
    detected_variants.update(pair_variants)
    words = re.findall('[A-Z0-9]+', input_without_brand)
    input_match_tokens = extract_variant_match_tokens(input_without_brand)
    input_descriptor_tokens = {token for token in input_match_tokens if token not in GENERIC_VARIANT_MATCH_TOKENS}
    input_forms_detected = detect_dosage_form_in_input(input_name, verbose=False)
    for variant in available_variants:
        if variant == 'PLAIN':
            continue
        variant_upper = variant.upper()
        if len(variant_upper) == 1 and _single_letter_variant_only_in_pack(variant_upper):
            continue
        variant_compact = compact_variant_token(variant_upper)
        if compound_token_letters and variant_upper in compound_token_letters:
            continue
        if variant_compact and variant_compact in compact_input_tokens:
            detected_variants.add(variant)
            continue
        if variant_upper in words:
            detected_variants.add(variant)
            continue
        patterns = ['[\\-\\s]' + re.escape(variant_upper) + '[\\-\\s]', '[\\-\\s]' + re.escape(variant_upper) + '\\b', '\\b' + re.escape(variant_upper) + '[\\-\\s]']
        for pattern in patterns:
            if re.search(pattern, input_without_brand):
                detected_variants.add(variant)
                break
    for variant in available_variants:
        if variant == 'PLAIN' or len(variant) <= 1:
            continue
        if compound_token_letters and variant.upper() in compound_token_letters:
            continue
        variant_upper = variant.upper()
        variant_normalized = normalize_variant_token(variant_upper)
        if variant_normalized and re.search('\\b' + re.escape(variant_normalized).replace('\\ ', '\\s+') + '\\b', normalized_input):
            detected_variants.add(variant)
            continue
        if re.search('\\b' + re.escape(variant_upper) + '\\b', input_without_brand):
            detected_variants.add(variant)
    for variant in available_variants:
        if variant == 'PLAIN':
            continue
        if len(str(variant).strip()) == 1 and _single_letter_variant_only_in_pack(variant):
            continue
        variant_match_tokens = extract_variant_match_tokens(variant)
        if not variant_match_tokens:
            continue
        if variant_match_tokens.issubset(input_match_tokens):
            if variant not in detected_variants:
                detected_variants.add(variant)
            continue
        if len(variant_match_tokens) >= 2 and input_descriptor_tokens and (len(input_descriptor_tokens) >= 2) and input_descriptor_tokens.issubset(variant_match_tokens):
            detected_variants.add(variant)
        variant_form_tokens = variant_match_tokens & GENERIC_VARIANT_MATCH_TOKENS
        variant_non_form_tokens = variant_match_tokens - GENERIC_VARIANT_MATCH_TOKENS
        if variant not in detected_variants and variant_non_form_tokens and variant_non_form_tokens.issubset(input_match_tokens) and input_forms_detected and variant_form_tokens and forms_share_group(variant_form_tokens, input_forms_detected):
            detected_variants.add(variant)
    input_form_tokens = input_match_tokens & GENERIC_VARIANT_MATCH_TOKENS
    for form_token in sorted(input_form_tokens):
        form_variants = []
        for variant in available_variants:
            if variant == 'PLAIN':
                continue
            variant_match_tokens = extract_variant_match_tokens(variant)
            if form_token in variant_match_tokens:
                form_variants.append(variant)
        if len(form_variants) == 1:
            variant = form_variants[0]
            non_form = extract_variant_match_tokens(variant) - GENERIC_VARIANT_MATCH_TOKENS
            if non_form and (not non_form.issubset(input_match_tokens)):
                continue
            if variant not in detected_variants:
                detected_variants.add(variant)
    input_collapsed = collapse_spaced_thousands(input_without_brand)
    input_strength_numbers = set(re.findall('\\d+(?:\\.\\d+)?', input_collapsed))
    if input_strength_numbers:
        for variant in available_variants:
            if variant in detected_variants:
                continue
            vnum = numeric_strength_variant_value(variant)
            if vnum and vnum in input_strength_numbers:
                detected_variants.add(variant)
    _INJ_FORM_VAR = {'IV', 'INJ', 'INJECTION', 'VIAL', 'PFS', 'AMP', 'AMPS', 'SYR', 'SYRINGE'}
    if input_forms_detected and forms_share_group({'INJECTION'}, input_forms_detected):
        for variant in available_variants:
            if variant == 'PLAIN' or variant in detected_variants:
                continue
            vt = {re.sub('[^A-Z]', '', t) for t in str(variant).upper().split()}
            vt = {t for t in vt if t}
            if vt and vt <= _INJ_FORM_VAR:
                if input_strength_numbers and brand_items is not None:
                    has_dose = any((str(it.get('sub_variant', '')).strip() in input_strength_numbers for it in brand_items if str(it.get('variant', '')).upper() == str(variant).upper()))
                    if not has_dose:
                        continue
                detected_variants.add(variant)
    if brand_items:
        input_combos = set(re.findall('(?<![\\d./])\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)+(?![\\d./])', input_without_brand))
        if input_combos:
            for it in brand_items:
                ivar = str(it.get('variant', 'PLAIN')).strip().upper()
                isub = str(it.get('sub_variant', '')).strip().upper()
                if ivar and ivar != 'PLAIN' and ('/' in isub) and (isub in input_combos):
                    detected_variants.add(ivar)
    variant_suffix_patterns = {v: '\\b' + re.escape(v) + '\\b' for v in available_variants if v != 'PLAIN'}
    input_upper_for_var = input_without_brand.upper()
    pack_positions = get_all_pack_positions(input_upper_for_var)

    def is_in_pack_pattern(start, end):
        for ps, pe in pack_positions:
            if ps <= start < pe or ps < end <= pe:
                return True
        return False
    for var_name, var_pattern in variant_suffix_patterns.items():
        if var_name not in available_variants:
            continue
        if len(var_name) == 1:
            for m in re.finditer(var_pattern, input_upper_for_var):
                if not is_in_pack_pattern(m.start(), m.end()):
                    detected_variants.add(var_name)
                    break
        elif re.search(var_pattern, input_upper_for_var):
            detected_variants.add(var_name)
    single_letter_variants = {v for v in available_variants if len(v) == 1 and v.isalpha()}
    for var_name in single_letter_variants:
        if var_name in available_variants:
            pattern_glued = '\\d(?:\\.\\d+)?' + re.escape(var_name) + '(?![A-Z])'
            for m in re.finditer(pattern_glued, input_upper_for_var):
                if not is_in_pack_pattern(m.start(), m.end()):
                    detected_variants.add(var_name)
                    break
            pattern_token_start = '(?<![A-Z0-9])' + re.escape(var_name) + '(?=\\d)'
            for m in re.finditer(pattern_token_start, input_upper_for_var):
                if not is_in_pack_pattern(m.start(), m.end()):
                    detected_variants.add(var_name)
                    break
            pattern_spaced = '\\d(?:\\.\\d+)?\\s+' + re.escape(var_name) + '(?![A-Z])'
            for m in re.finditer(pattern_spaced, input_upper_for_var):
                if not is_in_pack_pattern(m.start(), m.end()):
                    detected_variants.add(var_name)
                    break
    if not any((variant != 'PLAIN' for variant in detected_variants)):
        detected_variants.add('PLAIN')
    if verbose:
        print(f'  Final detected variants: {sorted(detected_variants)}')
    return detected_variants

def detect_sub_variants_in_input(input_name: str, brand_name: str, available_sub_variants: set, available_variants: set, compound_token_digits: set=None, verbose=False) -> set:
    input_upper = normalize_ocr_numeric_noise(input_name)
    detected_sub_variants = set()
    if brand_name:
        input_without_brand = strip_brand_from_input(input_upper, brand_name, available_variants=available_variants)
    else:
        input_without_brand = input_upper
    combo_dosage = extract_combination_dosage(input_upper)
    combo_components: set = set()
    if combo_dosage:
        detected_sub_variants.add(combo_dosage)
        combo_components = set(re.findall('\\d+(?:\\.\\d+)?', combo_dosage))
    _, split_sub_variants = split_strength_and_variant(input_without_brand, available_variants, verbose=verbose)
    if compound_token_digits:
        split_sub_variants -= compound_token_digits
    if combo_components:
        split_sub_variants -= combo_components
    detected_sub_variants.update(split_sub_variants)
    _, combo_sub_variants = detect_compact_variant_strength_combo(input_without_brand, available_variants)
    if combo_components:
        combo_sub_variants -= combo_components
    detected_sub_variants.update(combo_sub_variants)
    _, pair_sub_variants = detect_variant_numeric_pairs(input_without_brand, available_variants)
    if combo_components:
        pair_sub_variants -= combo_components
    detected_sub_variants.update(pair_sub_variants)
    pack_positions = get_all_pack_positions(input_upper)
    sorted_sub_variants = sorted([sv for sv in available_sub_variants if sv != 'PLAIN'], key=lambda x: (len(str(x).replace('.', '')), len(str(x))), reverse=True)
    matched_positions = []

    def is_position_matched(start, end):
        for ms, me in matched_positions:
            if ms <= start < me or ms < end <= me:
                return True
        return False
    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = extract_numeric_strength(sub_var)
        if not sub_str:
            continue
        unit_pattern = '(?<![\\d.])' + re.escape(sub_str) + '\\s*(?:MG|GM|MCG|G|M|IU|UNIT)(?:\\b|/)'
        for m in re.finditer(unit_pattern, input_upper):
            if not is_position_in_pack_pattern(m.start(), m.end(), pack_positions) and (not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break
    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = str(sub_var)
        pattern = '(?<![\\d.])' + re.escape(sub_str) + '(?![\\d.])'
        for m in re.finditer(pattern, input_upper):
            context_start = max(0, m.start() - 5)
            context_end = min(len(input_upper), m.end() + 5)
            context = input_upper[context_start:context_end]
            if '/' in context:
                continue
            if re.match('\\s*M?L\\b', input_upper[m.end():]):
                continue
            if m.start() > 0 and input_upper[m.start() - 1] == '.' or re.match('\\.\\d', input_upper[m.end():]):
                continue
            if is_position_in_pack_pattern(m.start(), m.end(), pack_positions):
                continue
            if not is_position_in_pack_pattern(m.start(), m.end(), pack_positions) and (not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break
    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants:
            continue
        if compound_token_digits and str(sub_var) in compound_token_digits:
            continue
        if combo_components and str(sub_var) in combo_components:
            continue
        sub_str = str(sub_var)
        pattern = '[-\\s/]' + re.escape(sub_str) + '(?![\\d.])'
        for m in re.finditer(pattern, input_upper):
            if re.match('\\s*M?L\\b', input_upper[m.end():]):
                continue
            if re.match('\\.\\d', input_upper[m.end():]):
                continue
            if not is_position_in_pack_pattern(m.start(), m.end(), pack_positions) and (not is_position_matched(m.start(), m.end())):
                detected_sub_variants.add(sub_var)
                matched_positions.append((m.start(), m.end()))
                break
    input_alpha_tokens = {t for t in re.findall('[A-Z]+', input_upper) if len(t) >= 3}
    input_alpha_signatures = {alpha_signature(t) for t in input_alpha_tokens}
    input_tokens_set_upper = set(re.findall('[A-Z]+', input_upper))
    for sub_var in sorted_sub_variants:
        if sub_var in detected_sub_variants or is_numeric_subvariant_like(sub_var):
            continue
        sub_var_upper = str(sub_var).strip().upper()
        if sub_var_upper not in input_tokens_set_upper:
            continue
        sub_signature = alpha_signature(sub_var)
        if len(sub_signature) >= 3 and sub_signature in input_alpha_signatures:
            detected_sub_variants.add(sub_var)
    _SV_ABBREV = {'FORT': 'FORTE', 'FORTS': 'FORTE'}
    for abbrev, expanded in _SV_ABBREV.items():
        if abbrev in input_tokens_set_upper and expanded in available_sub_variants:
            detected_sub_variants.add(expanded)
    if not any((is_numeric_subvariant_like(s) for s in detected_sub_variants if s != 'PLAIN')):
        rescue_text = normalize_ocr_numeric_noise(input_without_brand)
        input_numbers = set()
        for m in re.finditer('(\\d+(?:\\.\\d+)?)\\s*([A-Z]*)', rescue_text):
            unit = m.group(2)
            if unit == 'L' or unit.startswith('ML'):
                continue
            input_numbers.add(m.group(1))
        strength_subs = {s for s in available_sub_variants if is_numeric_subvariant_like(s) and '/' not in str(s)}
        hits = {s for s in strength_subs if str(s) in input_numbers}
        if len(hits) == 1:
            detected_sub_variants.add(next(iter(hits)))
    detected_sub_variants.add('PLAIN')
    if verbose:
        print(f'  Final detected sub-variants: {sorted(detected_sub_variants)}')
    return detected_sub_variants

def _specific_detected_variants(detected_variants: set) -> set:
    specific = {v for v in detected_variants if v and v != 'PLAIN'}
    if len(specific) <= 1:
        return detected_variants
    word_sets = {v: set(str(v).upper().split()) for v in specific}
    kept = set()
    for v, vset in word_sets.items():
        is_subset_of_other = any((v != other and vset < word_sets[other] for other in word_sets))
        if not is_subset_of_other:
            kept.add(v)
    if 'PLAIN' in detected_variants:
        kept.add('PLAIN')
    return kept if kept else detected_variants

def filter_items_by_variant(items: list, detected_variants: set, reference_items: list=None, verbose=False) -> list:
    if verbose:
        print('\n' + '=' * 80)
        print('STEP 2: VARIANT-BASED FILTERING')
        print('=' * 80)
        print(f'Detected variants in input: {sorted(detected_variants)}')
        print(f'Total products before filtering: {len(items)}')
    specific_variants = _specific_detected_variants(detected_variants)
    if verbose and specific_variants != detected_variants:
        print(f'Narrowed to most-specific: {sorted(specific_variants)}')

    def _partition(variant_set):
        vmatches, plains = ([], [])
        for item in items:
            iv = item.get('variant', 'PLAIN').upper()
            if iv == 'PLAIN':
                plains.append(item)
            elif iv in variant_set:
                vmatches.append(item)
        return (vmatches, plains)
    variant_matches, plain_products = _partition(specific_variants)
    if not variant_matches and specific_variants != detected_variants:
        if verbose:
            print(f'[WARN] No items for specific variants -- falling back to full detected set')
        variant_matches, plain_products = _partition(detected_variants)
    if variant_matches:
        filtered = variant_matches
        if verbose:
            print(f'\n[OK] Found {len(variant_matches)} specific variant matches')
            print(f'  Excluding {len(plain_products)} PLAIN products')
    elif reference_items and reference_items is not items:
        ref_matches, ref_plains = ([], [])
        for item in reference_items:
            iv = item.get('variant', 'PLAIN').upper()
            if iv == 'PLAIN':
                ref_plains.append(item)
            elif iv in specific_variants or iv in detected_variants:
                ref_matches.append(item)
        if ref_matches:
            filtered = ref_matches
            if verbose:
                print(f'\n[OK] Recovered {len(ref_matches)} variant matches from full brand list')
        else:
            filtered = plain_products if plain_products else ref_plains
            if verbose:
                print(f'\n[WARN] No variant matches; using {len(filtered)} PLAIN as fallback')
    else:
        filtered = plain_products
        if verbose:
            print(f'\n[WARN] No variant matches; using {len(plain_products)} PLAIN as fallback')
    if verbose:
        print(f'\nProducts after variant filtering: {len(filtered)}')
        if variant_matches:
            variants_kept = {item.get('variant', 'PLAIN') for item in filtered}
            for var in sorted(variants_kept):
                count = sum((1 for i in filtered if i.get('variant', 'PLAIN') == var))
                print(f'  - {var}: {count} products')
    if not filtered:
        return items
    return filtered

def filter_by_subvariant_in_product_name(items: list, detected_sub_variants: set, verbose=False) -> list:
    numeric_sub_variants = {sv for sv in detected_sub_variants if sv != 'PLAIN' and '/' not in str(sv)}
    if not numeric_sub_variants:
        return items
    matches_with_scores = []
    for item in items:
        product_name = item['product'].upper()
        best_score = 0
        matched_svs = set()
        for sub_var in numeric_sub_variants:
            sub_str = str(sub_var)
            score = 0
            if re.search('(?<![\\d.])' + re.escape(sub_str) + '\\s*(?:MG|GM|ML|MCG|IU)\\b', product_name):
                score = 100
            elif re.search('(?<![\\d.])' + re.escape(sub_str) + '(?![\\d.])', product_name):
                if not re.search('(?<![\\d.])' + re.escape(sub_str) + '\\s*[X*×]', product_name):
                    score = 90
            elif re.search('[A-Z]' + re.escape(sub_str) + '\\b', product_name):
                score = 75
            elif sub_str in product_name:
                is_false_match = False
                for m in re.finditer('(?<![\\d.])' + re.escape(sub_str) + '(?![\\d.])', product_name):
                    after = product_name[m.end():m.end() + 3]
                    if after and after[0].isdigit() and (not re.match('\\s*(?:MG|GM|ML|MCG|IU|G|L|KG)\\b', after)):
                        is_false_match = True
                        break
                    before = product_name[max(0, m.start() - 3):m.start()]
                    if before and before[-1].isdigit() and (not re.match('(?<![\\d.])\\d+$', before)):
                        is_false_match = True
                        break
                if not is_false_match:
                    score = 50
            if score > best_score:
                best_score = score
                matched_svs = {sub_var}
            elif score == best_score and score > 0:
                matched_svs.add(sub_var)
        if best_score >= 75:
            matches_with_scores.append((item, best_score, matched_svs))
    if matches_with_scores:
        matches_with_scores.sort(key=lambda x: x[1], reverse=True)
        best_score = matches_with_scores[0][1]
        return [item for item, score, _ in matches_with_scores if score == best_score]
    return items

def calculate_product_name_priority_score(product_name: str, input_name: str, detected_variants: set=None, detected_sub_variants: set=None) -> int:
    _TOKEN_SPLIT = '[\\s\\-+/]+'
    input_upper = input_name.upper().strip()
    pack_positions = get_all_pack_positions(input_upper)
    input_words_raw = {w for w in re.split(_TOKEN_SPLIT, input_upper) if w}
    input_words = set()
    for w in input_words_raw:
        is_pack_word = False
        for m in re.finditer('\\b' + re.escape(w) + '\\b', input_upper):
            if is_position_in_pack_pattern(m.start(), m.end(), pack_positions):
                is_pack_word = True
                break
        if not is_pack_word:
            input_words.add(w)
    input_raw_numbers = set(re.findall('\\d+(?:\\.\\d+)?', input_upper))
    input_numbers = set()
    for num in input_raw_numbers:
        is_pack_num = False
        for m in re.finditer('(?<![\\d.])' + re.escape(num) + '(?![\\d.])', input_upper):
            if is_position_in_pack_pattern(m.start(), m.end(), pack_positions):
                is_pack_num = True
                break
        if not is_pack_num:
            input_numbers.add(num)
    product_upper = product_name.upper().strip()
    product_words = {w for w in re.split(_TOKEN_SPLIT, product_upper) if w}
    _PACK_SHAPE = re.compile('^\\d+[Xx*]\\d+[A-Z]?$|^\\d+[STCP]$|^\\d+PCS?$|^\\d+NOS?$', re.IGNORECASE)
    _UNIT_SUFFIX = re.compile('^(?:MG|ML|GM|MCG|IU|KG)\\.?$', re.IGNORECASE)

    def _is_structural(token: str) -> bool:
        if _PACK_SHAPE.match(token) or _UNIT_SUFFIX.match(token):
            return True
        bare = token.rstrip('.')
        if bare in FORM_HINT_TOKENS:
            return True
        glued = re.match('^\\d+([A-Z]+)$', token)
        if glued and glued.group(1) in FORM_HINT_TOKENS:
            return True
        return False
    score = 0
    extra_words = product_words - input_words
    if extra_words:
        extra_words = {w for w in extra_words if not _is_structural(w)}
        numeric_re = re.compile('^\\d+(?:\\.\\d+)?(?:/\\d+(?:\\.\\d+)?)*$')
        extra_numeric = {w for w in extra_words if numeric_re.match(w)}
        extra_numeric_penalized = {w for w in extra_numeric if w not in input_numbers}
        extra_text = extra_words - extra_numeric
        score -= len(extra_text) * 40
        score -= len(extra_numeric_penalized) * 25
    if detected_variants:
        for var in detected_variants:
            if var != 'PLAIN' and var in product_upper:
                score += 100
    if detected_sub_variants:
        for sv in detected_sub_variants:
            if sv != 'PLAIN' and sv in product_upper:
                score += 50
    product_sub_var = 'PLAIN'
    strength_match = re.search('(\\d+(?:\\.\\d+)?)\\s*(?:MG|GM|ML|MCG|IU|G|KG)\\b', product_upper)
    if strength_match:
        product_sub_var = strength_match.group(1)
    else:
        form_words = ['TABLET', 'TABLETS', 'CAPSULE', 'CAPSULES', 'TAB', 'TABS', 'CAP', 'CAPS', 'INJECTION', 'INJ', 'SYRUP', 'SYP', 'SUSPENSION', 'SUSP', 'DROP', 'DROPS', 'CREAM', 'OINTMENT', 'GEL', 'LOTION', 'POWDER', 'GRANULES', 'GRANUL']
        for form in form_words:
            pattern = '(\\d+(?:\\.\\d+)?)\\s+' + re.escape(form) + '\\b'
            m = re.search(pattern, product_upper)
            if m:
                product_sub_var = m.group(1)
                break
    if product_sub_var != 'PLAIN' and product_sub_var not in ('', 'PLAIN'):
        strength_num = product_sub_var
        for m in re.finditer('(?<![\\d.])' + re.escape(strength_num) + '(?![\\d.])', input_upper):
            if is_position_in_pack_pattern(m.start(), m.end(), pack_positions):
                continue
            after = input_upper[m.end():m.end() + 3]
            if re.match('\\s*[STCP]', after) or re.match('\\s*[X*×]', after):
                continue
            score += 75
            break
    if product_sub_var != 'PLAIN' and product_sub_var not in ('', 'PLAIN'):
        strength_num = product_sub_var
        typical_tablet_counts = {'10', '15', '20', '30', '14', '21', '28', '7', '5', '30'}
        for m in re.finditer('(?<![\\d.])' + re.escape(strength_num) + '\\s*[X*×]\\s*(?:' + '|'.join(typical_tablet_counts) + ')(?![\\d.])', input_upper):
            before = input_upper[max(0, m.start() - 5):m.start()]
            after = input_upper[m.end():m.end() + 5]
            if not re.search('\\d\\s*$', before):
                if not re.match('\\s*[STCP]', after):
                    score += 100
                    break
    if 'SEMI' not in input_words and 'SEMI' in product_upper:
        score -= 20
    return score

def prioritize_by_product_name_match(items: list, input_name: str, verbose=False) -> list:
    if len(items) <= 1:
        return items
    scored_items = [(item, calculate_product_name_priority_score(item['product'], input_name)) for item in items]
    scored_items.sort(key=lambda x: x[1], reverse=True)
    if verbose:
        print(f'\n  Product name prioritization scores:')
        for item, score in scored_items:
            print(f"    {score}: {item['product'][:50]}")
    return [item for item, _ in scored_items]

def filter_items_by_sub_variant(items: list, detected_sub_variants: set, input_name: str='', reference_items: list=None, verbose=False) -> list:
    if verbose:
        print('\n' + '=' * 80)
        print('STEP 3: SUB-VARIANT-BASED FILTERING')
        print('=' * 80)
        print(f'Detected sub-variants: {sorted(detected_sub_variants)}')
    specific_detected = {sv for sv in detected_sub_variants if sv != 'PLAIN'}
    match_targets = specific_detected if specific_detected else set(detected_sub_variants)
    detected_forms = detect_dosage_form_in_input(input_name, verbose=False) if input_name else set()
    alpha_detected = {sv for sv in specific_detected if re.search('[A-Z]', str(sv)) and (not is_numeric_subvariant_like(str(sv)))}
    combo_dosages = {sv for sv in detected_sub_variants if '/' in str(sv)}
    if combo_dosages:
        combo_matches = []
        for item in items:
            for combo in combo_dosages:
                if product_contains_combination(item['product'], combo):
                    combo_matches.append(item)
                    break
        if combo_matches:
            if verbose:
                print(f'\n[OK] {len(combo_matches)} literal combo matches')
            return combo_matches
        ordered_components = []
        for combo in combo_dosages:
            parts = [p.strip() for p in str(combo).split('/') if p.strip()]
            for idx, part in enumerate(parts):
                if idx >= len(ordered_components):
                    ordered_components.append(set())
                ordered_components[idx].add(part.upper())
        input_combo_len = len(ordered_components)
        master_combo_lengths = set()
        for item in items:
            subv = str(item.get('sub_variant', 'PLAIN')).strip().upper()
            if '/' in subv:
                master_combo_lengths.add(len(subv.split('/')))
        is_2to3_combo = input_combo_len == 2 and any((l >= 3 for l in master_combo_lengths))
        if is_2to3_combo:
            if verbose:
                print(f'\n[2->3 COMBO] Input has 2 components, master has 3+ -- will match middle component')
            middle_components = set()
            for item in items:
                subv = str(item.get('sub_variant', 'PLAIN')).strip().upper()
                if '/' in subv:
                    parts = subv.split('/')
                    if len(parts) >= 3:
                        middle_components.add(parts[1].upper())
            if middle_components:
                comp_col_matches = [item for item in items if len(str(item.get('sub_variant', 'PLAIN')).strip().upper().split('/')) >= 3 if str(item.get('sub_variant', 'PLAIN')).strip().upper().split('/')[1].upper() in middle_components]
                if comp_col_matches:
                    if verbose:
                        print(f'\n[OK] 2->3 combo -- using middle component {sorted(middle_components)} (sub_variant column)')
                    middle_values = []
                    for item in comp_col_matches:
                        subv = str(item.get('sub_variant', 'PLAIN')).strip().upper()
                        parts = subv.split('/')
                        if len(parts) >= 3:
                            try:
                                middle_values.append((float(parts[1]), item))
                            except ValueError:
                                pass
                    if middle_values:
                        middle_values.sort(key=lambda x: x[0])
                        lowest_middle = middle_values[0][0]
                        lowest_items = [item for val, item in middle_values if val == lowest_middle]
                        if verbose:
                            print(f'\n[DEFAULT] Picking lowest middle component: {lowest_middle}')
                        if input_name:
                            return prioritize_by_product_name_match(lowest_items, input_name, verbose=verbose)
                        return lowest_items
        for position, components in enumerate(ordered_components):
            if not components:
                continue
            comp_col_matches = [item for item in items if str(item.get('sub_variant', 'PLAIN')).strip().upper() in components]
            if comp_col_matches:
                if verbose:
                    print(f'\n[OK] No literal combo -- using component #{position + 1} {sorted(components)} (sub_variant column)')
                if input_name:
                    return prioritize_by_product_name_match(comp_col_matches, input_name, verbose=verbose)
                return comp_col_matches
            comp_name_matches = []
            for item in items:
                pname = item['product'].upper()
                for comp in components:
                    if re.search('\\b' + re.escape(str(comp)) + '\\b', pname):
                        comp_name_matches.append(item)
                        break
            if comp_name_matches:
                if verbose:
                    print(f'\n[OK] No literal combo -- using component #{position + 1} {sorted(components)} (product name)')
                if input_name:
                    return prioritize_by_product_name_match(comp_name_matches, input_name, verbose=verbose)
                return comp_name_matches
        if len(ordered_components) >= 2:
            master_combo_lengths = set()
            for item in items:
                subv = str(item.get('sub_variant', 'PLAIN')).strip().upper()
                if '/' in subv:
                    master_combo_lengths.add(len(subv.split('/')))
            if any((l > len(ordered_components) for l in master_combo_lengths)):
                last_components = ordered_components[-1]
                if last_components:
                    comp_col_matches = [item for item in items if str(item.get('sub_variant', 'PLAIN')).strip().upper() in last_components]
                    if comp_col_matches:
                        if verbose:
                            print(f'\n[OK] Master has longer combo -- using last component {sorted(last_components)} (sub_variant column)')
                        if input_name:
                            return prioritize_by_product_name_match(comp_col_matches, input_name, verbose=verbose)
                        return comp_col_matches
                    comp_name_matches = []
                    for item in items:
                        pname = item['product'].upper()
                        for comp in last_components:
                            if re.search('\\b' + re.escape(str(comp)) + '\\b', pname):
                                comp_name_matches.append(item)
                                break
                    if comp_name_matches:
                        if verbose:
                            print(f'\n[OK] Master has longer combo -- using last component {sorted(last_components)} (product name)')
                        if input_name:
                            return prioritize_by_product_name_match(comp_name_matches, input_name, verbose=verbose)
                        return comp_name_matches
                if len(ordered_components) == 2 and len(comp_name_matches) > 1:
                    middle_values = []
                    for item in comp_name_matches:
                        subv = str(item.get('sub_variant', 'PLAIN')).strip().upper()
                        parts = subv.split('/')
                        if len(parts) == 3 and parts[0] == list(ordered_components[0])[0] and (parts[2] == list(ordered_components[-1])[0]):
                            try:
                                middle_values.append((float(parts[1]), item))
                            except ValueError:
                                pass
                    if middle_values:
                        middle_values.sort(key=lambda x: x[0])
                        lowest_middle = middle_values[0][0]
                        lowest_items = [item for val, item in middle_values if val == lowest_middle]
                        if verbose:
                            print(f'\n[DEFAULT] 2-component input matches multiple 3-component combos -- picking lowest middle component: {lowest_middle}')
                        if input_name:
                            return prioritize_by_product_name_match(lowest_items, input_name, verbose=verbose)
                        return lowest_items
    if alpha_detected:
        alpha_matches = [i for i in items if i.get('sub_variant', 'PLAIN').upper() in alpha_detected]
        if alpha_matches:
            narrowed = filter_by_subvariant_in_product_name(alpha_matches, detected_sub_variants, verbose=verbose)
            if input_name:
                return prioritize_by_product_name_match(narrowed, input_name, verbose=verbose)
            return narrowed
    exact_matches = []
    plain_products = []
    numeric_detected = {sv for sv in detected_sub_variants if sv != 'PLAIN' and is_numeric_subvariant_like(str(sv))}
    for item in items:
        item_sv = item.get('sub_variant', 'PLAIN').upper()
        if item_sv in match_targets:
            exact_matches.append(item)
        elif item_sv == 'PLAIN':
            plain_products.append(item)
    if exact_matches:
        if not specific_detected and detected_forms:
            form_matches = [item for item in exact_matches if detect_dosage_form_in_product(item['product']) & detected_forms or forms_share_group(detect_dosage_form_in_product(item['product']), detected_forms)]
            if form_matches:
                if input_name:
                    return prioritize_by_product_name_match(form_matches, input_name, verbose=verbose)
                return form_matches
        numeric_exact = {str(i.get('sub_variant', '')).strip() for i in exact_matches if is_numeric_subvariant_like(str(i.get('sub_variant', '')))}
        if input_name and len(numeric_exact) > 1:
            up = normalize_ocr_numeric_noise(input_name)

            def _pos(sub):
                m = re.search('(?<![\\d.])' + re.escape(str(sub)) + '(?![\\d.])', up)
                return m.start() if m else 10 ** 6
            best_pos = min((_pos(s) for s in numeric_exact))
            earliest = [i for i in exact_matches if _pos(str(i.get('sub_variant', '')).strip()) == best_pos]
            if earliest and len(earliest) < len(exact_matches):
                exact_matches = earliest
        if input_name:
            return prioritize_by_product_name_match(exact_matches, input_name, verbose=verbose)
        return exact_matches
    if numeric_detected:
        first_comp = [item for item in items if '/' in str(item.get('sub_variant', '')) and str(item.get('sub_variant', '')).split('/')[0].strip().upper() in numeric_detected]
        if first_comp and len(first_comp) < len(items):
            first_comp_sub_vars = {str(item.get('sub_variant', '')).strip().upper() for item in first_comp}
            if len(first_comp_sub_vars) > 1:
                if verbose:
                    print(f'\n[WARN] Multiple combo sub-variants share first component {sorted(numeric_detected)} -- not narrowing')
            else:
                if verbose:
                    print(f"\n[OK] Single number -> combo first-component match: {[i['product'] for i in first_comp]}")
                if input_name:
                    return prioritize_by_product_name_match(first_comp, input_name, verbose=verbose)
                return first_comp
    if numeric_detected:
        name_matches = filter_by_subvariant_in_product_name(items, detected_sub_variants, verbose=verbose)
        if name_matches and len(name_matches) < len(items):
            if input_name:
                return prioritize_by_product_name_match(name_matches, input_name, verbose=verbose)
            return name_matches
    if numeric_detected and plain_products:
        name_matches = filter_by_subvariant_in_product_name(plain_products, detected_sub_variants, verbose=verbose)
        if name_matches and len(name_matches) < len(plain_products):
            return name_matches
    if plain_products:
        return plain_products
    return items

def detect_dosage_form_in_input(input_name: str, verbose=False) -> set:
    input_upper = normalize_ocr_numeric_noise(input_name).upper()
    detected_forms = set()
    form_patterns = [('\\bDROPS?\\b|\\d+DROPS?\\b|DROPS?[A-Z]*\\b', 'DROPS'), ('\\bSUSPENSIONS?\\b|\\bSUSP\\.?\\b|\\bSUSPN\\b|\\bSYRUPS?\\b|\\bSYP\\.?\\b|\\d+SYP\\b', 'SUSPENSION'), ('\\bTABLETS?\\b|\\bTABS?\\.?\\b|\\bTA\\b|\\d+TABLETS?|\\d+TABS?\\b', 'TABLET'), ('\\bCAPSULES?\\b|\\bCAPS?\\.?\\b|\\d+CAPSULES?|\\d+CAPS?\\b', 'CAPSULE'), ('\\bINJECTIONS?\\b|\\bINJ\\.?\\b|\\bING\\.?\\b|\\d+INJ\\b|\\bVIAL\\b|\\bPFS\\b|\\bAMPS?\\b|\\bSYR\\.?\\b|\\d+SYR\\.?\\b|\\bI\\.?V\\.?\\b', 'INJECTION'), ('\\bCREAMS?\\b|\\bCRM\\.?\\b', 'CREAM'), ('\\bOINTMENTS?\\b|\\bOI\\.?\\b|\\bOINT\\.?\\b', 'OINTMENT'), ('\\bGELS?\\b', 'GEL'), ('\\bSOLUTIONS?\\b|\\bSOL\\.?\\b', 'SOLUTION'), ('\\bPOWDERS?\\b|\\bPOW\\.?\\b', 'POWDER'), ('\\bBARS?\\b|\\bSYNDET\\b', 'BAR'), ('\\bGRANULES?\\b|\\bGRANUL\\b|\\bGRANULS\\b', 'GRANULES'), ('\\bSACHETS?\\b', 'SACHET'), ('\\bLOTIONS?\\b|\\bLOT\\.?\\b|\\bLO\\b', 'LOTION'), ('\\bTONICS?\\b', 'TONIC')]
    for pattern, form_name in form_patterns:
        if re.search(pattern, input_upper):
            detected_forms.add(form_name)
    if not {'SUSPENSION', 'DROPS', 'SOLUTION', 'TONIC'} & detected_forms and re.search('\\bSY\\b', input_upper) and re.search('\\b\\d+(?:\\.\\d+)?\\s*ML\\b', input_upper):
        detected_forms.add('SUSPENSION')
    pack_positions = get_all_pack_positions(input_upper)

    def _all_matches_inside_pack(pattern: str) -> bool:
        matches = list(re.finditer(pattern, input_upper))
        if not matches:
            return False
        return all((is_position_in_pack_pattern(m.start(), m.end(), pack_positions) for m in matches))
    if 'CAPSULE' in detected_forms and 'TABLET' in detected_forms:
        tablet_pack_only = _all_matches_inside_pack('\\bTABLETS?\\b|\\bTABS?\\.?\\b|\\bTA\\b|\\d+TABLETS?|\\d+TABS?\\b')
        capsule_pack_only = _all_matches_inside_pack('\\bCAPSULES?\\b|\\bCAPS?\\.?\\b|\\d+CAPSULES?|\\d+CAPS?\\b')
        if tablet_pack_only and (not capsule_pack_only):
            detected_forms.discard('TABLET')
        elif capsule_pack_only and (not tablet_pack_only):
            detected_forms.discard('CAPSULE')
    return detected_forms

def detect_dosage_form_in_product(product_name: str) -> set:
    product_upper = product_name.upper()
    detected_forms = set()
    form_patterns = [('\\bDROPS?\\b', 'DROPS'), ('\\bSUSPENSIONS?\\b|\\bSUSP\\.?\\b|\\bSUSPN\\b|\\bSYRUPS?\\b|\\bSYP\\.?\\b|\\d+SYP\\b|\\bSYR\\.?\\b|\\d+SYR\\.?\\b', 'SUSPENSION'), ('\\bTABLETS?\\b|\\bTABS?\\.?\\b', 'TABLET'), ('\\bCAPSULES?\\b|\\bCAPS?\\.?\\b', 'CAPSULE'), ('\\bINJECTIONS?\\b|\\bINJ\\.?\\b|\\bING\\.?\\b|\\d+INJ\\b|\\bVIAL\\b|\\bPFS\\b|\\bAMPS?\\b|\\bI\\.?V\\.?\\b', 'INJECTION'), ('\\bCREAMS?\\b', 'CREAM'), ('\\bOINTMENTS?\\b', 'OINTMENT'), ('\\bGELS?\\b', 'GEL'), ('\\bSOLUTIONS?\\b|\\bSOL\\b', 'SOLUTION'), ('\\bPOWDERS?\\b', 'POWDER'), ('\\bBARS?\\b|\\bSYNDET\\b', 'BAR'), ('\\bGRANULES?\\b|\\bGRANUL\\b|\\bGRANULS\\b', 'GRANULES'), ('\\bSACHETS?\\b', 'SACHET'), ('\\bLOTIONS?\\b', 'LOTION'), ('\\bTONICS?\\b', 'TONIC')]
    for pattern, form_name in form_patterns:
        if re.search(pattern, product_upper):
            detected_forms.add(form_name)
    return detected_forms
_FORM_GROUPS = {'TABLET': 'solid_oral', 'CAPSULE': 'solid_oral', 'SUSPENSION': 'liquid_oral', 'SYRUP': 'liquid_oral', 'SOLUTION': 'liquid_oral', 'DROPS': 'liquid_oral', 'TONIC': 'liquid_oral', 'INJECTION': 'parenteral', 'CREAM': 'topical', 'GEL': 'topical', 'OINTMENT': 'topical', 'LOTION': 'topical', 'POWDER': 'powder', 'GRANULES': 'powder', 'SACHET': 'powder', 'BAR': 'bar'}

def forms_share_group(product_forms: set, detected_forms: set) -> bool:
    product_groups = {_FORM_GROUPS[f] for f in product_forms if f in _FORM_GROUPS}
    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    return bool(product_groups & detected_groups)

def prioritize_by_dosage_form(items: list, detected_forms: set, verbose=False) -> list:
    if not detected_forms:
        return items
    if verbose:
        print('\n' + '=' * 80)
        print('STEP 4.5: DOSAGE FORM PRIORITIZATION')
        print('=' * 80)
        print(f'Detected forms in input: {sorted(detected_forms)}')
    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    exact_form_matches = []
    other_products = []
    for item in items:
        product_forms = detect_dosage_form_in_product(item['product'])
        if product_forms & detected_forms:
            exact_form_matches.append(item)
        else:
            other_products.append(item)
    if exact_form_matches:
        if verbose:
            print(f'\n[OK] Found {len(exact_form_matches)} products matching dosage form')
        return exact_form_matches
    if detected_groups:
        non_conflicting = []
        excluded = []
        for item in items:
            product_forms = detect_dosage_form_in_product(item['product'])
            product_groups = {_FORM_GROUPS[f] for f in product_forms if f in _FORM_GROUPS}
            if product_groups and product_groups.isdisjoint(detected_groups):
                excluded.append(item)
            else:
                non_conflicting.append(item)
        if non_conflicting and len(non_conflicting) < len(items):
            if verbose:
                print(f'\n[WARN] No exact form match -- excluded {len(excluded)} conflicting-form products')
            return non_conflicting
    return items

def _form_compatible(items: list, detected_forms: set) -> list:
    detected_groups = {_FORM_GROUPS[f] for f in detected_forms if f in _FORM_GROUPS}
    if not detected_groups:
        return list(items)
    keep = []
    for item in items:
        product_groups = {_FORM_GROUPS[f] for f in detect_dosage_form_in_product(item['product']) if f in _FORM_GROUPS}
        if not product_groups or product_groups & detected_groups:
            keep.append(item)
    return keep

def filter_items_by_pack_size(items: list, input_pack_size: str, reference_items: list=None, verbose=False) -> list:
    if verbose:
        print('\n' + '=' * 80)
        print('STEP 5: PACK SIZE FILTERING')
        print('=' * 80)
        print(f"Input pack size: '{input_pack_size or 'NONE'}'")
    products_with_pack = []
    products_without_pack = []
    for item in items:
        pack = item['pack_size'].strip()
        if pack and pack not in ('', 'nan', 'NaN', 'None'):
            pack_num = parse_pack_size_value(pack)
            if pack_num is not None:
                products_with_pack.append((item, pack_num))
            else:
                products_without_pack.append(item)
        else:
            products_without_pack.append(item)
    if verbose:
        print(f'  Products with pack size:    {len(products_with_pack)}')
        print(f'  Products without pack size: {len(products_without_pack)}')
    if not input_pack_size:
        if products_without_pack:
            if verbose:
                print(f'\n[OK] Input has no pack -- returning {len(products_without_pack)} BASE products')
            return products_without_pack
        if products_with_pack:
            products_with_pack.sort(key=lambda x: x[1])
            return [item for item, _ in products_with_pack[:1]]
        return items
    try:
        input_pack_num = float(input_pack_size)
    except ValueError:
        return products_without_pack if products_without_pack else items
    exact_matches = [item for item, pn in products_with_pack if pn == input_pack_num]
    if exact_matches:
        return exact_matches
    if products_without_pack:
        if verbose:
            print(f'\n[OK] No exact pack match -- falling back to BASE products')
        return products_without_pack
    if products_with_pack:
        closest = min(products_with_pack, key=lambda x: abs(x[1] - input_pack_num))
        closest_diff = abs(closest[1] - input_pack_num)
        if closest_diff <= input_pack_num * 0.5 or closest_diff <= 20:
            if verbose:
                print(f'\n[OK] No exact pack match and no base products -- using closest pack {closest[1]} (diff={closest_diff})')
            return [closest[0]]
    return items

def _clean_input_for_fuzzy(input_name: str) -> str:
    s_upper = normalize_ocr_numeric_noise(input_name).upper()
    pack_positions = get_all_pack_positions(s_upper)
    _TOKEN_SPLIT = '[\\s\\-+/]+'
    tokens = []
    for m in re.finditer('\\S+', s_upper):
        token = m.group()
        is_pack = any((ps <= m.start() < pe or ps < m.end() <= pe for ps, pe in pack_positions))
        if not is_pack:
            tokens.append(token)
    return ' '.join(tokens)

def rank_local_candidates(items: list, input_name: str, detected_variants: set=None, detected_sub_variants: set=None) -> list:
    if not items:
        return []
    ranked_items = prioritize_by_product_name_match(items, input_name, verbose=False)
    if len(ranked_items) == 1:
        return [(100.0, ranked_items[0])]
    input_norm = norm(input_name)
    input_clean = _clean_input_for_fuzzy(input_name)
    input_tokens_upper = set(re.findall('[A-Z0-9]+', input_clean.upper())) | set(re.findall('[A-Z]+', input_clean.upper()))
    KNOWN_GHOST_TOKENS = {token for candidate in ranked_items for field in ('variant', 'sub_variant') for token in re.findall('[A-Z]+', str(candidate.get(field, '')).upper()) if token != 'PLAIN' and token not in FORM_HINT_TOKENS and (token not in DOSAGE_UNIT_SUFFIXES)}
    scored_items = []
    for item in ranked_items:
        product_norm = norm(item['product'])
        fuzzy_score = max(fuzz.token_sort_ratio(input_clean, product_norm), fuzz.token_set_ratio(input_clean, product_norm), fuzz.ratio(input_clean, product_norm))
        priority_score = calculate_product_name_priority_score(item['product'], input_name, detected_variants, detected_sub_variants)
        alpha_ghost_penalty = 0
        item_sv = str(item.get('sub_variant', 'PLAIN')).strip().upper() or 'PLAIN'
        if item_sv != 'PLAIN' and re.match('^[A-Z]+$', item_sv) and (item_sv not in input_tokens_upper):
            alpha_ghost_penalty = -200
        item_v = str(item.get('variant', 'PLAIN')).strip().upper() or 'PLAIN'
        if item_v != 'PLAIN' and re.match('^[A-Z ]+$', item_v):
            v_words = set(item_v.split())
            if not v_words.issubset(input_tokens_upper):
                alpha_ghost_penalty -= 150
        product_tokens = set(re.findall('[A-Z]+', item['product'].upper()))
        ghost_in_product = (product_tokens & KNOWN_GHOST_TOKENS) - input_tokens_upper
        if ghost_in_product:
            alpha_ghost_penalty -= 120 * len(ghost_in_product)
        score = fuzzy_score + priority_score + alpha_ghost_penalty
        scored_items.append((score, item))
    scored_items.sort(key=lambda x: x[0], reverse=True)
    return scored_items

def choose_best_local_candidate(items: list, input_name: str, detected_variants: set=None, detected_sub_variants: set=None) -> dict:
    ranked = rank_local_candidates(items, input_name, detected_variants, detected_sub_variants)
    return ranked[0][1] if ranked else None

def has_ambiguous_missing_strength(items: list, detected_sub_variants: set) -> bool:
    specific_detected = {sv for sv in detected_sub_variants if sv != 'PLAIN'}
    item_sub_variants = {str(item.get('sub_variant', 'PLAIN')).strip().upper() for item in items if str(item.get('sub_variant', 'PLAIN')).strip().upper() not in ('', 'PLAIN')}
    if len(item_sub_variants) <= 1:
        return False
    if not all((is_numeric_subvariant_like(sub_variant) for sub_variant in item_sub_variants)):
        return False
    item_variants = {str(item.get('variant', 'PLAIN')).strip().upper() for item in items}
    if len(item_variants) == 1:
        combo_subs = [sv for sv in item_sub_variants if '/' in sv]
        if len(combo_subs) >= 2:
            first_comps = {sv.split('/')[0].strip().upper() for sv in combo_subs}
            if len(first_comps) == 1:
                detected_combos = {sv for sv in detected_sub_variants if '/' in str(sv)}
                if not detected_combos:
                    return True
    if specific_detected and len(item_variants) == 1:
        combo_subs = [sv for sv in item_sub_variants if '/' in sv]
        if len(combo_subs) >= 2:
            first_comps = {sv.split('/')[0].strip().upper() for sv in combo_subs}
            if specific_detected.issubset(first_comps):
                detected_combos = {sv for sv in detected_sub_variants if '/' in str(sv)}
                if not detected_combos:
                    return True
    return False
TOP_N_SUGGESTIONS = 3
NO_BRAND_RECOVERY_MAX_EDITS = 1

def nearest_brand_within_edits(input_name, suggestions, max_edits=NO_BRAND_RECOVERY_MAX_EDITS):
    query = compact_brand_token(extract_brand_like_query(input_name))
    if not query:
        return (None, None)
    best, best_dist = (None, None)
    for s in suggestions or []:
        cand = compact_brand_token(s.get('brand', ''))
        if not cand:
            continue
        dist = Levenshtein.distance(query, cand)
        if best_dist is None or dist < best_dist:
            best, best_dist = (s, dist)
    if best is not None and best_dist is not None and (best_dist <= max_edits):
        return (best, best_dist)
    return (None, best_dist)

def extract_pack_notation(input_name: str) -> str:
    s_upper = normalize_ocr_numeric_noise(input_name).upper()
    m = re.search('\\b(\\d+)\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)(?=[A-Z]?\\b)', s_upper)
    if m:
        n1, n2 = m.groups()
        after = s_upper[m.end():m.end() + 5]
        form_suffix = ''
        if after:
            form_match = re.match('^([STCP]|PCS?|AB)', after)
            if form_match:
                form_suffix = form_match.group(1)
        return f'{n1}X{n2}{form_suffix}'
    m = re.search('\\b(\\d+)\\s*[X*×]\\s*(\\d+(?:\\.\\d+)?)\\s*(PCS?)\\b', s_upper)
    if m:
        n1, n2, suffix = m.groups()
        return f'{n1}X{n2}{suffix}'
    m = re.search('\\b(\\d+)\\s*([STCP])\\b', s_upper)
    if m:
        return f'{m.group(1)}{m.group(2)}'
    m = re.search('\\b(\\d+)\\s*(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\\b', s_upper)
    if m:
        return f'{m.group(1)}T'
    m = re.search('\\b(?:TA|TABS?|TABLETS?|STRIPS?|CAPS?|CAPSULES?)\\s*(\\d+)\\b', s_upper)
    if m:
        return f'{m.group(1)}T'
    return ''

def format_output_with_pack_size(product_name: str, input_name: str='', master_pack_size: str='') -> str:
    return product_name

def make_result(output, product_code='', status='MATCHED', confidence='HIGH', candidate_count=0, suggestions=None):
    return {'output': output, 'product_code': product_code, 'status': status, 'confidence': confidence, 'candidate_count': int(candidate_count or 0), 'suggestions': suggestions or []}

def _suggestions_from_ranked(ranked, top_n=TOP_N_SUGGESTIONS):
    out = []
    for score, item in ranked[:top_n]:
        out.append({'product': item.get('product', ''), 'product_code': item.get('product_code', ''), 'score': round(float(score), 1)})
    return out

def build_item_suggestions(items, input_name, top_n=TOP_N_SUGGESTIONS):
    if not items:
        return []
    return _suggestions_from_ranked(rank_local_candidates(items, input_name), top_n)

def suggest_nearest_brands(input_name, brand_map, top_n=TOP_N_SUGGESTIONS):
    if not brand_map:
        return []
    query = norm(extract_brand_like_query(input_name)) or norm(input_name)
    if not query:
        return []
    norm_to_brand = {}
    for brand in brand_map.keys():
        norm_to_brand.setdefault(norm(brand), brand)
    matches = process.extract(query, list(norm_to_brand.keys()), scorer=fuzz.token_sort_ratio, limit=top_n)
    brand_query = extract_brand_like_query(input_name)
    out = []
    for norm_brand, score, _ in matches:
        brand = norm_to_brand[norm_brand]
        if is_substring_brand_reject(brand_query, brand):
            continue
        brand_items = brand_map.get(brand, [])
        best = choose_best_local_candidate(brand_items, input_name) if brand_items else None
        out.append({'brand': brand, 'product': best['product'] if best else brand, 'product_code': best.get('product_code', '') if best else '', 'score': round(float(score), 1)})
    return out

def format_suggestions(suggestions):
    if not suggestions:
        return ''
    parts = []
    for s in suggestions:
        code = f" ({s['product_code']})" if s.get('product_code') else ''
        parts.append(f"{s.get('product', '')}{code} [{s.get('score', '')}]")
    return ' | '.join(parts)

def filter_items_by_name_strength(items: list, input_name: str, verbose=False) -> list:
    if len(items) <= 1:
        return items
    txt = collapse_spaced_thousands(normalize_ocr_numeric_noise(input_name))
    dose_pairs = re.findall('(\\d+(?:\\.\\d+)?)\\s*(MG|MCG|GM|IU|ML|UNIT)\\b', txt)
    if not dose_pairs:
        return items
    matched = []
    for it in items:
        name = collapse_spaced_thousands(str(it.get('product', '')).upper())
        if any((re.search('(?<![\\d.])' + re.escape(num) + '\\s*' + unit, name) for num, unit in dose_pairs)):
            matched.append(it)
    if matched and len(matched) < len(items):
        if verbose:
            print(f'  [NAME-STRENGTH] dose {dose_pairs} -> {len(items)}->{len(matched)} items')
        return matched
    return items

def filter_items_prefer_exact_brand_token(items: list, brand: str, input_name: str, verbose=False) -> list:
    if len(items) <= 1 or not brand:
        return items
    brand_first = brand.upper().replace('-', ' ').split()[0] if brand.split() else ''
    if not brand_first:
        return items
    input_tokens = set(re.findall('[A-Z]+', normalize_ocr_numeric_noise(input_name)))
    exact, glued = ([], [])
    for it in items:
        toks = re.findall('[A-Z]+', str(it.get('product', '')).upper())
        first = toks[0] if toks else ''
        if first != brand_first and first.startswith(brand_first) and (len(first) > len(brand_first)) and (first not in input_tokens) and (first[len(brand_first):] not in input_tokens):
            glued.append(it)
        else:
            exact.append(it)
    if exact and glued:
        if verbose:
            print(f"  [EXACT-BRAND] dropping glued sub-brand items ({[i['product'] for i in glued]})")
        return exact
    return items

def process_product(input_name: str, brand_map: dict, all_variants: set=None, all_sub_variants: set=None, client=None, verbose=False, _lower_variant_retry=False, forced_brand: str=None, auto_detect: bool=False) -> dict:
    input_name = str(input_name or '').strip()
    exact_items = _get_exact_catalog_index(brand_map).get(norm(input_name), [])
    if len(exact_items) == 1:
        item = exact_items[0]
        return make_result(item['product'], item.get('product_code', ''), status='MATCHED', confidence='HIGH', candidate_count=1, suggestions=[{'product': item['product'], 'product_code': item.get('product_code', ''), 'score': 100.0}])
    if len(exact_items) > 1:
        return make_result('NO_CLEAR_MATCH', '', status='AMBIGUOUS_EXACT_NAME', confidence='NONE', candidate_count=len(exact_items), suggestions=[{'product': item['product'], 'product_code': item.get('product_code', ''), 'score': 100.0} for item in exact_items[:TOP_N_SUGGESTIONS]])
    form_items = _get_form_catalog_index(brand_map).get(_canonical_form_name(input_name), [])
    if len(form_items) == 1:
        item = form_items[0]
        return make_result(item['product'], item.get('product_code', ''), status='MATCHED', confidence='HIGH', candidate_count=1, suggestions=[{'product': item['product'], 'product_code': item.get('product_code', ''), 'score': 100.0}])
    if len(form_items) > 1:
        return make_result('NO_CLEAR_MATCH', '', status='AMBIGUOUS_FORM_NAME', confidence='NONE', candidate_count=len(form_items), suggestions=[{'product': item['product'], 'product_code': item.get('product_code', ''), 'score': 100.0} for item in form_items[:TOP_N_SUGGESTIONS]])
    input_name = clean_duplicate_words(input_name)
    forced_brand_variant_tags = get_brand_variant_tags(brand_map, forced_brand) if forced_brand else None
    input_name = strip_parenthetical_noise(input_name, forced_brand_variant_tags)
    input_name = desegment_ocr_input(input_name, brand_map)
    if verbose:
        print('\n' + '=' * 80)
        print('PRODUCT MATCHING')
        print('=' * 80)
        print(f"Input: '{input_name}'")
    brand = None
    items = []
    if forced_brand:
        forced_upper = forced_brand.strip().upper()
        if forced_upper in brand_map:
            brand = forced_upper
            items = list(brand_map[brand])
            if verbose:
                print(f"\n  [forced_brand] Using hint '{brand}' -- {len(items)} candidates")
        else:
            matched_key = next((k for k in brand_map if k.upper() == forced_upper), None)
            if matched_key:
                brand = matched_key
                items = list(brand_map[brand])
                if verbose:
                    print(f"\n  [forced_brand] Using hint '{brand}' (normalised) -- {len(items)} candidates")
            elif auto_detect:
                if verbose:
                    print(f"  [forced_brand] Hint '{forced_brand}' not in brand_map -> auto-detecting brand")
                brand, items = find_best_brand_for_input(input_name, brand_map, verbose=verbose)
            else:
                if verbose:
                    print(f"  [forced_brand] Hint '{forced_brand}' not in brand_map -> NO_CLEAR_MATCH")
                return make_result('NO_CLEAR_MATCH', '', status='NO_BRAND', confidence='NONE', candidate_count=0, suggestions=[])
    elif auto_detect:
        if verbose:
            print(f"  [auto_detect] No brand hint -- running auto brand detection for '{input_name[:50]}'")
        brand, items = find_best_brand_for_input(input_name, brand_map, verbose=verbose)
    else:
        if verbose:
            print(f"  [forced_brand] No brand hint for '{input_name[:50]}' -> NO_CLEAR_MATCH")
        return make_result('NO_CLEAR_MATCH', '', status='NO_BRAND', confidence='NONE', candidate_count=0, suggestions=[])
    if not brand:
        suggestions = suggest_nearest_brands(input_name, brand_map)
        top, dist = nearest_brand_within_edits(input_name, suggestions)
        if top and top.get('product'):
            if verbose:
                print(f"\nRECOVERED (no brand) -> '{top['product']}' (brand '{top.get('brand', '')}' within {dist} edit(s))")
            return make_result(top['product'], top.get('product_code', ''), status='RECOVERED_NO_BRAND', confidence='LOW', candidate_count=0, suggestions=suggestions)
        if verbose:
            print(f'\nNO_CLEAR_MATCH: no brand within {NO_BRAND_RECOVERY_MAX_EDITS} edits (nearest gap {dist})')
        return make_result('NO_CLEAR_MATCH', '', status='NO_BRAND', confidence='NONE', candidate_count=0, suggestions=suggestions)
    _input_forms = detect_dosage_form_in_input(input_name, verbose=False)
    items = prioritize_by_dosage_form(items, _input_forms, verbose=verbose)
    compound_tokens = extract_brand_suffix_tokens(input_name, brand)
    compound_token_digits = get_compound_token_digits(compound_tokens)
    compound_token_letters = set()
    for token in compound_tokens:
        letters = re.sub('\\d', '', token)
        if letters:
            compound_token_letters.add(letters)
    if compound_tokens:
        if verbose:
            print(f'\n  Compound tokens: {compound_tokens}')
        items = filter_items_by_compound_tokens(items, compound_tokens, verbose=verbose)
    items = filter_items_by_product_name_abbrev(input_name, brand, items, verbose=verbose)
    items = filter_items_prefer_exact_brand_token(items, brand, input_name, verbose=verbose)

    def _is_semi_item(it):
        return bool(re.search('\\b(?:SEMI|HALF)\\b', str(it.get('product', '')).upper()))
    _input_has_semi = bool(re.search('\\b(?:SEMI|HALF)\\b', normalize_ocr_numeric_noise(input_name)))
    _semi_items = [it for it in items if _is_semi_item(it)]
    _plain_items = [it for it in items if not _is_semi_item(it)]
    if _semi_items and _plain_items:
        items = _semi_items if _input_has_semi else _plain_items
        if verbose:
            print(f'\n  SEMI/HALF prefix discriminator: input_has_semi={_input_has_semi} -> kept {len(items)} of {len(_semi_items) + len(_plain_items)} SKUs')
    brand_scope_items = list(items)
    available_variants, available_sub_variants, _ = extract_available_attributes_from_items(brand_scope_items)
    normalized_for_qual = normalize_ocr_numeric_noise(input_name)
    orphan_quals = find_unmatched_qualifier_tokens(normalized_for_qual, brand, available_variants, available_sub_variants, brand_items=brand_scope_items)
    if compound_token_letters:
        orphan_quals = [q for q in orphan_quals if q not in compound_token_letters]
    if orphan_quals:
        numeric_orphan = any((any((ch.isdigit() for ch in o)) for o in orphan_quals))
        if not _lower_variant_retry and (not numeric_orphan):
            stripped = input_name
            for o in orphan_quals:
                stripped = re.sub('\\b' + re.escape(o) + '\\b', ' ', stripped)
            stripped = re.sub('\\s+', ' ', stripped).strip()
            lower_vars = detect_variants_in_input(stripped, brand, available_variants, verbose=False)
            if stripped and stripped != input_name and any((v != 'PLAIN' for v in lower_vars)):
                res = process_product(stripped, brand_map, all_variants, all_sub_variants, client, verbose=verbose, _lower_variant_retry=True, forced_brand=forced_brand, auto_detect=auto_detect)
                if res['status'] == 'MATCHED' and res['output'] != 'NO_CLEAR_MATCH':
                    if verbose:
                        print(f"\nRECOVERED (lower variant): '{brand}' lacks {orphan_quals}; mapped to lower variant '{res['output']}'")
                    res['status'] = 'RECOVERED_LOWER_VARIANT'
                    res['confidence'] = 'LOW'
                return res
        suggestions = build_item_suggestions(brand_scope_items, input_name)
        if verbose:
            print(f"\nNO_CLEAR_MATCH: qualifier {orphan_quals} is not a known variant/sub-brand of '{brand}' in the master")
        return make_result('NO_CLEAR_MATCH', '', status='VARIANT_NOT_IN_MASTER', confidence='NONE', candidate_count=len(brand_scope_items), suggestions=suggestions)
    detected_variants = detect_variants_in_input(input_name, brand, available_variants, compound_token_letters=compound_token_letters, verbose=verbose, brand_items=brand_scope_items)
    _before_var = items
    items = filter_items_by_variant(items, detected_variants, reference_items=brand_scope_items, verbose=verbose)
    if _input_forms and items and (not _form_compatible(items, _input_forms)):
        _pool = _form_compatible(_before_var, _input_forms)
        if _pool:
            _retry = filter_items_by_variant(_pool, detected_variants, reference_items=_pool, verbose=verbose)
            if _retry:
                items = _retry
    detected_sub_variants = detect_sub_variants_in_input(input_name, brand, available_sub_variants, available_variants, compound_token_digits=compound_token_digits, verbose=verbose)
    _before_sv = items
    items = filter_items_by_sub_variant(items, detected_sub_variants, input_name, reference_items=brand_scope_items, verbose=verbose)
    if _input_forms and items and (not _form_compatible(items, _input_forms)):
        _pool = _form_compatible(_before_sv, _input_forms)
        if _pool:
            _retry = filter_items_by_sub_variant(_pool, detected_sub_variants, input_name, reference_items=_pool, verbose=verbose)
            if _retry:
                items = _retry
    _before_ns = items
    items = filter_items_by_name_strength(items, input_name, verbose=verbose)
    if _input_forms and items and (not _form_compatible(items, _input_forms)):
        _pool = _form_compatible(_before_ns, _input_forms)
        if _pool:
            _retry = filter_items_by_name_strength(_pool, input_name, verbose=verbose)
            if _retry:
                items = _retry
    input_pack_size = extract_pack_size_from_input(input_name)
    items = filter_items_by_pack_size(items, input_pack_size, reference_items=brand_scope_items, verbose=verbose)
    if has_ambiguous_missing_strength(items, detected_sub_variants):
        suggestions = build_item_suggestions(items, input_name)
        if verbose:
            print('\nNO_CLEAR_MATCH: ambiguous strength (input does not resolve to one SKU)')
        return make_result('NO_CLEAR_MATCH', '', status='AMBIGUOUS_STRENGTH', confidence='NONE', candidate_count=len(items), suggestions=suggestions)
    if not items:
        suggestions = build_item_suggestions(brand_scope_items, input_name)
        recovered = choose_best_local_candidate(brand_scope_items, input_name, detected_variants, detected_sub_variants)
        if recovered:
            if verbose:
                print(f"\nRECOVERED (over-filtered) -> '{recovered['product']}'")
            return make_result(recovered['product'], recovered.get('product_code', ''), status='RECOVERED_NO_CANDIDATES', confidence='LOW', candidate_count=len(brand_scope_items), suggestions=suggestions)
        if verbose:
            print('\n[FAIL] No products left after filtering')
        return make_result('NO_CLEAR_MATCH', '', status='NO_CANDIDATES', confidence='NONE', candidate_count=0, suggestions=suggestions)
    local_ranked = rank_local_candidates(items, input_name, detected_variants, detected_sub_variants)
    local_best_item = local_ranked[0][1] if local_ranked else None
    local_best_score = local_ranked[0][0] if local_ranked else 0
    local_second_score = local_ranked[1][0] if len(local_ranked) > 1 else None
    if verbose and local_ranked:
        print('\n' + '=' * 80)
        print('LOCAL FINAL CANDIDATE RANKING')
        print('=' * 80)
        for score, item in local_ranked[:5]:
            print(f"  {score:5.1f} -> {item['product']}")
    local_suggestions = _suggestions_from_ranked(local_ranked)
    if local_best_item:
        if len(items) == 1:
            if verbose:
                print('\n[OK] Single candidate -- skipping reranker')
            formatted_output = format_output_with_pack_size(local_best_item['product'], input_name, local_best_item.get('pack_size', ''))
            return make_result(formatted_output, local_best_item.get('product_code', ''), status='MATCHED', confidence='HIGH', candidate_count=len(items), suggestions=local_suggestions)
        if local_best_score >= 90 and (local_second_score is None or local_best_score - local_second_score >= 8):
            if verbose:
                print('\n[OK] Strong local match -- skipping reranker')
            formatted_output = format_output_with_pack_size(local_best_item['product'], input_name, local_best_item.get('pack_size', ''))
            return make_result(formatted_output, local_best_item.get('product_code', ''), status='MATCHED', confidence='HIGH', candidate_count=len(items), suggestions=local_suggestions)
    if local_best_item:
        return make_result(local_best_item['product'], local_best_item.get('product_code', ''), status='RECOVERED_LOCAL_ONLY', confidence='LOW', candidate_count=len(items), suggestions=local_suggestions)
    return make_result('NO_CLEAR_MATCH', '', status='LOCAL_ONLY_UNRESOLVED', confidence='NONE', candidate_count=len(items), suggestions=local_suggestions)

def suggest_local_match(input_name: str, brand_map: dict, verbose=False) -> tuple:
    result = process_product(input_name, brand_map, verbose=verbose, auto_detect=True)
    return (result['output'], result['product_code'])

def review_highlighted_output_file(output_xlsx_path: str, suggestions_xlsx_path: str='') -> str:
    print(f'Loading master product data...')
    df_master = load_master(MASTER_XLSX_PATH)
    brand_map = build_brand_product_map(df_master)
    print(f'Opening highlighted output file: {output_xlsx_path}')
    source_wb = load_workbook(output_xlsx_path)
    source_ws = source_wb[source_wb.sheetnames[0]]
    if not suggestions_xlsx_path:
        base = Path(output_xlsx_path)
        suggestions_xlsx_path = str(base.with_name(base.stem + '_suggestions.xlsx'))
    suggestions_wb = Workbook()
    suggestions_ws = suggestions_wb.active
    suggestions_ws.title = 'HighlightedSuggestions'
    suggestions_ws.append(['Row_No', 'Input_Column', 'Current_Output', 'Current_Product_Code', 'Suggested_Output', 'Suggested_Product_Code', 'Status'])
    highlighted_count = 0
    changed_count = 0
    for row_num in range(2, source_ws.max_row + 1):
        if not any((source_ws.cell(row_num, col).fill.fill_type for col in range(1, source_ws.max_column + 1))):
            continue
        highlighted_count += 1
        input_name = source_ws.cell(row_num, 1).value or ''
        current_output = source_ws.cell(row_num, 2).value or ''
        current_code = source_ws.cell(row_num, 3).value or ''
        suggested_output, suggested_code = suggest_local_match(input_name, brand_map, verbose=False)
        status = 'Needs manual review'
        if str(current_output).strip().upper() != str(suggested_output).strip().upper():
            status = 'Suggested change'
            changed_count += 1
        suggestions_ws.append([row_num, input_name, current_output, current_code, suggested_output, suggested_code, status])
    suggestions_wb.save(suggestions_xlsx_path)
    print(f'[OK] Reviewed {highlighted_count} rows, suggested {changed_count} changes')
    return suggestions_xlsx_path

def _save_to_mapping_sheet(df, workbook_path):
    try:
        wb = load_workbook(workbook_path)
    except FileNotFoundError:
        wb = Workbook()
        for name in wb.sheetnames:
            if name in ('Sheet', 'Sheet1'):
                del wb[name]
    if 'mapping' in wb.sheetnames:
        del wb['mapping']
    ws = wb.create_sheet('mapping')
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row in enumerate(df.itertuples(index=False, name=None), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=value if not (isinstance(value, float) and value != value) else '')
    wb.save(workbook_path)

def format_row_result(position: int, total: int, input_name: str, res: dict) -> str:
    mapped = str(res.get('output', '') or '')
    code = str(res.get('product_code', '') or '')
    tail = f'{mapped} [{code}]' if code else mapped
    counter = f'[{position:>{len(str(total))}}/{total}]'
    return f"{counter} {input_name[:44]:<44} -> {tail}  ({res.get('status', '')})"

def process_excel_file(workbook_path=None, master_path=None):
    workbook_path = Path(workbook_path or INPUT_OUTPUT_XLSX_PATH)
    master = load_master_data(master_path)
    brand_map = build_brand_product_map(master)
    all_variants = extract_all_variants_from_data(brand_map)
    all_sub_variants = extract_all_sub_variants_from_data(brand_map)
    source = pd.read_excel(workbook_path, sheet_name='garbage_check', engine='openpyxl', dtype=object)
    if len(source.columns) < 3:
        raise ValueError('The garbage_check sheet needs product, remark, and matched-brand columns.')
    product_column, garbage_column, brand_column = source.columns[:3]
    remarks = source[garbage_column].astype(str).str.strip()
    confirmed = source.loc[remarks.eq('0'), [product_column, brand_column]].copy()
    confirmed.rename(columns={product_column: INPUT_COLUMN, brand_column: '_brand_hint'}, inplace=True)
    confirmed['_auto_detect'] = False
    confirmed['_source_label'] = '0'
    maybe = source.loc[remarks.eq('MAYBE_PRODUCT'), [product_column]].copy()
    maybe.rename(columns={product_column: INPUT_COLUMN}, inplace=True)
    maybe['_brand_hint'] = None
    maybe['_auto_detect'] = True
    maybe['_source_label'] = 'MAYBE_PRODUCT'
    if PROCESS_CONFIRMED_ZERO_ONLY and PROCESS_MAYBE_PRODUCT_ONLY:
        raise ValueError('Confirmed-only and maybe-only processing are mutually exclusive.')
    if PROCESS_CONFIRMED_ZERO_ONLY:
        rows = confirmed
    elif PROCESS_MAYBE_PRODUCT_ONLY:
        rows = maybe
    else:
        rows = pd.concat([confirmed, maybe], ignore_index=True)
    rows = rows.reset_index(drop=True)
    if ROW_LIMIT is not None:
        rows = rows.iloc[:ROW_LIMIT].copy()
    output_columns = [OUTPUT_COLUMN, PRODUCT_CODE_COLUMN, STATUS_COLUMN, CONFIDENCE_COLUMN, CANDIDATES_COLUMN, SUGGESTIONS_COLUMN, SOURCE_COLUMN]
    for column in output_columns:
        rows[column] = pd.Series('', index=rows.index, dtype=object)
    print(f'Processing {len(rows)} product rows locally...')
    for idx, row in rows.iterrows():
        value = row[INPUT_COLUMN]
        input_name = '' if pd.isna(value) else str(value).strip()
        hint_value = row.get('_brand_hint')
        hint = None if pd.isna(hint_value) else str(hint_value).strip() or None
        if not input_name:
            result = make_result('SKIPPED_EMPTY_INPUT', status='SKIPPED_EMPTY', confidence='NONE')
        else:
            try:
                result = process_product(input_name, brand_map, all_variants, all_sub_variants, forced_brand=hint, auto_detect=bool(row['_auto_detect']))
            except Exception as exc:
                result = make_result(f'ERROR: {str(exc)[:100]}', status='ERROR', confidence='NONE')
        for column, key in ((OUTPUT_COLUMN, 'output'), (PRODUCT_CODE_COLUMN, 'product_code'), (STATUS_COLUMN, 'status'), (CONFIDENCE_COLUMN, 'confidence'), (CANDIDATES_COLUMN, 'candidate_count')):
            rows.at[idx, column] = result[key]
        rows.at[idx, SUGGESTIONS_COLUMN] = format_suggestions(result['suggestions'])
        rows.at[idx, SOURCE_COLUMN] = row['_source_label']
        print(format_row_result(idx + 1, len(rows), input_name, result))
        if (idx + 1) % 50 == 0:
            _save_to_mapping_sheet(rows.drop(columns=[c for c in rows if c.startswith('_')]), workbook_path)
    results = rows.drop(columns=[c for c in rows if c.startswith('_')])
    _save_to_mapping_sheet(results, workbook_path)
    print(f'Saved {len(results)} rows to {workbook_path} -> mapping')
    return results

def interactive_mode(master_path=None):
    master = load_master_data(master_path)
    brand_map = build_brand_product_map(master)
    print(f'Zuventus local product mapping: {len(brand_map)} brands')
    print('Commands: exit, brands, stats')
    while True:
        input_name = input('Enter INPUT PRODUCT NAME: ').strip()
        if input_name.lower() in {'exit', 'quit', 'q'}:
            break
        if not input_name:
            continue
        if input_name.lower() == 'brands':
            for brand, items in sorted(brand_map.items()):
                print(f'{brand}: {len(items)} products')
            continue
        if input_name.lower() == 'stats':
            print(f'Brands: {len(brand_map)}, products: {sum(map(len, brand_map.values()))}')
            continue
        result = process_product(input_name, brand_map, verbose=True, auto_detect=True)
        print(f"Result: {result['output']} [{result['product_code']}]")
        print(f"Status: {result['status']}; confidence: {result['confidence']}")
        if result['suggestions']:
            print('Suggestions: ' + format_suggestions(result['suggestions']))
_EXACT_CACHE_MAP = None
_EXACT_CACHE_INDEX = None

def _get_exact_catalog_index(brand_map: dict) -> dict:
    global _EXACT_CACHE_MAP, _EXACT_CACHE_INDEX
    if _EXACT_CACHE_MAP is brand_map:
        return _EXACT_CACHE_INDEX
    index = {}
    seen = set()
    for items in brand_map.values():
        for item in items:
            identity = (item['product'], item.get('product_code', ''), item.get('pack_size', ''))
            if identity not in seen:
                seen.add(identity)
                index.setdefault(norm(item['product']), []).append(item)
    _EXACT_CACHE_MAP, _EXACT_CACHE_INDEX = (brand_map, index)
    return index

def load_master_data(path=None) -> pd.DataFrame:
    return load_master(path or MASTER_XLSX_PATH)
_FORM_NAME_EQUIVALENTS = {'TABLET': 'TABLET', 'TABLETS': 'TABLET', 'TAB': 'TABLET', 'TABS': 'TABLET', 'TA': 'TABLET', 'CAPSULE': 'CAPSULE', 'CAPSULES': 'CAPSULE', 'CAP': 'CAPSULE', 'CAPS': 'CAPSULE', 'INJECTION': 'INJECTION', 'INJECTIONS': 'INJECTION', 'INJ': 'INJECTION', 'INJN': 'INJECTION', 'SYRUP': 'SYRUP', 'SYRUPS': 'SYRUP', 'SYP': 'SYRUP', 'SYR': 'SYRUP', 'SUSPENSION': 'SUSPENSION', 'SUSPENSIONS': 'SUSPENSION', 'SUSP': 'SUSPENSION', 'SUSPN': 'SUSPENSION', 'DROP': 'DROPS', 'DROPS': 'DROPS'}
_FORM_CACHE_MAP = None
_FORM_CACHE_INDEX = None

def _canonical_form_name(value: str) -> str:
    return ' '.join((_FORM_NAME_EQUIVALENTS.get(token, token) for token in norm(value).split()))

def _get_form_catalog_index(brand_map: dict) -> dict:
    global _FORM_CACHE_MAP, _FORM_CACHE_INDEX
    if _FORM_CACHE_MAP is brand_map:
        return _FORM_CACHE_INDEX
    index = {}
    for items in _get_exact_catalog_index(brand_map).values():
        for item in items:
            index.setdefault(_canonical_form_name(item['product']), []).append(item)
    _FORM_CACHE_MAP, _FORM_CACHE_INDEX = (brand_map, index)
    return index
if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(errors='backslashreplace')
    import argparse
    parser = argparse.ArgumentParser(description='Zuventus local product mapping')
    parser.add_argument('--master', type=Path, default=MASTER_XLSX_PATH)
    parser.add_argument('--workbook', type=Path, default=INPUT_OUTPUT_XLSX_PATH)
    parser.add_argument('--product', help='Map one product locally')
    parser.add_argument('--interactive', action='store_true')
    args = parser.parse_args()
    if args.product is not None:
        catalog = build_brand_product_map(load_master_data(args.master))
        result = process_product(args.product, catalog, auto_detect=True)
        for key in ('output', 'product_code', 'status', 'confidence'):
            print(f'{key}: {result[key]}')
    elif args.interactive:
        interactive_mode(args.master)
    else:
        process_excel_file(args.workbook, args.master)
