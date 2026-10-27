"""Bounded PRO-DIOF descriptions and master-backed presentation selection.

The blank-pack liquid is the user-confirmed 5 ML presentation. These rules do
not alter the generic form vocabulary, decimal parser, or pack comparator.
"""

from dataclasses import dataclass
import re
from rapidfuzz.distance import Levenshtein
from pack_selection import normalize_source_liquid_units


_BRAND = re.compile(
    r'(?<![A-Z0-9])(?:PRO[\W_]*DIOF(?=$|[^A-Z0-9]|LIQUID(?![A-Z]))'
    r'|DIOF[\W_]+PRO(?![A-Z0-9])|PEO[\s-]+DIOF(?![A-Z0-9]))', re.IGNORECASE)
_LIQUID_WORDS = (r'LLIQUID|LIQUIDS?|LIQUIDE|LIQUIED|LIQUD|LIQUI|LIQ|LIG|'
                 r'SUSPENSIONS?|SUSPESION|SUSPEN|SUSP|SYRUPS?|SYPR|SYP|SYR|SUS|SUP|'
                 r'SOLUSENT|SOLUTIONS?|SOL')
_FORM = re.compile(r'(?<![A-Z])(?:' + _LIQUID_WORDS + r')(?=$|[^A-Z])', re.I)
_DIOF = re.compile(r'(?<![A-Z0-9])DIOF(?:0[1-5](?![A-Z0-9])|'
                   r'(?=$|[^A-Z0-9]|(?:DS)?(?:' + _LIQUID_WORDS + r')(?=$|[^A-Z])|DS(?![A-Z])))', re.I)
_DOTTED_DS = re.compile(r'(?<![A-Z0-9])D\s*\.\s*S\s*\.?', re.I)
_SUS_D = re.compile(r'(?<![A-Z0-9])SUS\s*\.\s*D(?=$|[^A-Z]|ZUVENTUS\b)', re.I)
_FUSED_DS = re.compile(r'(?<![A-Z0-9])DS(?=(?:' + _LIQUID_WORDS + r')(?=$|[^A-Z]))', re.I)
_SOURCE_SUFFIX = re.compile(r'(?<=ML)(?:ZUVENTUS|ZU0053|ZUV)(?![A-Z0-9])', re.I)
_METADATA = re.compile(r'(?<![A-Z0-9])(?:ZUVENTUS|ZU0053|ZUV|BIG|CD\s*-\s*3)(?![A-Z0-9])', re.I)
_DPCO = re.compile(r'(?:(?<![A-Z])|(?<=ML))DPCO\s+PRODUCT(?![A-Z])', re.I)
_BOTTLE_OF = re.compile(r'\bBOTTLES?\s+OF\s*(?=\d+(?:\.\d+)?\s*ML\b)', re.I)
_ORAL = re.compile(r'\bORAL\b', re.I)


def match_pro_brand(value):
    """Complete PRO component or an explicitly approved source-brand alias."""
    return _BRAND.search(str(value or ''))


def match_diof_brand(value):
    """Bounded ordinary-family evidence, including approved numeric/fused forms."""
    return _DIOF.search(str(value or ''))


def _normalize_liquid_body(body, *, pro=False):
    if not pro:
        body = _SUS_D.sub(' DS SUSPENSION ', body)
        body = _DOTTED_DS.sub(' DS ', body)
        body = _FUSED_DS.sub('DS ', body)
    has_form = bool(_FORM.search(body))
    liquid = pro or has_form or bool(re.search(r'\d\s*ML', body))
    if liquid:
        body = _SOURCE_SUFFIX.sub(' ', body)
        body = _METADATA.sub(' ', body)
        if not pro:
            body = _DPCO.sub(' ', body)
    body = _FORM.sub(' LIQUID ' if pro else ' SUSPENSION ', body)
    if has_form:
        body = _ORAL.sub(' ', body)
    if liquid:
        approved = (30,) if pro else (60,)
        body = normalize_source_liquid_units(body, approved, single_container=pro, approved_volumes=approved)
    body = _BOTTLE_OF.sub(' ', body)
    return re.sub(r'\s+', ' ', body).strip()


def normalize_diof_source(value, *, confirmed=False):
    """Normalize ordinary DIOF text only; SKU selection remains in the mapper."""
    text = str(value or '').upper()
    if match_pro_brand(text):
        return None
    brand = match_diof_brand(text)
    if brand is None and confirmed:
        leading = re.match(r'\s*([A-Z]+)(?=$|[^A-Z])', text)
        if leading and Levenshtein.distance(leading[1], 'DIOF') == 1:
            brand = leading
    if brand is None:
        return None
    body = text[:brand.start()] + ' ' + text[brand.end():]
    return 'DIOF ' + _normalize_liquid_body(body)
_PRICE = re.compile(r'\bMRP\s*[-:=]?\s*\d+(?:\.\d+)?\b', re.IGNORECASE)
_SYRUP_SUFFIX = re.compile(r'\bSYRUPZU\b', re.IGNORECASE)
_VOLUME_SUFFIX = re.compile(r'(?<=ML)PRODA\b', re.IGNORECASE)
_VOLUME = re.compile(
    r'(?<![\d.])(?:(?P<count>\d+)\s*[X*×]\s*)?'
    r'(?P<number>\d+(?:\.\d+)?|\.\d+)(?P<gap>\s*[.*]?\s*)'
    r'(?P<unit>ML|M)(?![A-Z])', re.IGNORECASE)
_COUNT = re.compile(r'(?<![A-Z0-9.])\d+\s*[X*×]\s*\d+(?![A-Z0-9.])', re.IGNORECASE)
_WORDS = re.compile(r'[A-Z]+')
_NAME = re.compile(r'^PRODIOFLIQUID(?:(\d+(?:\.\d+)?)ML)?$')
_PACK = re.compile(r'^(\d+(?:\.\d+)?)\s*ML$', re.IGNORECASE)
_COMPACT = re.compile(r'[^A-Z0-9.]')
_DESCRIPTORS = frozenset({
    'LIQ', 'LIQUID', 'LIQUIDS', 'LIG', 'LIQUIDE', 'LIQUIED', 'LIQUD', 'MINI',
    'SUP', 'SUS', 'SUSP', 'SUSPENSION', 'SUSPENSIONS', 'SYP', 'SYR', 'SYRUP',
    'SOL', 'SOLUTION', 'SOLUTIONS', 'RESP', 'RESPULE', 'RESPULES',
    'AMP', 'AMPS', 'AMPL', 'AMPLE', 'AMPULE', 'AMPULES', 'AMPOULE', 'AMPOULES',
    'S', 'T', 'C', 'P', 'PCS', 'PC', 'NOS', 'NO', 'PACK', 'OF',
})


@dataclass(frozen=True, slots=True)
class Description:
    brand_text: str
    volumes: tuple
    five_ml: bool
    error: str = ''


def parse_description(value):
    """Recognize a complete family span and retain only approved descriptors.

    None means no complete PRO-DIOF identity. Unknown text and malformed unit
    expressions are refusals, including text after a recognized form. The
    literal .5ML override is distinct from numeric 0.5ML or 1.5ML.
    """
    text = str(value or '').upper()
    brand = _BRAND.search(text)
    if brand is None:
        return None
    if 'DUCT' in text:
        return Description(brand.group(), (), False, 'PRO_DIOF_DUCT_GUARD')
    body = text[:brand.start()] + ' ' + text[brand.end():]
    body = _PRICE.sub(' ', body)
    body = _SYRUP_SUFFIX.sub('SYRUP', body)
    body = _VOLUME_SUFFIX.sub('', body)
    body = _normalize_liquid_body(body, pro=True)
    volumes, five_ml = [], False
    error = ''

    def quantity(match):
        nonlocal five_ml, error
        number, gap, unit = match['number'], match['gap'].strip(), match['unit']
        volume = float(number)
        if match['count'] is not None and int(match['count']) <= 0 or volume <= 0:
            error = 'INVALID_PRO_DIOF_QUANTITY'
        elif not error and (gap and not (number == '5' and unit == 'ML')
                            or unit == 'M' and not (number == '30' and match['gap'].isspace())):
            error = 'UNSUPPORTED_PRO_DIOF_QUANTITY'
        volumes.append(volume)
        if unit == 'ML' and (volume == 5 or number == '.5'):
            five_ml = True
        return ' '

    body = _VOLUME.sub(quantity, body)
    body = _COUNT.sub(' ', body)
    words = set(_WORDS.findall(body))
    # A bare volume label carries no quantity; retain the approved base default.
    # Numeric malformed expressions such as 30;ML still require refusal.
    if 'ML' in words and not any(character.isdigit() for character in body):
        words.remove('ML')
    if not error and words & {'ML', 'M', 'MG', 'MCG', 'GM', 'IU', 'KG'}:
        error = 'UNSUPPORTED_PRO_DIOF_QUANTITY'
    if not error and not words.issubset(_DESCRIPTORS):
        error = 'VARIANT_NOT_IN_MASTER'
    return Description(brand.group(), tuple(volumes), five_ml, error)


_CACHE_MAP = None
_CACHE_PRESENTATIONS = None


def _presentations(brand_map):
    """Cache only references to the immutable, already-loaded master records."""
    global _CACHE_MAP, _CACHE_PRESENTATIONS
    if _CACHE_MAP is brand_map:
        return _CACHE_PRESENTATIONS
    keys = [key for key in brand_map if _COMPACT.sub('', key.upper()) == 'PRODIOF']
    presentations = {}
    if len(keys) == 1:
        for item in brand_map[keys[0]]:
            name = _NAME.fullmatch(_COMPACT.sub('', str(item.get('product', '')).upper()))
            if name is None or item.get('variant') != 'PLAIN' or item.get('sub_variant') != 'PLAIN':
                continue
            volume = float(name[1]) if name[1] else 5.0
            pack = str(item.get('pack_size', '')).strip()
            pack_match = _PACK.fullmatch(pack) if pack else None
            valid = (bool(str(item.get('product_code', '')).strip())
                     and str(item.get('product_code', '')).strip() not in {'0', '1', 'None', 'nan'}
                     and (not pack or name[1] is not None and pack_match is not None
                          and float(pack_match[1]) == volume))
            presentations.setdefault(volume, []).append((item, valid))
    _CACHE_MAP, _CACHE_PRESENTATIONS = brand_map, presentations
    return presentations


def select_presentation(description, brand_map):
    """Return (master record, refusal status) using the approved precedence.

    Alternate volumes select an actually available presentation at or below
    the request, then the confirmed base when no lower presentation exists.
    """
    if description.error:
        return None, description.error
    presentations = _presentations(brand_map)
    if description.five_ml:
        requested = {5.0}
    elif 30.0 in description.volumes:
        requested = {30.0}
    elif description.volumes:
        requested = set()
        for volume in description.volumes:
            lower = [available for available in presentations if available <= volume]
            requested.add(max(lower) if lower else 5.0)
        if len(requested) != 1:
            return None, 'AMBIGUOUS_PRO_DIOF_QUANTITY'
    else:
        requested = {5.0}
    matches = presentations.get(requested.pop(), [])
    if len(matches) > 1:
        return None, 'AMBIGUOUS_PRO_DIOF_TARGET'
    if not matches or not matches[0][1]:
        return None, 'MISSING_PRO_DIOF_TARGET'
    return matches[0][0], ''
