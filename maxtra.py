"""Bounded MAXTRA descriptor aliases shared by recognition and mapping.

This module normalizes text only. Catalog rows and the existing mapper still
own identity validation, pack selection, ambiguity checks and result creation.
"""
import re
from functools import lru_cache
from pack_selection import normalize_ml_punctuation


_GARGLE_WORDS = r'GARGLES|GARGAL|GARGEL|GARGIL|GARGLE'
_ORAL_WORDS = r'ORALDROPS|ORALDROP|ORALDR'
_NASAL_WORDS = r'NASALSPRAY|NASALSPARY|NASALSPREY|NASALSPRY|NASALSPR'
_FUSED_BRAND = re.compile(
    r'(?<![A-Z0-9])MAXTRA(?=(?:' + _GARGLE_WORDS + '|' + _ORAL_WORDS
    + '|' + _NASAL_WORDS + r'|DROPS?|TABLETS?|TABS?|SYRUP|SYP|COLD(?:\+|PLUS))(?![A-Z]))', re.I)
_BRAND = re.compile(r'(?<![A-Z0-9])MAXTRA(?![A-Z0-9])', re.I)
_SPACE = re.compile(r'\s+')
_TYPED_TABLETS = re.compile(r'(?<![A-Z0-9.])\d+\s*T(?![A-Z0-9])', re.I)
_BARE_P_TEN = re.compile(r'^\s*MAXTRA[\s-]+P[\s-]*10\s*\.?\s*$', re.I)
_SHORT_NASAL_SPRAY = re.compile(r'(?<![A-Z0-9])([OS])[\s-]+N\s*[/\.]\s*S(?![A-Z0-9])', re.I)
_GARGLE_PRESENT = re.compile(r'(?<![A-Z0-9])GARGLE(?![A-Z])', re.I)
_MOUTHWASH = re.compile(r'(?<![A-Z0-9])M\s*/\s*W(?![A-Z0-9])', re.I)
_PIECE_COUNT = re.compile(r'(?<![A-Z0-9.])([1-9]\d*(?:\s*[X*]\s*[1-9]\d*)*)\s*PIC(?![A-Z0-9])', re.I)
_SOURCE_CODE = re.compile(r'(ML)\s*(?:ZUV076|ZUV156|ZH028|ZH069|EZ0200|EZ0500|EZ0121|ZU0097|ZU0046|MAXSND|MAXD|ZUV)(?![A-Z0-9])|(?<=TAB)ZUV\b', re.I)
_METADATA = re.compile(r'(?<![A-Z])(?:ZUVENTUS|ZUV|IPCS|PHY|PH|PER|VVL|JMS0)(?![A-Z0-9])', re.I)
_PRICE = re.compile(r'\bMRP\s*[-:]\s*\d+(?:\.\d{1,2})?\b', re.I)
_PARENTHETICAL = re.compile(r'\(\s*([A-Z]{1,6})\s*\)', re.I)
_SPRAY_PERIOD = re.compile(r'\b(SPR|SPRY|SPRAY)\.(?=\d+\s*ML\b)', re.I)
_PACKAGING = re.compile(r'\bBOTTLES?\s+OF\s+(?=\d+(?:\.\d+)?\s*ML\b)|\bSTRIP\s+OF\s+(?=\d+\s*TABLETS?\b)', re.I)
_GLUED_VOLUME_DESCRIPTOR = re.compile(r'(\d+(?:\.\d+)?)\s*ML(?=(GARGLE|N[./]SPRAY)(?![A-Z]))', re.I)
_GARGLE_INGREDIENT = re.compile(r'\bGARGLE\s*BENZYDAMINE(?![A-Z])', re.I)
_SALINE_ANNOTATION = re.compile(r'(?<![\d.])0\.65\s*%?\s*W\s*/\s*V(?![A-Z])', re.I)
_S_NASAL_SPRAY = re.compile(r'\bS[\s-]+NASAL[\s-]+SPRAY(?![A-Z])', re.I)
_SHORT_ORAL = re.compile(r'\bO\s*[/\.]\s*(?:DROPS?|DR|D)\b', re.I)
_SHORT_NASAL = re.compile(r'\bN\s*[/\.]\s*(?:DROPS?|D)\b', re.I)
_M_VOLUME = re.compile(r'(?<![A-Z0-9.])((?:\d+\s*[X*]\s*)?\d+(?:\.\d+)?)\s*M(?:[.;]\s*L)?(?![A-Z])', re.I)
_BARE_VOLUME = re.compile(r'(?<![A-Z0-9.])(?:(\d+)\s*[X*]\s*)?(\d+)(?![A-Z0-9.])', re.I)
_FOLLOWING_UNIT = re.compile(r'\s*(?:MCG|MG|GM|ML|KG|G|L|IU|UNITS?|TABLETS?|TABS?|TA|CAPSULES?|CAPS?|PCS|PC|NOS|[STCP])\b', re.I)
_EXPLICIT_VOLUME = re.compile(r'(?<![A-Z0-9.])(?:(\d+)\s*[X*]\s*)?(\d+(?:\.\d+)?|\.\d+)\s*ML(?![A-Z])', re.I)
_CONTAINER_UNIT = re.compile(r'\b1\s+UNIT\b', re.I)
NUMERIC_OFFER = re.compile(r'\d+\s*\+\s*\d+')

# Finite source spellings, not a fuzzy vocabulary. Literal catalog variants win.
_ALIASES = {
    **dict.fromkeys(('GARGALE', 'GARGUL', 'GARRGAL', 'GARRGLE', 'GARGLR', 'GERGLE', 'GARGELS', 'GARGL', 'GRGLE', 'GURGLE'), 'GARGLE'),
    **dict.fromkeys(('ORL', 'OREL', 'ORLA', 'OROAL', 'ORALL', 'ORIAL'), 'ORAL'),
    **dict.fromkeys(('DRO', 'DP', 'DPS', 'DRP', 'DRAPS'), 'DROPS'),
    **dict.fromkeys(('NAS', 'NASL', 'NEASL', 'NASAAL', 'NOSEL', 'NASOL', 'NAAL'), 'NASAL'),
    **dict.fromkeys(('SYRUB', 'SYUP', 'SYRP'), 'SYRUP'),
    'SUS': 'SUSPENSION', 'SOLUSTION': 'SOLUTION', 'LIQ': 'LIQUID', 'TABL': 'TABLET',
    'CLOD': 'COLD', 'COLDTAB': 'COLD TABLET', 'PSYP': 'P SYRUP', 'PDS': 'P DS', 'DSSYP': 'DS SYRUP',
    'DSSYR': 'DS SYRUP', 'DSSYURP': 'DS SYRUP', 'HSDROP': 'HS DROPS',
    'CDROP': 'C DROPS', 'DCSYP': 'DC SYRUP', 'LMSYP': 'LM SYRUP',
    'ONASAL': 'O NASAL', 'SNASAL': 'S NASAL', 'NOSALSPRAY': 'NASAL SPRAY',
    'NASALSP': 'NASAL SPRAY',
    'DM': 'LM', 'IM': 'LM',
}


def split_supported_brand(text):
    """Separate only a complete catalog descriptor fused to the brand."""
    return _FUSED_BRAND.sub('MAXTRA ', str(text or ''))


@lru_cache(maxsize=32)
def _alias_patterns(variants):
    # Garbage-check evidence stores compact variants; the mapper stores words.
    compact = {re.sub('[^A-Z0-9]', '', value.upper()) for value in variants}
    patterns = []

    def add(expression, replacement):
        patterns.append((re.compile(r'(?<![A-Z0-9])(?:' + expression + r')(?![A-Z])', re.I), replacement))

    if 'GARGLE' in compact:
        add(_GARGLE_WORDS, 'GARGLE')
    if any('NASAL' in value for value in compact):
        add(_NASAL_WORDS, 'NASAL SPRAY')
        add(r'NARSAL|NASEL|NAZAL|NOSAL|NESAL', 'NASAL')
        # A complete expression, never an isolated N token.
        add(r'N\s*/\s*(?:SPRAY|SPARY|SPREY|SPRY|SPR)', 'NASAL SPRAY')
        add(r'NASAL[\s-]+(?:SPARY|SPREY|SPRY|SPR)', 'NASAL SPRAY')
    if any('ORAL' in value for value in compact):
        add(_ORAL_WORDS, 'ORAL DROPS')
    vocabulary = set(re.findall('[A-Z]+', ' '.join(variants))) | {'DROPS', 'SYRUP', 'SUSPENSION', 'SOLUTION', 'LIQUID', 'TABLET', 'SPRAY'}
    for alias, target in _ALIASES.items():
        if set(target.split()) <= vocabulary:
            add(re.escape(alias), target)
    if any('NASAL' in value for value in compact):
        add(r'NASAL[\s/-]+(?:SPARAY|SPARA|SPRA|SPAY|SYRAY|SSPRAY|SPREY|SPRY|SPY|SP)', 'NASAL SPRAY')
        add(r'N\s*[/.-]\s*(?:SPY|SSPRAY|SPRAY)', 'NASAL SPRAY')
    if any('HS' in value for value in compact):
        add(r'H\.S', 'HS')
    return tuple(patterns)


def normalize_description(text, variants):
    """Normalize approved whole descriptors only under a real MAXTRA brand."""
    text = split_supported_brand(text)
    if not _BRAND.search(text):
        return text
    literal = {value.upper() for value in variants}
    text = _GLUED_VOLUME_DESCRIPTOR.sub(
        lambda m: m[1] + ' ML ' if ('GARGLE' in literal if m[2].upper() == 'GARGLE'
                                     else any('NASAL' in value for value in literal)) else m[0], text)
    vocabulary = set(re.findall('[A-Z]+', ' '.join(literal))) | {'SYP', 'SYR', 'SYRUP', 'TAB', 'TABS', 'TABLET', 'DROPS', 'DROP'}
    if 'GARGLE' in literal and 'BENZYDAMINE' not in vocabulary:
        text = _GARGLE_INGREDIENT.sub('GARGLE', text)
    text = _PARENTHETICAL.sub(lambda m: ' ' + m[1] + ' ' if m[1].upper() in vocabulary else m[0], text)
    text = _SOURCE_CODE.sub(lambda m: (m[1] or '') + ' ', text)
    text = _METADATA.sub(lambda m: m[0] if m[0].upper() in vocabulary else ' ', text)
    text = _PRICE.sub(' ', text)
    text = _SPRAY_PERIOD.sub(r'\1 ', text)
    text = normalize_ml_punctuation(text)
    text = _PACKAGING.sub('', text)
    if any('ORAL' in value for value in literal):
        text = _SHORT_ORAL.sub('ORAL DROPS', text)
    if any('NASAL' in value for value in literal):
        text = _SHORT_NASAL.sub('NASAL DROPS', text)
    for pattern, replacement in _alias_patterns(tuple(sorted(literal))):
        text = pattern.sub(lambda match: match.group() if match.group().upper() in literal
                           or match.group().upper() in vocabulary else replacement, text)
    if any('NASAL' in value.replace(' ', '') for value in literal):
        text = _SHORT_NASAL_SPRAY.sub(r'\1 NASAL SPRAY', text)
        without_annotation = _SALINE_ANNOTATION.sub(' ', text)
        if _S_NASAL_SPRAY.search(without_annotation):
            text = without_annotation
    if 'GARGLE' in literal and _GARGLE_PRESENT.search(text) and not {'M/W', 'M W', 'MW'} & literal:
        text = _MOUTHWASH.sub('GARGLE', text)
        text = re.sub(r'\bMOUTH\s+WASH\b', 'GARGLE', text, flags=re.I)
    if 'PIC' not in literal:
        text = _PIECE_COUNT.sub(r'\1 PCS', text)
    return _SPACE.sub(' ', text).strip()


def normalize_liquid_quantity(text, available_volumes, *, nasal_spray=False):
    """Interpret approved source units only after compatible liquid identity."""
    text = _M_VOLUME.sub(r'\1 ML', text)

    def bare(match):
        if _FOLLOWING_UNIT.match(text, match.end()):
            return match[0]
        value = int(match[1] or 1) * int(match[2])
        return match[0] + ' ML' if value in available_volumes else match[0]

    text = _BARE_VOLUME.sub(bare, text)
    return _CONTAINER_UNIT.sub('1 PCS', text) if nasal_spray else text


def recover_explicit_volume(text, existing):
    """Recover count-shadowed ML without changing a valid physical measurement."""
    if existing and existing[1] != 'COUNT':
        return existing
    values = {float(m[2]) * int(m[1] or 1) for m in _EXPLICIT_VOLUME.finditer(text)} - {0}
    return (next(iter(values)), 'VOLUME') if len(values) == 1 else existing


def tablet_count_evidence(text):
    """Typed T counts, plus the specifically approved bare MAXTRA P 10."""
    return bool(_TYPED_TABLETS.search(text) or _BARE_P_TEN.fullmatch(text))


def fused_brand_evidence(text, variants, products=()):
    """Return bounded descriptor evidence for the confirmed-only app gate."""
    separated = split_supported_brand(text)
    if separated == text:
        return None
    normalized = normalize_description(separated, variants)
    compact = {re.sub('[^A-Z0-9]', '', value.upper()) for value in variants}
    if ('GARGLE' in compact and re.search(r'\bGARGLE(?=$|[^A-Z])', normalized, re.I)
            or any('ORAL' in value for value in compact)
            and re.search(r'\b(?:ORAL\s+)?DROPS?(?=$|[^A-Z])', normalized, re.I)):
        return normalized
    if (re.search(r'\b(?:SYRUP|SYP)(?=$|[^A-Z])', normalized, re.I)
            and any(not item['variant'] and 'SYP' in item['forms'] for item in products)):
        return normalized
    if (re.search(r'\bCOLD\s*(?:\+|PLUS\b)', normalized, re.I)
            and any(item['variant'] == 'COLD' and item['strength'] == 'PLUS' for item in products)):
        return normalized
    return None
