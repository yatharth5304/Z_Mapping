"""Catalog-bounded SOVENTUS source normalization and identity eligibility.

This module never selects a pack or manufactures a product. The mapper owns
measurement parsing, pack fallback, ranking and public result dictionaries.
"""
import re
from functools import lru_cache

from pack_selection import normalize_source_liquid_units


_SPACE = re.compile(r'\s+')
_JUNIOR_WORDS = (
    'JUNIOR', 'JNR', 'JUN', 'JN', 'JU', 'JUR', 'J', 'JUNOR', 'JUNIAR',
    'JUNIOUR', 'JUNIO', 'JUNIR', 'JUNR', 'JUNER', 'JUNNER', 'JUNIUR',
    'JUNOIR', 'JUNIER', 'JOUNIOR', 'JUNI',
)
_JUNIOR_EXPR = '|'.join(sorted(_JUNIOR_WORDS + ('JR',), key=len, reverse=True))
_JUNIOR = re.compile(r'(?<![A-Z])(?:' + _JUNIOR_EXPR + r')(?![A-Z])', re.I)
_FUSED_FORM = re.compile(
    r'(?<![A-Z])(?:LS(?:' + _JUNIOR_EXPR + r')|' + _JUNIOR_EXPR
    + r'|LS|DX|DC|TRIO)(?=(?:SYRUP|SYP|SYR|SUSPENSION|SUSP)(?![A-Z]))', re.I)
_LS_JUNIOR = re.compile(r'(?<![A-Z])LS(?=(?:' + _JUNIOR_EXPR + r')(?![A-Z]))', re.I)
_ORAL_FORM = re.compile(r'\bORAL(?=SUSP(?:ENSION)?(?![A-Z]))', re.I)
_TWELVE_ORAL = re.compile(r'(?<![\d.])12(?=ORAL\b)', re.I)
_TWELVE = re.compile(
    r'(?<![\d.])12(?![ST]\b)(?:\s*MG\b|(?!\s*(?:ML|MCG|MG|GM|G|KG|L|IU|PCS|PC|TAB)\b))(?![\d.])', re.I)
_FORM_PERIOD = re.compile(r'\b(SYP|SYRUP|SYR)\.(?=100\s*ML\b)', re.I)
_CHILD_PERIOD = re.compile(r'\b(SYP|SYRUP|SYR)\.(?=12(?:\s|$))', re.I)
_SPLIT_HUNDRED = re.compile(r'(?<![\d.])10\s+0\s*ML\b', re.I)
_SOURCE_CODE = re.compile(r'(?<![\d.])(100\s*ML)(?:EZ0416|EZ0076|NOS)(?![A-Z0-9])', re.I)
_VOLUME_BOTTLE = re.compile(r'(ML)(?=BOTTLES?\b)', re.I)
_SOURCE_UNIT = re.compile(r'(?<![\d.])(?:(100)\s*IS|(15|60|100)\s*MI|(15)\s*NL)(?![A-Z0-9])', re.I)
_UNAPPROVED_SOURCE_UNIT = re.compile(r'\d\s*(?:MIL|NL)\b', re.I)
_SUGAR_FREE = re.compile(r'\bS\s*/\s*F\b|\bSF\b|\bSUGAR[ -]*FREE\b', re.I)
_FUSED_SF = re.compile(r'\b(DC|DX|LS)SF\b', re.I)
_SHORT_LIQUID = re.compile(r'\b(?:S|SU)\b', re.I)
_ORAL_ALIAS = re.compile(r'\b(?:O\s*/\s*S|OREL|OR|O)\b', re.I)
_UNSUPPORTED_IDENTITY = re.compile(r'\b(?:HS|DS|SC|LD|EXP|EXPT)\b', re.I)
_NON_LIQUID = re.compile(
    r'\b(?:TABLETS?|TABS?|CAPSULES?|CAPS?|INJECTION|INJ|CREAM|OINTMENT|SACHETS?|SPRAY)\b', re.I)
_FORM_ALIASES = {
    **dict.fromkeys(('SYURP', 'SYRP', 'SYRU', 'SYS', 'SYRUB', 'SRUP', 'SRP', 'SYP', 'SYR',
                     'SY', 'STP', 'SYPS', 'SYPTML'), 'SYRUP'),
    **dict.fromkeys(('SUS', 'SUP', 'SUSP', 'SUSPEN', 'SUSPENSIO',
                     'SUSPE', 'SUSPENISION', 'SUSUP'), 'SUSPENSION'),
    **dict.fromkeys(('DRP', 'DR', 'DROP'), 'DROPS'),
    'TRION': 'TRIO', 'SYPZUVENTUS': 'SYRUP', 'SYPKS': 'SYRUP', 'LIQ': 'LIQUID',
}
_FORM_ALIAS = re.compile(r'\b(?:' + '|'.join(sorted(_FORM_ALIASES, key=len, reverse=True)) + r')(?![A-Z])', re.I)
_FORM = re.compile(r'\b(SYRUP|SYP|SUSPENSION|SUSP|DROPS?|LIQUID|SOLUTION|TONIC)\b', re.I)
_KNOWN = frozenset(('SOVENTUS', 'PLAIN', 'ORAL', 'SYRUP', 'SUSPENSION', 'DROPS',
                    'LIQUID', 'SOLUTION', 'TONIC', 'ML', 'L', 'MG', 'MCG', 'GM',
                    'G', 'KG', 'IU', 'UNIT', 'UNITS', 'M', 'PCS', 'PC', 'NOS',
                    'X', 'S', 'T', 'C', 'BOTTLE', 'BOTTLES', 'PACK', 'OF'))


def _forms(product):
    aliases = {'SYP': 'SYRUP', 'SUSP': 'SUSPENSION', 'DROP': 'DROPS'}
    return frozenset(aliases.get(m[1].upper(), m[1].upper()) for m in _FORM.finditer(product))


def _identity(value):
    return frozenset(re.findall(r'[A-Z]+', _SUGAR_FREE.sub(' ', value.upper()))) - {'PLAIN'}


@lru_cache(maxsize=32)
def _catalog(signature):
    """Cache only immutable catalog attributes; edits produce a different key."""
    vocabulary = frozenset(word for row in signature for value in row[1:3]
                           for word in re.findall(r'[A-Z]+', value))
    return vocabulary, tuple((_identity(row[1]), row[2], _forms(row[0]),
                               bool(_SUGAR_FREE.search(' '.join(row[:3])))) for row in signature)


def preserve_tags(text):
    """Keep approved identity tags before generic parenthetical cleanup."""
    text = re.sub(r'\(\s*(LS|JR|DX|DC|TRIO|SF|S/F|SUGAR\s+FREE)\s*\)', r' \1 ', text, flags=re.I)
    # These source annotations were explicitly left unresolved, not disposable.
    text = re.sub(r'\(\s*(BOT|BTL|PH|PHYL|PHY|PHL|B|GEN|INC|NA|EACH|SOVS)\s*\)', r' \1 ', text, flags=re.I)
    # Generic letter/digit splitting must not destroy these approved codes.
    return _SOURCE_CODE.sub(r'\1 ', text)


def prepare_candidates(text, items):
    """Return normalized text and compatible actual rows, or a refusal reason."""
    signature = tuple((item['product'], item['variant'], str(item['sub_variant'])) for item in items)
    vocabulary, attributes = _catalog(signature)
    text = preserve_tags(text).upper()

    def split_supported(match):
        word = re.match(r'[A-Z]+', text[match.start():])[0]
        return match[0] if word in vocabulary else match[0] + ' '

    text = _FUSED_FORM.sub(split_supported, text)
    text = _LS_JUNIOR.sub(split_supported, text)
    text = _ORAL_FORM.sub(split_supported, text)
    text = _TWELVE_ORAL.sub('12 ', text)
    text = _VOLUME_BOTTLE.sub(r'\1 ', text)
    text = re.sub(r'\bD C\b', 'DC', text)
    text = re.sub(r'\bD X\s*12\b', 'DX 12', text)
    text = re.sub(r'\bORL\s+SUSPEN\b', 'ORAL SUSPENSION', text)
    text = re.sub(r'\bDRO\s+15M\b', 'DROPS 15ML', text)
    text = re.sub(r'\bDROPSIDROP\b', 'DROPS', text)
    if 'JR' in vocabulary:
        text = _JUNIOR.sub(lambda m: m[0] if m[0] in vocabulary else 'JR', text)
    text = _FORM_ALIAS.sub(lambda m: m[0] if m[0] in vocabulary else _FORM_ALIASES[m[0]], text)
    text = _FUSED_SF.sub(lambda m: m[0] if m[0] in vocabulary else m[1] + ' SF', text)
    if _UNSUPPORTED_IDENTITY.search(text):
        return text, [], 'VARIANT_NOT_IN_MASTER'
    sugar_free = bool(_SUGAR_FREE.search(text))
    identity_text = _SUGAR_FREE.sub(' ', text)
    printed = frozenset(re.findall(r'[A-Z]+', identity_text))
    identity = printed & vocabulary - {'PLAIN', 'SF'}
    child_text = _CHILD_PERIOD.sub(r'\1 ', identity_text)
    child = bool('DX' in identity and _TWELVE.search(child_text))
    if child:
        if identity - {'DX', 'ORAL'}:
            return text, [], 'VARIANT_NOT_IN_MASTER'
        identity = frozenset(('DX', 'ORAL'))
        # Only this approved strength-tagged identity loses its printed MG unit.
        text = _TWELVE.sub('12', _CHILD_PERIOD.sub(r'\1 ', text))
        if not re.search(r'\bORAL\b', text):
            text = re.sub(r'\bDX\b', 'DX ORAL', text, count=1)
    selected = [(item, attr) for item, attr in zip(items, attributes)
                if attr[0] == identity and (not child or attr[1] == '12')]
    selected = [(item, attr) for item, attr in selected
                if attr[2] and not _NON_LIQUID.search(item['product'])]
    if not selected:
        return text, [], 'MISSING_SOVENTUS_TARGET'
    if identity == {'DX', 'ORAL'} and all(attr[1] == '12' for _, attr in selected):
        text = _ORAL_ALIAS.sub('ORAL', text)
    # Do not silently select a new child identity when the catalog is ambiguous.
    subs = {attr[1] for _, attr in selected}
    if not child:
        if 'PLAIN' in subs:
            selected = [(item, attr) for item, attr in selected if attr[1] == 'PLAIN']
        elif len(subs) != 1:
            return text, [], 'AMBIGUOUS_SOVENTUS_TARGET'
    sf_items = [(item, attr) for item, attr in selected if attr[3]]
    if sugar_free and sf_items:
        selected = sf_items
        # Keep the catalog distinction, but canonicalize slash/word spellings.
        sf_literal = any(re.search(r'\bSF\b', ' '.join((item['product'], item['variant'], str(item['sub_variant']))))
                         for item, _ in selected)
        text = _SUGAR_FREE.sub('SF' if sf_literal else 'SUGAR FREE', text)
    elif sugar_free:
        text = _SUGAR_FREE.sub(' ', text)
    else:
        selected = [(item, attr) for item, attr in selected if not attr[3]]
    if not selected:
        return text, [], 'MISSING_SOVENTUS_TARGET'
    forms = _forms(text)
    if _NON_LIQUID.search(text):
        return text, [], 'VARIANT_NOT_IN_MASTER'
    if identity == {'LS'}:
        syrup = bool(forms & {'SYRUP', 'SUSPENSION'})
        if 'DROPS' in forms and syrup:
            return text, [], 'VARIANT_NOT_IN_MASTER'
        if 'DROPS' in forms or syrup:
            drops = 'DROPS' in forms
            selected = [(item, attr) for item, attr in selected if ('DROPS' in attr[2]) == drops]
    if _SHORT_LIQUID.search(text):
        available_forms = {form for _, attr in selected for form in attr[2]}
        if len(available_forms) != 1:
            return text, [], 'VARIANT_NOT_IN_MASTER'
        text = _SHORT_LIQUID.sub(next(iter(available_forms)), text)
    if not selected:
        return text, [], 'MISSING_SOVENTUS_TARGET'
    text = _FORM_PERIOD.sub(r'\1 ', text)
    text = _SPLIT_HUNDRED.sub('100 ML', text)
    text = _SOURCE_CODE.sub(r'\1 ', text)
    text = _SOURCE_UNIT.sub(lambda match: next(value for value in match.groups() if value) + ' ML', text)
    if not _UNAPPROVED_SOURCE_UNIT.search(text):
        text = normalize_source_liquid_units(text, (100,), single_container=True, approved_volumes=(100,))
    unknown = set(re.findall(r'[A-Z]+', text)) - _KNOWN - vocabulary
    if sugar_free and sf_items:
        unknown -= {'SF', 'SUGAR', 'FREE', 'S', 'F'}
    if unknown:
        return text, [], 'VARIANT_NOT_IN_MASTER'
    return _SPACE.sub(' ', text).strip(), [item for item, _ in selected], ''
