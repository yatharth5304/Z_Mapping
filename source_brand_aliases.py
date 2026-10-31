"""Explicitly confirmed source spellings; never broaden fuzzy brand matching."""
import re

_CORLION = re.compile(r'(?<![A-Z0-9])CORLION(?![A-Z0-9])', re.I)
_HU_ENZYME_JOIN = re.compile(
    r'^\s*HU[\s-]*ENZYME\s*(?:\+|PLUS)\s*PROBIOTIC(?![A-Z0-9])', re.I)


def normalize_hu_enzyme_source(text, brands):
    """Treat the approved join spellings identically, retaining the full brand."""
    if 'ENZYME' not in text.upper() or not _HU_ENZYME_JOIN.match(text):
        return text
    targets = [brand for brand in brands
               if re.sub('[^A-Z0-9]', '', brand.upper())
               in {'HUENZYMEPROBIOTIC', 'HUENZYMEPLUSPROBIOTIC'}]
    if (len(targets) != 1
            or re.sub('[^A-Z0-9]', '', targets[0].upper()) != 'HUENZYMEPROBIOTIC'):
        return text
    return _HU_ENZYME_JOIN.sub('HU ENZYME PROBIOTIC', text, count=1)


def resolve_source_brand_alias(text, brands):
    if 'CORLION' not in text.upper() or not _CORLION.search(text):
        return None
    brands = list(brands)
    if any(re.sub('[^A-Z0-9]', '', brand.upper()) == 'CORLION' for brand in brands):
        return None
    targets = [brand for brand in brands if re.sub('[^A-Z0-9]', '', brand.upper()) == 'CORALIUM']
    return targets[0] if len(targets) == 1 else None


def replace_source_brand_alias(text, brand):
    return _CORLION.sub(lambda match: brand, text)
