"""Explicitly confirmed source spellings; never broaden fuzzy brand matching."""
import re

_CORLION = re.compile(r'(?<![A-Z0-9])CORLION(?![A-Z0-9])', re.I)


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
