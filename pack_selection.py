"""Unit-aware selection from already compatible, loaded master rows."""
import math
import re
from functools import lru_cache

_UNITS = {
    'ML': ('VOLUME', 1), 'L': ('VOLUME', 1000),
    'MCG': ('MASS', .001), 'MG': ('MASS', 1),
    'G': ('MASS', 1000), 'GM': ('MASS', 1000), 'KG': ('MASS', 1000000),
    'IU': ('ACTIVITY', 1), 'UNIT': ('ACTIVITY', 1),
    **{unit: ('COUNT_ORAL', 1) for unit in
       ('TA', 'TAB', 'TABS', 'TABLET', 'TABLETS', 'CAP', 'CAPS', 'CAPSULE', 'CAPSULES')},
    **{unit: ('COUNT_INJECTION', 1) for unit in ('VIAL', 'VIALS', 'AMP', 'AMPS', 'AMPOULE', 'AMPOULES')},
    **{unit: ('COUNT_PATCH', 1) for unit in ('PATCH', 'PATCHES')},
    **{unit: ('COUNT_SACHET', 1) for unit in ('SACHET', 'SACHETS')},
    **{unit: ('COUNT_STRIP', 1) for unit in ('STRIP', 'STRIPS')},
    **{unit: ('COUNT', 1) for unit in ('S', 'T', 'C', 'P', 'PCS', 'PC', 'NOS', 'COUNT')},
}
UNIT_PATTERN = '|'.join(sorted(_UNITS, key=len, reverse=True))
_PACK = re.compile(r'\s*(?:(\d+)\s*[X*]\s*)?(\d+(?:\.\d+)?|\.\d+)\s*(' + UNIT_PATTERN + r')\s*', re.I)
_ENDING_UNIT = re.compile(r'(' + UNIT_PATTERN + r')\s*$', re.I)
_FOLLOWING_UNIT = re.compile(r'\s*(' + UNIT_PATTERN + r')(?![A-Z])', re.I)
_ML_PUNCTUATION = re.compile(r'(?<![\d.])(\d+)\s*[.;*\-]\s*ML\b', re.I)
_SOURCE_VOLUME_UNIT = re.compile(
    r'(?<![A-Z0-9.])((?:\d+\s*[X*×]\s*)?(?:\d+(?:\.\d+)?|\.\d+))\s*(?:MIL|NL|M)(?![A-Z0-9])', re.I)
_SOURCE_BARE_VOLUME = re.compile(
    r'(?<![A-Z0-9.*×])(?:(\d+)\s*[X*×]\s*)?(\d+)(?![A-Z0-9.])', re.I)


def normalize_ml_punctuation(text, approved_volumes=None):
    """Repair approved number/unit separators; callers establish liquid context."""
    return _ML_PUNCTUATION.sub(lambda match: match[1] + ' ML'
                              if approved_volumes is None or int(match[1]) in approved_volumes else match[0], text)


def normalize_source_liquid_units(text, bare_volumes=(), *, single_container=False, approved_volumes=None):
    """Bounded source notation, never an unscoped quantity-parser fallback."""
    text = _SOURCE_VOLUME_UNIT.sub(
        lambda match: match[1] + ' ML' if approved_volumes is None
        or float(re.split(r'[X*×]', match[1])[-1]) in approved_volumes else match[0],
        normalize_ml_punctuation(text, approved_volumes))

    def bare(match):
        if _FOLLOWING_UNIT.match(text, match.end()):
            return match[0]
        if match[1] is not None and not (single_container and match[1] == '1'):
            return match[0]
        return match[0] + ' ML' if int(match[2]) in bare_volumes else match[0]

    return _SOURCE_BARE_VOLUME.sub(bare, text)


def measurement(value, unit):
    """Return quantity in a common unit and its measurement dimension."""
    spec = _UNITS.get(str(unit).upper())
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not spec or not math.isfinite(number) or number <= 0:
        return None
    dimension, factor = spec
    return number * factor, dimension


@lru_cache(maxsize=512)
def parse_master_pack(pack):
    if re.fullmatch(r'\s*\d+(?:\.\d+)?\s*', str(pack or '')):
        return measurement(pack, 'COUNT')
    match = _PACK.fullmatch(str(pack or ''))
    if not match:
        return None
    return measurement(float(match[2]) * int(match[1] or 1), match[3])


def matched_measurement(value, match, source, default_unit='COUNT'):
    """Attach units to the existing mapper's selected pack expression."""
    following = _FOLLOWING_UNIT.match(source, match.end())
    ending = _ENDING_UNIT.search(match.group())
    unit = ending[1] if ending else following[1] if following else default_unit
    return measurement(value, unit)


def select_pack_rows(numbered, blank, requested):
    """Exact, lower, unique blank base, then lowest; retain every numbered tie."""
    if not numbered:
        return (list(blank), '') if len(blank) == 1 else ([], 'AMBIGUOUS_PACK_TARGET' if blank else 'NO_COMPATIBLE_PACK')
    value, dimension = requested
    if dimension == 'COUNT':
        kinds = {pack[1] for _, pack in numbered if pack[1].startswith('COUNT')}
        if len(kinds) > 1:
            kinds.discard('COUNT')
        if len(kinds) != 1:
            return [], 'AMBIGUOUS_PACK_UNIT' if kinds else 'NO_COMPATIBLE_PACK'
        dimension = next(iter(kinds))
    compatible = [(item, pack[0]) for item, pack in numbered
                  if pack[1] == dimension or dimension.startswith('COUNT_') and pack[1] == 'COUNT']
    if not compatible:
        return [], 'NO_COMPATIBLE_PACK'
    exact = [item for item, size in compatible if math.isclose(size, value, rel_tol=1e-9, abs_tol=1e-9)]
    if exact:
        return exact, ''
    lower = [size for _, size in compatible if size < value]
    if not lower and blank:
        return (list(blank), '') if len(blank) == 1 else ([], 'AMBIGUOUS_PACK_TARGET')
    selected = max(lower) if lower else min(size for _, size in compatible)
    return [item for item, size in compatible if size == selected], ''
