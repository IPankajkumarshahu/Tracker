"""Find Indian vehicle registration numbers in free text (e.g. email subjects)."""

import re

# RTO state / union-territory codes. Restricting the first two letters to real
# codes keeps random words like "RE 12 AB 1234" from being picked up.
STATE_CODES = {
    "AN", "AP", "AR", "AS", "BR", "CG", "CH", "DD", "DL", "DN", "GA", "GJ",
    "HP", "HR", "JH", "JK", "KA", "KL", "LA", "LD", "MH", "ML", "MN", "MP",
    "MZ", "NL", "OD", "OR", "PB", "PY", "RJ", "SK", "TG", "TN", "TR", "TS",
    "UK", "UA", "UP", "WB",
}

_SEP = r"[\s\-./]*"

# Standard format: MH12AB1234, MH 12 AB 1234, MH-12-A-123, DL-3C-AB-1234 ...
_STANDARD = re.compile(
    rf"(?<![A-Z0-9])([A-Z]{{2}}){_SEP}(\d{{1,2}}){_SEP}([A-Z](?:{_SEP}[A-Z]){{0,2}}){_SEP}(\d{{1,4}})(?![A-Z0-9])"
)

# Bharat series: 22BH1234AB, 22 BH 1234 A ...
_BHARAT = re.compile(
    rf"(?<![A-Z0-9])(\d{{2}}){_SEP}(BH){_SEP}(\d{{4}}){_SEP}([A-Z]{{1,2}})(?![A-Z0-9])"
)

# Labelled numbers ("Regn. No. HR890648", "Registration No: DL 1 1234") may
# lack series letters; accept those only when the label is present.
_LABELLED = re.compile(
    rf"(?:REGN?|REGISTRATION|VEH(?:ICLE)?)\.?\s*(?:NO|NUMBER)\.?\s*[:#\-]?\s*"
    rf"([A-Z]{{2}}){_SEP}(\d{{1,2}}){_SEP}()(\d{{1,4}})(?![A-Z0-9])"
)


def find_reg_numbers(text):
    """Return unique registration numbers found in ``text``, normalised to
    upper case without spaces/hyphens (e.g. ``MH12AB1234``), in order of
    appearance."""
    if not text:
        return []
    upper = text.upper()
    found = []

    for m in _BHARAT.finditer(upper):
        found.append((m.start(), "".join(m.groups())))  # already separator-free

    standard = [(m.start(), m.groups()) for m in _STANDARD.finditer(upper)]
    labelled = [(m.start(1), m.groups()) for m in _LABELLED.finditer(upper)]
    for start, (state, district, series, number) in standard + labelled:
        if state not in STATE_CODES:
            continue
        series = re.sub(r"[^A-Z]", "", series)
        found.append((start, f"{state}{district.zfill(2)}{series}{number.zfill(4)}"))

    result = []
    for _, reg in sorted(found):
        if reg not in result:
            result.append(reg)
    return result
