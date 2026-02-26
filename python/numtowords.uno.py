# NumToWords LibreOffice Calc Add-In
# Python UNO component - mirrors the pattern used by libnumbertext
# Provides: =NUMTOWORDS(number, formatStyle)
#   formatStyle: 0 = cardinal (default), 1 = ordinal, 2 = currency (USD)

import uno
import unohelper
from com.sun.star.lang import XServiceInfo, XLocalizable, Locale
from com.sun.star.sheet import XAddIn
# Import the custom UNO interface from NumToWords.rdb — this is what makes
# numToWords() visible to LibreOffice's UNO introspection (same technique
# libnumbertext uses with XNumberText).
from com.numbertext.converter import NumToWordsConverter

IMPLEMENTATION_NAME = "com.numbertext.converter.NumToWordsPy"
SERVICE_NAME = "com.sun.star.sheet.AddIn"


def _ones():
    return ["", "one", "two", "three", "four", "five", "six", "seven",
            "eight", "nine", "ten", "eleven", "twelve", "thirteen",
            "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]


def _tens():
    return ["", "", "twenty", "thirty", "forty", "fifty",
            "sixty", "seventy", "eighty", "ninety"]


def _below_thousand(n):
    ones = _ones()
    tens = _tens()
    if n == 0:
        return ""
    elif n < 20:
        return ones[n]
    elif n < 100:
        t = tens[n // 10]
        o = ones[n % 10]
        return t + ("-" + o if o else "")
    else:
        h = ones[n // 100] + " hundred"
        rest = n % 100
        if rest:
            return h + " and " + _below_thousand(rest)
        return h


def _cardinal(n):
    """Convert integer part to cardinal words."""
    if n == 0:
        return "zero"
    parts = []
    scales = [
        (10 ** 12, "trillion"),
        (10 ** 9,  "billion"),
        (10 ** 6,  "million"),
        (10 ** 3,  "thousand"),
    ]
    for scale, name in scales:
        if n >= scale:
            parts.append(_below_thousand(n // scale) + " " + name)
            n %= scale
    if n > 0:
        parts.append(_below_thousand(n))
    return " ".join(parts)


_ORDINAL_MAP = {
    "one": "first", "two": "second", "three": "third", "five": "fifth",
    "eight": "eighth", "nine": "ninth", "twelve": "twelfth",
}


def _to_ordinal(cardinal):
    words = cardinal.split()
    last = words[-1]
    # handle hyphenated tens: "twenty-one" -> "twenty-first"
    if "-" in last:
        prefix, unit = last.rsplit("-", 1)
        words[-1] = prefix + "-" + _ordinal_suffix(unit)
    else:
        words[-1] = _ordinal_suffix(last)
    return " ".join(words)


def _ordinal_suffix(word):
    if word in _ORDINAL_MAP:
        return _ORDINAL_MAP[word]
    if word.endswith("t"):
        return word + "h"
    if word.endswith("e"):
        return word[:-1] + "th"
    if word.endswith("y"):
        return word[:-1] + "ieth"
    return word + "th"


def convert(number, fmt):
    """Core conversion. fmt: 0=cardinal, 1=ordinal, 2=currency."""
    negative = number < 0
    number = abs(number)
    int_part = int(number)
    frac_cents = round((number - int_part) * 100)

    words = _cardinal(int_part)

    if fmt == 1:
        words = _to_ordinal(words)
    elif fmt == 2:
        words += " dollar" + ("" if int_part == 1 else "s")
        if frac_cents:
            cent_words = _cardinal(frac_cents)
            words += " and " + cent_words + " cent" + ("" if frac_cents == 1 else "s")
    else:
        if frac_cents:
            # spell out decimal digits individually
            cent_str = f"{frac_cents:02d}"
            digit_words = " ".join(_ones()[int(d)] for d in cent_str)
            words += " point " + digit_words

    return ("minus " if negative else "") + words


class NumToWords(unohelper.Base, NumToWordsConverter, XAddIn, XServiceInfo, XLocalizable):
    """
    LibreOffice Calc Add-In exposing =NUMTOWORDS(number [, formatStyle]).
    Registered via XAddIn so Calc can find it without a custom IDL interface.
    """

    def __init__(self, ctx):
        self.ctx = ctx
        self.locale = Locale("en", "US", "")

    # ── XLocalizable ─────────────────────────────────────────────────────────

    def setLocale(self, locale):
        self.locale = locale

    def getLocale(self):
        return self.locale

    # ── XAddIn ───────────────────────────────────────────────────────────────

    def getProgrammaticCategoryName(self, name):
        return "Add-In"

    def getDisplayCategoryName(self, name):
        return "Add-In"

    def getProgrammaticFuntionName(self, display_name):
        return display_name

    def getDisplayFunctionName(self, prog_name):
        return prog_name.upper()

    def getFunctionDescription(self, name):
        return "Converts a number to its English word representation"

    def getDisplayArgumentName(self, name, idx):
        return ["Number", "FormatStyle"][idx] if idx < 2 else ""

    def getArgumentDescription(self, name, idx):
        descs = [
            "The number to convert to words",
            "Optional: 0=cardinal (default), 1=ordinal, 2=currency (USD)",
        ]
        return descs[idx] if idx < 2 else ""

    # ── XServiceInfo ─────────────────────────────────────────────────────────

    def getImplementationName(self):
        return IMPLEMENTATION_NAME

    def supportsService(self, name):
        return name == SERVICE_NAME

    def getSupportedServiceNames(self):
        return (SERVICE_NAME,)

    # ── The actual Calc function ──────────────────────────────────────────────
    # LibreOffice auto-injects the calling cell's XPropertySet as the first
    # argument (same pattern as libnumbertext's numbertext(self, prop, num, loc)).
    # This is NOT listed in the CalcAddIns XCU — it's injected transparently.

    def numToWords(self, number, formatStyle=None):
        try:
            # formatStyle arrives as void any when omitted — treat as 0 (cardinal)
            fmt = 0
            if formatStyle is not None:
                try:
                    fmt = int(formatStyle)
                except (TypeError, ValueError):
                    fmt = 0
            return convert(float(number), fmt)
        except Exception as e:
            return "Error: " + str(e)


# ── UNO component factory boilerplate ────────────────────────────────────────

def createInstance(ctx):
    return NumToWords(ctx)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    createInstance,
    IMPLEMENTATION_NAME,
    (SERVICE_NAME,),
)
