#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LibreOffice Calc Add-in: Number to Words Converter
Implements NUMTOWORDS() function for LibreOffice Calc
"""

import uno
import unohelper
from com.sun.star.lang import XServiceName, XLocalizable
from com.sun.star.sheet import XAddIn


# Implementation of the number-to-words converter
class NumToWordsAddIn(unohelper.Base, XServiceName, XAddIn, XLocalizable):
    def __init__(self, ctx):
        self.ctx = ctx
        self.locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        self.locale.Language = "en"
        self.locale.Country = "US"

    # XServiceName implementation
    def getServiceName(self):
        return "com.sun.star.sheet.addin.NumToWords"

    # XLocalizable implementation
    def setLocale(self, locale):
        self.locale = locale

    def getLocale(self):
        return self.locale

    # XAddIn implementation
    def getProgrammaticFunctionName(self, aDisplayName):
        # Map display name to programmatic name
        if aDisplayName == "NUMTOWORDS":
            return "NUMTOWORDS"
        return ""

    def getDisplayFunctionName(self, aProgrammaticName):
        # Map programmatic name to display name
        if aProgrammaticName == "NUMTOWORDS":
            return "NUMTOWORDS"
        return ""

    def getFunctionDescription(self, aProgrammaticName):
        if aProgrammaticName == "NUMTOWORDS":
            return "Converts a number to its word representation"
        return ""

    def getDisplayArgumentName(self, aProgrammaticName, nArgument):
        if aProgrammaticName == "NUMTOWORDS":
            if nArgument == 0:
                return "Number"
            elif nArgument == 1:
                return "Format"
        return ""

    def getArgumentDescription(self, aProgrammaticName, nArgument):
        if aProgrammaticName == "NUMTOWORDS":
            if nArgument == 0:
                return "The number to convert to words"
            elif nArgument == 1:
                return "Optional: Format style (0=cardinal, 1=ordinal, 2=currency)"
        return ""

    def getProgrammaticCategoryName(self, aProgrammaticName):
        if aProgrammaticName == "NUMTOWORDS":
            return "Text"
        return "Add-In"

    def getDisplayCategoryName(self, aProgrammaticName):
        if aProgrammaticName == "NUMTOWORDS":
            return "Text"
        return "Add-In"

    # The actual function implementation
    def NUMTOWORDS(self, number, format_style=0):
        """
        Convert a number to words.

        Args:
            number: The number to convert (float or integer)
            format_style: 0=cardinal (default), 1=ordinal, 2=currency

        Returns:
            String representation of the number in words
        """
        try:
            # Handle empty or error values
            if number is None:
                return ""

            # Convert to float if it's a string
            if isinstance(number, str):
                try:
                    num = float(number)
                except ValueError:
                    return "#VALUE!"
            else:
                num = float(number)

            # Call the actual conversion function
            result = self._number_to_words(num, format_style)
            return result

        except Exception as e:
            # Return error for any unexpected issues
            return "#ERROR!"

    def _number_to_words(self, num, format_style=0):
        """
        Convert number to words with support for large numbers and decimals.
        Uses a simpler iterative approach.
        """
        # Handle negative numbers
        if num < 0:
            return "minus " + self._number_to_words(abs(num), format_style)

        # Handle zero
        if num == 0:
            if format_style == 1:  # ordinal
                return "zeroth"
            return "zero"

        # Split into integer and decimal parts
        integer_part = int(num)
        decimal_part = int(round((num - integer_part) * 100))

        # Basic number words
        ones = [
            "",
            "one",
            "two",
            "three",
            "four",
            "five",
            "six",
            "seven",
            "eight",
            "nine",
        ]
        teens = [
            "ten",
            "eleven",
            "twelve",
            "thirteen",
            "fourteen",
            "fifteen",
            "sixteen",
            "seventeen",
            "eighteen",
            "nineteen",
        ]
        tens = [
            "",
            "",
            "twenty",
            "thirty",
            "forty",
            "fifty",
            "sixty",
            "seventy",
            "eighty",
            "ninety",
        ]

        def convert_below_thousand(n):
            """Convert numbers 0-999 to words."""
            if n == 0:
                return ""

            result = ""

            # Hundreds
            if n >= 100:
                result += ones[n // 100] + " hundred"
                n %= 100
                if n > 0:
                    result += " and "

            # Tens and ones
            if n >= 20:
                result += tens[n // 10]
                n %= 10
                if n > 0:
                    result += "-" + ones[n]
            elif n >= 10:
                result += teens[n - 10]
            elif n > 0:
                result += ones[n]

            return result

        # Large number scales
        scales = [
            (10**12, "trillion"),
            (10**9, "billion"),
            (10**6, "million"),
            (10**3, "thousand"),
            (1, ""),
        ]

        # Convert integer part
        if integer_part == 0:
            words = "zero"
        else:
            parts = []
            for scale_value, scale_name in scales:
                if integer_part >= scale_value:
                    scale_part = integer_part // scale_value
                    integer_part %= scale_value

                    if scale_part > 0:
                        scale_words = convert_below_thousand(scale_part)
                        if scale_words:
                            if scale_name:
                                parts.append(f"{scale_words} {scale_name}")
                            else:
                                parts.append(scale_words)

            words = " ".join(parts)

        # Handle format styles
        if format_style == 1:  # ordinal
            # Simple ordinal conversion rules
            if words.endswith("one"):
                words = words[:-3] + "first"
            elif words.endswith("two"):
                words = words[:-3] + "second"
            elif words.endswith("three"):
                words = words[:-5] + "third"
            elif words.endswith("ve"):
                words = words[:-2] + "fth"
            elif words.endswith("t"):
                words += "h"
            elif words.endswith("e"):
                words = words[:-1] + "th"
            else:
                words += "th"

        elif format_style == 2:  # currency
            words += " dollars"
            if decimal_part > 0:
                cents_words = self._number_to_words(decimal_part, 0)
                words += " and " + cents_words + " cents"

        elif format_style == 0 and decimal_part > 0:
            # Cardinal with decimal
            # Convert decimal part digit by digit
            decimal_str = str(decimal_part).rjust(2, "0")
            decimal_words = []
            for digit in decimal_str:
                if digit == "0":
                    decimal_words.append("zero")
                else:
                    decimal_words.append(ones[int(digit)])
            words += " point " + " ".join(decimal_words)

        return words.strip()


# Register the implementation
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    NumToWordsAddIn,
    "com.sun.star.sheet.addin.NumToWords",
    ("com.sun.star.sheet.AddIn",),
)
