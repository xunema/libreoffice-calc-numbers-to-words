#!/usr/bin/env python3
"""
Test script for number to words conversion logic.
This tests the core conversion algorithms without requiring LibreOffice UNO.
"""


def test_conversion_logic():
    """Test the number conversion logic used in the add-in."""

    # Test data: (input_number, format_style, expected_output)
    test_cases = [
        # Cardinal numbers (format 0)
        (0, 0, "zero"),
        (1, 0, "one"),
        (7, 0, "seven"),
        (10, 0, "ten"),
        (11, 0, "eleven"),
        (20, 0, "twenty"),
        (21, 0, "twenty-one"),
        (100, 0, "one hundred"),
        (101, 0, "one hundred and one"),
        (123, 0, "one hundred and twenty-three"),
        (1000, 0, "one thousand"),
        (1234, 0, "one thousand two hundred and thirty-four"),
        (1000000, 0, "one million"),
        (
            1234567,
            0,
            "one million two hundred and thirty-four thousand five hundred and sixty-seven",
        ),
        # Negative numbers
        (-5, 0, "minus five"),
        (-123, 0, "minus one hundred and twenty-three"),
        # Ordinal numbers (format 1)
        (1, 1, "first"),
        (2, 1, "second"),
        (3, 1, "third"),
        (4, 1, "fourth"),
        (5, 1, "fifth"),
        (21, 1, "twenty-first"),
        # Test numbers that would fail certain edge cases
        (1002, 0, "one thousand two"),
        (1100, 0, "one thousand one hundred"),
    ]

    # Simplified conversion logic for testing
    def simple_number_to_words(num, format_style=0):
        """Simplified version of the conversion logic for testing."""
        # Handle negative numbers
        if num < 0:
            return "minus " + simple_number_to_words(abs(num), format_style)

        # Handle zero
        if num == 0:
            if format_style == 1:  # ordinal
                return "zeroth"
            return "zero"

        # Basic number words (same as in numtowords.py)
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
        integer_part = int(num)
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

        # Handle format styles (simplified)
        if format_style == 1:  # ordinal
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

        return words.strip()

    # Run tests
    print("Testing number to words conversion logic...")
    print("=" * 60)

    passed = 0
    failed = 0

    for i, (num, style, expected) in enumerate(test_cases, 1):
        result = simple_number_to_words(num, style)
        if result.lower() == expected.lower():
            print(f"✓ Test {i:2d}: {num:10} (style={style}) -> '{result}'")
            passed += 1
        else:
            print(f"✗ Test {i:2d}: {num:10} (style={style})")
            print(f"  Expected: '{expected}'")
            print(f"  Got:      '{result}'")
            failed += 1

    print("=" * 60)
    print(f"Results: {passed} passed, {failed} failed")

    if failed == 0:
        print("✅ All tests passed!")
        return True
    else:
        print("❌ Some tests failed.")
        return False


if __name__ == "__main__":
    success = test_conversion_logic()
    exit(0 if success else 1)
