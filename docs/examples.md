# Usage Examples

This document provides comprehensive examples of using the `NUMTOWORDS()` function in LibreOffice Calc.

## Basic Syntax

```
=NUMTOWORDS(number, [format])
```

- `number`: The numeric value to convert (can be a number, cell reference, or formula)
- `format` (optional): Format style (0=cardinal, 1=ordinal, 2=currency). Default is 0.

## Example Table

| Formula | Result | Description |
|---------|--------|-------------|
| `=NUMTOWORDS(123)` | "one hundred and twenty-three" | Basic cardinal number |
| `=NUMTOWORDS(4567)` | "four thousand five hundred and sixty-seven" | Larger number |
| `=NUMTOWORDS(1000000)` | "one million" | Simple large number |
| `=NUMTOWORDS(123456789)` | "one hundred twenty-three million four hundred fifty-six thousand seven hundred eighty-nine" | Very large number |
| `=NUMTOWORDS(-42)` | "minus forty-two" | Negative number |
| `=NUMTOWORDS(0)` | "zero" | Zero |
| `=NUMTOWORDS(123, 1)` | "one hundred and twenty-third" | Ordinal number |
| `=NUMTOWORDS(21, 1)` | "twenty-first" | Ordinal with special rule |
| `=NUMTOWORDS(5, 1)` | "fifth" | Ordinal with "ve" to "fth" rule |
| `=NUMTOWORDS(123.45, 2)` | "one hundred and twenty-three dollars and forty-five cents" | Currency format |
| `=NUMTOWORDS(15.50, 2)` | "fifteen dollars and fifty cents" | Currency with .50 |
| `=NUMTOWORDS(0.99)` | "zero point nine nine" | Decimal number |
| `=NUMTOWORDS(12.34)` | "twelve point three four" | Decimal with two digits |

## Practical Examples

### 1. Invoice Amounts in Words

Create an invoice where the total amount is shown in words:

| Cell | Value | Formula | Result |
|------|-------|---------|--------|
| A1 | 1234.56 | (manual entry) | |
| B1 | =NUMTOWORDS(A1, 2) | =NUMTOWORDS(A1, 2) | "one thousand two hundred thirty-four dollars and fifty-six cents" |

### 2. Ranking System

Create a ranking system with ordinal numbers:

| Cell | Value | Formula | Result |
|------|-------|---------|--------|
| A1 | 1 | (manual entry) | |
| B1 | =NUMTOWORDS(A1, 1) | =NUMTOWORDS(A1, 1) | "first" |
| A2 | 2 | (manual entry) | |
| B2 | =NUMTOWORDS(A2, 1) | =NUMTOWORDS(A2, 1) | "second" |
| A3 | 3 | (manual entry) | |
| B3 | =NUMTOWORDS(A3, 1) | =NUMTOWORDS(A3, 1) | "third" |

### 3. Check Writing

Format amounts for check writing:

| Cell | Value | Formula | Result |
|------|-------|---------|--------|
| A1 | 150.75 | (manual entry) | |
| B1 | =NUMTOWORDS(A1, 2) | =NUMTOWORDS(A1, 2) | "one hundred fifty dollars and seventy-five cents" |

### 4. Combining with Other Functions

You can combine `NUMTOWORDS()` with other Calc functions:

| Formula | Result |
|---------|--------|
| `=NUMTOWORDS(SUM(A1:A10))` | Converts the sum of range A1:A10 to words |
| `=NUMTOWORDS(A1*B1, 2)` | Converts the product of A1 and B1 to currency words |
| `=NUMTOWORDS(ROUND(C1, 2), 2)` | Rounds C1 to 2 decimals then converts to currency |

### 5. Conditional Formatting in Words

Create a status report:

```
A1: 1
B1: =IF(A1=1, NUMTOWORDS(A1, 1) & " place", "Not first")
Result: "first place"
```

## Advanced Examples

### Large Number Conversion

The function handles numbers up to trillions:

| Number | Result |
|--------|--------|
| 1,000,000,000,000 | "one trillion" |
| 1,234,567,890,123 | "one trillion two hundred thirty-four billion five hundred sixty-seven million eight hundred ninety thousand one hundred twenty-three" |

### Decimal Precision

For currency, decimals are rounded to two places:

| Input | Result |
|-------|--------|
| 123.456 | "one hundred twenty-three dollars and forty-six cents" (rounded) |
| 123.454 | "one hundred twenty-three dollars and forty-five cents" (rounded) |

### Error Handling

The function handles errors gracefully:

| Input | Result |
|-------|--------|
| `=NUMTOWORDS("abc")` | `#VALUE!` |
| `=NUMTOWORDS("")` | "" (empty string) |

## Tips and Best Practices

1. **Cell References**: Always use cell references instead of hard-coded numbers for flexibility
2. **Currency Format**: Use format 2 for financial documents
3. **Error Checking**: Wrap in `IFERROR()` for cleaner sheets:
   ```
   =IFERROR(NUMTOWORDS(A1, 2), "Invalid amount")
   ```
4. **Combining Text**: Use `&` to combine with other text:
   ```
   ="Amount: " & NUMTOWORDS(B5, 2) & " only"
   ```
5. **Localization**: Currently English-only. For other languages, the source code would need modification.

## Common Use Cases

1. **Financial Documents**: Checks, invoices, receipts
2. **Legal Documents**: Contracts where amounts need to be written out
3. **Educational Materials**: Teaching number words
4. **Reports**: Making numbers more readable
5. **Accessibility**: Providing textual representation of numbers

## Limitations

1. **English Only**: Only supports English number words
2. **Scale Limit**: Up to trillions (can be extended in code)
3. **Decimal Handling**: For non-currency decimals, converts digit by digit ("point three four" not "thirty-four hundredths")

## Getting Help

If you encounter issues:
1. Check the formula syntax
2. Verify the extension is installed and enabled
3. Restart LibreOffice if the function doesn't appear
4. Check for error values in referenced cells

For more help, see the [README](../README.md) or open an issue on GitHub.