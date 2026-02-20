# LibreOffice Calc Number to Words Converter

A LibreOffice Calc add-in that provides a `NUMTOWORDS()` function to convert numbers to their word representation.

## Features

- Converts numbers to English words (cardinal form)
- Supports ordinal numbers (first, second, third, etc.)
- Supports currency format (dollars and cents)
- Handles large numbers up to trillions
- Supports negative numbers and decimals

## Installation

1. Download the latest `.oxt` extension file from the [Releases](https://github.com/xunema/libreoffice-calc-numbers-to-words/releases) page.

2. Open LibreOffice Calc.

3. Go to `Tools` → `Extension Manager` → `Add`.

4. Select the downloaded `.oxt` file and click `Open`.

5. Restart LibreOffice.

## Usage

In any LibreOffice Calc cell, use the `NUMTOWORDS()` function:

```
=NUMTOWORDS(number, [format])
```

### Parameters

- `number`: The number to convert to words
- `format` (optional): 
  - `0` or omitted: Cardinal number (default)
  - `1`: Ordinal number (first, second, third, etc.)
  - `2`: Currency format (dollars and cents)

### Examples

- `=NUMTOWORDS(123)` → "one hundred and twenty-three"
- `=NUMTOWORDS(123, 1)` → "one hundred and twenty-third"
- `=NUMTOWORDS(123.45, 2)` → "one hundred and twenty-three dollars and forty-five cents"
- `=NUMTOWORDS(-56.78)` → "minus fifty-six point seven eight"

## Building from Source

To build the extension from source:

```bash
# Create the OXT package
zip -r numtowords.oxt numtowords.py description.xml license.txt META-INF/
```

## Development

The main implementation is in `numtowords.py`. The add-in implements the LibreOffice UNO API with Python.

### Testing

Run the simple test script:

```bash
python3 simple_test.py
```

## License

GNU General Public License v3.0

## Contributing

Contributions are welcome! Please open an issue or pull request on GitHub.
