# LibreOffice Calc Number to Words Converter

A LibreOffice Calc add-in that provides a `NUMTOWORDS()` function to convert numbers to their word representation.

## Features

- **Cardinal Numbers**: Convert numbers to English words (e.g., 123 → "one hundred twenty-three")
- **Ordinal Numbers**: Convert to ordinal form (e.g., 3 → "third", 21 → "twenty-first")
- **Currency Format**: Format as dollars and cents (e.g., 123.45 → "one hundred twenty-three dollars and forty-five cents")
- **Large Numbers**: Handle numbers up to trillions
- **Decimal Support**: Convert decimal numbers (e.g., 12.34 → "twelve point three four")
- **Negative Numbers**: Support for negative values (e.g., -5 → "minus five")

## Quick Start

### Installation (Pre-built Extension)

1. **Download** the latest `.oxt` file from the [Releases](https://github.com/xunema/libreoffice-calc-numbers-to-words/releases) page.

2. **Open LibreOffice Calc** and go to `Tools` → `Extension Manager`.

3. **Click "Add"** and select the downloaded `numtowords.oxt` file.

4. **Restart LibreOffice** to activate the extension.

### Installation (Build from Source)

If you want to build the extension yourself:

```bash
# Clone the repository
git clone https://github.com/xunema/libreoffice-calc-numbers-to-words.git
cd libreoffice-calc-numbers-to-words

# Build the extension
./build.sh

# The extension file numtowords.oxt will be created
```

Then follow the installation steps above using your built `.oxt` file.

## Usage Examples

### Basic Usage

| Formula | Result |
|---------|--------|
| `=NUMTOWORDS(123)` | "one hundred and twenty-three" |
| `=NUMTOWORDS(4567)` | "four thousand five hundred and sixty-seven" |
| `=NUMTOWORDS(1000000)` | "one million" |
| `=NUMTOWORDS(-42)` | "minus forty-two" |

### Format Options

| Formula | Result |
|---------|--------|
| `=NUMTOWORDS(123, 0)` | "one hundred and twenty-three" (cardinal, default) |
| `=NUMTOWORDS(123, 1)` | "one hundred and twenty-third" (ordinal) |
| `=NUMTOWORDS(123.45, 2)` | "one hundred and twenty-three dollars and forty-five cents" (currency) |

### Advanced Examples

| Formula | Result |
|---------|--------|
| `=NUMTOWORDS(21, 1)` | "twenty-first" |
| `=NUMTOWORDS(1002)` | "one thousand two" |
| `=NUMTOWORDS(15.50, 2)` | "fifteen dollars and fifty cents" |
| `=NUMTOWORDS(0.99)` | "zero point nine nine" |
| `=NUMTOWORDS(123456789)` | "one hundred twenty-three million four hundred fifty-six thousand seven hundred eighty-nine" |

### Using with Cell References

You can also use the function with cell references:

```
A1: 123.45
B1: =NUMTOWORDS(A1, 2)  → "one hundred twenty-three dollars and forty-five cents"
```

## Development

### Project Structure

```
libreoffice-calc-numbers-to-words/
├── numtowords.py          # Main Python implementation
├── description.xml        # Extension metadata
├── META-INF/manifest.xml  # Package manifest for LibreOffice
├── build.sh              # Build script
├── LICENSE               # GNU GPL v3.0 license
├── README.md            # This documentation
└── docs/                 # Architecture and documentation
    ├── architecture.md   # Technical architecture
    └── wiki-home.md     # Wiki home page
```

### Building the Extension

The `build.sh` script creates the `.oxt` package:

```bash
./build.sh
```

This creates `numtowords.oxt` which contains:
- `numtowords.py` - Python implementation
- `description.xml` - Extension metadata
- `license.txt` - License information
- `META-INF/manifest.xml` - LibreOffice package manifest

### Technical Details

The add-in implements the LibreOffice UNO API with Python. For detailed architecture and implementation details, see the [Architecture Documentation](docs/architecture.md).

## Troubleshooting

### Extension Not Appearing
- Make sure you restarted LibreOffice after installation
- Check the Extension Manager to verify the add-in is listed
- Try reinstalling the extension

### Function Not Working
- Ensure you're using LibreOffice Calc (not Writer or other components)
- Check for syntax errors in your formula
- Verify the extension is enabled in Extension Manager

### Building Issues
- Ensure you have `zip` installed on your system
- Make sure all required files are present in the directory

## License

This project is licensed under the GNU General Public License v3.0. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

For bug reports or feature requests, please open an issue on GitHub.

## Support

- **GitHub Issues**: [Report issues or request features](https://github.com/xunema/libreoffice-calc-numbers-to-words/issues)
- **Documentation**: [Architecture and development documentation](docs/architecture.md)
