# LibreOffice Calc NumToWords Extension

A LibreOffice Calc add-in that converts numbers to their English word representation using a custom `=NUMTOWORDS()` spreadsheet function.

## Features

- **Cardinal numbers** — `123` → `one hundred and twenty-three`
- **Ordinal numbers** — `42` → `forty-second`
- **Currency format** — `99.99` → `ninety-nine dollars and ninety-nine cents`
- **Negative numbers** — `-5` → `minus five`
- **Decimals** — `12.34` → `twelve point three four`
- **Large numbers** — up to trillions

---

## Installation

### Option A — Extension Manager (GUI)

1. Download **[NumToWordsPy.oxt](https://github.com/xunema/libreoffice-calc-numbers-to-words/releases/download/v2.0.0/NumToWordsPy.oxt)** (latest release: v2.0.0)
2. Open LibreOffice Calc
3. Go to **Tools → Extension Manager**
4. Click **Add** and select `NumToWordsPy.oxt`
5. Restart LibreOffice Calc

### Option B — Command Line

```bash
# Download
curl -LO https://github.com/xunema/libreoffice-calc-numbers-to-words/releases/download/v2.0.0/NumToWordsPy.oxt

# Make sure LibreOffice is closed first, then install
unopkg add NumToWordsPy.oxt

# Verify it installed
unopkg list
```

Then open LibreOffice Calc and start using `=NUMTOWORDS()`.

---

## Usage

```
=NUMTOWORDS(number)
=NUMTOWORDS(number, formatStyle)
```

| formatStyle | Mode | Example |
|-------------|------|---------|
| `0` (default) | Cardinal | `=NUMTOWORDS(123)` → `one hundred and twenty-three` |
| `1` | Ordinal | `=NUMTOWORDS(42, 1)` → `forty-second` |
| `2` | Currency (USD) | `=NUMTOWORDS(99.99, 2)` → `ninety-nine dollars and ninety-nine cents` |

### More Examples

```
=NUMTOWORDS(0)            →  zero
=NUMTOWORDS(-5)           →  minus five
=NUMTOWORDS(1000000)      →  one million
=NUMTOWORDS(1001, 1)      →  one thousand first
=NUMTOWORDS(1, 2)         →  one dollar
=NUMTOWORDS(12.34)        →  twelve point three four
```

---

## Project Structure

```
libreoffice-calc-numbers-to-words/
├── python/                        # ← Working extension (Python UNO)
│   ├── NumToWordsPy.oxt           #   Ready-to-install extension package
│   ├── numtowords.uno.py          #   Python UNO component
│   ├── CalcAddIns.xcu             #   Calc function registration
│   ├── NumToWords.rdb             #   Compiled UNO type library
│   ├── description.xml            #   Extension metadata
│   └── META-INF/manifest.xml      #   OXT manifest
├── idl/
│   └── com/numbertext/converter/
│       └── NumToWords.idl         #   UNO interface definition (IDL source)
├── java/                          #   Java attempt (abandoned — see wiki)
└── README.md
```

---

## Reinstalling / Upgrading

If you are updating from a previous version, do a clean reinstall:

```bash
# Close LibreOffice first, then:
unopkg remove com.numbertext.numtowords-python
rm -rf ~/.config/libreoffice/4/user/uno_packages/cache/
rm -rf ~/.config/libreoffice/4/user/extensions/

unopkg add NumToWordsPy.oxt
```

---

## Technical Details

- **Language:** Python 3 (UNO bridge)
- **Pattern:** Implements a custom IDL interface (`NumToWordsConverter`) so LibreOffice's UNO introspection can discover and dispatch calls to `numToWords()` — the same technique used by [libnumbertext](https://github.com/Numbertext/libnumbertext)
- **No dependencies:** Pure Python standard library, no third-party packages required
- **Tested on:** LibreOffice 24.2 on Linux

---

## Wiki

For detailed documentation see the [GitHub Wiki](../../wiki):

- [How to Create a Custom LibreOffice Calc Function](../../wiki/How-to-Create-a-Custom-LibreOffice-Calc-Function) — step-by-step guide anyone can follow to build their own Calc add-in
- [Development Challenges and Debugging Log](../../wiki/Development-Challenges-and-Debugging-Log) — all the things that went wrong and how they were fixed

---

## License

GNU General Public License v3.0

---

## Version History

- **v2.0.0** — Rewritten as Python UNO extension. Working. Java approach abandoned due to LO 24.2 Java bridge instability.
- **v1.0.0** — Original Python prototype
