# Architecture: LibreOffice Calc Number to Words Add-in

## Overview

This document describes the architecture and implementation of the LibreOffice Calc Number to Words add-in. The add-in provides a `NUMTOWORDS()` function that converts numeric values to their English word representation.

## System Architecture

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    LibreOffice Calc                         │
├─────────────────────────────────────────────────────────────┤
│  UNO (Universal Network Objects) API                        │
│  ┌─────────────────────────────────────────────────────┐    │
│  │              NumToWordsAddIn Class                  │    │
│  │  ┌─────────────────┐  ┌─────────────────────────┐  │    │
│  │  │ XAddIn Interface│  │ Conversion Engine       │  │    │
│  │  └─────────────────┘  └─────────────────────────┘  │    │
│  └─────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────┘
```

### Component Breakdown

1. **UNO Component Interface** (`XAddIn`, `XServiceName`, `XLocalizable`)
2. **Add-in Implementation Class** (`NumToWordsAddIn`)
3. **Number Conversion Engine** (`_number_to_words` method)
4. **Packaging System** (OXT file with manifest and metadata)

## Detailed Component Design

### 1. UNO Component Interface

The add-in implements several UNO interfaces to integrate with LibreOffice:

#### `XAddIn` Interface
- **Purpose**: Provides metadata about the add-in function
- **Key Methods**:
  - `getProgrammaticFunctionName()`: Maps display names to internal names
  - `getDisplayFunctionName()`: Maps internal names to display names
  - `getFunctionDescription()`: Provides function description for UI
  - `getDisplayArgumentName()`: Names function arguments for UI
  - `getArgumentDescription()`: Describes function arguments
  - `getProgrammaticCategoryName()`: Function category (internal)
  - `getDisplayCategoryName()`: Function category (display)

#### `XServiceName` Interface
- **Purpose**: Identifies the service name
- **Key Method**: `getServiceName()`: Returns `"com.sun.star.sheet.addin.NumToWords"`

#### `XLocalizable` Interface
- **Purpose**: Supports localization (currently set to English/US)
- **Key Methods**: `setLocale()` and `getLocale()`

### 2. NumToWordsAddIn Class

The main class that implements all interfaces:

```python
class NumToWordsAddIn(unohelper.Base, XServiceName, XAddIn, XLocalizable):
    def __init__(self, ctx):
        self.ctx = ctx
        self.locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        self.locale.Language = "en"
        self.locale.Country = "US"
    
    # Interface implementations...
    def NUMTOWORDS(self, number, format_style=0):
        # Main function implementation
```

### 3. Number Conversion Engine

The core conversion logic is implemented in the `_number_to_words()` method:

#### Conversion Strategy

1. **Input Validation**: Handles None values, strings, and error cases
2. **Number Decomposition**: Splits into integer and decimal parts
3. **Scale Processing**: Processes number in chunks (trillions, billions, etc.)
4. **Word Mapping**: Converts numeric values to English words
5. **Format Application**: Applies cardinal, ordinal, or currency formatting

#### Data Structures

```python
# Number word mappings
ones = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]
teens = ["ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", 
         "sixteen", "seventeen", "eighteen", "nineteen"]
tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", 
        "seventy", "eighty", "ninety"]

# Scale definitions (supports up to trillions)
scales = [
    (10**12, "trillion"),
    (10**9, "billion"),
    (10**6, "million"),
    (10**3, "thousand"),
    (1, "")
]
```

#### Helper Functions

- `convert_below_thousand()`: Converts numbers 0-999 to words
- Handles hundreds, tens, and ones with proper hyphenation
- Manages "and" conjunction between hundreds and tens/ones

### 4. Format Styles

The add-in supports three format styles:

#### Cardinal (Format 0)
- Default format
- Example: `123` → "one hundred and twenty-three"
- Decimal handling: `12.34` → "twelve point three four"

#### Ordinal (Format 1)
- Converts to ordinal numbers
- Rules for different endings:
  - "one" → "first"
  - "two" → "second" 
  - "three" → "third"
  - "ve" → "fth" (five → fifth)
  - "t" → "th" (eight → eighth)
  - "e" → "th" (nine → ninth)
  - Default: add "th"

#### Currency (Format 2)
- Formats as dollars and cents
- Example: `123.45` → "one hundred twenty-three dollars and forty-five cents"
- Handles both integer and decimal parts

## Error Handling

### Input Validation
- **None values**: Returns empty string
- **Invalid strings**: Returns `#VALUE!` error
- **General exceptions**: Returns `#ERROR!`

### Edge Cases
- **Zero**: Properly handled as "zero" or "zeroth" for ordinal
- **Negative numbers**: Prefixed with "minus"
- **Large decimals**: Rounded to two decimal places for currency

## Packaging Architecture

### OXT File Structure
```
numtowords.oxt
├── numtowords.py           # Main Python implementation
├── description.xml         # Extension metadata
├── license.txt            # License information
└── META-INF/
    └── manifest.xml       # LibreOffice package manifest
```

### Manifest File
```xml
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
    <manifest:file-entry
        manifest:media-type="application/vnd.sun.star.uno-component;type=Python"
        manifest:full-path="numtowords.py"/>
    <manifest:file-entry
        manifest:media-type="application/vnd.sun.star.package-bundle-description"
        manifest:full-path="description.xml"/>
</manifest:manifest>
```

### Description Metadata
- **Identifier**: `org.xunema.libreoffice.calc.numtowords`
- **Version**: Semantic versioning
- **Platform**: All platforms supported
- **License**: GNU GPL v3.0

## Build System

### Build Script (`build.sh`)
```bash
#!/bin/bash
# Build script for LibreOffice Number to Words extension

set -e

echo "Building LibreOffice Calc Number to Words extension..."

# Create the OXT package
zip -r numtowords.oxt \
    numtowords.py \
    description.xml \
    license.txt \
    META-INF/manifest.xml

echo "Created numtowords.oxt"
echo "Size: $(du -h numtowords.oxt | cut -f1)"
```

## Integration with LibreOffice

### Registration Process
1. **Component Discovery**: LibreOffice scans OXT file for Python components
2. **Service Registration**: `g_ImplementationHelper.addImplementation()` registers the service
3. **Function Availability**: `NUMTOWORDS()` becomes available in Calc function wizard

### UNO Component Lifecycle
1. **Initialization**: Context passed during construction
2. **Function Call**: LibreOffice calls `NUMTOWORDS()` with parameters
3. **Result Return**: String result returned to Calc cell

## Code Walkthrough

### Key Methods

#### `NUMTOWORDS()` - Main Function
```python
def NUMTOWORDS(self, number, format_style=0):
    """
    Convert a number to words.
    
    Args:
        number: The number to convert (float or integer)
        format_style: 0=cardinal (default), 1=ordinal, 2=currency
    
    Returns:
        String representation of the number in words
    """
    # Input validation and type conversion
    # Call to internal conversion method
    # Error handling
```

#### `_number_to_words()` - Core Conversion
```python
def _number_to_words(self, num, format_style=0):
    """
    Convert number to words with support for large numbers and decimals.
    Uses a simpler iterative approach.
    """
    # Handle negative numbers and zero
    # Split into integer and decimal parts
    # Convert using scales and helper functions
    # Apply format-specific transformations
```

#### `convert_below_thousand()` - Helper Function
```python
def convert_below_thousand(n):
    """Convert numbers 0-999 to words."""
    # Handle hundreds
    # Handle tens and ones
    # Proper hyphenation and conjunctions
```

## Extension Points

### Localization
The current implementation is hardcoded for English (US). To add support for other languages:

1. **Locale Support**: Extend `XLocalizable` implementation
2. **Word Sets**: Create language-specific word mappings
3. **Grammar Rules**: Implement language-specific ordinal and plural rules

### Additional Number Formats
Easy to add new format styles by extending the format handling in `_number_to_words()`.

### Performance Optimizations
- **Caching**: Could cache frequently used conversions
- **Memoization**: Store results for repeated numbers

## Testing Strategy

### Unit Testing
- **Number Conversion**: Test individual conversion functions
- **Format Styles**: Verify cardinal, ordinal, and currency formats
- **Edge Cases**: Test zero, negative, large, and decimal numbers

### Integration Testing
- **LibreOffice Integration**: Test within actual Calc environment
- **Function Wizard**: Verify metadata appears correctly
- **Error Handling**: Test error conditions and error messages

## Dependencies

### Runtime Dependencies
- **LibreOffice 4.0+**: Required for UNO Python support
- **Python 3.x**: Included with LibreOffice
- **UNO Python Bindings**: Provided by LibreOffice installation

### Development Dependencies
- **Python 3.x**: For testing and development
- **zip**: For creating OXT packages
- **Git**: For version control

## Limitations and Future Improvements

### Current Limitations
1. **English Only**: Only supports English number words
2. **Scale Limit**: Supports up to trillions (could be extended)
3. **Decimal Precision**: Currency rounds to two decimal places

### Future Enhancements
1. **Multi-language Support**: Add support for other languages
2. **Custom Currency**: Allow different currency names
3. **Scientific Notation**: Support for very large/small numbers
4. **Roman Numerals**: Additional conversion options
5. **Performance**: Add caching for frequently used numbers

## Conclusion

The LibreOffice Calc Number to Words add-in demonstrates:
- **Clean Architecture**: Separation of concerns between UNO interface and conversion logic
- **Extensible Design**: Easy to add new formats or languages
- **Robust Implementation**: Comprehensive error handling and edge case management
- **Standard Compliance**: Follows LibreOffice extension standards

The modular design allows for easy maintenance and extension, making it a solid foundation for future enhancements.