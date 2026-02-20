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