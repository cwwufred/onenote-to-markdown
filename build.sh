#!/bin/bash
# Build script for OneNote2MD

echo "🚀 Building OneNote2MD..."

# Clean previous builds
echo "📦 Cleaning previous builds..."
rm -rf build/ dist/ *.egg-info

# Install build dependencies if needed
echo "📚 Installing build dependencies..."
pip install --upgrade build wheel 2>/dev/null || python3 -m pip install --upgrade build wheel 2>/dev/null

# Build package
echo "🔨 Building package..."
python3 -m build

# Check for dist
if [ -d "dist" ]; then
    echo "✅ Build complete! Files in dist/:"
    ls -la dist/
else
    echo "❌ Build failed"
fi

echo ""
echo "📦 To create a wheel (.whl):"
echo "   pip install twine"
echo "   twine check dist/*"
echo ""
echo "🐍 To create a standalone executable:"
echo "   pip install pyinstaller"
echo "   pyinstaller onenote2md/gui.py --onefile --name onenote2md"