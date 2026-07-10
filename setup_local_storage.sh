#!/bin/bash
# setup_local_storage.sh
# Creates local storage directory structure for testing

set -e

echo "🔧 Setting up local storage directories..."

# Create directory structure
mkdir -p local_storage/{input-json,templates,output-pptx}

echo "✅ Directories created:"
echo "   - local_storage/input-json"
echo "   - local_storage/templates"
echo "   - local_storage/output-pptx"

# Check if template exists
TEMPLATE="local_storage/templates/PowerpointTemplate_BoardMeetingGovernance.pptx"
if [ ! -f "$TEMPLATE" ]; then
    echo ""
    echo "⚠️  Template file not found!"
    echo "   Please copy your PowerPoint template to:"
    echo "   $TEMPLATE"
    echo ""
else
    echo ""
    echo "✅ Template found: $TEMPLATE"
fi

# Check for test JSON
JSON_FILES=$(find local_storage/input-json -name "*.json" 2>/dev/null | wc -l)
if [ "$JSON_FILES" -eq 0 ]; then
    echo ""
    echo "⚠️  No test JSON files found!"
    echo "   Please add test JSON files to:"
    echo "   local_storage/input-json/"
    echo ""
else
    echo ""
    echo "✅ Found $JSON_FILES JSON file(s) in input-json/"
fi

echo ""
echo "📋 Local storage structure:"
tree local_storage 2>/dev/null || find local_storage -type f

echo ""
echo "🚀 Ready to run! Use: func start"
