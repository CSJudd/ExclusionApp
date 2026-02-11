#!/bin/bash
# Initialize Git Repository for Exclusion Screening Tool

echo "=========================================="
echo "  Git Repository Setup"
echo "=========================================="
echo ""

cd "$(dirname "$0")"

# Initialize git if not already done
if [ ! -d ".git" ]; then
    echo "Initializing Git repository..."
    git init
    echo "✓ Git initialized"
else
    echo "Git repository already exists"
fi

# Add all files
echo ""
echo "Staging files..."
git add .

# Show what will be committed
echo ""
echo "Files to be committed:"
git status --short

echo ""
echo "=========================================="
echo "Next steps:"
echo "=========================================="
echo ""
echo "1. Review files to commit:"
echo "   git status"
echo ""
echo "2. Make initial commit:"
echo "   git commit -m 'Initial commit: Exclusion screening tool'"
echo ""
echo "3. Create GitHub repo, then:"
echo "   git remote add origin <your-repo-url>"
echo "   git branch -M main"
echo "   git push -u origin main"
echo ""
echo "Note: .gitignore is configured to EXCLUDE:"
echo "  - Client data files (CSV, Excel)"
echo "  - OIG/SAM databases"
echo "  - Virtual environments"
echo "  - Build artifacts"
echo ""
