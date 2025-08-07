#!/bin/bash

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}🚀 Starting publish process...${NC}"

# Check if we're in the right directory
if [ ! -f "package.json" ]; then
    echo -e "${RED}Error: package.json not found. Are you in the right directory?${NC}"
    exit 1
fi

# Get current version
CURRENT_VERSION=$(node -p "require('./package.json').version")
echo -e "${YELLOW}Current version: ${CURRENT_VERSION}${NC}"

# Build the project
echo -e "${GREEN}📦 Building project...${NC}"
npm run build
if [ $? -ne 0 ]; then
    echo -e "${RED}Build failed!${NC}"
    exit 1
fi

# Increment version
echo -e "${GREEN}📝 Incrementing version...${NC}"
npm version patch --no-git-tag-version
NEW_VERSION=$(node -p "require('./package.json').version")
echo -e "${GREEN}New version: ${NEW_VERSION}${NC}"

# Git operations
echo -e "${GREEN}🔧 Committing changes...${NC}"
git add -A
COMMIT_MESSAGE="Release v${NEW_VERSION}

Changes:
- Added detection and recovery of missing empty sheets
- Fixed Excel sheets with >>> characters being excluded from import
- Hardcoded addition of known missing sheets (Financial Model>>>, DCF>>>, LBO>>>)
- Sheets are now properly positioned even when not in workbook.xml
- Preserved console.log statements for debugging by disabling terser drop_console

🤖 Generated with automated publish script"

git commit -m "$COMMIT_MESSAGE"

echo -e "${GREEN}📤 Pushing to GitHub...${NC}"
git push origin master

# Publish to npm
echo -e "${GREEN}🎉 Publishing to npm...${NC}"
npm publish

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✅ Successfully published version ${NEW_VERSION}!${NC}"
    echo ""
    echo -e "${YELLOW}📋 Next steps:${NC}"
    echo "1. Go to alphafrontend directory: cd ../alphafrontend"
    echo "2. Update the package: pnpm remove @mertdeveci55/univer-import-export && pnpm add @mertdeveci55/univer-import-export@${NEW_VERSION}"
    echo "3. Restart dev server: npm run dev"
    echo "4. In browser: Hard refresh (Cmd+Shift+R on Mac, Ctrl+Shift+F5 on Windows)"
    echo "5. Clear browser cache if needed: DevTools > Application > Storage > Clear site data"
else
    echo -e "${RED}❌ npm publish failed!${NC}"
    exit 1
fi