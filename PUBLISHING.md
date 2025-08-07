# Publishing Instructions

Follow these steps to publish the package to GitHub and npm:

## 1. Push to GitHub

### First Time Setup:
```bash
# Initialize git (if not already done)
git init

# Add the remote repository
git remote add origin https://github.com/mertdeveci/univerjs-import-export.git

# Add all files
git add .

# Commit the changes
git commit -m "Initial commit: Complete import/export library for Univer"

# Push to GitHub
git push -u origin main
```

### Subsequent Updates:
```bash
# Add changes
git add .

# Commit with descriptive message
git commit -m "Your commit message"

# Push to GitHub
git push
```

## 2. Publish to npm

### First Time Setup:

1. **Create an npm account** (if you don't have one):
   - Go to https://www.npmjs.com/signup
   - Create your account

2. **Login to npm from terminal**:
```bash
npm login
# Enter your username, password, and email
```

3. **Check the package name availability**:
```bash
npm view @mertdeveci55/univer-import-export
# If it returns an error, the name is available
```

### Publishing:

1. **Ensure everything is built**:
```bash
npm run build
```

2. **Test the package locally** (optional but recommended):
```bash
npm pack
# This creates a .tgz file you can test in another project
```

3. **Publish to npm**:
```bash
# For private/restricted package (only you and invited users can install)
npm publish --access restricted

# Or if you want it public under your scope
npm publish --access public

# For subsequent releases (after updating version in package.json)
npm version patch  # or minor/major
npm publish
```

## 3. Using in alphafrontend

After publishing, install in your alphafrontend project:

```bash
# Remove old package if exists
npm uninstall @mertdeveci55/luckyexcel-fixed

# Install new package
npm install @mertdeveci55/univer-import-export
```

Then update your imports:
```javascript
// Old
import { LuckyExcel } from '@mertdeveci55/luckyexcel-fixed';

// New
import { LuckyExcel } from '@mertdeveci55/univer-import-export';
```

## Version Management

- **Patch** (0.1.0 → 0.1.1): Bug fixes and minor changes
- **Minor** (0.1.0 → 0.2.0): New features, backward compatible
- **Major** (0.1.0 → 1.0.0): Breaking changes

Use npm version commands:
```bash
npm version patch -m "Fix: %s"
npm version minor -m "Feature: %s"
npm version major -m "Breaking: %s"
```

## Troubleshooting

### If npm publish fails:

1. **Check if you're logged in**:
```bash
npm whoami
```

2. **Check registry**:
```bash
npm config get registry
# Should be: https://registry.npmjs.org/
```

3. **If package name is taken**, update package.json:
   - Change name to something unique like `@yourusername/univer-import-export`

4. **For scoped packages** (@univerjs/...):
   - You might need permission from the @univerjs organization
   - Or use your own scope: `@mertdeveci/univer-import-export`

### If GitHub push fails:

1. **Check remote**:
```bash
git remote -v
```

2. **Check branch**:
```bash
git branch
# Ensure you're on main/master
```

3. **Force push if needed** (careful!):
```bash
git push -f origin main
```

## Maintenance

1. **Keep dependencies updated**:
```bash
npm outdated
npm update
```

2. **Run tests before publishing**:
```bash
npm test  # if tests are configured
npm run build
```

3. **Update README** when adding features

4. **Tag releases on GitHub**:
```bash
git tag v0.1.0
git push --tags
```

## Summary Checklist

Before each release:
- [ ] All features tested
- [ ] Build successful (`npm run build`)
- [ ] Version bumped in package.json
- [ ] README updated if needed
- [ ] Committed to git
- [ ] Pushed to GitHub
- [ ] Published to npm
- [ ] Tagged release on GitHub