# TN-OfficeDebug Release Configuration

## Files to Include in Release Package

### Essential Application Files
- src/taskpane.html
- src/taskpane.js  
- src/taskpane.css
- src/commands.html

### Configuration Files
- config/manifest.xml (URLs will be sanitized)
- config/web.config

### Documentation
- README.md
- LICENSE
- DEPLOYMENT.md (auto-generated)

## Files to Exclude from Release

### Development Files
- .git/
- .github/ (except for reference)
- node_modules/
- scripts/
- docs/ (development docs)
- test-standalone.html
- *.log
- *.tmp

### IDE and Editor Files
- .vscode/
- .vs/
- *.code-workspace
- .editorconfig

### System Files
- .DS_Store
- Thumbs.db
- desktop.ini

## URL Sanitization Rules

### Development URLs to Replace
- `https://localhost:*` → `https://YOUR-DOMAIN.com/`
- `http://localhost:*` → `https://YOUR-DOMAIN.com/`
- Any hardcoded domain URLs → `https://YOUR-DOMAIN.com/`

### Path Adjustments for Flat Deployment
- `/src/` paths → `/` (root level)
- `/config/` paths → `/` (root level)

## Version Synchronization Requirements

Before creating a release, ensure versions match in:

1. **config/manifest.xml**: `<Version>X.Y.Z</Version>`
2. **src/taskpane.js**: `const ADDIN_VERSION = 'X.Y.Z'`
3. **Git tag**: `vX.Y.Z`

## Release Workflow Triggers

### Automatic Release (GitHub Actions)
- Push git tag matching pattern `v*` (e.g., `v1.0.14`)
- Validates version consistency
- Generates changelog from conventional commits
- Creates ZIP package with sanitized files
- Publishes GitHub release

### Manual Release (Local Testing)
```powershell
# PowerShell
.\scripts\prepare-release.ps1 -Version "1.0.14"
```

```bash
# Bash
./scripts/prepare-release.sh 1.0.14
```

## Deployment Package Structure

After processing, the release package contains:

```
TN-OfficeDebug-X.Y.Z/
├── taskpane.html
├── taskpane.js
├── taskpane.css
├── commands.html
├── manifest.xml          # URLs sanitized
├── web.config             # Production-ready
├── README.md
├── LICENSE
└── DEPLOYMENT.md          # Deployment instructions
```

## Quality Checks

The release process performs these validations:

- ✅ Version consistency across files
- ✅ Essential files present
- ✅ URLs properly sanitized
- ✅ No development artifacts included
- ✅ Manifest XML syntax validation
- ✅ Deployment instructions generated

## Post-Release Steps

After successful release:

1. Test deployment package in clean environment
2. Validate manifest loads in Office
3. Update documentation if needed
4. Announce release to stakeholders