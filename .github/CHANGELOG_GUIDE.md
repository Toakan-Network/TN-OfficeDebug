# Changelog Generation Configuration

This project uses conventional commits to automatically generate changelogs for releases.

## Conventional Commit Format

Commits should follow this format: `type(scope): description`

### Supported Types:

- **feat**: New features (✨ Feature)
- **fix**: Bug fixes (🐛 Fix)
- **docs**: Documentation changes (📚 Documentation)
- **style**: Code style changes (💎 Style)
- **refactor**: Code refactoring (♻️ Refactor)
- **test**: Test additions/modifications (🧪 Test)
- **chore**: Maintenance tasks (🔧 Maintenance)

### Examples:

```
feat(ui): add document bookmark detection
fix(api): resolve custom properties loading error
docs: update deployment instructions
style(css): improve card layout spacing
refactor(security): streamline domain detection
test(integration): add manifest validation tests
chore(deps): update Office.js version
```

## Scope Guidelines:

- **ui**: User interface changes
- **api**: Office.js API interactions
- **security**: Security features and domain restrictions  
- **config**: Configuration and manifest changes
- **manifest**: Manifest file updates
- **docs**: Documentation updates
- **release**: Release preparation and versioning

## Automatic Changelog Generation:

The release workflow automatically generates changelogs by:

1. Finding the previous git tag
2. Collecting all commits since that tag
3. Categorizing commits by type
4. Formatting with appropriate emojis
5. Including commit messages in the release notes

## Manual Changelog:

For major releases, you can also create manual changelog entries in `CHANGELOG.md` following this format:

```markdown
# Changelog

## [1.0.14] - 2025-10-03

### ✨ Added
- New diagnostic feature for bookmark detection
- Enhanced error reporting system

### 🐛 Fixed  
- Custom properties display issue
- Memory leak in document scanner

### 📚 Documentation
- Updated deployment guide
- Added troubleshooting section

### 🔧 Maintenance
- Updated dependencies
- Improved build process
```

## Release Process:

1. Ensure all commits follow conventional format
2. Update version in `config/manifest.xml` and `src/taskpane.js`
3. Run `scripts/prepare-release.ps1 -Version "x.y.z"` to test locally
4. Create and push version tag: `git tag vx.y.z && git push origin vx.y.z`
5. GitHub Actions automatically creates release with changelog