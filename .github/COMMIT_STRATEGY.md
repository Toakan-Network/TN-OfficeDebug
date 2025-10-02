# Git Commit Strategy for TN-OfficeDebug

This document outlines the commit strategy and guidelines for maintaining clean git history.

## Commit Message Format

Use conventional commits format:
```
<type>(<scope>): <subject>

<body>

<footer>
```

### Types
- **feat**: New feature
- **fix**: Bug fix
- **docs**: Documentation changes
- **style**: Code style changes (formatting, etc.)
- **refactor**: Code refactoring without feature changes
- **test**: Adding or updating tests
- **chore**: Maintenance tasks, build updates

### Scopes
- **ui**: User interface changes
- **api**: Office.js API integration
- **security**: Security-related changes
- **config**: Configuration changes
- **manifest**: Office manifest changes

## Examples

```bash
# New feature
git commit -m "feat(api): add Word document bookmark detection

- Implement bookmark enumeration in document info tab
- Add bookmark count and names display
- Include error handling for bookmark access"

# Bug fix
git commit -m "fix(ui): resolve text overflow in custom properties display

- Fix long property names wrapping correctly
- Adjust card layout for better mobile display
- Update CSS for consistent spacing"

# Documentation
git commit -m "docs: update troubleshooting guide with new error patterns

- Add solutions for bookmark access errors
- Include Office version compatibility notes
- Update API limitation documentation"

# Configuration
git commit -m "config(manifest): update production URLs for deployment

- Change localhost URLs to production domain
- Update security policy URLs
- Increment manifest version to 1.1.0"
```

## Automatic Commit Workflow

For future development, consider implementing automatic commits after significant changes:

```powershell
# Add this to development workflow
function Commit-Changes {
    param(
        [string]$Type,
        [string]$Scope = "",
        [string]$Message,
        [string]$Body = ""
    )
    
    $commitMsg = "$Type"
    if ($Scope) { $commitMsg += "($Scope)" }
    $commitMsg += ": $Message"
    if ($Body) { $commitMsg += "`n`n$Body" }
    
    git add -A
    git commit -m $commitMsg
}

# Usage examples:
# Commit-Changes -Type "feat" -Scope "ui" -Message "add new diagnostic panel"
# Commit-Changes -Type "fix" -Scope "api" -Message "resolve Office.js compatibility issue"
```

## Pre-commit Checklist

Before each commit, ensure:

1. [ ] Code is tested in Word desktop and online
2. [ ] No console errors in browser developer tools
3. [ ] All tabs load without errors
4. [ ] Documentation is updated if APIs changed
5. [ ] Version numbers updated if applicable

## Branch Strategy

- **main**: Production-ready code
- **develop**: Development integration branch
- **feature/***: Individual feature development
- **hotfix/***: Critical bug fixes
- **docs/***: Documentation updates

## Release Process

1. Complete feature development in feature branch
2. Merge to develop branch
3. Test thoroughly in develop
4. Create release branch from develop
5. Final testing and version updates in release branch
6. Merge release to main
7. Tag release with semantic version
8. Deploy from main branch

## Version Tagging

```bash
# Create version tags for releases
git tag -a v1.0.0 -m "Release version 1.0.0 - Initial production release"
git tag -a v1.1.0 -m "Release version 1.1.0 - Add custom properties support"
git push origin --tags
```

This strategy ensures clean history and traceable changes for the Office diagnostic tool.