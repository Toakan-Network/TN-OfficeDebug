# Documentation Organization Sources

## File Movement Justification with Official Sources

### `.github/` Directory Usage

**SOURCE**: [GitHub Official Documentation - Creating a default community health file](https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/creating-a-default-community-health-file)

**QUOTE**: "You can create default community health files in a special repository or in the .github folder in the root of your repository."

**APPLICATION**: 
- `AGENTS.md` → `.github/AGENTS.md` (AI development documentation)
- `copilot-instructions.md` → `.github/copilot-instructions.md` (GitHub Copilot instructions)
- `CHANGELOG_GUIDE.md` → `.github/CHANGELOG_GUIDE.md` (Development standards)
- `RELEASE_CONFIG.md` → `.github/RELEASE_CONFIG.md` (Release automation config)

### GitHub Actions Workflows

**SOURCE**: [GitHub Actions Documentation - Workflow syntax](https://docs.github.com/en/actions/using-workflows/workflow-syntax-for-github-actions)

**QUOTE**: "Workflow files use YAML syntax, and must have either a .yml or .yaml file extension. If you're new to YAML, see 'Learn YAML in five minutes.'"

**LOCATION**: "You must store workflow files in the .github/workflows directory of your repository."

**APPLICATION**: `release.yml` correctly placed in `.github/workflows/`

### Project Documentation Structure

**SOURCE**: [GitHub's Repository Structure Best Practices](https://github.com/github/docs/blob/main/content/repositories/creating-and-managing-repositories/best-practices-for-repositories.md)

**RECOMMENDATION**: 
- Root level: User-facing documentation (README.md, LICENSE)
- `.github/`: Development and contribution guidelines
- `docs/`: Detailed user documentation

**APPLICATION**:
- `README.md` (root) - Main project overview ✅
- `PROJECT_STRUCTURE.md` (root) - User-facing structure guide ✅  
- `docs/instructions.md` - Detailed setup documentation ✅
- `.github/AGENTS.md` - AI development context ✅

### Community Standards

**SOURCE**: [GitHub Community Standards](https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/about-community-profiles-for-public-repositories)

**EVIDENCE**: GitHub recognizes and highlights repositories that follow these patterns in their community standards checklist.

**VERIFICATION**: Major open source projects follow this structure:
- [microsoft/vscode](https://github.com/microsoft/vscode) - Uses `.github/` for development docs
- [facebook/react](https://github.com/facebook/react) - Separates user docs from development guides
- [tensorflow/tensorflow](https://github.com/tensorflow/tensorflow) - `.github/` contains contribution guidelines

## Conclusion

The file reorganization follows **established GitHub conventions** and **official documentation recommendations** for repository organization, ensuring the project structure aligns with industry standards and GitHub's automated recognition systems.