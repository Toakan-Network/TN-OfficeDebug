# TN-OfficeDebug - AI Agent Instructions

This document provides comprehensive context and instructions for AI agents working on the TN-OfficeDebug project. It captures the project's specific requirements, ar#### 5. Web Server Configuration
**Problem**: CORS issues and SSL conflicts with Node.js development
**Solution**: Migration to IIS with proper CORS configuration
**Prevention**: Use IIS for enterprise-ready hosting

#### 6. Office Manifest Caching
**Problem**: 404 errors when manifest file is moved to different directory location
**Solution**: Clear Office cache - Office applications cache manifest locations from previous sideloading
**Prevention**: Clear cache after moving manifest files or changing project structure

#### 7. Office.js API Compatibility Issues
**Problem**: Advanced APIs (Office.addin.hide, ExecuteFunction) don't work reliably across all Office environments
**Solution**: Use basic, guaranteed-to-work patterns like ShowTaskpane action instead of complex toggle logic
**Prevention**: Test API availability and have fallback to simpler approachesecture decisions, lessons learned, and development patterns established during the initial development phase.

## 📋 Project Context

### Project Overview
**TN-OfficeDebug** is a comprehensive diagnostic and debugging tool for Microsoft Office Add-ins, specifically designed for Word integration with customer support capabilities. The tool provides detailed diagnostic information about the Office environment, document properties, system configuration, and add-in status.

### Primary Use Case
- **Customer Support Engineering**: Diagnostic tool for troubleshooting Office add-in issues
- **Development & Testing**: Environment validation and API compatibility checking
- **Enterprise Deployment**: Secure diagnostic capabilities with domain-based access control

### Key Stakeholders
- **Customer Support Engineers**: Primary users needing quick, scannable diagnostic information
- **Enterprise IT**: Deployment and security configuration
- **Developers**: Add-in development and troubleshooting

## 🏗️ Architecture & Technical Stack

### Core Technologies
```
- Office.js API (1.1+ with Word-specific extensions)
- Vanilla JavaScript (ES6+)
- HTML5 with semantic structure
- CSS3 with responsive design
- IIS/Node.js hosting compatibility
```

### Project Structure
```
/src          - Source code (taskpane.html, taskpane.js, taskpane.css, commands.html)
/config       - Configuration (manifest.xml, web.config)
/docs         - Documentation (instructions.md, USAGE.md)
/assets       - Static resources (currently empty)
/.github      - GitHub workflows and templates
```

### Key Components
1. **Tabbed Interface**: Document Info, Office Environment, System Info, Add-in Info, Debug (restricted)
2. **Custom Properties Engine**: Full CRUD support for Word document custom properties
3. **Security System**: Domain-based access control (@bighand.services)
4. **Card-based Layout**: Prevents text overflow and improves readability
5. **Comprehensive Logging**: Real-time debug system with window.logDebug()

## 🔑 Critical Development Guidelines

### Office.js API Compatibility Rules
**NEVER use these properties** - they don't exist in Word API:
```javascript
// ❌ These will cause errors:
context.document.saved
context.document.isDirty
Office.context.auth (not available in Word)

// ✅ Use these patterns instead:
context.document.properties.load(['title', 'author'])
await context.sync()
```

### Error Handling Pattern
```javascript
try {
  // Office API calls here
  const properties = context.document.properties;
  properties.load(['propertyName']);
  await context.sync();
  
  container.appendChild(createInfoRow('Label', properties.propertyName));
} catch (error) {
  window.logDebug('Error context', { error: error.message });
  container.appendChild(createInfoRow('Label', 'Could not load'));
}
```

### UI Component Pattern
```javascript
// Use createInfoRow helper for consistency
container.appendChild(createInfoRow('Label', 'Value'));
container.appendChild(createInfoRow('Long Label Name', 'Value', true)); // true for long labels

// Create visual separators
const separator = document.createElement('div');
separator.style.marginTop = '20px';
separator.style.marginBottom = '15px';
separator.style.fontWeight = 'bold';
separator.style.borderTop = '1px solid #ccc';
separator.style.paddingTop = '10px';
separator.textContent = 'Section Title';
container.appendChild(separator);
```

## 🛡️ Security & Access Control

### Domain-Based Security
- **Debug tab visibility**: Restricted to @bighand.services domain users
- **Fallback methods**: URL-based detection, local development auto-allow, manual prompt
- **Production security**: HTTPS required, CORS properly configured

### Security Implementation Pattern
```javascript
// Multi-method user detection with graceful fallbacks
function checkDebugTabAccess() {
  // 1. Try Office.context.user (when available)
  // 2. Document author properties extraction
  // 3. URL-based hosting detection
  // 4. Local development auto-allow
  // 5. Manual domain prompt as final fallback
}
```

## 📊 Feature Requirements & User Experience

### Information Organization Principles
1. **Scannable Information**: Customer support needs quick visual scanning
2. **Logical Grouping**: Related information grouped in tabs and sections
3. **No Duplicates**: Each piece of information appears in its most logical location
4. **Progressive Disclosure**: Essential info first, detailed diagnostics in subsections

### Tab-Specific Guidelines

#### Document Info Tab
- **Core Properties**: title, author, dates, statistics
- **Custom Properties**: Highlighted with green styling for easy identification
- **Security Info**: Document mode, content controls, URL
- **Avoid**: Complex protection APIs that may not be available

#### Office Environment Tab
- **Key Diagnostics**: Version, license, API support level
- **Essential Only**: Host, platform, network status
- **API Capabilities**: WordApi version matrix, DialogAPI, SharedRuntime

#### System Info Tab
- **Critical Browser Info**: User agent, platform, security context
- **Performance Metrics**: Memory usage, screen resolution
- **Security Context**: HTTPS validation, storage availability

#### Add-in Info Tab
- **Developer Information**: Project details, repository links (clickable)
- **Support Resources**: GitHub issues, documentation links
- **Troubleshooting**: Load status, performance metrics, common issue detection

### UI/UX Requirements
- **Card-based Layout**: Prevents text overflow, improves readability
- **Consistent Styling**: Use helper functions and established patterns
- **Responsive Design**: Works in various Office environments
- **Professional Appearance**: Clean, organized, suitable for customer support

## 🔧 Development Patterns & Best Practices

### Function Organization
```javascript
// Each tab has dedicated load function
function loadDocumentInfo() { /* Document tab logic */ }
function loadOfficeInfo() { /* Office environment logic */ }
function loadSystemInfo() { /* System diagnostics logic */ }
function loadAddinsInfo() { /* Add-in information logic */ }

// Helper functions for consistency
function createInfoRow(label, value, longLabel = false) { /* UI helper */ }
function setupTabs() { /* Tab navigation */ }
```

### Logging & Debugging
```javascript
// Use consistent logging throughout
window.logDebug('Context description', { 
  relevant: data,
  error: error?.message 
});
```

### Git Workflow Requirements
- **Mandatory Commits**: ALWAYS git add and commit after each file change or set of related changes
- **Change Summary Required**: Always provide a summary of changes before creating commits
- **User Confirmation**: Get confirmation from user after explaining changes but before executing git commit
- **Conventional Commits**: feat, fix, docs, style, refactor, test, chore
- **Scopes**: ui, api, security, config, manifest
- **Concise Messages**: Keep commit messages professional and succinct, avoid unnecessary details
- **Clear Messages**: Include impact and reasoning in commit body only when essential
- **Commit Frequency**: Commit after each logical unit of work, never leave uncommitted changes
- **Change Documentation**: Each commit should represent a complete, testable change

### Agent Instruction Updates
- **Context Learning**: When additional context or guidance is provided during development sessions, update the AI agent instruction files (AGENTS.md and .github/copilot-instructions.md) to capture these learnings
- **Pattern Recognition**: Document new patterns, best practices, or requirements that emerge during development
- **Instruction Refinement**: Keep agent instructions current with evolving project needs and discovered workflows
- **Knowledge Preservation**: Ensure future AI assistance benefits from lessons learned in previous sessions
- **README Maintenance**: Update README.md when functionality changes, new features are added, or project structure is modified

### Verification and Problem-Solving Standards
- **No Assumptions**: Do not assume or guess - only use verified sources for information
- **Show Sources**: Always cite and show sources when providing factual information
- **No Repetition**: Do not repeat yourself when solving issues - try different approaches
- **Acknowledge Limitations**: If a problem occurs and cannot be resolved, inform the user clearly rather than continuing ineffective attempts

## 🚨 Critical Lessons Learned

### Major Issues Resolved

#### 1. Office.js API Compatibility
**Problem**: Properties like `context.document.saved` don't exist in Word API
**Solution**: Comprehensive API validation and proper error handling
**Prevention**: Always validate API availability before use

#### 2. Information Overload
**Problem**: Too much diagnostic information overwhelmed users
**Solution**: Strategic organization, elimination of duplicates, scannable layout
**Prevention**: Design for customer support workflow, not developer curiosity

#### 3. Domain Security Implementation
**Problem**: Office.context.auth not available in Word add-ins
**Solution**: Multi-method user detection with graceful fallbacks
**Prevention**: Test all Office.js APIs across different host applications

#### 4. Web Server Configuration
**Problem**: CORS issues and SSL conflicts with Node.js development
**Solution**: Migration to IIS with proper CORS configuration
**Prevention**: Use IIS for enterprise-ready hosting

#### 5. UI Text Overflow
**Problem**: Long property names caused layout issues
**Solution**: Card-based layout with proper text wrapping
**Prevention**: Test with edge cases (long property names, values)

### Technical Architecture Decisions

#### Why Card-Based Layout
- Prevents text overflow issues
- Improves visual hierarchy
- Better mobile/responsive behavior
- Professional appearance for customer support

#### Why Domain-Based Security
- Enterprise security requirement
- Prevents accidental exposure of debug features
- Maintains professional appearance for external users

#### Why Comprehensive Error Handling
- Office.js API availability varies by host and version
- Graceful degradation ensures partial functionality
- Better user experience when some features unavailable

## 🎯 Development Workflow Guidelines

### Before Adding New Features
1. **Identify appropriate tab** for the information
2. **Check for existing similar information** to avoid duplicates
3. **Validate Office.js API availability** in Word
4. **Test error scenarios** and implement proper fallbacks
5. **Consider customer support workflow** - is this information helpful?

### Code Quality Standards
- **Always use try/catch** around Office.js API calls
- **Load properties explicitly** before accessing them
- **Provide fallback values** for unavailable data
- **Use helper functions** for consistent UI components
- **Log debugging information** with proper context

### Testing Requirements
- **Test in Word Desktop** and Word Online
- **Verify error handling** with missing/restricted APIs
- **Check responsive layout** at different window sizes
- **Validate security restrictions** work correctly
- **Test with real documents** containing custom properties

## 🚀 Deployment & Configuration

### Manifest Configuration
```xml
<!-- Always update URLs to match hosting structure -->
<SourceLocation DefaultValue="https://domain.com/src/taskpane.html"/>
<bt:Url id="Commands.Url" DefaultValue="https://domain.com/src/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://domain.com/src/taskpane.html"/>
```

### Web Server Requirements
- **HTTPS required** for production Office Online integration
- **CORS headers** properly configured for cross-origin requests
- **IIS recommended** for enterprise deployment stability

### Security Configuration
- **Domain restrictions** implemented for debug features
- **Secure context validation** for sensitive operations
- **Proper error boundaries** to prevent information leakage

## 📚 Reference Documentation

### Essential Reading
- **Office.js API Documentation**: https://docs.microsoft.com/en-us/office/dev/add-ins/
- **Word JavaScript API**: https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview
- **Add-in Manifest**: https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests

### Project-Specific Documentation
- **Full Setup Guide**: `/docs/instructions.md`
- **Quick Reference**: `/docs/USAGE.md`
- **Project Structure**: `/PROJECT_STRUCTURE.md`
- **Git Workflow**: `/.github/COMMIT_STRATEGY.md`

## 🔮 Future Development Considerations

### Potential Enhancements
- **Additional Office Host Support**: Excel, PowerPoint integration
- **Advanced Document Analysis**: Bookmark detection, style analysis
- **Performance Monitoring**: Real-time performance metrics
- **Export Functionality**: Diagnostic report generation

### Architectural Improvements
- **Modular Tab System**: Plugin-based architecture for easier extension
- **Configuration Management**: User-customizable diagnostic sections
- **Caching Strategy**: Reduce API calls for static information

### Maintenance Areas
- **Office.js Version Updates**: Regular compatibility testing
- **Security Policy Updates**: Domain restriction management
- **UI/UX Refinements**: Based on customer support feedback

---

**Remember**: This project serves customer support engineers first. Every feature, UI decision, and architectural choice should prioritize their workflow and the diagnostic information they need to resolve Office add-in issues efficiently.

**Last Updated**: October 2025  
**Project Phase**: Production Ready v1.0