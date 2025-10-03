# GitHub Copilot Instructions for TN-OfficeDebug

This file provides GitHub Copilot with specific context and coding patterns for the TN-OfficeDebug project. It ensures consistent, secure, and maintainable code generation that follows established project patterns.

## Project Context

TN-OfficeDebug is a Microsoft Office Add-in diagnostic tool for Word, designed primarily for ### Office.js API Reliability
```javascript
// Office.js APIs can be inconsistent across hosts and environments
// Always test API availability and provide fallbacks
if (Office.context.manifest && Office.context.manifest.version) {
  // API is available
} else {
  // Graceful fallback with visible failure indicator
}

// CRITICAL LESSON: Office.context.manifest.version is NOT RELIABLE
// In practice, this API often returns undefined or is not available
// DO NOT attempt to read version from Office.context.manifest.version
// Use obvious failure indicators like 'x.x.x' to show when APIs fail
```upport engineers. It provides comprehensive diagnostic information through a secure, tabbed interface with domain-based access control.

## Critical Office.js API Patterns

### ✅ Safe Office.js Patterns
```javascript
// Always load properties explicitly
const properties = context.document.properties;
properties.load(['title', 'author', 'subject']);
await context.sync();

// Use proper error handling
try {
  const customProperties = context.document.properties.customProperties;
  customProperties.load(['items']);
  await context.sync();
  // Process properties here
} catch (error) {
  window.logDebug('Error loading custom properties', { error: error.message });
}

// Check API support before use
if (Office.context.requirements && Office.context.requirements.isSetSupported('WordApi', '1.3')) {
  // Use advanced API features
}
```

### ❌ Avoid These Patterns
```javascript
// These properties don't exist in Word API - will cause errors
context.document.saved
context.document.isDirty
Office.context.auth // Not available in Word add-ins

// Don't access properties without loading first
properties.title // Error: property not loaded
```

## UI Component Patterns

### Standard Info Row Creation
```javascript
// Use the established helper function
container.appendChild(createInfoRow('Label', 'Value'));
container.appendChild(createInfoRow('Long Property Name', 'Value', true)); // true for long labels

// For custom properties (use green styling)
const customRow = createInfoRow(prop.key, prop.value);
customRow.style.borderLeft = '4px solid #28a745';
container.appendChild(customRow);
```

### Section Separators
```javascript
// Create visual section separators
const separator = document.createElement('div');
separator.style.marginTop = '20px';
separator.style.marginBottom = '15px';
separator.style.fontWeight = 'bold';
separator.style.borderTop = '1px solid #ccc';
separator.style.paddingTop = '10px';
separator.textContent = 'Section Title';
container.appendChild(separator);
```

## Tab-Specific Guidelines

### Document Info Tab (`loadDocumentInfo`)
- Focus on document properties, statistics, and custom properties
- Always handle custom properties with green styling
- Include security information (document mode, content controls)
- Avoid complex protection APIs that may not be available

### Office Environment Tab (`loadOfficeInfo`)
- Essential Office version and license information
- API capability detection (WordApi versions)
- Network status for connectivity troubleshooting
- Keep information scannable for support engineers

### System Info Tab (`loadSystemInfo`)
- Browser and platform information
- Security context (HTTPS validation)
- Performance metrics (memory usage)
- Critical diagnostics for troubleshooting

### Add-in Info Tab (`loadAddinsInfo`)
- Developer and project information
- Clickable links to GitHub repository and issues
- Troubleshooting information and support resources
- Load performance and status indicators

## Security Patterns

### Domain-Based Access Control
```javascript
// Check user domain for debug tab access
function checkDebugTabAccess() {
  // Multiple detection methods with fallbacks
  // 1. Office.context.user (when available)
  // 2. Document author extraction
  // 3. URL-based detection (localhost for dev)
  // 4. Manual domain prompt
}

// Hide debug features for unauthorized users
function hideDebugTab() {
  const debugTab = document.querySelector('[data-tab="debug"]');
  const debugContent = document.getElementById('debug');
  if (debugTab) debugTab.style.display = 'none';
  if (debugContent) debugContent.style.display = 'none';
}
```

### Secure Context Validation
```javascript
// Always validate security context
const isSecure = window.isSecureContext;
const protocol = window.location.protocol;
container.appendChild(createInfoRow('Security Context', isSecure ? 'Secure (HTTPS)' : 'Insecure (HTTP)'));
```

## Error Handling Standards

### Comprehensive Error Boundaries
```javascript
// Wrap all Office API operations in try/catch
try {
  await Word.run(async (context) => {
    // Office API operations here
    const body = context.document.body;
    body.load(['text']);
    await context.sync();
    
    container.appendChild(createInfoRow('Body Text Length', body.text.length));
  });
} catch (error) {
  window.logDebug('Error in Word.run operation', { 
    error: error.message,
    stack: error.stack 
  });
  container.innerHTML = `<div class="error">Error loading information: ${error.message}</div>`;
}
```

### Graceful Degradation
```javascript
// Provide fallback values for unavailable data
const value = properties.title || 'Not available';
const diagnosticInfo = Office.context.diagnostics?.version || 'Unknown';
```

## Logging and Debugging

### Consistent Logging Pattern
```javascript
// Use window.logDebug throughout the application
window.logDebug('Operation context', {
  operation: 'loadCustomProperties',
  itemCount: items.length,
  timestamp: new Date().toISOString()
});

// Log errors with full context
window.logDebug('API Error occurred', {
  error: error.message,
  api: 'customProperties',
  context: 'document load'
});
```

## File Structure Awareness

### Import/Reference Patterns
```javascript
// Files are organized in structured directories
// /src - Source code (taskpane.html, taskpane.js, taskpane.css, commands.html)
// /config - Configuration (manifest.xml, web.config)
// /docs - Documentation

// When referencing files in HTML:
<link rel="stylesheet" type="text/css" href="taskpane.css" />
<script type="text/javascript" src="taskpane.js"></script>

// Manifest references include /src path:
<SourceLocation DefaultValue="https://domain.com/src/taskpane.html"/>
```

## Performance Considerations

### Efficient API Calls
```javascript
// Load multiple properties in single call
properties.load(['title', 'author', 'subject', 'creationDate', 'lastSaveTime']);
await context.sync(); // Single sync call

// Avoid repeated DOM queries
const container = document.getElementById('container-id'); // Query once
// Use container variable for all subsequent operations
```

### Memory Management
```javascript
// Display memory usage when available
if (window.performance && window.performance.memory) {
  const memory = window.performance.memory;
  const usedMB = Math.round(memory.usedJSHeapSize / 1048576);
  const limitMB = Math.round(memory.jsHeapSizeLimit / 1048576);
  container.appendChild(createInfoRow('Memory Usage', `${usedMB}MB / ${limitMB}MB`));
}
```

## Accessibility and UX

### Responsive Design Patterns
```css
/* Use card-based layout to prevent text overflow */
.info-row {
  display: flex;
  flex-direction: column;
  margin: 10px 0;
  padding: 12px;
  background: #f8f9fa;
  border-radius: 4px;
  border-left: 4px solid #007acc;
}

/* Ensure text wrapping for long content */
.info-value {
  word-wrap: break-word;
  overflow-wrap: break-word;
}
```

### Professional Appearance
```javascript
// Use consistent styling for customer support context
// Cards should be clean, scannable, and professional
// Avoid overwhelming users with too much information
// Group related information logically
```

## Git and Deployment Patterns

### Commit Message Standards
```bash
# Use conventional commits - keep messages concise and professional
feat(ui): add document bookmark detection
fix(api): resolve custom properties loading error  
docs: update troubleshooting guide
refactor(security): improve domain detection

# Avoid verbose descriptions and unnecessary details
# Focus on what changed, not why (unless critical context)
```

### Mandatory Git Workflow
```bash
# ALWAYS commit after making changes - NEVER leave work uncommitted
# Required workflow:
# 1. Make file changes
# 2. Provide summary of changes to user
# 3. Get confirmation before committing
# 4. MANDATORY: Increment version number in manifest.xml before commit
# 5. Execute git add and git commit
# 6. No exceptions - every change must be committed

# SEMANTIC VERSIONING STRATEGY (REQUIRED):
# Before EVERY commit, increment version in config/manifest.xml:
# - Major (x.0.0): Brand new components, breaking changes, major features
# - Minor (1.x.0): New features, functionality additions, significant enhancements
# - Patch (1.0.x): Bug fixes, cache busting, small improvements, code cleanup

# Office Cache Prevention:
# Version increments in manifest.xml automatically invalidate Office cache
# This is the ONLY cache prevention method - no external scripts needed

git add .
git commit -m "type(scope): description"
```

### Git Tool Consistency
```bash
# CRITICAL: Maintain consistent git tooling throughout each session
# Once a git workflow is established (run_in_terminal vs GitKraken), stick with it
# Do not switch git tools mid-session without explicit user request or clear justification
# Examples:
# - If session starts with run_in_terminal + git commands, continue using that
# - If session starts with GitKraken tools, continue using those
# - Switching confuses workflow and reduces user experience quality

# Preferred approach: Use run_in_terminal with standard git commands unless:
# 1. User specifically requests GitKraken
# 2. GitKraken-specific features are genuinely needed
# 3. Consistent with project's established workflow
```

### Agent Instruction Maintenance
```bash
# When new context or patterns emerge during development:
# 1. Update AGENTS.md and/or copilot-instructions.md with new learnings
# 2. Document additional context, patterns, or requirements
# 3. Update README.md when functionality changes
# 4. Commit instruction updates immediately
# 5. Ensure future AI sessions benefit from current session learnings

# Example:
git add AGENTS.md .github/copilot-instructions.md README.md
git commit -m "docs: capture new development pattern from session context"
```

### Automatic Instruction Update Triggers
```bash
# CRITICAL: Immediately update agent instructions when these occur:
# 1. User corrects an AI mistake or approach - capture the correct pattern
# 2. User provides guidance on workflow, tools, or methodology
# 3. User explains "why" something should be done differently
# 4. Discovery of new project patterns, constraints, or requirements
# 5. Resolution of issues that could affect future sessions

# Trigger phrases that should prompt instruction updates:
# - "You should have..."
# - "Why did you..."
# - "Don't do that, do this instead..."
# - "Going forward..."
# - "The correct way is..."

# Process: Recognize → Update Instructions → Commit → Continue
# Do NOT wait for user prompting - make this automatic
```

### Verification and Problem-Solving Standards
```javascript
// CRITICAL: Follow these standards for all AI assistance
// 1. Do not assume or guess - only use verified sources for information
// 2. Always show sources when providing factual information
// 3. Do not repeat yourself when solving issues - try different approaches
// 4. If a problem cannot be resolved, inform user clearly rather than continuing

// Example of proper source citation:
// Source: Microsoft Learn documentation - [specific URL]
// Source: netstat command output - [command result]
// Source: project file analysis - [file path and content]
```

### Configuration Updates
```javascript
// When updating manifest or configuration files:
// 1. Update URLs to match hosting structure (/src paths)
// 2. Increment version numbers appropriately  
// 3. Test in both development and production environments
// 4. Validate manifest with Office validation tools
```

## Version Management Best Practices

### Single Source of Truth for Versions
```javascript
// ✅ Read version dynamically from manifest
container.appendChild(createInfoRow('Version', Office.context.manifest?.version || 'x.x.x'));

// ❌ Never hardcode versions in multiple places
container.appendChild(createInfoRow('Version', '1.0.3')); // Creates dual maintenance
```

### Effective Fallback Strategies
```javascript
// ✅ Use obvious failure indicators
let version = 'x.x.x'; // Makes API failures immediately visible

// ❌ Hide problems with alternate hardcoded values
let version = '1.0.0'; // Masks when Office.context.manifest.version fails
```

### Office.js API Reliability
```javascript
// Office.js APIs can be inconsistent across hosts/environments
// Always test API availability before relying on them
if (Office.context.manifest && Office.context.manifest.version) {
  // API is available
} else {
  // Graceful fallback with visible failure indicator
}
```

## Testing Patterns

### Cross-Environment Testing
```javascript
// Always test in multiple environments:
// - Word Desktop (Windows/Mac)
// - Word Online
// - Different browser contexts
// - With and without network connectivity
// - Various document types and protection levels

// Test error scenarios:
// - Missing custom properties
// - Restricted document access
// - Network connectivity issues
// - Unsupported Office versions
```

### Office Cache Troubleshooting
```bash
# Office applications cache manifest locations - clear cache after moving files
# Remove Office add-in cache folder when getting 404 errors
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Force -ErrorAction SilentlyContinue

# Common scenario: manifest moved from root to /config/ directory
# Office still looks for old location until cache is cleared
```

### Office.js API Compatibility
```javascript
// Use reliable patterns - advanced APIs may not work across all environments
// ✅ Reliable: ShowTaskpane action in manifest
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>ButtonId1</TaskpaneId>
  <SourceLocation resid="Taskpane.Url"/>
</Action>

// ❌ Unreliable: ExecuteFunction with Office.addin APIs (requires SharedRuntime)
// May not work in all Office versions/environments
Office.addin.hide() // Requires SharedRuntime 1.1
Office.addin.showAsTaskpane() // Requires SharedRuntime 1.1

// Always test API availability and provide fallbacks
```

### Development Testing
```javascript
// Use test-standalone.html for UI testing outside Office
// Test domain restrictions with different user contexts
// Validate security features work as expected
// Ensure graceful degradation for missing APIs
```

## Code Style and Conventions

### Variable Naming
```javascript
// Use descriptive names for customer support context
const customPropertiesContainer = document.getElementById('custom-properties');
const officeVersionInfo = Office.context.diagnostics?.version;
const documentSecurityStatus = 'protected';

// Use camelCase for JavaScript
const loadDocumentProperties = () => { /* */ };
const createDiagnosticRow = (label, value) => { /* */ };
```

### Function Organization
```javascript
// Keep tab functions focused and maintainable
function loadDocumentInfo() {
  // Document-specific diagnostics only
}

function loadOfficeInfo() {
  // Office environment diagnostics only  
}

// Use helper functions for reusable patterns
function createInfoRow(label, value, longLabel = false) {
  // Consistent UI component creation
}
```

Remember: This project serves customer support engineers who need quick, reliable diagnostic information. Every code suggestion should prioritize clarity, reliability, and the support workflow over developer convenience.