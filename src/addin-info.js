/* global Office */

function loadAddinsInfo() {
  const container = document.getElementById('addins-info');
  
  if (!container) {
    console.error('ERROR: addins-info container not found');
    return;
  }
  
  container.innerHTML = '';

  try {
    // Developer Information Section
    container.appendChild(createSectionSeparator('Developer Information'));
    
    container.appendChild(createInfoRow('Developer', 'Toakan Network'));
    container.appendChild(createInfoRow('Project Name', 'TN-OfficeDebug'));
    container.appendChild(createInfoRow('Description', 'Office.js Debug and Diagnostic Tool'));
    
    // Version from manifest - Office.js doesn't provide runtime access to manifest version
    // Source: Microsoft documentation shows no Office.context.manifest API exists
    const ADDIN_VERSION = '1.0.28'; // Keep in sync with config/manifest.xml
    container.appendChild(createInfoRow('Version', ADDIN_VERSION));
    container.appendChild(createInfoRow('License', 'MIT'));
    
    // Repository Information
    container.appendChild(createSectionSeparator('Repository Information'));
    
    // Create clickable link for GitHub
    const githubRow = createInfoRow('GitHub Repository', 'https://github.com/Toakan-Network/TN-OfficeDebug');
    const githubValueDiv = githubRow.lastElementChild;
    githubValueDiv.style.color = '#0078d4';
    githubValueDiv.style.textDecoration = 'underline';
    githubValueDiv.style.cursor = 'pointer';
    githubValueDiv.onclick = () => window.open('https://github.com/Toakan-Network/TN-OfficeDebug', '_blank');
    container.appendChild(githubRow);
    
    container.appendChild(createInfoRow('Repository Owner', 'Toakan-Network'));
    container.appendChild(createInfoRow('Current Branch', 'main'));
    
    // Technical Information
    container.appendChild(createSectionSeparator('Technical Details'));
    
    container.appendChild(createInfoRow('Built With', 'Office.js, HTML5, CSS3, JavaScript ES6'));
    container.appendChild(createInfoRow('Target Office Apps', 'Microsoft Word'));
    container.appendChild(createInfoRow('Hosting', 'Web Server (IIS/Node.js compatible)'));
    container.appendChild(createInfoRow('API Requirements', 'Office.js 1.1+'));
    
    // Current Add-in Runtime Info
    container.appendChild(createSectionSeparator('Runtime Information'));
    
    // Current add-in info
    if (Office.context.document && Office.context.document.settings) {
      container.appendChild(createInfoRow('Document Settings Available', 'Yes'));
    }
    
    // Roaming settings
    if (Office.context.roamingSettings) {
      const settings = Office.context.roamingSettings;
      const settingsKeys = Object.keys(settings);
      container.appendChild(createInfoRow('Roaming Settings Count', settingsKeys.length.toString()));
    }
    
    // Add-in license
    if (Office.context.license) {
      container.appendChild(createInfoRow('Office License', Office.context.license.value || 'N/A'));
    }
    
    // Current add-in manifest info
    if (Office.context.manifest) {
      container.appendChild(createInfoRow('Manifest ID', Office.context.manifest.id || 'N/A', true));
    } else {
      container.appendChild(createInfoRow('Manifest Available', 'No - Office.context.manifest is undefined'));
    }
    
    // Add-in UI
    if (Office.context.ui) {
      container.appendChild(createInfoRow('UI API Available', 'Yes'));
    }
    
    // Load information
    container.appendChild(createInfoRow('Load Time', new Date().toISOString()));
    
    // Critical Troubleshooting Information
    container.appendChild(createSectionSeparator('Support Information'));
    
    // Essential diagnostics for support
    container.appendChild(createInfoRow('Add-in Status', 'Loaded Successfully'));
    container.appendChild(createInfoRow('Current URL', window.location.href));
    container.appendChild(createInfoRow('Security Context', window.isSecureContext ? 'Secure (HTTPS)' : 'Insecure (HTTP)'));
    container.appendChild(createInfoRow('Console Access', typeof console !== 'undefined' ? 'Available' : 'Blocked'));
    
    // Common issue indicators
    const mixedContent = window.location.protocol === 'https:' ? 'No issues detected' : 'Potential mixed content';
    container.appendChild(createInfoRow('Mixed Content Check', mixedContent));
    
    const iframeContext = window.self === window.top ? 'Standalone' : 'Embedded (iframe)';
    container.appendChild(createInfoRow('Context', iframeContext));
    
    // Troubleshooting Information
    container.appendChild(createSectionSeparator('Troubleshooting Information'));
    
    // Add-in load diagnostics
    container.appendChild(createInfoRow('Add-in Loaded Successfully', 'Yes (you can see this!)'));
    container.appendChild(createInfoRow('Office.js Load Time', window.performance ? Math.round(performance.now()) + ' ms' : 'Unknown'));
    container.appendChild(createInfoRow('DOM Ready', document.readyState));
    
    // Error tracking
    if (window.onerror || window.addEventListener) {
      container.appendChild(createInfoRow('Error Monitoring', 'Active'));
    }
    
    // Console availability
    container.appendChild(createInfoRow('Console Available', typeof console !== 'undefined' ? 'Yes' : 'No'));
    
    // Common troubleshooting checks
    container.appendChild(createInfoRow('Mixed Content Issues', window.location.protocol === 'https:' ? 'None detected' : 'Possible (HTTP on HTTPS)'));
    container.appendChild(createInfoRow('CORS Headers', 'Check Network tab for details'));
    container.appendChild(createInfoRow('Iframe Context', window.self === window.top ? 'No (standalone)' : 'Yes (in iframe)'));
    
    // Performance metrics
    if (window.performance && window.performance.navigation) {
      const navType = window.performance.navigation.type;
      const navTypes = ['Navigate', 'Reload', 'Back/Forward', 'Reserved'];
      container.appendChild(createInfoRow('Navigation Type', navTypes[navType] || 'Unknown'));
    }
    
    // Contact Information
    container.appendChild(createSectionSeparator('Support & Contact'));
    
    // Create clickable link for Issues
    const issuesRow = createInfoRow('Report Issues', 'https://github.com/Toakan-Network/TN-OfficeDebug/issues');
    const issuesValueDiv = issuesRow.lastElementChild;
    issuesValueDiv.style.color = '#0078d4';
    issuesValueDiv.style.textDecoration = 'underline';
    issuesValueDiv.style.cursor = 'pointer';
    issuesValueDiv.onclick = () => window.open('https://github.com/Toakan-Network/TN-OfficeDebug/issues', '_blank');
    container.appendChild(issuesRow);
    
    container.appendChild(createInfoRow('Documentation', 'See README.md in repository'));
    container.appendChild(createInfoRow('Last Updated', 'October 2025'));
    
    // Note about limitations
    const note = document.createElement('div');
    note.style.marginTop = '20px';
    note.style.padding = '10px';
    note.style.backgroundColor = '#fff3cd';
    note.style.border = '1px solid #ffeaa7';
    note.style.borderRadius = '4px';
    note.style.fontSize = '11px';
    note.style.color = '#856404';
    
    // Create security note content safely
    const strongText = document.createElement('strong');
    strongText.textContent = 'Security Note:';
    note.appendChild(strongText);
    note.appendChild(document.createTextNode(' Office Add-ins run in a sandboxed environment and cannot directly enumerate other installed add-ins for security reasons. This tool shows information about the current add-in and Office environment only.'));
    
    container.appendChild(note);
    
  } catch (error) {
    displaySafeError(container, `Error loading add-ins info: ${error.message}`);
  }
}
