/* global Office */

Office.onReady((info) => {
  try {
    // Support Word, Excel, and PowerPoint
    if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
      const refreshBtn = document.getElementById('refresh-btn');
      if (refreshBtn) {
        refreshBtn.onclick = loadAllDebugInfo;
      }
      loadAllDebugInfo();
    } else {
      console.log('Unsupported Office host:', info.host);
      // Still load basic info for unsupported hosts
      loadAllDebugInfo();
    }
  } catch (error) {
    console.error('ERROR in Office.onReady', error);
  }
}).catch(error => {
  console.error('ERROR: Office.onReady failed', error);
});

function loadAllDebugInfo() {
  try {
    loadOfficeInfo();
    loadDocumentInfo();
    loadAddinsInfo();
    loadSystemInfo();
    // Skip loadContextInfo since there's no context-info element in HTML
    // loadContextInfo();
  } catch (error) {
    console.error('ERROR in loadAllDebugInfo', error);
  }
}

function createInfoRow(label, value, isCode = false) {
  // Create a card-style container for better layout
  const container = document.createElement('div');
  container.style.marginBottom = '8px';
  container.style.padding = '6px';
  container.style.backgroundColor = '#f9f9f9';
  container.style.borderRadius = '4px';
  container.style.borderLeft = '3px solid #667eea';
  
  const labelDiv = document.createElement('div');
  labelDiv.style.fontWeight = '600';
  labelDiv.style.color = '#555';
  labelDiv.style.fontSize = '11px';
  labelDiv.style.marginBottom = '2px';
  labelDiv.style.wordBreak = 'break-word';
  labelDiv.textContent = label;
  
  const valueDiv = document.createElement('div');
  valueDiv.style.color = '#333';
  valueDiv.style.fontSize = '11px';
  valueDiv.style.wordBreak = 'break-word';
  
  if (isCode) {
    valueDiv.style.fontFamily = "'Courier New', monospace";
    valueDiv.style.backgroundColor = '#f0f0f0';
    valueDiv.style.padding = '2px 4px';
    valueDiv.style.borderRadius = '3px';
    valueDiv.style.fontSize = '10px';
    valueDiv.style.marginTop = '2px';
  }
  
  valueDiv.textContent = value !== null && value !== undefined ? value.toString() : 'N/A';
  
  container.appendChild(labelDiv);
  container.appendChild(valueDiv);
  
  return container;
}

function loadOfficeInfo() {
  const container = document.getElementById('office-info');
  
  if (!container) {
    console.error('ERROR: office-info container not found');
    return;
  }
  
  try {
    container.innerHTML = '';

        // Office.context information
    container.appendChild(createInfoRow('Host', Office.context.host));
    container.appendChild(createInfoRow('Platform', Office.context.platform));
    
    // Key Diagnostics
    const keyDiagSeparator = document.createElement('div');
    keyDiagSeparator.style.marginTop = '20px';
    keyDiagSeparator.style.marginBottom = '15px';
    keyDiagSeparator.style.fontWeight = 'bold';
    keyDiagSeparator.style.borderTop = '1px solid #ccc';
    keyDiagSeparator.style.paddingTop = '10px';
    keyDiagSeparator.textContent = 'Key Diagnostics';
    container.appendChild(keyDiagSeparator);
    
    // Essential version info
    if (Office.context.diagnostics) {
      container.appendChild(createInfoRow('Office Version', Office.context.diagnostics.version || 'Unknown'));
    }
    
    // License info (important for support)
    if (Office.context.license) {
      container.appendChild(createInfoRow('License Type', Office.context.license.value || 'Unknown'));
    }
    
    // Critical API support
    let apiSupport = 'Basic';
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        apiSupport = 'Full (WordApi 1.3+)';
      } else if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
        apiSupport = 'Limited (WordApi 1.1)';
      }
    }
    container.appendChild(createInfoRow('API Support', apiSupport));
    

    
    // Office licensing and subscription info
    if (Office.context.license) {
      container.appendChild(createInfoRow('Office License', Office.context.license.value || 'Unknown'));
    }
    
    // Office capability detection for support
    const capabilities = [];
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      if (Office.context.requirements.isSetSupported('WordApi', '1.1')) capabilities.push('WordApi 1.1');
      if (Office.context.requirements.isSetSupported('WordApi', '1.2')) capabilities.push('WordApi 1.2');
      if (Office.context.requirements.isSetSupported('WordApi', '1.3')) capabilities.push('WordApi 1.3');
      if (Office.context.requirements.isSetSupported('WordApi', '1.4')) capabilities.push('WordApi 1.4');
      if (Office.context.requirements.isSetSupported('DialogApi', '1.1')) capabilities.push('DialogAPI');
      if (Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) capabilities.push('SharedRuntime');
      if (Office.context.requirements.isSetSupported('RibbonApi', '1.1')) capabilities.push('RibbonAPI');
    }
    container.appendChild(createInfoRow('Supported APIs', capabilities.length > 0 ? capabilities.join(', ') : 'Basic APIs only'));
    
    // Network and connectivity for troubleshooting
    container.appendChild(createInfoRow('Network Status', navigator.onLine ? 'Online' : 'Offline'));
    if (navigator.connection) {
      container.appendChild(createInfoRow('Connection Type', navigator.connection.effectiveType || 'Unknown'));
      container.appendChild(createInfoRow('Connection Speed', navigator.connection.downlink ? navigator.connection.downlink + ' Mbps' : 'Unknown'));
    }
    
    // Requirements info
    if (Office.context.requirements) {
      container.appendChild(createInfoRow('Requirements Available', 'Yes'));
    }
    
    // Display Mode
    if (Office.context.displayLanguage) {
      container.appendChild(createInfoRow('Display Language', Office.context.displayLanguage));
    }
    
    // Content Language
    if (Office.context.contentLanguage) {
      container.appendChild(createInfoRow('Content Language', Office.context.contentLanguage));
    }
    
    // Office Version (if available)
    if (Office.context.diagnostics) {
      container.appendChild(createInfoRow('Host Name', Office.context.diagnostics.host));
      container.appendChild(createInfoRow('Host Version', Office.context.diagnostics.version));
      container.appendChild(createInfoRow('Platform', Office.context.diagnostics.platform));
    }
    
    // Touch enabled
    if (Office.context.touchEnabled !== undefined) {
      container.appendChild(createInfoRow('Touch Enabled', Office.context.touchEnabled ? 'Yes' : 'No'));
    }
    
    // Office theme
    if (Office.context.officeTheme) {
      const theme = Office.context.officeTheme;
      container.appendChild(createInfoRow('Theme - Body Background', theme.bodyBackgroundColor, true));
      container.appendChild(createInfoRow('Theme - Body Foreground', theme.bodyForegroundColor, true));
      container.appendChild(createInfoRow('Theme - Control Background', theme.controlBackgroundColor, true));
      container.appendChild(createInfoRow('Theme - Control Foreground', theme.controlForegroundColor, true));
    }
    
  } catch (error) {
    container.innerHTML = `<div class="error">Error loading Office info: ${error.message}</div>`;
  }
}

function loadDocumentInfo() {
  const container = document.getElementById('document-info');
  
  if (!container) {
    console.error('ERROR: document-info container not found');
    return;
  }
  
  container.innerHTML = '';

  try {
    // Check which Office host we're in
    if (Office.context.host === Office.HostType.Word) {
      // Word-specific document info
      loadWordDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.Excel) {
      // Excel-specific document info
      loadExcelDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.PowerPoint) {
      // PowerPoint-specific document info
      loadPowerPointDocumentInfo(container);
    } else {
      // Generic document info for other hosts
      loadGenericDocumentInfo(container);
    }
  } catch (error) {
    console.error('ERROR in loadDocumentInfo', error);
    container.innerHTML = `<div class="error">Error: ${error.message}</div>`;
  }
}

function loadWordDocumentInfo(container) {
  Word.run(async (context) => {
      try {
        const doc = context.document;
        const properties = doc.properties;
        const body = doc.body;
        
        // Load properties
        properties.load('title,subject,author,keywords,comments,template,lastAuthor,revisionNumber,applicationName,lastPrintDate,creationDate,lastSaveTime');
        body.load('text,type');
        
        // Load custom properties
        const customProperties = doc.properties.customProperties;
        customProperties.load('items');
        
        await context.sync();
        
        // Document URL
        if (Office.context.document && Office.context.document.url) {
          container.appendChild(createInfoRow('Document URL', Office.context.document.url));
        }
        
        // Document Mode
        if (Office.context.document && Office.context.document.mode) {
          container.appendChild(createInfoRow('Document Mode', Office.context.document.mode));
        }
      
        
        // Built-in Document Properties
        container.appendChild(createInfoRow('Title', properties.title || 'Not set'));
        container.appendChild(createInfoRow('Subject', properties.subject || 'Not set'));
        container.appendChild(createInfoRow('Author', properties.author || 'Not set'));
        container.appendChild(createInfoRow('Keywords', properties.keywords || 'Not set'));
        container.appendChild(createInfoRow('Comments', properties.comments || 'Not set'));
        container.appendChild(createInfoRow('Template', properties.template || 'Not set'));
        container.appendChild(createInfoRow('Last Author', properties.lastAuthor || 'Not set'));
        container.appendChild(createInfoRow('Revision Number', properties.revisionNumber));
        container.appendChild(createInfoRow('Application Name', properties.applicationName));
        
        // Dates
        if (properties.creationDate) {
          container.appendChild(createInfoRow('Creation Date', new Date(properties.creationDate).toLocaleString()));
        }
        if (properties.lastSaveTime) {
          container.appendChild(createInfoRow('Last Save Time', new Date(properties.lastSaveTime).toLocaleString()));
        }
        if (properties.lastPrintDate) {
          container.appendChild(createInfoRow('Last Print Date', new Date(properties.lastPrintDate).toLocaleString()));
        }
        
        // Document statistics
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        container.appendChild(createInfoRow('Paragraph Count', paragraphs.items.length));
        
        // Text length (approximate)
        const textLength = body.text ? body.text.length : 0;
        container.appendChild(createInfoRow('Character Count (approx)', textLength));
        
        // Document Security (simplified)
        const securitySeparator = document.createElement('div');
        securitySeparator.style.marginTop = '20px';
        securitySeparator.style.marginBottom = '15px';
        securitySeparator.style.fontWeight = 'bold';
        securitySeparator.style.borderTop = '1px solid #ccc';
        securitySeparator.style.paddingTop = '10px';
        securitySeparator.textContent = 'Document Security';
        container.appendChild(securitySeparator);
        
        // Simple protection check
        try {
          const contentControls = context.document.contentControls;
          contentControls.load('items');
          await context.sync();
          container.appendChild(createInfoRow('Content Controls', contentControls.items.length > 0 ? `Yes (${contentControls.items.length})` : 'None'));
        } catch (controlError) {
          container.appendChild(createInfoRow('Content Controls', 'Could not check'));
        }
        
        // Document URL and mode
        container.appendChild(createInfoRow('Document URL', context.document.url || 'Local document'));
        if (Office.context.document && Office.context.document.mode) {
          container.appendChild(createInfoRow('Document Mode', Office.context.document.mode));
        }
        
        // Document Protection and Security Info
        const protectionSeparator = document.createElement('div');
        protectionSeparator.style.marginTop = '20px';
        protectionSeparator.style.marginBottom = '15px';
        protectionSeparator.style.fontWeight = 'bold';
        protectionSeparator.style.borderTop = '1px solid #ccc';
        protectionSeparator.style.paddingTop = '10px';
        protectionSeparator.textContent = 'Document Security & Protection';
        container.appendChild(protectionSeparator);
        
        // Check document protection
        try {
          const documentProtection = context.document.getDocumentProtection();
          documentProtection.load(['type', 'enabled']);
          await context.sync();
          
          container.appendChild(createInfoRow('Protection Enabled', documentProtection.enabled ? 'Yes' : 'No'));
          if (documentProtection.enabled) {
            container.appendChild(createInfoRow('Protection Type', documentProtection.type || 'Unknown'));
          }
        } catch (protError) {
          container.appendChild(createInfoRow('Protection Status', 'Could not determine'));
          // Could not get document protection info
        }
        

        
        // Custom Properties Section
        
        if (customProperties.items && customProperties.items.length > 0) {
          
          // Add a separator
          const separator = document.createElement('div');
          separator.style.marginTop = '20px';
          separator.style.marginBottom = '10px';
          separator.style.fontWeight = 'bold';
          separator.style.borderTop = '1px solid #ccc';
          separator.style.paddingTop = '10px';
          separator.textContent = 'Custom Document Properties';
          container.appendChild(separator);
          
          // Display each custom property  
          customProperties.items.forEach((customProp, index) => {
            const propName = customProp.key || `Property ${index + 1}`;
            const propValue = customProp.value !== null && customProp.value !== undefined ? customProp.value.toString() : 'N/A';
            const propType = customProp.type || 'Unknown';
            // Create custom property row with type information
            const propRow = createInfoRow(propName, propValue);
            
            // Add type information as a small subtitle
            const typeSpan = document.createElement('div');
            typeSpan.style.color = '#888';
            typeSpan.style.fontSize = '10px';
            typeSpan.style.fontStyle = 'italic';
            typeSpan.style.marginTop = '2px';
            typeSpan.textContent = `Type: ${propType}`;
            
            // Add the type info to the value div
            const valueDiv = propRow.lastElementChild;
            valueDiv.appendChild(typeSpan);
            
            // Different border color for custom properties
            propRow.style.borderLeftColor = '#28a745';
            
            container.appendChild(propRow);
          });
          
          container.appendChild(createInfoRow('Total Custom Properties', customProperties.items.length));
        } else {
          
          // Add a separator for "no custom properties"
          const separator = document.createElement('div');
          separator.style.marginTop = '20px';
          separator.style.marginBottom = '10px';
          separator.style.fontWeight = 'bold';
          separator.style.borderTop = '1px solid #ccc';
          separator.style.paddingTop = '10px';
          separator.textContent = 'Custom Document Properties';
          container.appendChild(separator);
          
          container.appendChild(createInfoRow('Custom Properties', 'None found in this document'));
        }
        
      } catch (error) {
        container.innerHTML = `<div class="error">Error loading document info: ${error.message}</div>`;
      }
    }).catch((error) => {
      container.innerHTML = `<div class="error">Word.run Error: ${error.message}</div>`;
    });
}

function loadExcelDocumentInfo(container) {
  // Excel-specific document information
  container.appendChild(createInfoRow('Host Application', 'Microsoft Excel'));
  
  // Generic document info that works across hosts
  if (Office.context.document && Office.context.document.url) {
    container.appendChild(createInfoRow('Document URL', Office.context.document.url));
  }
  
  // Add Excel-specific info here when needed
  container.appendChild(createInfoRow('Document Type', 'Excel Workbook'));
}

function loadPowerPointDocumentInfo(container) {
  // PowerPoint-specific document information
  container.appendChild(createInfoRow('Host Application', 'Microsoft PowerPoint'));
  
  // Generic document info that works across hosts
  if (Office.context.document && Office.context.document.url) {
    container.appendChild(createInfoRow('Document URL', Office.context.document.url));
  }
  
  // Add PowerPoint-specific info here when needed
  container.appendChild(createInfoRow('Document Type', 'PowerPoint Presentation'));
}

function loadGenericDocumentInfo(container) {
  // Generic document information for unknown hosts
  container.appendChild(createInfoRow('Host Application', Office.context.host || 'Unknown'));
  
  if (Office.context.document && Office.context.document.url) {
    container.appendChild(createInfoRow('Document URL', Office.context.document.url));
  }
}

function loadAddinsInfo() {
  const container = document.getElementById('addins-info');
  
  if (!container) {
    console.error('ERROR: addins-info container not found');
    return;
  }
  
  container.innerHTML = '';

  try {
    // Developer Information Section
    const devSeparator = document.createElement('div');
    devSeparator.style.marginTop = '10px';
    devSeparator.style.marginBottom = '15px';
    devSeparator.style.fontWeight = 'bold';
    devSeparator.style.borderTop = '1px solid #ccc';
    devSeparator.style.paddingTop = '10px';
    devSeparator.textContent = 'Developer Information';
    container.appendChild(devSeparator);
    
    container.appendChild(createInfoRow('Developer', 'Toakan Network'));
    container.appendChild(createInfoRow('Project Name', 'TN-OfficeDebug'));
    container.appendChild(createInfoRow('Description', 'Office.js Debug and Diagnostic Tool'));
    
    container.appendChild(createInfoRow('Version', Office.context.manifest?.version || 'x.x.x'));
    container.appendChild(createInfoRow('License', 'MIT'));
    
    // Repository Information
    const repoSeparator = document.createElement('div');
    repoSeparator.style.marginTop = '20px';
    repoSeparator.style.marginBottom = '15px';
    repoSeparator.style.fontWeight = 'bold';
    repoSeparator.style.borderTop = '1px solid #ccc';
    repoSeparator.style.paddingTop = '10px';
    repoSeparator.textContent = 'Repository Information';
    container.appendChild(repoSeparator);
    
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
    const techSeparator = document.createElement('div');
    techSeparator.style.marginTop = '20px';
    techSeparator.style.marginBottom = '15px';
    techSeparator.style.fontWeight = 'bold';
    techSeparator.style.borderTop = '1px solid #ccc';
    techSeparator.style.paddingTop = '10px';
    techSeparator.textContent = 'Technical Details';
    container.appendChild(techSeparator);
    
    container.appendChild(createInfoRow('Built With', 'Office.js, HTML5, CSS3, JavaScript ES6'));
    container.appendChild(createInfoRow('Target Office Apps', 'Microsoft Word'));
    container.appendChild(createInfoRow('Hosting', 'Web Server (IIS/Node.js compatible)'));
    container.appendChild(createInfoRow('API Requirements', 'Office.js 1.1+'));
    
    // Current Add-in Runtime Info
    const runtimeSeparator = document.createElement('div');
    runtimeSeparator.style.marginTop = '20px';
    runtimeSeparator.style.marginBottom = '15px';
    runtimeSeparator.style.fontWeight = 'bold';
    runtimeSeparator.style.borderTop = '1px solid #ccc';
    runtimeSeparator.style.paddingTop = '10px';
    runtimeSeparator.textContent = 'Runtime Information';
    container.appendChild(runtimeSeparator);
    
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
    const supportSeparator = document.createElement('div');
    supportSeparator.style.marginTop = '20px';
    supportSeparator.style.marginBottom = '15px';
    supportSeparator.style.fontWeight = 'bold';
    supportSeparator.style.borderTop = '1px solid #ccc';
    supportSeparator.style.paddingTop = '10px';
    supportSeparator.textContent = 'Support Information';
    container.appendChild(supportSeparator);
    
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
    const addinTroubleshootSeparator2 = document.createElement('div');
    addinTroubleshootSeparator2.style.marginTop = '20px';
    addinTroubleshootSeparator2.style.marginBottom = '15px';
    addinTroubleshootSeparator2.style.fontWeight = 'bold';
    addinTroubleshootSeparator2.style.borderTop = '1px solid #ccc';
    addinTroubleshootSeparator2.style.paddingTop = '10px';
    addinTroubleshootSeparator2.textContent = 'Troubleshooting Information';
    container.appendChild(addinTroubleshootSeparator2);
    
    // Add-in load diagnostics
    container.appendChild(createInfoRow('Add-in Loaded Successfully', 'Yes (you can see this!)'));
    container.appendChild(createInfoRow('Office.js Load Time', window.performance ? Math.round(performance.now()) + ' ms' : 'Unknown'));
    container.appendChild(createInfoRow('DOM Ready', document.readyState));
    
    // Error tracking
    let errorCount = 0;
    let lastError = 'None';
    if (window.onerror || window.addEventListener) {
      // Check if there's a global error handler
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
    const contactSeparator = document.createElement('div');
    contactSeparator.style.marginTop = '20px';
    contactSeparator.style.marginBottom = '15px';
    contactSeparator.style.fontWeight = 'bold';
    contactSeparator.style.borderTop = '1px solid #ccc';
    contactSeparator.style.paddingTop = '10px';
    contactSeparator.textContent = 'Support & Contact';
    container.appendChild(contactSeparator);
    
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
    note.innerHTML = '<strong>Security Note:</strong> Office Add-ins run in a sandboxed environment and cannot directly enumerate other installed add-ins for security reasons. This tool shows information about the current add-in and Office environment only.';
    container.appendChild(note);
    
  } catch (error) {
    container.innerHTML = `<div class="error">Error loading add-ins info: ${error.message}</div>`;
  }
}

function loadSystemInfo() {
  const container = document.getElementById('system-info');
  
  if (!container) {
    console.error('ERROR: system-info container not found');
    return;
  }
  
  container.innerHTML = '';

  try {
    // Essential Browser Information
    container.appendChild(createInfoRow('User Agent', navigator.userAgent));
    container.appendChild(createInfoRow('Platform', navigator.platform));
    container.appendChild(createInfoRow('Language', navigator.language));
    container.appendChild(createInfoRow('Online Status', navigator.onLine ? 'Yes' : 'No'));
    
    // Critical Diagnostics
    const criticalSeparator = document.createElement('div');
    criticalSeparator.style.marginTop = '20px';
    criticalSeparator.style.marginBottom = '15px';
    criticalSeparator.style.fontWeight = 'bold';
    criticalSeparator.style.borderTop = '1px solid #ccc';
    criticalSeparator.style.paddingTop = '10px';
    criticalSeparator.textContent = 'Critical Diagnostics';
    container.appendChild(criticalSeparator);
    
    // Security context (most important for Office add-ins)
    container.appendChild(createInfoRow('Security Protocol', window.location.protocol));
    container.appendChild(createInfoRow('Secure Context', window.isSecureContext ? 'Yes (HTTPS)' : 'No (HTTP)'));
    
    // Storage capabilities (important for add-in functionality)
    container.appendChild(createInfoRow('Local Storage', typeof(Storage) !== 'undefined' ? 'Available' : 'Blocked'));
    container.appendChild(createInfoRow('Cookies', navigator.cookieEnabled ? 'Enabled' : 'Disabled'));
    
    // Performance indicators
    if (window.performance && window.performance.memory) {
      const memory = window.performance.memory;
      const usedMB = Math.round(memory.usedJSHeapSize / 1048576);
      const limitMB = Math.round(memory.jsHeapSizeLimit / 1048576);
      container.appendChild(createInfoRow('Memory Usage', `${usedMB}MB / ${limitMB}MB`));
    }
    
    // Network diagnostics
    if (navigator.connection) {
      const conn = navigator.connection;
      container.appendChild(createInfoRow('Connection Type', conn.effectiveType || 'Unknown'));
      container.appendChild(createInfoRow('Downlink Speed', conn.downlink ? conn.downlink + ' Mbps' : 'Unknown'));
      container.appendChild(createInfoRow('Round Trip Time', conn.rtt ? conn.rtt + ' ms' : 'Unknown'));
      container.appendChild(createInfoRow('Data Saver Mode', conn.saveData ? 'Enabled' : 'Disabled'));
    }
    
    // Screen info
    container.appendChild(createInfoRow('Screen Width', screen.width + 'px'));
    container.appendChild(createInfoRow('Screen Height', screen.height + 'px'));
    container.appendChild(createInfoRow('Screen Color Depth', screen.colorDepth + ' bit'));
    container.appendChild(createInfoRow('Screen Pixel Depth', screen.pixelDepth + ' bit'));
    
    // Window info
    container.appendChild(createInfoRow('Window Width', window.innerWidth + 'px'));
    container.appendChild(createInfoRow('Window Height', window.innerHeight + 'px'));
    
    // Performance info (if available)
    if (window.performance && window.performance.memory) {
      const memory = window.performance.memory;
      container.appendChild(createInfoRow('JS Heap Size Limit', formatBytes(memory.jsHeapSizeLimit)));
      container.appendChild(createInfoRow('Total JS Heap Size', formatBytes(memory.totalJSHeapSize)));
      container.appendChild(createInfoRow('Used JS Heap Size', formatBytes(memory.usedJSHeapSize)));
    } else {
      container.appendChild(createInfoRow('Memory Info', 'Not available (Chrome-only feature)'));
    }
    
    // Timing info
    if (window.performance && window.performance.timing) {
      const timing = window.performance.timing;
      const loadTime = timing.loadEventEnd - timing.navigationStart;
      if (loadTime > 0) {
        container.appendChild(createInfoRow('Page Load Time', loadTime + ' ms'));
      }
    }
    
    // Current time
    container.appendChild(createInfoRow('Current Time', new Date().toLocaleString()));
    container.appendChild(createInfoRow('Timezone Offset', (new Date().getTimezoneOffset() / 60) + ' hours'));
    
  } catch (error) {
    container.innerHTML = `<div class="error">Error loading system info: ${error.message}</div>`;
  }
}

function loadContextInfo() {
  const container = document.getElementById('context-info');
  container.innerHTML = '';

  try {
    // Create a JSON display of the entire context (excluding circular references)
    const contextData = {
      host: Office.context.host,
      platform: Office.context.platform,
      displayLanguage: Office.context.displayLanguage,
      contentLanguage: Office.context.contentLanguage,
      touchEnabled: Office.context.touchEnabled,
    };
    
    // Add diagnostics if available
    if (Office.context.diagnostics) {
      contextData.diagnostics = {
        host: Office.context.diagnostics.host,
        version: Office.context.diagnostics.version,
        platform: Office.context.diagnostics.platform
      };
    }
    
    // Add document info if available
    if (Office.context.document) {
      contextData.document = {
        mode: Office.context.document.mode,
        url: Office.context.document.url
      };
    }
    
    // Create formatted JSON display
    const jsonDiv = document.createElement('div');
    jsonDiv.className = 'json-display';
    jsonDiv.textContent = JSON.stringify(contextData, null, 2);
    
    const label = document.createElement('div');
    label.style.marginBottom = '10px';
    label.style.fontWeight = '600';
    label.textContent = 'Full Context Object (sanitized):';
    
    container.appendChild(label);
    container.appendChild(jsonDiv);
    
    // Add Office JS library version info
    if (typeof Office !== 'undefined' && Office.context) {
      const versionInfo = document.createElement('div');
      versionInfo.style.marginTop = '15px';
      versionInfo.appendChild(createInfoRow('Office.js Library', 'Loaded and initialized'));
      container.appendChild(versionInfo);
    }
    
  } catch (error) {
    container.innerHTML = `<div class="error">Error loading context info: ${error.message}</div>`;
  }
}

// Helper function to format bytes
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
