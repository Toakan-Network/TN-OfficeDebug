/* global Office */

// Debug logging function
window.logDebug = function(message, data = null) {
  if (data) {
    console.log(`[DEBUG] ${message}:`, data);
  } else {
    console.log(`[DEBUG] ${message}`);
  }
};

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
    loadDMSInfo();
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

// Security helper function to safely display errors without XSS
function displaySafeError(container, errorMessage) {
  const errorDiv = document.createElement('div');
  errorDiv.className = 'error';
  errorDiv.textContent = errorMessage; // Use textContent to prevent XSS
  container.appendChild(errorDiv);
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
    displaySafeError(container, `Error loading Office info: ${error.message}`);
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
    // Debug host detection
    window.logDebug('loadDocumentInfo called', {
      host: Office.context.host,
      hostTypeWord: Office.HostType.Word,
      hostTypeExcel: Office.HostType.Excel,
      hostTypePowerPoint: Office.HostType.PowerPoint,
      isWord: Office.context.host === Office.HostType.Word,
      isExcel: Office.context.host === Office.HostType.Excel,
      isPowerPoint: Office.context.host === Office.HostType.PowerPoint
    });
    
    // Check which Office host we're in
    if (Office.context.host === Office.HostType.Word) {
      // Word-specific document info
      window.logDebug('Calling loadWordDocumentInfo');
      loadWordDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.Excel) {
      // Excel-specific document info
      window.logDebug('Calling loadExcelDocumentInfo');
      loadExcelDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.PowerPoint) {
      // PowerPoint-specific document info
      window.logDebug('Calling loadPowerPointDocumentInfo');
      loadPowerPointDocumentInfo(container);
    } else {
      // Generic document info for other hosts
      window.logDebug('Calling loadGenericDocumentInfo for unknown host');
      loadGenericDocumentInfo(container);
    }
  } catch (error) {
    console.error('ERROR in loadDocumentInfo', error);
    displaySafeError(container, `Error: ${error.message}`);
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
        displaySafeError(container, `Error loading document info: ${error.message}`);
      }
    }).catch((error) => {
      displaySafeError(container, `Word.run Error: ${error.message}`);
    });
}

function loadExcelDocumentInfo(container) {
  Excel.run(async (context) => {
    try {
      const workbook = context.workbook;
      const properties = workbook.properties;
      
      // Load built-in properties
      properties.load(['title', 'subject', 'author', 'keywords', 'comments', 'creationDate', 'lastAuthor']);
      
      // Load custom properties
      const customProperties = properties.custom;
      customProperties.load('items');
      
      await context.sync();
      
      // Host application info
      container.appendChild(createInfoRow('Host Application', 'Microsoft Excel'));
      container.appendChild(createInfoRow('Document Type', 'Excel Workbook'));
      
      // Document URL
      if (Office.context.document && Office.context.document.url) {
        container.appendChild(createInfoRow('Document URL', Office.context.document.url));
      }
      
      // Built-in Document Properties
      container.appendChild(createInfoRow('Title', properties.title || 'Not set'));
      container.appendChild(createInfoRow('Subject', properties.subject || 'Not set'));
      container.appendChild(createInfoRow('Author', properties.author || 'Not set'));
      container.appendChild(createInfoRow('Keywords', properties.keywords || 'Not set'));
      container.appendChild(createInfoRow('Comments', properties.comments || 'Not set'));
      container.appendChild(createInfoRow('Last Author', properties.lastAuthor || 'Not set'));
      
      if (properties.creationDate) {
        container.appendChild(createInfoRow('Creation Date', new Date(properties.creationDate).toLocaleString()));
      }
      
      // Custom Properties Section
      const separator = document.createElement('div');
      separator.style.marginTop = '20px';
      separator.style.marginBottom = '15px';
      separator.style.fontWeight = 'bold';
      separator.style.borderTop = '1px solid #ccc';
      separator.style.paddingTop = '10px';
      separator.textContent = 'Custom Properties';
      container.appendChild(separator);
      
      if (customProperties.items && customProperties.items.length > 0) {
        customProperties.items.forEach((customProp, index) => {
          try {
            const propRow = createInfoRow(customProp.key, customProp.value);
            // Different border color for custom properties (green)
            propRow.style.borderLeft = '4px solid #28a745';
            container.appendChild(propRow);
          } catch (propError) {
            window.logDebug('Error displaying custom property', { 
              index: index, 
              error: propError.message 
            });
          }
        });
        
        container.appendChild(createInfoRow('Total Custom Properties', customProperties.items.length));
      } else {
        const noPropsRow = createInfoRow('Custom Properties', 'None found in this workbook');
        container.appendChild(noPropsRow);
      }
      
    } catch (error) {
      window.logDebug('Error in loadExcelDocumentInfo', { error: error.message });
      displaySafeError(container, `Error loading Excel document information: ${error.message}`);
    }
  });
}

function loadPowerPointDocumentInfo(container) {
  window.logDebug('PowerPoint document info loading started', {
    host: Office.context.host,
    hostType: Office.HostType.PowerPoint,
    hostMatch: Office.context.host === Office.HostType.PowerPoint,
    requirementsAvailable: !!Office.context.requirements,
    powerPointRunAvailable: typeof PowerPoint !== 'undefined' && typeof PowerPoint.run === 'function'
  });
  
  // Add basic info first - this always works
  container.appendChild(createInfoRow('Host Application', 'Microsoft PowerPoint'));
  container.appendChild(createInfoRow('Document Type', 'PowerPoint Presentation'));
  
  // Use only reliable Office.context APIs - avoid PowerPoint.run due to GeneralException issues
  try {
    window.logDebug('Loading document info using Office.context APIs only');
    
    // Document URL and basic info - most reliable source
    if (Office.context.document && Office.context.document.url) {
      const url = Office.context.document.url;
      const fileName = url.split('/').pop().split('\\').pop();
      container.appendChild(createInfoRow('File Name', fileName || 'Unknown'));
      container.appendChild(createInfoRow('Document URL', url));
    }
    
    // Document mode - available in basic Office APIs
    if (Office.context.document && Office.context.document.mode !== undefined) {
      const mode = Office.context.document.mode === Office.DocumentMode.ReadOnly ? 'Read-Only' : 'Read-Write';
      container.appendChild(createInfoRow('Document Mode', mode));
    }
    
    // Try to get any additional document settings
    if (Office.context.document && Office.context.document.settings) {
      window.logDebug('Document settings available');
      // Settings API is available but may not contain useful document properties
    }
    
    // Try basic custom properties via Office.context - this is more reliable than PowerPoint.run
    if (Office.context.document && Office.context.document.customProperties) {
      window.logDebug('Attempting custom properties via Office.context');
      
      Office.context.document.customProperties.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          window.logDebug('Custom properties loaded via Office.context', { 
            count: result.value ? result.value.length : 0,
            properties: result.value
          });
          
          if (result.value && result.value.length > 0) {
            // Add separator for custom properties
            const separator = document.createElement('div');
            separator.style.marginTop = '20px';
            separator.style.marginBottom = '15px';
            separator.style.fontWeight = 'bold';
            separator.style.borderTop = '1px solid #ccc';
            separator.style.paddingTop = '10px';
            separator.textContent = 'Custom Properties';
            container.appendChild(separator);
            
            // Display each custom property
            result.value.forEach(prop => {
              const row = createInfoRow(prop.name, prop.value);
              row.style.borderLeft = '4px solid #28a745';
              container.appendChild(row);
            });
            
            container.appendChild(createInfoRow('Total Custom Properties', result.value.length));
          } else {
            container.appendChild(createInfoRow('Custom Properties', 'None found'));
          }
        } else {
          window.logDebug('Custom properties loading failed', { 
            error: result.error ? result.error.message : 'Unknown error',
            status: result.status
          });
          container.appendChild(createInfoRow('Custom Properties', 'Unable to load (API limitation)'));
        }
      });
    } else {
      window.logDebug('Office.context.document.customProperties not available');
      container.appendChild(createInfoRow('Custom Properties', 'Not supported in this PowerPoint version'));
    }
    
    // Add informational note about PowerPoint API limitations
    const infoNote = document.createElement('div');
    infoNote.style.marginTop = '20px';
    infoNote.style.padding = '12px';
    infoNote.style.backgroundColor = '#d1ecf1';
    infoNote.style.border = '1px solid #bee5eb';
    infoNote.style.borderRadius = '4px';
    infoNote.style.fontSize = '12px';
    infoNote.style.color = '#0c5460';
    infoNote.innerHTML = '<strong>PowerPoint API Information:</strong><br>' +
                         'PowerPoint provides limited document property access compared to Word and Excel. ' +
                         'Advanced features like detailed document statistics, content analysis, and ' +
                         'comprehensive metadata are not available through PowerPoint\'s Office.js APIs.';
    container.appendChild(infoNote);
    
    // Add a note about successful basic info loading
    container.appendChild(createInfoRow('Status', 'Basic document information loaded successfully'));
    
    window.logDebug('PowerPoint document info loading completed using basic APIs');
    
  } catch (error) {
    window.logDebug('Error loading PowerPoint document info via Office.context', { 
      error: error.message,
      stack: error.stack,
      name: error.name 
    });
    
    displaySafeError(container, `Error accessing document information: ${error.message}`);
    
    // Still add a basic status message
    const fallbackNote = document.createElement('div');
    fallbackNote.style.marginTop = '15px';
    fallbackNote.style.padding = '10px';
    fallbackNote.style.backgroundColor = '#f8d7da';
    fallbackNote.style.border = '1px solid #f5c6cb';
    fallbackNote.style.borderRadius = '4px';
    fallbackNote.style.fontSize = '11px';
    fallbackNote.style.color = '#721c24';
    fallbackNote.textContent = 'PowerPoint document information could not be loaded due to API restrictions.';
    container.appendChild(fallbackNote);
  }
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
    
    // Version from manifest - Office.js doesn't provide runtime access to manifest version
    // Source: Microsoft documentation shows no Office.context.manifest API exists
    const ADDIN_VERSION = '1.0.28'; // Keep in sync with config/manifest.xml
    container.appendChild(createInfoRow('Version', ADDIN_VERSION));
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
    displaySafeError(container, `Error loading system info: ${error.message}`);
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
    displaySafeError(container, `Error loading context info: ${error.message}`);
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

// DMS (Document Management System) Detection and Integration
function loadDMSInfo() {
  window.logDebug('DMS info loading started');
  const container = document.getElementById('dms-info');
  
  if (!container) {
    console.error('ERROR: dms-info container not found');
    return;
  }
  
  container.innerHTML = '';
  container.appendChild(createInfoRow('Status', 'Analyzing document...'));
  
  // Detect DMS based on custom properties
  detectDMS()
    .then(dmsInfo => {
      container.innerHTML = '';
      displayDMSInfo(container, dmsInfo);
    })
    .catch(error => {
      window.logDebug('Error in DMS detection', { error: error.message });
      displaySafeError(container, `Error detecting DMS: ${error.message}`);
    });
}

async function detectDMS() {
  window.logDebug('Starting DMS detection');
  
  return new Promise((resolve, reject) => {
    try {
      // Get custom properties to detect DMS
      if (Office.context.host === Office.HostType.Word) {
        detectDMSWord(resolve, reject);
      } else if (Office.context.host === Office.HostType.Excel) {
        detectDMSExcel(resolve, reject);
      } else if (Office.context.host === Office.HostType.PowerPoint) {
        detectDMSPowerPoint(resolve, reject);
      } else {
        resolve({
          type: 'unknown',
          detected: false,
          message: 'DMS detection not supported for this Office application'
        });
      }
    } catch (error) {
      reject(error);
    }
  });
}

function detectDMSWord(resolve, reject) {
  window.logDebug('Detecting DMS in Word document');
  
  // First check if Word API is available
  if (typeof Word === 'undefined' || typeof Word.run !== 'function') {
    window.logDebug('Word API not available, using fallback');
    detectDMSFallback(resolve, reject);
    return;
  }
  
  Word.run(async (context) => {
    try {
      window.logDebug('Inside Word.run context for DMS detection');
      
      // Use the exact same approach as loadWordDocumentInfo which works
      const doc = context.document;
      const properties = doc.properties;
      window.logDebug('Got document and properties objects', { 
        hasDoc: !!doc, 
        hasProperties: !!properties 
      });
      
      // Load custom properties using the exact same pattern as Document Info tab
      const customProperties = properties.customProperties;
      window.logDebug('Got custom properties object', { hasCustomProperties: !!customProperties });
      
      // Use 'items' string instead of array (match working code exactly)
      customProperties.load('items');
      window.logDebug('Custom properties load called, about to sync');
      
      await context.sync();
      window.logDebug('Context sync completed for DMS detection');
      
      window.logDebug('Custom properties items loaded', {
        hasItems: !!customProperties.items,
        itemCount: customProperties.items ? customProperties.items.length : 0,
        items: customProperties.items ? customProperties.items.map(item => ({
          key: item.key,
          value: item.value,
          type: item.type
        })) : null
      });
      
      const dmsInfo = analyzeDMSProperties(customProperties.items);
      window.logDebug('DMS analysis completed in Word', dmsInfo);
      resolve(dmsInfo);
      
    } catch (innerError) {
      window.logDebug('Error inside Word.run for DMS detection', {
        error: innerError.message,
        stack: innerError.stack,
        name: innerError.name
      });
      throw innerError;
    }
  }).catch(error => {
    window.logDebug('Word.run failed in DMS detection', { 
      error: error.message,
      name: error.name,
      stack: error.stack
    });
    // Fallback to Office.context
    window.logDebug('Falling back to Office.context for DMS detection');
    detectDMSFallback(resolve, reject);
  });
}

function detectDMSExcel(resolve, reject) {
  window.logDebug('Detecting DMS in Excel document');
  
  Excel.run(async (context) => {
    try {
      const properties = context.workbook.properties;
      const customProperties = properties.custom;
      customProperties.load(['items']);
      
      await context.sync();
      
      const dmsInfo = analyzeDMSProperties(customProperties.items);
      resolve(dmsInfo);
      
    } catch (error) {
      window.logDebug('Excel DMS detection failed', { error: error.message });
      detectDMSFallback(resolve, reject);
    }
  }).catch(error => {
    window.logDebug('Excel.run failed in DMS detection', { error: error.message });
    detectDMSFallback(resolve, reject);
  });
}

function detectDMSPowerPoint(resolve, reject) {
  window.logDebug('Detecting DMS in PowerPoint document');
  
  // PowerPoint has limited API support, use fallback approach
  detectDMSFallback(resolve, reject);
}

function detectDMSFallback(resolve, reject) {
  window.logDebug('Using DMS detection fallback - Office.context.document.customProperties API not available');
  
  try {
    window.logDebug('Office.context availability check', {
      hasOfficeContext: !!Office.context,
      hasDocument: !!(Office.context && Office.context.document),
      hasCustomProperties: !!(Office.context && Office.context.document && Office.context.document.customProperties),
      documentUrl: Office.context && Office.context.document ? Office.context.document.url : 'not available'
    });
    
    // Office.context.document.customProperties doesn't exist in this Office version
    // This is expected - Word custom properties should be accessed via Word.run API
    window.logDebug('Office.context.document.customProperties API not available (this is normal)');
    
    resolve({
      type: 'unknown',
      detected: false,
      message: 'Custom properties must be accessed via Word API - Office.context.document.customProperties not available',
      debugInfo: {
        reason: 'Office.context.document.customProperties API not supported',
        suggestion: 'Word.run should be used instead',
        hasOfficeContext: !!Office.context,
        hasDocument: !!(Office.context && Office.context.document),
        availableDocumentProperties: Office.context && Office.context.document ? Object.keys(Office.context.document) : null
      }
    });
    
  } catch (error) {
    window.logDebug('Exception in DMS fallback detection', {
      error: error.message,
      stack: error.stack,
      name: error.name
    });
    reject(error);
  }
}

function analyzeDMSProperties(customProperties) {
  window.logDebug('Analyzing custom properties for DMS signatures', { count: customProperties.length });
  
  let netDocumentsId = null;
  let iManageWork = null;
  let documentNumber = null;
  let allProperties = [];
  
  customProperties.forEach(prop => {
    const key = prop.key.toLowerCase();
    const value = prop.value;
    
    allProperties.push({ key: prop.key, value: value });
    
    // NetDocuments detection
    if (key === 'nddocumentid') {
      netDocumentsId = value;
      window.logDebug('NetDocuments ID found', { id: value });
    }
    
    // iManage detection
    if (key === 'isimanagework') {
      iManageWork = value;
      window.logDebug('iManage Work indicator found', { value: value });
    }
    
    if (key === 'documentnumber') {
      documentNumber = value;
      window.logDebug('Document Number found', { number: value });
    }
  });
  
  // Determine DMS type
  if (netDocumentsId) {
    return {
      type: 'netdocuments',
      detected: true,
      properties: {
        ndDocumentId: netDocumentsId
      },
      allProperties: allProperties,
      message: 'NetDocuments document detected'
    };
  }
  
  if (iManageWork && documentNumber) {
    return {
      type: 'imanage',
      detected: true,
      properties: {
        IsiManageWork: iManageWork,
        DocumentNumber: documentNumber
      },
      allProperties: allProperties,
      message: 'iManage document detected'
    };
  }
  
  return {
    type: 'unknown',
    detected: false,
    allProperties: allProperties,
    message: 'No DMS signatures found in document custom properties'
  };
}

function analyzeDMSPropertiesFallback(customProperties) {
  window.logDebug('Analyzing custom properties for DMS signatures (fallback)', { count: customProperties.length });
  
  let netDocumentsId = null;
  let iManageWork = null;
  let documentNumber = null;
  let allProperties = [];
  
  customProperties.forEach(prop => {
    const key = prop.name.toLowerCase();
    const value = prop.value;
    
    allProperties.push({ key: prop.name, value: value });
    
    // NetDocuments detection
    if (key === 'nddocumentid') {
      netDocumentsId = value;
      window.logDebug('NetDocuments ID found (fallback)', { id: value });
    }
    
    // iManage detection
    if (key === 'isimanagework') {
      iManageWork = value;
      window.logDebug('iManage Work indicator found (fallback)', { value: value });
    }
    
    if (key === 'documentnumber') {
      documentNumber = value;
      window.logDebug('Document Number found (fallback)', { number: value });
    }
  });
  
  // Determine DMS type
  if (netDocumentsId) {
    return {
      type: 'netdocuments',
      detected: true,
      properties: {
        ndDocumentId: netDocumentsId
      },
      allProperties: allProperties,
      message: 'NetDocuments document detected'
    };
  }
  
  if (iManageWork && documentNumber) {
    return {
      type: 'imanage',
      detected: true,
      properties: {
        IsiManageWork: iManageWork,
        DocumentNumber: documentNumber
      },
      allProperties: allProperties,
      message: 'iManage document detected'
    };
  }
  
  return {
    type: 'unknown',
    detected: false,
    allProperties: allProperties,
    message: 'No DMS signatures found in document custom properties'
  };
}

function displayDMSInfo(container, dmsInfo) {
  window.logDebug('Displaying DMS info', dmsInfo);
  
  // DMS Detection Results
  container.appendChild(createInfoRow('DMS Type', dmsInfo.detected ? dmsInfo.type.toUpperCase() : 'Not Detected'));
  container.appendChild(createInfoRow('Detection Status', dmsInfo.message));
  
  // Add debug information
  if (dmsInfo.debugInfo) {
    const debugSeparator = document.createElement('div');
    debugSeparator.style.marginTop = '20px';
    debugSeparator.style.marginBottom = '15px';
    debugSeparator.style.fontWeight = 'bold';
    debugSeparator.style.borderTop = '1px solid #ccc';
    debugSeparator.style.paddingTop = '10px';
    debugSeparator.textContent = 'Debug Information';
    container.appendChild(debugSeparator);
    
    Object.entries(dmsInfo.debugInfo).forEach(([key, value]) => {
      const row = createInfoRow(key, JSON.stringify(value));
      row.style.borderLeft = '4px solid #ff6b6b';
      container.appendChild(row);
    });
  }
  
  if (dmsInfo.detected) {
    // DMS-specific information
    const separator = document.createElement('div');
    separator.style.marginTop = '20px';
    separator.style.marginBottom = '15px';
    separator.style.fontWeight = 'bold';
    separator.style.borderTop = '1px solid #ccc';
    separator.style.paddingTop = '10px';
    separator.textContent = `${dmsInfo.type.toUpperCase()} Properties`;
    container.appendChild(separator);
    
    // Display DMS-specific properties
    Object.entries(dmsInfo.properties).forEach(([key, value]) => {
      const row = createInfoRow(key, value);
      row.style.borderLeft = '4px solid #007acc';
      container.appendChild(row);
    });
    
    // Add DMS actions section
    if (dmsInfo.type === 'netdocuments') {
      addNetDocumentsActions(container, dmsInfo.properties);
    } else if (dmsInfo.type === 'imanage') {
      addIManageActions(container, dmsInfo.properties);
    }
  }
  
  // All Custom Properties Section (for debugging)
  if (dmsInfo.allProperties && dmsInfo.allProperties.length > 0) {
    const allPropsSeparator = document.createElement('div');
    allPropsSeparator.style.marginTop = '20px';
    allPropsSeparator.style.marginBottom = '15px';
    allPropsSeparator.style.fontWeight = 'bold';
    allPropsSeparator.style.borderTop = '1px solid #ccc';
    allPropsSeparator.style.paddingTop = '10px';
    allPropsSeparator.textContent = 'All Custom Properties';
    container.appendChild(allPropsSeparator);
    
    dmsInfo.allProperties.forEach(prop => {
      const row = createInfoRow(prop.key, prop.value);
      row.style.borderLeft = '4px solid #28a745';
      container.appendChild(row);
    });
    
    container.appendChild(createInfoRow('Total Custom Properties', dmsInfo.allProperties.length));
  } else {
    // Show that no custom properties were found
    const noPropsSeparator = document.createElement('div');
    noPropsSeparator.style.marginTop = '20px';
    noPropsSeparator.style.marginBottom = '15px';
    noPropsSeparator.style.fontWeight = 'bold';
    noPropsSeparator.style.borderTop = '1px solid #ccc';
    noPropsSeparator.style.paddingTop = '10px';
    noPropsSeparator.textContent = 'Custom Properties Status';
    container.appendChild(noPropsSeparator);
    
    container.appendChild(createInfoRow('Custom Properties Found', '0'));
    container.appendChild(createInfoRow('Note', 'No custom properties detected - check console for detailed debugging info'));
  }
}

function addNetDocumentsActions(container, properties) {
  window.logDebug('Adding NetDocuments actions', properties);
  
  const actionsSeparator = document.createElement('div');
  actionsSeparator.style.marginTop = '20px';
  actionsSeparator.style.marginBottom = '15px';
  actionsSeparator.style.fontWeight = 'bold';
  actionsSeparator.style.borderTop = '1px solid #ccc';
  actionsSeparator.style.paddingTop = '10px';
  actionsSeparator.textContent = 'NetDocuments Actions';
  container.appendChild(actionsSeparator);
  
  // Test API Connection Button
  const testButton = document.createElement('button');
  testButton.textContent = 'Test NetDocuments API Connection';
  testButton.style.margin = '10px 0';
  testButton.style.padding = '8px 16px';
  testButton.style.backgroundColor = '#007acc';
  testButton.style.color = 'white';
  testButton.style.border = 'none';
  testButton.style.borderRadius = '4px';
  testButton.style.cursor = 'pointer';
  testButton.onclick = () => testNetDocumentsAPI(properties);
  container.appendChild(testButton);
  
  // Get Document Info Button
  const docInfoButton = document.createElement('button');
  docInfoButton.textContent = 'Get Document Information';
  docInfoButton.style.margin = '10px 0 10px 10px';
  docInfoButton.style.padding = '8px 16px';
  docInfoButton.style.backgroundColor = '#28a745';
  docInfoButton.style.color = 'white';
  docInfoButton.style.border = 'none';
  docInfoButton.style.borderRadius = '4px';
  docInfoButton.style.cursor = 'pointer';
  docInfoButton.onclick = () => getNetDocumentsInfo(properties);
  container.appendChild(docInfoButton);
  
  // API Results area
  const resultsArea = document.createElement('div');
  resultsArea.id = 'netdocs-results';
  resultsArea.style.marginTop = '15px';
  container.appendChild(resultsArea);
}

function addIManageActions(container, properties) {
  window.logDebug('Adding iManage actions', properties);
  
  const actionsSeparator = document.createElement('div');
  actionsSeparator.style.marginTop = '20px';
  actionsSeparator.style.marginBottom = '15px';
  actionsSeparator.style.fontWeight = 'bold';
  actionsSeparator.style.borderTop = '1px solid #ccc';
  actionsSeparator.style.paddingTop = '10px';
  actionsSeparator.textContent = 'iManage Actions';
  container.appendChild(actionsSeparator);
  
  const placeholder = document.createElement('div');
  placeholder.style.padding = '10px';
  placeholder.style.backgroundColor = '#f8f9fa';
  placeholder.style.border = '1px solid #dee2e6';
  placeholder.style.borderRadius = '4px';
  placeholder.style.fontStyle = 'italic';
  placeholder.textContent = 'iManage API integration will be implemented in future version';
  container.appendChild(placeholder);
}

// NetDocuments API Functions (Real REST API Implementation)
async function testNetDocumentsAPI(properties) {
  window.logDebug('Testing NetDocuments API connection', properties);
  
  const resultsArea = document.getElementById('netdocs-results');
  resultsArea.innerHTML = '<div style="color: #666; font-style: italic;">Testing API connection...</div>';
  
  try {
    // NetDocuments REST API Configuration
    const config = getNetDocumentsConfig();
    
    if (!config.isConfigured) {
      showNetDocumentsConfigPrompt(resultsArea, properties);
      return;
    }
    
    // Test API connectivity with a simple repository list call
    const testEndpoint = `${config.baseUrl}/v2/repository`;
    
    window.logDebug('Testing NetDocuments API endpoint', { endpoint: testEndpoint });
    
    const response = await fetch(testEndpoint, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    });
    
    window.logDebug('NetDocuments API response received', { 
      status: response.status,
      statusText: response.statusText,
      ok: response.ok
    });
    
    resultsArea.innerHTML = '';
    
    if (response.ok) {
      const data = await response.json();
      window.logDebug('NetDocuments API test successful', data);
      
      resultsArea.appendChild(createInfoRow('API Test Status', 'Success'));
      resultsArea.appendChild(createInfoRow('Response Status', `${response.status} ${response.statusText}`));
      resultsArea.appendChild(createInfoRow('Base URL', config.baseUrl));
      resultsArea.appendChild(createInfoRow('Test Endpoint', '/v2/repository'));
      resultsArea.appendChild(createInfoRow('Repositories Found', data.length || 'N/A'));
      resultsArea.appendChild(createInfoRow('Test Time', new Date().toISOString()));
      
      const successNote = document.createElement('div');
      successNote.style.marginTop = '10px';
      successNote.style.padding = '8px';
      successNote.style.backgroundColor = '#d4edda';
      successNote.style.border = '1px solid #c3e6cb';
      successNote.style.borderRadius = '4px';
      successNote.style.fontSize = '11px';
      successNote.style.color = '#155724';
      successNote.textContent = 'NetDocuments API connection successful! You can now retrieve document information.';
      resultsArea.appendChild(successNote);
      
    } else {
      window.logDebug('NetDocuments API test failed', { 
        status: response.status,
        statusText: response.statusText
      });
      
      const errorText = await response.text();
      window.logDebug('NetDocuments API error response', errorText);
      
      resultsArea.appendChild(createInfoRow('API Test Status', 'Failed'));
      resultsArea.appendChild(createInfoRow('Response Status', `${response.status} ${response.statusText}`));
      resultsArea.appendChild(createInfoRow('Error Details', errorText || 'No error details available'));
      
      displayNetDocumentsError(resultsArea, response.status, errorText);
    }
    
  } catch (error) {
    window.logDebug('NetDocuments API test failed with exception', { error: error.message });
    resultsArea.innerHTML = '';
    displaySafeError(resultsArea, `API test failed: ${error.message}`);
    
    if (error.message.includes('fetch')) {
      const networkNote = document.createElement('div');
      networkNote.style.marginTop = '10px';
      networkNote.style.padding = '8px';
      networkNote.style.backgroundColor = '#f8d7da';
      networkNote.style.border = '1px solid #f5c6cb';
      networkNote.style.borderRadius = '4px';
      networkNote.style.fontSize = '11px';
      networkNote.style.color = '#721c24';
      networkNote.textContent = 'Network error: Check CORS settings and ensure the NetDocuments API endpoints are accessible from this domain.';
      resultsArea.appendChild(networkNote);
    }
  }
}

async function getNetDocumentsInfo(properties) {
  window.logDebug('Getting NetDocuments document information', properties);
  
  const resultsArea = document.getElementById('netdocs-results');
  resultsArea.innerHTML = '<div style="color: #666; font-style: italic;">Fetching document information...</div>';
  
  try {
    const config = getNetDocumentsConfig();
    
    if (!config.isConfigured) {
      showNetDocumentsConfigPrompt(resultsArea, properties);
      return;
    }
    
    // NetDocuments REST API endpoint for document information
    // Format: /v2/document/{documentId}
    const documentEndpoint = `${config.baseUrl}/v2/document/${properties.ndDocumentId}`;
    
    window.logDebug('Fetching document info from NetDocuments', { 
      endpoint: documentEndpoint,
      documentId: properties.ndDocumentId
    });
    
    const response = await fetch(documentEndpoint, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    });
    
    window.logDebug('NetDocuments document info response', { 
      status: response.status,
      statusText: response.statusText,
      ok: response.ok
    });
    
    resultsArea.innerHTML = '';
    
    if (response.ok) {
      const docInfo = await response.json();
      window.logDebug('NetDocuments document info retrieved', docInfo);
      
      const infoSeparator = document.createElement('div');
      infoSeparator.style.marginTop = '10px';
      infoSeparator.style.marginBottom = '10px';
      infoSeparator.style.fontWeight = 'bold';
      infoSeparator.style.borderTop = '1px solid #ccc';
      infoSeparator.style.paddingTop = '8px';
      infoSeparator.textContent = 'Document Information from NetDocuments';
      resultsArea.appendChild(infoSeparator);
      
      // Display key document properties
      const displayFields = [
        { key: 'id', label: 'Document ID' },
        { key: 'name', label: 'Document Name' },
        { key: 'extension', label: 'File Extension' },
        { key: 'size', label: 'File Size' },
        { key: 'created', label: 'Created Date' },
        { key: 'modified', label: 'Modified Date' },
        { key: 'author', label: 'Author' },
        { key: 'version', label: 'Version' },
        { key: 'status', label: 'Status' },
        { key: 'repository', label: 'Repository' },
        { key: 'cabinet', label: 'Cabinet' },
        { key: 'workspace', label: 'Workspace' }
      ];
      
      displayFields.forEach(field => {
        if (docInfo[field.key] !== undefined && docInfo[field.key] !== null) {
          let value = docInfo[field.key];
          
          // Format specific fields
          if (field.key === 'size' && typeof value === 'number') {
            value = formatBytes(value);
          } else if ((field.key === 'created' || field.key === 'modified') && value) {
            value = new Date(value).toLocaleString();
          } else if (typeof value === 'object') {
            value = JSON.stringify(value);
          }
          
          const row = createInfoRow(field.label, value);
          row.style.borderLeft = '4px solid #007acc';
          resultsArea.appendChild(row);
        }
      });
      
      // Display custom attributes if available
      if (docInfo.attributes && Array.isArray(docInfo.attributes) && docInfo.attributes.length > 0) {
        const attrSeparator = document.createElement('div');
        attrSeparator.style.marginTop = '15px';
        attrSeparator.style.marginBottom = '10px';
        attrSeparator.style.fontWeight = 'bold';
        attrSeparator.style.borderTop = '1px solid #ccc';
        attrSeparator.style.paddingTop = '8px';
        attrSeparator.textContent = 'Document Attributes';
        resultsArea.appendChild(attrSeparator);
        
        docInfo.attributes.forEach(attr => {
          if (attr.name && attr.value !== undefined) {
            const row = createInfoRow(attr.name, attr.value);
            row.style.borderLeft = '4px solid #28a745';
            resultsArea.appendChild(row);
          }
        });
      }
      
      // Add success note
      const successNote = document.createElement('div');
      successNote.style.marginTop = '10px';
      successNote.style.padding = '8px';
      successNote.style.backgroundColor = '#d1ecf1';
      successNote.style.border = '1px solid #bee5eb';
      successNote.style.borderRadius = '4px';
      successNote.style.fontSize = '11px';
      successNote.style.color = '#0c5460';
      successNote.textContent = 'Document information successfully retrieved from NetDocuments REST API.';
      resultsArea.appendChild(successNote);
      
    } else {
      window.logDebug('NetDocuments document info request failed', { 
        status: response.status,
        statusText: response.statusText
      });
      
      const errorText = await response.text();
      window.logDebug('NetDocuments document info error response', errorText);
      
      resultsArea.appendChild(createInfoRow('Request Status', `${response.status} ${response.statusText}`));
      resultsArea.appendChild(createInfoRow('Document ID', properties.ndDocumentId));
      resultsArea.appendChild(createInfoRow('Error Details', errorText || 'No error details available'));
      
      displayNetDocumentsError(resultsArea, response.status, errorText);
    }
    
  } catch (error) {
    window.logDebug('NetDocuments document info retrieval failed', { error: error.message });
    resultsArea.innerHTML = '';
    displaySafeError(resultsArea, `Failed to get document info: ${error.message}`);
    
    if (error.message.includes('fetch')) {
      const networkNote = document.createElement('div');
      networkNote.style.marginTop = '10px';
      networkNote.style.padding = '8px';
      networkNote.style.backgroundColor = '#f8d7da';
      networkNote.style.border = '1px solid #f5c6cb';
      networkNote.style.borderRadius = '4px';
      networkNote.style.fontSize = '11px';
      networkNote.style.color = '#721c24';
      networkNote.textContent = 'Network error: Ensure the NetDocuments API is accessible and CORS is properly configured.';
      resultsArea.appendChild(networkNote);
    }
  }
}

// NetDocuments API Configuration and Helper Functions
function getNetDocumentsConfig() {
  // In a production environment, these would be stored securely
  // For POC, we'll prompt for configuration or use environment variables
  
  const config = {
    baseUrl: '',
    accessToken: '',
    isConfigured: false
  };
  
  // Try to get from sessionStorage first (temporary for testing)
  const storedConfig = sessionStorage.getItem('netdocs-config');
  if (storedConfig) {
    try {
      const parsed = JSON.parse(storedConfig);
      config.baseUrl = parsed.baseUrl;
      config.accessToken = parsed.accessToken;
      config.isConfigured = !!(config.baseUrl && config.accessToken);
    } catch (e) {
      window.logDebug('Failed to parse stored NetDocuments config', { error: e.message });
    }
  }
  
  // Default to EU region if no base URL configured
  if (!config.baseUrl) {
    config.baseUrl = 'https://api.eu.netdocuments.com';
  }
  
  window.logDebug('NetDocuments config loaded', { 
    hasBaseUrl: !!config.baseUrl,
    hasAccessToken: !!config.accessToken,
    isConfigured: config.isConfigured
  });
  
  return config;
}

function showNetDocumentsConfigPrompt(container, properties) {
  container.innerHTML = '';
  
  const configSeparator = document.createElement('div');
  configSeparator.style.marginTop = '10px';
  configSeparator.style.marginBottom = '15px';
  configSeparator.style.fontWeight = 'bold';
  configSeparator.style.borderTop = '1px solid #ccc';
  configSeparator.style.paddingTop = '8px';
  configSeparator.textContent = 'NetDocuments API Configuration Required';
  container.appendChild(configSeparator);
  
  const configForm = document.createElement('div');
  configForm.style.padding = '15px';
  configForm.style.backgroundColor = '#f8f9fa';
  configForm.style.border = '1px solid #dee2e6';
  configForm.style.borderRadius = '4px';
  configForm.style.marginBottom = '10px';
  
  // Base URL input
  const urlLabel = document.createElement('div');
  urlLabel.style.fontWeight = 'bold';
  urlLabel.style.marginBottom = '5px';
  urlLabel.textContent = 'NetDocuments API Base URL:';
  configForm.appendChild(urlLabel);
  
  const urlInput = document.createElement('input');
  urlInput.type = 'text';
  urlInput.placeholder = 'https://api.eu.netdocuments.com';
  urlInput.value = 'https://api.eu.netdocuments.com';
  urlInput.style.width = '100%';
  urlInput.style.padding = '5px';
  urlInput.style.marginBottom = '10px';
  urlInput.style.borderRadius = '3px';
  urlInput.style.border = '1px solid #ccc';
  configForm.appendChild(urlInput);
  
  // Access Token input
  const tokenLabel = document.createElement('div');
  tokenLabel.style.fontWeight = 'bold';
  tokenLabel.style.marginBottom = '5px';
  tokenLabel.textContent = 'Access Token:';
  configForm.appendChild(tokenLabel);
  
  const tokenInput = document.createElement('input');
  tokenInput.type = 'password';
  tokenInput.placeholder = 'Enter your NetDocuments access token';
  tokenInput.style.width = '100%';
  tokenInput.style.padding = '5px';
  tokenInput.style.marginBottom = '10px';
  tokenInput.style.borderRadius = '3px';
  tokenInput.style.border = '1px solid #ccc';
  configForm.appendChild(tokenInput);
  
  // Save Configuration Button
  const saveButton = document.createElement('button');
  saveButton.textContent = 'Save Configuration';
  saveButton.style.padding = '8px 16px';
  saveButton.style.backgroundColor = '#007acc';
  saveButton.style.color = 'white';
  saveButton.style.border = 'none';
  saveButton.style.borderRadius = '4px';
  saveButton.style.cursor = 'pointer';
  saveButton.style.marginRight = '10px';
  saveButton.onclick = () => {
    const config = {
      baseUrl: urlInput.value.trim(),
      accessToken: tokenInput.value.trim()
    };
    
    if (config.baseUrl && config.accessToken) {
      sessionStorage.setItem('netdocs-config', JSON.stringify(config));
      window.logDebug('NetDocuments configuration saved', { baseUrl: config.baseUrl });
      
      container.innerHTML = '<div style="color: #666; font-style: italic;">Configuration saved. You can now test the API connection.</div>';
      
      setTimeout(() => {
        testNetDocumentsAPI(properties);
      }, 1000);
    } else {
      alert('Please provide both Base URL and Access Token');
    }
  };
  configForm.appendChild(saveButton);
  
  container.appendChild(configForm);
  
  // Configuration Instructions
  const instructions = document.createElement('div');
  instructions.style.marginTop = '10px';
  instructions.style.padding = '10px';
  instructions.style.backgroundColor = '#e2e3e5';
  instructions.style.border = '1px solid #d6d8db';
  instructions.style.borderRadius = '4px';
  instructions.style.fontSize = '11px';
  instructions.innerHTML = '<strong>Configuration Instructions:</strong><br>' +
                          '1. Obtain an access token from your NetDocuments administrator<br>' +
                          '2. Choose the appropriate API base URL for your region<br>' +
                          '3. Ensure CORS is configured to allow requests from this domain<br>' +
                          '4. Access tokens are stored temporarily in browser session storage';
  container.appendChild(instructions);
}

function displayNetDocumentsError(container, statusCode, errorText) {
  const errorSeparator = document.createElement('div');
  errorSeparator.style.marginTop = '15px';
  errorSeparator.style.marginBottom = '10px';
  errorSeparator.style.fontWeight = 'bold';
  errorSeparator.style.borderTop = '1px solid #ccc';
  errorSeparator.style.paddingTop = '8px';
  errorSeparator.textContent = 'Error Details';
  container.appendChild(errorSeparator);
  
  let errorMessage = 'Unknown error occurred';
  let suggestion = 'Please check your configuration and try again';
  
  switch (statusCode) {
    case 401:
      errorMessage = 'Authentication failed';
      suggestion = 'Check your access token and ensure it has not expired';
      break;
    case 403:
      errorMessage = 'Access forbidden';
      suggestion = 'Your account may not have permission to access this document or API endpoint';
      break;
    case 404:
      errorMessage = 'Document or endpoint not found';
      suggestion = 'Verify the document ID and API base URL are correct';
      break;
    case 429:
      errorMessage = 'Rate limit exceeded';
      suggestion = 'Too many requests. Please wait before trying again';
      break;
    case 500:
      errorMessage = 'NetDocuments server error';
      suggestion = 'There may be a temporary issue with the NetDocuments service';
      break;
  }
  
  const errorNote = document.createElement('div');
  errorNote.style.marginTop = '10px';
  errorNote.style.padding = '10px';
  errorNote.style.backgroundColor = '#f8d7da';
  errorNote.style.border = '1px solid #f5c6cb';
  errorNote.style.borderRadius = '4px';
  errorNote.style.fontSize = '11px';
  errorNote.style.color = '#721c24';
  errorNote.innerHTML = `<strong>${errorMessage}</strong><br>${suggestion}`;
  container.appendChild(errorNote);
}
