/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById('refresh-btn').onclick = loadAllDebugInfo;
    loadAllDebugInfo();
  }
});

function loadAllDebugInfo() {
  loadOfficeInfo();
  loadDocumentInfo();
  loadAddinsInfo();
  loadSystemInfo();
  loadContextInfo();
}

function createInfoRow(label, value, isCode = false) {
  const row = document.createElement('div');
  row.className = 'info-row';
  
  const labelDiv = document.createElement('div');
  labelDiv.className = 'info-label';
  labelDiv.textContent = label + ':';
  
  const valueDiv = document.createElement('div');
  valueDiv.className = isCode ? 'info-value code' : 'info-value';
  valueDiv.textContent = value !== null && value !== undefined ? value.toString() : 'N/A';
  
  row.appendChild(labelDiv);
  row.appendChild(valueDiv);
  
  return row;
}

function loadOfficeInfo() {
  const container = document.getElementById('office-info');
  container.innerHTML = '';

  try {
    // Office.context information
    container.appendChild(createInfoRow('Host', Office.context.host));
    container.appendChild(createInfoRow('Platform', Office.context.platform));
    
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
  container.innerHTML = '';

  Word.run(async (context) => {
    try {
      const doc = context.document;
      const properties = doc.properties;
      const body = doc.body;
      
      // Load properties
      properties.load('title,subject,author,keywords,comments,template,lastAuthor,revisionNumber,applicationName,lastPrintDate,creationDate,lastSaveTime');
      body.load('text,type');
      
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
      
    } catch (error) {
      container.innerHTML = `<div class="error">Error loading document info: ${error.message}</div>`;
    }
  }).catch((error) => {
    container.innerHTML = `<div class="error">Error: ${error.message}</div>`;
  });
}

function loadAddinsInfo() {
  const container = document.getElementById('addins-info');
  container.innerHTML = '';

  try {
    // Current add-in info
    if (Office.context.document && Office.context.document.settings) {
      container.appendChild(createInfoRow('Settings Available', 'Yes'));
    }
    
    // Roaming settings
    if (Office.context.roamingSettings) {
      const settings = Office.context.roamingSettings;
      const settingsKeys = Object.keys(settings);
      container.appendChild(createInfoRow('Roaming Settings Count', settingsKeys.length));
    }
    
    // Add-in license
    if (Office.context.license) {
      container.appendChild(createInfoRow('License Value', Office.context.license.value || 'N/A'));
    }
    
    // Current add-in manifest info
    if (Office.context.manifest) {
      container.appendChild(createInfoRow('Manifest ID', Office.context.manifest.id || 'N/A', true));
      container.appendChild(createInfoRow('Manifest Version', Office.context.manifest.version || 'N/A'));
    }
    
    // Add-in UI
    if (Office.context.ui) {
      container.appendChild(createInfoRow('UI API Available', 'Yes'));
    }
    
    // Note about limitations
    const note = document.createElement('div');
    note.className = 'info-value';
    note.style.marginTop = '15px';
    note.style.fontStyle = 'italic';
    note.style.color = '#666';
    note.innerHTML = '<strong>Note:</strong> Office Add-ins run in a sandboxed environment and cannot directly enumerate other installed add-ins for security reasons. This section shows information about the current add-in only.';
    container.appendChild(note);
    
  } catch (error) {
    container.innerHTML = `<div class="error">Error loading add-ins info: ${error.message}</div>`;
  }
}

function loadSystemInfo() {
  const container = document.getElementById('system-info');
  container.innerHTML = '';

  try {
    // Browser/User Agent info
    container.appendChild(createInfoRow('User Agent', navigator.userAgent));
    container.appendChild(createInfoRow('Platform', navigator.platform));
    container.appendChild(createInfoRow('Language', navigator.language));
    container.appendChild(createInfoRow('Online', navigator.onLine ? 'Yes' : 'No'));
    container.appendChild(createInfoRow('Cookie Enabled', navigator.cookieEnabled ? 'Yes' : 'No'));
    
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
