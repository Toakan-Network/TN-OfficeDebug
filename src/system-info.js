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
    container.appendChild(createSectionSeparator('Critical Diagnostics'));
    
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
  if (!container) return;
  
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
