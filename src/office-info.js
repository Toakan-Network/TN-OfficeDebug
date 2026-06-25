/* global Office */

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
    container.appendChild(createSectionSeparator('Key Diagnostics'));
    
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
