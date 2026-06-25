/* global Office, Word, Excel, PowerPoint */

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
      
      const doc = context.document;
      const properties = doc.properties;
      window.logDebug('Got document and properties objects', { 
        hasDoc: !!doc, 
        hasProperties: !!properties 
      });
      
      // Load custom properties using the exact same pattern as Document Info tab
      const customProperties = properties.customProperties;
      window.logDebug('Got custom properties object', { hasCustomProperties: !!customProperties });
      
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
    
    if (key === 'nddocumentid') {
      netDocumentsId = value;
      window.logDebug('NetDocuments ID found (fallback)', { id: value });
    }
    
    if (key === 'isimanagework') {
      iManageWork = value;
      window.logDebug('iManage Work indicator found (fallback)', { value: value });
    }
    
    if (key === 'documentnumber') {
      documentNumber = value;
      window.logDebug('Document Number found (fallback)', { number: value });
    }
  });
  
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
    container.appendChild(createSectionSeparator('Debug Information'));
    
    Object.entries(dmsInfo.debugInfo).forEach(([key, value]) => {
      const row = createInfoRow(key, JSON.stringify(value));
      row.style.borderLeft = '4px solid #ff6b6b';
      container.appendChild(row);
    });
  }
  
  if (dmsInfo.detected) {
    container.appendChild(createSectionSeparator(`${dmsInfo.type.toUpperCase()} Properties`));
    
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
    container.appendChild(createSectionSeparator('All Custom Properties'));
    
    dmsInfo.allProperties.forEach(prop => {
      const row = createInfoRow(prop.key, prop.value);
      row.style.borderLeft = '4px solid #28a745';
      container.appendChild(row);
    });
    
    container.appendChild(createInfoRow('Total Custom Properties', dmsInfo.allProperties.length));
  } else {
    container.appendChild(createSectionSeparator('Custom Properties Status'));
    
    container.appendChild(createInfoRow('Custom Properties Found', '0'));
    container.appendChild(createInfoRow('Note', 'No custom properties detected - check console for detailed debugging info'));
  }
}

function addNetDocumentsActions(container, properties) {
  window.logDebug('Adding NetDocuments actions', properties);
  
  container.appendChild(createSectionSeparator('NetDocuments Actions'));
  
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
  
  container.appendChild(createSectionSeparator('iManage Actions'));
  
  const placeholder = document.createElement('div');
  placeholder.style.padding = '10px';
  placeholder.style.backgroundColor = '#f8f9fa';
  placeholder.style.border = '1px solid #dee2e6';
  placeholder.style.borderRadius = '4px';
  placeholder.style.fontStyle = 'italic';
  placeholder.textContent = 'iManage API integration will be implemented in future version';
  container.appendChild(placeholder);
}

async function refreshNetDocumentsToken() {
  const config = getNetDocumentsConfig();
  
  if (!config.refreshToken) {
    throw new Error('No refresh token available');
  }
  
  const tokenEndpoint = `${config.oauthUrl}/v2/oauth2/token`;
  
  const refreshRequest = {
    grant_type: 'refresh_token',
    refresh_token: config.refreshToken,
    client_id: config.clientId,
    client_secret: config.clientSecret
  };
  
  window.logDebug('Refreshing NetDocuments access token');
  
  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Accept': 'application/json'
    },
    body: new URLSearchParams(refreshRequest)
  });
  
  if (response.ok) {
    const tokenData = await response.json();
    
    // Calculate token expiry
    const expiryTime = new Date();
    expiryTime.setSeconds(expiryTime.getSeconds() + (tokenData.expires_in || 3600));
    
    // Store new tokens
    const tokens = {
      accessToken: tokenData.access_token,
      refreshToken: tokenData.refresh_token || config.refreshToken,
      tokenExpiry: expiryTime.toISOString()
    };
    
    sessionStorage.setItem('netdocs-tokens', JSON.stringify(tokens));
    window.logDebug('NetDocuments token refreshed successfully');
    
    return tokenData.access_token;
  } else {
    const errorText = await response.text();
    window.logDebug('NetDocuments token refresh failed', { 
      status: response.status,
      error: errorText
    });
    throw new Error(`Token refresh failed: ${response.status} ${response.statusText}`);
  }
}

async function makeNetDocumentsAPICall(endpoint, options = {}) {
  let config = getNetDocumentsConfig();
  
  // If not configured or token expired, try to refresh
  if (!config.isConfigured && config.refreshToken) {
    try {
      await refreshNetDocumentsToken();
      config = getNetDocumentsConfig();
    } catch (error) {
      window.logDebug('Token refresh failed', { error: error.message });
      throw new Error('Authentication required - please re-authenticate');
    }
  }
  
  if (!config.isConfigured) {
    throw new Error('NetDocuments not configured - authentication required');
  }
  
  const url = `${config.baseUrl}${endpoint}`;
  const requestOptions = {
    method: options.method || 'GET',
    headers: {
      'Authorization': `Bearer ${config.accessToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      ...options.headers
    },
    ...options
  };
  
  window.logDebug('Making NetDocuments API call', { 
    endpoint: endpoint,
    method: requestOptions.method
  });
  
  const response = await fetch(url, requestOptions);
  
  // If unauthorized, try to refresh token once
  if (response.status === 401 && config.refreshToken) {
    window.logDebug('API call unauthorized, attempting token refresh');
    try {
      await refreshNetDocumentsToken();
      config = getNetDocumentsConfig();
      
      requestOptions.headers['Authorization'] = `Bearer ${config.accessToken}`;
      return await fetch(url, requestOptions);
    } catch (refreshError) {
      window.logDebug('Token refresh on 401 failed', { error: refreshError.message });
      throw new Error('Authentication expired - please re-authenticate');
    }
  }
  
  return response;
}

// NetDocuments API Functions (OAuth Implementation)
async function testNetDocumentsAPI(properties) {
  window.logDebug('Testing NetDocuments API connection', properties);
  
  const resultsArea = document.getElementById('netdocs-results');
  resultsArea.innerHTML = '<div style="color: #666; font-style: italic;">Testing API connection...</div>';
  
  try {
    const config = getNetDocumentsConfig();
    
    if (!config.isConfigured) {
      showNetDocumentsConfigPrompt(resultsArea, properties);
      return;
    }
    
    // Test API connectivity with a simple repository list call
    const response = await makeNetDocumentsAPICall('/v2/repository');
    
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
      resultsArea.appendChild(createInfoRow('OAuth Server', config.oauthUrl));
      resultsArea.appendChild(createInfoRow('API Server', config.baseUrl));
      resultsArea.appendChild(createInfoRow('Client ID', config.clientId));
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
      successNote.textContent = 'NetDocuments OAuth API connection successful! You can now retrieve document information.';
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
    
    if (error.message.includes('authentication required')) {
      showNetDocumentsConfigPrompt(resultsArea, properties);
    } else {
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
    
    window.logDebug('Fetching document info from NetDocuments', { 
      documentId: properties.ndDocumentId
    });
    
    const response = await makeNetDocumentsAPICall(`/v2/document/${properties.ndDocumentId}`);
    
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
    
    if (error.message.includes('authentication required')) {
      showNetDocumentsConfigPrompt(resultsArea, properties);
    } else {
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
}

// NetDocuments API Configuration and Helper Functions
function getNetDocumentsConfig() {
  // For Office Add-ins, secure authentication requires:
  // 1. Dialog API for OAuth authorization code flow (most secure)
  // 2. Server-side proxy for API calls (avoids CORS)
  // 3. Pre-configured API keys for read-only operations (if available)
  
  const config = {
    baseUrl: 'https://api.eu.netdocuments.com',
    oauthUrl: 'https://eu.netdocuments.com',
    clientId: 'AP-9L3NZFJO',
    
    // Note: In production, authentication should use:
    // - Office Dialog API for secure OAuth flow
    // - Server-side proxy to handle API calls and avoid CORS
    // - Token storage in secure, server-side session storage
    
    isConfigured: false,
    accessToken: '',
    authMethod: 'demonstration-only'
  };
  
  return config;
}

function showNetDocumentsConfigPrompt(container, properties) {
  window.logDebug('Showing NetDocuments authentication information');
  
  container.innerHTML = '';
  
  container.appendChild(createSectionSeparator('NetDocuments Authentication Required'));
  
  // Security notice
  const securityNotice = document.createElement('div');
  securityNotice.style.marginBottom = '15px';
  securityNotice.style.padding = '12px';
  securityNotice.style.backgroundColor = '#fff3cd';
  securityNotice.style.border = '1px solid #ffeaa7';
  securityNotice.style.borderRadius = '4px';
  securityNotice.style.fontSize = '12px';
  securityNotice.style.color = '#856404';
  securityNotice.innerHTML = `
    <strong>🔒 Security Notice:</strong><br>
    For production use, NetDocuments authentication should be implemented using:<br>
    • Office Dialog API for secure OAuth authorization code flow<br>
    • Server-side proxy to handle API calls and avoid CORS restrictions<br>
    • Secure token storage on the server side
  `;
  container.appendChild(securityNotice);
  
  // Current limitations
  const limitationsSection = document.createElement('div');
  limitationsSection.style.marginBottom = '15px';
  limitationsSection.style.padding = '12px';
  limitationsSection.style.backgroundColor = '#f8d7da';
  limitationsSection.style.border = '1px solid #f5c6cb';
  limitationsSection.style.borderRadius = '4px';
  limitationsSection.style.fontSize = '12px';
  limitationsSection.style.color = '#721c24';
  limitationsSection.innerHTML = `
    <strong>⚠️ Current Limitations:</strong><br>
    • CORS policies prevent direct OAuth calls from Office add-ins<br>
    • Resource Owner Password Credentials flow is not secure for this context<br>
    • NetDocuments API requires proper authentication infrastructure
  `;
  container.appendChild(limitationsSection);
  
  // Document information we can show
  container.appendChild(createInfoRow('Detected Document System', 'NetDocuments'));
  container.appendChild(createInfoRow('Document ID Property', 'ndDocumentId'));
  container.appendChild(createInfoRow('Document ID Value', properties.ndDocumentId || 'Not found'));
  container.appendChild(createInfoRow('API Client ID', 'AP-9L3NZFJO (configured)'));
  container.appendChild(createInfoRow('API Base URL', 'https://api.eu.netdocuments.com'));
  container.appendChild(createInfoRow('OAuth URL', 'https://eu.netdocuments.com'));
  
  // Recommended implementation approach
  const implementationSection = document.createElement('div');
  implementationSection.style.marginTop = '15px';
  implementationSection.style.marginBottom = '10px';
  implementationSection.style.fontWeight = 'bold';
  implementationSection.style.borderTop = '1px solid #ccc';
  implementationSection.style.paddingTop = '8px';
  implementationSection.textContent = 'Recommended Implementation for Production';
  container.appendChild(implementationSection);
  
  const recommendations = [
    'Use Office.context.ui.displayDialogAsync() for OAuth authorization code flow',
    'Implement server-side API proxy to handle NetDocuments REST calls',
    'Store tokens securely on server side, not in browser storage',
    'Configure CORS properly on your API proxy server',
    'Use HTTPS for all authentication flows and API communications'
  ];
  
  recommendations.forEach((rec, index) => {
    const recRow = document.createElement('div');
    recRow.style.margin = '5px 0';
    recRow.style.padding = '8px';
    recRow.style.backgroundColor = '#d1ecf1';
    recRow.style.border = '1px solid #bee5eb';
    recRow.style.borderRadius = '4px';
    recRow.style.fontSize = '11px';
    recRow.style.color = '#0c5460';
    recRow.textContent = `${index + 1}. ${rec}`;
    container.appendChild(recRow);
  });
  
  // Note about current status
  const statusNote = document.createElement('div');
  statusNote.style.marginTop = '15px';
  statusNote.style.padding = '10px';
  statusNote.style.backgroundColor = '#e2e3e5';
  statusNote.style.border = '1px solid #d6d8db';
  statusNote.style.borderRadius = '4px';
  statusNote.style.fontSize = '11px';
  statusNote.style.color = '#383d41';
  statusNote.innerHTML = `
    <strong>📋 Current Status:</strong><br>
    This implementation demonstrates NetDocuments detection and shows the proper architecture needed for secure API integration. 
    The actual API calls require a server-side implementation to handle authentication and avoid CORS restrictions.
  `;
  container.appendChild(statusNote);
}
