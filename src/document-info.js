/* global Office, Word, Excel, PowerPoint */

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
      window.logDebug('Calling loadWordDocumentInfo');
      loadWordDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.Excel) {
      window.logDebug('Calling loadExcelDocumentInfo');
      loadExcelDocumentInfo(container);
    } else if (Office.context.host === Office.HostType.PowerPoint) {
      window.logDebug('Calling loadPowerPointDocumentInfo');
      loadPowerPointDocumentInfo(container);
    } else {
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
        
        // Document Security
        container.appendChild(createSectionSeparator('Document Security'));
        
        // Simple protection check
        try {
          const contentControls = context.document.contentControls;
          contentControls.load('items');
          await context.sync();
          container.appendChild(createInfoRow('Content Controls', contentControls.items.length > 0 ? `Yes (${contentControls.items.length})` : 'None'));
        } catch (controlError) {
          container.appendChild(createInfoRow('Content Controls', 'Could not check'));
        }
        
        // Document Protection and Security Info
        container.appendChild(createSectionSeparator('Document Security & Protection'));
        
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
        }
        
        // Custom Properties Section
        if (customProperties.items && customProperties.items.length > 0) {
          container.appendChild(createSectionSeparator('Custom Document Properties'));
          
          // Display each custom property  
          customProperties.items.forEach((customProp, index) => {
            const propName = customProp.key || `Property ${index + 1}`;
            const propValue = customProp.value !== null && customProp.value !== undefined ? customProp.value.toString() : 'N/A';
            const propType = customProp.type || 'Unknown';
            const propRow = createInfoRow(propName, propValue);
            
            // Add type information as a small subtitle
            const typeSpan = document.createElement('div');
            typeSpan.className = 'info-meta';
            typeSpan.textContent = `Type: ${propType}`;
            
            const valueDiv = propRow.lastElementChild;
            valueDiv.appendChild(typeSpan);
            
            // Different border color for custom properties
            propRow.classList.add('custom-property');
            
            container.appendChild(propRow);
          });
          
          container.appendChild(createInfoRow('Total Custom Properties', customProperties.items.length));
        } else {
          container.appendChild(createSectionSeparator('Custom Document Properties'));
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
      container.appendChild(createSectionSeparator('Custom Properties'));
      
      if (customProperties.items && customProperties.items.length > 0) {
        customProperties.items.forEach((customProp, index) => {
          try {
            const propRow = createInfoRow(customProp.key, customProp.value);
            propRow.classList.add('custom-property');
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
    
    // Document mode
    if (Office.context.document && Office.context.document.mode !== undefined) {
      const mode = Office.context.document.mode === Office.DocumentMode.ReadOnly ? 'Read-Only' : 'Read-Write';
      container.appendChild(createInfoRow('Document Mode', mode));
    }
    
    // Try to get any additional document settings
    if (Office.context.document && Office.context.document.settings) {
      window.logDebug('Document settings available');
    }
    
    // Try basic custom properties via Office.context
    if (Office.context.document && Office.context.document.customProperties) {
      window.logDebug('Attempting custom properties via Office.context');
      
      Office.context.document.customProperties.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          window.logDebug('Custom properties loaded via Office.context', { 
            count: result.value ? result.value.length : 0,
            properties: result.value
          });
          
          if (result.value && result.value.length > 0) {
            container.appendChild(createSectionSeparator('Custom Properties'));
            
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
    
    container.appendChild(createInfoRow('Status', 'Basic document information loaded successfully'));
    
    window.logDebug('PowerPoint document info loading completed using basic APIs');
    
  } catch (error) {
    window.logDebug('Error loading PowerPoint document info via Office.context', { 
      error: error.message,
      stack: error.stack,
      name: error.name 
    });
    
    displaySafeError(container, `Error accessing document information: ${error.message}`);
    
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
  container.appendChild(createInfoRow('Host Application', Office.context.host || 'Unknown'));
  
  if (Office.context.document && Office.context.document.url) {
    container.appendChild(createInfoRow('Document URL', Office.context.document.url));
  }
}
