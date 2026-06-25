/* global Office */
// Main orchestrator - loads modules from helpers.js, office-info.js, 
// document-info.js, system-info.js, addin-info.js, dms-info.js

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

function loadDebugSection(sectionName, loadFn, containerId) {
  if (typeof loadFn !== 'function') {
    console.warn(`Skipping ${sectionName}: loader not available`);
    return;
  }

  try {
    loadFn();
  } catch (error) {
    console.error(`ERROR in ${sectionName}`, error);

    if (containerId) {
      const container = document.getElementById(containerId);
      if (container) {
        container.innerHTML = '';
        displaySafeError(container, `Error loading ${sectionName}: ${error.message}`);
      }
    }
  }
}

function loadAllDebugInfo() {
  loadDebugSection('Office info', loadOfficeInfo, 'office-info');
  loadDebugSection('Document info', loadDocumentInfo, 'document-info');
  loadDebugSection('Add-ins info', loadAddinsInfo, 'addins-info');
  loadDebugSection('System info', loadSystemInfo, 'system-info');
  loadDebugSection('DMS info', loadDMSInfo, 'dms-info');
  loadDebugSection('Actions', loadActions, 'actions-info');
}