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

function loadAllDebugInfo() {
  try {
    loadOfficeInfo();
    loadDocumentInfo();
    loadAddinsInfo();
    loadSystemInfo();
    loadDMSInfo();
    loadActions();
  } catch (error) {
    console.error('ERROR in loadAllDebugInfo', error);
  }
}