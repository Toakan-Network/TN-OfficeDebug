/* global Office */

// Debug logging function
window.logDebug = function(message, data = null) {
  if (data) {
    console.log(`[DEBUG] ${message}:`, data);
  } else {
    console.log(`[DEBUG] ${message}`);
  }
};

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

// Helper function to format bytes
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Helper function to create section separators
function createSectionSeparator(title) {
  const separator = document.createElement('div');
  separator.style.marginTop = '20px';
  separator.style.marginBottom = '15px';
  separator.style.fontWeight = 'bold';
  separator.style.borderTop = '1px solid #ccc';
  separator.style.paddingTop = '10px';
  separator.textContent = title;
  return separator;
}
