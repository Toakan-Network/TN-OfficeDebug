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
  const container = document.createElement('div');
  container.className = 'info-row';

  const labelDiv = document.createElement('div');
  labelDiv.className = 'info-label';
  labelDiv.textContent = label;

  const valueDiv = document.createElement('div');
  valueDiv.className = isCode ? 'info-value code' : 'info-value';

  if (isCode) {
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
  separator.className = 'section-separator';
  separator.textContent = title;
  return separator;
}
