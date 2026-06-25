/* global Office, Word */

// Footer Content Control Configuration
const FOOTER_CC_TAG = 'TN-DocRef';
const FOOTER_CC_TITLE = 'Document Reference';

// Track AfterSave handler registration
let afterSaveRegistered = false;

/**
 * Load the Actions panel UI
 */
function loadActions() {
  window.logDebug('Actions panel loading started');
  const container = document.getElementById('actions-info');

  if (!container) {
    console.error('ERROR: actions-info container not found');
    return;
  }

  container.innerHTML = '';

  // --- Footer Document Reference Section ---
  container.appendChild(createSectionSeparator('Footer Document Reference'));

  const description = document.createElement('div');
  description.style.padding = '8px';
  description.style.marginBottom = '10px';
  description.style.fontSize = '11px';
  description.style.color = '#555';
  description.textContent =
    'Inserts or updates a content control in the document footer (right-aligned) with the document reference ID. Uses NetDocuments ID, iManage Document Number, or the document filename as a fallback.';
  container.appendChild(description);

  // Status area – shows detected reference
  const statusRow = document.createElement('div');
  statusRow.id = 'docref-status';
  container.appendChild(statusRow);
  updateDocRefStatus();

  // Apply / Update button
  const applyBtn = document.createElement('button');
  applyBtn.textContent = 'Apply Document Reference to Footer';
  applyBtn.className = 'action-btn primary';
  applyBtn.onclick = handleApplyDocRef;
  container.appendChild(applyBtn);

  // Result area
  const resultArea = document.createElement('div');
  resultArea.id = 'docref-result';
  resultArea.style.marginTop = '10px';
  container.appendChild(resultArea);

  // --- Auto-Update on Save Section ---
  container.appendChild(createSectionSeparator('Auto-Update on Save'));

  const autoDesc = document.createElement('div');
  autoDesc.style.padding = '8px';
  autoDesc.style.marginBottom = '10px';
  autoDesc.style.fontSize = '11px';
  autoDesc.style.color = '#555';
  autoDesc.textContent =
    'When available, automatically re-applies the document reference after each save.';
  container.appendChild(autoDesc);

  const autoSaveStatus = document.createElement('div');
  autoSaveStatus.id = 'autosave-status';
  container.appendChild(autoSaveStatus);

  updateAutoSaveStatus('Initializing...');
  registerAfterSaveHandler();

  // --- Heading 1 Style Override Section ---
  container.appendChild(createSectionSeparator('Heading 1 Style Override'));

  const h1Desc = document.createElement('div');
  h1Desc.style.padding = '8px';
  h1Desc.style.marginBottom = '10px';
  h1Desc.style.fontSize = '11px';
  h1Desc.style.color = '#555';
  h1Desc.textContent =
    'Changes the built-in Heading 1 style font to Wingdings at size 20 and refreshes the document.';
  container.appendChild(h1Desc);

  const h1Btn = document.createElement('button');
  h1Btn.textContent = 'Apply Wingdings to Heading 1';
  h1Btn.className = 'action-btn primary';
  h1Btn.onclick = handleApplyWingdingsHeading1;
  container.appendChild(h1Btn);

  const h1Result = document.createElement('div');
  h1Result.id = 'h1style-result';
  h1Result.style.marginTop = '10px';
  container.appendChild(h1Result);
}

// ─── Heading 1 Style Override ────────────────────────────────────────────────

async function handleApplyWingdingsHeading1() {
  const resultArea = document.getElementById('h1style-result');
  if (!resultArea) return;

  resultArea.innerHTML = '';
  resultArea.appendChild(createInfoRow('Status', 'Applying style change...'));

  try {
    if (Office.context.host !== Office.HostType.Word) {
      throw new Error('Style modification is only supported in Word');
    }

    await Word.run(async (context) => {
      const heading1Style = context.document.getStyles().getByNameOrNullObject('Heading 1');
      heading1Style.load('nameLocal');
      await context.sync();

      if (heading1Style.isNullObject) {
        throw new Error('Heading 1 style not found in document');
      }

      heading1Style.font.name = 'Wingdings';
      heading1Style.font.size = 20;
      await context.sync();

      // Force a visual refresh by toggling a no-op body edit
      const body = context.document.body;
      body.load('text');
      await context.sync();
    });

    resultArea.innerHTML = '';
    const successDiv = document.createElement('div');
    successDiv.className = 'success';
    successDiv.textContent = 'Heading 1 style updated to Wingdings, size 20';
    resultArea.appendChild(successDiv);
    window.logDebug('Heading 1 style updated to Wingdings size 20');
  } catch (error) {
    resultArea.innerHTML = '';
    window.logDebug('Error applying Wingdings to Heading 1', { error: error.message, stack: error.stack });
    displaySafeError(resultArea, 'Error: ' + error.message);
  }
}

// ─── Document Reference Detection ────────────────────────────────────────────

async function getDocumentReference() {
  try {
    // Reuse detectDMS() from dms-info.js (loaded globally)
    const dmsInfo = await detectDMS();

    if (dmsInfo.type === 'netdocuments' && dmsInfo.properties && dmsInfo.properties.ndDocumentId) {
      return { docRef: dmsInfo.properties.ndDocumentId, source: 'NetDocuments (ndDocumentId)' };
    }

    if (dmsInfo.type === 'imanage' && dmsInfo.properties && dmsInfo.properties.DocumentNumber) {
      return { docRef: dmsInfo.properties.DocumentNumber, source: 'iManage (DocumentNumber)' };
    }
  } catch (error) {
    window.logDebug('DMS detection failed for document reference', { error: error.message });
  }

  // Fallback – document filename from URL
  const docUrl = Office.context.document ? Office.context.document.url : '';
  if (docUrl) {
    const fileName = docUrl.split(/[/\\]/).pop() || 'Unknown Document';
    return { docRef: fileName, source: 'Document filename (no DMS detected)' };
  }

  return { docRef: 'Unknown Document', source: 'Fallback (no DMS or URL available)' };
}

async function updateDocRefStatus() {
  const statusContainer = document.getElementById('docref-status');
  if (!statusContainer) return;

  statusContainer.innerHTML = '';
  statusContainer.appendChild(createInfoRow('Status', 'Detecting document reference...'));

  try {
    const { docRef, source } = await getDocumentReference();
    statusContainer.innerHTML = '';
    statusContainer.appendChild(createInfoRow('Reference', docRef));
    statusContainer.appendChild(createInfoRow('Source', source));
  } catch (error) {
    statusContainer.innerHTML = '';
    statusContainer.appendChild(createInfoRow('Status', 'Error: ' + error.message));
  }
}

// ─── Footer Content Control Operations ───────────────────────────────────────

async function applyDocumentReferenceToFooter() {
  if (Office.context.host !== Office.HostType.Word) {
    throw new Error('Footer content controls are only supported in Word');
  }

  const { docRef, source } = await getDocumentReference();
  window.logDebug('Applying document reference to footer', { docRef, source });

  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();

    if (sections.items.length === 0) {
      throw new Error('No sections found in document');
    }

    const footer = sections.items[0].getFooter(Word.HeaderFooterType.primary);

    // Look for an existing content control with our tag
    const existingControls = footer.contentControls.getByTag(FOOTER_CC_TAG);
    existingControls.load('items');
    await context.sync();

    if (existingControls.items.length > 0) {
      // Update the existing content control text
      const cc = existingControls.items[0];
      cc.insertText(docRef, Word.InsertLocation.replace);
      await context.sync();
      window.logDebug('Updated existing footer content control', { docRef });
    } else {
      // Insert a new right-aligned paragraph at the end of the footer
      const paragraph = footer.insertParagraph(docRef, Word.InsertLocation.end);
      paragraph.alignment = Word.Alignment.right;
      paragraph.font.size = 8;
      paragraph.font.color = '#888888';

      // Wrap paragraph in a tagged content control
      const cc = paragraph.insertContentControl();
      cc.tag = FOOTER_CC_TAG;
      cc.title = FOOTER_CC_TITLE;
      cc.appearance = Word.ContentControlAppearance.boundingBox;

      await context.sync();
      window.logDebug('Created new footer content control', { docRef });
    }
  });

  return { docRef, source };
}

async function handleApplyDocRef() {
  const resultArea = document.getElementById('docref-result');
  if (!resultArea) return;

  resultArea.innerHTML = '';
  resultArea.appendChild(createInfoRow('Status', 'Applying...'));

  try {
    const { docRef, source } = await applyDocumentReferenceToFooter();
    resultArea.innerHTML = '';

    const successDiv = document.createElement('div');
    successDiv.className = 'success';
    successDiv.textContent = 'Footer updated with: ' + docRef;
    resultArea.appendChild(successDiv);
    resultArea.appendChild(createInfoRow('Source', source));

    // Refresh the status area too
    updateDocRefStatus();
  } catch (error) {
    resultArea.innerHTML = '';
    window.logDebug('Error applying document reference', { error: error.message, stack: error.stack });
    displaySafeError(resultArea, 'Error: ' + error.message);
  }
}

// ─── AfterSave Event Handler ─────────────────────────────────────────────────

async function registerAfterSaveHandler() {
  if (afterSaveRegistered) {
    window.logDebug('AfterSave handler already registered');
    return;
  }

  if (Office.context.host !== Office.HostType.Word) {
    window.logDebug('AfterSave handler only supported in Word');
    updateAutoSaveStatus('Not available (Word only)');
    return;
  }

  // Attempt 1: Word JS API onSaved event (requires newer WordApi)
  try {
    await Word.run(async (context) => {
      context.document.onSaved.add(handleDocumentSaved);
      await context.sync();
    });
    afterSaveRegistered = true;
    window.logDebug('AfterSave handler registered via Word API onSaved');
    updateAutoSaveStatus('Active – footer will auto-update on save');
    return;
  } catch (e) {
    window.logDebug('Word API onSaved not available', { error: e.message });
  }

  // Attempt 2: Common API documentSaved handler
  try {
    await new Promise((resolve, reject) => {
      Office.context.document.addHandlerAsync(
        'documentSaved',
        handleDocumentSavedCommon,
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error ? result.error.message : 'Unknown error'));
          }
        }
      );
    });
    afterSaveRegistered = true;
    window.logDebug('AfterSave handler registered via common API');
    updateAutoSaveStatus('Active – footer will auto-update on save');
    return;
  } catch (e) {
    window.logDebug('Common API documentSaved not available', { error: e.message });
  }

  // Neither approach worked
  updateAutoSaveStatus('Not available in this Office version – use the button above after saving');
}

function handleDocumentSaved(event) {
  window.logDebug('Document saved event fired (Word API)', event);
  applyDocumentReferenceToFooter()
    .then(({ docRef }) => {
      window.logDebug('Footer reference updated after save', { docRef });
      updateAutoSaveStatus('Active – last updated: ' + new Date().toLocaleTimeString());
    })
    .catch(error => {
      window.logDebug('Error updating footer after save', { error: error.message });
      updateAutoSaveStatus('Error on last save: ' + error.message);
    });
}

function handleDocumentSavedCommon() {
  window.logDebug('Document saved event fired (common API)');
  applyDocumentReferenceToFooter()
    .then(({ docRef }) => {
      window.logDebug('Footer reference updated after save', { docRef });
      updateAutoSaveStatus('Active – last updated: ' + new Date().toLocaleTimeString());
    })
    .catch(error => {
      window.logDebug('Error updating footer after save', { error: error.message });
      updateAutoSaveStatus('Error on last save: ' + error.message);
    });
}

function updateAutoSaveStatus(message) {
  const statusContainer = document.getElementById('autosave-status');
  if (!statusContainer) return;

  statusContainer.innerHTML = '';
  statusContainer.appendChild(createInfoRow('AfterSave Listener', message || 'Initializing...'));
}
