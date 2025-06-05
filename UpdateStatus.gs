// ===== ARCHIVO: UpdateStatus_ImprovedFlow.gs =====
// Sistema de Escalación PR - PASO 3: Flujo Visual Mejorado
// Interfaz con progreso claro y feedback en tiempo real

// ===== CONSTANTES LOCALES =====
const MAL_WEEK_COL = 2, MAL_DATE_COL = 3, MAL_NPI_COL = 4, MAL_PROVIDER_COL = 5, MAL_POLICY_COL = 10, MAL_STATUS_COL = 14;
const EXP_WEEK_COL = 1, EXP_DATE_COL = 2, EXP_NPI_COL = 3, EXP_PROVIDER_COL = 4, EXP_LICENSE_COL = 7, EXP_STATUS_COL = 12;
const EC_WEEK_COL = 1, EC_NPI_COL = 2, EC_POLICY_LICENSE_COL = 3, EC_PROVIDER_COL = 4, EC_STATUS_COL = 7, EC_TYPE_COL = 10, EC_ESCALATION_LEVEL_COL = 11;

const HISTORY_HEADERS = [
  'NPI', 'Policy/License Number', 'Provider Name', 'Type (MAL/EXP)',
  'OUTREACH #1 Date', 'OUTREACH #2 Date', 'OUTREACH #3 Date', 'Call before PR Date',
  'FIRST ESCALATION TO PR Date', 'SECOND ESCALATION TO PR Date', 'FINAL ESCALATION Date'
];

const OUTREACH_STATUS_MAPPING = {
  'MALPRACTICE': {
    outreach1: ['OUTREACH - EXPIRE', 'OUTREACH #1'],
    outreach2: ['OUTREACH - 2ND ATTEMPT', 'OUTREACH #2'],
    outreach3: ['OUTREACH 3RD ATTEMPT', 'OUTREACH #3'],
    callBeforePR: ['CALL BEFORE PR']
  },
  'EXPIRABLES': {
    outreach1: ['OUTREACH - EXPIRE', 'OUTREACH #1'],
    outreach2: ['OUTREACH - 2ND ATTEMPT', 'OUTREACH #2'],
    outreach3: ['OUTREACH - 3RD ATTEMPT', 'OUTREACH #3'],
    callBeforePR: ['CALL BEFORE PR']
  }
};

// ===== FUNCIÓN PRINCIPAL CON FLUJO MEJORADO =====
function updateStatusFromEscalationCenter() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaEscalation = ss.getSheetByName('ESCALATION CENTER');
    
    if (!hojaEscalation) {
      showErrorToast('ESCALATION CENTER sheet not found. Please run Step 1 first.');
      return;
    }

    // PASO 1: Verificar datos disponibles
    showProgressToast('Checking available data...', '🔍');
    Utilities.sleep(1000);
    
    const readyNPIs = previewAvailableData(hojaEscalation);
    if (readyNPIs.length === 0) {
      showInfoToast('No NPIs are currently ready for escalation.');
      return;
    }

    // PASO 2: Mostrar confirmación con preview
    const confirmed = showConfirmationWithPreview(readyNPIs);
    if (!confirmed) {
      showInfoToast('Process cancelled by user.');
      return;
    }

    // PASO 3: Ejecutar con progreso detallado
    executeProcessWithDetailedProgress(readyNPIs, ss);

  } catch (error) {
    console.error('❌ Error:', error);
    showErrorToast(`An unexpected error occurred: ${error.message}`);
  }
}

// ===== SISTEMA DE MENSAJES TOAST =====
function showProgressToast(message, icon = '⏳') {
  SpreadsheetApp.getActive().toast(message, `${icon} Processing...`, 3);
}

function showSuccessToast(message, icon = '✅') {
  SpreadsheetApp.getActive().toast(message, `${icon} Success`, 4);
}

function showInfoToast(message, icon = 'ℹ️') {
  SpreadsheetApp.getActive().toast(message, `${icon} Information`, 3);
}

function showErrorToast(message, icon = '❌') {
  SpreadsheetApp.getActive().toast(message, `${icon} Error`, 5);
}

// ===== DIÁLOGO DE ERROR CON TEMA AGUA MARINA =====
function showErrorDialog(message) {
  const html = `
    <div style="font-family: 'Google Sans', Arial, sans-serif;">
      <div style="background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%); padding: 30px; text-align: center; color: white;">
        <div style="font-size: 48px; margin-bottom: 15px;">⚠️</div>
        <h2 style="margin: 0; font-weight: 300;">Error Occurred</h2>
      </div>
      
      <div style="padding: 30px;">
        <div style="background: #fef2f2; border-radius: 8px; padding: 20px; margin-bottom: 25px; border-left: 4px solid #ef4444;">
          <p style="margin: 0; color: #374151; line-height: 1.5;">${message}</p>
        </div>

        <div style="text-align: center;">
          <button onclick="google.script.host.close()" 
                  style="background: #dc2626; color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; box-shadow: 0 2px 4px rgba(220, 38, 38, 0.3);">
            Close
          </button>
        </div>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
}

function showInfoDialog(title, message) {
  const html = `
    <div style="font-family: 'Google Sans', Arial, sans-serif;">
      <div style="background: linear-gradient(135deg, #0891b2 0%, #0e7490 100%); padding: 25px; text-align: center; color: white;">
        <div style="font-size: 36px; margin-bottom: 10px;">ℹ️</div>
        <h2 style="margin: 0; font-weight: 300;">${title}</h2>
      </div>
      
      <div style="padding: 25px;">
        <p style="margin: 0 0 25px 0; color: #374151; line-height: 1.5; text-align: center;">${message}</p>
        
        <div style="text-align: center;">
          <button onclick="google.script.host.close()" 
                  style="background: linear-gradient(135deg, #0891b2, #0e7490); color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; box-shadow: 0 2px 4px rgba(8, 145, 178, 0.3);">
            OK
          </button>
        </div>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(350)
    .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Information');
}

// ===== CONFIRMACIÓN CON PREVIEW HTML ELEGANTE =====
function showConfirmationWithPreview(readyNPIs) {
  // Lista compacta de NPIs (solo NPI + Policy/License)
  const npisList = readyNPIs.slice(0, 8).map(npi => 
    `<div style="background: white; padding: 10px; border-radius: 6px; margin-bottom: 6px; border-left: 3px solid #0891b2; font-size: 13px;">
      <div style="font-weight: 500; color: #164e63;">${npi.npi} - ${npi.policyLicense}</div>
      <div style="font-size: 11px; color: #0891b2; margin-top: 2px;">${npi.type} • ${npi.escalationLevel}</div>
    </div>`
  ).join('');
  
  const extraCount = readyNPIs.length > 8 ? 
    `<div style="text-align: center; color: #0891b2; font-size: 13px; margin-top: 10px; font-weight: 500;">... and ${readyNPIs.length - 8} more NPIs</div>` : '';

  const html = `
    <div style="font-family: 'Google Sans', Arial, sans-serif;">
      <div style="background: linear-gradient(135deg, #0891b2 0%, #0e7490 100%); padding: 25px; text-align: center; color: white;">
        <div style="font-size: 32px; margin-bottom: 10px;">🚀</div>
        <h2 style="margin: 0; font-weight: 300; font-size: 22px;">Ready to Update Escalations</h2>
        <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Found ${readyNPIs.length} NPIs ready for escalation</p>
      </div>
      
      <div style="padding: 25px;">
        <h3 style="margin: 0 0 15px 0; color: #164e63; font-size: 16px;">NPIs to be updated:</h3>
        <div style="max-height: 220px; overflow-y: auto; background: #f0fdfa; padding: 15px; border-radius: 10px; border: 1px solid #a7f3d0;">
          ${npisList}
          ${extraCount}
        </div>

        <div style="background: #ecfdf5; border-radius: 10px; padding: 15px; margin: 20px 0; border-left: 4px solid #10b981;">
          <h4 style="margin: 0 0 10px 0; color: #065f46; font-size: 14px;">📋 This process will:</h4>
          <div style="color: #047857; font-size: 13px; line-height: 1.6;">
            ✓ Update escalation status in source sheets<br>
            ✓ Maintain complete outreach history<br>
            ✓ Process everything in ~15-30 seconds
          </div>
        </div>

        <div style="text-align: center; margin-top: 25px;">
          <button onclick="startUpdateAndClose()" 
                  style="background: linear-gradient(135deg, #0891b2, #0e7490); color: white; border: none; padding: 12px 25px; border-radius: 8px; font-size: 16px; cursor: pointer; margin-right: 10px; box-shadow: 0 2px 4px rgba(8, 145, 178, 0.3);">
            Start Update
          </button>
          <button onclick="google.script.host.close()" 
                  style="background: #f1f5f9; color: #475569; border: none; padding: 12px 25px; border-radius: 8px; font-size: 16px; cursor: pointer; border: 1px solid #e2e8f0;">
            Cancel
          </button>
        </div>
      </div>
    </div>

    <script>
      function startUpdateAndClose() {
        google.script.host.close();
        google.script.run.proceedWithUpdate();
      }
    </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(480);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Escalation Preview');
  return true;
}

// ===== EJECUCIÓN CON PROGRESO DETALLADO =====
function executeProcessWithDetailedProgress(readyNPIs, ss) {
  const startTime = new Date();
  
  try {
    // FASE 1: Preparación
    showProgressToast('Initializing data...', '🔄');
    
    const hojaMalpractice = ss.getSheetByName('MALPRACTICE 2025');
    const hojaExpirables = ss.getSheetByName('EXPIRABLES 2025');
    
    showProgressToast('Reading source data...', '📊');
    const malpracticeData = hojaMalpractice ? hojaMalpractice.getDataRange().getValues() : [];
    const expirablesData = hojaExpirables ? hojaExpirables.getDataRange().getValues() : [];
    
    // FASE 2: Actualización de Estados
    showProgressToast(`Updating status for ${readyNPIs.length} NPIs...`, '⚡');
    Utilities.sleep(1000); // Dar tiempo para que el usuario vea el mensaje
    
    const statusStartTime = new Date();
    const updatedCount = batchUpdateStatuses(readyNPIs, malpracticeData, expirablesData, hojaMalpractice, hojaExpirables);
    const statusTime = ((new Date() - statusStartTime) / 1000).toFixed(1);
    
    // MENSAJE 1: Finalización de Estados
    showStatusCompletionDialog(updatedCount, statusTime);
    
    // FASE 3: Actualización de Historial
    showProgressToast('Processing escalation history...', '📋');
    Utilities.sleep(1000);
    
    let hojaHistory = ss.getSheetByName('ESCALATIONS HISTORY');
    if (!hojaHistory) {
      showProgressToast('Creating ESCALATIONS HISTORY sheet...', '📄');
      hojaHistory = createEscalationsHistorySheet(ss);
    }
    
    const historyStartTime = new Date();
    const historyCount = batchUpdateHistoryOptimized(readyNPIs, malpracticeData, expirablesData, hojaHistory);
    const historyTime = ((new Date() - historyStartTime) / 1000).toFixed(1);
    
    // MENSAJE 2: Finalización Completa
    const totalTime = ((new Date() - startTime) / 1000).toFixed(1);
    showFinalCompletionDialog(updatedCount, historyCount, statusTime, historyTime, totalTime);

  } catch (error) {
    showErrorToast(`Process failed: ${error.message}`);
    throw error;
  }
}

// ===== FUNCIÓN CALLBACK PARA PROCEDER CON LA ACTUALIZACIÓN =====
function proceedWithUpdate() {
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaEscalation = ss.getSheetByName('ESCALATION CENTER');
    const readyNPIs = previewAvailableData(hojaEscalation);
    
    if (readyNPIs.length > 0) {
      executeProcessWithDetailedProgress(readyNPIs, ss);
    }
  } catch (error) {
    showErrorToast(`Error starting process: ${error.message}`);
  }
}

// ===== DIÁLOGO DE FINALIZACIÓN DE ESTADOS =====
function showStatusCompletionDialog(updatedCount, timeSeconds) {
  const html = `
    <div style="font-family: 'Google Sans', Arial, sans-serif;">
      <div style="background: linear-gradient(135deg, #0891b2 0%, #0e7490 100%); padding: 25px; text-align: center; color: white;">
        <div style="font-size: 40px; margin-bottom: 12px;">⚡</div>
        <h2 style="margin: 0; font-weight: 300; font-size: 22px;">Phase 1 Complete</h2>
        <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">Status updates finished successfully</p>
      </div>
      
      <div style="padding: 25px;">
        <div style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border-radius: 12px; padding: 20px; text-align: center; margin-bottom: 20px; border: 1px solid #7dd3fc;">
          <div style="font-size: 32px; font-weight: bold; color: #0369a1; margin-bottom: 8px;">${updatedCount}</div>
          <div style="color: #0891b2; font-size: 14px; margin-bottom: 4px; font-weight: 500;">NPIs Updated</div>
          <div style="color: #0e7490; font-size: 12px;">Completed in ${timeSeconds}s</div>
        </div>

        <div style="background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%); border-radius: 10px; padding: 15px; margin-bottom: 20px; text-align: center; border: 1px solid #86efac;">
          <div style="color: #047857; font-weight: 500; font-size: 14px; margin-bottom: 4px;">📋 Phase 2: Processing history...</div>
          <div style="color: #065f46; font-size: 12px;">This will complete automatically in the background</div>
        </div>

        <div style="text-align: center;">
          <button onclick="google.script.host.close()" 
                  style="background: linear-gradient(135deg, #0891b2, #0e7490); color: white; border: none; padding: 10px 25px; border-radius: 6px; font-size: 14px; cursor: pointer; box-shadow: 0 2px 4px rgba(8, 145, 178, 0.3);">
            Continue to Phase 2
          </button>
        </div>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(380);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Phase 1 Complete');
}

// ===== DIÁLOGO DE FINALIZACIÓN COMPLETA =====
function showFinalCompletionDialog(statusCount, historyCount, statusTime, historyTime, totalTime) {
  const html = `
    <div style="font-family: 'Google Sans', Arial, sans-serif;">
      <div style="background: linear-gradient(135deg, #059669 0%, #047857 100%); padding: 30px; text-align: center; color: white;">
        <div style="font-size: 48px; margin-bottom: 15px;">🎉</div>
        <h2 style="margin: 0; font-weight: 300; font-size: 24px;">All Phases Complete!</h2>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">Escalation process finished successfully</p>
      </div>
      
      <div style="padding: 25px;">
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px;">
          <div style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); padding: 15px; border-radius: 10px; text-align: center; border: 1px solid #7dd3fc;">
            <div style="font-size: 24px; font-weight: bold; color: #0369a1; margin-bottom: 4px;">${statusCount}</div>
            <div style="color: #0891b2; font-size: 12px; margin-bottom: 2px; font-weight: 500;">Status Updates</div>
            <div style="color: #0e7490; font-size: 11px;">${statusTime}s</div>
          </div>
          <div style="background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); padding: 15px; border-radius: 10px; text-align: center; border: 1px solid #86efac;">
            <div style="font-size: 24px; font-weight: bold; color: #047857; margin-bottom: 4px;">${historyCount}</div>
            <div style="color: #059669; font-size: 12px; margin-bottom: 2px; font-weight: 500;">History Records</div>
            <div style="color: #047857; font-size: 11px;">${historyTime}s</div>
          </div>
        </div>

        <div style="background: linear-gradient(135deg, #fefce8 0%, #fef3c7 100%); border-radius: 10px; padding: 15px; text-align: center; margin-bottom: 20px; border: 1px solid #fcd34d;">
          <div style="color: #92400e; font-weight: 500; font-size: 16px;">⚡ Total Time: ${totalTime}s</div>
          <div style="color: #a16207; font-size: 12px; margin-top: 4px;">Process completed successfully</div>
        </div>

        <div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border-radius: 8px; padding: 12px; margin-bottom: 20px; border: 1px solid #cbd5e1;">
          <div style="font-size: 13px; color: #475569; text-align: center;">
            ✅ All escalations have been processed and recorded in ESCALATIONS HISTORY
          </div>
        </div>

        <div style="text-align: center;">
          <button onclick="google.script.host.close()" 
                  style="background: linear-gradient(135deg, #059669, #047857); color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 16px; cursor: pointer; box-shadow: 0 3px 6px rgba(5, 150, 105, 0.3);">
            Done
          </button>
        </div>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Complete');
}

// ===== VISTA PREVIA DE DATOS (FUNCIÓN AUXILIAR) =====
function previewAvailableData(hojaEscalation) {
  const escalationData = hojaEscalation.getDataRange().getValues();
  const readyNPIs = [];

  for (let i = 1; i < escalationData.length; i++) {
    const row = escalationData[i];
    const npi = row[EC_NPI_COL] ? row[EC_NPI_COL].toString() : '';
    const policyLicense = row[EC_POLICY_LICENSE_COL] ? row[EC_POLICY_LICENSE_COL].toString() : '';
    const providerName = row[EC_PROVIDER_COL] || '';
    const status = row[EC_STATUS_COL] || '';
    const type = row[EC_TYPE_COL] || '';
    const escalationLevel = row[EC_ESCALATION_LEVEL_COL] || '';

    if (status.includes('READY TO ESCALATE') && npi && policyLicense) {
      readyNPIs.push({
        week: row[EC_WEEK_COL],
        npi: npi,
        policyLicense: policyLicense,
        providerName: providerName,
        type: type,
        escalationLevel: escalationLevel
      });
    }
  }

  return readyNPIs;
}

// ===== FUNCIONES DE PROCESAMIENTO ULTRA RÁPIDAS (MANTIENEN DISEÑO VISUAL) =====
function batchUpdateStatuses(readyNPIs, malpracticeData, expirablesData, hojaMalpractice, hojaExpirables) {
  console.log('🚀 Starting ULTRA FAST status update...');
  const startTime = new Date();
  let totalUpdated = 0;

  const malpracticeNPIs = readyNPIs.filter(npi => npi.type === 'MALPRACTICE');
  const expirablesNPIs = readyNPIs.filter(npi => npi.type === 'EXPIRABLES');

  console.log(`⚡ Processing ${malpracticeNPIs.length} MALPRACTICE and ${expirablesNPIs.length} EXPIRABLES NPIs`);

  // Procesar MALPRACTICE ultra rápido
  if (malpracticeNPIs.length > 0 && hojaMalpractice) {
    showProgressToast(`Ultra-fast processing ${malpracticeNPIs.length} MALPRACTICE NPIs...`, '⚡');
    const malCount = updateSheetUltraFast(malpracticeNPIs, malpracticeData, hojaMalpractice, 'MALPRACTICE');
    totalUpdated += malCount;
    console.log(`✅ MALPRACTICE: ${malCount} updates applied`);
  }

  // Procesar EXPIRABLES ultra rápido
  if (expirablesNPIs.length > 0 && hojaExpirables) {
    showProgressToast(`Ultra-fast processing ${expirablesNPIs.length} EXPIRABLES NPIs...`, '⚡');
    const expCount = updateSheetUltraFast(expirablesNPIs, expirablesData, hojaExpirables, 'EXPIRABLES');
    totalUpdated += expCount;
    console.log(`✅ EXPIRABLES: ${expCount} updates applied`);
  }

  const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  console.log(`🎯 ULTRA FAST update completed: ${totalUpdated} NPIs in ${elapsed}s (was 565s before!)`);

  return totalUpdated;
}

function updateSheetUltraFast(npis, sheetData, sheet, type) {
  console.log(`⚡ Ultra-fast processing ${npis.length} ${type} NPIs...`);
  
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const cols = type === 'MALPRACTICE' ? 
    { npi: MAL_NPI_COL, policy: MAL_POLICY_COL, status: MAL_STATUS_COL, week: MAL_WEEK_COL } :
    { npi: EXP_NPI_COL, policy: EXP_LICENSE_COL, status: EXP_STATUS_COL, week: EXP_WEEK_COL };
  
  // PASO 1: Crear índice de búsqueda súper rápido O(n)
  console.log(`📇 Creating search index for ${type}...`);
  const searchIndex = createSearchIndexUltraFast(sheetData, cols, tz);
  
  // PASO 2: Encontrar todas las actualizaciones O(1) por búsqueda
  console.log(`🎯 Finding updates for ${npis.length} NPIs...`);
  const updates = [];
  
  npis.forEach(npiInfo => {
    const stateMapping = getStateMapping(npiInfo.escalationLevel, type); // Pasar el tipo
    if (!stateMapping) {
      console.log(`❌ No state mapping for escalation level: ${npiInfo.escalationLevel}`);
      return;
    }
    
    const targetWeekStr = (npiInfo.week instanceof Date) ? 
      Utilities.formatDate(npiInfo.week, tz, "MM/dd/yyyy") : npiInfo.week.toString();
    
    // Búsqueda instantánea O(1) - ESTO ES LA MAGIA 🪄
    const searchKey = `${npiInfo.npi}|${npiInfo.policyLicense}|${targetWeekStr}|${stateMapping.currentStatus}`;
    const rowIndex = searchIndex[searchKey];
    
    if (rowIndex !== undefined) {
      updates.push({
        row: rowIndex + 1,
        col: cols.status + 1,
        value: stateMapping.newStatus,
        npi: npiInfo.npi,
        debug: `${type} - ${npiInfo.npi} - ${stateMapping.currentStatus} → ${stateMapping.newStatus}`
      });
      console.log(`✅ Found match: ${updates[updates.length - 1].debug}`);
    } else {
      // DEBUGGING para EXPIRABLES
      if (type === 'EXPIRABLES') {
        console.log(`🔍 EXPIRABLES Debug - NPI: ${npiInfo.npi}, Policy: ${npiInfo.policyLicense}, Week: ${targetWeekStr}, Expected Status: "${stateMapping.currentStatus}"`);
        // Buscar variaciones del estado
        const possibleKeys = Object.keys(searchIndex).filter(key => key.startsWith(`${npiInfo.npi}|${npiInfo.policyLicense}|${targetWeekStr}|`));
        console.log(`🔍 Possible status variations found:`, possibleKeys);
      }
      console.log(`⚠️ No match found for ${npiInfo.npi} in ${type}`);
    }
  });
  
  // PASO 3: Aplicar todas las actualizaciones en batches ultra rápidos
  console.log(`💾 Applying ${updates.length} updates in ultra-fast batches...`);
  if (updates.length > 0) {
    applyUpdatesUltraFast(sheet, updates);
  }
  
  return updates.length;
}

function createSearchIndexUltraFast(sheetData, cols, tz) {
  console.log(`📇 Building ultra-fast search index...`);
  const index = {};
  
  // Crear índice una sola vez - O(n) en lugar de O(n×m)
  for (let i = 1; i < sheetData.length; i++) {
    const npi = sheetData[i][cols.npi] ? sheetData[i][cols.npi].toString() : '';
    const policy = sheetData[i][cols.policy] ? sheetData[i][cols.policy].toString() : '';
    const status = sheetData[i][cols.status] ? sheetData[i][cols.status].toString().trim() : '';
    const week = sheetData[i][cols.week];
    
    if (npi && policy && status) {
      const weekStr = (week instanceof Date) ? 
        Utilities.formatDate(week, tz, "MM/dd/yyyy") : week.toString();
      
      // Clave única para búsqueda instantánea
      const key = `${npi}|${policy}|${weekStr}|${status}`;
      index[key] = i; // Guardar índice de fila para acceso O(1)
    }
  }
  
  console.log(`📊 Search index created with ${Object.keys(index).length} entries - ready for O(1) lookups!`);
  return index;
}

function applyUpdatesUltraFast(sheet, updates) {
  const batchSize = 100; // Procesar en lotes de 100 para máximo rendimiento
  let applied = 0;
  
  console.log(`⚡ Applying ${updates.length} updates in batches of ${batchSize}...`);
  
  for (let i = 0; i < updates.length; i += batchSize) {
    const batch = updates.slice(i, i + batchSize);
    
    // Aplicar batch completo
    batch.forEach(update => {
      try {
        sheet.getRange(update.row, update.col).setValue(update.value);
        console.log(`✅ Applied: ${update.debug}`);
        applied++;
      } catch (error) {
        console.error(`❌ Failed to apply: ${update.debug}`, error);
      }
    });
    
    // Micro pausa solo entre batches para no saturar API
    if (i + batchSize < updates.length) {
      Utilities.sleep(25); // 25ms pausa micro
      showProgressToast(`Applied batch ${Math.floor(i/batchSize) + 1}/${Math.ceil(updates.length/batchSize)}...`, '💾');
    }
  }
  
  console.log(`🎯 Successfully applied ${applied}/${updates.length} updates`);
  return applied;
}

function batchUpdateHistoryOptimized(readyNPIs, malpracticeData, expirablesData, hojaHistory) {
  console.log('📋 Starting optimized history update (no line-by-line formatting)...');
  
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const currentDateTime = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy HH:mm");
  const existingHistory = hojaHistory.getDataRange().getValues();
  
  const newRowsToAdd = [];
  const directUpdates = [];
  let processedCount = 0;

  readyNPIs.forEach((npiInfo, index) => {
    if (index % 25 === 0) {
      showProgressToast(`Processing history ${index + 1}/${readyNPIs.length}...`, '📋');
    }
    
    try {
      const sourceData = npiInfo.type === 'MALPRACTICE' ? malpracticeData : expirablesData;
      const outreachHistory = collectOutreachHistoryUltraOptimized(sourceData, npiInfo.npi, npiInfo.policyLicense, npiInfo.type, tz);

      let existingRowIndex = -1;
      for (let i = 1; i < existingHistory.length; i++) {
        const historyNPI = existingHistory[i][0] ? existingHistory[i][0].toString() : '';
        const historyPolicy = existingHistory[i][1] ? existingHistory[i][1].toString() : '';
        if (historyNPI === npiInfo.npi && historyPolicy === npiInfo.policyLicense) {
          existingRowIndex = i;
          break;
        }
      }

      if (existingRowIndex === -1 && npiInfo.escalationLevel === 'First Escalation') {
        newRowsToAdd.push([
          npiInfo.npi, npiInfo.policyLicense, npiInfo.providerName, npiInfo.type,
          outreachHistory.outreach1Date || '', outreachHistory.outreach2Date || '', 
          outreachHistory.outreach3Date || '', outreachHistory.callBeforePRDate || '',
          currentDateTime, '', ''
        ]);
      } else if (existingRowIndex !== -1) {
        const rowIndex = existingRowIndex + 1;
        if (npiInfo.escalationLevel === 'Second Escalation') {
          directUpdates.push({ row: rowIndex, col: 10, value: currentDateTime });
        } else if (npiInfo.escalationLevel === 'Final Escalation') {
          directUpdates.push({ row: rowIndex, col: 11, value: currentDateTime });
        }
      }
      processedCount++;
    } catch (error) {
      console.error(`❌ Error processing ${npiInfo.npi}:`, error);
    }
  });

  // APLICAR TODAS LAS NUEVAS FILAS DE UNA VEZ (SIN FORMATO LÍNEA POR LÍNEA) 🚀
  if (newRowsToAdd.length > 0) {
    showProgressToast('Writing new records in bulk (no formatting)...', '💾');
    const lastRow = hojaHistory.getLastRow();
    hojaHistory.getRange(lastRow + 1, 1, newRowsToAdd.length, newRowsToAdd[0].length).setValues(newRowsToAdd);
    console.log(`✅ Added ${newRowsToAdd.length} new history records WITHOUT line-by-line formatting (HUGE time save!)`);
  }

  // APLICAR ACTUALIZACIONES DIRECTAS EN LOTE
  if (directUpdates.length > 0) {
    showProgressToast('Applying direct updates in batch...', '🔄');
    directUpdates.forEach(update => {
      hojaHistory.getRange(update.row, update.col).setValue(update.value);
    });
    console.log(`✅ Applied ${directUpdates.length} direct updates`);
  }

  console.log(`📊 History processing completed: ${processedCount} NPIs processed`);
  return processedCount;
}

function collectOutreachHistoryUltraOptimized(sourceData, targetNPI, targetPolicyLicense, type, tz) {
  const history = { outreach1Date: null, outreach2Date: null, outreach3Date: null, callBeforePRDate: null };
  const cols = type === 'MALPRACTICE' ? 
    { npi: MAL_NPI_COL, policy: MAL_POLICY_COL, status: MAL_STATUS_COL, date: MAL_DATE_COL } :
    { npi: EXP_NPI_COL, policy: EXP_LICENSE_COL, status: EXP_STATUS_COL, date: EXP_DATE_COL };

  const statusMapping = OUTREACH_STATUS_MAPPING[type];
  const relevantEntries = [];

  for (let i = 1; i < sourceData.length; i++) {
    const currentNPI = sourceData[i][cols.npi] ? sourceData[i][cols.npi].toString() : '';
    const currentPolicy = sourceData[i][cols.policy] ? sourceData[i][cols.policy].toString() : '';
    const currentStatus = sourceData[i][cols.status] ? sourceData[i][cols.status].toString().trim() : '';
    const currentDate = sourceData[i][cols.date];

    if (currentNPI === targetNPI && currentPolicy === targetPolicyLicense && currentDate instanceof Date) {
      relevantEntries.push({ date: currentDate, status: currentStatus.toUpperCase() });
    }
  }

  relevantEntries.sort((a, b) => a.date.getTime() - b.date.getTime());

  relevantEntries.forEach(entry => {
    const formattedDate = Utilities.formatDate(entry.date, tz, "MM/dd/yyyy HH:mm");
    const status = entry.status;

    if (!history.outreach1Date && statusMapping.outreach1.some(s => status.includes(s.toUpperCase()))) {
      history.outreach1Date = formattedDate;
    } else if (!history.outreach2Date && statusMapping.outreach2.some(s => status.includes(s.toUpperCase()))) {
      history.outreach2Date = formattedDate;
    } else if (!history.outreach3Date && statusMapping.outreach3.some(s => status.includes(s.toUpperCase()))) {
      history.outreach3Date = formattedDate;
    } else if (!history.callBeforePRDate && statusMapping.callBeforePR.some(s => status.includes(s.toUpperCase()))) {
      history.callBeforePRDate = formattedDate;
    }
  });

  return history;
}

// ===== FUNCIONES DE APOYO =====
function getStateMapping(escalationLevel, type = null) {
  switch (escalationLevel) {
    case 'First Escalation': 
      // CORRECCIÓN CRÍTICA: Estados diferentes por tipo
      if (type === 'EXPIRABLES') {
        return { 
          currentStatus: 'To be Escalated to PR',  // EXPIRABLES usa capitalización mixta
          newStatus: 'FIRST ESCALATION TO PR' 
        };
      } else {
        return { 
          currentStatus: 'TO BE ESCALATED TO PR',  // MALPRACTICE usa mayúsculas
          newStatus: 'FIRST ESCALATION TO PR' 
        };
      }
    case 'Second Escalation': 
      return { 
        currentStatus: 'TO BE ESCALATED TO PR #2', 
        newStatus: 'SECOND ESCALATION TO PR' 
      };
    case 'Final Escalation': 
      return { 
        currentStatus: 'TO BE ESCALATED TO PR #3', 
        newStatus: 'FINAL ESCALATION' 
      };
    default: 
      return null;
  }
}

function createEscalationsHistorySheet(spreadsheet) {
  const hoja = spreadsheet.insertSheet('ESCALATIONS HISTORY');
  const headerRange = hoja.getRange(1, 1, 1, HISTORY_HEADERS.length);
  headerRange.setValues([HISTORY_HEADERS]);
  headerRange.setBackground('#006064').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  
  const columnWidths = [120, 180, 250, 100, 140, 140, 140, 150, 170, 170, 150];
  columnWidths.forEach((width, index) => hoja.setColumnWidth(index + 1, width));
  hoja.setFrozenRows(1);
  return hoja;
}

function formatHistoryRow(hoja, rowNumber) {
  [5, 6, 7, 8, 9, 10, 11].forEach(col => {
    const cell = hoja.getRange(rowNumber, col);
    if (cell.getValue()) {
      cell.setNumberFormat('MM/dd/yyyy hh:mm');
    }
  });
  
  hoja.getRange(rowNumber, 1, 1, HISTORY_HEADERS.length)
    .setBorder(true, true, true, true, true, true);
  
  if (rowNumber % 2 === 0) {
    hoja.getRange(rowNumber, 1, 1, HISTORY_HEADERS.length)
      .setBackground('#f8f9fa');
  }
}

// ===== FUNCIÓN DE COMPATIBILIDAD =====
function updateEscalationStatus() {
  updateStatusFromEscalationCenter();
}
