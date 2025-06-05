// ===== ARCHIVO: EmailPreparation_Updated.gs =====
// Sistema de Preparación de Correos para Escalaciones - ACTUALIZADO PARA IPAS AFILIADAS

// ===== CONFIGURACIÓN DE IPAs AFILIADAS =====
const IPA_GROUPS = {
  'AAMG': { name: 'All American Medical Group' },
  'ACPAZ': { name: 'Astranacare Partners Of Arizona' },
  'AHC': { name: 'Accountable Health Care IPA' },
  'AHISP': { name: 'Associated Hispanic Partners' },
  'ALPHA': { name: 'Alpha Care Medical Group' },
  'AMTX': { name: 'Astrana Care Partners Of Texas' },
  'APC': { 
    name: 'Allied Pacific (Grupo de 5)', 
    subgroups: [
      'Allied Pacific of California IPA',
      'Accountable Health Care IPA',
      'Arroyo Vista Family Health Center', 
      'Beverly Alliance IPA',
      'Diamond Bar Medical Group'
    ]
  },
  'APCMG': { name: 'Access Primary Care Medical Group' },
  'AVISTA': { name: 'Arroyo Vista Health Medical Group' },
  'BACP': { name: 'Bay Area Care Partners' },
  'BAIPA': { name: 'Beverly Alliance IPA' },
  'CCPP': { name: 'Central California Physicians Partners' },
  'CFC': { name: 'Community Family Care Medical Group' },
  'CFCHP': { name: 'Community Family Care Health Plan' },
  'CVMG': { name: 'Central Valley Medical Group' },
  'GOM': { name: 'Diamond Bar Medical Group' },
  'HHMG': { name: 'Hana Hou Medical Group' },
  'JADE': { name: 'Jade Health Care Medical Group' },
  'MDPTN': { name: 'MD Partners' },
  'SEEN': { name: 'Seen Health San Gabriel Valley' },
  'FYB': { name: 'For Your Benefit' }
};

// ===== MAPEO FLEXIBLE DE VARIACIONES DE NOMBRES =====
const IPA_NAME_VARIATIONS = {
  // Variaciones de Texas
  'ASTRANA HEALTH OF TEXAS': 'Astrana Care Partners Of Texas',
  'ASTRANA CARE OF TEXAS': 'Astrana Care Partners Of Texas',
  'ASTRANA CARE PARTNERS OF TEXAS': 'Astrana Care Partners Of Texas',
  
  // Variaciones de Arizona
  'ASTRANACARE PARTNERS OF ARIZONA': 'Astranacare Partners Of Arizona',
  'ASTRANA CARE OF ARIZONA': 'Astranacare Partners Of Arizona',
  
  // Otras variaciones comunes
  'COMMUNITY FAMILY CARE IPA': 'Community Family Care Medical Group',
  'ACCOUNTABLE HEALTH CARE': 'Accountable Health Care IPA',
  'ALLIED PACIFIC OF CALIFORNIA': 'Allied Pacific of California IPA',
  'ARROYO VISTA HEALTH MEDICAL GROUP': 'Arroyo Vista Health Medical Group',
  'ARROYO VISTA FAMILY HEALTH CENTER': 'Arroyo Vista Health Medical Group',
  'BEVERLY ALLIANCE': 'Beverly Alliance IPA',
  'BEVERLY ALIANZA IPA': 'Beverly Alliance IPA',
  'DIAMOND BAR MEDICAL GROUP': 'Diamond Bar Medical Group',
  'GREATER ORANGE COUNTY MEDICAL GROUP': 'Diamond Bar Medical Group'
};

// ID del spreadsheet externo con Office Locations
const OFFICE_LOCATIONS_SPREADSHEET_ID = '1G5AZnl01eTF01rba4uvh5cMLQcjnwbEprndXJsIpXoI';
const OFFICE_LOCATIONS_SHEET_NAME = 'DB';

// ===== FUNCIÓN PRINCIPAL =====
function prepararCorreosEscalacion() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Verificar que exista data en ESCALATION CENTER
    const ss = SpreadsheetApp.getActive();
    const escalationSheet = ss.getSheetByName('ESCALATION CENTER');
    
    if (!escalationSheet || escalationSheet.getLastRow() <= 1) {
      ui.alert(
        '⚠️ Warning', 
        'Before preparing emails, you must execute "📊 Process Escalations per Week" first.\n\nNo data found in ESCALATION CENTER.', 
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Mostrar disclaimer antes de continuar
    const response = ui.alert(
      '⚠️ Important Notice',
      'Before preparing emails, make sure you have executed "📊 Process Escalations per Week" first.\n\nThis process requires the escalation data to identify which IPAs are active.\n\nDo you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Obtener IPAs activas antes de mostrar el diálogo
    const ipasActivas = obtenerIPAsAfiliadas();
    
    if (ipasActivas.length === 0) {
      ui.alert(
        '⚠️ Warning',
        'No affiliated IPAs found with READY TO ESCALATE or MISSING PR CONTACT status.\n\nPlease execute "📊 Process Escalations per Week" first.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Crear diálogo de selección HTML con IPAs afiliadas
    const htmlOutput = HtmlService.createHtmlOutput(getSelectionHTML(ipasActivas))
      .setWidth(600)
      .setHeight(500);
    
    ui.showModalDialog(htmlOutput, '📧 Select Affiliated IPAs to Prepare Emails');
    
  } catch (error) {
    console.error('Error:', error);
    ui.alert('❌ Error', 'An error occurred: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===== OBTENER IPAs AFILIADAS CON STATUS VÁLIDOS =====
function obtenerIPAsAfiliadas() {
  const ss = SpreadsheetApp.getActive();
  const escalationSheet = ss.getSheetByName('ESCALATION CENTER');
  const escalationData = escalationSheet.getDataRange().getValues();
  
  const ipasAfiliadas = new Set();
  
  console.log('🔍 Buscando IPAs afiliadas con status válidos...');
  
  // Procesar datos desde fila 2 (saltando header)
  for (let i = 1; i < escalationData.length; i++) {
    const row = escalationData[i];
    const ipasActivasCell = row[5]; // Columna F: IPAs Activas
    const status = row[7]; // Columna H: Status
    
    if (!ipasActivasCell || !status) continue;
    
    // Solo procesar si el status es válido
    const statusStr = status.toString();
    const isValidStatus = statusStr.includes('READY TO ESCALATE') || statusStr.includes('MISSING PR CONTACT');
    
    if (!isValidStatus) {
      continue;
    }
    
    console.log(`✅ Procesando fila ${i + 1} con status: ${statusStr}`);
    
    // Verificar cada IPA configurada
    Object.entries(IPA_GROUPS).forEach(([ipaCode, ipaInfo]) => {
      if (ipaCode === 'APC') {
        // Caso especial: grupo de 5 IPAs
        const tieneAPC = ipaInfo.subgroups.some(subgroup => 
          tieneIPAEnLista(ipasActivasCell.toString(), subgroup)
        );
        if (tieneAPC) {
          ipasAfiliadas.add(ipaCode);
          console.log(`🎯 Encontrada IPA afiliada: ${ipaCode} (grupo APC)`);
        }
      } else {
        // IPAs individuales - usar mapeo flexible
        if (tieneIPAEnLista(ipasActivasCell.toString(), ipaInfo.name)) {
          ipasAfiliadas.add(ipaCode);
          console.log(`🎯 Encontrada IPA afiliada: ${ipaCode} - ${ipaInfo.name}`);
        }
      }
    });
  }
  
  const resultado = Array.from(ipasAfiliadas);
  console.log(`📊 Total IPAs afiliadas encontradas: ${resultado.length}`);
  console.log(`📋 Lista: ${resultado.join(', ')}`);
  
  return resultado;
}

// ===== FUNCIÓN MEJORADA PARA VERIFICAR IPA EN LISTA =====
function tieneIPAEnLista(listaIPAs, nombreIPA) {
  const listaUpper = listaIPAs.toUpperCase();
  const nombreUpper = nombreIPA.toUpperCase();
  
  // Búsqueda directa
  if (listaUpper.includes(nombreUpper)) {
    return true;
  }
  
  // Búsqueda por variaciones
  for (const [variacion, nombreCanonico] of Object.entries(IPA_NAME_VARIATIONS)) {
    if (nombreCanonico.toUpperCase() === nombreUpper) {
      if (listaUpper.includes(variacion)) {
        console.log(`🔄 Mapeo encontrado: "${variacion}" → "${nombreCanonico}"`);
        return true;
      }
    }
  }
  
  // Búsqueda flexible por palabras clave
  const palabrasClave = extraerPalabrasClave(nombreUpper);
  return palabrasClave.length > 0 && palabrasClave.every(palabra => listaUpper.includes(palabra));
}

// ===== EXTRAER PALABRAS CLAVE PARA BÚSQUEDA FLEXIBLE =====
function extraerPalabrasClave(nombre) {
  // Remover palabras comunes y extraer palabras significativas
  const palabrasComunes = ['IPA', 'MEDICAL', 'GROUP', 'HEALTH', 'CARE', 'OF', 'THE'];
  const palabras = nombre.split(/\s+/).filter(palabra => 
    palabra.length > 2 && !palabrasComunes.includes(palabra)
  );
  
  // Para nombres muy específicos, usar palabras más distintivas
  if (nombre.includes('ASTRANA')) {
    if (nombre.includes('TEXAS')) return ['ASTRANA', 'TEXAS'];
    if (nombre.includes('ARIZONA')) return ['ASTRANA', 'ARIZONA'];
  }
  
  return palabras.slice(0, 2); // Máximo 2 palabras clave
}

// ===== HTML PARA SELECCIÓN ACTUALIZADO =====
function getSelectionHTML(ipasAfiliadas) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
          }
          .container {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          h2 {
            color: #006064;
            margin-bottom: 20px;
          }
          .info-banner {
            background-color: #E8F5E8;
            border-left: 4px solid #4CAF50;
            padding: 12px;
            margin-bottom: 20px;
            border-radius: 4px;
          }
          .info-banner strong {
            color: #2E7D32;
          }
          .affiliated-note {
            background-color: #E3F2FD;
            border-left: 4px solid #2196F3;
            padding: 12px;
            margin-bottom: 20px;
            border-radius: 4px;
            font-size: 14px;
          }
          .affiliated-note strong {
            color: #1565C0;
          }
          .ipa-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
            margin-bottom: 20px;
          }
          .ipa-item {
            padding: 10px;
            border: 1px solid #e0e0e0;
            border-radius: 5px;
            cursor: pointer;
            transition: all 0.3s;
          }
          .ipa-item.active {
            background-color: #ffffff;
            border-color: #00ACC1;
          }
          .ipa-item.inactive {
            background-color: #f5f5f5;
            border-color: #e0e0e0;
            opacity: 0.6;
            cursor: not-allowed;
          }
          .ipa-item.active:hover {
            background-color: #E0F7FA;
            border-color: #00ACC1;
          }
          .ipa-item input {
            margin-right: 10px;
          }
          .ipa-item input:disabled {
            cursor: not-allowed;
          }
          .ipa-item label {
            cursor: pointer;
            display: block;
          }
          .ipa-item.inactive label {
            cursor: not-allowed;
          }
          .ipa-code {
            font-weight: bold;
            color: #006064;
          }
          .ipa-code.inactive {
            color: #999;
          }
          .buttons {
            margin-top: 20px;
            text-align: center;
          }
          button {
            padding: 10px 20px;
            margin: 0 10px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s;
          }
          .btn-primary {
            background-color: #00ACC1;
            color: white;
          }
          .btn-primary:hover {
            background-color: #00838F;
          }
          .btn-secondary {
            background-color: #e0e0e0;
            color: #333;
          }
          .btn-secondary:hover {
            background-color: #bdbdbd;
          }
          .select-all {
            margin-bottom: 15px;
            padding: 10px;
            background-color: #E0F7FA;
            border-radius: 5px;
          }
          .status-note {
            font-size: 12px;
            color: #666;
            margin-top: 10px;
            font-style: italic;
          }
          .loading {
            display: none;
            text-align: center;
            margin-top: 20px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #00ACC1;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .progress-text {
            margin-top: 15px;
            color: #006064;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🌊 Select Affiliated IPAs to Prepare Emails</h2>
          
          <div class="info-banner">
            <strong>Affiliated IPAs Found:</strong> ${ipasAfiliadas.length} IPAs with valid escalation status
          </div>
          
          <div class="affiliated-note">
            <strong>ℹ️ Note:</strong> Only showing affiliated IPAs with "READY TO ESCALATE" or "MISSING PR CONTACT" status. Non-affiliated IPAs are automatically excluded.
          </div>
          
          <div class="select-all">
            <label>
              <input type="checkbox" id="selectAll" onchange="toggleAll()">
              <strong>Select All Affiliated IPAs</strong>
            </label>
          </div>
          
          <div class="ipa-grid">
            ${Object.entries(IPA_GROUPS).map(([code, ipa]) => {
              const isActive = ipasAfiliadas.includes(code);
              return `
                <div class="ipa-item ${isActive ? 'active' : 'inactive'}">
                  <label>
                    <input type="checkbox" name="ipa" value="${code}" ${!isActive ? 'disabled' : ''}>
                    <span class="ipa-code ${!isActive ? 'inactive' : ''}">${code}</span> - ${ipa.name}
                    ${!isActive ? '<br><small style="color: #999;">No valid escalations found</small>' : ''}
                  </label>
                </div>
              `;
            }).join('')}
          </div>
          
          <div class="status-note">
            <strong>Note:</strong> Only affiliated IPAs with escalations ready for processing are selectable. All affiliated IPAs from selected NPIs will be included in tables, regardless of individual PR contact status.
          </div>
          
          <div class="buttons">
            <button class="btn-primary" onclick="procesarSeleccion()">
              📧 Generate Tables
            </button>
            <button class="btn-secondary" onclick="google.script.host.close()">
              Cancel
            </button>
          </div>
          
          <div class="loading" id="loading">
            <div class="spinner"></div>
            <div class="progress-text">Generating email tables, please wait...</div>
          </div>
        </div>
        
        <script>
          function toggleAll() {
            const selectAll = document.getElementById('selectAll');
            const checkboxes = document.querySelectorAll('input[name="ipa"]:not(:disabled)');
            checkboxes.forEach(cb => cb.checked = selectAll.checked);
          }
          
          function procesarSeleccion() {
            const checkboxes = document.querySelectorAll('input[name="ipa"]:checked');
            if (checkboxes.length === 0) {
              alert('Please select at least one affiliated IPA');
              return;
            }
            
            const selectedIPAs = Array.from(checkboxes).map(cb => cb.value);
            
            document.getElementById('loading').style.display = 'block';
            document.querySelector('.buttons').style.display = 'none';
            
            google.script.run
              .withSuccessHandler(() => {
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error: ' + error);
                document.getElementById('loading').style.display = 'none';
                document.querySelector('.buttons').style.display = 'block';
              })
              .procesarIPAsSeleccionadas(selectedIPAs);
          }
        </script>
      </body>
    </html>
  `;
}

// ===== PROCESAR IPAs SELECCIONADAS (ACTUALIZADO) =====
function procesarIPAsSeleccionadas(selectedIPAs) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Mostrar toast de inicio
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Generating email tables for ' + selectedIPAs.length + ' affiliated IPAs...', 
      '🔄 Processing', 
      -1
    );
    
    const ss = SpreadsheetApp.getActive();
    const escalationSheet = ss.getSheetByName('ESCALATION CENTER');
    const malpracticeSheet = ss.getSheetByName('MALPRACTICE 2025');
    const expirablesSheet = ss.getSheetByName('EXPIRABLES 2025');
    const ipasDBSheet = ss.getSheetByName('IPAS DB');
    
    // Crear o limpiar hoja de cuadros
    let cuadrosSheet = ss.getSheetByName('EMAILS TO PR');
    if (!cuadrosSheet) {
      cuadrosSheet = ss.insertSheet('EMAILS TO PR');
    } else {
      cuadrosSheet.clear();
    }
    
    // Obtener datos de ESCALATION CENTER
    const escalationData = escalationSheet.getDataRange().getValues();
    
    // Agrupar NPIs por IPA y tipo (SOLO IPAs afiliadas con status válidos)
    const npisPorIPAyTipo = agruparNPIsAfiliados(escalationData, selectedIPAs);
    
    // Toast de progreso
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Loading office locations...', 
      '📍 Processing', 
      3
    );
    
    // Obtener office locations
    const officeLocations = obtenerOfficeLocations();
    
    // Generar cuadros para cada IPA (lado a lado)
    let currentRow = 1;
    let cuadrosGenerados = 0;
    
    selectedIPAs.forEach((ipaCode, index) => {
      const tieneMAL = npisPorIPAyTipo[ipaCode]?.MALPRACTICE?.length > 0;
      const tieneEXP = npisPorIPAyTipo[ipaCode]?.EXPIRABLES?.length > 0;
      
      if (tieneMAL || tieneEXP) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Generating tables ${index + 1}/${selectedIPAs.length}: ${ipaCode}`, 
          '📊 Creating Tables', 
          2
        );
        
        currentRow = generarCuadrosLadoALado(
          cuadrosSheet, 
          ipaCode, 
          npisPorIPAyTipo[ipaCode],
          currentRow,
          malpracticeSheet,
          expirablesSheet,
          ipasDBSheet,
          officeLocations
        );
        currentRow += 3; // Espacio entre grupos de IPAs
        cuadrosGenerados++;
      }
    });
    
    // Toast de formato
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Applying final formatting...', 
      '🎨 Formatting', 
      2
    );
    
    // Aplicar formato final
    aplicarFormatoCuadros(cuadrosSheet);
    
    // Activar la hoja de resultados
    cuadrosSheet.activate();
    
    // Toast de éxito
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${cuadrosGenerados} affiliated IPA groups ready in "EMAILS TO PR" sheet`, 
      '✅ Complete', 
      5
    );
    
    // Mostrar mensaje de éxito
    ui.alert(
      '✅ Tables Generated Successfully',
      `Tables for ${cuadrosGenerados} affiliated IPAs have been generated in the "EMAILS TO PR" sheet.\n\n` +
      'MALPRACTICE tables are on the left, EXPIRABLES tables are on the right.\n\n' +
      'All affiliated IPAs are included regardless of individual PR contact status.\n\n' +
      'The tables are ready to copy and paste into emails.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error processing affiliated IPAs:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error: ' + error.toString(), 
      '❌ Error', 
      10
    );
    ui.alert('❌ Error', 'An error occurred while processing: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===== AGRUPAR NPIs POR IPA Y TIPO (SOLO AFILIADAS CON STATUS VÁLIDOS) =====
function agruparNPIsAfiliados(escalationData, selectedIPAs) {
  const grupos = {};
  
  // Inicializar grupos
  selectedIPAs.forEach(ipa => {
    grupos[ipa] = {
      MALPRACTICE: [],
      EXPIRABLES: []
    };
  });
  
  console.log('🔍 Agrupando NPIs de IPAs afiliadas seleccionadas...');
  
  // Procesar datos desde fila 2 (saltando header)
  for (let i = 1; i < escalationData.length; i++) {
    const row = escalationData[i];
    const ipasActivas = row[5]; // Columna F: IPAs Activas
    const npi = row[2]; // Columna C: NPI
    const tipo = row[10]; // Columna K: Tipo
    const status = row[7]; // Columna H: Status
    
    if (!ipasActivas || !npi || !tipo || !status) continue;
    
    // Solo procesar si el status es válido
    const statusStr = status.toString();
    const isValidStatus = statusStr.includes('READY TO ESCALATE') || statusStr.includes('MISSING PR CONTACT');
    
    if (!isValidStatus) {
      console.log(`⏭️ Saltando NPI ${npi} - status no válido: ${statusStr}`);
      continue;
    }
    
    // Verificar cada IPA seleccionada
    selectedIPAs.forEach(ipaCode => {
      const ipaInfo = IPA_GROUPS[ipaCode];
      
      let perteneceAIPA = false;
      
      if (ipaCode === 'APC') {
        // Caso especial: grupo de 5 IPAs
        perteneceAIPA = ipaInfo.subgroups.some(subgroup => 
          tieneIPAEnLista(ipasActivas.toString(), subgroup)
        );
      } else {
        // IPAs individuales - usar mapeo flexible
        perteneceAIPA = tieneIPAEnLista(ipasActivas.toString(), ipaInfo.name);
      }
      
      if (perteneceAIPA) {
        const npiData = {
          npi: npi,
          rowIndex: i
        };
        
        if (tipo === 'MALPRACTICE') {
          grupos[ipaCode].MALPRACTICE.push(npiData);
          console.log(`✅ Agregado NPI ${npi} a ${ipaCode} MALPRACTICE`);
        } else if (tipo === 'EXPIRABLES') {
          grupos[ipaCode].EXPIRABLES.push(npiData);
          console.log(`✅ Agregado NPI ${npi} a ${ipaCode} EXPIRABLES`);
        }
      }
    });
  }
  
  // Log de resultados
  selectedIPAs.forEach(ipaCode => {
    const malCount = grupos[ipaCode].MALPRACTICE.length;
    const expCount = grupos[ipaCode].EXPIRABLES.length;
    console.log(`📊 ${ipaCode}: ${malCount} MALPRACTICE, ${expCount} EXPIRABLES`);
  });
  
  return grupos;
}

// ===== GENERAR CUADROS LADO A LADO (sin cambios) =====
function generarCuadrosLadoALado(sheet, ipaCode, npisPorTipo, startRow, malpracticeSheet, expirablesSheet, ipasDBSheet, officeLocations) {
  const ipaInfo = IPA_GROUPS[ipaCode];
  const separacionColumnas = 2; // Espacio entre las dos tablas
  
  // Título general para la IPA
  sheet.getRange(startRow, 1, 1, 17).merge()
    .setValue(`📧 ${ipaCode} - ${ipaInfo.name}`)
    .setBackground('#006064')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');
  
  startRow += 2;
  
  const npisMalpractice = npisPorTipo.MALPRACTICE || [];
  const npisExpirables = npisPorTipo.EXPIRABLES || [];
  
  // Subtítulos
  if (npisMalpractice.length > 0) {
    sheet.getRange(startRow, 1, 1, 8).merge()
      .setValue('MALPRACTICE')
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  if (npisExpirables.length > 0) {
    sheet.getRange(startRow, 10, 1, 8).merge()
      .setValue('EXPIRABLES')
      .setBackground('#FF9800')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  startRow++;
  
  // Headers MALPRACTICE
  if (npisMalpractice.length > 0) {
    const headersMal = [
      'Provider',
      'NPI',
      'Groups/Office Locations',
      'Email',
      'Malpractice Insurance Carrier',
      'Malpractice Ins. Carrier Code',
      'Policy Number',
      'Expiration Date'
    ];
    
    sheet.getRange(startRow, 1, 1, headersMal.length)
      .setValues([headersMal])
      .setBackground('#00ACC1')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  // Headers EXPIRABLES
  if (npisExpirables.length > 0) {
    const headersExp = [
      'Provider',
      'NPI',
      'Groups/Office Locations',
      'Email',
      'License/Credential Type',
      'Institution',
      'License Number / Internal ID',
      'Newest Expiration Date'
    ];
    
    sheet.getRange(startRow, 10, 1, headersExp.length)
      .setValues([headersExp])
      .setBackground('#00ACC1')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  startRow++;
  
  // Obtener datos de ambas hojas
  const malpracticeData = malpracticeSheet.getDataRange().getValues();
  const expirablesData = expirablesSheet.getDataRange().getValues();
  const ipasData = ipasDBSheet.getDataRange().getValues();
  
  // Crear mapa de emails desde IPAS DB
  const emailMap = {};
  for (let i = 1; i < ipasData.length; i++) {
    const npi = ipasData[i][3]; // Columna D
    const email = ipasData[i][2]; // Columna C
    if (npi) emailMap[npi] = email;
  }
  
  // Datos MALPRACTICE
  let malStartRow = startRow;
  if (npisMalpractice.length > 0) {
    npisMalpractice.forEach(npiInfo => {
      let providerData = null;
      for (let i = 1; i < malpracticeData.length; i++) {
        if (malpracticeData[i][4] == npiInfo.npi) { // Columna E: NPI
          providerData = malpracticeData[i];
          break;
        }
      }
      
      if (providerData) {
        const rowData = [
          providerData[5] || '', // F: Provider Name
          npiInfo.npi,
          officeLocations[npiInfo.npi] || 'Pending',
          emailMap[npiInfo.npi] || '',
          providerData[8] || '', // I: Insurance Carrier
          providerData[9] || '', // J: Carrier Code
          providerData[10] || '', // K: Policy Number
          providerData[12] ? formatDate(providerData[12]) : '' // M: Expiration Date
        ];
        
        sheet.getRange(malStartRow, 1, 1, rowData.length)
          .setValues([rowData])
          .setBackground('#FFFFFF');
        
        malStartRow++;
      }
    });
  }
  
  // Datos EXPIRABLES
  let expStartRow = startRow;
  if (npisExpirables.length > 0) {
    npisExpirables.forEach(npiInfo => {
      let providerData = null;
      for (let i = 1; i < expirablesData.length; i++) {
        if (expirablesData[i][3] == npiInfo.npi) { // Columna D: NPI
          providerData = expirablesData[i];
          break;
        }
      }
      
      if (providerData) {
        const rowData = [
          providerData[4] || '', // E: Provider Name
          npiInfo.npi,
          officeLocations[npiInfo.npi] || 'Pending',
          emailMap[npiInfo.npi] || '',
          providerData[5] || '', // F: License/Credential Type
          providerData[6] || '', // G: Institution
          providerData[7] || '', // H: License Number
          providerData[9] ? formatDate(providerData[9]) : '' // J: Expiration Date
        ];
        
        sheet.getRange(expStartRow, 10, 1, rowData.length)
          .setValues([rowData])
          .setBackground('#FFFFFF');
        
        expStartRow++;
      }
    });
  }
  
  // Aplicar bordes a las tablas
  const maxRows = Math.max(malStartRow - startRow + 1, expStartRow - startRow + 1);
  
  if (npisMalpractice.length > 0) {
    const malTableRange = sheet.getRange(startRow - 1, 1, maxRows, 8);
    malTableRange.setBorder(true, true, true, true, true, true);
  }
  
  if (npisExpirables.length > 0) {
    const expTableRange = sheet.getRange(startRow - 1, 10, maxRows, 8);
    expTableRange.setBorder(true, true, true, true, true, true);
  }
  
  return Math.max(malStartRow, expStartRow);
}

// ===== OBTENER OFFICE LOCATIONS (sin cambios) =====
function obtenerOfficeLocations() {
  try {
    console.log('Intentando conectar con spreadsheet externo...');
    const externalSS = SpreadsheetApp.openById(OFFICE_LOCATIONS_SPREADSHEET_ID);
    const sheet = externalSS.getSheetByName(OFFICE_LOCATIONS_SHEET_NAME);
    
    if (!sheet) {
      console.error('No se encontró la hoja:', OFFICE_LOCATIONS_SHEET_NAME);
      return {};
    }
    
    const data = sheet.getDataRange().getValues();
    console.log('Filas obtenidas de office locations:', data.length);
    
    const locations = {};
    
    // Procesar datos desde fila 2 (saltando header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const npi = row[3]; // Columna D: NPI
      const officeName = row[4] || ''; // Columna E: Office name
      const address = row[5] || ''; // Columna F: Address
      const city = row[6] || ''; // Columna G: City
      const state = row[7] || ''; // Columna H: State
      const zipcode = row[8] || ''; // Columna I: Zipcode
      const phone = row[9] || ''; // Columna J: Phone number
      
      if (npi) {
        // Si el NPI ya existe, agregar nueva ubicación
        if (locations[npi]) {
          locations[npi] += '\n'; // Separador entre ubicaciones
        } else {
          locations[npi] = '';
        }
        
        // Formato: Office name + Address + City + State + Zipcode + Phone
        const locationInfo = [
          officeName,
          address,
          city,
          state,
          zipcode,
          phone
        ].filter(item => item.toString().trim() !== '').join('\n');
        
        locations[npi] += locationInfo;
      }
    }
    
    console.log('NPIs procesados para office locations:', Object.keys(locations).length);
    return locations;
    
  } catch (error) {
    console.error('Error obteniendo office locations:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Could not load office locations from external database. Showing as "Pending"', 
      '⚠️ Warning', 
      5
    );
    return {};
  }
}

// ===== APLICAR FORMATO A CUADROS (AJUSTADO PARA LADO A LADO) =====
function aplicarFormatoCuadros(sheet) {
  // Ajustar anchos de columna para MALPRACTICE
  sheet.setColumnWidth(1, 200); // Provider
  sheet.setColumnWidth(2, 120); // NPI
  sheet.setColumnWidth(3, 250); // Groups/Office Locations
  sheet.setColumnWidth(4, 200); // Email
  sheet.setColumnWidth(5, 200); // Malpractice Insurance Carrier
  sheet.setColumnWidth(6, 150); // Carrier Code
  sheet.setColumnWidth(7, 150); // Policy Number
  sheet.setColumnWidth(8, 120); // Expiration Date
  
  // Ajustar anchos de columna para EXPIRABLES
  sheet.setColumnWidth(10, 200); // Provider
  sheet.setColumnWidth(11, 120); // NPI
  sheet.setColumnWidth(12, 250); // Groups/Office Locations
  sheet.setColumnWidth(13, 200); // Email
  sheet.setColumnWidth(14, 200); // License/Credential Type
  sheet.setColumnWidth(15, 200); // Institution
  sheet.setColumnWidth(16, 150); // License Number / Internal ID
  sheet.setColumnWidth(17, 120); // Newest Expiration Date
  
  // Congelar primera fila si hay datos
  if (sheet.getLastRow() > 0) {
    sheet.setFrozenRows(1);
  }
  
  // Aplicar wrap text a todas las celdas con contenido
  if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
      .setWrap(true)
      .setVerticalAlignment('top');
  }
}

// ===== UTILIDADES =====
function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  }
  return date;
}
