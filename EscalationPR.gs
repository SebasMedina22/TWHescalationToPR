// ===== INICIO DEL CÓDIGO =====
// ===== ARCHIVO: EscalationPR_Step1_Complete.gs =====
// Sistema de Escalación PR - PASO 1: Poblar ESCALATION CENTER (Complete with IPA Affiliation & PR Contacts)
// Function: Load NPIs according to selected escalation type and show active IPAs with affiliation and contact status

// ===== UPDATED CONSTANTS FOR COLUMN INDICES =====
// MALPRACTICE 2025
const MALPRACTICE_WEEK_COL = 2; // Column C
const MALPRACTICE_NPI_COL = 4; // Column E
const MALPRACTICE_PROVIDER_COL = 5; // Column F
const MALPRACTICE_POLICY_COL = 10; // Column K
const MALPRACTICE_ISSUE_DATE_COL = 11; // Column L - Malpractice Ins. Current Issued Date
const MALPRACTICE_STATUS_COL = 14; // Column O

// EXPIRABLES 2025
const EXPIRABLES_WEEK_COL = 1; // Column B
const EXPIRABLES_NPI_COL = 3; // Column D
const EXPIRABLES_PROVIDER_COL = 4; // Column E
const EXPIRABLES_LICENSE_COL = 7; // Column H - License Number / Internal ID
const EXPIRABLES_ISSUE_DATE_COL = 8; // Column I - Issue Date
const EXPIRABLES_STATUS_COL = 12; // Column M

// IPAS DB
const IPAS_DB_FIRST_NAME_COL = 0; // Column A
const IPAS_DB_LAST_NAME_COL = 1; // Column B
const IPAS_DB_EMAIL_COL = 2; // Column C
const IPAS_DB_NPI_COL = 3; // Column D
const IPAS_DB_FACILITIES_START_COL = 4; // Column E

// ESCALATION CENTER Headers - UPDATED WITH NEW COLUMNS
const ESCALATION_CENTER_HEADERS = [
  '#',                      // A - Order number
  'Week',                   // B - Week processed
  'NPI',                    // C - NPI
  'Policy/License Number',  // D - Policy/license number
  'Provider Name',          // E - Provider name
  'Active IPAS',           // F - Active IPAs
  '#IPAS activas',         // G - Number of active IPAs
  'Status',                // H - Status
  'Astrana status',        // I - Astrana status
  'Date of process',       // J - Process date
  'Type',                  // K - Type (MALPRACTICE/EXPIRABLES)
  'Escalation Level',      // L - Escalation type (First/Second/Final)
  'Non-Affiliated IPAs',   // M - IPAs not affiliated with Astrana
  'Missing PR Contact'     // N - IPAs without PR contact
];

// ===== IPA AFFILIATION DATABASE =====
const IPA_AFFILIATION_DB = new Map([
  ['American Acupuncture Chinese Medicine', 'NO'],
  ['All American Medical Group', 'YES'],
  ['AstranaCare Partners Of Arizona', 'YES'],
  ['Advantage Health Network', 'NO'],
  ['Accountable Health Care IPA', 'YES'],
  ['Associated Hispanic Partners', 'YES'],
  ['Alpha Care Medical Group', 'YES'],
  ['Apollo Care Partners Of Nevada', 'NO'],
  ['Astrana Care Partners Of Texas', 'YES'],
  ['Allied Pacific of California IPA', 'YES'],
  ['Access Primary Care Medical Group', 'YES'],
  ['Arroyo Vista Health Medical Group', 'YES'],
  ['Bay Area Care Partners', 'YES'],
  ['Beverly Alianza IPA', 'YES'],
  ['Caipa MSO LLC', 'NO'],
  ['Central California Physicians Partners', 'YES'],
  ['Community Family Care IPA', 'YES'],
  ['Community Family Care Health Plan', 'YES'],
  ['Chesapeake Independet Physicians Association', 'NO'],
  ['Connecticut State Medical Sociaty IPA', 'NO'],
  ['Emanate Health IPA', 'NO'],
  ['Central Valley Medical Group', 'YES'],
  ['Diamond Bar Medical Group', 'YES'],
  ['Greater San Gabriel Valley Physicians', 'NO'],
  ['Golden Triangle Physicians Alliance', 'NO'],
  ['Hana Hou Medical Group', 'YES'],
  ['Heritage Physicians Networks', 'NO'],
  ['Jade Health IPA', 'YES'],
  ['La Salle Medical Associates IPA', 'NO'],
  ['MD Partners', 'YES'],
  ['Northern California Physicians Network', 'NO'],
  ['Provider Health Link LLC', 'NO'],
  ['Seen Health San Gabriel Valley', 'YES'],
  ['For Your Benefit', 'YES'],
  // Astrana specific entities
  ['Astrana Care of Arizona', 'YES'],
  ['Astrana Care of Nevada', 'YES'],
  ['Astrana Care of Texas', 'YES']
]);

// ===== PR CONTACTS DATABASE =====
const PR_CONTACTS_DB = new Set([
  'Alpha Care Medical Group',
  'Allied Pacific of California IPA', 
  'Accountable Health Care IPA',
  'Arroyo Vista Health Medical Group',
  'Beverly Alianza IPA',
  'Diamond Bar Medical Group',
  'Community Family Care IPA',
  'Bay Area Care Partners',
  'Central California Physicians Partners',
  'Jade Health IPA',
  'All American Medical Group',
  'American Acupuncture Chinese Medicine',
  'Access Primary Care Medical Group',
  'Associated Hispanic Partners',
  'Central Valley Medical Group',
  'Astrana Care of Arizona',
  'Astrana Care of Nevada', 
  'Astrana Care of Texas',
  'For Your Benefit',
  'Hana Hou Medical Group'
]);

// Global variable to track progress
let progressToast = null;

// ===== MAIN FUNCTION - Shows HTML dialog to select escalation level =====
function showEscalationLevelSelection() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('EscalationLevel')
      .setWidth(650)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Escalation Level');
}

// ===== NEW FUNCTION CALLED FROM HTML WITH SEPARATED PARAMETERS =====
function processEscalationsWithLevelAndParams(escalationLevel, typeSelection, dateInput) {
  try {
    console.log('🚀 Starting escalation process with parameters:', {
      escalationLevel,
      typeSelection,
      dateInput
    });
    
    // Determine filter states based on escalation level
    let statusFilterMalpractice = '';
    let statusFilterExpirables = '';
    let escalationTypeToRecord = '';

    if (escalationLevel === 'new') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR';
      statusFilterExpirables = 'To be Escalated to PR';
      escalationTypeToRecord = 'First Escalation';
    } else if (escalationLevel === 'second') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR #2';
      statusFilterExpirables = 'TO BE ESCALATED TO PR #2';
      escalationTypeToRecord = 'Second Escalation';
    } else if (escalationLevel === 'final') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR #3';
      statusFilterExpirables = 'TO BE ESCALATED TO PR #3';
      escalationTypeToRecord = 'Final Escalation';
    } else {
      throw new Error('Invalid escalation level selected.');
    }

    // Determine processing types
    let procesarMalpractice = false;
    let procesarExpirables = false;
    
    if (typeSelection === 'malpractice') {
      procesarMalpractice = true;
    } else if (typeSelection === 'expirables') {
      procesarExpirables = true;
    } else if (typeSelection === 'both') {
      procesarMalpractice = true;
      procesarExpirables = true;
    } else {
      throw new Error('Invalid type selection.');
    }
    
    // Validate date format
    const dateRegex = /^\d{4}\/\d{2}\/\d{2}$/;
    if (!dateInput || !dateRegex.test(dateInput)) {
      throw new Error('Invalid date format. Please use YYYY/MM/DD format.');
    }
    
    // Convert date format for week filter
    const parts = dateInput.split('/');
    const weekFormat = `${parts[1]}/${parts[2]}/${parts[0]}`;
    
    console.log('📊 Parameters validated successfully');
    
    // Execute main escalation process
    const startTime = new Date();
    const results = ejecutarProcesoEscalacionWithProgress(
      weekFormat, 
      procesarMalpractice, 
      procesarExpirables, 
      statusFilterMalpractice, 
      statusFilterExpirables,
      escalationTypeToRecord
    );
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);

    // VALIDAR QUE LOS RESULTADOS NO ESTÉN UNDEFINED Y ASEGURAR VALORES
    const malpracticeStats = results.malpractice || {};
    const expirablesStats = results.expirables || {};
    
    console.log('🔍 Validating results:', { malpracticeStats, expirablesStats });

    const successData = {
      success: true,
      duration: duration,
      dateProcessed: dateInput,
      weekFormat: weekFormat,
      types: procesarMalpractice && procesarExpirables ? 'MALPRACTICE + EXPIRABLES' : 
             procesarMalpractice ? 'MALPRACTICE' : 'EXPIRABLES',
      escalationLevel: escalationTypeToRecord,
      results: {
        malpractice: {
          count: malpracticeStats.count || 0,
          readyToEscalate: malpracticeStats.readyToEscalate || 0,
          requiresReview: malpracticeStats.requiresReview || 0,
          missingPRContact: malpracticeStats.missingPRContact || 0,
          notAffiliated: malpracticeStats.notAffiliated || 0,
          astranaInactive: malpracticeStats.astranaInactive || 0,
          noActiveIPAs: malpracticeStats.noActiveIPAs || 0,
          notFound: malpracticeStats.notFound || 0,
          duplicates: malpracticeStats.duplicates || 0,
          fallbacks: malpracticeStats.fallbacks || 0
        },
        expirables: {
          count: expirablesStats.count || 0,
          readyToEscalate: expirablesStats.readyToEscalate || 0,
          requiresReview: expirablesStats.requiresReview || 0,
          missingPRContact: expirablesStats.missingPRContact || 0,
          notAffiliated: expirablesStats.notAffiliated || 0,
          astranaInactive: expirablesStats.astranaInactive || 0,
          noActiveIPAs: expirablesStats.noActiveIPAs || 0,
          notFound: expirablesStats.notFound || 0,
          duplicates: expirablesStats.duplicates || 0,
          fallbacks: expirablesStats.fallbacks || 0
        }
      }
    };

    console.log('✅ Process completed successfully');
    console.log('📊 Final validated results:', successData.results);
    
    return successData;
    
  } catch (error) {
    console.error('❌ Error in processEscalationsWithLevelAndParams:', error);
    throw {
      name: error.name || 'ProcessingError',
      message: error.message || 'An unknown error occurred',
      stack: error.stack || 'No stack trace available'
    };
  }
}

// ===== FUNCTION TO UPDATE HTML PROGRESS =====
function updateHTMLProgress(percentage, message, details) {
  try {
    // This function will be called by the HTML via google.script.run
    console.log(`Progress Update: ${percentage}% - ${message}`);
    // Note: Direct HTML updates from server-side are not possible in Google Apps Script
    // Progress updates will be handled through the simulation in HTML
  } catch (error) {
    console.log(`Progress update failed: ${error.message}`);
  }
}

// ===== MAIN PROCESS WITH PROGRESS UPDATES =====
function ejecutarProcesoEscalacionWithProgress(weekFilter, procesarMalpractice, procesarExpirables, statusFilterMalpractice, statusFilterExpirables, escalationTypeToRecord) {
  try {
    console.log('📂 Step 1/7: Loading spreadsheet data...');
    const ss = SpreadsheetApp.getActive();
    
    // Get required sheets
    const hojaMalpractice = procesarMalpractice ? ss.getSheetByName('MALPRACTICE 2025') : null;
    const hojaExpirables = procesarExpirables ? ss.getSheetByName('EXPIRABLES 2025') : null;
    const hojaIPASDB = ss.getSheetByName('IPAS DB');
    let hojaEscalation = ss.getSheetByName('ESCALATION CENTER');

    // Verify sheets exist
    if (procesarMalpractice && !hojaMalpractice) {
      throw new Error('MALPRACTICE 2025 sheet not found.');
    }
    if (procesarExpirables && !hojaExpirables) {
      throw new Error('EXPIRABLES 2025 sheet not found.');
    }
    if (!hojaIPASDB) {
      throw new Error('IPAS DB sheet not found.');
    }
    
    console.log('🏗️ Step 2/7: Preparing ESCALATION CENTER...');
    
    // Create ESCALATION CENTER if it doesn't exist and clear if it exists
    if (!hojaEscalation) {
      hojaEscalation = crearHojaEscalationCenter(ss);
    } else {
      // Clear existing content (except headers)
      const lastRow = hojaEscalation.getLastRow();
      if (lastRow > 1) {
        const range = hojaEscalation.getRange(2, 1, lastRow - 1, hojaEscalation.getLastColumn());
        range.clear();
      }
    }
    
    console.log('✅ Sheets loaded successfully');
    console.log('🔍 Step 3/7: Searching for NPIs...');
    
    // STEP 1: Get NPIs from both sheets with unique key detection
    let npisMalpractice = [];
    let npisExpirables = [];
    
    if (procesarMalpractice) {
      npisMalpractice = obtenerNPIsMalpracticeWithUniqueKeys(hojaMalpractice, weekFilter, statusFilterMalpractice);
      console.log(`✅ NPIs found in MALPRACTICE: ${npisMalpractice.length}`);
    }
    
    if (procesarExpirables) {
      npisExpirables = obtenerNPIsExpirablesWithUniqueKeys(hojaExpirables, weekFilter, statusFilterExpirables);
      console.log(`✅ NPIs found in EXPIRABLES: ${npisExpirables.length}`);
    }
    
    const totalNPIs = npisMalpractice.length + npisExpirables.length;
    if (totalNPIs === 0) {
      throw new Error(`No NPIs found with status "${statusFilterMalpractice}" or "${statusFilterExpirables}" for week ${weekFilter}.`);
    }
    
    console.log(`📊 Step 4/7: Loading IPAS database for ${totalNPIs} NPIs...`);
    
    // STEP 2: Load IPAS DB
    const ipasDatabase = cargarBaseDatosIPAS(hojaIPASDB);
    console.log(`✅ IPAS DB loaded: ${Object.keys(ipasDatabase).length} records`);
    
    console.log('⚙️ Step 5/7: Processing NPIs and analyzing IPAs...');
    
    // STEP 3: Process NPIs from both types with duplicate detection and IPA analysis
    const resultadosMalpractice = [];
    const resultadosExpirables = [];
    
    // Process MALPRACTICE with duplicate handling and IPA affiliation
    if (procesarMalpractice && npisMalpractice.length > 0) {
      const processedMalpractice = procesarNPIsConDuplicados(npisMalpractice, ipasDatabase, 'MALPRACTICE', escalationTypeToRecord);
      resultadosMalpractice.push(...processedMalpractice);
      console.log(`✅ Processed ${resultadosMalpractice.length} MALPRACTICE NPIs (duplicates and affiliation handled)`);
    }
    
    // Process EXPIRABLES with duplicate handling and IPA affiliation
    if (procesarExpirables && npisExpirables.length > 0) {
      const processedExpirables = procesarNPIsConDuplicados(npisExpirables, ipasDatabase, 'EXPIRABLES', escalationTypeToRecord);
      resultadosExpirables.push(...processedExpirables);
      console.log(`✅ Processed ${resultadosExpirables.length} EXPIRABLES NPIs (duplicates and affiliation handled)`);
    }
    
    console.log('📝 Step 6/7: Writing results to ESCALATION CENTER...');
    
    // STEP 4: Combine and write results to ESCALATION CENTER
    const todosLosResultados = [...resultadosMalpractice, ...resultadosExpirables];
    escribirResultadosOptimized(todosLosResultados, hojaEscalation);
    
    console.log('🎨 Step 7/7: Applying formatting...');
    
    // STEP 5: Apply formatting
    aplicarFormatoBonitoOptimized(hojaEscalation);
    
    console.log('✅ Process completed successfully');
    
    // Return separate statistics
    return {
      malpractice: procesarMalpractice ?
        calcularEstadisticasCompletas(resultadosMalpractice) : 
        { count: 0, readyToEscalate: 0, requiresReview: 0, missingPRContact: 0, notAffiliated: 0, astranaInactive: 0, noActiveIPAs: 0, notFound: 0, duplicates: 0, fallbacks: 0 },
      expirables: procesarExpirables ?
        calcularEstadisticasCompletas(resultadosExpirables) : 
        { count: 0, readyToEscalate: 0, requiresReview: 0, missingPRContact: 0, notAffiliated: 0, astranaInactive: 0, noActiveIPAs: 0, notFound: 0, duplicates: 0, fallbacks: 0 }
    };
    
  } catch (error) {
    console.error('❌ Error in ejecutarProcesoEscalacionWithProgress:', error);
    throw new Error(`Processing failed: ${error.message}`);
  }
}

// ===== NEW: GET EXPIRABLES NPIs WITH UNIQUE KEYS - UPDATED =====
function obtenerNPIsExpirablesWithUniqueKeys(hoja, weekFilter, statusFilter) {
  try {
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    
    // Get all data in one operation
    const datos = hoja.getDataRange().getValues();
    const npisEncontrados = [];
    
    console.log(`🔍 Searching EXPIRABLES NPIs for week: ${weekFilter} with status: '${statusFilter}'`);
    
    // Process all rows in memory
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      
      const fecha = fila[EXPIRABLES_WEEK_COL]; 
      const npi = fila[EXPIRABLES_NPI_COL]; 
      const provider = fila[EXPIRABLES_PROVIDER_COL]; 
      const licenseNumber = fila[EXPIRABLES_LICENSE_COL];
      const issueDate = fila[EXPIRABLES_ISSUE_DATE_COL];
      const status = fila[EXPIRABLES_STATUS_COL]; 
      
      // Skip empty rows early
      if (!npi || !fecha) continue;
      
      const fechaStr = (fecha instanceof Date) ? 
        Utilities.formatDate(fecha, tz, "MM/dd/yyyy") : "";
      
      const estadoNormalizado = status ? status.toString().trim() : '';
      
      if (fechaStr === weekFilter && estadoNormalizado === statusFilter) {
        // Create unique key hierarchy
        let uniqueKey = '';
        let policyLicenseDisplay = '';
        let usedFallback = false;
        
        if (licenseNumber && licenseNumber.toString().trim() !== '') {
          // Option 1: NPI + License Number
          uniqueKey = `${npi.toString()}_${licenseNumber.toString().trim()}`;
          policyLicenseDisplay = licenseNumber.toString().trim();
        } else if (issueDate) {
          // Option 2: NPI + Issue Date
          const issueDateStr = (issueDate instanceof Date) ? 
            Utilities.formatDate(issueDate, tz, "MM/dd/yyyy") : issueDate.toString();
          uniqueKey = `${npi.toString()}_${issueDateStr}`;
          policyLicenseDisplay = `No license number - Issue date: ${issueDateStr}`;
          usedFallback = true;
        } else {
          // Fallback: Just NPI
          uniqueKey = npi.toString();
          policyLicenseDisplay = 'No license number - No issue date';
          usedFallback = true;
        }
        
        npisEncontrados.push({
          npi: npi.toString(),
          providerName: provider || '',
          week: fecha,
          rowIndex: i + 1,
          policyLicenseNumber: policyLicenseDisplay,
          uniqueKey: uniqueKey,
          usedFallback: usedFallback,
          tipo: 'EXPIRABLES'
        });
      }
    }
    
    console.log(`✅ Found ${npisEncontrados.length} EXPIRABLES NPIs`);
    return npisEncontrados;
    
  } catch (error) {
    throw new Error(`Error retrieving EXPIRABLES NPIs: ${error.message}`);
  }
}

// ===== NEW: GET MALPRACTICE NPIs WITH UNIQUE KEYS - UPDATED =====
function obtenerNPIsMalpracticeWithUniqueKeys(hoja, weekFilter, statusFilter) {
  try {
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    
    // Get all data in one operation
    const datos = hoja.getDataRange().getValues();
    const npisEncontrados = [];
    
    console.log(`🔍 Searching MALPRACTICE NPIs for week: ${weekFilter} with status: '${statusFilter}'`);
    
    // Process all rows in memory
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      
      const fecha = fila[MALPRACTICE_WEEK_COL]; 
      const npi = fila[MALPRACTICE_NPI_COL]; 
      const provider = fila[MALPRACTICE_PROVIDER_COL]; 
      const policyNumber = fila[MALPRACTICE_POLICY_COL];
      const issueDate = fila[MALPRACTICE_ISSUE_DATE_COL];
      const status = fila[MALPRACTICE_STATUS_COL]; 
      
      // Skip empty rows early
      if (!npi || !fecha) continue;
      
      const fechaStr = (fecha instanceof Date) ? 
        Utilities.formatDate(fecha, tz, "MM/dd/yyyy") : "";
      
      const estadoNormalizado = status ? status.toString().trim() : '';
      
      if (fechaStr === weekFilter && estadoNormalizado === statusFilter) {
        // Create unique key hierarchy
        let uniqueKey = '';
        let policyLicenseDisplay = '';
        let usedFallback = false;
        
        if (policyNumber && policyNumber.toString().trim() !== '') {
          // Option 1: NPI + Policy Number
          uniqueKey = `${npi.toString()}_${policyNumber.toString().trim()}`;
          policyLicenseDisplay = policyNumber.toString().trim();
        } else if (issueDate) {
          // Option 2: NPI + Issue Date
          const issueDateStr = (issueDate instanceof Date) ? 
            Utilities.formatDate(issueDate, tz, "MM/dd/yyyy") : issueDate.toString();
          uniqueKey = `${npi.toString()}_${issueDateStr}`;
          policyLicenseDisplay = `No policy number - Issue date: ${issueDateStr}`;
          usedFallback = true;
        } else {
          // Fallback: Just NPI
          uniqueKey = npi.toString();
          policyLicenseDisplay = 'No policy number - No issue date';
          usedFallback = true;
        }
        
        npisEncontrados.push({
          npi: npi.toString(),
          providerName: provider || '',
          week: fecha,
          rowIndex: i + 1,
          policyLicenseNumber: policyLicenseDisplay,
          uniqueKey: uniqueKey,
          usedFallback: usedFallback,
          tipo: 'MALPRACTICE'
        });
      }
    }
    
    console.log(`✅ Found ${npisEncontrados.length} MALPRACTICE NPIs`);
    return npisEncontrados;
    
  } catch (error) {
    throw new Error(`Error retrieving MALPRACTICE NPIs: ${error.message}`);
  }
}

// ===== NEW: PROCESS NPIs WITH DUPLICATE DETECTION =====
function procesarNPIsConDuplicados(npisArray, ipasDatabase, tipo, escalationLevel) {
  const resultados = [];
  const uniqueKeysFound = new Set();
  const duplicateKeys = new Set();
  
  // First pass: Identify duplicates
  npisArray.forEach(npiInfo => {
    if (uniqueKeysFound.has(npiInfo.uniqueKey)) {
      duplicateKeys.add(npiInfo.uniqueKey);
    } else {
      uniqueKeysFound.add(npiInfo.uniqueKey);
    }
  });
  
  console.log(`📋 Found ${duplicateKeys.size} duplicate unique keys in ${tipo}`);
  
  // Second pass: Process NPIs, keeping only first occurrence of duplicates
  const processedKeys = new Set();
  
  npisArray.forEach((npiInfo, index) => {
    let shouldProcess = true;
    let orderNumber = index + 1;
    
    // Check if this is a duplicate
    if (duplicateKeys.has(npiInfo.uniqueKey)) {
      if (processedKeys.has(npiInfo.uniqueKey)) {
        // Skip this duplicate
        shouldProcess = false;
      } else {
        // First occurrence of duplicate - mark it
        orderNumber = `${index + 1}. Duplicated`;
        processedKeys.add(npiInfo.uniqueKey);
      }
    }
    
    if (shouldProcess) {
      const npiData = ipasDatabase[npiInfo.npi];
      const resultado = procesarNPIWithAffiliation(npiInfo, npiData, tipo);
      resultado.escalationLevel = escalationLevel;
      resultado.orderNumber = orderNumber;
      resultado.isDuplicate = duplicateKeys.has(npiInfo.uniqueKey);
      resultados.push(resultado);
    }
  });
  
  console.log(`✅ Processed ${resultados.length} unique ${tipo} NPIs (${npisArray.length - resultados.length} duplicates filtered)`);
  return resultados;
}

// ===== LOAD IPAS DATABASE - OPTIMIZED =====
function cargarBaseDatosIPAS(hoja) {
  try {
    // Get all data in one operation
    const datos = hoja.getDataRange().getValues();
    const database = {};
    
    // Use for loop for better performance
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const npi = fila[IPAS_DB_NPI_COL]; 
      
      if (npi) {
        const npiStr = npi.toString();
        database[npiStr] = {
          firstName: fila[IPAS_DB_FIRST_NAME_COL] || '', 
          lastName: fila[IPAS_DB_LAST_NAME_COL] || '', 
          email: fila[IPAS_DB_EMAIL_COL] || '', 
          facilities: []
        };
        
        // Process facilities (columns E-T, every 2 columns)
        for (let col = IPAS_DB_FACILITIES_START_COL; col < 20; col += 2) {
          const facility = fila[col];
          const status = fila[col + 1];
          
          if (facility && facility.toString().trim() !== '') {
            database[npiStr].facilities.push({
              name: facility.toString().trim(),
              status: status ? status.toString().trim() : '',
              columnIndex: col,
              facilityNumber: Math.floor((col - IPAS_DB_FACILITIES_START_COL) / 2) + 1
            });
          }
        }
      }
    }
    
    return database;
    
  } catch (error) {
    throw new Error(`Error loading IPAS database: ${error.message}`);
  }
}

// ===== NEW: PROCESS NPI WITH AFFILIATION AND PR CONTACT ANALYSIS - CORRECTED LOGIC =====
function procesarNPIWithAffiliation(npiInfo, datosNPI, tipo) {
  const resultado = {
    ...npiInfo,
    activeIPAs: [],
    inactiveIPAs: [],
    nonAffiliatedIPAs: [],
    missingContactIPAs: [],
    astranaStatus: 'N/A',
    statusFinal: '',
    statusEmoji: '',
    processedDate: new Date(),
    tipo: tipo || npiInfo.tipo || 'UNKNOWN'
  };
  
  if (!datosNPI) {
    resultado.statusFinal = 'NO_ENCONTRADO';
    resultado.statusEmoji = '⚫ NOT FOUND IN DB';
    return resultado;
  }
  
  // Check Astrana status (first facility)
  if (datosNPI.facilities.length > 0) {
    const primeraFacility = datosNPI.facilities[0];
    resultado.astranaStatus = primeraFacility.status || 'N/A';
  }
  
  // Process other facilities (from second onwards) with affiliation and contact analysis
  const affiliatedActiveIPAs = [];
  let hasAtLeastOneAffiliatedIPA = false;
  
  for (let i = 1; i < datosNPI.facilities.length; i++) {
    const facility = datosNPI.facilities[i];
    
    if (facility.status === 'Active' || facility.status === 'Applicant') {
      const ipaName = identificarIPAConAstranaCare(facility.name);
      const isAffiliated = checkIPAAffiliation(ipaName);
      const hasContact = checkPRContact(ipaName);
      
      console.log(`🔍 Processing IPA: ${ipaName}, Affiliated: ${isAffiliated}, Has Contact: ${hasContact}`);
      
      resultado.activeIPAs.push({
        name: ipaName,
        facility: facility.name,
        status: facility.status,
        isAffiliated: isAffiliated,
        hasContact: hasContact
      });
      
      // LÓGICA CORREGIDA: Separar por tipo de IPA
      if (!isAffiliated) {
        // COLUMNA M: IPAs NO afiliadas
        resultado.nonAffiliatedIPAs.push(ipaName);
        console.log(`➡️ Added to Non-Affiliated (Column M): ${ipaName}`);
      } else {
        // Es una IPA afiliada
        hasAtLeastOneAffiliatedIPA = true;
        affiliatedActiveIPAs.push({
          name: ipaName,
          hasContact: hasContact
        });
        console.log(`✅ Affiliated IPA with contact status: ${ipaName} - Has Contact: ${hasContact}`);
      }
      
  // COLUMNA N: TODAS las IPAs activas que NO tienen contacto PR (afiliadas o no)
        if (!hasContact) {
          resultado.missingContactIPAs.push(ipaName);
          console.log(`➡️ Added to Missing Contact (Column N): ${ipaName}`);
        }
      } else if (facility.status === 'Inactive') {
        resultado.inactiveIPAs.push({
          name: identificarIPAConAstranaCare(facility.name),
          facility: facility.name,
          status: facility.status
        });
      }
    }  // ← ESTA LLAVE CIERRA EL LOOP FOR
    
    console.log(`📊 Summary for NPI ${npiInfo.npi}:`);
    console.log(`  - Has at least one affiliated IPA: ${hasAtLeastOneAffiliatedIPA}`);
    console.log(`  - Affiliated IPAs with contact: ${affiliatedActiveIPAs.filter(ipa => ipa.hasContact).length}`);
    console.log(`  - Non-affiliated IPAs (Column M): ${resultado.nonAffiliatedIPAs.length}`);
    console.log(`  - Missing contact IPAs (Column N): ${resultado.missingContactIPAs.length}`);
    
    // Determinar status final basado en análisis corregido
    resultado.statusFinal = determinarStatusFinalCorregido(resultado, affiliatedActiveIPAs);
    resultado.statusEmoji = obtenerEmojiStatus(resultado.statusFinal);
    
    return resultado;
  }

// ===== CORRECTED: CHECK IPA AFFILIATION =====
function checkIPAAffiliation(ipaName) {
  console.log(`🔍 Checking affiliation for IPA: "${ipaName}"`);
  
  // First check direct match in affiliation database
  if (IPA_AFFILIATION_DB.has(ipaName)) {
    const result = IPA_AFFILIATION_DB.get(ipaName) === 'YES';
    console.log(`✅ Direct match - Affiliated: ${result}`);
    return result;
  }
  
  // If in PR contacts but not in affiliation list, consider as affiliated
  if (PR_CONTACTS_DB.has(ipaName)) {
    console.log(`✅ Found in PR contacts - considering as affiliated`);
    return true;
  }
  
  // Check for partial matches (flexible matching)
  for (const [affiliatedIPA, status] of IPA_AFFILIATION_DB) {
    if (status === 'YES' && (
      ipaName.toUpperCase().includes(affiliatedIPA.toUpperCase()) ||
      affiliatedIPA.toUpperCase().includes(ipaName.toUpperCase())
    )) {
      console.log(`✅ Partial match found: "${affiliatedIPA}"`);
      return true;
    }
  }
  
  console.log(`❌ Not affiliated: "${ipaName}"`);
  return false; // Default to not affiliated
}

// ===== NEW: CHECK PR CONTACT =====
function checkPRContact(ipaName) {
  console.log(`🔍 Checking PR contact for IPA: "${ipaName}"`);
  
  // Direct match
  if (PR_CONTACTS_DB.has(ipaName)) {
    console.log(`✅ Direct match found in PR_CONTACTS_DB`);
    return true;
  }
  
  // Check for partial matches (flexible matching)
  for (const contactIPA of PR_CONTACTS_DB) {
    if (ipaName.toUpperCase().includes(contactIPA.toUpperCase()) ||
        contactIPA.toUpperCase().includes(ipaName.toUpperCase())) {
      console.log(`✅ Partial match found: "${contactIPA}"`);
      return true;
    }
  }
  
  console.log(`❌ No PR contact found for: "${ipaName}"`);
  return false;
}

// ===== CORRECTED: DETERMINE FINAL STATUS =====
function determinarStatusFinalCorregido(resultado, affiliatedActiveIPAs) {
  console.log(`🔍 Determinando status para NPI: ${resultado.npi}`);
  console.log(`  - Active IPAs: ${resultado.activeIPAs.length}`);
  console.log(`  - Affiliated Active IPAs: ${affiliatedActiveIPAs.length}`);
  console.log(`  - Astrana Status: ${resultado.astranaStatus}`);
  
  // Priority 1: Not found in DB
  if (resultado.statusFinal === 'NO_ENCONTRADO') {
    console.log(`  ➡️ Status: NO_ENCONTRADO (not in database)`);
    return 'NO_ENCONTRADO';
  }
  
  // Priority 2: No active IPAs at all
  if (resultado.activeIPAs.length === 0) {
    if (resultado.astranaStatus === 'Inactive') {
      console.log(`  ➡️ Status: ASTRANA_INACTIVE (no IPAs + Astrana inactive)`);
      return 'ASTRANA_INACTIVE';
    }
    console.log(`  ➡️ Status: SIN_IPAS (no active IPAs)`);
    return 'SIN_IPAS';
  }
  
  // Priority 3: Check Astrana status first for cases with IPAs
  if (resultado.astranaStatus === 'Inactive') {
    console.log(`  ➡️ Status: ASTRANA_INACTIVE (Astrana inactive with IPAs)`);
    return 'ASTRANA_INACTIVE';
  }
  
  // Priority 4: LÓGICA CLAVE - Si tiene al menos una IPA afiliada = READY TO ESCALATE
  if (affiliatedActiveIPAs.length > 0) {
    // Si tiene al menos una IPA afiliada, SIEMPRE es READY TO ESCALATE
    // No importa si hay otras no afiliadas o sin contacto
    console.log(`  ➡️ Status: LISTO (has ${affiliatedActiveIPAs.length} affiliated IPA(s))`);
    return 'LISTO';
  }
  
  // Priority 5: Solo tiene IPAs no afiliadas
  if (affiliatedActiveIPAs.length === 0 && resultado.activeIPAs.length > 0) {
    console.log(`  ➡️ Status: NOT_AFFILIATED (only non-affiliated IPAs)`);
    return 'NOT_AFFILIATED';
  }
  
  console.log(`  ➡️ Status: SIN_IPAS (fallback)`);
  return 'SIN_IPAS';
}

// ===== NEW: GET STATUS EMOJI =====
function obtenerEmojiStatus(statusFinal) {
  const statusMap = {
    'LISTO': '🟢 READY TO ESCALATE',
    'REVISION': '🟡 REQUIRES REVIEW',
    'MISSING_PR_CONTACT': '🟠 MISSING PR CONTACT',
    'NOT_AFFILIATED': '🔴 NOT AFFILIATED IPA',
    'ASTRANA_INACTIVE': '🟠 ASTRANA INACTIVE',
    'SIN_IPAS': '🔴 NO ACTIVE IPAs',
    'NO_ENCONTRADO': '⚫ NOT FOUND IN DB'
  };
  
  return statusMap[statusFinal] || '❓ UNKNOWN STATUS';
}

// ===== UPDATED: IDENTIFY IPA WITH ASTRANA CARE CORRECTION =====
function identificarIPAConAstranaCare(facilityName) {
  if (!facilityName) return 'Unknown';
  
  const upper = facilityName.toUpperCase();
  
  // First check for specific Astrana Care mappings
  if (upper.includes('ASTRANA HEALTH OF TEXAS')) {
    return 'Astrana Care of Texas';
  }
  if (upper.includes('ASTRANACARE PARTNERS OF ARIZONA')) {
    return 'Astrana Care of Arizona';
  }
  if (upper.includes('ASTRANACARE PARTNERS OF NEVADA')) {
    return 'Astrana Care of Nevada';
  }
  
  // Use Map for faster lookups for other IPAs
  const mappings = new Map([
    ['COMMUNITY FAMILY', 'Community Family Care IPA'],
    ['ACCOUNTABLE', 'Accountable Health Care IPA'],
    ['ALPHA', 'Alpha Care Medical Group'],
    ['ALLIED PACIFIC', 'Allied Pacific of California IPA'],
    ['DIAMOND BAR', 'Diamond Bar Medical Group'],
    ['ARROYO', 'Arroyo Vista Health Medical Group'],
    ['JADE', 'Jade Health IPA'],
    ['HISPANIC', 'Associated Hispanic Partners'],
    ['CENTRAL VALLEY', 'Central Valley Medical Group'],
    ['ACCESS', 'Access Primary Care Medical Group'],
    ['BAY AREA', 'Bay Area Care Partners'],
    ['BEVERLY', 'Beverly Alianza IPA'],
    ['ALL AMERICAN', 'All American Medical Group'],
    ['LA SALLE', 'La Salle Medical Associates IPA']
    // Note: Removed generic ASTRANA mapping to handle specific cases above
  ]);
  
  for (const [key, value] of mappings) {
    if (upper.includes(key)) return value;
  }
  
  return facilityName;
}

// ===== UPDATED: WRITE RESULTS TO ESCALATION CENTER WITH NEW COLUMNS =====
function escribirResultadosOptimized(datos, hoja) {
  try {
    if (datos.length === 0) return;
    
    // Prepare all rows in memory first
    const filas = [];
    
    for (let index = 0; index < datos.length; index++) {
      const item = datos[index];
      
      // Format active IPAs list
      const listaIPAs = item.activeIPAs.map(ipa => 
        `• ${ipa.name}\n  ✓ ${ipa.status}`
      ).join('\n\n');
      
      // Format non-affiliated IPAs list (Column M)
      const nonAffiliatedList = item.nonAffiliatedIPAs && item.nonAffiliatedIPAs.length > 0 ?
        item.nonAffiliatedIPAs.map(ipa => `• ${ipa}`).join('\n') : '';
      
      // Format missing contact IPAs list (Column N)
      const missingContactList = item.missingContactIPAs && item.missingContactIPAs.length > 0 ?
        item.missingContactIPAs.map(ipa => `• ${ipa}`).join('\n') : '';
      
      // Use the custom order number (handles duplicates)
      const orderNumber = item.orderNumber || (index + 1);

      filas.push([
        orderNumber,                  // # - A (can be "X. Duplicated")
        item.week,                    // Week - B
        item.npi,                     // NPI - C
        item.policyLicenseNumber,     // Policy/License Number - D (includes fallback messages)
        item.providerName,            // Provider Name - E
        listaIPAs || '---',          // Active IPAS - F
        item.activeIPAs.length,       // #IPAS activas - G
        item.statusEmoji,             // Status - H
        item.astranaStatus,           // Astrana status - I
        item.processedDate,           // Date of process - J
        item.tipo,                    // Type - K
        item.escalationLevel || '',   // Escalation Level - L
        nonAffiliatedList,            // Non-Affiliated IPAs - M
        missingContactList            // Missing PR Contact - N
      ]);
    }
    
    // Write all rows in a single operation
    hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
    console.log(`✅ Written ${filas.length} rows to ESCALATION CENTER`);
    
  } catch (error) {
    throw new Error(`Error writing results: ${error.message}`);
  }
}

// ===== UPDATED: CREATE ESCALATION CENTER SHEET WITH NEW COLUMNS =====
function crearHojaEscalationCenter(spreadsheet) {
  try {
    const hoja = spreadsheet.insertSheet('ESCALATION CENTER');
    
    const headers = ESCALATION_CENTER_HEADERS;
    const headerRange = hoja.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Header format
    headerRange.setBackground('#006064');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(12);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Updated column widths to accommodate new columns
    const columnWidths = [80, 100, 120, 250, 250, 400, 80, 220, 120, 140, 120, 150, 200, 200];
    columnWidths.forEach((width, index) => {
      hoja.setColumnWidth(index + 1, width);
    });
    
    hoja.setFrozenRows(1);
    return hoja;
    
  } catch (error) {
    throw new Error(`Error creating ESCALATION CENTER sheet: ${error.message}`);
  }
}

// ===== UPDATED: APPLY VISUAL FORMAT WITH NEW STATUS COLORS =====
function aplicarFormatoBonitoOptimized(hoja) {
  try {
    const lastRow = hoja.getLastRow();
    if (lastRow <= 1) return;

    const COLORS = {
      READY: '#00ACC1',
      REVIEW: '#FFB300',
      MISSING_CONTACT: '#FF9800',
      NOT_AFFILIATED: '#FF5252',
      INACTIVE: '#FF9800',
      NO_IPAS: '#FF5252',
      NO_DATA: '#9E9E9E',
      MALPRACTICE_BG: '#E8F5E9',
      EXPIRABLES_BG: '#FFF3E0',
      DUPLICATE_BG: '#FFEBEE',
      FALLBACK_BG: '#FFF9C4',
      NON_AFFILIATED_BG: '#FFCDD2',
      MISSING_PR_BG: '#FFE0B2'
    };
    
    // Get all data at once for batch processing
    const dataRange = hoja.getRange(1, 1, lastRow, hoja.getLastColumn());
    const allData = dataRange.getValues();
    
    // Prepare batch formatting operations
    const typeFormats = [];
    const escalationFormats = [];
    const statusFormats = [];
    const duplicateFormats = [];
    const fallbackFormats = [];
    const nonAffiliatedFormats = [];
    const missingContactFormats = [];
    
    for (let i = 1; i < lastRow; i++) { // Start from row 1 (index 1) to skip headers
      const rowData = allData[i];
      const orderNumber = rowData[0]; // # - A (index 0)
      const policyLicense = rowData[3]; // Policy/License Number - D (index 3)
      const tipo = rowData[10]; // Type - K (index 10)
      const escalationLevel = rowData[11]; // Escalation Level - L (index 11)
      const status = rowData[7]; // Status - H (index 7)
      const nonAffiliated = rowData[12]; // Non-Affiliated IPAs - M (index 12)
      const missingContact = rowData[13]; // Missing PR Contact - N (index 13)
      
      // Check for duplicates (order number contains "Duplicated")
      if (orderNumber && orderNumber.toString().includes('Duplicated')) {
        duplicateFormats.push({
          row: i + 1,
          background: COLORS.DUPLICATE_BG,
          fontColor: '#C62828'
        });
      }
      
      // Check for fallback policy/license (contains "No policy" or "No license")
      if (policyLicense && (policyLicense.toString().includes('No policy') || 
                           policyLicense.toString().includes('No license'))) {
        fallbackFormats.push({
          row: i + 1,
          background: COLORS.FALLBACK_BG,
          fontColor: '#F57F17'
        });
      }
      
      // Non-affiliated IPAs formatting (Column M)
      if (nonAffiliated && nonAffiliated.toString().trim() !== '') {
        nonAffiliatedFormats.push({
          row: i + 1,
          background: COLORS.NON_AFFILIATED_BG,
          fontColor: '#C62828'
        });
      }
      
      // Missing contact IPAs formatting (Column N)
      if (missingContact && missingContact.toString().trim() !== '') {
        missingContactFormats.push({
          row: i + 1,
          background: COLORS.MISSING_PR_BG,
          fontColor: '#E65100'
        });
      }
      
      // Type formatting
      if (tipo === 'MALPRACTICE') {
        typeFormats.push({
          row: i + 1,
          background: COLORS.MALPRACTICE_BG,
          fontColor: '#2E7D32'
        });
      } else if (tipo === 'EXPIRABLES') {
        typeFormats.push({
          row: i + 1,
          background: COLORS.EXPIRABLES_BG,
          fontColor: '#E65100'
        });
      }
      
      // Escalation Level formatting
      if (escalationLevel) {
        let bgColor = '#FFFFFF';
        let fontColor = '#000000';
        
        if (escalationLevel.includes('First')) {
          bgColor = '#E3F2FD';
          fontColor = '#0D47A1';
        } else if (escalationLevel.includes('Second')) {
          bgColor = '#FFF3E0';
          fontColor = '#E65100';
        } else if (escalationLevel.includes('Final')) {
          bgColor = '#FFEBEE';
          fontColor = '#C62828';
        }
        
        escalationFormats.push({
          row: i + 1,
          background: bgColor,
          fontColor: fontColor
        });
      }
      
      // Status formatting with new statuses
      if (status) {
        let bgColor = null;
        
        if (status.includes('READY')) {
          bgColor = COLORS.READY;
        } else if (status.includes('REVIEW')) {
          bgColor = COLORS.REVIEW;
        } else if (status.includes('MISSING PR CONTACT')) {
          bgColor = COLORS.MISSING_CONTACT;
        } else if (status.includes('NOT AFFILIATED')) {
          bgColor = COLORS.NOT_AFFILIATED;
        } else if (status.includes('ASTRANA INACTIVE')) {
          bgColor = COLORS.INACTIVE;
        } else if (status.includes('NO ACTIVE IPAs')) {
          bgColor = COLORS.NO_IPAS;
        } else if (status.includes('NOT FOUND')) {
          bgColor = COLORS.NO_DATA;
        }
        
        if (bgColor) {
          statusFormats.push({
            row: i + 1,
            background: bgColor
          });
        }
      }
    }
    
    // Apply formatting in batches
    
    // Order number column formatting for duplicates (A)
    duplicateFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 1);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontWeight('bold');
    });
    
    // Policy/License column formatting for fallbacks (D)
    fallbackFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 4);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontStyle('italic');
    });
    
    // Non-Affiliated IPAs column formatting (M)
    nonAffiliatedFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 13);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontWeight('bold');
    });
    
    // Missing Contact IPAs column formatting (N)
    missingContactFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 14);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontWeight('bold');
    });
    
    // Type column formatting (K)
    typeFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 11);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontWeight('bold');
    });
    
    // Escalation Level column formatting (L)
    escalationFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 12);
      cell.setBackground(format.background);
      cell.setFontColor(format.fontColor);
      cell.setFontWeight('bold');
      cell.setHorizontalAlignment('center');
    });
    
    // Status column formatting (H)
    statusFormats.forEach(format => {
      const cell = hoja.getRange(format.row, 8);
      cell.setBackground(format.background);
      cell.setFontColor('white');
      cell.setFontWeight('bold');
      cell.setHorizontalAlignment('center');
    });
    
    // Apply other formatting in batch operations
    // Active IPAS, Non-Affiliated, and Missing Contact columns - text wrapping
    const textWrapColumns = [6, 13, 14]; // F, M, N
    textWrapColumns.forEach(col => {
      const range = hoja.getRange(2, col, lastRow - 1, 1);
      range.setWrap(true);
      range.setVerticalAlignment('top');
    });
    
    // Center content for specific columns
    const centerColumns = [1, 3, 7, 9, 11, 12]; // #, NPI, #IPAS, Astrana, Type, Escalation
    centerColumns.forEach(col => {
      hoja.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('center');
    });
    
    // Date format for Date of process (J)
    hoja.getRange(2, 10, lastRow - 1, 1).setNumberFormat('MM/dd/yyyy hh:mm');
    
    // Apply borders to entire data range
    dataRange.setBorder(true, true, true, true, true, true);
    
    console.log(`✅ Applied formatting to ${lastRow} rows with affiliation and contact indicators`);
    
  } catch (error) {
    throw new Error(`Error applying formatting: ${error.message}`);
  }
}

// ===== UPDATED: CALCULATE COMPREHENSIVE STATISTICS - CORRECTED NAMES =====
function calcularEstadisticasCompletas(resultados) {
  const stats = {
    count: resultados.length,
    readyToEscalate: resultados.filter(r => r.statusFinal === 'LISTO').length,
    requiresReview: resultados.filter(r => r.statusFinal === 'REVISION').length,
    missingPRContact: resultados.filter(r => r.statusFinal === 'MISSING_PR_CONTACT').length,
    notAffiliated: resultados.filter(r => r.statusFinal === 'NOT_AFFILIATED').length,
    astranaInactive: resultados.filter(r => r.statusFinal === 'ASTRANA_INACTIVE').length,
    noActiveIPAs: resultados.filter(r => r.statusFinal === 'SIN_IPAS').length,
    notFound: resultados.filter(r => r.statusFinal === 'NO_ENCONTRADO').length,
    duplicates: resultados.filter(r => r.isDuplicate).length,
    fallbacks: resultados.filter(r => r.usedFallback).length
  };
  
  // AGREGAR LOG PARA DEBUG
  console.log(`📊 Estadísticas detalladas para ${resultados.length} resultados:`);
  console.log(`  - Ready to escalate (LISTO): ${stats.readyToEscalate}`);
  console.log(`  - Requires review (REVISION): ${stats.requiresReview}`);
  console.log(`  - Missing PR contact: ${stats.missingPRContact}`);
  console.log(`  - Not affiliated: ${stats.notAffiliated}`);
  console.log(`  - Astrana inactive: ${stats.astranaInactive}`);
  console.log(`  - No active IPAs: ${stats.noActiveIPAs}`);
  console.log(`  - Not found: ${stats.notFound}`);
  
  return stats;
}

// ===== KEEP OLD FUNCTION FOR COMPATIBILITY =====
function processEscalationsWithLevel(escalationLevel) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    showProgress('🔄 Initializing escalation process...');
    
    // Determine filter states based on escalation level
    let statusFilterMalpractice = '';
    let statusFilterExpirables = '';
    let escalationTypeToRecord = '';

    if (escalationLevel === 'new') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR';
      statusFilterExpirables = 'To be Escalated to PR';
      escalationTypeToRecord = 'First Escalation';
    } else if (escalationLevel === 'second') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR #2';
      statusFilterExpirables = 'TO BE ESCALATED TO PR #2';
      escalationTypeToRecord = 'Second Escalation';
    } else if (escalationLevel === 'final') {
      statusFilterMalpractice = 'TO BE ESCALATED TO PR #3';
      statusFilterExpirables = 'TO BE ESCALATED TO PR #3';
      escalationTypeToRecord = 'Final Escalation';
    } else {
      hideProgress();
      ui.alert('❌ Error', 'Invalid escalation level.', ui.ButtonSet.OK);
      return;
    }

    showProgress('📋 Validating escalation parameters...');
    
    // Request process type (MALPRACTICE, EXPIRABLES or both)
    const tipoResponse = ui.alert(
      '🌊 Select Escalation Type',
      'Please select the type of escalation to process:\n\n' +
      '• Yes = MALPRACTICE 2025\n' +
      '• No = EXPIRABLES 2025\n' +
      '• Cancel = Both types',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    let procesarMalpractice = false;
    let procesarExpirables = false;
    
    if (tipoResponse === ui.Button.YES) {
      procesarMalpractice = true;
    } else if (tipoResponse === ui.Button.NO) {
      procesarExpirables = true;
    } else if (tipoResponse === ui.Button.CANCEL) {
      procesarMalpractice = true;
      procesarExpirables = true;
    } else {
      hideProgress();
      return;
    }
    
    showProgress('📅 Processing date input...');
    
    // Request date from user
    const response = ui.prompt(
      '📅 Enter Date',
      'Please enter the date to process in YYYY/MM/DD format (e.g.: 2025/05/19):',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      hideProgress();
      return;
    }
    
    const dateInput = response.getResponseText().trim();
    
    // Validate date format
    const dateRegex = /^\d{4}\/\d{2}\/\d{2}$/;
    if (!dateInput || !dateRegex.test(dateInput)) {
      hideProgress();
      ui.alert('❌ Error', 'You must enter a valid date in YYYY/MM/DD format.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert date format for week filter
    const parts = dateInput.split('/');
    const weekFormat = `${parts[1]}/${parts[2]}/${parts[0]}`;
    
    showProgress('🚀 Starting escalation analysis...');
    
    // Execute main escalation process - CORRECTED VARIABLE NAME
    const startTime = new Date();
    const results = ejecutarProcesoEscalacionWithProgress(
      weekFormat, 
      procesarMalpractice, 
      procesarExpirables, 
      statusFilterMalpractice, 
      statusFilterExpirables,
      escalationTypeToRecord
    );
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);

    hideProgress();

    // Show final results with comprehensive statistics
    const tiposText = procesarMalpractice && procesarExpirables ?
      'MALPRACTICE + EXPIRABLES' : 
      procesarMalpractice ?
      'MALPRACTICE' : 'EXPIRABLES';

    let resultadosTexto = `⏱️ Processing time: ${duration} seconds\n\n`;
    resultadosTexto += `📊 Processing details for ${dateInput}\n`;
    resultadosTexto += `📅 Week: ${weekFormat} | Type: ${tiposText}\n`;
    resultadosTexto += `🚀 Level: ${escalationTypeToRecord}\n\n`;
    
    if (procesarMalpractice) {
      resultadosTexto += `📊 MALPRACTICE RESULTS:\n` +
        `  • Total processed: ${results.malpractice.count}\n` +
        `  • Ready to escalate: ${results.malpractice.readyToEscalate}\n` +
        `  • Requires review: ${results.malpractice.requiresReview}\n` +
        `  • Missing PR contact: ${results.malpractice.missingPRContact}\n` +
        `  • Not affiliated: ${results.malpractice.notAffiliated}\n` +
        `  • Astrana inactive: ${results.malpractice.astranaInactive}\n` +
        `  • No active IPAs: ${results.malpractice.noActiveIPAs}\n` +
        `  • Not found in DB: ${results.malpractice.notFound}\n` +
        `  • Duplicates handled: ${results.malpractice.duplicates}\n` +
        `  • Used fallback keys: ${results.malpractice.fallbacks}\n\n`;
    }
    
    if (procesarExpirables) {
      resultadosTexto += `📋 EXPIRABLES RESULTS:\n` +
        `  • Total processed: ${results.expirables.count}\n` +
        `  • Ready to escalate: ${results.expirables.readyToEscalate}\n` +
        `  • Requires review: ${results.expirables.requiresReview}\n` +
        `  • Missing PR contact: ${results.expirables.missingPRContact}\n` +
        `  • Not affiliated: ${results.expirables.notAffiliated}\n` +
        `  • Astrana inactive: ${results.expirables.astranaInactive}\n` +
        `  • No active IPAs: ${results.expirables.noActiveIPAs}\n` +
        `  • Not found in DB: ${results.expirables.notFound}\n` +
        `  • Duplicates handled: ${results.expirables.duplicates}\n` +
        `  • Used fallback keys: ${results.expirables.fallbacks}\n\n`;
    }
    
    resultadosTexto += '✅ Please review the ESCALATION CENTER sheet for detailed results.\n\n';
    resultadosTexto += '🔍 Look for:\n';
    resultadosTexto += '  • Red background in # column = Duplicates\n';
    resultadosTexto += '  • Yellow background in Policy/License = Fallback keys\n';
    resultadosTexto += '  • Red background in Non-Affiliated IPAs = Not affiliated\n';
    resultadosTexto += '  • Orange background in Missing PR Contact = No contact info\n';
    resultadosTexto += '  • Italic text = Missing policy/license numbers';
    
    ui.alert('🎉 Process Completed Successfully', resultadosTexto, ui.ButtonSet.OK);
    
  } catch (error) {
    hideProgress();
    console.error('❌ Error in processEscalationsWithLevel:', error);
    const errorDetails = `Error occurred during escalation processing:\n\n` +
      `Error type: ${error.name || 'Unknown'}\n` +
      `Error message: ${error.message || 'No details available'}\n` +
      `Error location: ${error.stack ? error.stack.split('\n')[1] : 'Unknown location'}\n\n` +
      `Please check the console logs for more details.`;
    ui.alert('❌ Processing Error', errorDetails, ui.ButtonSet.OK);
  }
}

// ===== MAIN PROCESS - POPULATE ESCALATION CENTER ONLY =====
function ejecutarProcesoEscalacion(weekFilter, procesarMalpractice, procesarExpirables, statusFilterMalpractice, statusFilterExpirables, escalationTypeToRecord) {
  try {
    showProgress('📂 Loading spreadsheet data...');
    const ss = SpreadsheetApp.getActive();
    
    // Get required sheets
    const hojaMalpractice = procesarMalpractice ? ss.getSheetByName('MALPRACTICE 2025') : null;
    const hojaExpirables = procesarExpirables ? ss.getSheetByName('EXPIRABLES 2025') : null;
    const hojaIPASDB = ss.getSheetByName('IPAS DB');
    let hojaEscalation = ss.getSheetByName('ESCALATION CENTER');

    // Verify sheets exist
// CORRECTED LOGIC: Separate tracking for columns M and N
      if (procesarExpirables && !hojaExpirables) {
      throw new Error('EXPIRABLES 2025 sheet not found.');
    }
    if (!hojaIPASDB) {
      throw new Error('IPAS DB sheet not found.');
    }
    
    showProgress('🏗️ Preparing ESCALATION CENTER...');
    
    // Create ESCALATION CENTER if it doesn't exist and clear if it exists
    if (!hojaEscalation) {
      hojaEscalation = crearHojaEscalationCenter(ss);
    } else {
      // Clear existing content (except headers)
      const lastRow = hojaEscalation.getLastRow();
      if (lastRow > 1) {
        const range = hojaEscalation.getRange(2, 1, lastRow - 1, hojaEscalation.getLastColumn());
        range.clear();
      }
    }
    
    console.log('✅ Sheets loaded successfully');
    
    showProgress('🔍 Searching for NPIs...');
    
    // STEP 1: Get NPIs from both sheets with unique key detection
    let npisMalpractice = [];
    let npisExpirables = [];
    
    if (procesarMalpractice) {
      npisMalpractice = obtenerNPIsMalpracticeWithUniqueKeys(hojaMalpractice, weekFilter, statusFilterMalpractice);
      console.log(`✅ NPIs found in MALPRACTICE: ${npisMalpractice.length}`);
    }
    
    if (procesarExpirables) {
      npisExpirables = obtenerNPIsExpirablesWithUniqueKeys(hojaExpirables, weekFilter, statusFilterExpirables);
      console.log(`✅ NPIs found in EXPIRABLES: ${npisExpirables.length}`);
    }
    
    const totalNPIs = npisMalpractice.length + npisExpirables.length;
    if (totalNPIs === 0) {
      throw new Error(`No NPIs found with status "${statusFilterMalpractice}" or "${statusFilterExpirables}" for week ${weekFilter}.`);
    }
    
    showProgress(`📊 Loading IPAS database for ${totalNPIs} NPIs...`);
    
    // STEP 2: Load IPAS DB
    const ipasDatabase = cargarBaseDatosIPAS(hojaIPASDB);
    console.log(`✅ IPAS DB loaded: ${Object.keys(ipasDatabase).length} records`);
    
    showProgress('⚙️ Processing NPIs and analyzing IPAs...');
    
    // STEP 3: Process NPIs from both types with duplicate detection and IPA analysis
    const resultadosMalpractice = [];
    const resultadosExpirables = [];
    
    // Process MALPRACTICE with duplicate handling and IPA affiliation
    if (procesarMalpractice && npisMalpractice.length > 0) {
      const processedMalpractice = procesarNPIsConDuplicados(npisMalpractice, ipasDatabase, 'MALPRACTICE', escalationTypeToRecord);
      resultadosMalpractice.push(...processedMalpractice);
      console.log(`✅ Processed ${resultadosMalpractice.length} MALPRACTICE NPIs (duplicates and affiliation handled)`);
    }
    
    // Process EXPIRABLES with duplicate handling and IPA affiliation
    if (procesarExpirables && npisExpirables.length > 0) {
      const processedExpirables = procesarNPIsConDuplicados(npisExpirables, ipasDatabase, 'EXPIRABLES', escalationTypeToRecord);
      resultadosExpirables.push(...processedExpirables);
      console.log(`✅ Processed ${resultadosExpirables.length} EXPIRABLES NPIs (duplicates and affiliation handled)`);
    }
    
    showProgress('📝 Writing results to ESCALATION CENTER...');
    
    // STEP 4: Combine and write results to ESCALATION CENTER
    const todosLosResultados = [...resultadosMalpractice, ...resultadosExpirables];
    escribirResultadosOptimized(todosLosResultados, hojaEscalation);
    
    showProgress('🎨 Applying formatting...');
    
    // STEP 5: Apply formatting
    aplicarFormatoBonitoOptimized(hojaEscalation);
    
    console.log('✅ Process completed successfully');
    
    // Return separate statistics with comprehensive data
    return {
      malpractice: procesarMalpractice ?
        calcularEstadisticasCompletas(resultadosMalpractice) : 
        { count: 0, readyToEscalate: 0, requiresReview: 0, missingPRContact: 0, notAffiliated: 0, astranaInactive: 0, noActiveIPAs: 0, notFound: 0, duplicates: 0, fallbacks: 0 },
      expirables: procesarExpirables ?
        calcularEstadisticasCompletas(resultadosExpirables) : 
        { count: 0, readyToEscalate: 0, requiresReview: 0, missingPRContact: 0, notAffiliated: 0, astranaInactive: 0, noActiveIPAs: 0, notFound: 0, duplicates: 0, fallbacks: 0 }
    };
    
  } catch (error) {
    console.error('❌ Error in ejecutarProcesoEscalacion:', error);
    throw new Error(`Processing failed at step: ${error.message}`);
  }
}

// ===== PROGRESS TRACKING FUNCTIONS =====
function showProgress(message) {
  try {
    console.log(`🔄 PROGRESS: ${message}`);
    SpreadsheetApp.flush();
  } catch (error) {
    console.log(`Progress update: ${message}`);
  }
}

function hideProgress() {
  try {
    console.log('✅ Process completed');
    SpreadsheetApp.flush();
  } catch (error) {
    console.log('Process completed');
  }
}
