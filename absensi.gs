// ===== KONFIGURASI SHEET =====
const ANGGOTA_SHEET = 'Sheet1';    // Sheet untuk data anggota (nama, kelas, ekstra)
const ABSENSI_SHEET = 'Sheet2';    // Sheet untuk data absensi (tanggal, ekstra, kelas, nama, status)
const PENILAIAN_SHEET = 'Sheet3';  // Sheet untuk data penilaian (ekstra, kelas, nama, nilai, keterangan)

// ===== MAIN HANDLER =====
function doGet(e) {
  try {
    debugLog('=== REQUEST START ===');
    debugLog('Parameters received:', e.parameter);
    
    const action = e.parameter.action;
    const callback = e.parameter.callback;
    
    if (!action) {
      throw new Error('Action parameter is required');
    }
    
    let result;
    
    switch (action) {
      // Test & Debug
      case 'test':
        result = handleTest();
        break;
      case 'debugSheetStructure':
        result = debugSheetStructure();
        break;
        
      // Get Data
      case 'getAnggota':
        result = getAnggota();
        break;
      case 'getAbsensi':
        result = getAbsensi();
        break;
      case 'getPenilaian':
        result = getPenilaian();
        break;
        
      // Anggota Operations
      case 'addAnggota':
        result = addAnggota(e.parameter);
        break;
      case 'updateAnggota':
        result = updateAnggota(e.parameter);
        break;
      case 'deleteAnggota':
        result = deleteAnggota(e.parameter);
        break;
        
      // Absensi Operations
      case 'addAbsensi':
        result = addAbsensi(e.parameter);
        break;
      case 'addAbsensiBatch':
        result = addAbsensiBatch(e.parameter);
        break;
      case 'updateAbsensi':
        result = updateAbsensi(e.parameter);
        break;
      case 'deleteAbsensi':
        result = deleteAbsensi(e.parameter);
        break;
        
      // Penilaian Operations
      case 'addPenilaian':
        result = addPenilaian(e.parameter);
        break;
      case 'addPenilaianBatch':
        result = addPenilaianBatch(e.parameter);
        break;
      case 'updatePenilaian':
        result = updatePenilaian(e.parameter);
        break;
      case 'deletePenilaian':
        result = deletePenilaian(e.parameter);
        break;
        
      default:
        throw new Error(`Unknown action: ${action}`);
    }
    
    debugLog('=== REQUEST SUCCESS ===');
    debugLog('Result:', result);
    
    // Return JSONP response
    if (callback) {
      return ContentService
        .createTextOutput(`${callback}(${JSON.stringify(result)})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    debugLog('=== REQUEST ERROR ===');
    debugLog('Error:', error.toString());
    
    const errorResult = {
      success: false,
      error: error.toString(),
      message: error.message || 'Unknown error occurred'
    };
    
    if (e.parameter.callback) {
      return ContentService
        .createTextOutput(`${e.parameter.callback}(${JSON.stringify(errorResult)})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService
        .createTextOutput(JSON.stringify(errorResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

// ===== HELPER FUNCTIONS =====
function debugLog(message, data = null) {
  const timestamp = new Date().toISOString();
  if (data) {
    console.log(`[${timestamp}] ${message}`, data);
  } else {
    console.log(`[${timestamp}] ${message}`);
  }
}

// ===== TEST & DEBUG FUNCTIONS =====
function handleTest() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    return {
      success: true,
      message: 'Connection successful! Google Apps Script is working.',
      timestamp: new Date().toISOString(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetName: spreadsheet.getName(),
      availableSheets: sheets,
      expectedSheets: [ANGGOTA_SHEET, ABSENSI_SHEET, PENILAIAN_SHEET]
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      message: 'Connection test failed'
    };
  }
}

function debugSheetStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const structure = {};
    
    [ANGGOTA_SHEET, ABSENSI_SHEET, PENILAIAN_SHEET].forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        structure[sheetName] = {
          exists: true,
          rows: data.length,
          columns: data.length > 0 ? data[0].length : 0,
          headers: data.length > 0 ? data[0] : [],
          sampleData: data.length > 1 ? data[1] : []
        };
      } else {
        structure[sheetName] = {
          exists: false,
          error: 'Sheet not found'
        };
      }
    });
    
    return {
      success: true,
      structure: structure
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ===== GET DATA FUNCTIONS =====
function getAnggota() {
  try {
    debugLog('=== GET ANGGOTA START ===');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANGGOTA_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ANGGOTA_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Raw data rows:', data.length);
    
    if (data.length <= 1) {
      debugLog('No data found (only headers or empty)');
      return { success: true, data: [] };
    }
    
    // Skip header row and convert to objects
    const anggotaData = data.slice(1).map((row, index) => ({
      nama: row[0] || '',
      kelas: row[1] || '',
      ekstra: row[2] || '',
      row_number: index + 2 // +2 because we skip header and arrays are 0-indexed
    })).filter(item => item.nama.trim() !== ''); // Filter out empty rows
    
    debugLog('Processed anggota data:', anggotaData.length);
    debugLog('=== GET ANGGOTA END ===');
    
    return {
      success: true,
      data: anggotaData
    };
  } catch (error) {
    debugLog('❌ GET ANGGOTA ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      data: []
    };
  }
}

function getAbsensi() {
  try {
    debugLog('=== GET ABSENSI START ===');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABSENSI_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ABSENSI_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Raw data rows:', data.length);
    
    if (data.length <= 1) {
      debugLog('No data found (only headers or empty)');
      return { success: true, data: [] };
    }
    
    // Skip header row and convert to objects
    const absensiData = data.slice(1).map((row, index) => ({
      tanggal: formatDate(row[0]),
      ekstra: row[1] || '',
      kelas: row[2] || '',
      nama: row[3] || '',
      status: row[4] || '',
      row_number: index + 2
    })).filter(item => item.nama.trim() !== ''); // Filter out empty rows
    
    debugLog('Processed absensi data:', absensiData.length);
    debugLog('=== GET ABSENSI END ===');
    
    return {
      success: true,
      data: absensiData
    };
  } catch (error) {
    debugLog('❌ GET ABSENSI ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      data: []
    };
  }
}

function getPenilaian() {
  try {
    debugLog('=== GET PENILAIAN START ===');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PENILAIAN_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${PENILAIAN_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Raw data rows:', data.length);
    
    if (data.length <= 1) {
      debugLog('No data found (only headers or empty)');
      return { success: true, data: [] };
    }
    
    // Skip header row and convert to objects
    const penilaianData = data.slice(1).map((row, index) => ({
      ekstra: row[0] || '',
      kelas: row[1] || '',
      nama: row[2] || '',
      nilai: row[3] || '',
      keterangan: row[4] || '',
      row_number: index + 2
    })).filter(item => item.nama.trim() !== ''); // Filter out empty rows
    
    debugLog('Processed penilaian data:', penilaianData.length);
    debugLog('=== GET PENILAIAN END ===');
    
    return {
      success: true,
      data: penilaianData
    };
  } catch (error) {
    debugLog('❌ GET PENILAIAN ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      data: []
    };
  }
}

// ===== ANGGOTA FUNCTIONS =====
function addAnggota(params) {
  try {
    debugLog('=== ADD ANGGOTA START ===');
    debugLog('Params:', params);
    
    if (!params.nama || !params.kelas || !params.ekstra) {
      throw new Error('Nama, kelas, and ekstra are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANGGOTA_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ANGGOTA_SHEET} not found`);
    }
    
    // Add new row
    sheet.appendRow([params.nama, params.kelas, params.ekstra]);
    
    debugLog('✅ Anggota added successfully');
    debugLog('=== ADD ANGGOTA END ===');
    
    return {
      success: true,
      message: `Anggota ${params.nama} added successfully`
    };
  } catch (error) {
    debugLog('❌ ADD ANGGOTA ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to add anggota'
    };
  }
}

function updateAnggota(params) {
  try {
    debugLog('=== UPDATE ANGGOTA START ===');
    debugLog('Update params:', params);
    
    if (!params.originalNama || !params.originalKelas || !params.originalEkstra) {
      throw new Error('Original nama, kelas, and ekstra are required to identify the record');
    }
    
    if (!params.newNama || !params.newKelas || !params.newEkstra) {
      throw new Error('New nama, kelas, and ekstra are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANGGOTA_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ANGGOTA_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to update (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === params.originalNama && 
          row[1] === params.originalKelas && 
          row[2] === params.originalEkstra) {
        targetRow = i + 1; // +1 because sheet rows are 1-indexed
        debugLog(`Found target row: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Anggota record not found for: ${params.originalNama} (${params.originalKelas}, ${params.originalEkstra})`);
    }
    
    // Update the row
    sheet.getRange(targetRow, 1, 1, 3).setValues([[
      params.newNama,
      params.newKelas,
      params.newEkstra
    ]]);
    
    debugLog(`✅ Anggota updated successfully at row ${targetRow}`);
    debugLog('=== UPDATE ANGGOTA END ===');
    
    return {
      success: true,
      message: `Anggota ${params.newNama} updated successfully`,
      updatedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ UPDATE ANGGOTA ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to update anggota'
    };
  }
}

function deleteAnggota(params) {
  try {
    debugLog('=== DELETE ANGGOTA START ===');
    debugLog('Delete params:', params);
    
    if (!params.nama || !params.kelas || !params.ekstra) {
      throw new Error('Nama, kelas, and ekstra are required to identify the record');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANGGOTA_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ANGGOTA_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to delete (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === params.nama && 
          row[1] === params.kelas && 
          row[2] === params.ekstra) {
        targetRow = i + 1; // +1 because sheet rows are 1-indexed
        debugLog(`Found target row for deletion: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Anggota record not found for: ${params.nama} (${params.kelas}, ${params.ekstra})`);
    }
    
    // Delete the row
    sheet.deleteRow(targetRow);
    
    debugLog(`✅ Anggota deleted successfully from row ${targetRow}`);
    debugLog('=== DELETE ANGGOTA END ===');
    
    return {
      success: true,
      message: `Anggota ${params.nama} deleted successfully`,
      deletedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ DELETE ANGGOTA ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to delete anggota'
    };
  }
}

// ===== ABSENSI FUNCTIONS =====
function addAbsensi(params) {
  try {
    debugLog('=== ADD ABSENSI START ===');
    debugLog('Params:', params);
    
    if (!params.tanggal || !params.ekstra || !params.kelas || !params.nama || !params.status) {
      throw new Error('Tanggal, ekstra, kelas, nama, and status are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABSENSI_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ABSENSI_SHEET} not found`);
    }
    
    // Add new row
    sheet.appendRow([params.tanggal, params.ekstra, params.kelas, params.nama, params.status]);
    
    debugLog('✅ Absensi added successfully');
    debugLog('=== ADD ABSENSI END ===');
    
    return {
      success: true,
      message: `Absensi for ${params.nama} added successfully`
    };
  } catch (error) {
    debugLog('❌ ADD ABSENSI ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to add absensi'
    };
  }
}

function addAbsensiBatch(params) {
  try {
    debugLog('=== ADD ABSENSI BATCH START ===');
    debugLog('Batch params:', params);
    
    if (!params.absensiData) {
      throw new Error('absensiData parameter is required');
    }
    
    const absensiArray = JSON.parse(params.absensiData);
    debugLog('Parsed absensi array:', absensiArray);
    
    if (!Array.isArray(absensiArray) || absensiArray.length === 0) {
      throw new Error('Invalid absensi data array');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABSENSI_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ABSENSI_SHEET} not found`);
    }
    
    // Prepare batch data
    const batchData = [];
    absensiArray.forEach((absensi, index) => {
      debugLog(`Processing absensi ${index + 1}:`, absensi);
      
      if (!absensi.tanggal || !absensi.ekstra || !absensi.kelas || !absensi.nama || !absensi.status) {
        throw new Error(`Missing required fields in absensi ${index + 1}`);
      }
      
      batchData.push([
        absensi.tanggal,
        absensi.ekstra,
        absensi.kelas,
        absensi.nama,
        absensi.status
      ]);
    });
    
    debugLog('Batch data prepared:', batchData);
    
    // Insert all data at once
    const lastRow = sheet.getLastRow();
    const startRow = lastRow + 1;
    
    sheet.getRange(startRow, 1, batchData.length, 5).setValues(batchData);
    
    debugLog(`✅ Batch insert completed: ${batchData.length} records added starting from row ${startRow}`);
    debugLog('=== ADD ABSENSI BATCH END ===');
    
    return {
      success: true,
      message: `Successfully added ${batchData.length} absensi records`,
      recordsAdded: batchData.length,
      startRow: startRow
    };
  } catch (error) {
    debugLog('❌ ADD ABSENSI BATCH ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to add absensi batch'
    };
  }
}

function updateAbsensi(params) {
  try {
    debugLog('=== UPDATE ABSENSI START ===');
    debugLog('Update params:', params);
    
    if (!params.originalTanggal || !params.originalEkstra || !params.originalKelas || !params.originalNama) {
      throw new Error('Original tanggal, ekstra, kelas, and nama are required to identify the record');
    }
    
    if (!params.newTanggal || !params.newEkstra || !params.newKelas || !params.newNama || !params.newStatus) {
      throw new Error('New tanggal, ekstra, kelas, nama, and status are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABSENSI_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ABSENSI_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to update (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDate(row[0]);
      if (rowDate === params.originalTanggal && 
          row[1] === params.originalEkstra && 
          row[2] === params.originalKelas && 
          row[3] === params.originalNama) {
        targetRow = i + 1;
        debugLog(`Found target row: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Absensi record not found for: ${params.originalNama} on ${params.originalTanggal}`);
    }
    
    // Update the row
    sheet.getRange(targetRow, 1, 1, 5).setValues([[
      params.newTanggal,
      params.newEkstra,
      params.newKelas,
      params.newNama,
      params.newStatus
    ]]);
    
    debugLog(`✅ Absensi updated successfully at row ${targetRow}`);
    debugLog('=== UPDATE ABSENSI END ===');
    
    return {
      success: true,
      message: `Absensi for ${params.newNama} updated successfully`,
      updatedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ UPDATE ABSENSI ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to update absensi'
    };
  }
}

function deleteAbsensi(params) {
  try {
    debugLog('=== DELETE ABSENSI START ===');
    debugLog('Delete params:', params);
    
    if (!params.tanggal || !params.ekstra || !params.kelas || !params.nama) {
      throw new Error('Tanggal, ekstra, kelas, and nama are required to identify the record');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABSENSI_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${ABSENSI_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to delete (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDate(row[0]);
      if (rowDate === params.tanggal && 
          row[1] === params.ekstra && 
          row[2] === params.kelas && 
          row[3] === params.nama) {
        targetRow = i + 1;
        debugLog(`Found target row for deletion: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Absensi record not found for: ${params.nama} on ${params.tanggal}`);
    }
    
    // Delete the row
    sheet.deleteRow(targetRow);
    
    debugLog(`✅ Absensi deleted successfully from row ${targetRow}`);
    debugLog('=== DELETE ABSENSI END ===');
    
    return {
      success: true,
      message: `Absensi for ${params.nama} deleted successfully`,
      deletedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ DELETE ABSENSI ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to delete absensi'
    };
  }
}

// ===== PENILAIAN FUNCTIONS =====
function addPenilaian(params) {
  try {
    debugLog('=== ADD PENILAIAN START ===');
    debugLog('Params:', params);
    
    if (!params.ekstra || !params.kelas || !params.nama || !params.nilai) {
      throw new Error('Ekstra, kelas, nama, and nilai are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PENILAIAN_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${PENILAIAN_SHEET} not found`);
    }
    
    // Add new row
    sheet.appendRow([params.ekstra, params.kelas, params.nama, params.nilai, params.keterangan || '']);
    
    debugLog('✅ Penilaian added successfully');
    debugLog('=== ADD PENILAIAN END ===');
    
    return {
      success: true,
      message: `Penilaian for ${params.nama} added successfully`
    };
  } catch (error) {
    debugLog('❌ ADD PENILAIAN ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to add penilaian'
    };
  }
}

function addPenilaianBatch(params) {
  try {
    debugLog('=== ADD PENILAIAN BATCH START ===');
    debugLog('Batch params:', params);
    
    if (!params.penilaianData) {
      throw new Error('penilaianData parameter is required');
    }
    
    const penilaianArray = JSON.parse(params.penilaianData);
    debugLog('Parsed penilaian array:', penilaianArray);
    
    if (!Array.isArray(penilaianArray) || penilaianArray.length === 0) {
      throw new Error('Invalid penilaian data array');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PENILAIAN_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${PENILAIAN_SHEET} not found`);
    }
    
    // Prepare batch data
    const batchData = [];
    penilaianArray.forEach((penilaian, index) => {
      debugLog(`Processing penilaian ${index + 1}:`, penilaian);
      
      if (!penilaian.ekstra || !penilaian.kelas || !penilaian.nama || !penilaian.nilai) {
        throw new Error(`Missing required fields in penilaian ${index + 1}`);
      }
      
      batchData.push([
        penilaian.ekstra,
        penilaian.kelas,
        penilaian.nama,
        penilaian.nilai,
        penilaian.keterangan || ''
      ]);
    });
    
    debugLog('Batch data prepared:', batchData);
    
    // Insert all data at once
    const lastRow = sheet.getLastRow();
    const startRow = lastRow + 1;
    
    sheet.getRange(startRow, 1, batchData.length, 5).setValues(batchData);
    
    debugLog(`✅ Batch insert completed: ${batchData.length} records added starting from row ${startRow}`);
    debugLog('=== ADD PENILAIAN BATCH END ===');
    
    return {
      success: true,
      message: `Successfully added ${batchData.length} penilaian records`,
      recordsAdded: batchData.length,
      startRow: startRow
    };
  } catch (error) {
    debugLog('❌ ADD PENILAIAN BATCH ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to add penilaian batch'
    };
  }
}

function updatePenilaian(params) {
  try {
    debugLog('=== UPDATE PENILAIAN START ===');
    debugLog('Update params:', params);
    
    if (!params.originalEkstra || !params.originalKelas || !params.originalNama) {
      throw new Error('Original ekstra, kelas, and nama are required to identify the record');
    }
    
    if (!params.newEkstra || !params.newKelas || !params.newNama || !params.newNilai) {
      throw new Error('New ekstra, kelas, nama, and nilai are required');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PENILAIAN_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${PENILAIAN_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to update (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === params.originalEkstra && 
          row[1] === params.originalKelas && 
          row[2] === params.originalNama) {
        targetRow = i + 1;
        debugLog(`Found target row: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Penilaian record not found for: ${params.originalNama} (${params.originalKelas}, ${params.originalEkstra})`);
    }
    
    // Update the row
    sheet.getRange(targetRow, 1, 1, 5).setValues([[
      params.newEkstra,
      params.newKelas,
      params.newNama,
      params.newNilai,
      params.newKeterangan || ''
    ]]);
    
    debugLog(`✅ Penilaian updated successfully at row ${targetRow}`);
    debugLog('=== UPDATE PENILAIAN END ===');
    
    return {
      success: true,
      message: `Penilaian for ${params.newNama} updated successfully`,
      updatedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ UPDATE PENILAIAN ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to update penilaian'
    };
  }
}

function deletePenilaian(params) {
  try {
    debugLog('=== DELETE PENILAIAN START ===');
    debugLog('Delete params:', params);
    
    if (!params.ekstra || !params.kelas || !params.nama) {
      throw new Error('Ekstra, kelas, and nama are required to identify the record');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PENILAIAN_SHEET);
    if (!sheet) {
      throw new Error(`Sheet ${PENILAIAN_SHEET} not found`);
    }
    
    const data = sheet.getDataRange().getValues();
    debugLog('Sheet data loaded, rows:', data.length);
    
    // Find the row to delete (skip header row)
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === params.ekstra && 
          row[1] === params.kelas && 
          row[2] === params.nama) {
        targetRow = i + 1;
        debugLog(`Found target row for deletion: ${targetRow}`, row);
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`Penilaian record not found for: ${params.nama} (${params.kelas}, ${params.ekstra})`);
    }
    
    // Delete the row
    sheet.deleteRow(targetRow);
    
    debugLog(`✅ Penilaian deleted successfully from row ${targetRow}`);
    debugLog('=== DELETE PENILAIAN END ===');
    
    return {
      success: true,
      message: `Penilaian for ${params.nama} deleted successfully`,
      deletedRow: targetRow
    };
  } catch (error) {
    debugLog('❌ DELETE PENILAIAN ERROR:', error.toString());
    return {
      success: false,
      error: error.toString(),
      message: error.message || 'Failed to delete penilaian'
    };
  }
}

// ===== UTILITY FUNCTIONS =====
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else {
      return String(dateValue);
    }
    
    if (isNaN(date.getTime())) {
      return String(dateValue);
    }
    
    // Format to YYYY-MM-DD
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (error) {
    debugLog('Date formatting error:', error.toString());
    return String(dateValue);
  }
}