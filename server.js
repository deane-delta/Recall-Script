const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const XLSX = require('xlsx');
const FordScraper = require('./scraper/fordScraper');
const DocSearchScraper = require('./scraper/docsearchScraper');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.csv', '.xls', '.xlsx'];
    const fileExt = path.extname(file.originalname).toLowerCase();
    
    if (allowedTypes.includes(fileExt)) {
      cb(null, true);
    } else {
      cb(new Error('Only CSV, XLS, and XLSX files are allowed!'), false);
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const fileName = req.file.originalname;
    const vinColumn = req.body.vinColumn || 'auto';
    const sessionId = req.body.sessionId || Date.now().toString();

    // Process the uploaded file with session ID for progress updates
    const result = await processExcelFile(filePath, fileName, vinColumn, sessionId);
    
    // Clean up uploaded file
    fs.unlinkSync(filePath);

    // Close SSE connection
    emitProgress(sessionId, { type: 'complete', data: result });
    const client = sseClients.get(sessionId);
    if (client) {
      client.end();
    }
    sseClients.delete(sessionId);

    res.json({
      success: true,
      message: 'File processed successfully',
      data: result
    });

  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ 
      error: 'Error processing file', 
      details: error.message 
    });
  }
});

// New endpoint to get column information without processing
app.post('/get-columns', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const filePath = req.file.path;
    const fileName = req.file.originalname;

    // Read the Excel file to get column information
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    // Get available columns
    const availableColumns = Object.keys(data[0] || {}).map((key, index) => ({
      letter: String.fromCharCode(65 + index), // A, B, C, etc.
      name: key
    }));

    // Clean up uploaded file
    fs.unlinkSync(filePath);

    res.json({
      success: true,
      columns: availableColumns
    });

  } catch (error) {
    console.error('Error reading file columns:', error);
    res.status(500).json({ 
      error: 'Error reading file columns', 
      details: error.message 
    });
  }
});

// Store for SSE connections
const sseClients = new Map();

// SSE endpoint for progress updates
app.get('/progress/:sessionId', (req, res) => {
  const sessionId = req.params.sessionId;
  
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  sseClients.set(sessionId, res);
  
  req.on('close', () => {
    sseClients.delete(sessionId);
  });
});

// Helper function to emit progress
function emitProgress(sessionId, data) {
  const client = sseClients.get(sessionId);
  if (client) {
    client.write(`data: ${JSON.stringify(data)}\n\n`);
  }
}

app.get('/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'downloads', filename);
  
  if (fs.existsSync(filePath)) {
    res.download(filePath, filename, (err) => {
      if (err) {
        console.error('Download error:', err);
        res.status(500).json({ error: 'Download failed' });
      }
    });
  } else {
    res.status(404).json({ error: 'File not found' });
  }
});

// Comparison endpoint
app.post('/compare', upload.single('comparisonFile'), async (req, res) => {
  try {
    if (!req.file || !req.body.outputFile) {
      return res.status(400).json({ error: 'Missing required files' });
    }

    const comparisonFilePath = req.file.path;
    const outputFileName = req.body.outputFile;
    const outputFilePath = path.join(__dirname, 'downloads', outputFileName);

    if (!fs.existsSync(outputFilePath)) {
      // Clean up uploaded file
      fs.unlinkSync(comparisonFilePath);
      return res.status(404).json({ error: 'Output file not found' });
    }

    // Read the output file to extract all recall/satisfaction numbers
    console.log('Reading output file to extract recall/satisfaction numbers...');
    const outputWorkbook = XLSX.readFile(outputFilePath);
    const outputSheetName = outputWorkbook.SheetNames[0];
    const outputWorksheet = outputWorkbook.Sheets[outputSheetName];
    const outputData = XLSX.utils.sheet_to_json(outputWorksheet);

    // Extract all recall/satisfaction numbers with their types and asset numbers from output file
    // Store each recall-asset combination separately
    const recallSatisfactionList = []; // Array of {recallNumber, type, assetNo}
    outputData.forEach(row => {
      const recallNumber = row['Ford Recall Number'];
      const type = row['Type'] || 'Recall'; // Default to 'Recall' if not specified
      const assetNo = row['ASSET NO'] || '';
      if (recallNumber && 
          recallNumber !== 'No recall information' && 
          recallNumber !== 'No recall information available') {
        const trimmedRecall = recallNumber.toString().trim();
        // Store each recall-asset combination
        recallSatisfactionList.push({
          recallNumber: trimmedRecall,
          type: type,
          assetNo: assetNo.toString().trim()
        });
      }
    });

    // Create a Set of unique recall numbers for checking
    const uniqueRecallNumbers = new Set(recallSatisfactionList.map(item => item.recallNumber));
    console.log(`Found ${uniqueRecallNumbers.size} unique recall/satisfaction numbers in output file`);

    // Read the comparison file
    console.log('Reading comparison file...');
    const comparisonWorkbook = XLSX.readFile(comparisonFilePath);
    const comparisonSheetName = comparisonWorkbook.SheetNames[0];
    const comparisonWorksheet = comparisonWorkbook.Sheets[comparisonSheetName];
    const comparisonData = XLSX.utils.sheet_to_json(comparisonWorksheet);

    // Check if "Title" column exists
    if (comparisonData.length === 0) {
      fs.unlinkSync(comparisonFilePath);
      return res.status(400).json({ error: 'Comparison file is empty' });
    }

    const firstRow = comparisonData[0];
    const titleColumn = Object.keys(firstRow).find(key => 
      key.toLowerCase() === 'title'
    );

    if (!titleColumn) {
      fs.unlinkSync(comparisonFilePath);
      return res.status(400).json({ error: 'No "Title" column found in comparison file' });
    }

    // Extract all Title column values as an array for substring searching
    const titleValues = [];
    comparisonData.forEach(row => {
      const titleValue = row[titleColumn];
      if (titleValue) {
        titleValues.push(titleValue.toString().trim());
      }
    });

    console.log(`Found ${titleValues.length} entries in Title column of comparison file`);

    // Find missing recall/satisfaction numbers
    // Check if recall number appears anywhere in any Title value (substring search)
    const missingRecalls = [];
    const foundRecallNumbers = new Set(); // Track which recall numbers were found
    
    // First, check which recall numbers exist in the comparison file
    uniqueRecallNumbers.forEach(recallNumber => {
      const found = titleValues.some(titleValue => {
        const titleLower = titleValue.toLowerCase();
        const recallLower = recallNumber.toLowerCase();
        return titleLower.includes(recallLower);
      });
      
      if (found) {
        foundRecallNumbers.add(recallNumber);
      }
    });
    
    // Then, add all recall-asset combinations that are missing
    recallSatisfactionList.forEach(item => {
      if (!foundRecallNumbers.has(item.recallNumber)) {
        missingRecalls.push({
          recallNumber: item.recallNumber,
          type: item.type,
          assetNo: item.assetNo
        });
      }
    });

    console.log(`Found ${missingRecalls.length} missing recall/satisfaction numbers`);

    // Clean up uploaded comparison file
    fs.unlinkSync(comparisonFilePath);

    // Create Excel file with missing recall/satisfaction numbers
    if (missingRecalls.length > 0) {
      const missingWorkbook = XLSX.utils.book_new();
      const missingData = missingRecalls.map(item => ({
        'ASSET NO': item.assetNo,
        'Recall/Satisfaction Number': item.recallNumber,
        'Type': item.type
      }));

      const missingWorksheet = XLSX.utils.json_to_sheet(missingData);
      
      // Set column widths
      missingWorksheet['!cols'] = [
        { wch: 18 }, // ASSET NO
        { wch: 30 }, // Recall/Satisfaction Number
        { wch: 12 }  // Type
      ];
      
      XLSX.utils.book_append_sheet(missingWorkbook, missingWorksheet, 'Missing Recalls');
      
      const missingFileName = `missing_recalls_${Date.now()}.xlsx`;
      const missingFilePath = path.join(__dirname, 'downloads', missingFileName);
      
      XLSX.writeFile(missingWorkbook, missingFilePath);
      
      // Clean up old files, keeping only the 5 most recent
      cleanupOldFiles(5);

      res.json({
        success: true,
        missingFile: missingFileName,
        missingCount: missingRecalls.length,
        totalChecked: uniqueRecallNumbers.size
      });
    } else {
      res.json({
        success: true,
        missingFile: null,
        missingCount: 0,
        totalChecked: uniqueRecallNumbers.size,
        message: 'All recall/satisfaction numbers were found in the reference file'
      });
    }

  } catch (error) {
    console.error('Error comparing files:', error);
    res.status(500).json({ 
      error: 'Error comparing files', 
      details: error.message 
    });
  }
});

// Function to clean up old files, keeping only the most recent ones
function cleanupOldFiles(maxFiles = 5) {
  try {
    const downloadsDir = path.join(__dirname, 'downloads');
    
    // Check if downloads directory exists
    if (!fs.existsSync(downloadsDir)) {
      return;
    }
    
    // Get all files in the downloads directory
    const files = fs.readdirSync(downloadsDir);
    
    // Filter only Excel files
    const excelFiles = files
      .filter(file => file.startsWith('recall_data_') && file.endsWith('.xlsx'))
      .map(file => ({
        name: file,
        path: path.join(downloadsDir, file),
        stats: fs.statSync(path.join(downloadsDir, file))
      }));
    
    // Sort by modification time (newest first)
    excelFiles.sort((a, b) => b.stats.mtime - a.stats.mtime);
    
    // If we have more files than the limit, delete the oldest ones
    if (excelFiles.length > maxFiles) {
      const filesToDelete = excelFiles.slice(maxFiles);
      console.log(`Cleaning up ${filesToDelete.length} old file(s), keeping the 5 most recent...`);
      
      filesToDelete.forEach(file => {
        try {
          fs.unlinkSync(file.path);
          console.log(`Deleted old file: ${file.name}`);
        } catch (error) {
          console.error(`Error deleting file ${file.name}:`, error);
        }
      });
    }
  } catch (error) {
    console.error('Error cleaning up old files:', error);
  }
}

// Function to process Excel file
async function processExcelFile(filePath, fileName, vinColumn = 'auto', sessionId = null) {
  try {
    // Read the Excel file
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    // Emit progress: Reading Excel complete
    if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Reading Excel file...', progress: 10 });
    
    // Extract VIN numbers based on column selection
    const vinNumbers = [];
    const vinMap = new Map(); // Track VINs with their dates for duplicate handling
    
    // Helper function to parse date from "DATETIME OPEN" column (format: M/D/YYYY, e.g., "4/29/2021")
    const parseDate = (dateValue) => {
      if (!dateValue) return null;
      
      // Handle Excel date numbers (days since 1900-01-01)
      if (typeof dateValue === 'number') {
        try {
          const excelEpoch = new Date(1899, 11, 30);
          const date = new Date(excelEpoch.getTime() + dateValue * 86400000);
          if (!isNaN(date.getTime())) return date;
        } catch (e) {
          // Not a valid Excel date
        }
      }
      
      // Try parsing as string
      const strValue = dateValue.toString().trim();
      if (!strValue) return null;
      
      // Parse M/D/YYYY format (e.g., "4/29/2021")
      const dateMatch = strValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dateMatch) {
        const month = parseInt(dateMatch[1]) - 1; // Month is 0-indexed
        const day = parseInt(dateMatch[2]);
        const year = parseInt(dateMatch[3]);
        const date = new Date(year, month, day);
        if (!isNaN(date.getTime())) return date;
      }
      
      // Try other common formats
      const nativeDate = new Date(strValue);
      if (!isNaN(nativeDate.getTime())) return nativeDate;
      
      return null;
    };
    
    // Helper function to get date from "DATETIME OPEN" column
    const getDateFromRow = (row) => {
      const dateColumnKeys = ['DATETIME OPEN', 'DATETIME_OPEN', 'DATE TIME OPEN', 'DATE_TIME_OPEN'];
      
      for (const key of dateColumnKeys) {
        if (row[key]) {
          const parsedDate = parseDate(row[key]);
          if (parsedDate) {
            return { date: parsedDate, column: key };
          }
        }
      }
      
      // Try case-insensitive match
      for (const [key, value] of Object.entries(row)) {
        const upperKey = key.toUpperCase().replace(/\s+/g, ' ');
        if (upperKey.includes('DATETIME') && upperKey.includes('OPEN')) {
          const parsedDate = parseDate(value);
          if (parsedDate) {
            return { date: parsedDate, column: key };
          }
        }
      }
      
      return null;
    };
    
    if (vinColumn === 'auto') {
      // Auto-detect VIN column
      data.forEach(row => {
        // Look for VIN in various possible column names
        const vinKeys = ['SERIAL NO', 'VIN', 'vin', 'Vin', 'VIN NO'];
        let vin = null;
        let foundKey = null;
        
        for (const key of vinKeys) {
          if (row[key]) {
            const value = row[key].toString().trim();
            // Validate VIN format (exactly 17 characters matching VIN pattern)
            if (value.length === 17 && /^[A-HJ-NPR-Z0-9]{17}$/.test(value)) {
              vin = value;
              foundKey = key;
              break;
            } else if (value) {
              // Log skipped value that's not a valid VIN
              console.log(`Skipping non-VIN value in "${key}" column: "${value}" (length: ${value.length})`);
            }
          }
        }
        
        // If no VIN found, try to find any column that looks like a VIN (17 characters)
        if (!vin) {
          for (const [key, value] of Object.entries(row)) {
            if (value) {
              const strValue = value.toString().trim();
              if (strValue.length === 17 && /^[A-HJ-NPR-Z0-9]{17}$/.test(strValue)) {
                vin = strValue;
                foundKey = key;
                break;
              }
            }
          }
        }
        
        if (vin) {
          // Get date from "DATETIME OPEN" column
          const dateInfo = getDateFromRow(row);
          
          if (vinMap.has(vin)) {
            // Duplicate VIN found - compare dates
            const existing = vinMap.get(vin);
            const existingDate = existing.dateInfo ? existing.dateInfo.date : null;
            const currentDate = dateInfo ? dateInfo.date : null;
            
            if (currentDate && existingDate) {
              // Both have dates - keep the one with the latest date
              if (currentDate > existingDate) {
                console.log(`Duplicate VIN "${vin}": Keeping row with later date (${dateInfo.column}: ${dateInfo.date.toLocaleDateString()} vs ${existing.dateInfo.column}: ${existingDate.toLocaleDateString()})`);
                vinMap.set(vin, { vin, originalRow: row, dateInfo });
              } else {
                console.log(`Duplicate VIN "${vin}": Keeping existing row with later date (${existing.dateInfo.column}: ${existingDate.toLocaleDateString()} vs ${dateInfo.column}: ${dateInfo.date.toLocaleDateString()})`);
              }
            } else if (currentDate && !existingDate) {
              // Current row has date, existing doesn't - keep current
              console.log(`Duplicate VIN "${vin}": Keeping row with date (${dateInfo.column}: ${dateInfo.date.toLocaleDateString()})`);
              vinMap.set(vin, { vin, originalRow: row, dateInfo });
            } else if (!currentDate && existingDate) {
              // Existing has date, current doesn't - keep existing
              console.log(`Duplicate VIN "${vin}": Keeping existing row with date (${existing.dateInfo.column}: ${existingDate.toLocaleDateString()})`);
            } else {
              // Neither has date - keep the first one encountered
              console.log(`Duplicate VIN "${vin}": No date found in either row, keeping first occurrence`);
            }
          } else {
            // First occurrence of this VIN
            vinMap.set(vin, { vin, originalRow: row, dateInfo });
          }
        }
      });
      
      // Convert map to array
      vinMap.forEach((value) => {
        vinNumbers.push({
          vin: value.vin,
          originalRow: value.originalRow
        });
      });
    } else {
      // Use specified column (A=0, B=1, C=2, etc.)
      const columnIndex = vinColumn.charCodeAt(0) - 65; // Convert A=0, B=1, etc.
      const columnKeys = Object.keys(data[0] || {});
      
      if (columnIndex < columnKeys.length) {
        const columnKey = columnKeys[columnIndex];
        console.log(`Looking for VINs in column ${vinColumn} (${columnKey})`);
        
        data.forEach(row => {
          if (row[columnKey]) {
            const vin = row[columnKey].toString().trim();
            // Validate VIN format (exactly 17 characters matching VIN pattern)
            if (vin.length === 17 && /^[A-HJ-NPR-Z0-9]{17}$/.test(vin)) {
              // Get date from "DATETIME OPEN" column
              const dateInfo = getDateFromRow(row);
              
              if (vinMap.has(vin)) {
                // Duplicate VIN found - compare dates
                const existing = vinMap.get(vin);
                const existingDate = existing.dateInfo ? existing.dateInfo.date : null;
                const currentDate = dateInfo ? dateInfo.date : null;
                
                if (currentDate && existingDate) {
                  // Both have dates - keep the one with the latest date
                  if (currentDate > existingDate) {
                    console.log(`Duplicate VIN "${vin}": Keeping row with later date (${dateInfo.column}: ${dateInfo.date.toLocaleDateString()} vs ${existing.dateInfo.column}: ${existingDate.toLocaleDateString()})`);
                    vinMap.set(vin, { vin, originalRow: row, dateInfo });
                  } else {
                    console.log(`Duplicate VIN "${vin}": Keeping existing row with later date (${existing.dateInfo.column}: ${existingDate.toLocaleDateString()} vs ${dateInfo.column}: ${dateInfo.date.toLocaleDateString()})`);
                  }
                } else if (currentDate && !existingDate) {
                  // Current row has date, existing doesn't - keep current
                  console.log(`Duplicate VIN "${vin}": Keeping row with date (${dateInfo.column}: ${dateInfo.date.toLocaleDateString()})`);
                  vinMap.set(vin, { vin, originalRow: row, dateInfo });
                } else if (!currentDate && existingDate) {
                  // Existing has date, current doesn't - keep existing
                  console.log(`Duplicate VIN "${vin}": Keeping existing row with date (${existing.dateInfo.column}: ${existingDate.toLocaleDateString()})`);
                } else {
                  // Neither has date - keep the first one encountered
                  console.log(`Duplicate VIN "${vin}": No date found in either row, keeping first occurrence`);
                }
              } else {
                // First occurrence of this VIN
                vinMap.set(vin, { vin, originalRow: row, dateInfo });
              }
            } else if (vin) {
              // Log skipped value that's not a valid VIN
              console.log(`Skipping non-VIN value in column ${vinColumn} (${columnKey}): "${vin}" (length: ${vin.length})`);
            }
          }
        });
        
        // Convert map to array
        vinMap.forEach((value) => {
          vinNumbers.push({
            vin: value.vin,
            originalRow: value.originalRow
          });
        });
      } else {
        throw new Error(`Column ${vinColumn} not found in the file. Available columns: ${columnKeys.join(', ')}`);
      }
    }

    console.log(`Found ${vinNumbers.length} VIN numbers in ${fileName}`);
    
    // Determine which column was actually used for detection
    let detectedColumn = 'auto';
    let detectedColumnName = 'SERIAL NO';
    if (vinNumbers.length > 0 && vinColumn === 'auto') {
      // Find which column the VINs came from
      const firstVin = vinNumbers[0];
      const columnKeys = Object.keys(data[0] || {});
      
      for (const [key, value] of Object.entries(firstVin.originalRow)) {
        if (value && value.toString().trim() === firstVin.vin) {
          const columnIndex = columnKeys.indexOf(key);
          detectedColumn = String.fromCharCode(65 + columnIndex); // A, B, C, etc.
          detectedColumnName = key; // Store the actual column name used
          break;
        }
      }
    }
    
    // Emit progress: VINs extracted with actual column name
    if (sessionId) {
      const progressMessage = detectedColumnName === 'SERIAL NO' 
        ? 'Extracted VINs from "SERIAL NO" column'
        : `Extracted VINs from "${detectedColumnName}" column`;
      emitProgress(sessionId, { type: 'progress', message: progressMessage, progress: 15 });
    }
    
    if (vinNumbers.length === 0) {
      const columnInfo = vinColumn === 'auto' 
        ? 'using auto-detection' 
        : `in column ${vinColumn}`;
      
      // Get available columns for user reference
      const availableColumns = Object.keys(data[0] || {}).map((key, index) => ({
        letter: String.fromCharCode(65 + index), // A, B, C, etc.
        name: key
      }));
      
      // Check if "SERIAL NO" column exists
      const hasSerialNo = availableColumns.some(col => col.name === 'SERIAL NO');
      
      let message;
      if (vinColumn === 'auto' && !hasSerialNo) {
        message = `No VIN numbers found. Please ensure your Excel file has a column named "SERIAL NO" containing VIN numbers. Available columns: ${availableColumns.map(col => `${col.letter} (${col.name})`).join(', ')}`;
      } else {
        message = `No VIN numbers found ${columnInfo}. Available columns: ${availableColumns.map(col => `${col.letter} (${col.name})`).join(', ')}`;
      }
      
      return {
        fileName: fileName,
        totalRows: data.length,
        vinCount: 0,
        vinNumbers: [],
        columnUsed: vinColumn,
        detectedColumn: detectedColumn,
        availableColumns: availableColumns,
        message: message
      };
    }

    // Scrape data from Ford and DocSearch
    const scrapedData = await scrapeVinData(vinNumbers.map(item => item.vin), sessionId);
    
    // Create a map of VIN to originalRow for efficient lookup
    const vinToRowMap = new Map();
    vinNumbers.forEach(item => {
      vinToRowMap.set(item.vin, item.originalRow);
    });
    
    // Add originalRow data to each scraped result for Excel output
    // Match by VIN instead of index to ensure correct pairing
    // Also convert docsearchDataByRecall Map to plain object for JSON serialization
    const scrapedDataWithRows = scrapedData.map((item) => {
      const originalRow = vinToRowMap.get(item.vin);
      if (!originalRow) {
        console.warn(`⚠️ Warning: Could not find originalRow for VIN ${item.vin}`);
      }
      
      // Convert docsearchDataByRecall Map to plain object for JSON serialization
      let docsearchDataByRecallObj = {};
      if (item.docsearchDataByRecall && item.docsearchDataByRecall instanceof Map) {
        item.docsearchDataByRecall.forEach((value, key) => {
          docsearchDataByRecallObj[key] = value;
        });
      }
      
      return {
        ...item,
        originalRow: originalRow || {},
        docsearchDataByRecall: docsearchDataByRecallObj // Replace Map with plain object
      };
    });
    
    // Emit progress: Starting Excel creation
    if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Creating output file...', progress: 90 });
    
    // Create output Excel file
    const outputFileName = `recall_data_${Date.now()}.xlsx`;
    const outputPath = path.join(__dirname, 'downloads', outputFileName);
    
    await createOutputExcel(scrapedDataWithRows, outputPath);
    
    // Clean up old files, keeping only the 5 most recent
    cleanupOldFiles(5);
    
    // Convert docsearchDataByRecall Maps to plain objects for JSON serialization in scrapedData
    const scrapedDataForResponse = scrapedData.map((item) => {
      let docsearchDataByRecallObj = {};
      if (item.docsearchDataByRecall && item.docsearchDataByRecall instanceof Map) {
        item.docsearchDataByRecall.forEach((value, key) => {
          docsearchDataByRecallObj[key] = value;
        });
      }
      return {
        ...item,
        docsearchDataByRecall: docsearchDataByRecallObj // Replace Map with plain object
      };
    });
    
    return {
      fileName: fileName,
      totalRows: data.length,
      vinCount: vinNumbers.length,
      vinNumbers: vinNumbers.map(item => item.vin),
      columnUsed: vinColumn,
      detectedColumn: detectedColumn,
      scrapedData: scrapedDataForResponse,
      downloadFile: outputFileName,
      message: ` Successfully processed ${vinNumbers.length} VIN numbers and scraped recall data.`
    };

  } catch (error) {
    throw new Error(`Error reading Excel file: ${error.message}`);
  }
}

// Function to scrape VIN data from Ford and DocSearch
async function scrapeVinData(vinNumbers, sessionId = null) {
  const fordScraper = new FordScraper();
  // Initialize DocSearch scraper (credentials not required for manual sign-in)
  const docsearchScraper = new DocSearchScraper(
    process.env.DOCSEARCH_USERNAME || '',
    process.env.DOCSEARCH_PASSWORD || ''
  );

  const results = [];

  try {
    // Initialize Ford scraper first
    console.log('Initializing Ford scraper...');
    if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Initializing scrapers...', progress: 20 });
    const fordInitialized = await fordScraper.initialize();

    if (!fordInitialized) {
      console.warn('Ford scraper initialization failed');
    }

    // PHASE 1: Process all VINs with Ford scraper
    console.log(`\n=== PHASE 1: FORD SCRAPING (${vinNumbers.length} VINs) ===`);
    
    // Estimate time: ~5-10 seconds per VIN + 3 second delay = ~8-13 seconds per VIN
    const estimatedMinutes = Math.ceil((vinNumbers.length * 10) / 60);
    if (vinNumbers.length > 50) {
      console.log(`⏱️  Estimated time: ~${estimatedMinutes} minutes for ${vinNumbers.length} VINs`);
    }
    
    if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Scraping Ford recall data...', progress: 30 });
    
    const BATCH_SIZE = 50; // Restart browser every 50 VINs to prevent memory issues
    const REQUEST_TIMEOUT = 60000; // 60 second timeout per VIN
    
    for (let i = 0; i < vinNumbers.length; i++) {
      const vin = vinNumbers[i];
      const progressPercent = 30 + Math.floor((i / vinNumbers.length) * 30); // 30-60% progress
      console.log(`\nFord scraping VIN ${i + 1}/${vinNumbers.length}: ${vin}`);
      
      if (sessionId) {
        emitProgress(sessionId, { 
          type: 'progress', 
          message: `Scraping Ford recall data... (${i + 1}/${vinNumbers.length})`, 
          progress: progressPercent 
        });
      }

      const vinResult = {
        vin: vin,
        fordData: null,
        docsearchData: null, // Kept for backwards compatibility
        docsearchDataByRecall: new Map(), // Map of recall number -> DocSearch data
        processedAt: new Date().toISOString()
      };

      // Scrape Ford data with timeout and retry logic
      if (fordInitialized) {
        let retryCount = 0;
        const MAX_RETRIES = 1; // One retry with browser restart
        let scrapingSuccess = false;
        
        while (retryCount <= MAX_RETRIES && !scrapingSuccess) {
          try {
            // Add timeout wrapper for individual VIN scraping
            const scrapingPromise = fordScraper.scrapeVinRecallData(vin);
            const timeoutPromise = new Promise((_, reject) => 
              setTimeout(() => reject(new Error('Request timeout after 60 seconds')), REQUEST_TIMEOUT)
            );
            
            vinResult.fordData = await Promise.race([scrapingPromise, timeoutPromise]);
            
            // Check if scraping was successful
            if (vinResult.fordData && vinResult.fordData.success !== false) {
              scrapingSuccess = true;
              console.log(`✅ Ford data scraped for VIN: ${vin}`);
            } else {
              throw new Error(vinResult.fordData?.error || 'Scraping failed');
            }
          } catch (error) {
            const errorMessage = error.message || '';
            const isVinInputError = errorMessage.includes('Could not find VIN input field');
            const isNsErrorAbort = errorMessage.includes('NS_ERROR_ABORT');
            const needsBrowserRestart = isVinInputError || isNsErrorAbort;
            
            if (needsBrowserRestart && retryCount < MAX_RETRIES) {
              console.error(`❌ Error scraping Ford data for VIN ${vin}: ${errorMessage}`);
              console.log(`🔄 Restarting browser and retrying (attempt ${retryCount + 1}/${MAX_RETRIES + 1})...`);
              
              try {
                await fordScraper.close();
                await new Promise(resolve => setTimeout(resolve, 2000));
                const reinitialized = await fordScraper.initialize();
                
                if (!reinitialized) {
                  console.error('❌ Failed to restart browser. Marking as failed...');
                  vinResult.fordData = {
                    vin: vin,
                    success: false,
                    error: `Failed to restart browser after error: ${errorMessage}`
                  };
                  scrapingSuccess = true; // Stop retrying
                } else {
                  console.log('✅ Browser restarted successfully. Retrying...');
                  retryCount++;
                  // Continue to retry
                }
              } catch (restartError) {
                console.error('❌ Error during browser restart:', restartError);
                vinResult.fordData = {
                  vin: vin,
                  success: false,
                  error: `Browser restart failed: ${restartError.message}`
                };
                scrapingSuccess = true; // Stop retrying
              }
            } else {
              // No retry needed or max retries reached
              console.error(`❌ Error scraping Ford data for VIN ${vin}:`, errorMessage);
              vinResult.fordData = {
                vin: vin,
                success: false,
                error: errorMessage
              };
              scrapingSuccess = true; // Stop retrying
            }
          }
        }
      }

      results.push(vinResult);

      // Restart browser every BATCH_SIZE VINs to prevent memory issues and crashes
      if ((i + 1) % BATCH_SIZE === 0 && i < vinNumbers.length - 1 && fordInitialized) {
        console.log(`\n⚠️ Restarting browser after ${i + 1} VINs to maintain stability...`);
        await fordScraper.close();
        const reinitialized = await fordScraper.initialize();
        if (!reinitialized) {
          console.error('❌ Failed to restart browser. Continuing with existing instance...');
        } else {
          console.log('✅ Browser restarted successfully');
        }
      }

      // Add delay between VINs to be respectful to Ford website
      if (i < vinNumbers.length - 1) {
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
    }

    // Close Ford scraper after completing all VINs
    if (fordInitialized) {
      console.log('\n=== FORD SCRAPING COMPLETE ===');
      await fordScraper.close();
    }

    // PHASE 2: Initialize and authenticate DocSearch scraper
    console.log('\n=== PHASE 2: DOCSEARCH SCRAPING ===');
    console.log('Initializing DocSearch scraper...');
    if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Initializing DocSearch...', progress: 60 });
    const docsearchInitialized = await docsearchScraper.initialize();

    if (!docsearchInitialized) {
      console.warn('DocSearch scraper initialization failed');
    }

    // Authenticate with DocSearch
    let docsearchAuthenticated = false;
    if (docsearchInitialized) {
      const checkSignedIn = await docsearchScraper.checkIfAlreadySignedIn();
      
      if (checkSignedIn) {
        console.log('✅ User is already signed into DocSearch');
        docsearchAuthenticated = true;
      } else {
        console.log('⚠️ User is not signed into DocSearch');
        console.log('📝 ACTION REQUIRED: Please sign in to DocSearch manually in the browser window that opened');
        if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Waiting for manual DocSearch sign-in...', progress: 65 });
        
        // Wait for manual sign-in
        docsearchAuthenticated = await docsearchScraper.authenticate();
        if (!docsearchAuthenticated) {
          console.warn('❌ DocSearch authentication failed');
          if (sessionId) emitProgress(sessionId, { type: 'error', message: 'DocSearch sign-in required. Please sign in and try again.' });
        } else {
          console.log('✅ DocSearch authentication successful');
          if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Scraping DocSearch data...', progress: 70 });
        }
      }
    }

    // PHASE 3: Process only VINs with Ford recall numbers through DocSearch scraper
    if (docsearchAuthenticated) {
      // STEP 1: Collect all unique recall/satisfaction numbers across all VINs
      const uniqueRecallNumbers = new Set();
      const recallToVinsMap = new Map(); // Map recall number to array of VINs that have it
      
      // Initialize docsearchDataByRecall map for each result
      for (const result of results) {
        if (result.fordData && 
            result.fordData.success && 
            result.fordData.recallData && 
            result.fordData.recallData.recalls) {
          
          result.docsearchDataByRecall = new Map();
          
          // Get all valid recall numbers for this VIN
          const validRecalls = result.fordData.recallData.recalls.filter(recall => 
            recall.recallNumber && 
            recall.recallNumber !== 'No recall information' && 
            recall.recallNumber !== 'No recall information available'
          );
          
          // Track which VINs have which recall numbers
          for (const recall of validRecalls) {
            const recallNum = recall.recallNumber;
            uniqueRecallNumbers.add(recallNum);
            
            if (!recallToVinsMap.has(recallNum)) {
              recallToVinsMap.set(recallNum, []);
            }
            recallToVinsMap.get(recallNum).push({
              vin: result.vin,
              result: result
            });
          }
        }
      }

      // STEP 2: Convert Set to Array for iteration
      const uniqueRecallsArray = Array.from(uniqueRecallNumbers);
      const totalRecallsBeforeDedup = Array.from(recallToVinsMap.values()).reduce((sum, vins) => sum + vins.length, 0);
      
      console.log(`\n=== PHASE 3: DOCSEARCH SCRAPING ===`);
      console.log(`📊 Total recall/satisfaction numbers found: ${totalRecallsBeforeDedup}`);
      console.log(`✅ Unique recall/satisfaction numbers to search: ${uniqueRecallsArray.length}`);
      console.log(`⚡ Efficiency improvement: ${totalRecallsBeforeDedup - uniqueRecallsArray.length} duplicate searches avoided`);
      
      if (sessionId) emitProgress(sessionId, { type: 'progress', message: 'Scraping DocSearch data...', progress: 75 });
      
      // STEP 3: Search each unique recall number once and store results
      const recallToDocsearchDataMap = new Map(); // Map recall number to DocSearch result
      const DOCSEARCH_BATCH_SIZE = 100; // Restart browser every 100 requests
      const DOCSEARCH_REQUEST_TIMEOUT = 60000; // 60 second timeout per request
      
      for (let i = 0; i < uniqueRecallsArray.length; i++) {
        const recallNumber = uniqueRecallsArray[i];
        const progressPercent = 75 + Math.floor((i / uniqueRecallsArray.length) * 15); // 75-90% progress
        const vinsWithThisRecall = recallToVinsMap.get(recallNumber);
        
        console.log(`\nDocSearch scraping ${i + 1}/${uniqueRecallsArray.length}: Recall ${recallNumber} (affects ${vinsWithThisRecall.length} VIN(s))`);
        
        if (sessionId) {
          emitProgress(sessionId, { 
            type: 'progress', 
            message: `Scraping DocSearch data... (${i + 1}/${uniqueRecallsArray.length})`, 
            progress: progressPercent 
          });
        }

        try {
          // Pass the recall number to DocSearch with timeout
          const scrapingPromise = docsearchScraper.searchVinData(recallNumber);
          const timeoutPromise = new Promise((_, reject) => 
            setTimeout(() => reject(new Error('Request timeout after 60 seconds')), DOCSEARCH_REQUEST_TIMEOUT)
          );
          
          const docsearchData = await Promise.race([scrapingPromise, timeoutPromise]);
          
          // Store DocSearch data for this recall number
          recallToDocsearchDataMap.set(recallNumber, docsearchData);
          
          console.log(`✅ DocSearch data scraped for Recall ${recallNumber} (EA Exists: ${docsearchData.eaExists}, EA Number: ${docsearchData.eaNumber || 'NONE'})`);
          console.log(`   → This result will be applied to ${vinsWithThisRecall.length} VIN(s)`);
        } catch (error) {
          console.error(`❌ Error scraping DocSearch data for Recall ${recallNumber}:`, error.message);
          recallToDocsearchDataMap.set(recallNumber, {
            recallNumber: recallNumber,
            success: false,
            error: error.message,
            eaExists: false,
            eaNumber: null
          });
        }

        // Restart browser every BATCH_SIZE requests to prevent memory issues
        if ((i + 1) % DOCSEARCH_BATCH_SIZE === 0 && i < uniqueRecallsArray.length - 1 && docsearchInitialized) {
          console.log(`\n⚠️ Restarting DocSearch browser after ${i + 1} requests to maintain stability...`);
          await docsearchScraper.close();
          const reinitialized = await docsearchScraper.initialize();
          if (!reinitialized) {
            console.error('❌ Failed to restart DocSearch browser. Continuing with existing instance...');
          } else {
            console.log('✅ DocSearch browser restarted successfully');
            // Re-authenticate if needed
            const checkSignedIn = await docsearchScraper.checkIfAlreadySignedIn();
            if (!checkSignedIn) {
              console.log('⚠️ DocSearch session expired. Please sign in again.');
              if (sessionId) emitProgress(sessionId, { type: 'error', message: 'DocSearch session expired. Please sign in and try again.' });
              break; // Stop processing if authentication is lost
            }
          }
        }

        // Add delay between requests to be respectful to DocSearch website
        if (i < uniqueRecallsArray.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 3000));
        }
      }

      // STEP 4: Map DocSearch results back to each VIN's recall data
      console.log(`\n=== MAPPING DOCSEARCH RESULTS TO VINS ===`);
      for (const [recallNumber, vins] of recallToVinsMap.entries()) {
        const docsearchData = recallToDocsearchDataMap.get(recallNumber);
        
        if (docsearchData) {
          // Assign the DocSearch data to all VINs that have this recall number
          for (const { vin, result } of vins) {
            result.docsearchDataByRecall.set(recallNumber, docsearchData);
          }
          console.log(`✅ Mapped DocSearch data for Recall ${recallNumber} to ${vins.length} VIN(s)`);
        } else {
          // If no DocSearch data (error case), assign error result to all VINs
          const errorData = {
            recallNumber: recallNumber,
            success: false,
            error: 'DocSearch data not available',
            eaExists: false,
            eaNumber: null
          };
          for (const { vin, result } of vins) {
            result.docsearchDataByRecall.set(recallNumber, errorData);
          }
          console.log(`⚠️ No DocSearch data available for Recall ${recallNumber} (affects ${vins.length} VIN(s))`);
        }
      }
    }

    // Close DocSearch scraper
    if (docsearchInitialized) {
      console.log('\n=== DOCSEARCH SCRAPING COMPLETE ===');
      await docsearchScraper.close();
    }

  } catch (error) {
    console.error('Error during scraping process:', error);
    throw error;
  } finally {
    // Clean up scrapers
    try {
      await fordScraper.close();
      await docsearchScraper.close();
    } catch (error) {
      console.error('Error closing scrapers:', error);
    }
  }

  return results;
}

// Function to create output Excel file
async function createOutputExcel(scrapedData, outputPath) {
  try {
    const workbook = XLSX.utils.book_new();
    
    // Prepare data for Excel
    const excelData = [];
    
    console.log(`\n=== CREATING EXCEL OUTPUT ===`);
    console.log(`Total scraped items: ${scrapedData.length}`);
    
    let itemsWithRecalls = 0;
    let itemsWithoutRecalls = 0;
    let itemsWithoutOriginalRow = 0;
    
    scrapedData.forEach((item, index) => {
      // Debug: Check if originalRow exists
      if (!item.originalRow) {
        itemsWithoutOriginalRow++;
        console.warn(`⚠️ Warning: Item ${index} (VIN: ${item.vin}) has no originalRow`);
      }
      
      // Debug: Log Ford data status
      if (!item.fordData) {
        console.log(`   VIN ${item.vin}: No Ford data`);
      } else if (!item.fordData.success) {
        console.log(`   VIN ${item.vin}: Ford scraping failed - ${item.fordData.error || 'Unknown error'}`);
      } else if (!item.fordData.recallData) {
        console.log(`   VIN ${item.vin}: No recallData in Ford response`);
      } else if (!item.fordData.recallData.recalls) {
        console.log(`   VIN ${item.vin}: No recalls array in Ford response`);
      } else {
        console.log(`   VIN ${item.vin}: Found ${item.fordData.recallData.recalls.length} recall(s)`);
      }
      
      // Check if this VIN has recall numbers
      // Note: We check recallData even if success is false, in case data was partially extracted
      const hasRecalls = item.fordData && 
                        item.fordData.recallData && 
                        item.fordData.recallData.recalls && 
                        Array.isArray(item.fordData.recallData.recalls) &&
                        item.fordData.recallData.recalls.length > 0 &&
                        item.fordData.recallData.recalls.some(recall => 
                          recall && 
                          recall.recallNumber && 
                          typeof recall.recallNumber === 'string' &&
                          recall.recallNumber.trim() !== '' &&
                          recall.recallNumber !== 'No recall information' && 
                          recall.recallNumber !== 'No recall information available'
                        );
      
      if (hasRecalls) {
        itemsWithRecalls++;
      } else {
        itemsWithoutRecalls++;
      }

      // Only include VINs that have recall numbers
      if (hasRecalls) {
        // Helper function to safely get column value with fallback variations
        const getColumnValue = (row, primaryKey, ...alternateKeys) => {
          if (!row) return '';
          
          // Normalize function to trim and uppercase for comparison
          const normalize = (str) => str ? str.toString().trim().toUpperCase().replace(/\s+/g, ' ') : '';
          
          const normalizedPrimary = normalize(primaryKey);
          const normalizedAlternates = alternateKeys.map(k => normalize(k));
          
          // Try primary key first (exact match)
          if (row[primaryKey]) return row[primaryKey];
          
          // Try alternate keys (exact match)
          for (const key of alternateKeys) {
            if (row[key]) return row[key];
          }
          
          // Try case-insensitive and whitespace-tolerant match
          for (const [key, value] of Object.entries(row)) {
            const normalizedKey = normalize(key);
            
            // Check against primary key
            if (normalizedKey === normalizedPrimary) {
              return value;
            }
            
            // Check against alternate keys
            for (const normalizedAlt of normalizedAlternates) {
              if (normalizedKey === normalizedAlt) {
                return value;
              }
            }
          }
          
          return '';
        };
        
        // Helper function to format date values (handles Excel serial numbers and date strings)
        const formatDate = (dateValue) => {
          if (!dateValue) return '';
          
          // Handle Excel date serial numbers (days since 1900-01-01)
          if (typeof dateValue === 'number') {
            try {
              // Excel epoch is December 30, 1899 (not January 1, 1900)
              const excelEpoch = new Date(1899, 11, 30);
              const date = new Date(excelEpoch.getTime() + dateValue * 86400000);
              if (!isNaN(date.getTime())) {
                // Format as M/D/YYYY
                const month = date.getMonth() + 1;
                const day = date.getDate();
                const year = date.getFullYear();
                return `${month}/${day}/${year}`;
              }
            } catch (e) {
              // Not a valid Excel date
            }
          }
          
          // If it's already a string, return it as-is (might already be formatted)
          if (typeof dateValue === 'string') {
            return dateValue.trim();
          }
          
          // Try to parse as Date object
          if (dateValue instanceof Date) {
            const month = dateValue.getMonth() + 1;
            const day = dateValue.getDate();
            const year = dateValue.getFullYear();
            return `${month}/${day}/${year}`;
          }
          
          return dateValue.toString();
        };
        
        // Get valid recall objects (keep full objects to access type)
        const validRecalls = item.fordData && item.fordData.recallData && item.fordData.recallData.recalls
          ? item.fordData.recallData.recalls
              .filter(recall => recall.recallNumber && 
                               recall.recallNumber !== 'No recall information' && 
                               recall.recallNumber !== 'No recall information available')
          : [];

        // Create a separate row for each recall number
        if (validRecalls.length > 0) {
          for (const recall of validRecalls) {
            // Get EA number from DocSearch data for this specific recall number
            let eaNumber = 'NONE';
            if (item.docsearchDataByRecall) {
              let docsearchData = null;
              if (item.docsearchDataByRecall instanceof Map) {
                // Handle as Map (if not yet converted)
                docsearchData = item.docsearchDataByRecall.get(recall.recallNumber);
              } else if (typeof item.docsearchDataByRecall === 'object') {
                // Handle as plain object (converted from Map for JSON serialization)
                docsearchData = item.docsearchDataByRecall[recall.recallNumber];
              }
              if (docsearchData && docsearchData.eaNumber) {
                eaNumber = docsearchData.eaNumber;
              }
            } else if (item.docsearchData && item.docsearchData.eaNumber) {
              // Fallback to old format for backwards compatibility
              eaNumber = item.docsearchData.eaNumber;
            }

            // Get recall type (Recall or Satisfaction), default to "Recall" if not specified
            const recallType = recall.type || 'Recall';

            const row = {
              'ASSET NO': getColumnValue(item.originalRow, 'ASSET NO', 'EQ EQUIP NO', 'EQ EQUIPMENT NO', 'EQUIPMENT NO', 'EQUIP NO'),
              'YEAR': getColumnValue(item.originalRow, 'YEAR'),
              'MODEL': getColumnValue(item.originalRow, 'MODEL'),
              'MANUFACTURER': getColumnValue(item.originalRow, 'MANUFACTURER', 'MAKE'),
              'STATION': getColumnValue(item.originalRow, 'STATION', 'LOC ASSIGN PM LOC', 'LOC', 'PM LOC', 'LOCATION'),
              'VIN': item.vin,
              'Ford Recall Number': recall.recallNumber,
              'Type': recallType,
              'EA Number': eaNumber,
              'Work Order': getColumnValue(item.originalRow, 'Work Order', 'WORK ORDER', 'WO', 'WORK ORDER NO', 'WORK ORDER NUMBER', 'WO NUMBER', 'WorkOrder', 'WORKORDER'),
              'WORK ORDER STATUS': (() => {
                const status = getColumnValue(item.originalRow, 'WORK ORDER STATUS', 'WO STATUS', 'WO Status', 'Work Order Status', 'WORK ORDER STAT', 'WO STAT', 'WorkOrderStatus', 'WORKORDERSTATUS', 'WOStatus', 'WOSTATUS');
                return status && status.toString().trim() !== '' ? status : 'NONE';
              })()
            };

            excelData.push(row);
            console.log(`✅ Added VIN ${item.vin} to Excel with recall number: ${recall.recallNumber} (Type: ${recallType})`);
          }
          console.log(`   Created ${validRecalls.length} row(s) for VIN ${item.vin} (${validRecalls.length} recall number(s))`);
        }
      } else {
        console.log(`❌ Skipped VIN ${item.vin} - no recall numbers found`);
      }
    });
    
    // Summary before creating Excel
    console.log(`\n=== EXCEL DATA SUMMARY ===`);
    console.log(`📊 Total scraped items: ${scrapedData.length}`);
    console.log(`✅ Items with recalls: ${itemsWithRecalls}`);
    console.log(`❌ Items without recalls: ${itemsWithoutRecalls}`);
    console.log(`⚠️ Items without originalRow: ${itemsWithoutOriginalRow}`);
    console.log(`📝 Rows to write to Excel: ${excelData.length}`);

    // Create worksheet
    if (excelData.length === 0) {
      console.warn('⚠️ Warning: No data to write to Excel. File will be empty.');
      console.warn('   This usually means:');
      console.warn('   1. No VINs were found in the input file, OR');
      console.warn('   2. No VINs had recall numbers from Ford, OR');
      console.warn('   3. There was an error during scraping');
    }
    
    // SHEET 1: Grouped by Recall Number (New Design)
    // Group all rows by "Ford Recall Number"
    const recallGroups = new Map();
    
    for (const row of excelData) {
      const recallNumber = row['Ford Recall Number'];
      if (recallNumber) {
        if (!recallGroups.has(recallNumber)) {
          recallGroups.set(recallNumber, []);
        }
        recallGroups.get(recallNumber).push(row);
      }
    }
    
    // Separate recall numbers into multi-occurrence and single-occurrence
    const multiOccurrenceRecalls = [];
    const singleOccurrenceRecalls = [];
    
    for (const [recallNumber, vehicles] of recallGroups.entries()) {
      if (vehicles.length > 1) {
        multiOccurrenceRecalls.push({ recallNumber, vehicles });
      } else {
        singleOccurrenceRecalls.push({ recallNumber, vehicles });
      }
    }
    
    // Sort both groups by recall number
    multiOccurrenceRecalls.sort((a, b) => a.recallNumber.localeCompare(b.recallNumber));
    singleOccurrenceRecalls.sort((a, b) => a.recallNumber.localeCompare(b.recallNumber));
    
    // Create grouped data: multi-occurrence first, then single-occurrence at bottom
    const groupedData = [];
    
    // Add multi-occurrence recalls first
    for (const { recallNumber, vehicles } of multiOccurrenceRecalls) {
      for (const vehicle of vehicles) {
        // Reorder columns to put "Ford Recall Number" first
        const groupedRow = {
          'Ford Recall Number': vehicle['Ford Recall Number'],
          'Type': vehicle['Type'],
          'EA Number': vehicle['EA Number'],
          'ASSET NO': vehicle['ASSET NO'],
          'YEAR': vehicle['YEAR'],
          'MODEL': vehicle['MODEL'],
          'MANUFACTURER': vehicle['MANUFACTURER'],
          'STATION': vehicle['STATION'],
          'VIN': vehicle['VIN'],
          'Work Order': vehicle['Work Order'],
          'WORK ORDER STATUS': vehicle['WORK ORDER STATUS']
        };
        groupedData.push(groupedRow);
      }
    }
    
    // Add single-occurrence recalls at the bottom
    for (const { recallNumber, vehicles } of singleOccurrenceRecalls) {
      for (const vehicle of vehicles) {
        // Reorder columns to put "Ford Recall Number" first
        const groupedRow = {
          'Ford Recall Number': vehicle['Ford Recall Number'],
          'Type': vehicle['Type'],
          'EA Number': vehicle['EA Number'],
          'ASSET NO': vehicle['ASSET NO'],
          'YEAR': vehicle['YEAR'],
          'MODEL': vehicle['MODEL'],
          'MANUFACTURER': vehicle['MANUFACTURER'],
          'STATION': vehicle['STATION'],
          'VIN': vehicle['VIN'],
          'Work Order': vehicle['Work Order'],
          'WORK ORDER STATUS': vehicle['WORK ORDER STATUS']
        };
        groupedData.push(groupedRow);
      }
    }
    
    console.log(`📊 Recall grouping: ${multiOccurrenceRecalls.length} multi-occurrence, ${singleOccurrenceRecalls.length} single-occurrence`);
    
    // Create grouped worksheet
    let groupedWorksheet;
    if (groupedData.length > 0) {
      groupedWorksheet = XLSX.utils.json_to_sheet(groupedData);
      
      // Set column widths for grouped sheet (Ford Recall Number first)
      const groupedColumnWidths = [
        { wch: 25 }, // Ford Recall Number (first column)
        { wch: 12 }, // Type
        { wch: 25 }, // EA Number
        { wch: 18 }, // ASSET NO
        { wch: 8 },  // YEAR
        { wch: 15 }, // MODEL
        { wch: 15 }, // MANUFACTURER
        { wch: 20 }, // STATION
        { wch: 20 }, // VIN
        { wch: 18 }, // Work Order
        { wch: 20 }  // WORK ORDER STATUS
      ];
      groupedWorksheet['!cols'] = groupedColumnWidths;
      
      // Set up row grouping/outlining for collapsible groups
      // Initialize !rows array (row 0 is header, so data starts at row 1)
      if (!groupedWorksheet['!rows']) {
        groupedWorksheet['!rows'] = [];
      }
      
      // Track current recall number and group boundaries
      let currentRecallNumber = null;
      let groupStartRow = null;
      const rowGroups = []; // Array of {startRow, endRow, recallNumber}
      
      // Identify groups (consecutive rows with same recall number)
      for (let i = 0; i < groupedData.length; i++) {
        const rowRecallNumber = groupedData[i]['Ford Recall Number'];
        
        if (rowRecallNumber !== currentRecallNumber) {
          // New group starting
          if (currentRecallNumber !== null && groupStartRow !== null) {
            // Save previous group
            rowGroups.push({
              startRow: groupStartRow,
              endRow: i, // End row (exclusive)
              recallNumber: currentRecallNumber
            });
          }
          currentRecallNumber = rowRecallNumber;
          groupStartRow = i + 1; // +1 because row 0 is header
        }
      }
      
      // Save last group
      if (currentRecallNumber !== null && groupStartRow !== null) {
        rowGroups.push({
          startRow: groupStartRow,
          endRow: groupedData.length, // End row (exclusive)
          recallNumber: currentRecallNumber
        });
      }
      
      // Set outline levels for each row to create collapsible groups
      // Row 0 is header (level 0, no outline)
      groupedWorksheet['!rows'][0] = { level: 0, hidden: false };
      
      // Initialize all rows
      for (let i = 1; i <= groupedData.length; i++) {
        if (!groupedWorksheet['!rows'][i]) {
          groupedWorksheet['!rows'][i] = {};
        }
      }
      
      // Set outline levels for each group
      // First row of each group = level 1 (visible, acts as group header)
      // Remaining rows = level 2 (hidden by default, collapsed)
      for (const group of rowGroups) {
        // First row of group is at level 1 (visible, acts as summary/header)
        const firstRowIndex = group.startRow;
        groupedWorksheet['!rows'][firstRowIndex].level = 1;
        groupedWorksheet['!rows'][firstRowIndex].hidden = false;
        
        // Remaining rows in group are at level 2 (hidden by default - collapsed)
        for (let i = firstRowIndex + 1; i < group.endRow; i++) {
          groupedWorksheet['!rows'][i].level = 2;
          groupedWorksheet['!rows'][i].hidden = true; // Hidden by default = collapsed
        }
      }
      
      // Configure outline settings for row grouping
      // This tells Excel to create collapsible groups for rows at level 2
      groupedWorksheet['!outline'] = {
        above: false,      // Groups are below (rows are grouped, not columns)
        below: true,       // Groups are below
        left: false,
        right: false,
        summaryBelow: true,  // Summary rows are below the detail
        summaryRight: false
      };
      
      console.log(`📋 Created ${rowGroups.length} collapsible recall groups`);
      console.log(`   Groups are collapsed by default - use outline controls to expand`);
    } else {
      // Create empty worksheet with headers
      groupedWorksheet = XLSX.utils.json_to_sheet([{
        'Ford Recall Number': '',
        'Type': '',
        'EA Number': '',
        'ASSET NO': '',
        'YEAR': '',
        'MODEL': '',
        'MANUFACTURER': '',
        'STATION': '',
        'VIN': '',
        'Work Order': '',
        'WORK ORDER STATUS': ''
      }]);
      const groupedColumnWidths = [
        { wch: 25 }, { wch: 12 }, { wch: 25 }, { wch: 18 }, { wch: 8 },
        { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 18 }, { wch: 20 }
      ];
      groupedWorksheet['!cols'] = groupedColumnWidths;
    }
    
    // SHEET 2: Original format (Current Design)
    // Create worksheet with data (will have headers automatically from keys)
    let worksheet;
    if (excelData.length > 0) {
      worksheet = XLSX.utils.json_to_sheet(excelData);
      
      // Set column widths for original sheet
      const columnWidths = [
        { wch: 18 }, // ASSET NO
        { wch: 8 },  // YEAR
        { wch: 15 }, // MODEL
        { wch: 15 }, // MANUFACTURER
        { wch: 20 }, // STATION
        { wch: 20 }, // VIN
        { wch: 25 }, // Ford Recall Number
        { wch: 12 }, // Type
        { wch: 25 }, // EA Number
        { wch: 18 }, // Work Order
        { wch: 20 }  // WORK ORDER STATUS
      ];
      worksheet['!cols'] = columnWidths;
    } else {
      // Create empty worksheet with headers
      worksheet = XLSX.utils.json_to_sheet([{
        'ASSET NO': '',
        'YEAR': '',
        'MODEL': '',
        'MANUFACTURER': '',
        'STATION': '',
        'VIN': '',
        'Ford Recall Number': '',
        'Type': '',
        'EA Number': '',
        'Work Order': '',
        'WORK ORDER STATUS': ''
      }]);
      const columnWidths = [
        { wch: 18 }, { wch: 8 }, { wch: 15 }, { wch: 15 }, { wch: 20 },
        { wch: 20 }, { wch: 25 }, { wch: 12 }, { wch: 25 }, { wch: 18 }, { wch: 20 }
      ];
      worksheet['!cols'] = columnWidths;
    }

    // SHEET 3: Needs EA - Recalls without EA numbers
    const needsEAData = [];
    const needsEASet = new Set(); // Track unique recall+type combinations
    
    for (const row of excelData) {
      const recallNumber = row['Ford Recall Number'];
      const recallType = row['Type'] || 'Recall';
      const eaNumber = row['EA Number'] || '';
      
      // Check if EA Number is missing or 'NONE'
      const hasNoEA = !eaNumber || 
                      eaNumber.toString().trim() === '' || 
                      eaNumber.toString().trim().toUpperCase() === 'NONE';
      
      if (hasNoEA && recallNumber) {
        // Create unique key for recall number + type combination
        const uniqueKey = `${recallNumber}|${recallType}`;
        
        // Only add if we haven't seen this combination before
        if (!needsEASet.has(uniqueKey)) {
          needsEASet.add(uniqueKey);
          needsEAData.push({
            'Ford Recall Number': recallNumber,
            'Type': recallType
          });
        }
      }
    }
    
    // Sort by recall number
    needsEAData.sort((a, b) => {
      const recallCompare = (a['Ford Recall Number'] || '').localeCompare(b['Ford Recall Number'] || '');
      if (recallCompare !== 0) return recallCompare;
      return (a['Type'] || '').localeCompare(b['Type'] || '');
    });
    
    let needsEAWorksheet;
    if (needsEAData.length > 0) {
      needsEAWorksheet = XLSX.utils.json_to_sheet(needsEAData);
      needsEAWorksheet['!cols'] = [
        { wch: 25 }, // Ford Recall Number
        { wch: 12 }  // Type
      ];
    } else {
      // Create empty worksheet with headers
      needsEAWorksheet = XLSX.utils.json_to_sheet([{
        'Ford Recall Number': '',
        'Type': ''
      }]);
      needsEAWorksheet['!cols'] = [
        { wch: 25 }, // Ford Recall Number
        { wch: 12 }  // Type
      ];
    }
    
    // SHEET 4: Needs WO - Vehicles with EA but no Work Order
    const needsWOData = [];
    
    for (const row of excelData) {
      const eaNumber = row['EA Number'] || '';
      const workOrder = row['Work Order'] || '';
      
      // Check if EA Number exists and is not 'NONE', but Work Order is missing
      const hasEA = eaNumber && 
                    eaNumber.toString().trim() !== '' && 
                    eaNumber.toString().trim().toUpperCase() !== 'NONE';
      const hasNoWO = !workOrder || 
                       workOrder.toString().trim() === '' || 
                       workOrder.toString().trim().toUpperCase() === 'NONE';
      
      if (hasEA && hasNoWO) {
        needsWOData.push({
          'EA Number': eaNumber,
          'Asset NO': row['ASSET NO'] || '',
          'YEAR': row['YEAR'] || '',
          'MODEL': row['MODEL'] || '',
          'STATION': row['STATION'] || '',
          'VIN': row['VIN'] || ''
        });
      }
    }
    
    // Sort by EA Number, then by Asset NO
    needsWOData.sort((a, b) => {
      const eaCompare = (a['EA Number'] || '').localeCompare(b['EA Number'] || '');
      if (eaCompare !== 0) return eaCompare;
      return (a['Asset NO'] || '').localeCompare(b['Asset NO'] || '');
    });
    
    let needsWOWorksheet;
    if (needsWOData.length > 0) {
      needsWOWorksheet = XLSX.utils.json_to_sheet(needsWOData);
      needsWOWorksheet['!cols'] = [
        { wch: 25 }, // EA Number
        { wch: 18 }, // Asset NO
        { wch: 8 },  // YEAR
        { wch: 15 }, // MODEL
        { wch: 20 }, // STATION
        { wch: 20 }  // VIN
      ];
    } else {
      // Create empty worksheet with headers
      needsWOWorksheet = XLSX.utils.json_to_sheet([{
        'EA Number': '',
        'Asset NO': '',
        'YEAR': '',
        'MODEL': '',
        'STATION': '',
        'VIN': ''
      }]);
      needsWOWorksheet['!cols'] = [
        { wch: 25 }, // EA Number
        { wch: 18 }, // Asset NO
        { wch: 8 },  // YEAR
        { wch: 15 }, // MODEL
        { wch: 20 }, // STATION
        { wch: 20 }  // VIN
      ];
    }

    // Add all worksheets to workbook
    XLSX.utils.book_append_sheet(workbook, groupedWorksheet, 'Grouped by Recall');
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Recall Data');
    XLSX.utils.book_append_sheet(workbook, needsEAWorksheet, 'Needs EA');
    XLSX.utils.book_append_sheet(workbook, needsWOWorksheet, 'Needs WO');

    // Write file
    XLSX.writeFile(workbook, outputPath);
    
    console.log(`\n=== EXCEL FILE SUMMARY ===`);
    console.log(`✅ VINs with recalls included: ${excelData.length}`);
    console.log(`❌ VINs without recalls skipped: ${scrapedData.length - excelData.length}`);
    console.log(`📊 Total VINs processed: ${scrapedData.length}`);
    console.log(`📋 Unique recall numbers: ${recallGroups.size}`);
    console.log(`📁 Output Excel file created: ${outputPath}`);
    console.log(`   - Sheet 1: "Grouped by Recall" (Ford Recall Number first, grouped by recall)`);
    console.log(`   - Sheet 2: "Recall Data" (original format)`);
    console.log(`   - Sheet 3: "Needs EA" (${needsEAData.length} recall/satisfaction numbers without EA)`);
    console.log(`   - Sheet 4: "Needs WO" (${needsWOData.length} vehicles with EA but no Work Order)`);
    
  } catch (error) {
    console.error('Error creating output Excel file:', error);
    throw error;
  }
}

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Make sure to install Playwright browsers: npx playwright install');
});
