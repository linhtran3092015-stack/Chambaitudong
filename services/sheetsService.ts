/**
 * Google Sheets API Helper
 */

/**
 * Extracts spreadsheet ID from any valid Google Sheets URL
 */
export function extractSpreadsheetId(url: string): string | null {
  if (!url) return null;
  // Match standard link format
  // https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit...
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  // Fallback if they just paste the ID
  if (/^[a-zA-Z0-9-_]{30,}$/.test(url)) {
    return url;
  }
  return null;
}

/**
 * Fetches the tab/sheet names of the Google Spreadsheet
 */
export async function fetchSpreadsheetSheets(spreadsheetId: string, accessToken: string): Promise<string[]> {
  try {
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties.title`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: 'application/json',
        },
      }
    );

    if (!response.ok) {
      const errBody = await response.json().catch(() => ({}));
      throw new Error(errBody.error?.message || `Failed to fetch sheets structure (${response.status})`);
    }

    const data = await response.json();
    if (!data.sheets) return [];
    
    return data.sheets.map((s: any) => s.properties?.title || '').filter(Boolean);
  } catch (err: any) {
    console.error('Error in fetchSpreadsheetSheets:', err);
    throw err;
  }
}

/**
 * Fetches sheet data for a specific sheet/tab and formats it as a TSV string
 */
export async function fetchSheetDataAsTsv(
  spreadsheetId: string,
  sheetTitle: string,
  accessToken: string
): Promise<string> {
  try {
    // Range covers the whole sheet
    const encodedSheetTitle = encodeURIComponent(sheetTitle);
    const range = `${encodedSheetTitle}!A:Z`;
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueRenderOption=FORMATTED_VALUE`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: 'application/json',
        },
      }
    );

    if (!response.ok) {
      const errBody = await response.json().catch(() => ({}));
      throw new Error(errBody.error?.message || `Failed to fetch sheet values (${response.status})`);
    }

    const data = await response.json();
    const values: any[][] = data.values || [];
    if (values.length === 0) return '';

    // Convert rows to TSV string
    const tsvLines = values.map(row => {
      // Clean up column values, escaping tabs and trimming whitespace
      return row
        .map(cell => {
          if (cell === null || cell === undefined) return '';
          const strCell = String(cell);
          // Replace actual tab characters and newlines internally to avoid splitting errors
          return strCell.replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
        })
        .join('\t');
    });

    return tsvLines.join('\n');
  } catch (err: any) {
    console.error('Error in fetchSheetDataAsTsv:', err);
    throw err;
  }
}

/**
 * Helper to split a single CSV row safely, respecting quotes, escaped quotes, and alternate separators (e.g., semicolon, comma)
 */
function parseCSVRow(rowText: string, separator: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < rowText.length; i++) {
    const char = rowText[i];
    if (inQuotes) {
      if (char === '"') {
        if (i + 1 < rowText.length && rowText[i + 1] === '"') {
          // Escaped double quote
          current += '"';
          i++; // Skip the next quote
        } else {
          // Ending quote
          inQuotes = false;
        }
      } else {
        current += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === separator) {
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
  }
  result.push(current);
  return result;
}

/**
 * Parses multi-line CSV/Semicolon text and handles line breaks that are potentially enclosed inside quotes.
 */
export function convertCsvToTsv(csvText: string, separator: string = ','): string {
  if (!csvText) return '';
  
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];
    
    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          currentField += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        currentField += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === separator) {
        currentRow.push(currentField);
        currentField = '';
      } else if (char === '\n' || char === '\r') {
        currentRow.push(currentField);
        currentField = '';
        rows.push(currentRow);
        currentRow = [];
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
      } else {
        currentField += char;
      }
    }
  }
  
  if (currentField || inQuotes) {
    currentRow.push(currentField);
  }
  if (currentRow.length > 0) {
    rows.push(currentRow);
  }
  
  return rows.map(row => 
    row.map(cell => {
      const trimmed = String(cell).trim();
      // Remove any internal tab/newline characters within cells to form a valid single-line TSV record
      return trimmed.replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
    }).join('\t')
  ).join('\n');
}

/**
 * Detects format and automatically normalizes spreadsheet cell data to standardized TSV format.
 */
export function autoNormalizeToTsv(text: string): string {
  if (!text) return '';
  const trimmed = text.trim();
  if (trimmed === '') return '';

  const lines = trimmed.split('\n');
  const sampleLines = lines.slice(0, 5);
  
  // Count frequency of common delimiters
  const tabCounts = sampleLines.map(line => (line.match(/\t/g) || []).length);
  const commaCounts = sampleLines.map(line => (line.match(/,/g) || []).length);
  const semicolonCounts = sampleLines.map(line => (line.match(/;/g) || []).length);

  const avgTabs = tabCounts.reduce((sum, val) => sum + val, 0) / sampleLines.length;
  const avgCommas = commaCounts.reduce((sum, val) => sum + val, 0) / sampleLines.length;
  const avgSemicolons = semicolonCounts.reduce((sum, val) => sum + val, 0) / sampleLines.length;

  // Let's decide on the delimiter
  if (avgTabs > 0 && avgTabs >= avgCommas && avgTabs >= avgSemicolons) {
    // Already TSV, just map-trim and clean up row line breaks or carriage returns
    return lines.map(line => 
      line.split('\t').map(cell => {
        const trimmedCell = cell.trim();
        return trimmedCell.replace(/\r/g, '').replace(/\n/g, ' ');
      }).join('\t')
    ).join('\n');
  }

  if (avgSemicolons > 0 && avgSemicolons > avgCommas && avgSemicolons > avgTabs) {
    // Semicolon-delimited values
    return convertCsvToTsv(trimmed, ';');
  }

  if (avgCommas > 0) {
    // Standard Comma-separated values (CSV)
    return convertCsvToTsv(trimmed, ',');
  }

  // Fallback if no delimiter can be detected or if text is single field:
  return trimmed;
}
