const cheerio = require('cheerio');
const xl = require('excel4node');
const puppeteer = require('puppeteer');
const fs = require('fs').promises;

// Input HTML file (replace with your file path)
const htmlFile = 'school_list.html';

// Output files
const excelOutput = 'school_list.xlsx';
const pdfOutput = 'school_list.pdf';

// Step 1: Parse HTML and extract table
async function parseHtmlTable() {
  const html = await fs.readFile(htmlFile, 'utf-8');
  const $ = cheerio.load(html);

  // Find the table with id="DataGrid1"
  const table = $('#DataGrid1');
  if (!table.length) {
    throw new Error("Table with id='DataGrid1' not found in HTML");
  }

  // Extract headers
  const headers = [];
  table.find('tr.MISDataGridHeaderFont td').each((i, th) => {
    headers.push($(th).text().trim());
  });

  // Extract rows
  const rows = [];
  table.find('tr').slice(1).each((i, tr) => { // Skip header row
    const cells = [];
    $(tr).find('td').each((j, td) => {
      cells.push($(td).text().trim());
    });
    rows.push(cells);
  });

  return { headers, rows };
}

// Step 2: Convert to Excel
async function generateExcel(headers, rows) {
  const wb = new xl.Workbook();
  const ws = wb.addWorksheet('Schools');

  // Style for headers
  const headerStyle = wb.createStyle({
    font: { bold: true },
    alignment: { horizontal: 'center' },
  });

  // Write headers
  headers.forEach((header, i) => {
    ws.cell(1, i + 1).string(header).style(headerStyle);
  });

  // Write rows
  rows.forEach((row, i) => {
    row.forEach((cell, j) => {
      ws.cell(i + 2, j + 1).string(cell);
    });
  });

  // Save Excel file
  await wb.write(excelOutput);
  console.log(`Excel file saved as ${excelOutput}`);
}

// Step 3: Convert to PDF
async function generatePdf(html) {
  // Clean HTML and add custom CSS
  const cleanHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        table {
          width: 100%;
          border-collapse: collapse;
          font-family: Arial, sans-serif;
          font-size: 10px;
        }
        th, td {
          border: 1px solid black;
          padding: 5px;
          text-align: left;
        }
        th {
          background-color: #f2f2f2;
          font-weight: bold;
        }
        tr:nth-child(even) {
          background-color: #f9f9f9;
        }
        @media print {
          table { page-break-inside: auto; }
          tr { page-break-inside: avoid; page-break-after: auto; }
        }
      </style>
    </head>
    <body>
      <h1>DOE Unaided Schools List</h1>
      ${html.match(/<table[^>]*id="DataGrid1"[^>]*>[\s\S]*<\/table>/)[0]}
    </body>
    </html>
  `;

  // Launch Puppeteer
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
  const page = await browser.newPage();

  // Set HTML content
  await page.setContent(cleanHtml, { waitUntil: 'networkidle0' });

  // Generate PDF
  await page.pdf({
    path: pdfOutput,
    format: 'A4',
    margin: { top: '1cm', right: '1cm', bottom: '1cm', left: '1cm' },
    printBackground: true,
  });

  await browser.close();
  console.log(`PDF file saved as ${pdfOutput}`);
}

// Main function
async function main() {
  try {
    // Parse HTML
    const { headers, rows } = await parseHtmlTable();

    // Generate Excel
    await generateExcel(headers, rows);

    // Read HTML for PDF
    const html = await fs.readFile(htmlFile, 'utf-8');

    // Generate PDF
    await generatePdf(html);
  } catch (error) {
    console.error('Error:', error.message);
  }
}

// Run the script
main();