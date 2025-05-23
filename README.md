# HTML Table to PDF & Excel Parser

A lightweight Node.js tool to convert HTML tables into Excel (.xlsx) and styled, paginated PDF files. Built with `cheerio` for HTML parsing, `excel4node` for Excel generation, and `puppeteer` for high-quality PDF output. Ideal for automating data extraction and reporting in web or desktop applications.

## Features
- Parses HTML tables with any number of rows and columns.
- Generates structured Excel files with bold headers.
- Creates A4-sized PDFs with custom CSS, pagination, and alternating row styles.
- Optimized for large datasets and seamless integration.

## Prerequisites
- **Node.js** (v16 or later)
- **Puppeteer Dependencies** (for PDF generation):
  - **Linux**: Install Chrome dependencies:
    ```bash
    sudo apt-get update
    sudo apt-get install -y libx11-xcb1 libxcomposite1 libxdamage1 libxi6 libxtst6 libnss3 libcups2 libxss1 libxrandr2 libasound2 libpangocairo-1.0-0 libatk1.0-0 libatk-bridge2.0-0 libgtk-3-0
    ```
  - **Windows/macOS**: Puppeteer automatically installs Chromium.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/html-table-to-pdf-excel.git
   cd html-table-to-pdf-excel
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Ensure your HTML file (e.g., `school_list.html`) contains a table with `id="DataGrid1"`.

## Usage
1. Place your HTML file in the project directory (default: `school_list.html`).
2. Run the script:
   ```bash
   node html_to_excel_pdf.js
   ```
3. Outputs:
   - **Excel**: `school_list.xlsx` (table data in a spreadsheet).
   - **PDF**: `school_list.pdf` (styled table in A4 format).

## Script Details
- **Input**: HTML file with a table (`id="DataGrid1"`).
- **Libraries**:
  - `cheerio`: Parses HTML and extracts table data.
  - `excel4node`: Creates Excel files with formatted headers.
  - `puppeteer`: Renders HTML to PDF with custom CSS.
- **Customization**:
  - Modify CSS in `html_to_excel_pdf.js` for PDF styling (e.g., fonts, colors).
  - Adjust Excel formatting (e.g., column widths) in the `generateExcel` function.

## Example HTML Input
```html
<table id="DataGrid1" class="MISDataGridBody" border="1">
  <tr class="MISDataGridHeaderFont">
    <td>S.No</td><td>Name</td><td>Address</td>
  </tr>
  <tr>
    <td>1</td><td>John Doe</td><td>123 Main St</td>
  </tr>
</table>
```

## Output
- **Excel**: `school_list.xlsx` with headers and data.
- **PDF**: `school_list.pdf` with a styled, paginated table and a title.

## Deployment on EC2
1. Copy files to your EC2 server.
2. Install Node.js:
   ```bash
   curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
   sudo apt-get install -y nodejs
   ```
3. Install dependencies and run:
   ```bash
   npm install
   node html_to_excel_pdf.js
   ```
4. Optionally, upload outputs to AWS S3:
   ```javascript
   const AWS = require('aws-sdk');
   const s3 = new AWS.S3();
   s3.upload({ Bucket: 'your-bucket', Key: 'school_list.xlsx', Body: require('fs').readFileSync('school_list.xlsx') }).promise();
   ```

## Troubleshooting
- **Table Not Found**: Ensure the table has `id="DataGrid1"`. Update the script if the ID differs.
- **PDF Issues**: Adjust CSS (e.g., reduce `font-size`) or check Puppeteer logs for memory errors.
- **Excel Errors**: Verify `excel4node` compatibility (`^3.0.0`).
- **EC2**: Ensure Chrome dependencies are installed for Puppeteer.

## Contributing
Feel free to submit issues or pull requests for enhancements, such as additional formatting options or alternative libraries.

## License
[MIT License](LICENSE)

---
*Last updated: May 17, 2025*