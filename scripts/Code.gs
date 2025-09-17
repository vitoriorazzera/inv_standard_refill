// === MAIN MENU SETUP ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('W1LD Actions')
    .addItem('01_Reset Export Flags', 'resetExportedFlags')
    .addItem('02_Download orders as CSV', 'downloadOrdersAsCSV')
    .addItem('03_Send POs to Suppliers', 'sendSuppliersReports_FromCsvPdfTemplate')
    .addToUi();
}


// === RESET EXPORT FLAGS ===
function resetExportedFlags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = [
    { sheetName: "PO Template", rangeFunc: sheet => sheet.getLastRow() >= 2 ? sheet.getRange("V2:V" + sheet.getLastRow()) : null },
    { sheetName: "Refill_Calculator", rangeFunc: sheet => sheet.getLastRow() >= 3 ? sheet.getRange("N3:N" + sheet.getLastRow()) : null },
    { sheetName: "Emails", rangeFunc: sheet => sheet.getLastRow() >= 2 ? sheet.getRange("D2:D" + sheet.getLastRow()) : null }
  ];
  tasks.forEach(({ sheetName, rangeFunc }) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const range = rangeFunc(sheet);
      if (range) range.clearContent();
    } else {
      console.warn(`Sheet "${sheetName}" not found.`);
    }
  });
  showAlert("Export flags and relevant sheet columns cleared.");
}


// === CONFIG ===
const INCLUDE_PDF = true;
const CSV_PDF_SHEET_NAME = 'CSV-PDF Template';
const EMAILS_SHEET_NAME = 'Emails';


// === SEND POs TO SUPPLIERS ===
function sendSuppliersReports_FromCsvPdfTemplate() {
  return sendSuppliersReports_BySupplierAndPO();
}


function sendSuppliersReports_BySupplierAndPO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tpl = ss.getSheetByName(CSV_PDF_SHEET_NAME);
  const emails = ss.getSheetByName(EMAILS_SHEET_NAME);
  if (!tpl || !emails) return showAlert('Missing CSV-PDF Template or Emails');


  const [eh, ...eRows] = emails.getDataRange().getValues();
  const idx = {
    name: eh.indexOf('Name'),
    email: eh.indexOf('Email'),
    status: eh.indexOf('Status'),
    supplierName: eh.indexOf('SupplierName')
  };
  for (const [k, v] of Object.entries(idx)) if (v < 0) return showAlert('Emails sheet missing ' + k);


  const [th, ...tRows] = tpl.getDataRange().getDisplayValues();
  const headerCI = th.slice(2, 9); // C to I
  const COL = { PO: 0, SUP: 1 };

  // Quote only the 3rd column to protect commas and quotes
  function protectThirdColumn(row) {
    const c = row.slice();
    c[2] = `"${String(c[2] ?? "").replace(/\r\n?|\n/g, " ").replace(/"/g, '""')}"`;
    return c;
  }


  const groups = new Map();
  for (const r of tRows) {
    const poOriginal = (r[COL.PO] || '').toString().trim();
    const sup = toKey(r[COL.SUP]);
    if (!poOriginal || !sup) continue;
    if (!groups.has(sup)) groups.set(sup, new Map());
    const byPO = groups.get(sup);
    if (!byPO.has(poOriginal)) byPO.set(poOriginal, []);
    byPO.get(poOriginal).push(r.slice(2, 9));
  }


  let sent = 0;
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');


  for (let i = 0; i < eRows.length; i++) {
    const er = eRows[i];
    const supId = toKey(er[idx.name]);
    const email = er[idx.email];
    const status = er[idx.status];
    const supplierName = cleanSupplierName(er[idx.supplierName] || '');
    if (!supId || !email) { logSkip(i + 2, 'missing id or email'); continue; }
    if (status) { logSkip(i + 2, 'already sent'); continue; }


    const byPO = groups.get(supId);
    if (!byPO) { logSkip(i + 2, 'no rows for supplier ' + supId); continue; }


    for (const [poNumber, rows] of byPO.entries()) {
      const csvText = [headerCI, ...rows.map(protectThirdColumn)]
        .map(r => r.join(','))
        .join('\n');
      const csvBlob = Utilities.newBlob(csvText, 'text/csv', `${supId}_${poNumber}.csv`);
      const atts = [csvBlob];


      if (INCLUDE_PDF) {
        const pdfBlob = createStyledPdfFromHtml(poNumber, supplierName, headerCI, rows);
        atts.push(pdfBlob);
      }


      const html = `
        <html><body style="font-family:Arial,sans-serif;font-size:14px">
          <p>Dear ${supplierName} team,</p>
          <p>Attached is purchase order ${poNumber} in CSV${INCLUDE_PDF ? ' and PDF' : ''}.</p>
          <p>Questions to <a href="mailto:orders@wildearth.com.au">orders@wildearth.com.au</a></p>
          <p>Best regards<br><strong>Wild Earth</strong></p>
          <hr><p style="font-size:11px;color:#777"><em>Automated message</em></p>
        </body></html>`;


      GmailApp.sendEmail(email, `Wild Earth Purchase Order ${poNumber}`, '', {
        htmlBody: html,
        attachments: atts
      });


      sent++;
      console.log(`Sent PO ${poNumber} to ${email} for supplier ${supId}`);
    }


    emails.getRange(i + 2, idx.status + 1).setValue(`Email sent ${nowStr}`);
  }


  showAlert(`${sent} email(s) sent.`);
}


// === PDF GENERATION HELPER ===
function createStyledPdfFromHtml(poNumber, supplierName, headers, rows) {
  let subtotal = 0;
  let gstTotal = 0;


  // Parse columns correctly by header position
  const headerMap = headers.reduce((map, h, i) => {
    map[h.trim().toLowerCase()] = i;
    return map;
  }, {});


  const idxLineTotal = headerMap['line total'];
  const idxGst = headerMap['gst'];


  if (idxLineTotal === undefined || idxGst === undefined) {
    throw new Error('Missing "Line Total" or "GST" columns in header');
  }


  rows.forEach(row => {
    const lineTotal = parseFloat((row[idxLineTotal] || '').toString().replace(/[^0-9.]/g, '')) || 0;
    const gst = parseFloat((row[idxGst] || '').toString().replace(/[^0-9.]/g, '')) || 0;


    subtotal += lineTotal;
    gstTotal += gst;
  });


  const grandTotal = subtotal + gstTotal;
  const format = (num) => `$${num.toFixed(2)}`;


  const htmlContent = `
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            font-size: 11px;
            margin: 20px;
            color: #333;
          }


          .header {
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
          }


          .info {
            font-size: 12px;
            line-height: 1.5;
          }


          .summary {
            margin-top: 25px;
            font-size: 12px;
            text-align: right;
            line-height: 1.6;
          }


          table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 12px;
            font-size: 10px;
            color: #333;
          }


          th {
            background-color: #f8f8f8;
            font-weight: 600;
            text-align: left;
            padding: 6px 8px;
            border-bottom: 1px solid #ccc;
            white-space: nowrap;
          }


          td {
            padding: 5px 8px;
            border-bottom: 1px solid #eee;
            vertical-align: top;
          }


          tr:hover {
            background-color: #f4f4f4;
          }


          table thead tr {
            border-bottom: 2px solid #ccc;
          }


          table tbody tr:last-child td {
            border-bottom: none;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="info">
            <h2>Purchase Order</h2>
            <p><strong>PO Number:</strong> ${poNumber}</p>
            <p><strong>Supplier:</strong> ${supplierName}</p>
            <p><strong>Date:</strong> ${formatDate(new Date())}</p>
          </div>
          <div class="info">
            <h3>Wild Earth</h3>
            <p><strong>Phone:</strong> 0755934180</p>
            <p><strong>Website:</strong> wildearth.com.au</p>
            <p><strong>Deliver To:</strong><br>
              Unit 2 - 27 Central Drive<br>
              Burleigh Heads QLD 4220
            </p>
          </div>
        </div>


        <table>
          <thead>
            <tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${rows.map(r =>
              `<tr>${r.map(c => `<td>${c ?? ''}</td>`).join('')}</tr>`
            ).join('')}
          </tbody>
        </table>


        <div class="summary">
          <p><strong>Subtotal:</strong> ${format(subtotal)}</p>
          <p><strong>GST:</strong> ${format(gstTotal)}</p>
          <p><strong>Total:</strong> ${format(grandTotal)}</p>
        </div>
      </body>
    </html>
  `;


  const blob = Utilities.newBlob(htmlContent, 'text/html', `PO_${poNumber}.html`);
  const file = DriveApp.createFile(blob).setName(`PO_${poNumber}.html`);
  const pdf = file.getAs('application/pdf').setName(`PO_${poNumber}.pdf`);
  file.setTrashed(true);


  return pdf;
}


function getColumnSum(rows, colIndex) {
  return rows.reduce((sum, r) => {
    const val = parseFloat(r[colIndex]);
    return sum + (isNaN(val) ? 0 : val);
  }, 0);
}




// === UTILS ===
function toKey(v) { return (v == null ? '' : v.toString()).trim().toLowerCase(); }
function cleanSupplierName(name) { return name.replace(/\s*\((BO|NBO)?\)\s*$/, '').trim(); }
function showAlert(msg) { SpreadsheetApp.getUi().alert(msg); }
function logSkip(row, why) { console.log(`Skip Emails row ${row} ${why}`); }

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}


// === FUNCTION 2: Download Orders as CSV (H to U only) ===
function downloadOrdersAsCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PO Template');
  if (!sheet) {
    showAlert('Sheet "PO Template" not found.');
    return;
  }


  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    showAlert('No data to export.');
    return;
  }


  // Filter unexported rows (Column V = index 21)
  const header = data[0].slice(7, 21); // Columns H to U
  const rows = data.slice(1).filter(r => !r[21]);


  if (rows.length === 0) {
    showAlert('No unexported rows found.');
    return;
  }


  // Format dates to DD/MM/YYYY for trimmed rows (Columns H to U)
  const trimmedRows = rows.map(r => {
    return r.slice(7, 21).map(cell => {
      if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell)) {
        // It's a valid Date object
        return formatDate(cell);
      }
      return cell;
    });
  });



  // Build CSV content
  const csv = [header].concat(trimmedRows).map(r => r.join(',')).join('\n');
  const csvEncoded = encodeURIComponent(csv);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const filename = `Orders_Export_${timestamp}.csv`;


  const htmlContent = `
    <html><body>
      <p style="font-family:Arial;">Export complete. Your file is downloading...</p>
      <a id="download" href="data:text/csv;charset=utf-8,${csvEncoded}" download="${filename}"></a>
      <script>
        document.getElementById('download').click();
        setTimeout(() => google.script.host.close(), 1000);
      </script>
    </body></html>
  `;


  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(300)
    .setHeight(100);


  SpreadsheetApp.getUi().showModalDialog(html, 'Download Orders');


  // Mark exported rows in Column V with timestamp
  const exportCol = 22;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  for (let i = 0; i < data.length - 1; i++) {
    if (!data[i + 1][21]) {
      sheet.getRange(i + 2, exportCol).setValue(now);
    }
  }
}