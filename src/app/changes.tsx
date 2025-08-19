const handleDownload = () => {
  if (sheets.length === 0) return;

  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    if (sheet.merges.length > 0) ws["!merges"] = sheet.merges;
    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  });

  XLSX.writeFile(workbook, "Edited_Sheets.xlsx");
};


<button
  className="mt-4 px-4 py-2 bg-green-600 text-white rounded"
  onClick={handleDownload}
>
  Download Edited XLSX
</button>



const handleDownloadAllCSV = () => {
  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", `${sheet.name}.csv`);
    link.click();
  });
};


const handleDownloadAllJSON = () => {
  sheets.forEach((sheet) => {
    const [headers, ...rows] = sheet.data;
    const json: Record<string, any[]> = {};
    headers.forEach((h) => (json[h as string] = []));
    rows.forEach((row) => {
      headers.forEach((h, idx) => {
        json[h as string].push(row[idx] ?? null);
      });
    });

    const blob = new Blob([JSON.stringify(json, null, 2)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", `${sheet.name}.json`);
    link.click();
  });
};


<div className="mt-4 flex space-x-2">
  <button
    className="px-4 py-2 bg-blue-600 text-white rounded"
    onClick={handleDownloadAllCSV}
  >
    Download All CSV
  </button>

  <button
    className="px-4 py-2 bg-yellow-500 text-white rounded"
    onClick={handleDownloadAllJSON}
  >
    Download All JSON
  </button>
</div>



const handleDownloadStyledXLSX = () => {
  if (sheets.length === 0) return;

  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);

    // Apply merges
    if (sheet.merges.length > 0) ws["!merges"] = sheet.merges;

    // Apply header style (row 1)
    const headers = sheet.data[0];
    headers.forEach((_, cIdx) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: cIdx });
      if (!ws[cellAddress]) ws[cellAddress] = { t: "s", v: headers[cIdx] };
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFFFF" } }, // white font
        fill: { fgColor: { rgb: "FF1F77B4" } }, // blue background
        alignment: { horizontal: "center", vertical: "center" },
      };
    });

    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  });

  XLSX.writeFile(workbook, "Edited_Styled_Sheets.xlsx");
};


npm install exceljs


import ExcelJS from "exceljs";

const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // Add rows
    sheet.data.forEach((row) => ws.addRow(row));

    // Style header row
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};



import ExcelJS from "exceljs";

const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // Add rows
    sheet.data.forEach((row) => ws.addRow(row));

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // Apply merges
    sheet.merges.forEach((merge) => {
      const startRow = merge.s.r + 1; // ExcelJS is 1-indexed
      const startCol = merge.s.c + 1;
      const endRow = merge.e.r + 1;
      const endCol = merge.e.c + 1;
      ws.mergeCells(startRow, startCol, endRow, endCol);
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};


import ExcelJS from "exceljs";

const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      ws.addRow(normalizedRow);
    });

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // Apply merges (ExcelJS is 1-indexed)
    sheet.merges.forEach((merge) => {
      const startRow = merge.s.r + 1;
      const startCol = merge.s.c + 1;
      const endRow = merge.e.r + 1;
      const endCol = merge.e.c + 1;
      ws.mergeCells(startRow, startCol, endRow, endCol);
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};



const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      ws.addRow(normalizedRow);
    });

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // 👉 Header row height
    ws.getRow(1).height = 25;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 20 });

    // 👉 Conditional formatting for "Probability" column
    const headers = sheet.data[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "probability"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "HIGH") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MEDIUM") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "LOW") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed)
    sheet.merges.forEach((merge) => {
      const startRow = merge.s.r + 1;
      const startCol = merge.s.c + 1;
      const endRow = merge.e.r + 1;
      const endCol = merge.e.c + 1;
      ws.mergeCells(startRow, startCol, endRow, endCol);
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};



const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      ws.addRow(normalizedRow);
    });

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // 👉 Header row height
    ws.getRow(1).height = 25;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "Probability" column
    const headers = sheet.data[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "probability"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "HIGH") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MEDIUM") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "LOW") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // 👉 Conditional formatting for "Impact" column
    const impactIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "impact"
    );
    if (impactIndex !== -1) {
      ws.getColumn(impactIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "HIGH") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MEDIUM") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "LOW") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed)
    sheet.merges.forEach((merge) => {
      const startRow = merge.s.r + 1;
      const startCol = merge.s.c + 1;
      const endRow = merge.e.r + 1;
      const endCol = merge.e.c + 1;
      ws.mergeCells(startRow, startCol, endRow, endCol);
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};



const handleDownloadStyledXLSX = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // default row height for the whole sheet
    ws.properties.defaultRowHeight = 60;

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      ws.addRow(normalizedRow);
    });

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    // 👉 Header row height
    ws.getRow(1).height = 60;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "Probability" column
    const headers = sheet.data[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "probability"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "HIGH") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MEDIUM") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "LOW") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // 👉 Conditional formatting for "Impact" column
    const impactIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "impact"
    );
    if (impactIndex !== -1) {
      ws.getColumn(impactIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "HIGH") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MEDIUM") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "LOW") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed)
    sheet.merges.forEach((merge) => {
      const startRow = merge.s.r + 1;
      const startCol = merge.s.c + 1;
      const endRow = merge.e.r + 1;
      const endCol = merge.e.c + 1;
      ws.mergeCells(startRow, startCol, endRow, endCol);
    });
  });

  // Generate XLSX blob
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.setAttribute("download", "Edited_Styled_Sheets.xlsx");
  link.click();
};

