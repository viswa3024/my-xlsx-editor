"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

type SheetData = {
  name: string;
  data: (string | number)[][];
};

export default function Home() {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [sheetJson, setSheetJson] = useState<Record<string, any[]>>({});
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [tempRowData, setTempRowData] = useState<Record<string, any>>({});
  const [fileUrl, setFileUrl] = useState("");

  const editableColumns = ["Product line", "Gender", "Payment", "Customer type", "Branch", "Unit price"];

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const csvText = evt.target?.result as string;

      // Parse CSV into workbook
      const workbook = XLSX.read(csvText, { type: "string" });
      const sheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[sheetName];

      const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        blankrows: true,
      }) as any;

      const sheetData: SheetData[] = [{ name: file.name.replace(".csv", ""), data }];
      setSheets(sheetData);
      setActiveSheet(0);

      convertSheetToJson(sheetData[0]);
    };
    reader.readAsText(file);
  };

   const handleLoadUrlbkp = async () => {
  if (!fileUrl) return;

  try {
    const res = await fetch(`/api/fetch-csv?url=${encodeURIComponent(fileUrl)}`);
    if (!res.ok) throw new Error("Failed to fetch CSV");
    const text = await res.text();

    const workbook = XLSX.read(text, { type: "string" });
    const sheetsData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number)[][];
        return { name, data };
      });
    setSheets(sheetsData);
    setActiveSheet(0);
  } catch (err) {
    console.error("Error loading CSV:", err);
  }
};

const handleLoadUrl = async () => {
  debugger
  if (!fileUrl) return;
  try {
    // const response = await fetch(fileUrl);
    // if (!response.ok) throw new Error("Failed to fetch CSV");
    // const csvText = await response.text();

    const res = await fetch(`/api/fetch-csv?url=${encodeURIComponent(fileUrl)}`);
    if (!res.ok) {
      alert("Failed to load file from URL");
      return;
    }

    const csvText = await res.text();

    // Parse CSV into workbook
    const workbook = XLSX.read(csvText, { type: "string" });
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];

    const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      blankrows: true,
    }) as any;

    const sheetData: SheetData[] = [{ name: fileUrl.split("/").pop() || "Sheet1", data }];
    setSheets(sheetData);
    setActiveSheet(0);

    convertSheetToJson(sheetData[0]);
  } catch (error) {
    console.error("Error loading CSV from URL:", error);
  }
};

  const convertSheetToJson = (sheet: SheetData) => {
    const [headers, ...rows] = sheet.data;
    const json: Record<string, any[]> = {};
    headers.forEach((h) => (json[h as string] = []));
    rows.forEach((row) => {
      headers.forEach((h, idx) => {
        json[h as string].push(row[idx] ?? null);
      });
    });
    setSheetJson(json);
  };

  const updateJsonValue = (header: string, rowIndex: number, value: any) => {
    const newJson = { ...sheetJson };
    newJson[header][rowIndex] = value;
    setSheetJson(newJson);
  };

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

  const startEditRow = (rIdx: number) => {
    setEditingRow(rIdx);
    const tempData: Record<string, any> = {};
    editableColumns.forEach((col) => {
      tempData[col] = sheetJson[col]?.[rIdx - 1] ?? "";
    });
    setTempRowData(tempData);
  };

  const saveRow = (rIdx: number) => {
    editableColumns.forEach((col) => {
      updateJsonValue(col, rIdx - 1, tempRowData[col]);
      const colIndex = sheets[activeSheet].data[0].findIndex((h) => h === col);
      if (colIndex >= 0) {
        sheets[activeSheet].data[rIdx][colIndex] = tempRowData[col];
      }
    });
    setEditingRow(null);
    setTempRowData({});
  };

  const closeEdit = () => {
    setEditingRow(null);
    setTempRowData({});
  };


  const handleDownloadbkp = () => {
    const wb = XLSX.utils.book_new();
    sheets.forEach((sheet) => {
      const ws = XLSX.utils.aoa_to_sheet(sheet.data);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(wb, "edited.xlsx");
  };

  const handleDownload = () => {
  const wb = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    // Add Project Title & Description on top
    const projectTitle = ["Project Title:", "AI-based Generative QA System"];
    const projectDescription = ["Project Description:", "Fine-tuned models for QA and email subject generation"];

    // Combine title, description, and actual sheet data
    const newData = [projectTitle, projectDescription, [], ...sheet.data];

    const ws = XLSX.utils.aoa_to_sheet(newData);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });

  XLSX.writeFile(wb, "edited.xlsx");
};

const handleDownloadbkp2 = () => {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    // Add Project Title and Description
    const dataWithMeta = [
      ["Project Title", "", "", "", ""],
      ["Project Description", "", "", "", ""],
      ...sheet.data,
    ];

    const ws = XLSX.utils.aoa_to_sheet(dataWithMeta);

    // Merge cells for Project Title (first row across 5 columns)
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Merge A1:E1
      { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } }, // Merge A2:E2
    ];

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });
  XLSX.writeFile(wb, "edited.xlsx");
};

const handleDownloadbkp4 = () => {
  const wb = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["AI-based Generative QA System"], // Project Title
      ["Fine-tuned models for QA and email subject generation"], // Project Description
      [],
      ...sheet.data,
    ]);

    // Find number of columns to merge title row properly
    const colCount = Math.max(...sheet.data.map((r) => r.length));

    if (!ws["!merges"]) ws["!merges"] = [];

    ws["!merges"].push({
      s: { r: 0, c: 0 }, // start cell
      e: { r: 0, c: colCount - 1 }, // end cell
    });

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });

  XLSX.writeFile(wb, "edited.xlsx");
};


const handleDownloadbkp1 = () => {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet([]);

    // project title row with 2-colspan
    XLSX.utils.sheet_add_aoa(ws, [["Project Title:", "AI-based Generative QA System"]], { origin: "A1" });
    ws["!merges"] = ws["!merges"] || [];
    ws["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }); // merge A1:B1

    // project description row with 2-colspan
    XLSX.utils.sheet_add_aoa(ws, [["Project Description:", "Fine-tuned models for QA and email subject generation"]], { origin: "A2" });
    ws["!merges"].push({ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }); // merge A2:B2

    // push actual sheet data starting from row 4
    XLSX.utils.sheet_add_aoa(ws, sheet.data, { origin: "A4" });

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });
  XLSX.writeFile(wb, "edited.xlsx");
};


  const handleDownloadStyledXLSXbkp = async () => {
    if (sheets.length === 0) return;

    const workbook = new ExcelJS.Workbook();

    sheets.forEach((sheet) => {
      const ws = workbook.addWorksheet(sheet.name);

      ws.properties.defaultRowHeight = 50;

      sheet.data.forEach((row, rowIndex) => {
        const newRow = ws.addRow(row);
        newRow.height = 50;

        row.forEach((_, colIndex) => {
          const cell = newRow.getCell(colIndex + 1);
          cell.alignment = { vertical: "middle", horizontal: "center" };
          cell.font = { size: 12, name: "Arial" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });
    });

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buf]), "styled_edited.xlsx");
  };


  const handleDownloadStyledXLSXbkp1 = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // default row height for the whole sheet
    ws.properties.defaultRowHeight = 50;

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      const addedRow = ws.addRow(normalizedRow);

      // 👉 Apply wrap text for all row cells
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      });
    });

    // Apply header styles
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    });

    // 👉 Header row height
    ws.getRow(1).height = 50;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "Probability" column
    const headers = sheet.data[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "branch"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "A") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "B") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        } else if (val === "C") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // green
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // 👉 Conditional formatting for "Impact" column
    const impactIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "gender"
    );
    if (impactIndex !== -1) {
      ws.getColumn(impactIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "FEMALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }; // red
          cell.font = { color: { argb: "FFFFFFFF" } }; // white text
        } else if (val === "MALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }; // yellow
          cell.font = { color: { argb: "FF000000" } }; // black text
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed)
    if ((sheet as any).merges) {
      (sheet as any).merges.forEach((merge: any) => {
        const startRow = merge.s.r + 1;
        const startCol = merge.s.c + 1;
        const endRow = merge.e.r + 1;
        const endCol = merge.e.c + 1;
        ws.mergeCells(startRow, startCol, endRow, endCol);
      });
    }
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


const handleDownloadStyledXLSXwr1 = async () => {
  if (sheets.length === 0) return;

  const workbook = new ExcelJS.Workbook();

  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);

    // 👉 Project Title
    ws.addRow([`Project Title: My Sample Project`]);
    ws.mergeCells(1, 1, 1, 5);
    ws.getRow(1).font = { bold: true, size: 16, color: { argb: "FF000000" } };
    ws.getRow(1).alignment = { vertical: "middle", horizontal: "center" };
    ws.getRow(1).height = 30;

    // 👉 Project Description
    ws.addRow([`Project Description: This is a demo project description text`]);
    ws.mergeCells(2, 1, 2, 5);
    ws.getRow(2).font = { italic: true, size: 12, color: { argb: "FF333333" } };
    ws.getRow(2).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.getRow(2).height = 40;

    // Add an empty spacer row
    ws.addRow([]);

    // default row height for the whole sheet
    ws.properties.defaultRowHeight = 50;

    // 👉 Find max columns count
    const maxCols = Math.max(...sheet.data.map((row) => row.length));

    // Add rows with padding
    sheet.data.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      const addedRow = ws.addRow(normalizedRow);

      // 👉 Apply wrap text for all row cells
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      });
    });

    // Apply header styles (row 4 because title+description+spacer used 3 rows)
    const headerRowIndex = 4;
    ws.getRow(headerRowIndex).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    });

    // 👉 Header row height
    ws.getRow(headerRowIndex).height = 50;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "Branch" column
    const headers = sheet.data[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "branch"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === headerRowIndex) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "A") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
          cell.font = { color: { argb: "FFFFFFFF" } };
        } else if (val === "B") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { color: { argb: "FF000000" } };
        } else if (val === "C") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
          cell.font = { color: { argb: "FF000000" } };
        }
      });
    }

    // 👉 Conditional formatting for "Gender" column
    const impactIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "gender"
    );
    if (impactIndex !== -1) {
      ws.getColumn(impactIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === headerRowIndex) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "FEMALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
          cell.font = { color: { argb: "FFFFFFFF" } };
        } else if (val === "MALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { color: { argb: "FF000000" } };
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed)
    if ((sheet as any).merges) {
      (sheet as any).merges.forEach((merge: any) => {
        const startRow = merge.s.r + 4; // shift merges because of 3 extra rows
        const startCol = merge.s.c + 1;
        const endRow = merge.e.r + 4;
        const endCol = merge.e.c + 1;
        ws.mergeCells(startRow, startCol, endRow, endCol);
      });
    }
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
    ws.properties.defaultRowHeight = 50;

    // 👉 Remove "Rating" column from sheet data
    const ratingIndex = sheet.data[0].findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "rating"
    );
    const filteredData =
      ratingIndex !== -1
        ? sheet.data.map((row) => row.filter((_, idx) => idx !== ratingIndex))
        : sheet.data;

    // 👉 Find max columns count
    const maxCols = Math.max(...filteredData.map((row) => row.length));

    // Add rows with padding (main data first)
    filteredData.forEach((row) => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxCols) {
        normalizedRow.push(""); // pad empty cells
      }
      const addedRow = ws.addRow(normalizedRow);

      // 👉 Apply wrap text for all row cells
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
      });
    });

    // Apply header styles (row 1 since now no title above)
    const headerRowIndex = 1;
    ws.getRow(headerRowIndex).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // white font
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F77B4" } }; // blue background
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    });

    // 👉 Header row height
    ws.getRow(headerRowIndex).height = 50;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "Branch" column
    const headers = filteredData[0];
    const probabilityIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "branch"
    );
    if (probabilityIndex !== -1) {
      ws.getColumn(probabilityIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === headerRowIndex) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "A") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
          cell.font = { color: { argb: "FFFFFFFF" } };
        } else if (val === "B") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { color: { argb: "FF000000" } };
        } else if (val === "C") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
          cell.font = { color: { argb: "FF000000" } };
        }
      });
    }

    // 👉 Conditional formatting for "Gender" column
    const impactIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "gender"
    );
    if (impactIndex !== -1) {
      ws.getColumn(impactIndex + 1).eachCell((cell, rowNumber) => {
        if (rowNumber === headerRowIndex) return; // skip header
        const val = (cell.value || "").toString().toUpperCase();
        if (val === "FEMALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
          cell.font = { color: { argb: "FFFFFFFF" } };
        } else if (val === "MALE") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { color: { argb: "FF000000" } };
        }
      });
    }

    // Apply merges (ExcelJS is 1-indexed, no shift now)
    if ((sheet as any).merges) {
      (sheet as any).merges.forEach((merge: any) => {
        const startRow = merge.s.r + 1;
        const startCol = merge.s.c + 1;
        const endRow = merge.e.r + 1;
        const endCol = merge.e.c + 1;
        ws.mergeCells(startRow, startCol, endRow, endCol);
      });
    }

    // 👉 Spacer row before project info
    ws.addRow([]);

    const titleRowIndex = (ws.lastRow ? ws.lastRow.number : 0) + 1;
    console.log("titleRowIndex: ", titleRowIndex);
    ws.addRow([`Project Title:\nAI-based Generative QA System`]);
    ws.mergeCells(titleRowIndex, 1, titleRowIndex, maxCols);
    ws.getRow(titleRowIndex).font = { bold: true, size: 16, color: { argb: "FF000000" } };
    ws.getRow(titleRowIndex).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    ws.getRow(titleRowIndex).height = 40;

    // 👉 Project Description (after data, with 2 lines)
    const descRowIndex = (ws.lastRow ? ws.lastRow.number : 0) + 1;
    console.log("descRowIndex: ", descRowIndex);
    ws.addRow([`Project Description:\nFine-tuned models for QA and email subject generation`]);
    ws.mergeCells(descRowIndex, 1, descRowIndex, maxCols);
    ws.getRow(descRowIndex).font = { italic: true, size: 12, color: { argb: "FF333333" } };
    ws.getRow(descRowIndex).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    ws.getRow(descRowIndex).height = 50;
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





  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 CSV Editor (Editable Risk Columns with Row Edit)</h1>

      <input type="file" accept=".csv" onChange={handleFileUpload} className="mb-4" />

      <br />

      <div className="mb-4 flex gap-2">
        <input
          type="text"
          placeholder="Enter CSV file URL"
          value={fileUrl}
          onChange={(e) => setFileUrl(e.target.value)}
          className="border p-2 rounded w-full"
        />
        <button
          onClick={handleLoadUrl}
          className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          Load
        </button>
      </div>

      {sheets.length > 0 && (
        <>
          {/* Sheet Tabs */}
          <div className="flex space-x-2 mb-4">
            {sheets.map((sheet, idx) => (
              <button
                key={idx}
                onClick={() => {
                  setActiveSheet(idx);
                  convertSheetToJson(sheets[idx]);
                }}
                className={`px-4 py-2 rounded ${
                  idx === activeSheet ? "bg-blue-600 text-white" : "bg-gray-200"
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>

          <div className="overflow-auto border rounded-lg">
            <table className="border-collapse w-full">
              <tbody>
                {sheets[activeSheet].data.map((row, rIdx) => {
                  const isHeader = rIdx === 0;
                  const headers = sheets[activeSheet].data[0];

                  return (
                    <tr key={rIdx}>
                      {row.map((cell, cIdx) => {
                        const headerName = headers[cIdx] as string;
                        const isEditable = !isHeader && editableColumns.includes(headerName);
                        const inEditMode = editingRow === rIdx;

                        return (
                          <td
                            key={cIdx}
                            className={`border p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}
                          >
                            {isHeader ? (
                              cell
                            ) : inEditMode && isEditable ? (
                              headerName === "Unit price" || headerName === "Payment" || headerName === "Customer type" ? (
                                <input
                                  type="text"
                                  className="w-full border-none p-1 focus:outline-none text-center"
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Product line" ? (
                                <textarea
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Branch" || headerName === "Impact" ? (
                                <select
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? "A"}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                >
                                  <option value="A">A</option>
                                  <option value="B">B</option>
                                  <option value="C">C</option>
                                </select>
                              ) : headerName === "Gender" ? (
                                <select
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? "Female"}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                >
                                  <option value="Female">Female</option>
                                  <option value="Male">Male</option>
                                </select>
                              ) : (
                                cell
                              )
                            ) : (
                              cell
                            )}
                          </td>
                        );
                      })}

                      {/* Last column: Edit Changes */}
                      <td className={`border p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}>
                        {isHeader
                          ? "Edit Changes"
                          : editingRow === rIdx ? (
                              <>
                                <button
                                  className="px-2 py-1 bg-green-500 text-white rounded mr-1"
                                  onClick={() => saveRow(rIdx)}
                                >
                                  Save
                                </button>
                                <button
                                  className="px-2 py-1 bg-red-500 text-white rounded"
                                  onClick={closeEdit}
                                >
                                  Close
                                </button>
                              </>
                            ) : (
                              <button
                                className="px-2 py-1 bg-blue-600 text-white rounded"
                                onClick={() => startEditRow(rIdx)}
                              >
                                Edit
                              </button>
                            )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Download buttons */}
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

             <button
            className="mt-4 px-4 py-2 bg-green-600 text-white rounded"
            onClick={handleDownload}
          >
            Download Edited XLSX
          </button>
          <button
            className="mt-4 ml-4 px-4 py-2 bg-green-600 text-white rounded"
            onClick={handleDownloadStyledXLSX}
          >
            Download Styled Edited XLSX
          </button>
          </div>
        </>
      )}
    </main>
  );
}
