"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

type SheetData = {
  name: string;
  data: (string | number)[][];
  merges: XLSX.Range[];
};

export default function Home() {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [sheetJson, setSheetJson] = useState<Record<string, any[]>>({});
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [tempRowData, setTempRowData] = useState<Record<string, any>>({});
  const [fileUrl, setFileUrl] = useState<string>("");

  const [editableColumns, setEditableColumns] = useState<string[]>(["Product line", "Risk Description", "Probability", "Impact", "Response Plan"])

  const masterEditableColumns = ["Product line", "Risk Description", "Probability", "Impact", "Response Plan"];

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const workbook = XLSX.read(bstr, { type: "binary" });

      const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true }) as any;
        const merges: XLSX.Range[] = ws["!merges"] || [];
        return { name, data, merges };
      });

      setSheets(sheetData);
      setActiveSheet(0);
       const headers = sheetData[0].data[0];
      const validEditableColumns = masterEditableColumns.filter((col) => headers.includes(col));
      setEditableColumns(validEditableColumns);

      convertSheetToJson(sheetData[0]);
    };
    reader.readAsBinaryString(file);
  };

  const handleLoadUrl = async () => {
    if (!fileUrl) return;
    try {
      // const res = await fetch(fileUrl);
      // const arrayBuffer = await res.arrayBuffer();
      // const workbook = XLSX.read(arrayBuffer, { type: "array" });

       const response = await fetch("/api/load-xlsx", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: fileUrl }),
          });
       
      
         
         
        const result = await response.json();
        const data = new Uint8Array(result.data);
        const workbook = XLSX.read(data, { type: "array" });

      const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true }) as any;
        const merges: XLSX.Range[] = ws["!merges"] || [];
        return { name, data, merges };
      });

      setSheets(sheetData);
      setActiveSheet(0);
      const headers = sheetData[0].data[0];
      const validEditableColumns = masterEditableColumns.filter((col) => headers.includes(col));
      setEditableColumns(validEditableColumns);

      convertSheetToJson(sheetData[0]);
    } catch (err) {
      console.error("Failed to load XLSX from URL:", err);
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

  const getMergedCell = (r: number, c: number, merges: XLSX.Range[]) => {
    for (const merge of merges) {
      const { s, e } = merge;
      if (r >= s.r && r <= e.r && c >= s.c && c <= e.c) {
        if (r === s.r && c === s.c) return { topLeft: true, rowSpan: e.r - s.r + 1, colSpan: e.c - s.c + 1 };
        return { topLeft: false };
      }
    }
    return null;
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
  if (sheets.length === 0) return;

  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    if (sheet.merges.length > 0) ws["!merges"] = sheet.merges;
    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  });

  XLSX.writeFile(workbook, "Edited_Sheets.xlsx");
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

const handleDownloadStyledXLSXbkp = async () => {
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
    ws.getRow(1).height = 60;

    // 👉 Column widths
    ws.columns = new Array(maxCols).fill({ width: 25 });

    // 👉 Conditional formatting for "branch" column
    const headers = sheet.data[0];
    const branchIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "branch"
    );
    if (branchIndex !== -1) {
      ws.getColumn(branchIndex + 1).eachCell((cell, rowNumber) => {
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

    // 👉 Conditional formatting for "rating" column
    const ratingIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "rating"
    );
    if (ratingIndex !== -1) {
      ws.getColumn(ratingIndex + 1).eachCell((cell, rowNumber) => {
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

const handleDownloadStyledXLSX_TitleTop = async () => {
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
    const branchIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "branch"
    );
    if (branchIndex !== -1) {
      ws.getColumn(branchIndex + 1).eachCell((cell, rowNumber) => {
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
    const genderIndex = headers.findIndex(
      (h) => typeof h === "string" && h.toLowerCase() === "gender"
    );
    if (genderIndex !== -1) {
      ws.getColumn(genderIndex + 1).eachCell((cell, rowNumber) => {
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

    // // 👉 Spacer row before project info
    // ws.addRow([]);

    // // 👉 Spacer row before project info
    // const spacerRow = ws.addRow([]);
    // spacerRow.eachCell((cell, colNumber) => {
    //   cell.fill = {
    //     type: "pattern",
    //     pattern: "solid",
    //     fgColor: { argb: "FF1F77B4" }, // Blue background
    //   };
    // });
    // // ensure spacer row spans all columns with background
    // ws.mergeCells(spacerRow.number, 1, spacerRow.number, maxCols);

    // add spacer row with background
// const spacerRow = ws.addRow(new Array(maxCols).fill("")); // create empty cells

// spacerRow.height = 40;

// spacerRow.eachCell((cell) => {
//   cell.fill = {
//     type: "pattern",
//     pattern: "solid",
//     fgColor: { argb: "FF1F77B4" }, // Blue background
//   };
// });

// // merge across all columns
// ws.mergeCells(spacerRow.number, 1, spacerRow.number, maxCols);

const spacerRow = ws.addRow([]);

// set row height = 40
spacerRow.height = 20;

// apply blue background to all cells in the row
for (let i = 1; i <= maxCols; i++) {
  const cell = spacerRow.getCell(i);
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F77B4" }, // Blue background
  };
}

// ensure spacer row spans all columns (even though it's empty)
ws.mergeCells(spacerRow.number, 1, spacerRow.number, maxCols);

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
      <h1 className="text-2xl font-bold mb-4">XLSX Editor</h1>

      <input type="file" accept=".xlsx" onChange={handleFileUpload} 
        className="mb-4 block w-full text-sm text-gray-900 file:mr-4 file:py-2 file:px-4 file:border file:border-gray-300 file:rounded-lg file:cursor-pointer file:bg-gray-50 file:focus:outline-none file:focus:ring-2 file:focus:ring-blue-500"
       />

      <div className="mb-4 flex gap-2">
        <input
          type="text"
          placeholder="Enter XLSX file URL"
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
                  const headers = sheets[idx].data[0];
                  const validEditableColumns = masterEditableColumns.filter((col) => headers.includes(col));
                  setEditableColumns(validEditableColumns);
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
                  const renderedCells: Set<string> = new Set();

                  return (
                    <tr key={rIdx}>
                      {row.map((cell, cIdx) => {
                        if (renderedCells.has(`${rIdx}-${cIdx}`)) return null;

                        const merged = getMergedCell(rIdx, cIdx, sheets[activeSheet].merges);
                        let rowSpan = 1;
                        let colSpan = 1;

                        if (merged && merged.topLeft) {
                          rowSpan = merged.rowSpan ?? 1;
                          colSpan = merged.colSpan ?? 1;
                          for (let r = rIdx; r < rIdx + rowSpan; r++) {
                            for (let c = cIdx; c < cIdx + colSpan; c++) {
                              renderedCells.add(`${r}-${c}`);
                            }
                          }
                        } else if (merged && !merged.topLeft) {
                          return null;
                        }

                        const headerName = headers[cIdx] as string;
                        const isEditable = !isHeader && editableColumns.includes(headerName);
                        const inEditMode = editingRow === rIdx;

                        return (
                          <td
                            key={cIdx}
                            className={`border p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}
                            rowSpan={rowSpan}
                            colSpan={colSpan}
                          >
                            {isHeader ? (
                              cell
                            ) : inEditMode && isEditable ? (
                              headerName === "Product line" ? (
                                <input
                                  type="text"
                                  className="w-full border-none p-1 focus:outline-none text-center"
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Risk Description" ? (
                                <textarea
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Probability" || headerName === "Impact" ? (
                                <select
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? "Low"}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                >
                                  <option value="High">High</option>
                                  <option value="Low">Low</option>
                                </select>
                              ) : headerName === "Response Plan" ? (
                                <select
                                  className="w-full border-none p-1 focus:outline-none"
                                  value={tempRowData[headerName] ?? "FALSE"}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                >
                                  <option value="TRUE">TRUE</option>
                                  <option value="FALSE">FALSE</option>
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
              className="mt-4 px-4 py-2 bg-green-600 text-white rounded"
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
