"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { CreditCard, Save, Edit, X } from "lucide-react";  
import CustomSelect from "./CustomSelect";

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

  const [editableColumns, setEditableColumns] = useState<string[]>(["Product line", "Unit price", "Customer type", "City", "Gender", "Quantity"])

  const masterEditableColumns = ["Product line", "Unit price", "Customer type", "City", "Gender", "Quantity"];

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const workbook = XLSX.read(bstr, { type: "binary" });

      const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, raw: false }) as any;
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
        //raw: false in sheet_to_json, XLSX will automatically try to parse dates into human-readable form.
        const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, raw: false }) as any;
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

const formatHeader = (header: string) => {
  // Replace underscores with space
  let str = header.replace(/_/g, " ");
  // Capitalize first letter of each word
  str = str.replace(/\w\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
  return str;
};

const getBranchColor = (value: string) => {
  switch (value) {
    case "A":
      return "bg-red-500 text-white";
    case "B":
      return "bg-orange-400 text-white";
    case "C":
      return "bg-green-500 text-white";
    case "D":
      return "bg-red-800 text-white";
    default:
      return "";
  }
};

const headerWidths: Record<string, string> = {
  "Invoice ID": "150px",
  "Branch": "120px",
  "Customer type": "100px",
  "Gender": "130px",
  "Product line": "200px",
  "Unit price": "120px",
  "Quantity": "120px",
  "Tax 5%": "120px",
  "Total": "120px",
  "Date": "120px",
  "Time": "120px",
  "Payment": "120px",
  "cogs": "120px",
  "gross income": "120px",
  "Rating": "120px",
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

          <div className="overflow-auto shadow-lg p-4 bg-white">
            <table className="min-w-full border-collapse">
              {/* <thead>
                <tr>
                  <th colSpan={sheets[activeSheet].data[0]?.length || 1} className="p-2">
                    <div className="flex space-x-2 mb-2">
                      {sheets.map((sheet, idx) => (
                        <button
                          key={idx}
                          onClick={() => {
                            setActiveSheet(idx);
                            convertSheetToJson(sheets[idx]);
                            const headers = sheets[idx].data[0];
                            const validEditableColumns = masterEditableColumns.filter((col) =>
                              headers.includes(col)
                            );
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
                  </th>
                </tr>
              </thead> */}

              {/* <thead>
    <tr>
      {sheets.map((sheet, idx) => (
        <th
          key={idx}
          onClick={() => {
            setActiveSheet(idx);
            convertSheetToJson(sheets[idx]);
            const headers = sheets[idx].data[0];
            const validEditableColumns = masterEditableColumns.filter((col) =>
              headers.includes(col)
            );
            setEditableColumns(validEditableColumns);
          }}
          className={`px-4 py-2 cursor-pointer rounded-t ${
            idx === activeSheet
              ? "bg-blue-600 text-white"
              : "bg-gray-200 text-gray-700 hover:bg-gray-300"
          }`}
        >
          {sheet.name}
        </th>
      ))}
    </tr>
  </thead> */}

  <thead className="mb-4">
  <tr className="">
    {sheets.map((sheet, idx) => (
      <th
        key={idx}
        // colSpan={sheets[activeSheet]?.data[0]?.length || 1} // span full tbody width
        colSpan={
          Math.floor(
            (sheets[activeSheet]?.data[0]?.length + 2 || 1) / sheets.length
          )
        }
        className={`px-4 py-2 mx-1 cursor-pointer text-sm font-medium border-b-[4px] ${
          idx === activeSheet ? "border-blue-600 text-blue-600" : "border-black-300 text-gray-600"
        }`}
        onClick={() => {
          setActiveSheet(idx);
          convertSheetToJson(sheets[idx]);
          const headers = sheets[idx].data[0];
          const validEditableColumns = masterEditableColumns.filter((col) =>
            headers.includes(col)
          );
          setEditableColumns(validEditableColumns);
        }}
      >
        {sheet.name}
      </th>
    ))}
  </tr>
</thead>
              <tbody>
                {sheets[activeSheet].data.map((row, rIdx) => {
                  const isHeader = rIdx === 0;
                  const headers = sheets[activeSheet].data[0];
                  const renderedCells: Set<string> = new Set();

                  return (
                    <tr key={rIdx} 
                      //className={`${!isHeader && rIdx % 2 === 0 ? "bg-white" : "bg-gray-50"} hover:bg-gray-100 transition-colors`}
                    >
                      {row.map((cell, cIdx) => {

                        const headerName = headers[cIdx] as string;
                        if (headerName === "Payment") return null; // Skip Payment for now
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

                        const isEditable = !isHeader && editableColumns.includes(headerName);
                        const inEditMode = editingRow === rIdx;

                        const displayHeader = isHeader ? formatHeader(headerName) : cell;

                        return (
                          <td
                            key={cIdx}
                            //className={`border p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : "border-b hover:bg-gray-50"}`}

                            className={`border border-[#ddd] p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}
                            rowSpan={rowSpan}
                            colSpan={colSpan}
                            style={{
                              width: headerWidths[cell as string] || "auto", // fallback if header not in map
                            }}
                          >
                            {isHeader ? (
                              displayHeader
                            ) : inEditMode && isEditable ? (
                              headerName === "Unit price" ? (
                                <input
                                  type="text"
                                  className="w-full border border-gray-300 rounded-md p-1 text-center shadow-sm placeholder-gray-400 focus:outline-none focus:ring-0"
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Product line" || headerName === "City" ? (
                                <textarea
                                //resize-none focus:border-none 
                                  className="
                                    w-full 
                                    border 
                                    border-gray-300 
                                    rounded-md 
                                    p-2 
                                    text-sm 
                                    text-gray-800 
                                    bg-white 
                                    focus:outline-none 
                                    focus:ring-0
                                  "
                                  rows={2}
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Customer type" ? (
                                // <select
                                //   className="w-full p-2 border border-gray-300 rounded-md bg-white text-gray-700 shadow-sm 
                                //     focus:outline-none focus:ring-0  transition cursor-pointer"
                                //   value={tempRowData[headerName] ?? "Member"}
                                //   onChange={(e) =>
                                //     setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                //   }
                                // >
                                //   <option value="Member">Member</option>
                                //   <option value="Normal">Normal</option>
                                // </select>

                                // <CustomSelect
                                //     value={tempRowData[headerName] ?? "Member"}
                                //     options={["Member", "Normal"]}
                                //     onChange={(val) =>
                                //       setTempRowData((prev) => ({ ...prev, [headerName]: val }))
                                //     }
                                //   />
                                <CustomSelect
                                    value={tempRowData[headerName] ?? "Member"}
                                    options={[
                                      { key: "Member", label: "Member" },
                                      { key: "Normal", label: "Normal" },
                                    ]}
                                    onChange={(val) => setTempRowData((prev) => ({ ...prev, [headerName]: val }))}
                                  />
                              ) : headerName === "Quantity" ? (
                                  <input
                                    type="number"
                                    className="w-full border border-gray-300 rounded-md p-1 text-center shadow-sm placeholder-gray-400 focus:outline-none focus:ring-0"
                                    value={tempRowData[headerName] ?? ""}
                                    onChange={(e) =>
                                      setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                    }
                                  />
                              ) : headerName === "Gender" ? (
                                <select
                                   className="w-full p-2 border border-gray-300 rounded-md bg-white text-gray-700 shadow-sm 
               focus:outline-none focus:ring-0  transition cursor-pointer"
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
                            ) :
                            // ✅ Non-editable rows wrapped
                            (
                              headerName.toLowerCase() === "gender" ? (
                              <div className="flex justify-center items-center h-full w-full">
                                <div
                                  className={`flex justify-center items-center 
                                              text-[14px] tracking-[0.1px] w-[90px] pt-[4px] pb-[4px] 
                                              font-bold rounded-[6px] ${
                                                cell === "Male"
                                                  ? "text-green-600 bg-green-100"
                                                  : cell === "Female"
                                                  ? "text-red-600 bg-red-100"
                                                  : "text-black bg-gray-100"
                                              }`}
                                >
                                  {cell}
                                </div>
                              </div>
                            ) :
                              headerName.toLowerCase() === "branch" ? (
                                  <div className="flex justify-center items-center h-full w-full">
                                    <div className={`flex justify-center items-center text-[14px] 
                                                    tracking-[0.1px] w-[40px] pt-[4px] pb-[4px] font-bold rounded-[6px] 
                                                    ${getBranchColor(cell as string)} `}>
                                      {cell}
                                    </div>
                                  </div>
                              ) : headerName.toLowerCase() === "product line" ? (
                                // ✅ Non-editable Product Line with ellipsis & hover tooltip
                                <div
                                  className="font-bold text-[14px] line-clamp-3 overflow-hidden text-ellipsis cursor-pointer"
                                  title={String(cell)} // shows full text on hover
                                >
                                  {cell}
                                </div>
                              ) : (
                              // ✅ Non-editable rows wrapped in div with bold and font-size 14px
                              <div className="font-bold text-[14px]">{cell}</div>
                            )
                          )}
                          </td>
                        );
                      })}

                      {/* Render Payment column at the end */}
                      {(() => {
                        const paymentIndex = headers.findIndex((h) => h === "Payment");
                        if (paymentIndex === -1) return null;
                        const cell = row[paymentIndex];
                        const merged = getMergedCell(rIdx, paymentIndex, sheets[activeSheet].merges);
                        let rowSpan = 1;
                        let colSpan = 1;
                        if (merged && merged.topLeft) {
                          rowSpan = merged.rowSpan ?? 1;
                          colSpan = merged.colSpan ?? 1;
                        }
                        return (
                          <td
                            key="payment"
                            className={`border border-[#ddd] p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}
                            rowSpan={rowSpan}
                            colSpan={colSpan}
                          >
                            {/* {cell} */}
                            {isHeader ? (
                                  <div className="inline-flex items-center justify-center w-10 h-10 bg-blue-200 rounded-full mx-auto">
                                    <CreditCard className="text-blue-600 w-5 h-5" />
                                  </div>
                                ) : (
                                  cell
                                )}
                          </td>
                        );
                      })()}

                      {/* Last column: Edit Changes */}
                      <td className={`border border-[#ddd] p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}>
                        {isHeader
                          ? "Edit Changes"
                          : editingRow === rIdx ? (
                              <>
                              <div className="flex gap-[8px] justify-center items-center">
                                <button
                                  className="py-1 cursor-pointer"
                                  onClick={() => saveRow(rIdx)}
                                >
                                  {/* <Save size={16} /> */}
                                  <Save size={20} className="text-green-500 hover:text-green-700" />
                                </button>
                                <button
                                  className="py-1 cursor-pointer"
                                  onClick={closeEdit}
                                >
                                  {/* <X size={16} /> */} {/* Close icon */}
                                  <X size={20} className="text-red-500 hover:text-red-700" />
                                </button>
                              </div>
                                
                              </>
                            ) : (
                              <button
                                className="px-2 py-1 cursor-pointer"
                                onClick={() => startEditRow(rIdx)}
                              >
                                {/* <Edit size={16} />  */}
                                <Edit size={20} className="text-blue-600 hover:text-blue-800" />
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
