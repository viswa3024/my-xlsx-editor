"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { saveAs } from "file-saver";

type SheetData = {
  name: string;
  data: (string | number)[][];
};

export default function Home() {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [activeJsonData, setActiveJsonData] = useState<Record<string, any[]> | null>(null);

  // Handle file upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const workbook = XLSX.read(bstr, { type: "binary" });

      const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number)[][];
        return { name, data };
      });

      console
      setSheets(sheetData);
      setActiveSheet(0);
    };
    reader.readAsBinaryString(file);
  };

  const exportAllSheetsToCSVasZip = async () => {
  if (sheets.length === 0) return;

  const zip = new JSZip();

  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    zip.file(`${sheet.name}.csv`, csv);
  });

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, "all_sheets_csv.zip");
};

  // Handle cell edit
  const handleEdit = (row: number, col: number, value: string) => {
    setSheets((prev) => {
      const newSheets = [...prev];
      newSheets[activeSheet].data[row][col] = value;
      return newSheets;
    });
  };

  // Export current sheet to CSV
  const exportToCSV = () => {
    if (sheets.length === 0) return;

    const ws = XLSX.utils.aoa_to_sheet(sheets[activeSheet].data);
    const csv = XLSX.utils.sheet_to_csv(ws);

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${sheets[activeSheet].name}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportAllSheetsToCSV = () => {
  if (sheets.length === 0) return;

  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    const csv = XLSX.utils.sheet_to_csv(ws);

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${sheet.name}.csv`;
    a.click();

    URL.revokeObjectURL(url);
  });
};


const convertActiveSheetToColumnJSON = () => {
  if (sheets.length === 0) return;

  const [headers, ...rows] = sheets[activeSheet].data;

  const columnData: Record<string, any[]> = {};
  headers.forEach((header) => {
    columnData[header as string] = [];
  });

  rows.forEach((row) => {
    headers.forEach((header, idx) => {
      columnData[header as string].push(row[idx] ?? null);
    });
  });

  setActiveJsonData(columnData);
};

const uploadAllSheetsToAPI = async () => {
  if (sheets.length === 0) return;

  // Prepare CSVs for each sheet
  const payload = sheets.map((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    return { name: sheet.name, csv };
  });

  try {
    const res = await fetch("http://localhost:8000/upload-csvs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheets: payload }),
    });

    if (res.ok) {
      alert("Sheets uploaded successfully!");
    } else {
      alert("Upload failed.");
    }
  } catch (err) {
    console.error(err);
    alert("Error uploading sheets.");
  }
};

const uploadAllSheetsToAPIRouter = async () => {
  if (sheets.length === 0) return;

  // Prepare CSVs for each sheet
  const payload = sheets.map((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    return { name: sheet.name, csv };
  });

  try {
    const res = await fetch("/api/upload-csvs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheets: payload }),
    });

    if (res.ok) {
      alert("Sheets uploaded successfully!");
    } else {
      alert("Upload failed.");
    }
  } catch (err) {
    console.error(err);
    alert("Error uploading sheets.");
  }
};


const downloadAllSheetsAsJSON = () => {
  if (sheets.length === 0) return;

  const jsonData = sheets.map((sheet) => ({
    name: sheet.name,
    data: sheet.data,
  }));

  const blob = new Blob([JSON.stringify(jsonData, null, 2)], {
    type: "application/json",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "all_sheets.json";
  a.click();
  URL.revokeObjectURL(url);
};


const downloadSheetsAsJSONFiles = () => {
  if (sheets.length === 0) return;

  sheets.forEach((sheet) => {
    const jsonContent = JSON.stringify(sheet.data, null, 2); // save only data
    // If you want { name, data } format:
    // const jsonContent = JSON.stringify({ name: sheet.name, data: sheet.data }, null, 2);

    const blob = new Blob([jsonContent], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${sheet.name}.json`;
    a.click();

    URL.revokeObjectURL(url);
  });
};

const downloadSheetsAsJSONFilesKeys = () => {
  if (sheets.length === 0) return;

  sheets.forEach((sheet) => {
    const [headers, ...rows] = sheet.data;

    // Map rows to objects using headers
    const jsonObjects = rows.map((row) => {
      const obj: Record<string, any> = {};
      headers.forEach((header, idx) => {
        obj[header as string] = row[idx] ?? null;
      });
      return obj;
    });

    const jsonContent = JSON.stringify(jsonObjects, null, 2);

    const blob = new Blob([jsonContent], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${sheet.name}.json`;
    a.click();

    URL.revokeObjectURL(url);
  });
};


const downloadSheetsAsColumnJSONFiles = () => {
  if (sheets.length === 0) return;

  sheets.forEach((sheet) => {
    const [headers, ...rows] = sheet.data;

    // Build column-based JSON
    const columnData: Record<string, any[]> = {};
    headers.forEach((header) => {
      columnData[header as string] = [];
    });

    rows.forEach((row) => {
      headers.forEach((header, idx) => {
        columnData[header as string].push(row[idx] ?? null);
      });
    });

    const jsonContent = JSON.stringify(columnData, null, 2);

    const blob = new Blob([jsonContent], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${sheet.name}.json`;
    a.click();

    URL.revokeObjectURL(url);
  });
};
  

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 XLSX Editor (Multiple Sheets)</h1>

      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileUpload}
        className="mb-4"
      />

      {sheets.length > 0 && (
        <>
          {/* Sheet Tabs */}
          <div className="flex space-x-2 mb-4">
            {sheets.map((sheet, idx) => (
              <button
                key={idx}
                onClick={() => setActiveSheet(idx)}
                className={`px-4 py-2 rounded ${
                  idx === activeSheet ? "bg-blue-600 text-white" : "bg-gray-200"
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>

          {/* Editable Table */}
          <div className="overflow-auto border rounded-lg">
            <table className="border-collapse w-full">
              <tbody>
                {sheets[activeSheet].data.map((row, rIdx) => (
                  <tr key={rIdx}>
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} className="border p-1">
                        <input
                          value={cell}
                          onChange={(e) => handleEdit(rIdx, cIdx, e.target.value)}
                          className="w-full border-none outline-none p-1"
                        />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <br />

          {activeJsonData && (
  <div className="mt-6 overflow-auto border rounded-lg">
    <table className="border-collapse w-full text-sm">
      <thead>
        <tr>
          {Object.keys(activeJsonData).map((key) => (
            <th key={key} className="border p-2 bg-gray-200">{key}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {Array.from({ length: Math.max(...Object.values(activeJsonData).map(arr => arr.length)) }).map((_, rowIdx) => (
          <tr key={rowIdx}>
            {Object.keys(activeJsonData).map((key) => (
              <td key={key} className="border p-2">
                {activeJsonData[key][rowIdx] ?? ""}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  </div>
)}

          <br />

          <button
            onClick={exportToCSV}
            className="mt-4 px-4 py-2 bg-green-600 text-white rounded"
          >
            Export Active Sheet to CSV
          </button>
          <br/>
          <button
            onClick={exportAllSheetsToCSVasZip}
            className="mt-4 px-4 py-2 bg-purple-600 text-white rounded"
          >
            Export ALL Sheets to CSV (ZIP)
          </button>
          <br />
          <button
            onClick={exportAllSheetsToCSV}
            className="mt-4 px-4 py-2 bg-red-600 text-white rounded"
          >
            Export ALL Sheets as CSV Files
          </button>
          <br />

          <button
            onClick={uploadAllSheetsToAPI}
            className="mt-4 px-4 py-2 bg-blue-600 text-white rounded"
          >
            Upload ALL Sheets to FastAPI
          </button>
          <br />
          <button
            onClick={uploadAllSheetsToAPIRouter}
            className="mt-4 px-4 py-2 bg-blue-600 text-white rounded"
          >
            Upload ALL Sheets to FastAPI Router
          </button>
          <br />
          <button
            onClick={downloadAllSheetsAsJSON}
            className="mt-4 px-4 py-2 bg-orange-600 text-white rounded"
          >
            Download ALL Sheets as JSON
          </button>
          <br />
          <button
            onClick={downloadSheetsAsJSONFiles}
            className="mt-4 px-4 py-2 bg-yellow-600 text-white rounded"
          >
            Download ALL Sheets as JSON Files
          </button>
          <br />
          <button
            onClick={downloadSheetsAsJSONFilesKeys}
            className="mt-4 px-4 py-2 bg-pink-600 text-white rounded"
          >
            Download ALL Sheets as JSON Files (Keyed)
          </button>
          <br />
          <button
            onClick={downloadSheetsAsColumnJSONFiles}
            className="mt-4 px-4 py-2 bg-indigo-600 text-white rounded"
          >
            Download ALL Sheets as Column JSON
          </button>
          <br />

          <button
            onClick={convertActiveSheetToColumnJSON}
            className="mt-4 px-4 py-2 bg-teal-600 text-white rounded"
          >
            Convert Active Sheet to Column JSON
          </button>
        </>
      )}
    </main>
  );
}
