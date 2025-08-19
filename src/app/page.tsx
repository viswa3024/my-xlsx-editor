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
  const [fileUrl, setFileUrl] = useState<string>("");
  const editableKeys = ["Customer type", "Rating", "Gender", "Product line", "Unit price"];
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [rowEdits, setRowEdits] = useState<Record<string, any>>({});

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

  const handleLoadUrl = async () => {
  if (!fileUrl) return;

  try {
    const response = await fetch("/api/load-xlsx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: fileUrl }),
    });
    // if (!response.ok) throw new Error("Failed to fetch file");

    // const arrayBuffer = await response.arrayBuffer();
    // const data = new Uint8Array(arrayBuffer);
    const result = await response.json();
    const data = new Uint8Array(result.data);
    const workbook = XLSX.read(data, { type: "array" });
        

    const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
      const ws = workbook.Sheets[name];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as (string | number)[][];
      return { name, data };
    });

    setSheets(sheetData);
    setActiveSheet(0);
  } catch (err) {
    console.error("Error loading XLSX from URL:", err);
    alert("Failed to load XLSX file. Check URL.");
  }
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
  const handleEditBkp = (row: number, col: number, value: string) => {
    setSheets((prev) => {
      const newSheets = [...prev];
      newSheets[activeSheet].data[row][col] = value;
      return newSheets;
    });
  };

  // rowIndex: number, key: string, value: string | number
const handleEdit = (rowIndex: number, key: string, value: string | number) => {
  if (!activeJsonData) return;

  // Clone the specific column array
  const updatedData = { ...activeJsonData };
  updatedData[key] = [...updatedData[key]];

  // Update the cell
  updatedData[key][rowIndex] = value;

  // Update state
  setActiveJsonData(updatedData);
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
                onClick={() => setActiveSheet(idx)}
                className={`px-4 py-2 rounded ${
                  idx === activeSheet ? "bg-blue-600 text-white" : "bg-gray-200"
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>

    
          <br />

         

{activeJsonData && Object.keys(activeJsonData).length > 0 && (
  <div className="mt-6 overflow-auto border rounded-lg">
    <table className="border-collapse w-full text-sm">
      <thead>
        <tr>
          {editableKeys.map((key) => (
            <th key={key} className="border p-2 bg-gray-100 text-left">
              {key}
            </th>
          ))}
          <th className="border p-2 bg-gray-100 text-left">Actions</th>
        </tr>
      </thead>
      <tbody>
        {Array.from({ length: activeJsonData[Object.keys(activeJsonData)[0]].length }).map(
          (_, rowIndex) => {
            const isEditing = editingRow === rowIndex;
            return (
              <tr key={rowIndex}>
                {editableKeys.map((key) => (
                  <td key={key} className="border p-2">
                    {isEditing ? (
                      key === "Gender" ? (
                        <select
                          className="border p-1 rounded"
                          value={rowEdits[key]}
                          onChange={(e) =>
                            setRowEdits((prev) => ({ ...prev, [key]: e.target.value }))
                          }
                        >
                          <option value="Male">Male</option>
                          <option value="Female">Female</option>
                        </select>
                      ) : key === "Rating" ? (
                        <input
                          type="number"
                          className="border p-1 rounded w-full"
                          value={rowEdits[key]}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (/^\d*$/.test(val) && Number(val) <= 100) {
                              setRowEdits((prev) => ({ ...prev, [key]: Number(val) }));
                            }
                          }}
                        />
                      ) : key === "Unit price" ? (
                        <input
                          type="number"
                          className="border p-1 rounded w-full"
                          value={rowEdits[key]}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (/^\d*\.?\d*$/.test(val)) {
                              setRowEdits((prev) => ({ ...prev, [key]: Number(val) }));
                            }
                          }}
                        />
                      ) : key === "Product line" ? (
                        <textarea
                          className="border p-1 rounded w-full"
                          value={rowEdits[key]}
                          onChange={(e) =>
                            setRowEdits((prev) => ({ ...prev, [key]: e.target.value }))
                          }
                        />
                      ) : (
                        <input
                          type="text"
                          className="border p-1 rounded w-full"
                          value={rowEdits[key]}
                          onChange={(e) =>
                            setRowEdits((prev) => ({ ...prev, [key]: e.target.value }))
                          }
                        />
                      )
                    ) : (
                      activeJsonData[key][rowIndex]
                    )}
                  </td>
                ))}

                <td className="border p-2">
                  {isEditing ? (
                    <div className="flex gap-2">
                      <button
                        className="px-2 py-1 bg-green-500 text-white rounded hover:bg-green-600"
                        onClick={() => {
                          if (!activeJsonData) return;

                          // Clone activeJsonData immutably
                          const updatedData = { ...activeJsonData };
                          editableKeys.forEach((key) => {
                            updatedData[key] = [...updatedData[key]];
                            updatedData[key][rowIndex] = rowEdits[key];
                          });

                          setActiveJsonData(updatedData);
                          setEditingRow(null);
                        }}
                      >
                        Save
                      </button>
                      <button
                        className="px-2 py-1 bg-red-500 text-white rounded hover:bg-red-600"
                        onClick={() => setEditingRow(null)}
                      >
                        Cancel
                      </button>
                    </div>
                  ) : (
                    <button
                      className="px-2 py-1 bg-blue-500 text-white rounded hover:bg-blue-600"
                      onClick={() => {
                        setEditingRow(rowIndex);
                        const initialEdits: Record<string, any> = {};
                        editableKeys.forEach((key) => {
                          initialEdits[key] = activeJsonData[key][rowIndex];
                        });
                        setRowEdits(initialEdits);
                      }}
                    >
                      Edit
                    </button>
                  )}
                </td>
              </tr>
            );
          }
        )}
      </tbody>
    </table>

    <div className="mt-4 flex gap-4">
      <button
        onClick={() => {
          const filteredData: Record<string, any[]> = {};
          editableKeys.forEach((key) => {
            filteredData[key] = activeJsonData[key];
          });

          const blob = new Blob([JSON.stringify(filteredData, null, 2)], {
            type: "application/json",
          });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = "data.json";
          a.click();
        }}
        className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
      >
        Save as JSON
      </button>

      <button
        onClick={() => {
          const filteredData: Record<string, any[]> = {};
          editableKeys.forEach((key) => {
            filteredData[key] = activeJsonData[key];
          });

          const keys = Object.keys(filteredData);
          const rows = Array.from({ length: filteredData[keys[0]].length }).map(
            (_, rowIndex) =>
              keys.map((key) => filteredData[key][rowIndex]).join(",")
          );
          const csvContent = [keys.join(","), ...rows].join("\n");

          const blob = new Blob([csvContent], { type: "text/csv" });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = "data.csv";
          a.click();
        }}
        className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
      >
        Save as CSV
      </button>
    </div>
  </div>
)}




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
