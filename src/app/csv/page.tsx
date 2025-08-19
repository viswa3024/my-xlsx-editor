"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

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

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 CSV Editor (Editable Risk Columns with Row Edit)</h1>

      <input type="file" accept=".csv" onChange={handleFileUpload} className="mb-4" />

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
          </div>
        </>
      )}
    </main>
  );
}
