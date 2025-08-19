"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

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

  const editableColumns = ["Risk Title", "Risk Description", "Probability", "Impact", "Response Plan"];

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

      convertSheetToJson(sheetData[0]);
    };
    reader.readAsBinaryString(file);
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

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 XLSX Editor (Editable Risk Columns with Row Edit)</h1>

      <input type="file" accept=".xlsx" onChange={handleFileUpload} className="mb-4" />

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
                              headerName === "Risk Title" ? (
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
          </div>
        </>
      )}
    </main>
  );
}
