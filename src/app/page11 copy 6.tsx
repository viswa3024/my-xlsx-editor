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

    // Reflect changes back to 2D array
    const updatedSheets = [...sheets];
    const data = updatedSheets[activeSheet].data;
    data[rowIndex + 1][data[0].findIndex((h) => h === header)] = value; // +1 because first row = header
    setSheets(updatedSheets);
  };

  // Render merged cells info
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

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 XLSX Editor (Risk Title Editable + Merges)</h1>

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
                  const riskColIndex = headers.findIndex((h) => h === "Risk Title");

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

                        const isEditable = !isHeader && cIdx === riskColIndex;

                        return (
                          <td
                            key={cIdx}
                            className={`border p-1 text-center ${isHeader ? "bg-gray-200 font-bold" : ""}`}
                            rowSpan={rowSpan}
                            colSpan={colSpan}
                          >
                            {isEditable ? (
                              <input
                                type="text"
                                className="w-full border-none p-1 focus:outline-none text-center"
                                value={sheetJson["Risk Title"][rIdx - 1] ?? ""}
                                onChange={(e) => updateJsonValue("Risk Title", rIdx - 1, e.target.value)}
                              />
                            ) : (
                              cell
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
    </main>
  );
}
