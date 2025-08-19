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

  // Expand merged cells
  const expandMergedCells = (ws: XLSX.WorkSheet) => {
    const data: (string | number)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true }) as any;
    const merges = ws["!merges"] || [];

    merges.forEach((merge) => {
      const { s, e } = merge; // start and end of merge
      const value = data[s.r][s.c];
      for (let r = s.r; r <= e.r; r++) {
        for (let c = s.c; c <= e.c; c++) {
          data[r][c] = value;
        }
      }
    });

    return data;
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const workbook = XLSX.read(bstr, { type: "binary" });

      const sheetData: SheetData[] = workbook.SheetNames.map((name) => {
        const ws = workbook.Sheets[name];
        const data = expandMergedCells(ws);
        return { name, data };
      });

      setSheets(sheetData);
      setActiveSheet(0);
    };
    reader.readAsBinaryString(file);
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

          <div className="overflow-auto border rounded-lg">
            <table className="border-collapse w-full">
              <tbody>
                {sheets[activeSheet].data.map((row, rIdx) => (
                  <tr key={rIdx}>
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} className="border p-1">
                        {typeof cell === "boolean" ? (cell ? "true" : "false") : cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </main>
  );
}
