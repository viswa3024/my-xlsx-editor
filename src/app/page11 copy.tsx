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
    };
    reader.readAsBinaryString(file);
  };

  // Check if a cell is part of a merge (not top-left)
  const isMergedCell = (r: number, c: number, merges: XLSX.Range[]) => {
    for (const merge of merges) {
      const { s, e } = merge;
      if (r >= s.r && r <= e.r && c >= s.c && c <= e.c) {
        if (r === s.r && c === s.c) return { topLeft: true, merge };
        return { topLeft: false };
      }
    }
    return null;
  };

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 XLSX Editor (Merged Cells with colSpan/rowSpan)</h1>

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
                    {row.map((cell, cIdx) => {
                      const merged = isMergedCell(rIdx, cIdx, sheets[activeSheet].merges);
                      if (merged && !merged.topLeft) return null; // skip merged cells
                      const rowSpan = merged?.merge ? merged.merge.e.r - merged.merge.s.r + 1 : 1;
                      const colSpan = merged?.merge ? merged.merge.e.c - merged.merge.s.c + 1 : 1;

                      return (
                        <td
                          key={cIdx}
                          className="border p-1 text-center"
                          rowSpan={rowSpan}
                          colSpan={colSpan}
                        >
                          {typeof cell === "boolean" ? (cell ? "true" : "false") : cell}
                        </td>
                      );
                    })}
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
