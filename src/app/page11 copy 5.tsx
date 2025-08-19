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
  const [activeJsonData, setActiveJsonData] = useState<Record<string, any[]>>({});

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

  // Convert active sheet to column-based JSON
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

  // Update value for a specific header and row
  const handleJsonChange = (header: string, rowIndex: number, value: any) => {
    const updatedJson = { ...activeJsonData };
    updatedJson[header][rowIndex] = value;
    setActiveJsonData(updatedJson);
  };

  // Convert JSON back to 2D array to update sheet
  const updateSheetFromJson = () => {
    if (sheets.length === 0 || Object.keys(activeJsonData).length === 0) return;

    const headers = Object.keys(activeJsonData);
    const rowCount = activeJsonData[headers[0]].length;

    const newData: (string | number)[][] = [headers];

    for (let i = 0; i < rowCount; i++) {
      const row: (string | number)[] = [];
      headers.forEach((header) => {
        row.push(activeJsonData[header][i]);
      });
      newData.push(row);
    }

    const updatedSheets = [...sheets];
    updatedSheets[activeSheet].data = newData;
    setSheets(updatedSheets);
  };

  return (
    <main className="p-6">
      <h1 className="text-2xl font-bold mb-4">📊 XLSX Editor (JSON-based Risk Title Edit)</h1>

      <input type="file" accept=".xlsx" onChange={handleFileUpload} className="mb-4" />

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

          <button
            className="mb-4 px-4 py-2 bg-green-500 text-white rounded"
            onClick={convertActiveSheetToColumnJSON}
          >
            Convert to JSON
          </button>

          {/* Editable JSON for Risk Title */}
          {Object.keys(activeJsonData).length > 0 && (
            <div className="overflow-auto border rounded-lg p-2">
              <table className="border-collapse w-full">
                <thead>
                  <tr>
                    {Object.keys(activeJsonData).map((header) => (
                      <th key={header} className="border p-1 bg-gray-200 font-bold">
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: activeJsonData[Object.keys(activeJsonData)[0]].length }).map((_, rowIdx) => (
                    <tr key={rowIdx}>
                      {Object.keys(activeJsonData).map((header) => (
                        <td key={header} className="border p-1 text-center">
                          {header === "Risk Title" ? (
                            <input
                              type="text"
                              className="w-full border-none p-1 focus:outline-none text-center"
                              value={activeJsonData[header][rowIdx]}
                              onChange={(e) => handleJsonChange(header, rowIdx, e.target.value)}
                            />
                          ) : (
                            activeJsonData[header][rowIdx]
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>

              <button
                className="mt-4 px-4 py-2 bg-blue-600 text-white rounded"
                onClick={updateSheetFromJson}
              >
                Save Changes to Sheet
              </button>
            </div>
          )}
        </>
      )}
    </main>
  );
}
