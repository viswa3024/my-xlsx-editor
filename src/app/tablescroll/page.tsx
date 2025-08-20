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
      {sheets.length > 0 && (
        <>
         
        

          <div className="overflow-auto shadow-lg p-4 overflow-auto" 
              style={{ maxHeight: "600px", minWidth: "400px", maxWidth: "600px", minHeight: "400px" }}
          
          
          >
            <table className="min-w-full border-collapse table-fixed">
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
                                  rows={2}
                                  value={tempRowData[headerName] ?? ""}
                                  onChange={(e) =>
                                    setTempRowData((prev) => ({ ...prev, [headerName]: e.target.value }))
                                  }
                                />
                              ) : headerName === "Customer type" ? (
                               
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
        </>
      )}
    </main>
  );
}
