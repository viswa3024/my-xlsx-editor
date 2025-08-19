const handleDownload = () => {
  if (sheets.length === 0) return;

  const workbook = XLSX.utils.book_new();
  sheets.forEach((sheet) => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);
    if (sheet.merges.length > 0) ws["!merges"] = sheet.merges;
    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  });

  XLSX.writeFile(workbook, "Edited_Sheets.xlsx");
};


<button
  className="mt-4 px-4 py-2 bg-green-600 text-white rounded"
  onClick={handleDownload}
>
  Download Edited XLSX
</button>



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
