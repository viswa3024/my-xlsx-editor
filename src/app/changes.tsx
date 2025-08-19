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
