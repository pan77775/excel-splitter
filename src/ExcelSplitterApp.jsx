import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ExcelSplitterApp() {
  const [file, setFile] = useState(null);
  const [splitColumn, setSplitColumn] = useState("");
  const [columns, setColumns] = useState([]);
  const [selectedCols, setSelectedCols] = useState([]);
  const [status, setStatus] = useState("");

  const handleFileUpload = async (e) => {
    const uploadedFile = e.target.files[0];
    setFile(uploadedFile);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);

      if (json.length > 0) {
        const keys = Object.keys(json[0]);
        setColumns(keys);
      }
    };
    reader.readAsArrayBuffer(uploadedFile);
  };

  const handleSplit = async () => {
    if (!file || !splitColumn || selectedCols.length === 0) {
      setStatus("請選擇檔案、分頁欄位及至少一個輸出欄位");
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);

      const groups = json.reduce((acc, row) => {
        const key = row[splitColumn];
        if (!acc[key]) acc[key] = [];
        acc[key].push(row);
        return acc;
      }, {});

      const newWorkbook = {
        SheetNames: [],
        Sheets: {}
      };

      Object.entries(groups).forEach(([key, rows]) => {
        const filtered = rows.map((r) => {
          const filteredRow = {};
          selectedCols.forEach((col) => {
            filteredRow[col] = r[col];
          });
          return filteredRow;
        });
        const ws = XLSX.utils.json_to_sheet(filtered);
        const sheetName = key.substring(0, 31);
        newWorkbook.SheetNames.push(sheetName);
        newWorkbook.Sheets[sheetName] = ws;
      });

      XLSX.writeFile(newWorkbook, "分頁結果.xlsx");
      setStatus("處理完成，檔案已下載！");
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSplitColumnChange = (value) => {
    setSplitColumn(value);
    // 自動勾選該欄位為輸出欄位之一，且鎖定
    setSelectedCols((prev) =>
      prev.includes(value) ? prev : [...prev, value]
    );
  };

  const handleSelectAll = () => {
    const allExceptSplit = columns.filter((col) => col !== splitColumn);
    setSelectedCols([splitColumn, ...allExceptSplit]);
  };

  return (
    <div className="p-6 max-w-xl mx-auto space-y-6 text-gray-800">
      <header className="text-center">
        <h1 className="text-2xl font-bold">Excel 分頁小幫手</h1>
        <p className="text-sm text-gray-500">快速將 Excel 按欄位自動分頁</p>
      </header>
      <div className="shadow-md border rounded p-4 space-y-4">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="w-full border p-2 rounded"
        />

        {columns.length > 0 && (
          <>
            <div>
              <label className="block text-sm font-medium mb-1">分頁依據欄位</label>
              <select
                className="border p-2 w-full rounded"
                value={splitColumn}
                onChange={(e) => handleSplitColumnChange(e.target.value)}
              >
                <option value="">--請選擇--</option>
                {columns.map((col) => (
                  <option key={col} value={col}>
                    {col}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium mb-1">輸出欄位</label>
              <button
                onClick={handleSelectAll}
                className="mb-2 text-blue-600 hover:underline text-sm"
              >
                全選
              </button>
              <div className="grid grid-cols-2 gap-2">
                {columns.map((col) => (
                  <label key={col} className="flex items-center text-sm">
                    <input
                      type="checkbox"
                      className="mr-2"
                      value={col}
                      checked={selectedCols.includes(col)}
                      disabled={col === splitColumn} // 鎖定分頁欄位
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setSelectedCols((prev) =>
                          checked
                            ? [...prev, col]
                            : prev.filter((c) => c !== col)
                        );
                      }}
                    />
                    {col} {col === splitColumn && "(分頁欄位)"}
                  </label>
                ))}
              </div>
            </div>

            <button
              onClick={handleSplit}
              className="mt-2 w-full bg-blue-600 text-white py-2 px-4 rounded hover:bg-blue-700"
            >
              分頁並下載
            </button>
          </>
        )}

        {status && <p className="text-sm text-green-600 font-medium">{status}</p>}
      </div>
    </div>
  );
}
