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
        // 檢查 splitColumn 的值是否存在且不為空
        const key = row[splitColumn];
        if (key === undefined || key === null || key === "") {
          // 跳過空值或將它們分配到特定分頁名稱
          if (!acc["未分類"]) acc["未分類"] = [];
          acc["未分類"].push(row);
        } else {
          // 對於非空值，正常處理
          const keyStr = String(key); // 確保 key 是字串
          if (!acc[keyStr]) acc[keyStr] = [];
          acc[keyStr].push(row);
        }
        return acc;
      }, {});
  
      const wb = XLSX.utils.book_new();
  
      Object.entries(groups).forEach(([key, rows]) => {
        if (rows.length > 0) { // 只為有資料的群組創建工作表
          const filtered = rows.map((r) => {
            const filteredRow = {};
            selectedCols.forEach((col) => {
              filteredRow[col] = r[col];
            });
            return filteredRow;
          });
          const ws = XLSX.utils.json_to_sheet(filtered);
          
          // 確保工作表名稱有效且不超過 Excel 限制的 31 個字元
          let sheetName = key;
          if (sheetName === "") sheetName = "未命名";
          sheetName = sheetName.substring(0, 31).replace(/[\[\]\*\?\/\\]/g, "_");
          
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }
      });
  
      XLSX.writeFile(wb, "分頁結果.xlsx");
      setStatus("處理完成，請下載檔案！");
    };
    reader.readAsArrayBuffer(file);
  };
  

  const handleSplitColumnChange = (value) => {
    setSplitColumn(value);
    setSelectedCols((prev) => (prev.includes(value) ? prev : [...prev, value]));
  };

  const handleSelectAll = () => {
    const allExceptSplit = columns.filter((col) => col !== splitColumn);
    setSelectedCols([splitColumn, ...allExceptSplit]);
  };

  return (
    <div className="min-h-screen bg-gray-50 py-10 px-4 sm:px-6 lg:px-8">
      <div className="max-w-2xl mx-auto bg-white shadow-xl rounded-xl p-8 space-y-8">
        <header className="text-center">
          <h1 className="text-3xl font-bold text-gray-800">Excel 分頁小幫手</h1>
          <p className="text-sm text-gray-500 mt-1">快速將 Excel 按欄位自動分頁並匯出</p>
        </header>

        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">上傳 Excel 檔案</label>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileUpload}
              className="block w-full text-sm text-gray-700 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
          </div>

          {columns.length > 0 && (
            <>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">分頁依據欄位</label>
                <select
                  className="w-full border border-gray-300 rounded-md p-2 text-sm"
                  value={splitColumn}
                  onChange={(e) => handleSplitColumnChange(e.target.value)}
                >
                  <option value="">--請選擇--</option>
                  {columns.map((col) => (
                    <option key={col} value={col}>{col}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">輸出欄位</label>
                <button
                  onClick={handleSelectAll}
                  className="text-blue-600 hover:underline text-sm mb-2"
                >全選</button>
                <div className="grid grid-cols-2 gap-2">
                  {columns.map((col) => (
                    <label key={col} className="flex items-center text-sm text-gray-700">
                      <input
                        type="checkbox"
                        className="mr-2 rounded"
                        value={col}
                        checked={selectedCols.includes(col)}
                        disabled={col === splitColumn}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setSelectedCols((prev) =>
                            checked ? [...prev, col] : prev.filter((c) => c !== col)
                          );
                        }}
                      />
                      {col} {col === splitColumn && <span className="text-xs text-gray-400">(分頁欄位)</span>}
                    </label>
                  ))}
                </div>
              </div>

              <button
                onClick={handleSplit}
                className="w-full mt-4 bg-blue-600 text-white font-semibold py-2 px-4 rounded-md hover:bg-blue-700"
              >
                分頁並下載
              </button>
            </>
          )}

          {status && <p className="text-green-600 text-sm font-medium text-center">{status}</p>}
        </div>
      </div>
    </div>
  );
}
