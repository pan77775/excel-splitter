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
    // 每次選擇新檔案時，重置所有狀態
    setFile(uploadedFile);
    setColumns([]);
    setSelectedCols([]);
    setSplitColumn("");
    setStatus("");

    if (!uploadedFile) return;

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
        if (key === undefined || key === null || key === "") {
          if (!acc["未分類"]) acc["未分類"] = [];
          acc["未分類"].push(row);
        } else {
          const keyStr = String(key);
          if (!acc[keyStr]) acc[keyStr] = [];
          acc[keyStr].push(row);
        }
        return acc;
      }, {});

      const wb = XLSX.utils.book_new();

      Object.entries(groups).forEach(([key, rows]) => {
        if (rows.length > 0) {
          const filtered = rows.map((r) => {
            const filteredRow = {};
            const userSelected = selectedCols.filter(col => col !== splitColumn);
            const finalCols = [...userSelected, splitColumn];
            finalCols.forEach((col) => {
              filteredRow[col] = r[col];
            });
            return filteredRow;
          });
          const ws = XLSX.utils.json_to_sheet(filtered);

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
      <div className="max-w-3xl mx-auto bg-white shadow-xl rounded-xl p-8 space-y-8">
        <header className="text-center">
          <h1 className="text-3xl font-bold text-gray-800">Excel 分頁小幫手</h1>
          <p className="text-sm text-gray-500 mt-1">快速將 Excel 按欄位自動分頁並匯出</p>
        </header>

        <div className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">上傳 Excel 檔案</label>
          <div className="flex items-center">
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileUpload}
              className="block w-full text-sm text-gray-700 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
          </div>
          {file && (
            <div className="flex items-center mt-1">
              {/* 下方檔名顯示已移除 */}
            </div>
          )}
        </div>

          {file && columns.length > 0 && (
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
                className="text-blue-600 hover:underline text-sm mb-2 mr-4"
                disabled={!splitColumn}
              >
                全選
              </button>
              <button
                onClick={() => setSelectedCols(splitColumn ? [splitColumn] : [])}
                className="text-red-500 hover:underline text-sm mb-2"
                disabled={!splitColumn}
              >
                全部清除
              </button>
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-2">
                {columns.map((col) => {
                  const isSplitCol = col === splitColumn;
                  const order = selectedCols.indexOf(col);
                  // 只有非分頁欄位才顯示編號，且從1開始
                  const orderLabel = (!isSplitCol && order >= 0)
                    ? selectedCols.filter(c => c !== splitColumn).indexOf(col) + 1
                    : null;
                  const disabled = !splitColumn || isSplitCol;
                  return (
                    <label 
                      key={col} 
                      className={`flex items-center text-sm text-gray-700 border border-gray-200 rounded p-2 truncate ${disabled ? 'opacity-50' : ''}`}
                      title={col}
                    >
                      {/* 藍色方塊，分頁欄位不顯示數字 */}
                      {orderLabel !== null && (
                        <span
                          className="inline-flex items-center justify-center w-6 h-6 text-xs font-bold text-white bg-blue-500 rounded-md mr-2"
                          style={{ minWidth: "1.5rem", minHeight: "1.5rem" }}
                        >
                          {orderLabel}
                        </span>
                      )}
                      {isSplitCol && (
                        <span
                          className="inline-flex items-center justify-center w-6 h-6 bg-blue-500 rounded-md mr-2"
                          style={{ minWidth: "1.5rem", minHeight: "1.5rem" }}
                        ></span>
                      )}
                      <input
                        type="checkbox"
                        className="hidden"
                        value={col}
                        checked={selectedCols.includes(col)}
                        disabled={disabled}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setSelectedCols((prev) =>
                            checked ? [...prev, col] : prev.filter((c) => c !== col)
                          );
                        }}
                      />
                      <span className="truncate">
                        {col} {isSplitCol && <span className="text-xs text-gray-400">(分頁欄位)</span>}
                      </span>
                    </label>
                  );
                })}
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
