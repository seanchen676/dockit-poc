"use client";

import { useState, useCallback } from "react";
import { parseExcel } from "./lib/excel-parser";
import { encode } from "@toon-format/toon";

export default function Home() {
  const [isParsing, setIsParsing] = useState(false);
  const [jsonResult, setJsonResult] = useState<string | null>(null);
  const [toonResult, setToonResult] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>("");

  const handleFileChange = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      setFileName(file.name);
      setErrorMsg(null);
      setJsonResult(null);
      setToonResult(null);

      // Basic Validation
      if (!file.name.endsWith(".xlsx")) {
        setErrorMsg("請上傳正確的 Excel (.xlsx) 檔案。");
        return;
      }

      setIsParsing(true);

      try {
        const buffer = await file.arrayBuffer();

        // Use setTimeout to allow UI to update (show Loading) before heavy sync operation starts
        setTimeout(async () => {
          try {
            const result = await parseExcel(buffer, file.name);
            setJsonResult(JSON.stringify(result, null, 2));
            setToonResult(encode(result));
          } catch (err: unknown) {
            console.error(err);
            const message =
              err instanceof Error
                ? err.message
                : "未知錯誤，可能是檔案損毀或加密。";
            setErrorMsg("解析失敗: " + message);
          } finally {
            setIsParsing(false);
          }
        }, 100);
      } catch (err: unknown) {
        console.error(err);
        setErrorMsg("讀取檔案失敗。");
        setIsParsing(false);
      }
    },
    [],
  );

  const handleCopy = () => {
    if (jsonResult) {
      navigator.clipboard.writeText(jsonResult);
      alert("已複製到剪貼簿！");
    }
  };

  const handleDownload = () => {
    if (!jsonResult) return;
    const blob = new Blob([jsonResult], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${fileName.replace(".xlsx", "")}_parsed.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleCopyToon = () => {
    if (toonResult) {
      navigator.clipboard.writeText(toonResult);
      alert("已複製 Toon 到剪貼簿！");
    }
  };

  const handleDownloadToon = () => {
    if (!toonResult) return;
    const blob = new Blob([toonResult], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${fileName.replace(".xlsx", "")}_parsed.toon`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="flex min-h-screen flex-col items-center p-8 bg-gray-50 dark:bg-gray-900">
      <div className="w-full max-w-4xl bg-white dark:bg-gray-800 rounded-xl shadow-lg p-6 space-y-6">
        {/* Header */}
        <div className="border-b pb-4 border-gray-200 dark:border-gray-700">
          <h1 className="text-2xl font-bold text-gray-900 dark:text-gray-100">
            Excel to LLM-Ready JSON Parser
          </h1>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
            純前端解析工具，將 XLSX 轉換為結構化 JSON
          </p>
        </div>

        {/* Upload Area */}
        <div className="space-y-2">
          <label className="block text-sm font-medium text-gray-700 dark:text-gray-300">
            選擇 Excel 檔案 (.xlsx)
          </label>
          <input
            type="file"
            accept=".xlsx"
            onChange={handleFileChange}
            disabled={isParsing}
            className="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-gray-50 dark:text-gray-400 focus:outline-none dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 p-2"
          />
          {isParsing && (
            <div className="flex items-center text-blue-600 mt-2">
              <svg
                className="animate-spin -ml-1 mr-3 h-5 w-5 text-blue-600"
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
              >
                <circle
                  className="opacity-25"
                  cx="12"
                  cy="12"
                  r="10"
                  stroke="currentColor"
                  strokeWidth="4"
                ></circle>
                <path
                  className="opacity-75"
                  fill="currentColor"
                  d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                ></path>
              </svg>
              <span>正在解析中，請稍候... (大型檔案可能需要幾秒鐘)</span>
            </div>
          )}
          {errorMsg && (
            <div className="text-red-600 bg-red-50 p-3 rounded-md text-sm mt-2 border border-red-200">
              ❌ {errorMsg}
            </div>
          )}
        </div>

        {/* Result Area */}
        {jsonResult && !isParsing && (
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <h2 className="text-lg font-semibold text-gray-800 dark:text-gray-200">
                解析結果 Preview
              </h2>
              <div className="space-x-2">
                <button
                  onClick={handleCopy}
                  className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-md text-sm font-medium transition-colors"
                >
                  複製 JSON
                </button>
                <button
                  onClick={handleDownload}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-md text-sm font-medium transition-colors"
                >
                  下載 .json
                </button>
              </div>
            </div>

            <div className="relative">
              <textarea
                readOnly
                value={jsonResult}
                className="w-full h-96 p-4 font-mono text-sm bg-gray-900 text-green-400 rounded-lg overflow-auto focus:outline-none border border-gray-700 resize-y"
              />
              <div className="absolute top-2 right-4 text-xs text-gray-500">
                Size: {(jsonResult.length / 1024).toFixed(2)} KB
              </div>
            </div>
          </div>
        )}

        {/* Toon Result Area */}
        {toonResult && !isParsing && (
          <div className="space-y-4 pt-6 border-t border-gray-200 dark:border-gray-700">
            <div className="flex justify-between items-center">
              <div>
                <h2 className="text-lg font-semibold text-gray-800 dark:text-gray-200">
                  Toon Result
                </h2>
                <p className="text-xs text-gray-500 dark:text-gray-400">
                  Optimized for LLMs
                </p>
              </div>
              <div className="space-x-2">
                <button
                  onClick={handleCopyToon}
                  className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-md text-sm font-medium transition-colors"
                >
                  複製 Toon
                </button>
                <button
                  onClick={handleDownloadToon}
                  className="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-md text-sm font-medium transition-colors"
                >
                  下載 .toon
                </button>
              </div>
            </div>

            <div className="relative">
              <textarea
                readOnly
                value={toonResult}
                className="w-full h-96 p-4 font-mono text-sm bg-gray-900 text-purple-300 rounded-lg overflow-auto focus:outline-none border border-gray-700 resize-y"
              />
              <div className="absolute top-2 right-4 text-xs text-gray-500">
                Size: {(toonResult.length / 1024).toFixed(2)} KB
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
