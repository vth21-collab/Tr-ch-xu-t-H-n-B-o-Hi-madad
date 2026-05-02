/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */
import { useState, useRef, ChangeEvent } from 'react';
import { Upload, FileText, AlertCircle, Loader2, CheckCircle, X, Trash2, Download } from 'lucide-react';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

interface FileTask {
  id: string;
  file: File;
  status: 'idle' | 'processing' | 'success' | 'error';
  result?: string;
  error?: string;
}

export default function App() {
  const [tasks, setTasks] = useState<FileTask[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addFiles = (filesList: FileList | null) => {
    if (!filesList) return;
    const newTasks: FileTask[] = [];
    Array.from(filesList).forEach((file) => {
      if (file.name.endsWith('.docx')) {
        newTasks.push({
          id: Math.random().toString(36).substring(7) + Date.now(),
          file,
          status: 'idle',
        });
      }
    });

    if (newTasks.length > 0) {
      setTasks((prev) => [...prev, ...newTasks]);
    } else if (filesList.length > 0) {
      alert('Vui lòng chỉ chọn các file .docx (Word).');
    }
    
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    addFiles(e.target.files);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    addFiles(e.dataTransfer.files);
  };

  const removeTask = (id: string) => {
    setTasks((prev) => prev.filter((t) => t.id !== id));
  };

  const processAllFiles = async () => {
    const pendingTasks = tasks.filter((t) => t.status === 'idle' || t.status === 'error');
    if (pendingTasks.length === 0) return;

    setIsProcessing(true);

    for (const task of tasks) {
      if (task.status === 'success') continue;

      setTasks((prev) =>
        prev.map((t) => (t.id === task.id ? { ...t, status: 'processing', error: undefined } : t))
      );

      try {
        const arrayBuffer = await task.file.arrayBuffer();
        const extractedInfo = await mammoth.extractRawText({ arrayBuffer });
        const text = extractedInfo.value;

        if (!text || text.trim().length === 0) {
          throw new Error('Không thể đọc nội dung văn bản từ file này hoặc file rỗng.');
        }

        const response = await ai.models.generateContent({
          model: 'gemini-3.1-pro-preview',
          contents: `Trích xuất thông tin "thời gian bảo hiểm cháy nổ hết hạn" (ngày hết hạn hoặc thời hạn bảo hiểm cháy nổ) từ tài liệu dưới đây.\nNếu tìm thấy, hãy xuất ra thời gian bảo hiểm cháy nổ hết hạn một cách ngắn gọn, rõ ràng (ví dụ: "Ngày hết hạn bảo hiểm cháy nổ: 31/12/2024").\nNếu không tìm thấy, hãy phản hồi: "Không tìm thấy thông tin về thời gian bảo hiểm cháy nổ hết hạn trong tài liệu này".\n\nTài liệu:\n${text}`,
        });

        setTasks((prev) =>
          prev.map((t) =>
            t.id === task.id
              ? { ...t, status: 'success', result: response.text || 'Không có phản hồi từ AI.' }
              : t
          )
        );
      } catch (err: any) {
        console.error(err);
        setTasks((prev) =>
          prev.map((t) =>
            t.id === task.id
              ? { ...t, status: 'error', error: err.message || 'Đã xảy ra lỗi trong quá trình xử lý file.' }
              : t
          )
        );
      }
    }

    setIsProcessing(false);
  };

  const exportToExcel = () => {
    const completedTasks = tasks.filter(t => t.status === 'success');
    if (completedTasks.length === 0) {
      alert("Chưa có kết quả phân tích thành công để xuất ra Excel.");
      return;
    }

    const data = completedTasks.map(task => ({
      'Tên file': task.file.name,
      'Kết quả': task.result || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    
    // Automatically adjust column widths
    const maxWidths = data.reduce((acc, row) => {
      acc[0] = Math.max(acc[0], row['Tên file'].length);
      acc[1] = Math.max(acc[1], row['Kết quả'].length);
      return acc;
    }, [10, 10]);
    
    worksheet['!cols'] = [
      { wch: Math.min(maxWidths[0] + 5, 50) }, // Tên file col width
      { wch: Math.min(maxWidths[1] + 5, 100) } // Kết quả col width
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Kết Quả");
    
    XLSX.writeFile(workbook, "Ket_Qua_Han_Bao_Hiem.xlsx");
  };

  const hasPendingTasks = tasks.some((t) => t.status === 'idle' || t.status === 'error');

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center py-10 px-4 font-sans">
      <div className="max-w-3xl w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-gray-100">
        <div className="p-8 text-center bg-blue-50/50 border-b border-gray-100">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Trích xuất Hạn Bảo Hiểm</h1>
          <p className="text-gray-500">Tải lên nhiều file Word (.docx) để tìm kiếm thời gian bảo hiểm cháy nổ hết hạn</p>
        </div>

        <div className="p-8 space-y-6">
          <div
            className="border-2 border-dashed border-gray-300 rounded-xl p-8 text-center hover:border-blue-500 hover:bg-blue-50/30 transition-colors cursor-pointer group"
            onClick={() => fileInputRef.current?.click()}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
          >
            <input
              type="file"
              ref={fileInputRef}
              className="hidden"
              multiple
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={handleFileChange}
            />
            <div className="mx-auto w-16 h-16 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
              <Upload className="w-8 h-8" />
            </div>
            <h3 className="text-lg font-medium text-gray-900 mb-1">Kéo thả các file vào đây</h3>
            <p className="text-sm text-gray-500">Hoặc click để chọn nhiều file từ máy tính</p>
            <p className="text-xs text-gray-400 mt-2">Chỉ hỗ trợ file Word (.docx)</p>
          </div>

          {tasks.length > 0 && (
            <div className="space-y-4">
              <div className="flex items-center justify-between border-b pb-4">
                <h3 className="text-lg font-semibold text-gray-800">
                  Danh sách file ({tasks.length})
                </h3>
                <div className="flex items-center gap-2">
                  <button
                    onClick={exportToExcel}
                    disabled={tasks.filter(t => t.status === 'success').length === 0 || isProcessing}
                    className="px-4 py-2 bg-green-600 text-white text-sm font-medium rounded-lg hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2"
                  >
                    <Download className="w-4 h-4" />
                    Xuất Excel
                  </button>
                  <button
                    onClick={processAllFiles}
                    disabled={!hasPendingTasks || isProcessing}
                    className="px-5 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2"
                  >
                    {isProcessing ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        Đang phân tích...
                      </>
                    ) : (
                      'Phân tích tất cả'
                    )}
                  </button>
                </div>
              </div>

              <div className="space-y-4">
                {tasks.map((task) => (
                  <div key={task.id} className="bg-white rounded-lg p-4 border border-gray-200 shadow-sm">
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex items-center space-x-3 overflow-hidden flex-1">
                        <FileText className="w-8 h-8 text-blue-500 flex-shrink-0" />
                        <div className="truncate">
                          <p className="text-sm font-medium text-gray-900 truncate" title={task.file.name}>
                            {task.file.name}
                          </p>
                          <p className="text-xs text-gray-500">{(task.file.size / 1024).toFixed(1)} KB</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-3">
                        {task.status === 'processing' && (
                          <span className="flex items-center text-blue-600 text-sm font-medium">
                            <Loader2 className="w-4 h-4 animate-spin mr-1" />
                            Đang xử lý
                          </span>
                        )}
                        {task.status === 'success' && (
                          <span className="flex items-center text-green-600 text-sm font-medium">
                            <CheckCircle className="w-4 h-4 mr-1" />
                            Hoàn thành
                          </span>
                        )}
                        {task.status === 'error' && (
                          <span className="flex items-center text-red-600 text-sm font-medium">
                            <X className="w-4 h-4 mr-1" />
                            Lỗi
                          </span>
                        )}
                        
                        <button
                          onClick={() => removeTask(task.id)}
                          disabled={isProcessing && task.status === 'processing'}
                          className="p-1 hover:bg-red-50 text-gray-400 hover:text-red-500 rounded-md transition-colors disabled:opacity-50"
                          title="Xóa file"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>

                    {task.status === 'success' && task.result && (
                      <div className="mt-3 pt-3 border-t border-gray-100">
                        <div className="bg-green-50/50 p-3 rounded-lg border border-green-100 text-gray-800 text-sm prose prose-sm max-w-none">
                          {task.result}
                        </div>
                      </div>
                    )}
                    
                    {task.status === 'error' && task.error && (
                      <div className="mt-3 pt-3 border-t border-gray-100">
                        <div className="bg-red-50 text-red-700 p-3 rounded-lg flex items-start gap-2 border border-red-100 text-sm">
                          <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
                          <p>{task.error}</p>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
