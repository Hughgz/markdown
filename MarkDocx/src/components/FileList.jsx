import { useState } from 'react'

export default function FileList({ files, onRemove }) {
  const [expandedFile, setExpandedFile] = useState(null)

  if (!files.length) {
    return (
      <div className="mt-6 text-center text-gray-500 py-8 bg-gray-50 rounded-lg border-2 border-dashed">
        <svg 
          className="mx-auto h-12 w-12 text-gray-400"
          fill="none" 
          stroke="currentColor" 
          viewBox="0 0 24 24"
        >
          <path 
            strokeLinecap="round" 
            strokeLinejoin="round" 
            strokeWidth={2} 
            d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V7a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" 
          />
        </svg>
        <p className="mt-2">Chưa có file nào được tải lên</p>
      </div>
    )
  }

  return (
    <div className="mt-6">
      <h3 className="text-lg font-medium text-gray-900 mb-3">
        Files đã tải lên ({files.length})
      </h3>
      <div className="bg-white rounded-lg border border-gray-200 divide-y divide-gray-200">
        {files.map((file, index) => (
          <div 
            key={index}
            className={`p-4 transition-all duration-200 hover:bg-gray-50
              ${expandedFile === index ? 'bg-blue-50' : ''}`}
          >
            <div className="flex items-center justify-between">
              <div 
                className="flex items-center space-x-3 cursor-pointer"
                onClick={() => setExpandedFile(expandedFile === index ? null : index)}
              >
                {/* File Icon */}
                <div className="flex-shrink-0">
                  <svg 
                    className="h-8 w-8 text-blue-500"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path 
                      strokeLinecap="round" 
                      strokeLinejoin="round" 
                      strokeWidth={2} 
                      d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" 
                    />
                  </svg>
                </div>

                {/* File Info */}
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-gray-900 truncate">
                    {file.name}
                  </p>
                  <p className="text-sm text-gray-500">
                    {(file.size / 1024 / 1024).toFixed(2)} MB
                  </p>
                </div>
              </div>

              {/* Actions */}
              <div className="flex items-center space-x-2">
                <button
                  onClick={() => setExpandedFile(expandedFile === index ? null : index)}
                  className="p-1 rounded-full hover:bg-gray-200 transition-colors"
                  title={expandedFile === index ? "Thu gọn" : "Xem chi tiết"}
                >
                  <svg 
                    className={`h-5 w-5 text-gray-500 transform transition-transform
                      ${expandedFile === index ? 'rotate-180' : ''}`}
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path 
                      strokeLinecap="round" 
                      strokeLinejoin="round" 
                      strokeWidth={2} 
                      d="M19 9l-7 7-7-7" 
                    />
                  </svg>
                </button>
                <button
                  onClick={() => onRemove(index)}
                  className="p-1 rounded-full hover:bg-red-100 transition-colors group"
                  title="Xóa file"
                >
                  <svg 
                    className="h-5 w-5 text-red-400 group-hover:text-red-500"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path 
                      strokeLinecap="round" 
                      strokeLinejoin="round" 
                      strokeWidth={2} 
                      d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" 
                    />
                  </svg>
                </button>
              </div>
            </div>

            {/* Expanded Details */}
            {expandedFile === index && (
              <div className="mt-4 pl-11">
                <div className="bg-white rounded-lg border border-gray-200 p-4">
                  <dl className="grid grid-cols-1 gap-x-4 gap-y-3 sm:grid-cols-2">
                    <div>
                      <dt className="text-sm font-medium text-gray-500">Tên file</dt>
                      <dd className="mt-1 text-sm text-gray-900">{file.name}</dd>
                    </div>
                    <div>
                      <dt className="text-sm font-medium text-gray-500">Kích thước</dt>
                      <dd className="mt-1 text-sm text-gray-900">
                        {(file.size / 1024 / 1024).toFixed(2)} MB
                      </dd>
                    </div>
                    <div>
                      <dt className="text-sm font-medium text-gray-500">Loại file</dt>
                      <dd className="mt-1 text-sm text-gray-900">{file.type}</dd>
                    </div>
                    <div>
                      <dt className="text-sm font-medium text-gray-500">Ngày tải lên</dt>
                      <dd className="mt-1 text-sm text-gray-900">
                        {new Date().toLocaleDateString('vi-VN')}
                      </dd>
                    </div>
                  </dl>
                </div>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  )
} 