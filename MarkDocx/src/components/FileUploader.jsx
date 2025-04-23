// src/components/FileUploader.jsx
import { useCallback } from 'react'
import { useDropzone } from 'react-dropzone'

export default function FileUploader({ onFilesChange }) {
  const onDrop = useCallback((acceptedFiles) => {
    // Validate files
    const validFiles = acceptedFiles.filter(
      file => file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
              file.type === 'application/msword' ||
              // Một số trình duyệt có thể không nhận diện chính xác mime type, 
              // kiểm tra thêm phần mở rộng
              file.name.toLowerCase().endsWith('.docx') || 
              file.name.toLowerCase().endsWith('.doc')
    )
    
    if (validFiles.length > 0) {
      onFilesChange(validFiles)
    }
  }, [onFilesChange])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      'application/msword': ['.doc']
    }
  })

  return (
    <div 
      {...getRootProps()} 
      className={`p-10 border-2 border-dashed rounded-lg text-center cursor-pointer
        ${isDragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-300'}`}
    >
      <input {...getInputProps()} />
      {isDragActive ? (
        <p className="text-blue-500">Thả file vào đây ...</p>
      ) : (
        <div>
          <p className="text-gray-600">Kéo và thả file DOCX / DOC vào đây</p>
          <p className="text-sm text-gray-500">hoặc click để chọn file</p>
        </div>
      )}
    </div>
  )
}
