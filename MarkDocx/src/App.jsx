// src/App.jsx
import { useState, useEffect } from 'react'
import { supabase } from './supabase'
import FileUploader from './components/FileUploader'
import FileList from './components/FileList'
import Auth from './components/Auth'

// Địa chỉ URL của backend Flask
const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || 'http://127.0.0.1:5000'

export default function App() {
  const [files, setFiles] = useState([])
  const [isUploading, setIsUploading] = useState(false)
  const [isConverting, setIsConverting] = useState(false)
  const [error, setError] = useState(null)
  const [user, setUser] = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    // Kiểm tra người dùng hiện tại
    const getCurrentUser = async () => {
      try {
        const { data } = await supabase.auth.getSession()
        setUser(data.session?.user || null)
      } catch (error) {
        console.error('Error checking auth:', error)
      } finally {
        setLoading(false)
      }
    }

    // Lắng nghe sự kiện thay đổi trạng thái xác thực
    const { data: authListener } = supabase.auth.onAuthStateChange(
      async (event, session) => {
        setUser(session?.user || null)
      }
    )

    getCurrentUser()

    // Cleanup listener khi component unmount
    return () => {
      authListener?.subscription?.unsubscribe()
    }
  }, [])

  const handleSignOut = async () => {
    try {
      await supabase.auth.signOut()
    } catch (error) {
      console.error('Error signing out:', error)
    }
  }

  const handleUpload = async (newFiles) => {
    setIsUploading(true)
    setError(null)

    try {
      setFiles(prev => [...prev, ...newFiles])

      for (const file of newFiles) {
        const { data, error } = await supabase.storage
          .from('docx-files')
          .upload(`${user.id}/${Date.now()}-${file.name}`, file)

        if (error) throw error
      }
    } catch (err) {
      setError('Error uploading files: ' + err.message)
    } finally {
      setIsUploading(false)
    }
  }

  const handleRemove = (index) => {
    setFiles(prev => prev.filter((_, i) => i !== index))
  }

  // Hàm gửi file đến backend và tải file kết quả về
  const handleConvertAndDownload = async (format) => {
    if (files.length === 0) return
    
    setIsConverting(true)
    setError(null)
    
    try {
      // Tạo FormData để gửi file
      const formData = new FormData()
      
      // Thêm tất cả file vào formData
      files.forEach(file => {
        formData.append('files[]', file)
      })
      
      // Thêm tham số format
      formData.append('format', format)
      
      // Gửi request đến backend
      const response = await fetch(`${BACKEND_URL}/merge-and-convert`, {
        method: 'POST',
        body: formData,
      })
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Error converting files')
      }
      
      // Lấy filename từ Content-Disposition header
      const contentDisposition = response.headers.get('Content-Disposition')
      let filename = 'merged'
      if (format === 'markdown') {
        filename += '.md'
      } else {
        filename += '.txt'
      }
      
      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename="?([^"]*)"?/)
        if (filenameMatch && filenameMatch[1]) {
          filename = filenameMatch[1]
        }
      }
      
      // Tạo blob từ response
      const blob = await response.blob()
      
      // Tạo URL cho blob để tải về
      const url = window.URL.createObjectURL(blob)
      
      // Tạo một element a và trigger click để tải file
      const a = document.createElement('a')
      a.style.display = 'none'
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      
      // Cleanup
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)
      
    } catch (err) {
      setError('Error converting files: ' + err.message)
      console.error(err)
    } finally {
      setIsConverting(false)
    }
  }

  const handleDownloadMarkdown = () => handleConvertAndDownload('markdown')
  const handleDownloadTxt = () => handleConvertAndDownload('text')

  // Hiển thị loading trong khi kiểm tra xác thực
  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="animate-spin rounded-full h-12 w-12 border-4 border-blue-500 border-t-transparent"></div>
      </div>
    )
  }

  // Hiển thị trang đăng nhập nếu chưa xác thực
  if (!user) {
    return <Auth />
  }

  // Hiển thị ứng dụng chính khi đã đăng nhập
  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header với thông tin người dùng */}
      <header className="bg-white shadow">
        <div className="container mx-auto px-4 py-4 max-w-6xl flex justify-between items-center">
          <h1 className="text-xl font-bold text-gray-800">
            Chuyển đổi file DOCX / DOC sang Markdown và Text
          </h1>
          <div className="flex items-center space-x-4">
            <div className="text-sm text-gray-600">
              <span>Đã đăng nhập với: </span>
              <span className="font-medium">{user.email}</span>
            </div>
            <button
              onClick={handleSignOut}
              className="text-sm px-3 py-1.5 bg-gray-100 hover:bg-gray-200 rounded-md transition-colors"
            >
              Đăng xuất
            </button>
          </div>
        </div>
      </header>

      <div className="container mx-auto px-4 py-8 max-w-3xl">
        {/* Main Content */}
        <div className="bg-white rounded-lg shadow-md p-6">
          {/* Upload Area */}
          <FileUploader onFilesChange={handleUpload} />
          
          {/* Error Message */}
          {error && (
            <div className="mt-4 p-4 bg-red-50 text-red-500 rounded-md border border-red-200">
              <p className="flex items-center">
                <svg className="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                </svg>
                {error}
              </p>
            </div>
          )}

          {/* File List */}
          <FileList files={files} onRemove={handleRemove} />
          
          {/* Download Buttons */}
          <div className="mt-6 flex flex-col sm:flex-row justify-center gap-4">
            <button
              onClick={handleDownloadMarkdown}
              disabled={files.length === 0 || isUploading || isConverting}
              className={`px-6 py-2 rounded-md text-white font-medium flex items-center justify-center
                ${files.length === 0 || isUploading || isConverting
                  ? 'bg-gray-300 cursor-not-allowed'
                  : 'bg-blue-500 hover:bg-blue-600 transition-colors'
                }`}
            >
              <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              Tải Markdown
            </button>

            <button
              onClick={handleDownloadTxt}
              disabled={files.length === 0 || isUploading || isConverting}
              className={`px-6 py-2 rounded-md text-white font-medium flex items-center justify-center
                ${files.length === 0 || isUploading || isConverting
                  ? 'bg-gray-300 cursor-not-allowed'
                  : 'bg-green-500 hover:bg-green-600 transition-colors'
                }`}
            >
              <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              Tải Text
            </button>
          </div>
        </div>

        {/* Loading Indicator */}
        {(isUploading || isConverting) && (
          <div className="mt-4 text-center text-gray-600">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-4 border-blue-500 border-t-transparent"></div>
            <p className="mt-2">
              {isUploading ? 'Đang tải file lên...' : 'Đang chuyển đổi file...'}
            </p>
          </div>
        )}
      </div>
    </div>
  )
}
