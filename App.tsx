
import React, { useState, useRef, useEffect } from 'react';
import { ClassData, GradingResult, GradingReport } from './types';
import { processClassGrading } from './services/geminiService';
import * as XLSX from 'xlsx';
import { User } from 'firebase/auth';
import { initAuth, googleSignIn, logout } from './services/auth';
import { extractSpreadsheetId, fetchSpreadsheetSheets, fetchSheetDataAsTsv, autoNormalizeToTsv } from './services/sheetsService';
import { 
  Sparkles,
  Play,
  Video,
  MessageSquareHeart,
  Loader2,
  Copy,
  Check,
  Link as LinkIcon,
  RefreshCw,
  Image as ImageIcon,
  Download,
  UserCheck,
  GraduationCap,
  Trophy,
  Award,
  Medal,
  Frown,
  Clock,
  Target,
  Zap,
  Star,
  Quote,
  ChevronRight,
  PlusCircle,
  FileText,
  Trash2,
  VideoOff,
  ShieldCheck,
  BookOpenCheck,
  BrainCircuit,
  AlertTriangle,
  FileSearch,
  ExternalLink,
  ClipboardList,
  LogOut,
  User as LucideUser
} from 'lucide-react';

const FEEDBACK_META = [
  { label: 'KẾT QUẢ', icon: <Target className="w-3.5 h-3.5" />, color: 'bg-blue-500', text: 'text-blue-600', ring: 'ring-blue-100' },
  { label: 'HÀNH ĐỘNG', icon: <Zap className="w-3.5 h-3.5" />, color: 'bg-amber-500', text: 'text-amber-600', ring: 'ring-amber-100' },
  { label: 'CAMERA', icon: <Video className="w-3.5 h-3.5" />, color: 'bg-purple-500', text: 'text-purple-600', ring: 'ring-purple-100' },
  { label: 'TƯƠNG TÁC', icon: <MessageSquareHeart className="w-3.5 h-3.5" />, color: 'bg-rose-500', text: 'text-rose-600', ring: 'ring-rose-100' },
  { label: 'LỜI CHÚC & LỜI NGỎ', icon: <Star className="w-3.5 h-3.5" />, color: 'bg-emerald-500', text: 'text-emerald-600', ring: 'ring-emerald-100' },
];

const App: React.FC = () => {
  const [classData, setClassData] = useState<ClassData>({
    markingGuide: '',
    sheetUrl: '',
    sheetData: '',
    camVisibleList: '',
    camHiddenList: '',
    praiseList: ''
  });
  
  const [results, setResults] = useState<GradingResult[]>([]);
  const [showResultsView, setShowResultsView] = useState(false);
  const [validationWarnings, setValidationWarnings] = useState<string[]>([]);
  const [answerKey, setAnswerKey] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState('');
  const [isFetching, setIsFetching] = useState(false);
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [copiedTable, setCopiedTable] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Auth and Google Sheets integration states
  const [user, setUser] = useState<User | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [needsAuth, setNeedsAuth] = useState(true);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [sheetsList, setSheetsList] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [isFetchingSheets, setIsFetchingSheets] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isNormalized, setIsNormalized] = useState(false);

  useEffect(() => {
    if (classData.sheetData && classData.sheetData.trim() !== '') {
      const hasTabs = classData.sheetData.includes('\t');
      setIsNormalized(hasTabs);
    } else {
      setIsNormalized(false);
    }
  }, [classData.sheetData]);

  useEffect(() => {
    const unsubscribe = initAuth(
      (currentUser, accessToken) => {
        setUser(currentUser);
        setToken(accessToken);
        setNeedsAuth(false);
      },
      () => {
        setUser(null);
        setToken(null);
        setNeedsAuth(true);
      }
    );
    return () => {
      if (typeof unsubscribe === 'function') {
        unsubscribe();
      }
    };
  }, []);

  const handleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setUser(result.user);
        setToken(result.accessToken);
        setNeedsAuth(false);
      }
    } catch (err: any) {
      console.error('Login failed:', err);
      setErrorMessage('Đăng nhập thất bại: ' + (err.message || err));
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    try {
      await logout();
      setUser(null);
      setToken(null);
      setNeedsAuth(true);
      setSheetsList([]);
      setSelectedSheet('');
    } catch (err: any) {
      console.error('Logout failed:', err);
    }
  };

  // Automatically fetch sheet names when URL or token changes
  useEffect(() => {
    const loadSheetsList = async () => {
      const id = extractSpreadsheetId(classData.sheetUrl);
      if (id && token) {
        setIsFetchingSheets(true);
        try {
          const sheets = await fetchSpreadsheetSheets(id, token);
          setSheetsList(sheets);
          if (sheets.length > 0) {
            setSelectedSheet(prev => prev && sheets.includes(prev) ? prev : sheets[0]);
          } else {
            setSelectedSheet('');
          }
        } catch (err) {
          console.error('Error fetching sheets list:', err);
          setSheetsList([]);
          setSelectedSheet('');
        } finally {
          setIsFetchingSheets(false);
        }
      } else {
        setSheetsList([]);
        setSelectedSheet('');
      }
    };
    loadSheetsList();
  }, [classData.sheetUrl, token]);

  const resizeAndCompressImage = (fileOrBlob: File | Blob): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const maxDim = 1000;
          let width = img.width;
          let height = img.height;

          if (width > maxDim || height > maxDim) {
            if (width > height) {
              height = Math.round((height * maxDim) / width);
              width = maxDim;
            } else {
              width = Math.round((width * maxDim) / height);
              height = maxDim;
            }
          }

          const canvas = document.createElement("canvas");
          canvas.width = width;
          canvas.height = height;

          const ctx = canvas.getContext("2d");
          if (!ctx) {
            resolve(e.target?.result as string);
            return;
          }

          ctx.drawImage(img, 0, 0, width, height);
          const compressedBase64 = canvas.toDataURL("image/jpeg", 0.75); // 75% quality JPEG
          resolve(compressedBase64);
        };
        img.onerror = () => {
          reject(new Error("Không thể đọc định dạng ảnh này. Vui lòng chọn ảnh khác."));
        };
        img.src = e.target?.result as string;
      };
      reader.onerror = () => reject(new Error("Không thể đọc file."));
      reader.readAsDataURL(fileOrBlob);
    });
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      try {
        setProcessingStatus('Đang tối ưu dung lượng ảnh...');
        const optimizedSrc = await resizeAndCompressImage(file);
        setClassData(prev => ({ ...prev, testImage: optimizedSrc }));
      } catch (err: any) {
        console.error("Optimizing uploaded image failed:", err);
        setErrorMessage(err.message || "Không thể tải lên và tối ưu hóa hình ảnh.");
      }
    }
  };

  useEffect(() => {
    const handlePaste = async (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
          const blob = items[i].getAsFile();
          if (blob) {
            try {
              setProcessingStatus('Đang tối ưu dung lượng ảnh...');
              const optimizedSrc = await resizeAndCompressImage(blob);
              setClassData(prev => ({ ...prev, testImage: optimizedSrc }));
            } catch (err: any) {
              console.error("Optimizing pasted image failed:", err);
              setErrorMessage(err.message || "Không thể dán và tối ưu hóa hình ảnh.");
            }
          }
          break;
        }
      }
    };

    window.addEventListener('paste', handlePaste);
    return () => window.removeEventListener('paste', handlePaste);
  }, []);

  const fetchSheetData = async (targetSheetName?: string) => {
    if (!classData.sheetUrl) return;
    setIsFetching(true);
    try {
      const spreadsheetId = extractSpreadsheetId(classData.sheetUrl);
      if (spreadsheetId && token) {
        let activeSheet = targetSheetName || selectedSheet;
        if (!activeSheet) {
          const sheets = sheetsList.length > 0 ? sheetsList : await fetchSpreadsheetSheets(spreadsheetId, token);
          if (sheets.length > 0) {
            activeSheet = sheets[0];
            setSelectedSheet(sheets[0]);
            setSheetsList(sheets);
          } else {
            throw new Error("Không tìm thấy trang tính nào trong tệp này.");
          }
        }
        const tsvText = await fetchSheetDataAsTsv(spreadsheetId, activeSheet, token);
        setClassData(prev => ({ ...prev, sheetData: autoNormalizeToTsv(tsvText) }));
      } else {
        // Fallback for public shared CSV link
        const convertToCsvUrl = (url: string) => {
          if (url.includes('/pubhtml')) return url.replace('/pubhtml', '/pub?output=csv');
          return url.includes('?') ? `${url}&output=csv` : `${url}?output=csv`;
        };
        const response = await fetch(convertToCsvUrl(classData.sheetUrl));
        if (!response.ok) throw new Error("Lỗi tải.");
        const text = await response.text();
        setClassData(prev => ({ ...prev, sheetData: autoNormalizeToTsv(text) }));
      }
    } catch (err: any) {
      setErrorMessage(err.message || "Không thể tải dữ liệu tự động. Hãy dán thủ công nội dung từ Sheet.");
    } finally { setIsFetching(false); }
  };

  const handleSheetTabChange = async (sheetTitle: string) => {
    setSelectedSheet(sheetTitle);
    const id = extractSpreadsheetId(classData.sheetUrl);
    if (id && token) {
      setIsFetching(true);
      try {
        const tsvText = await fetchSheetDataAsTsv(id, sheetTitle, token);
        setClassData(prev => ({ ...prev, sheetData: autoNormalizeToTsv(tsvText) }));
      } catch (err: any) {
        setErrorMessage(err.message || "Lỗi khi tải dữ liệu từ trang tính.");
      } finally {
        setIsFetching(false);
      }
    }
  };

  const startGrading = async () => {
    setErrorMessage(null);
    if (!classData.sheetData || !classData.testImage) {
      setErrorMessage("Vui lòng tải ảnh đề bài và dán dữ liệu Sheet học sinh.");
      return;
    }
    setIsProcessing(true);
    setShowResultsView(false);
    setValidationWarnings([]);
    
    const statuses = [
      'Đang nhận diện cấu trúc TSV/Tab...',
      'Đang ánh xạ cột dữ liệu học sinh...',
      'Đang tái cấu trúc bài làm viết tay...',
      'Đang nháp ngầm & Tính toán lại...',
      'Đang rà soát sai sót & Đối chiếu...',
      'Đang hoàn thiện nhận xét cá nhân hóa...'
    ];
    let sIdx = 0;
    const interval = setInterval(() => {
      setProcessingStatus(statuses[sIdx % statuses.length]);
      sIdx++;
    }, 3000);
    
    try {
      const report = await processClassGrading(classData);
      setResults(report.results);
      setAnswerKey(report.answerKey);
      if (report.validationWarnings) {
        setValidationWarnings(report.validationWarnings);
      }
      setShowResultsView(true);
    } catch (error: any) {
      console.error(error);
      setErrorMessage(`Lỗi: ${error.message || "Có lỗi xảy ra trong quá trình chấm bài. Vui lòng thử lại."}`);
    } finally {
      clearInterval(interval);
      setIsProcessing(false);
    }
  };

  const exportToExcel = () => {
    if (results.length === 0) return;
    
    // Prepare the data for Excel
    const data = results.map((r, i) => ({
      'STT': i + 1,
      'Họ và tên học sinh': r.studentName,
      'Tên học sinh': r.firstName,
      'Bài làm của học sinh': r.studentAnswer,
      'Điểm': r.score,
      'Thời gian nộp bài': r.submissionTime,
      'Xếp hạng': r.rank,
      'Nhận xét chi tiết': r.feedback.join('\n\n')
    }));

    // Create a new worksheet and workbook
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Kết quả chấm bài");

    // Adjust column widths for better visual experience
    const wscols = [
      { wch: 6 },   // STT
      { wch: 25 },  // Họ và tên học sinh
      { wch: 15 },  // Tên học sinh
      { wch: 50 },  // Bài làm của học sinh
      { wch: 10 },  // Điểm
      { wch: 20 },  // Thời gian nộp bài
      { wch: 15 },  // Xếp hạng
      { wch: 100 }  // Nhận xét chi tiết
    ];
    worksheet['!cols'] = wscols;

    // Generate and download the file
    XLSX.writeFile(workbook, `KetQua_ThayVinh_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const copyResultsToClipboard = () => {
    if (results.length === 0) return;
    
    // Format cells in standard Tab-Separated Values (TSV) format
    const formatCell = (val: any) => {
      if (val === undefined || val === null) return '';
      let str = String(val);
      // Double quotes are escaped inside cells, and cells with tabs/newlines are wrapped in double quotes
      if (str.includes('\t') || str.includes('\n') || str.includes('"') || str.includes('\r')) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };

    const rows = results.map((res, idx) => {
      const feedbackText = res.feedback.join('\n');
      const cols = [
        formatCell(idx + 1),                      // STT
        formatCell(res.studentName),             // Họ tên HS
        formatCell(res.firstName),               // Tên
        formatCell(res.studentAnswer),           // Bài thi đầu ra của HS
        formatCell(res.score),                   // Điểm thi
        formatCell(''),                          // Điểm vở (trống)
        formatCell(res.submissionTime),          // Thời điểm nộp bài
        formatCell(res.rank),                    // Xếp hạng
        formatCell(feedbackText)                 // Nhận xét chi tiết
      ];
      return cols.join('\t');
    });

    const tsvContent = rows.join('\n');
    
    navigator.clipboard.writeText(tsvContent).then(() => {
      setCopiedTable(true);
      setTimeout(() => setCopiedTable(false), 2000);
    }).catch(err => {
      console.error('Failed to copy table data:', err);
      setErrorMessage('Không thể sao chép dữ liệu vào clipboard. Vui lòng thử lại.');
    });
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900 font-sans selection:bg-indigo-100 pb-24 transition-all">
      {/* Header Navigation */}
      <nav className="bg-white/95 backdrop-blur-xl border-b border-slate-200 sticky top-0 z-50 shadow-sm transition-all">
        <div className="max-w-[1440px] mx-auto px-4 md:px-10 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3 md:gap-5">
            <div className="w-10 h-10 md:w-12 md:h-12 bg-indigo-600 rounded-xl md:rounded-2xl flex items-center justify-center shadow-lg shadow-indigo-200 ring-2 md:ring-4 ring-indigo-50">
              <GraduationCap className="w-5 h-5 md:w-7 md:h-7 text-white" />
            </div>
            <div>
              <h1 className="text-lg md:text-2xl font-black tracking-tight text-slate-900 uppercase leading-none">Vinh<span className="text-indigo-600">Grading</span></h1>
            </div>
          </div>
          <div className="flex items-center gap-3 md:gap-6">
            <div className="hidden lg:flex items-center gap-8 text-[11px] font-black uppercase tracking-widest text-slate-400">
              <a href="#" className="hover:text-indigo-600 transition-colors">Hướng dẫn</a>
              <a href="#" className="hover:text-indigo-600 transition-colors">Mẫu Sheet</a>
            </div>

            {/* Google Authentication States */}
            {user ? (
              <div className="flex items-center gap-2 md:gap-3 bg-slate-50 border border-slate-200 pl-2 pr-4 py-1.5 rounded-full shadow-inner">
                {user.photoURL ? (
                  <img src={user.photoURL} alt={user.displayName || ''} className="w-7 h-7 md:w-8 md:h-8 rounded-full border border-slate-300" referrerPolicy="no-referrer" />
                ) : (
                  <div className="w-7 h-7 md:w-8 md:h-8 bg-indigo-100 text-indigo-700 rounded-full flex items-center justify-center font-bold text-xs" title={user.email || ''}>
                    {(user.displayName || 'V').charAt(0)}
                  </div>
                )}
                <div className="hidden md:block text-left">
                  <p className="text-[10px] font-black text-slate-800 leading-none">{user.displayName || 'Giáo viên'}</p>
                </div>
                <button onClick={handleLogout} className="p-1 text-slate-400 hover:text-rose-500 rounded-full transition-colors" title="Đăng xuất">
                  <LogOut className="w-4 h-4" />
                </button>
              </div>
            ) : (
              <button 
                onClick={handleLogin}
                disabled={isLoggingIn}
                className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 border border-slate-200 px-4 py-2 rounded-full text-[11px] font-bold shadow-sm transition-all hover:shadow hover:-translate-y-0.5"
              >
                {isLoggingIn ? (
                  <Loader2 className="w-4 h-4 animate-spin text-indigo-600" />
                ) : (
                  <svg className="w-4 h-4 shrink-0" viewBox="0 0 48 48">
                    <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                    <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                    <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                    <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                  </svg>
                )}
                <span>{isLoggingIn ? 'Đang kết nối...' : 'Đăng nhập Google'}</span>
              </button>
            )}

            <button 
              disabled={isProcessing}
              onClick={startGrading}
              className={`flex items-center gap-2 md:gap-3 px-4 md:px-8 py-2.5 md:py-3.5 rounded-xl md:rounded-2xl font-black text-xs md:text-sm shadow-2xl transition-all active:scale-95 group ${isProcessing ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-indigo-200 md:hover:-translate-y-0.5'}`}
            >
              {isProcessing ? (
                <>
                  <Loader2 className="w-4 h-4 md:w-5 md:h-5 animate-spin" />
                  <span className="animate-pulse hidden sm:inline">{processingStatus}</span>
                  <span className="animate-pulse sm:hidden">...</span>
                </>
              ) : (
                <>
                  <Play className="w-4 h-4 md:w-5 md:h-5 fill-current group-hover:scale-110 transition-transform" /> <span className="hidden sm:inline">Bắt đầu chấm bài</span><span className="sm:hidden">Chấm bài</span>
                </>
              )}
            </button>
          </div>
        </div>
      </nav>

      <main className={`mx-auto px-4 md:px-10 py-6 md:py-12 ${showResultsView ? 'w-full max-w-none' : 'max-w-[1440px]'}`}>
        {/* Error Banner */}
        {errorMessage && (
          <div className="mb-6 bg-red-50 border border-red-200 rounded-2xl p-4 md:p-6 shadow-md animate-in slide-in-from-top-4 duration-300 relative overflow-hidden flex items-start gap-3">
            <div className="w-8 h-8 rounded-full bg-red-100 flex items-center justify-center text-red-600 shrink-0">
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div className="flex-1 min-w-0">
              <h4 className="text-sm font-black text-red-800 uppercase tracking-wider mb-1">Cảnh báo lỗi</h4>
              <p className="text-xs md:text-sm font-medium text-red-700 leading-relaxed whitespace-pre-line">{errorMessage}</p>
            </div>
            <button 
              onClick={() => setErrorMessage(null)} 
              className="text-red-400 hover:text-red-600 p-1 rounded-lg transition-colors ml-auto"
              title="Đóng"
            >
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          </div>
        )}

        {/* Deep Thinking Header Banner */}
        {isProcessing && (
          <div className="mb-6 md:mb-12 bg-indigo-600 rounded-2xl md:rounded-[40px] p-5 md:p-8 shadow-2xl shadow-indigo-200 animate-in fade-in zoom-in duration-700 relative overflow-hidden">
            <div className="absolute top-0 right-0 w-64 h-64 bg-white/10 rounded-full -mr-20 -mt-20 blur-3xl"></div>
            <div className="absolute bottom-0 left-0 w-48 h-48 bg-indigo-400/20 rounded-full -ml-10 -mb-10 blur-2xl"></div>
            <div className="flex flex-col md:flex-row items-center gap-4 md:gap-8 text-white relative z-10">
              <div className="w-16 h-16 md:w-20 md:h-20 bg-white/20 rounded-2xl md:rounded-[32px] flex items-center justify-center backdrop-blur-md shadow-xl border border-white/20 shrink-0">
                <BrainCircuit className="w-8 h-8 md:w-10 md:h-10 animate-pulse" />
              </div>
              <div className="text-center md:text-left">
                <h3 className="text-lg md:text-2xl font-black tracking-tight uppercase">Rà soát chuyên sâu</h3>
                <p className="text-xs md:text-sm font-bold text-indigo-100 mt-1 md:mt-2 max-w-2xl opacity-90 leading-relaxed">
                  Đang tái cấu trúc dữ liệu TSV, kiểm tra logic các bước làm và đối chiếu sai số. Hệ thống đảm bảo kết quả chính xác tuyệt đối.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Validation Warnings Alert */}
        {validationWarnings.length > 0 && (
          <div className="mb-6 md:mb-12 bg-amber-50 border-2 border-amber-200 rounded-2xl md:rounded-[40px] p-5 md:p-8 shadow-sm animate-in slide-in-from-top-4 duration-500">
            <div className="flex items-start gap-4 md:gap-6">
              <div className="bg-amber-100 p-2.5 md:p-3.5 rounded-xl shadow-sm border border-amber-200/50">
                <AlertTriangle className="w-5 h-5 md:w-7 md:h-7 text-amber-600" />
              </div>
              <div className="flex-1">
                <h3 className="text-sm md:text-lg font-black text-amber-900 uppercase tracking-wider">Cảnh báo dữ liệu</h3>
                <div className="mt-2 grid grid-cols-1 gap-2">
                  {validationWarnings.map((warn, i) => (
                    <div key={i} className="bg-white/60 p-2 rounded-lg text-[10px] md:text-xs text-amber-800 font-bold border border-amber-100 flex items-center gap-2">
                      <span className="w-4 h-4 bg-amber-200 text-amber-700 rounded-full flex items-center justify-center text-[8px] shrink-0">{i+1}</span>
                      {warn}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {!showResultsView && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 md:gap-12">
          {/* Left Panel: Inputs */}
          <div className="lg:col-span-4 space-y-6 md:space-y-10">
            {/* Step 1: Handwriting Upload */}
            <div className="bg-white rounded-2xl md:rounded-[44px] shadow-sm border border-slate-200 p-6 md:p-8 transition-all hover:shadow-xl hover:shadow-slate-100">
              <h2 className="text-[10px] md:text-[12px] font-black uppercase text-slate-400 tracking-[0.2em] mb-4 md:mb-6 flex items-center gap-3">
                <span className="w-6 h-6 md:w-7 md:h-7 rounded-lg md:rounded-xl bg-slate-100 flex items-center justify-center text-slate-500 text-[10px] md:text-[11px]">01</span>
                Đề bài viết tay
              </h2>
              {!classData.testImage ? (
                <div onClick={() => fileInputRef.current?.click()} className="w-full aspect-video border-2 border-dashed border-slate-200 rounded-2xl md:rounded-[32px] flex flex-col items-center justify-center cursor-pointer hover:bg-indigo-50 hover:border-indigo-300 transition-all group relative overflow-hidden bg-slate-50/50">
                  <input type="file" ref={fileInputRef} onChange={handleImageUpload} accept="image/*" className="hidden" />
                  <div className="w-12 h-12 md:w-16 md:h-16 bg-white rounded-2xl shadow-md border border-slate-100 flex items-center justify-center mb-3 md:mb-4 group-hover:scale-110 group-hover:-rotate-3 transition-all">
                    <ImageIcon className="w-6 h-6 md:w-8 md:h-8 text-indigo-500" />
                  </div>
                  <span className="text-xs md:text-sm font-black text-slate-500 tracking-tight">Tải ảnh hoặc dán (Ctrl+V) đề bài</span>
                </div>
              ) : (
                <div className="relative rounded-2xl md:rounded-[32px] overflow-hidden group border-2 md:border-4 border-white shadow-xl aspect-video">
                  <img src={classData.testImage} className="w-full h-full object-cover" />
                  <div className="absolute inset-0 bg-black/60 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-4">
                    <button onClick={() => setClassData({...classData, testImage: undefined})} className="bg-white text-rose-500 p-3 md:p-4 rounded-full shadow-2xl hover:scale-110 active:scale-95 transition-all">
                      <Trash2 className="w-5 h-5 md:w-6 md:h-6" />
                    </button>
                  </div>
                </div>
              )}
              <div className="mt-6 md:mt-8 space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-2 flex items-center gap-2">
                  <ClipboardList className="w-3 h-3" /> Biểu mẫu chấm
                </label>
                <textarea 
                  className="w-full h-24 md:h-32 p-4 md:p-5 bg-slate-50 border border-slate-200 rounded-xl md:rounded-[28px] text-[12px] md:text-[13px] font-semibold text-slate-700 outline-none focus:ring-4 focus:ring-indigo-500/5 focus:border-indigo-500 transition-all resize-none shadow-inner" 
                  placeholder="Ví dụ: Câu 1 (5đ), Câu 2 (5đ)..." 
                  value={classData.markingGuide} 
                  onChange={e => setClassData({...classData, markingGuide: e.target.value})} 
                />
              </div>
            </div>

            {/* Step 2: Interactive Rewards */}
            <div className="bg-white rounded-2xl md:rounded-[44px] shadow-sm border border-slate-200 p-6 md:p-8 space-y-6 transition-all hover:shadow-xl hover:shadow-slate-100">
              <h2 className="text-[10px] md:text-[12px] font-black uppercase text-slate-400 tracking-[0.2em] mb-2 flex items-center gap-3">
                <span className="w-6 h-6 md:w-7 md:h-7 rounded-lg md:rounded-xl bg-slate-100 flex items-center justify-center text-slate-500 text-[10px] md:text-[11px]">02</span>
                Tương tác & Thưởng
              </h2>
              <div className="space-y-4">
                <div className="group">
                  <span className="text-[10px] font-black text-amber-500 uppercase mb-2 flex items-center gap-2 pl-2">
                    <Star className="w-3 h-3 fill-current animate-pulse" /> Danh sách khen (+1đ)
                  </span>
                  <p className="text-[10px] text-slate-400 font-medium pl-2 mb-2">Nhập danh sách học sinh tích cực phát biểu, tương tác để được cộng thêm 1 điểm.</p>
                  <textarea placeholder="Nhập tên học sinh (mỗi tên một dòng hoặc cách nhau bởi dấu phẩy)..." value={classData.praiseList} onChange={e => setClassData({...classData, praiseList: e.target.value})} className="w-full h-32 p-4 bg-amber-50/10 border border-amber-100 rounded-xl md:rounded-[24px] text-[11px] md:text-[12px] font-bold outline-none focus:border-amber-500 focus:ring-4 focus:ring-amber-500/5 transition-all resize-none" />
                </div>
              </div>
            </div>
          </div>

          {/* Right Panel: Results */}
          <div className="lg:col-span-8 space-y-6 md:space-y-12">
            {/* Step 3: Student Data */}
            <div className="bg-white rounded-2xl md:rounded-[44px] shadow-sm border border-slate-200 p-6 md:p-8 transition-all hover:shadow-xl hover:shadow-slate-100">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div className="flex items-center gap-3">
                  <h2 className="text-[10px] md:text-[12px] font-black uppercase text-slate-400 tracking-[0.2em] flex items-center gap-3">
                    <span className="w-6 h-6 md:w-7 md:h-7 rounded-lg md:rounded-xl bg-slate-100 flex items-center justify-center text-slate-500 text-[10px] md:text-[11px]">03</span>
                    Dữ liệu bài làm
                  </h2>
                  {isNormalized && classData.sheetData && (
                    <span className="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded-full text-[9px] md:text-[10px] font-extrabold border border-emerald-100 uppercase tracking-wider">
                      <ShieldCheck className="w-3.5 h-3.5 text-emerald-600 animate-pulse" /> Đã chuẩn hoá TSV
                    </span>
                  )}
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  {/* Google Sheets Tab Dropdown/Selector when logged in */}
                  {user && sheetsList.length > 0 && (
                    <div className="flex items-center gap-1.5 bg-indigo-50 border border-indigo-100/50 rounded-lg md:rounded-[20px] px-3 py-1.5 md:py-2">
                      <span className="text-[9px] md:text-[11px] font-black text-indigo-500 uppercase tracking-wider">Trang:</span>
                      <select 
                        value={selectedSheet} 
                        onChange={e => handleSheetTabChange(e.target.value)} 
                        className="bg-transparent text-[11px] md:text-[13px] font-black text-indigo-700 outline-none cursor-pointer pr-1"
                      >
                        {sheetsList.map((st, sI) => (
                          <option key={sI} value={st} className="font-bold text-slate-800">{st}</option>
                        ))}
                      </select>
                      {isFetchingSheets && <Loader2 className="w-3.5 h-3.5 animate-spin text-indigo-500" />}
                    </div>
                  )}

                  <div className="relative flex-1 sm:flex-initial group">
                    <LinkIcon className="absolute left-3 md:left-4 top-1/2 -translate-y-1/2 w-3 md:w-4 h-3 md:h-4 text-slate-400" />
                    <input type="text" placeholder="Link Google Sheet..." value={classData.sheetUrl} onChange={e => setClassData({...classData, sheetUrl: e.target.value})} className="bg-slate-50 border border-slate-200 rounded-lg md:rounded-[20px] pl-9 md:pl-11 pr-4 md:pr-5 py-2 md:py-3 text-[11px] md:text-[13px] font-bold w-full sm:w-60 lg:w-72 outline-none focus:border-indigo-500 transition-all" />
                  </div>
                  <button onClick={() => fetchSheetData()} disabled={isFetching || !classData.sheetUrl} className="p-2.5 md:p-3.5 bg-indigo-50 text-indigo-600 rounded-lg md:rounded-[20px] hover:bg-indigo-100 disabled:opacity-50 transition-all border border-indigo-100" title="Tải dữ liệu">
                    {isFetching ? <Loader2 className="w-4 h-4 md:w-5 md:h-5 animate-spin" /> : <RefreshCw className="w-4 h-4 md:w-5 md:h-5" />}
                  </button>
                </div>
              </div>

              {!user && classData.sheetUrl && extractSpreadsheetId(classData.sheetUrl) && (
                <div className="mb-4 bg-amber-50/50 border border-amber-100 rounded-xl p-3 flex items-center justify-between gap-4 animate-in fade-in duration-300">
                  <div className="flex items-center gap-2 text-xs text-amber-800 font-bold">
                    <span>💡 Đăng nhập Google ở góc phải để tải trực tiếp từ trang tính của bạn.</span>
                  </div>
                </div>
              )}

              <textarea 
                placeholder="Dán nội dung từ Google Sheet tại đây..." 
                value={classData.sheetData} 
                onChange={e => setClassData({...classData, sheetData: e.target.value})}
                onBlur={() => {
                  if (classData.sheetData) {
                    const normalized = autoNormalizeToTsv(classData.sheetData);
                    if (normalized !== classData.sheetData) {
                      setClassData(prev => ({ ...prev, sheetData: normalized }));
                    }
                  }
                }}
                onPaste={(e) => {
                  const pastedText = e.clipboardData.getData('text');
                  if (pastedText) {
                    e.preventDefault();
                    const normalized = autoNormalizeToTsv(pastedText);
                    setClassData(prev => ({ ...prev, sheetData: normalized }));
                  }
                }}
                className="w-full h-32 md:h-52 p-4 md:p-6 bg-slate-50 border border-slate-200 rounded-2xl md:rounded-[40px] text-[11px] md:text-[13px] font-mono font-bold outline-none focus:border-indigo-500 transition-all resize-none shadow-inner" 
              />
            </div>
          </div>
        </div>
        )}

        {showResultsView && (
          <div className="space-y-8 md:space-y-12 animate-in fade-in slide-in-from-bottom-6 duration-700 w-full relative z-10">
            {/* Answer Key Display */}
                {answerKey && (
                  <div className="bg-indigo-50 border border-indigo-100 rounded-2xl md:rounded-[44px] p-6 md:p-10 shadow-sm relative overflow-hidden group">
                    <div className="absolute top-0 right-0 p-8 opacity-5 md:opacity-10 group-hover:opacity-20 transition-opacity">
                      <BookOpenCheck className="w-24 md:w-32 h-24 md:h-32 text-indigo-900" />
                    </div>
                    <div className="flex items-center gap-3 md:gap-4 mb-4 md:mb-6 relative z-10">
                      <div className="bg-indigo-600 text-white p-2 md:p-3 rounded-lg md:rounded-2xl shadow-lg ring-2 md:ring-4 ring-indigo-100">
                        <BookOpenCheck className="w-5 h-5 md:w-6 md:h-6" />
                      </div>
                      <div>
                        <h3 className="text-sm md:text-lg font-black text-indigo-900 uppercase tracking-widest leading-none">Đáp án AI</h3>
                        <p className="text-[9px] md:text-[11px] font-black text-indigo-400 uppercase tracking-[0.2em] mt-1">Đối chiếu logic</p>
                      </div>
                    </div>
                    <div className="bg-white/70 backdrop-blur-md rounded-xl md:rounded-[32px] p-4 md:p-8 border border-indigo-100 shadow-inner relative z-10">
                      <p className="text-[12px] md:text-[14px] text-indigo-800 font-mono font-bold leading-relaxed whitespace-pre-wrap">{answerKey}</p>
                    </div>
                  </div>
                )}

                {/* Sorting and Control Bar */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 md:gap-8 sticky top-[80px] bg-[#F8FAFC]/95 backdrop-blur-md z-40 py-4 md:py-6 border-b border-slate-200">
                  <div className="flex items-center gap-3 md:gap-5">
                    <div className="bg-indigo-600 text-white p-2 md:p-3 rounded-lg md:rounded-2xl shadow-xl shadow-indigo-200">
                      <FileSearch className="w-5 h-5 md:w-6 md:h-6" />
                    </div>
                    <div>
                      <h3 className="text-lg md:text-2xl font-black text-slate-900 tracking-tighter uppercase leading-none">Kết quả rà soát</h3>
                      <div className="flex items-center gap-3 md:gap-5 mt-1.5 md:mt-2">
                        <span className="text-[9px] md:text-[11px] font-black text-slate-400 uppercase tracking-[0.2em] flex items-center gap-1.5">
                          <PlusCircle className="w-3.5 h-3.5" /> {results.length} HS
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-3 md:gap-5">
                    <button onClick={() => setShowResultsView(false)} className="flex items-center justify-center gap-2 bg-slate-100 text-slate-600 px-4 md:px-6 py-3 md:py-4 rounded-xl md:rounded-2xl text-[10px] md:text-[12px] font-black uppercase tracking-[0.2em] hover:bg-slate-200 transition-all shadow-sm">
                      Quay lại sửa dữ liệu
                    </button>
                    <button 
                      onClick={copyResultsToClipboard} 
                      className={`flex items-center justify-center gap-3 px-6 md:px-10 py-3 md:py-4 rounded-xl md:rounded-2xl text-[10px] md:text-[12px] font-black uppercase tracking-[0.2em] transition-all shadow-xl hover:-translate-y-0.5 active:scale-95 group ${copiedTable ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'bg-slate-800 text-white hover:bg-slate-900'}`}
                      title="Sao chép toàn bộ dòng dữ liệu từ hàng số 1 (không bao gồm dòng tiêu đề) dưới định dạng tệp Tabs (TSV) để dán trực tiếp vào Google Sheets / Excel."
                    >
                      {copiedTable ? (
                        <>
                          <Check className="w-4 h-4 md:w-5 md:h-5 text-emerald-300 animate-bounce" /> <span>Đã copy!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-4 h-4 md:w-5 md:h-5 text-indigo-300" /> <span>Copy Từ Hàng 1</span>
                        </>
                      )}
                    </button>
                    <button onClick={exportToExcel} className="flex items-center justify-center gap-3 bg-emerald-600 text-white px-6 md:px-10 py-3 md:py-4 rounded-xl md:rounded-2xl text-[10px] md:text-[12px] font-black uppercase tracking-[0.2em] hover:bg-emerald-700 transition-all shadow-xl hover:-translate-y-0.5 active:scale-95 group">
                      <Download className="w-4 h-4 md:w-5 md:h-5" /> <span>Xuất Excel</span>
                    </button>
                  </div>
                </div>

                {/* Table Layout */}
                <div className="overflow-x-auto bg-white rounded-xl shadow-sm border border-slate-300">
                  <table className="w-full text-left border-collapse text-[13px] md:text-[14px]">
                    <thead>
                      <tr className="bg-[#2d5845] text-white font-medium text-[12px] whitespace-nowrap">
                        <th className="p-3 md:p-4 border border-[#40755d] text-center">STT</th>
                        <th className="p-3 md:p-4 border border-[#40755d]">Họ tên HS</th>
                        <th className="p-3 md:p-4 border border-[#40755d]">Tên</th>
                        <th className="p-3 md:p-4 border border-[#40755d]">Bài thi đầu ra của HS</th>
                        <th className="p-3 md:p-4 border border-[#40755d] text-center">Điểm thi</th>
                        <th className="p-3 md:p-4 border border-[#40755d] text-center">Điểm vở</th>
                        <th className="p-3 md:p-4 border border-[#40755d] text-center">Thời điểm nộp bài</th>
                        <th className="p-3 md:p-4 border border-[#40755d] text-center">Xếp hạng</th>
                        <th className="p-3 md:p-4 border border-[#40755d] min-w-[400px]">
                          THẦY VINH VÀ THẦY CÔ PCN KÍNH GỬI PHỤ HUYNH NHẬN XÉT TÌNH HÌNH CỦA CON BUỔI HỌC HÔM NAY
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {results.map((res, idx) => {
                        return (
                          <tr key={idx} className="border-b border-dotted border-slate-400 hover:bg-slate-50 transition-colors">
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 text-center">{idx + 1}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 font-medium whitespace-nowrap text-slate-800">{res.studentName}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 text-slate-800">{res.firstName}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 whitespace-pre-wrap text-slate-700 min-w-[250px]">{res.studentAnswer}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 text-center font-bold text-slate-900">{res.score}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400"></td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 text-center text-slate-600 whitespace-nowrap">{res.submissionTime}</td>
                            <td className="p-3 md:p-4 border-r border-dotted border-slate-400 text-center">
                              <span className="text-slate-800">
                                {res.rank}
                              </span>
                            </td>
                            <td className="p-3 md:p-4 text-slate-800 align-top">
                              <div className="space-y-4">
                                {res.feedback.map((line, lIdx) => (
                                  <p key={lIdx} className="leading-relaxed whitespace-pre-wrap">{line}</p>
                                ))}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
          </div>
        )}
      </main>

      <footer className="fixed bottom-0 left-0 right-0 bg-white/95 backdrop-blur-xl border-t border-slate-200 py-3 md:py-4 px-4 md:px-12 z-50">
        <div className="max-w-[1440px] mx-auto flex items-center justify-between text-[9px] md:text-[11px] font-black text-slate-400 uppercase tracking-[0.2em]">
          <div className="flex items-center gap-4 md:gap-8">
            <span className="flex items-center gap-2">
              <span className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></span> 
              <span className="hidden sm:inline">System v3.7 Active</span>
              <span className="sm:hidden">v3.7 Active</span>
            </span>
            <span className="text-slate-200">|</span>
          </div>
          <p className="text-right">© 2026 THẦY VINH MATH</p>
        </div>
      </footer>
    </div>
  );
};

export default App;
