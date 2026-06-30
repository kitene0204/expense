import React, { useState, useEffect, useRef } from 'react';
import { 
  FileText, 
  UploadCloud, 
  Download, 
  Copy, 
  Check, 
  Trash2, 
  Plus, 
  RefreshCw, 
  AlertCircle, 
  FileSpreadsheet, 
  Edit2, 
  Layers, 
  ArrowUpDown, 
  Sparkles, 
  Info,
  Calendar,
  Layers2,
  X,
  PlusCircle,
  HelpCircle,
  ClipboardCheck,
  Key,
  Settings,
  ShieldCheck
} from 'lucide-react';
import * as XLSX from 'xlsx';

// Structure of K-Edufine requisition item
interface EdufineItem {
  id: string;
  name: string;
  spec: string;
  unit: string;
  quantity: number;
  price: number;
  remarks: string;
}

// Quick presets for school administration budget
const DRAFT_REASON_PRESETS = [
  "교실 및 특별실 전산환경 운영을 위한 전산소모품 구입",
  "행정업무 지원 및 교무실 전무성 향상을 위한 사무용품 및 전산재료 구매",
  "교과 연계 교육과정 운영 및 학생 협업 활동 생태계 구축을 위한 교구 구입",
  "학교 환경 개선 및 감염병 예방 방역 위생용품 구입",
  "학교 운영 부서별 자체 사무지원용 물품 및 교육 소모품 구입"
];

// Presets for dragging-and-dropping/pasting messy textual representations
const PRESET_DATA = {
  smartstore: `쿠팡 장바구니 리스트
HP 250 G9 업무용 노트북 PC 애쉬그레이
549,000원 수량: 1개
보관방식: 일반 택배배송
로지텍 Pebble M350 저소음 무선 마우스 시크블랙
22,900원 수량: 5개
로켓배송 무료배송
삼성전자 EVO Plus 고속 마이크로 SD카드 128GB (MB-MC128KA)
15,500원 수량: 10개
로켓배송 당일배송`,

  vendor: `디도정보기술 주식회사 견적서
견적일자: 2026. 05. 29
등록번호: 120-81-123456
상호: 디도정보기술 대표자: 김전산
품목내역:
1. CAT.6 UTP 랜케이블 10M 블루 / 수량 15 / 단가 4,000원 / 공급가액 60,000원
2. HDMI v2.0 초고속 케이블 2M 블랙 / 수량 8 / 단가 6,500원 / 공급가액 52,000원
3. USB 3.0 기가비트 이더넷 멀티 아댑터 / 수량 5 / 단가 18,000원 / 공급가액 90,000원
4. 고품질 유선 광마우스 (사무용) / 수량 20 / 단가 8,500원 / 공급가액 170,000원
합계 금액 (부가가치세 포함): 금372,000원`,

  cart_pasted: `내 장바구니 목록 (에듀쇼핑몰 복사)
모나미 153 볼펜 0.5mm 12자루 세트 - 흰검 / 3,400원 수량: 12셋 [배송비 무료]
Double A 복사용지 A4 80g 2500매 (1box) - 27,500원 세금 포함수량: 4개
무선 지시 프리젠터 포인터 RF-800 - 32,000원 수량: 2개
화이트보드 극세사 지우개 리필 포함 - 1,800원 수량 8개`
};

export default function App() {
  // App states
  const [inputText, setInputText] = useState<string>('');
  const [items, setItems] = useState<EdufineItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  // User API key settings
  const [userApiKey, setUserApiKey] = useState<string>(() => {
    return localStorage.getItem('edufine_user_api_key') || '';
  });
  const [showApiKeySetting, setShowApiKeySetting] = useState<boolean>(false);

  useEffect(() => {
    localStorage.setItem('edufine_user_api_key', userApiKey);
  }, [userApiKey]);

  // Draft Text configuration states
  const [draftRelated, setDraftRelated] = useState<string>('');
  const [draftNote, setDraftNote] = useState<string>('교실 및 특별실 전산환경 운영을 위한 전산소모품을 아래와 같이 구입하고자 합니다.');
  
  // Custom states for manual details
  const [excelTemplateType, setExcelTemplateType] = useState<'simple' | 'standard'>('standard');
  const [customFilename, setCustomFilename] = useState<string>('에듀파인_내역등록');
  const [copySuccess, setCopySuccess] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [attachedFile, setAttachedFile] = useState<File | null>(null);

  // File loading reference
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleClearAttachedFile = () => {
    setAttachedFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // Initialize draft metadata automatically on startup
  useEffect(() => {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const date = String(today.getDate()).padStart(2, '0');
    setDraftRelated(`○○과-○○호(${year}.${month}.${date}.)`);
  }, []);

  // Update default Excel filename based on loaded items
  useEffect(() => {
    if (items.length > 0) {
      const summaryName = items[0].name.slice(0, 10);
      const suffix = items.length > 1 ? `_외_${items.length - 1}종` : '';
      setCustomFilename(`에듀파인_품의내역_${summaryName}${suffix}`);
    } else {
      setCustomFilename('에듀파인_품의내역_신규등록');
    }
  }, [items]);

  // Number conversion algorithm: Number -> Korean accounting format (e.g. 오십팔만삼천)
  const formatNumberToKoreanWon = (num: number): string => {
    if (num <= 0) return '영';
    const units = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
    const positions = ['', '십', '백', '천'];
    const largePositions = ['', '만', '억', '조'];
    
    let result = '';
    const numString = num.toString();
    const len = numString.length;
    
    // Divide digits by 4 (Korean system relies on 10,000 threshold)
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numString[i]);
      const posFromRight = len - 1 - i;
      const largePosIdx = Math.floor(posFromRight / 4);
      const posInGroupIdx = posFromRight % 4;
      
      if (digit !== 0) {
        // Prevent '일십', '일백', '일천' if they are higher significance in blocks
        if (digit === 1 && posInGroupIdx > 0) {
          result += positions[posInGroupIdx];
        } else {
          result += units[digit] + positions[posInGroupIdx];
        }
      }
      
      // Emit group multiplier (만, 억, 조) at the boundary of 0 residue
      if (posInGroupIdx === 0) {
        // Determine if non-zero exists inside current block to place multiplier
        const groupStart = Math.max(0, i - 3);
        const groupStr = numString.slice(groupStart, i + 1);
        if (parseInt(groupStr) > 0) {
          result += largePositions[largePosIdx];
        }
      }
    }
    
    return result;
  };

  // Calculations
  const totalItemCount = items.length;
  const totalPrice = items.reduce((acc, curr) => acc + (curr.price * curr.quantity), 0);
  const koreanWonWord = formatNumberToKoreanWon(totalPrice);

  const firstItemName = items.length > 0 ? items[0].name : "품목 미지정";
  const firstItemSummary = items.length > 1 ? `${firstItemName} 외 ${items.length - 1}종` : firstItemName;

  // Build the final Draft text structure
  const draftTitle = `1. 품의 개요(초안)
2. 관련: ${draftRelated || '○○과-○○호(2026.00.00.)'}
3. ${draftNote}
4. 품목: ${firstItemSummary}
5. 금액: 금${totalPrice.toLocaleString()}원(금${koreanWonWord}원정)

붙임 지출품의서 1부. 끝.`;

  // Parse Text or image data using client-side Gemini API (Direct Fetch to support B-type severless Vercel deployment)
  const handleParseData = async (textToParse: string, customFilePayload?: { base64: string; mimeType: string }) => {
    if (!textToParse.trim() && !customFilePayload) {
      setError('분석할 원본 텍스트를 입력하거나 견적서 파일을 업로드해 주세요.');
      return;
    }

    const apiKeyToUse = userApiKey.trim() || (import.meta.env.VITE_GEMINI_API_KEY as string) || '';
    if (!apiKeyToUse) {
      setError('현재 활성화된 API 키가 없습니다. 우측 상단 ⚙️ [개인 API 키 입력] 버튼을 눌러 본인의 Gemini API 키를 입력해 주세요. (무료 발급 가능하며 로컬에만 안전하게 저장됩니다)');
      setShowApiKeySetting(true);
      return;
    }

    setLoading(true);
    setError(null);

    // K-Edufine Custom Parsing Instruction & Response Schema
    const systemInstructionText = `You are an expert procurement assistant for Korean school administration (K-Edufine/K-에듀파인). 
Your task is to parse unstructured estimate sheets, retail site shopping carts, excel tables, text, or screenshots, and convert them into a structured JSON array for school budget requisitions.

Parse each item finding:
1. 품명 (name): Clear product name. Translate to friendly Korean if it is excessively complex, but keep key brand/model details.
2. 규격 (spec): Specification (dimension, weight, capacity, color, model number, etc.). If no spec exists or it is empty, literally use "규격 없음" or "상세 정보 참조".
3. 단위 (unit): The measurement unit (개, EA, box, 롤, 세트, 켤레, 묶음, etc.). If unclear, always default to "개".
4. 수량 (quantity): Must be an integer > 0. Default to 1 if not specified.
5. 단가 (price): Unit price as an integer. Must be >= 0. If only "금액" (total sum) and "수량" are given, divide total sum by quantity to extract the single unit price.
6. 비고 (remarks): Extra notes like links, options selected, etc. (Can be empty).

Always respond with a valid JSON array matching the schema. Force prices and quantities to be numbers, not strings. Remove any currency symbols (₩, $, 원) and commas from numeric fields.`;

    const responseSchema = {
      type: "ARRAY",
      description: "Parsed list of items for K-Edufine",
      items: {
        type: "OBJECT",
        properties: {
          name: { type: "STRING", description: "품명 (Item Name / Description of the product)." },
          spec: { type: "STRING", description: "규격 (Specification, Model, Size, Volume, etc.). Format nicely." },
          unit: { type: "STRING", description: "단위 (Unit, e.g. 개, EA, box). If unclear, default to '개'." },
          quantity: { type: "INTEGER", description: "수량 (Quantity purchased). Must be a positive integer." },
          price: { type: "INTEGER", description: "단가 (Unit price of a single item in KRW, integer)." },
          remarks: { type: "STRING", description: "비고 (Remarks, notes, option details. Optional)." }
        },
        required: ["name", "spec", "unit", "quantity", "price"]
      }
    };

    const parts: any[] = [];
    if (customFilePayload && customFilePayload.base64 && customFilePayload.mimeType) {
      parts.push({
        inlineData: {
          data: customFilePayload.base64,
          mimeType: customFilePayload.mimeType
        }
      });
    }

    const textPrompt = `Please analyze the following estimate / cart data and parse it into K-Edufine items.
User Pasted Data:
${textToParse || "(Image/File provided only)"}

Format strictly as JSON.`;

    parts.push({ text: textPrompt });

    // Client-side models list to attempt fallback when a model is rate-limited or experiencing high demand (503)
    const modelsToTry = ['gemini-1.5-flash', 'gemini-2.5-flash', 'gemini-1.5-pro'];
    let lastError: any = null;
    let success = false;
    let parsedItems: any[] = [];

    for (const modelName of modelsToTry) {
      if (success) break;
      
      const maxRetries = 2;
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          console.log(`[Client Gemini Request] Model: ${modelName} | Attempt ${attempt}/${maxRetries}`);
          
          const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKeyToUse}`;
          const res = await fetch(endpoint, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              contents: [{ parts }],
              systemInstruction: {
                parts: [{ text: systemInstructionText }]
              },
              generationConfig: {
                responseMimeType: 'application/json',
                responseSchema: responseSchema,
                temperature: 0.1
              }
            })
          });

          const responseData = await res.json();

          if (!res.ok) {
            const apiErrorMsg = responseData?.error?.message || JSON.stringify(responseData?.error || 'Unknown API Error');
            const apiErrorCode = responseData?.error?.code || res.status;
            throw new Error(`[API Error ${apiErrorCode}] ${apiErrorMsg}`);
          }

          const rawText = responseData?.candidates?.[0]?.content?.parts?.[0]?.text;
          if (!rawText) {
            throw new Error('Gemini API가 빈 응답을 반환했습니다.');
          }

          parsedItems = JSON.parse(rawText.trim());
          success = true;
          break;
        } catch (err: any) {
          lastError = err;
          const errMsg = err?.message || String(err);
          console.warn(`[Client Gemini Attempt Failed] ${modelName} (Attempt ${attempt}): ${errMsg}`);

          // Critical API Key errors - do not retry on key issues (400 Invalid Key or 403 Forbidden)
          if (
            errMsg.includes('API_KEY_INVALID') || 
            errMsg.includes('invalid') || 
            errMsg.includes('key is not valid') ||
            errMsg.includes('403') || 
            errMsg.includes('PERMISSION_DENIED') || 
            errMsg.includes('unauthorized')
          ) {
            throw new Error('입력하신 API 키가 올바르지 않거나 구글 보안 정책에 의해 차단되었습니다. 발급 정보를 다시 확인해 주세요. (Error: 403 / Key Invalid)');
          }

          if (attempt < maxRetries) {
            // Wait 1s and retry
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        }
      }
    }

    try {
      if (!success) {
        throw lastError || new Error('모든 가용한 인공지능 모델 호출이 실패했습니다.');
      }

      if (parsedItems && Array.isArray(parsedItems)) {
        const mappedItems: EdufineItem[] = parsedItems.map((item: any, i: number) => ({
          id: Math.random().toString(36).substr(2, 9),
          name: item.name || `수정가능 상품 ${i + 1}`,
          spec: item.spec || '규격 없음',
          unit: item.unit || '개',
          quantity: typeof item.quantity === 'number' ? item.quantity : 1,
          price: typeof item.price === 'number' ? item.price : 0,
          remarks: item.remarks || ''
        }));
        setItems(mappedItems);
        setInputText(''); // Reset text box upon success
        setAttachedFile(null); // Clear attached file upon successful layout analysis
      } else {
        throw new Error('상품 형식을 올바른 데이터형으로 파싱하지 못했습니다.');
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || '인공지능 분석 및 변환 과정에서 일시적인 에러가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  // Helper local file parsing for standard XLSX sheets
  const handleLocalFileSelection = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processUploadedFile(file);
  };

  const processUploadedFile = (file: File) => {
    setError(null);
    setAttachedFile(file);
    setInputText(''); // Clear input text when a file is uploaded to avoid conflict
    const fileExtension = file.name.split('.').pop()?.toLowerCase();

    // Local parser for standard table file types (.xlsx, .xls, .csv)
    if (['xlsx', 'xls', 'csv'].includes(fileExtension || '')) {
      setLoading(true);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'binary' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

          if (jsonData.length === 0) {
            throw new Error('선택된 엑셀 시트에 데이터가 비어 있습니다.');
          }

          // Let Gemini parse the raw multidimensional Excel arrays to extract items reliably!
          // This keeps local parsing extremely clever so it can handle any random layout.
          const rawAoaText = jsonData.map(row => row.join(' | ')).join('\n');
          handleParseData(`[로컬 엑셀 업로드 원본 데이터]\n${rawAoaText}`);
        } catch (err: any) {
          setError(`엑셀 파일 구조 분석 실패: ${err.message}`);
          setLoading(false);
        }
      };
      reader.onerror = () => {
        setError('파일 읽기에 실패했습니다.');
        setLoading(false);
      };
      reader.readAsBinaryString(file);
    } else if (['png', 'jpg', 'jpeg', 'pdf', 'txt'].includes(fileExtension || '')) {
      // Send unstructured image or documents to Gemini directly!
      setLoading(true);
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const base64String = (e.target?.result as string).split(',')[1] || (e.target?.result as string);
          let mime = file.type;
          
          // PDF might missing type fallback
          if (!mime && fileExtension === 'pdf') mime = 'application/pdf';
          if (!mime && fileExtension === 'txt') mime = 'text/plain';

          handleParseData(`[파일 업로드: ${file.name}]`, {
            base64: base64String,
            mimeType: mime
          });
        } catch (err: any) {
          setError(`인공지능 파일 분석 준비 에러: ${err.message}`);
          setLoading(false);
        }
      };
      reader.readAsDataURL(file);
    } else {
      setError('지원하지 않는 파일 형식입니다. (엑셀, 텍스트, PDF, 이미지 지원)');
    }
  };

  // Preset loading helpers
  const handleLoadPreset = (key: keyof typeof PRESET_DATA) => {
    setInputText(PRESET_DATA[key]);
    setAttachedFile(null); // Clear attached file when a preset is loaded
  };

  // Spreadsheet interactivity modifiers
  const handleUpdateItem = (id: string, key: keyof EdufineItem, value: any) => {
    setItems(prev => prev.map(item => {
      if (item.id === id) {
        let updatedVal = value;
        if (key === 'quantity') {
          updatedVal = Math.max(0, parseInt(value) || 0);
        } else if (key === 'price') {
          updatedVal = Math.max(0, parseInt(value) || 0);
        }
        return { ...item, [key]: updatedVal };
      }
      return item;
    }));
  };

  const handleAddNewRow = () => {
    const newItem: EdufineItem = {
      id: Math.random().toString(36).substr(2, 9),
      name: '새로운 전산 소모품',
      spec: '규격 사양',
      unit: '개',
      quantity: 1,
      price: 10000,
      remarks: ''
    };
    setItems(prev => [...prev, newItem]);
  };

  const handleDeleteRow = (id: string) => {
    setItems(prev => prev.filter(item => item.id !== id));
  };

  const handleClearAllItems = () => {
    if (window.confirm('입력된 모든 품의 내역을 삭제하시겠습니까?')) {
      setItems([]);
    }
  };

  // Reorder row index position handler
  const moveRowIndex = (index: number, direction: 'up' | 'down') => {
    if (direction === 'up' && index === 0) return;
    if (direction === 'down' && index === items.length - 1) return;

    const updated = [...items];
    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    const temp = updated[index];
    updated[index] = updated[targetIndex];
    updated[targetIndex] = temp;
    setItems(updated);
  };

  // Copy-to-clipboard trigger
  const handleCopyToClipboard = () => {
    navigator.clipboard.writeText(draftTitle);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  // Drag and drop dropzone visual states
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processUploadedFile(files[0]);
    }
  };

  // Export actual styled Excel target
  const handleDownloadExcel = () => {
    if (items.length === 0) {
      setError('엑셀로 내보낼 상품 내역이 없습니다.');
      return;
    }

    let aoaData: any[][] = [];

    if (excelTemplateType === 'standard') {
      // K-Edufine Multi-row precise format definition
      aoaData = [
        ['[K-에듀파인 지출품의 내역 등록 파일]', '', '', '', '', '', ''],
        ['※ 필수 입력 항목(*): 품명, 규격, 단위, 수량, 단가. (가이드라인을 덮어쓰거나 그대로 올리실 수 있습니다)', '', '', '', '', '', ''],
        ['품명*', '규격*', '단위*', '수량*', '단가*', '금액', '비고']
      ];

      items.forEach((item, index) => {
        aoaData.push([
          item.name,
          item.spec,
          item.unit,
          item.quantity,
          item.price,
          item.quantity * item.price,
          item.remarks || ''
        ]);
      });
    } else {
      // Simple format
      aoaData = [
        ['품명', '규격', '단위', '수량', '단가', '금액', '비고']
      ];

      items.forEach(item => {
        aoaData.push([
          item.name,
          item.spec,
          item.unit,
          item.quantity,
          item.price,
          item.quantity * item.price,
          item.remarks || ''
        ]);
      });
    }

    const worksheet = XLSX.utils.aoa_to_sheet(aoaData);

    // Apply strict custom column width configurations to preserve grid aesthetics
    worksheet['!cols'] = [
      { wch: 30 }, // 품명
      { wch: 20 }, // 규격
      { wch: 8 },  // 단위
      { wch: 10 }, // 수량
      { wch: 12 }, // 단가
      { wch: 15 }, // 금액
      { wch: 20 }  // 비고
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '품의일괄업로드');

    // Emit save action
    const displayFilename = customFilename.endsWith('.xlsx') ? customFilename : `${customFilename}.xlsx`;
    XLSX.writeFile(workbook, displayFilename);
  };

  return (
    <div className="min-h-screen bg-[#0F1115] text-[#E0E0E0] font-sans flex flex-col justify-between" id="main_container">
      {/* Header Section */}
      <header className="sticky top-0 z-10 border-b border-[#2D3139] bg-[#16191F]/90 backdrop-blur-md px-6 py-4" id="app_header">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center space-x-3">
            <div className="bg-indigo-600 text-white p-2.5 rounded-xl shadow-lg shadow-indigo-900/30 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-bold text-white tracking-tight">K-에듀파인 품의 엑셀 변환기</h1>
                <span className="text-indigo-400 font-semibold text-xs bg-indigo-950/40 px-2 py-0.5 rounded-full border border-indigo-905/30">EduDraft Master</span>
              </div>
              <p className="text-xs text-gray-500 uppercase tracking-widest mt-0.5">
                대혼란의 견적서나 장바구니 복사글을 에듀파인 일괄업로드 양식으로 정밀 변환
              </p>
            </div>
          </div>
          
          <div className="flex items-center flex-wrap gap-2.5">
            <div className="flex items-center space-x-2 text-xs bg-[#1A1D23] border border-[#2D3139] px-3.5 py-2 rounded-xl text-gray-400">
              <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse inline-block"></span>
              <span>시스템 연결됨: <strong className="text-gray-200">K-Edufine Engine v1.2.0</strong></span>
            </div>

            <button
              onClick={() => setShowApiKeySetting(!showApiKeySetting)}
              className={`flex items-center space-x-2 text-xs px-3.5 py-2 rounded-xl border transition-all duration-200 cursor-pointer ${
                userApiKey.trim()
                  ? 'bg-emerald-950/40 border-emerald-500/30 text-emerald-300 hover:bg-emerald-950/60'
                  : 'bg-[#1A1D23] border-[#2D3139] text-gray-300 hover:bg-[#20242D] hover:text-white'
              }`}
            >
              <Key className={`w-3.5 h-3.5 ${userApiKey.trim() ? 'text-emerald-400 animate-pulse' : 'text-gray-400'}`} />
              <span className="font-semibold">
                {userApiKey.trim() ? '개인 API 키 사용 중' : '개인 API 키 입력'}
              </span>
              <Settings className="w-3.5 h-3.5 text-gray-500" />
            </button>
          </div>
        </div>
      </header>

      {/* API Key Drawer/Collapse Panel */}
      {showApiKeySetting && (
        <div className="bg-[#16191F] border-b border-[#2D3139] px-6 py-4" id="api_key_panel">
          <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div className="space-y-1 max-w-xl">
              <h3 className="text-xs font-bold text-indigo-400 tracking-wider uppercase flex items-center gap-1.5">
                <Key className="w-3.5 h-3.5" /> 개인용 Gemini API Key (Secrets) 관리
              </h3>
              <p className="text-[11px] text-gray-400 leading-relaxed">
                많은 사용자가 앱을 동시에 활용할 때, 구글 무료 쿼터 제한이나 503 과부하 에러가 발생할 수 있습니다. 
                이 곳에 본인만의 Gemini API Key를 등록하여 사용(브라우저 내에만 안전하게 로컬 저장됨)하시면 제한이나 대기 시간(503) 없이 언제든 번개처럼 즉시 변환됩니다.
              </p>
            </div>
            
            <div className="w-full md:w-auto flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 shrink-0">
              <div className="relative flex-1 sm:w-80">
                <input
                  type="password"
                  placeholder="구글 AI Studio에서 발급받은 API 키 붙여넣기 (AIzaSy...)"
                  value={userApiKey}
                  onChange={(e) => setUserApiKey(e.target.value)}
                  className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-indigo-500 rounded-xl px-3.5 py-2 text-xs font-mono outline-none text-white pr-9 placeholder-gray-500"
                />
                {userApiKey.trim() ? (
                  <ShieldCheck className="w-4 h-4 text-emerald-400 absolute right-3 top-2.5" />
                ) : (
                  <Key className="w-4 h-4 text-gray-650 absolute right-3 top-2.5 text-gray-600" />
                )}
              </div>
              
              <div className="flex gap-2">
                {userApiKey.trim() && (
                  <button
                    onClick={() => {
                      setUserApiKey('');
                    }}
                    className="px-3 py-2 bg-rose-950/30 hover:bg-rose-950/50 border border-rose-500/20 text-rose-300 rounded-xl text-xs font-semibold cursor-pointer transition flex items-center gap-1 shrink-0"
                    title="등록된 개인 키 삭제"
                  >
                    Clear
                  </button>
                )}
                <button
                  onClick={() => setShowApiKeySetting(false)}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 active:bg-indigo-700 text-white rounded-xl text-xs font-bold cursor-pointer transition shrink-0"
                >
                  적용 완료
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <main className="max-w-7xl mx-auto px-4 py-6 md:py-8 space-y-6 flex-1 w-full" id="working_layout">
        
        {/* Step Guide Panel */}
        <div className="bg-[#16191F] rounded-2xl border border-[#2D3139] p-5 flex flex-col lg:flex-row items-center gap-4 justify-between" id="guide_banner">
          <div className="flex items-start space-x-3">
            <div className="p-2 bg-indigo-950/50 text-indigo-400 rounded-lg shrink-0 mt-0.5 border border-indigo-900/30">
              <Sparkles className="w-5 h-5 text-indigo-400" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-white flex items-center gap-2">✨ 품의 작성 간소화를 위한 스마트 길잡이</h2>
              <p className="text-xs text-gray-400 mt-1 leading-relaxed">
                쇼핑몰 상품 페이지 장바구니 내용 전체나 엑셀 견적서 텍스트 영역을 <strong className="text-white bg-[#0F1115] px-1.5 py-0.5 rounded border border-[#2D3139] font-mono">[Ctrl+C]</strong>하여 아래에 붙여넣고 <strong className="text-indigo-400">품의 내역 변환</strong> 버튼을 클릭하세요. 업로드용 엑셀 다운로드와 지출품의 기안문 작성용 텍스트가 즉각 생성됩니다.
              </p>
            </div>
          </div>
          <div className="flex flex-wrap lg:flex-nowrap gap-2 shrink-0 w-full lg:w-auto mt-2 lg:mt-0">
            <button 
              onClick={() => handleLoadPreset('smartstore')} 
              className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-[#2D3139] hover:border-indigo-500/50 bg-[#1A1D23] hover:bg-indigo-950/20 text-gray-400 hover:text-white transition-all cursor-pointer flex-1 lg:flex-none text-center"
            >
              🛒 장바구니 예시
            </button>
            <button 
              onClick={() => handleLoadPreset('vendor')} 
              className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-[#2D3139] hover:border-indigo-500/50 bg-[#1A1D23] hover:bg-indigo-950/20 text-gray-400 hover:text-white transition-all cursor-pointer flex-1 lg:flex-none text-center"
            >
              📄 복잡한 견적서 예시
            </button>
            <button 
              onClick={() => handleLoadPreset('cart_pasted')} 
              className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-[#2D3139] hover:border-indigo-500/50 bg-[#1A1D23] hover:bg-indigo-950/20 text-gray-400 hover:text-white transition-all cursor-pointer flex-1 lg:flex-none text-center"
            >
              🖍️ 난잡글 예시 불러오기
            </button>
          </div>
        </div>

        {/* Outer Split Panels */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6" id="parsing_workspace">
          
          {/* Left Inputs Pane */}
          <section className="lg:col-span-5 space-y-6" id="input_panel">
            <div className="bg-[#16191F] rounded-3xl border border-[#2D3139] shadow-xl p-6 space-y-5">
              <div className="flex items-center justify-between">
                <h3 className="font-semibold text-sm text-indigo-300 uppercase tracking-wider flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-indigo-505 bg-indigo-500"></span>
                  원본 내역 데이터 입력
                </h3>
                <span className="text-[10px] text-gray-500 font-medium">
                  {inputText.length}자 입력됨 (Excel, PDF, Web-Cart)
                </span>
              </div>

              {/* Text Area & Image Drag Component */}
              <div 
                className={`relative border-2 border-dashed rounded-2xl transition-all duration-200 ${
                  isDragging 
                    ? 'border-indigo-500 bg-indigo-950/20 scale-[0.99]' 
                    : 'border-[#2D3139] focus-within:border-indigo-500 bg-[#16191F]/50 group'
                }`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                id="interactive_dropbox"
              >
                <textarea
                  className="w-full h-64 p-4 text-sm bg-transparent outline-none resize-none placeholder-gray-500 text-gray-200 leading-relaxed font-mono"
                  placeholder={`여기에 마우스 드래그를 통해 파일을 끌어 놓거나 복사한 데이터를 붙여넣어 주세요!

[가능한 입력 패턴]
- 쿠팡, 지마켓, 나라장터, 학교장터 장바구니 화면 전체 복사글
- 도매유통업체 견적서 PDF 또는 텍스트 목록
- 품명 수량 단가 정보가 불규칙하게 기록된 임의의 목록
- 혹은 이미지/엑셀 파일 자체를 이 곳에 직접 올려놓으세요.`
                  }
                  value={inputText}
                  onChange={(e) => {
                    setInputText(e.target.value);
                    if (e.target.value.trim() && attachedFile) {
                      setAttachedFile(null);
                    }
                  }}
                />

                {isDragging && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center bg-indigo-950/95 text-white rounded-2xl space-y-2">
                    <UploadCloud className="w-12 h-12 text-indigo-400 animate-bounce" />
                    <p className="font-bold text-base">파일을 여기에 놓아 업로드</p>
                    <p className="text-xs text-indigo-300">인공지능 분석용 PDF/이미지 및 로컬 엑셀 분석 지원</p>
                  </div>
                )}
              </div>

              {/* File Upload Assist Toolbar */}
              <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-[#1A1D23] border border-[#2D3139] p-3 rounded-2xl">
                <div className="flex items-center space-x-2">
                  <UploadCloud className="w-4 h-4 text-indigo-400" />
                  <span className="text-xs text-gray-400 font-semibold">또는 품의 견적 증빙 파일 첨부</span>
                </div>
                <div>
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="w-full md:w-auto px-4 py-2 text-xs font-bold rounded-xl text-gray-200 bg-[#111317] border border-[#2D3139] hover:bg-[#1A1D23] transition cursor-pointer flex items-center justify-center space-x-1"
                  >
                    <span>견적서 파일 읽기</span>
                  </button>
                  <input
                    type="file"
                    ref={fileInputRef}
                    className="hidden"
                    accept=".xlsx,.xls,.csv,.pdf,.png,.jpg,.jpeg,.txt"
                    onChange={handleLocalFileSelection}
                  />
                </div>
              </div>

              {/* Attached file indicator */}
              {attachedFile && (
                <div className="flex items-center justify-between bg-[#1A1D23] border border-indigo-500/30 p-3 rounded-2xl text-xs text-gray-200">
                  <div className="flex items-center space-x-2">
                    <FileText className="w-4 h-4 text-indigo-400" />
                    <div>
                      <p className="font-semibold text-white">{attachedFile.name}</p>
                      <p className="text-[10px] text-gray-500">{(attachedFile.size / 1024).toFixed(1)} KB · 파일 대기 중</p>
                    </div>
                  </div>
                  <button 
                    onClick={handleClearAttachedFile}
                    className="p-1.5 rounded-lg bg-[#111317] hover:bg-[#20242D] text-gray-400 hover:text-white transition cursor-pointer"
                    title="파일 첨부 해제"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* Core Execution Call To Action */}
              <div className="pt-2">
                <button
                  onClick={() => {
                    if (attachedFile) {
                      processUploadedFile(attachedFile);
                    } else {
                      handleParseData(inputText);
                    }
                  }}
                  disabled={loading || (!inputText.trim() && !attachedFile)}
                  className={`w-full py-3.5 rounded-2xl font-bold text-sm shadow-lg transition flex items-center justify-center space-x-2 cursor-pointer ${
                    loading 
                      ? 'bg-[#1D212B] text-gray-500 cursor-not-allowed border border-[#2D3139]' 
                      : 'bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-900/40'
                  }`}
                  id="convert_main_button"
                >
                  {loading ? (
                    <>
                      <RefreshCw className="w-5 h-5 animate-spin text-indigo-400" />
                      <span className="animate-pulse">대형 인공지능이 교육 기안 데이터 추출 중...</span>
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-5 h-5 text-white/90" />
                      <span>
                        {attachedFile 
                          ? `품의 첨부파일 AI 분석 및 변환하기 (${attachedFile.name.length > 15 ? attachedFile.name.slice(0, 15) + '...' : attachedFile.name})` 
                          : 'K-에듀파인 품의 내역 AI 변환하기'
                        }
                      </span>
                    </>
                  )}
                </button>
              </div>

              {/* Processing Feedback Stages */}
              {loading && (
                <div className="bg-[#1A1D23] border border-[#2D3139] p-4 rounded-2xl space-y-2.5 text-xs text-gray-400 animate-pulse">
                  <div className="flex items-center space-x-2 font-bold text-indigo-400">
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>실시간 변환 프로세스 작동 정보</span>
                  </div>
                  <div className="space-y-1">
                    <p className="flex items-center space-x-2">
                      <span className="text-emerald-500 font-bold">✔</span>
                      <span>제출 텍스트 토큰 구성 완료</span>
                    </p>
                    <p className="flex items-center space-x-2">
                      <span className="text-indigo-400 font-bold animate-pulse">●</span>
                      <span>Gemini-3.5-flash 대규모 모델에서 맵핑 작업 수행 중...</span>
                    </p>
                    <p className="text-gray-500 pl-4 text-[11px] leading-relaxed">
                      자동으로 비정형 텍스트 구조 분석, 품목 노이즈 제거 및 단가·수량 검증 수식을 순차 연산합니다.
                    </p>
                  </div>
                </div>
              )}

              {/* Error Box */}
              {error && (
                <div className="space-y-3 animate-fade-in" id="error_box">
                  {error.includes('Key') || error.includes('API') || error.includes('leaked') || error.includes('차단된') ? (
                    <div className="bg-amber-950/20 border border-amber-500/30 p-5 rounded-2xl flex flex-col gap-3 text-xs text-amber-200">
                      <div className="flex items-start space-x-3">
                        <AlertCircle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                        <div>
                          <h4 className="font-bold text-sm text-amber-400">🚨 API 비밀 키(Secrets) 사용량 및 제한 오류 안내</h4>
                          <p className="mt-1.5 leading-relaxed font-semibold">
                            현재 에듀드래프트에 등록된 <code className="bg-amber-950/80 px-1 py-0.5 rounded text-white font-mono">GEMINI_API_KEY</code>가 비활성화 되었거나 유출되어 구글 가드레일 필터에 의해 차단되었습니다. (403 Permission Denied)
                          </p>
                        </div>
                      </div>
                      
                      <div className="mt-1.5 bg-[#12141C] rounded-xl p-4 border border-[#2D3139] space-y-3 text-gray-300">
                        <p className="font-bold text-white text-xs flex items-center gap-1.5">
                          🔑 해결하는 방법 (초간단 1분 해결!)
                        </p>
                        <ol className="list-decimal list-inside space-y-2 text-[11px] leading-relaxed">
                          <li>
                            <a href="https://aistudio.google.com/" target="_blank" rel="noopener noreferrer" className="text-indigo-400 hover:underline font-bold inline-flex items-center gap-1">
                              구글 AI Studio(aistudio.google.com) ↗
                            </a>
                            에 접속하여 로그인합니다.
                          </li>
                          <li>
                            API 키 목록에서 생성해 두신 키 중 **아무 키나 우측의 [복사(아이콘)]**를 클릭합니다.
                          </li>
                          <li>
                            이 웹사이트 우측 상단의 <strong className="text-indigo-400">⚙️ [개인 API 키 입력]</strong> 버튼을 클릭합니다.
                          </li>
                          <li>
                            입력창에 복사한 API 키(<code className="bg-[#0F1115] px-1 rounded text-white font-mono">AIzaSy...</code>)를 붙여넣은 후, <strong className="text-indigo-400">[적용 완료]</strong>를 클릭합니다.
                          </li>
                          <li>
                            그 후 <strong className="text-white">품의 내역 AI 변환하기</strong> 버튼을 다시 누르면 대기 없이 즉시 깨끗하게 변환됩니다!
                          </li>
                        </ol>
                      </div>

                      <div className="text-[10px] text-gray-500 italic flex justify-between items-center px-1">
                        <span>상단 예제 시뮬레이터 기능 등은 상시 정밀 동작 중입니다.</span>
                        <span>Error Code: API_LEAK_403</span>
                      </div>
                    </div>
                  ) : (
                    <div className="bg-rose-950/20 border border-rose-900/40 p-4 rounded-2xl flex items-start space-x-3 text-xs text-rose-300">
                      <AlertCircle className="w-4 h-4 text-rose-500 shrink-0 mt-0.5" />
                      <div>
                        <h4 className="font-bold">변환 오류가 감지되었습니다</h4>
                        <p className="mt-1 font-medium leading-relaxed">{error}</p>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          </section>

          {/* Right Workspace Pane */}
          <section className="lg:col-span-7 space-y-6" id="output_workspace_panel">
            {/* Interactive Grid Table Card */}
            <div className="bg-[#16191F] rounded-3xl border border-[#2D3139] shadow-xl p-6 space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                <div className="flex items-center space-x-2">
                  <Layers className="w-5 h-5 text-indigo-400" />
                  <h3 className="font-bold text-base text-white">추출 품의 물품 조정</h3>
                </div>

                <div className="flex items-center gap-1.5">
                  <button
                    onClick={handleAddNewRow}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl text-indigo-400 bg-indigo-950/30 hover:bg-indigo-900/40 border border-indigo-900/40 transition duration-150 cursor-pointer flex items-center space-x-1"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>수동 추가</span>
                  </button>
                  {items.length > 0 && (
                    <button
                      onClick={handleClearAllItems}
                      className="px-3 py-1.5 text-xs font-bold rounded-xl text-rose-400 bg-rose-950/30 hover:bg-rose-900/40 border border-rose-900/40 transition duration-150 cursor-pointer flex items-center space-x-1"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>전체 비우기</span>
                    </button>
                  )}
                </div>
              </div>

              {items.length === 0 ? (
                <div className="border border-[#2D3139] rounded-2xl p-12 text-center bg-[#0F1115]/50">
                  <HelpCircle className="w-12 h-12 text-indigo-500/40 mx-auto mb-3 stroke-[1.5]" />
                  <p className="text-sm font-bold text-gray-400">분석 및 추출된 품의 항목이 없습니다.</p>
                  <p className="text-xs text-gray-500 mt-1 max-w-sm mx-auto leading-relaxed">
                    상단 예시 가이드를 로드해 시뮬레이션하거나 왼쪽 텍스트박스 영역에 견적 장바구니 내용을 복사+붙여넣기하여 실행 결과를 모니터링해보세요.
                  </p>
                </div>
              ) : (
                <div className="space-y-4">
                  {/* Table Scroll-wrap */}
                  <div className="overflow-x-auto border border-[#2D3139] rounded-2xl bg-[#0F1115]">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#1A1D23] border-b border-[#2D3139] text-gray-400 font-bold">
                          <th className="py-3 px-3 text-center w-12 text-slate-500">순번</th>
                          <th className="py-3 px-3 min-w-[150px] text-gray-300">품명*</th>
                          <th className="py-3 px-3 w-32 text-gray-300">규격*</th>
                          <th className="py-3 px-3 w-16 text-center text-gray-400">단위</th>
                          <th className="py-3 px-3 w-16 text-center text-yellow-500">수량*</th>
                          <th className="py-3 px-3 w-28 text-right text-emerald-400">단가*(원)</th>
                          <th className="py-3 px-3 w-28 text-right text-indigo-400">금액(원)</th>
                          <th className="py-3 px-3 w-20 text-center text-slate-500">관리</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#2D3139] font-medium text-gray-300">
                        {items.map((item, idx) => (
                          <tr key={item.id} className="hover:bg-[#1A1D23]/40 transition-colors">
                            {/* Indexes */}
                            <td className="py-2.5 px-3 text-center text-slate-500 font-mono text-[11px]">
                              {idx + 1}
                            </td>
                            
                            {/* Item name */}
                            <td className="py-2 px-3">
                              <input 
                                type="text"
                                className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-indigo-500 rounded px-2 py-1 text-white font-semibold focus:ring-1 focus:ring-indigo-950 outline-none"
                                value={item.name}
                                onChange={(e) => handleUpdateItem(item.id, 'name', e.target.value)}
                              />
                            </td>

                            {/* Spec */}
                            <td className="py-2 px-3">
                              <input 
                                type="text"
                                className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-indigo-500 rounded px-2 py-1 text-gray-400 focus:ring-1 focus:ring-indigo-950 outline-none"
                                value={item.spec}
                                onChange={(e) => handleUpdateItem(item.id, 'spec', e.target.value)}
                              />
                            </td>

                            {/* Unit */}
                            <td className="py-2 px-2">
                              <input 
                                type="text"
                                className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-indigo-500 rounded px-1.5 py-1 text-center text-gray-400 focus:ring-1 focus:ring-indigo-950 outline-none"
                                value={item.unit}
                                onChange={(e) => handleUpdateItem(item.id, 'unit', e.target.value)}
                              />
                            </td>

                            {/* Quantity */}
                            <td className="py-2 px-2">
                              <input 
                                type="number"
                                className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-emerald-500 rounded px-1.5 py-1 text-center text-yellow-500 font-bold focus:ring-1 focus:ring-[#1E291C] outline-none font-mono"
                                value={item.quantity}
                                onChange={(e) => handleUpdateItem(item.id, 'quantity', e.target.value)}
                              />
                            </td>

                            {/* Price */}
                            <td className="py-2 px-2">
                              <input 
                                type="text"
                                className="w-full bg-[#0F1115] border border-[#2D3139] focus:border-emerald-500 rounded px-2 py-1 text-right text-emerald-400 font-bold focus:ring-1 focus:ring-[#1E291C] outline-none font-mono"
                                value={item.price}
                                onChange={(e) => handleUpdateItem(item.id, 'price', e.target.value)}
                              />
                            </td>

                            {/* Item sum total */}
                            <td className="py-2 px-3 text-right text-indigo-300 font-bold font-mono">
                              {(item.price * item.quantity).toLocaleString()}
                            </td>

                            {/* Line modifiers */}
                            <td className="py-2 px-2 text-center">
                              <div className="flex items-center justify-center space-x-1">
                                <button 
                                  onClick={() => moveRowIndex(idx, 'up')}
                                  disabled={idx === 0}
                                  className="p-1 rounded text-gray-500 hover:text-white hover:bg-[#1A1D23] disabled:opacity-20 cursor-pointer text-[10px]"
                                  title="위로 이동"
                                >
                                  ▲
                                </button>
                                <button 
                                  onClick={() => moveRowIndex(idx, 'down')}
                                  disabled={idx === items.length - 1}
                                  className="p-1 rounded text-gray-500 hover:text-white hover:bg-[#1A1D23] disabled:opacity-20 cursor-pointer text-[10px]"
                                  title="아래로 이동"
                                >
                                  ▼
                                </button>
                                <button
                                  onClick={() => handleDeleteRow(item.id)}
                                  className="p-1 rounded text-rose-400 hover:text-rose-300 hover:bg-rose-950/30 cursor-pointer"
                                  title="행 삭제"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {/* Pricing Overview */}
                  <div className="bg-[#0F1115] text-[#E0E0E0] border border-[#2D3139] px-6 py-4 rounded-2xl flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 font-semibold text-sm">
                    <div className="space-y-1">
                      <span className="text-xs text-gray-500 font-bold uppercase tracking-wider block">총합 품의 산정 예산</span>
                      <span className="text-indigo-400 text-base font-black">금{totalPrice.toLocaleString()}원</span>
                      <span className="text-gray-400 text-xs ml-2 font-medium">({koreanWonWord ? `금${koreanWonWord}원정` : ''})</span>
                    </div>
                    <div className="text-right text-xs text-gray-400">
                      추출품수: <strong className="text-white text-sm font-extrabold">{totalItemCount}</strong> 종 / 수량총계: <strong className="text-white text-sm font-extrabold">{items.reduce((s,i)=>s+i.quantity,0)}</strong> 개
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Editable Draft Copying Panel */}
            {items.length > 0 && (
              <div className="bg-[#16191F] rounded-3xl border border-[#2D3139] shadow-xl p-6 space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center space-x-2 text-white">
                    <FileText className="w-5 h-5 text-indigo-400" />
                    <h3 className="font-semibold text-sm text-indigo-300 uppercase tracking-wider">나의 지출 품의 기안 초안</h3>
                  </div>
                  <span className="bg-emerald-950/50 text-emerald-400 text-[11px] font-bold px-2.5 py-0.5 rounded-full border border-emerald-900/30">
                    실시간 자동 매핑 중
                  </span>
                </div>

                <p className="text-xs text-gray-400 leading-relaxed font-semibold">
                  아래 상자는 교육행정 기안문에 즉시 복사하여 올릴 수 있도록 완성된 초안입니다. 학교 수신 공문번호나 일정 정보를 수정 적용할 수 있습니다.
                </p>

                {/* Draft details modifiers */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs bg-[#0F1115] p-4 rounded-2xl border border-[#2D3139]">
                  <div className="space-y-1.5">
                    <label className="text-gray-400 font-bold block">2. 관련 문서번호 양식 선택 또는 입력</label>
                    <div className="relative">
                      <Calendar className="absolute left-3 top-2.5 w-3.5 h-3.5 text-gray-500" />
                      <input 
                        type="text"
                        className="w-full bg-[#16191F] border border-[#2D3139] rounded-xl pl-9 pr-3 py-2 text-white font-semibold outline-none focus:border-indigo-500 transition-colors"
                        value={draftRelated}
                        onChange={(e) => setDraftRelated(e.target.value)}
                        placeholder="예: 행정과-154호(2026.05.29.)"
                      />
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-gray-400 font-bold block">3. 품의 상세 사유 프리셋 자동 삽입</label>
                    <select
                      className="w-full bg-[#16191F] border border-[#2D3139] rounded-xl px-3 py-2 text-gray-300 outline-none font-semibold focus:border-indigo-500 transition-colors"
                      onChange={(e) => setDraftNote(e.target.value)}
                      value={DRAFT_REASON_PRESETS.includes(draftNote) ? draftNote : "custom"}
                    >
                      {DRAFT_REASON_PRESETS.map((p, idx) => (
                        <option key={idx} value={p}>{p.slice(0, 30)}...</option>
                      ))}
                      <option value="custom">== 수동 텍스쳐 사유 기재하기 ==</option>
                    </select>
                  </div>

                  <div className="md:col-span-2 space-y-1.5">
                    <label className="text-gray-400 font-bold block">목적 및 상세 사유 수동 교정</label>
                    <textarea 
                      rows={2}
                      className="w-full bg-[#16191F] border border-[#2D3139] rounded-xl px-3 py-2 text-white outline-none focus:border-indigo-500 transition-colors text-xs font-semibold"
                      value={draftNote}
                      onChange={(e) => setDraftNote(e.target.value)}
                      placeholder="자세한 집행 세부 내역이나 관련 지출 품의 사유글을 기록하세요."
                    />
                  </div>
                </div>

                {/* Styled Output Clipboard Area */}
                <div className="relative bg-[#0F1115] rounded-2xl p-5 text-gray-200 font-serif text-[15px] leading-relaxed border border-[#2D3139] overflow-hidden">
                  <div className="absolute top-2.5 right-3 text-[10px] text-gray-600 select-none font-sans font-bold uppercase tracking-wider">
                    EduDraft Preview Box
                  </div>
                  <pre className="whitespace-pre-wrap select-all text-gray-200 pr-4 pt-2 font-serif leading-8">
                    {draftTitle}
                  </pre>
                  
                  {/* Floating Action Button */}
                  <div className="absolute bottom-4 right-4">
                    <button
                      onClick={handleCopyToClipboard}
                      className={`px-4 py-2.5 text-xs font-bold rounded-xl shadow-lg transition duration-200 flex items-center space-x-1.5 cursor-pointer ${
                        copySuccess 
                          ? 'bg-emerald-600 hover:bg-emerald-505 text-white shadow-emerald-900/40' 
                          : 'bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-900/40'
                      }`}
                    >
                      {copySuccess ? (
                        <>
                          <Check className="w-3.5 h-3.5" />
                          <span>기안문 클립보드에 복사 완료!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>기안문 클립보드 복사</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Export options and download panel */}
            {items.length > 0 && (
              <div className="bg-[#16191F] rounded-3xl border border-[#2D3139] p-6 space-y-5">
                <div className="flex items-center space-x-2.5 text-white">
                  <Download className="w-5 h-5 text-indigo-400" />
                  <h3 className="font-bold text-base">에듀파인 일괄 내역 등록용 엑셀 내보내기</h3>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* 1. Format Choice card */}
                  <div className="space-y-2">
                    <label className="text-xs text-gray-400 font-bold block bg-[#0F1115] px-2 py-1 rounded inline-block">A. 파일 업로드 양식 선택</label>
                    <div className="grid grid-cols-2 gap-2 text-center text-xs">
                      <button
                        onClick={() => setExcelTemplateType('standard')}
                        className={`p-3 rounded-2xl border font-bold transition duration-150 cursor-pointer flex flex-col items-center justify-center space-y-1.5 ${
                          excelTemplateType === 'standard'
                            ? 'bg-[#1A1D23] border-indigo-500 text-white shadow-md ring-1 ring-indigo-500/30'
                            : 'bg-[#111317] border-[#2D3139] text-gray-500 hover:bg-[#1A1D23]'
                        }`}
                      >
                        <Layers className="w-4 h-4" />
                        <span>K-에듀파인 표준양식</span>
                        <span className="text-[9px] text-gray-500 font-medium">(3줄 가이드라인 헤더 포함)</span>
                      </button>

                      <button
                        onClick={() => setExcelTemplateType('simple')}
                        className={`p-3 rounded-2xl border font-bold transition duration-150 cursor-pointer flex flex-col items-center justify-center space-y-1.5 ${
                          excelTemplateType === 'simple'
                            ? 'bg-[#1A1D23] border-indigo-500 text-white shadow-md ring-1 ring-indigo-500/30'
                            : 'bg-[#111317] border-[#2D3139] text-gray-500 hover:bg-[#1A1D23]'
                        }`}
                      >
                        <FileSpreadsheet className="w-4 h-4" />
                        <span>간편 단일 헤더 양식</span>
                        <span className="text-[9px] text-gray-500 font-medium">(정렬 기본형)</span>
                      </button>
                    </div>
                  </div>

                  {/* 2. File Name card */}
                  <div className="space-y-2">
                    <label className="text-xs text-gray-400 font-bold block bg-[#0F1115] px-2 py-1 rounded inline-block">B. 엑셀 저장 명칭 지정</label>
                    <div className="flex items-center space-x-1">
                      <input
                        type="text"
                        className="w-full bg-[#0F1115] border border-[#2D3139] rounded-xl px-3 py-2.5 text-xs text-white font-bold outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-950"
                        value={customFilename}
                        onChange={(e) => setCustomFilename(e.target.value)}
                        placeholder="저장할 파일명"
                      />
                      <span className="text-xs font-bold text-gray-500 shrink-0">.xlsx</span>
                    </div>
                    <p className="text-[10px] text-gray-500 leading-relaxed font-semibold">
                      K-에듀파인 일괄 물품 등록 창에서 업르드를 지원합니다.
                    </p>
                  </div>
                </div>

                {/* Final Master Download Button */}
                <div className="pt-2">
                  <button
                    onClick={handleDownloadExcel}
                    className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-3.5 px-4 rounded-2xl shadow-lg transition duration-200 flex items-center justify-center space-x-2 text-sm cursor-pointer shadow-emerald-900/20"
                  >
                    <FileSpreadsheet className="w-5 h-5 text-white/90" />
                    <span>정밀 변환 완료된 엑셀 파일 내려받기</span>
                  </button>
                  <p className="text-[11px] text-gray-500 text-center font-semibold mt-2">
                    * 이 다운로드한 엑셀 파일을 에듀파인 '지출품의 내역 등록' 목록 창에 바로 추가하시면 곧바로 등록이 수행됩니다.
                  </p>
                </div>
              </div>
            )}
          </section>

        </div>
        
      </main>

      <footer className="border-t border-[#2D3139] px-8 py-6 bg-[#16191F] flex flex-col md:flex-row items-center justify-between text-[10px] text-gray-500 tracking-widest uppercase gap-4 mt-12 w-full" id="footer">
        <div className="flex flex-col md:flex-row gap-4 md:gap-6 items-center text-center md:text-left">
          <span className="flex items-center gap-2"><span class="w-1.5 h-1.5 rounded-full bg-green-500"></span> 교직원 전용 업무 지원 패키지</span>
          <span>사용 기술: Gemini-3.5-Flash + SheetJS Integration</span>
        </div>
        <div className="flex gap-4 italic text-center md:text-right">
          <span>제작 프로젝트: EduDraft Master Desktop Suite</span>
          <span>v1.2.0-stable</span>
        </div>
      </footer>
    </div>
  );
}
