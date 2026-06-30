import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';

// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3000;

// Body parser with higher limit for file uploads (e.g. screenshots)
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));

// Set up static files serving from /dist after Vite build
app.use(express.static(path.join(__dirname, 'dist')));

// Helper: safe model initialization
function getGeminiClient(customApiKey?: string) {
  const apiKey = customApiKey || process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('API 키가 비어 있습니다. 개인 API 키를 등록하거나, 서버의 GEMINI_API_KEY 환경변수를 설정해 주세요.');
  }
  return new GoogleGenAI({
    apiKey: apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

// Robust text generation with retry and transparent fallback models (flash-latest, flash-lite) to guard against any temporary 503 high demand spikes.
async function generateContentWithFallback(
  ai: ReturnType<typeof getGeminiClient>,
  options: {
    contents: any[];
    systemInstruction: string;
    responseSchema: any;
  }
) {
  // We prioritize the primary model, then transparently fallback to high-availability stable/preview flash models
  const modelsToTry = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-3.1-flash-lite'];
  let lastError: any = null;

  for (const modelName of modelsToTry) {
    const maxRetries = 2;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        console.log(`[Gemini Request] Model: ${modelName} | Attempt ${attempt}/${maxRetries}`);
        const response = await ai.models.generateContent({
          model: modelName,
          contents: options.contents,
          config: {
            systemInstruction: options.systemInstruction,
            responseMimeType: 'application/json',
            responseSchema: options.responseSchema,
            temperature: 0.1
          }
        });
        
        console.log(`[Gemini Success] Successfully completed request with model: ${modelName}`);
        return response;
      } catch (err: any) {
        lastError = err;
        const errMsg = err?.message || String(err);
        const errStatus = err?.status || '';
        
        console.warn(`[Gemini Warning] Model ${modelName} (Attempt ${attempt}) failed: ${errMsg} (Status: ${errStatus})`);

        // If it's a 403 (leaked, unauthorized, PERMISSION_DENIED) or other authentication/key errors,
        // do not retry as it represents a permanent configuration error with the key itself.
        if (
          errMsg.includes('leaked') || 
          errMsg.includes('leak') ||
          errMsg.includes('API key') || 
          errMsg.includes('403') || 
          errMsg.includes('unauthorized') || 
          errMsg.includes('PERMISSION_DENIED') ||
          errMsg.includes('API_KEY')
        ) {
          throw err;
        }

        // Wait before retrying (backoff)
        if (attempt < maxRetries) {
          const delay = attempt * 1000;
          await new Promise((resolve) => setTimeout(resolve, delay));
        }
      }
    }
  }

  // If we exhausted all options, throw the final cumulative error
  throw lastError;
}

// REST API for parsing shopping cart / estimate sheets using Gemini
app.post('/api/parse', async (req, res) => {
  try {
    const { text, fileData, userApiKey } = req.body;
    
    if (!text && !fileData) {
      return res.status(400).json({ error: '분석할 텍스트나 파일 데이터가 필요합니다.' });
    }

    const ai = getGeminiClient(userApiKey);

    const systemInstruction = `You are an expert procurement assistant for Korean school administration (K-Edufine/K-에듀파인). 
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
      type: Type.ARRAY,
      description: "Parsed list of items for K-Edufine",
      items: {
        type: Type.OBJECT,
        properties: {
          name: { type: Type.STRING, description: "품명 (Item Name / Description of the product)." },
          spec: { type: Type.STRING, description: "규격 (Specification, Model, Size, Volume, etc.). Format nicely." },
          unit: { type: Type.STRING, description: "단위 (Unit, e.g. 개, EA, box). If unclear, default to '개'." },
          quantity: { type: Type.INTEGER, description: "수량 (Quantity purchased). Must be a positive integer." },
          price: { type: Type.INTEGER, description: "단가 (Unit price of a single item in KRW, integer)." },
          remarks: { type: Type.STRING, description: "비고 (Remarks, notes, option details. Optional)." }
        },
        required: ["name", "spec", "unit", "quantity", "price"]
      }
    };

    const contents: any[] = [];

    // If file image/document is uploaded, send it as inline data part
    if (fileData && fileData.base64 && fileData.mimeType) {
      contents.push({
        inlineData: {
          data: fileData.base64,
          mimeType: fileData.mimeType
        }
      });
    }

    // Add prompt text
    const textPrompt = `Please analyze the following estimate / cart data and parse it into K-Edufine items.
User Pasted Data:
${text || "(Image/File provided only)"}

Format strictly as JSON.`;

    contents.push({ text: textPrompt });

    const response = await generateContentWithFallback(ai, {
      contents: contents,
      systemInstruction: systemInstruction,
      responseSchema: responseSchema
    });

    const parsedText = response.text || "[]";
    const parsedData = JSON.parse(parsedText);

    res.json({ success: true, items: parsedData });
  } catch (error: any) {
    console.error('Gemini processing error:', error);
    
    let errMsg = '';
    if (error && typeof error === 'object') {
      errMsg = error.message || JSON.stringify(error);
    } else {
      errMsg = String(error);
    }
    
    let friendlyMessage = errMsg;
    
    // Check for Gemini API key leak or permission issues
    if (
      errMsg.includes('leaked') || 
      errMsg.includes('leak') || 
      errMsg.includes('API key') || 
      errMsg.includes('403') || 
      errMsg.includes('unauthorized') || 
      errMsg.includes('PERMISSION_DENIED') ||
      errMsg.includes('API_KEY')
    ) {
      friendlyMessage = '차단된 API Key 감지: 본 앱에 설정된 Gemini API Key가 외부에 유출되거나 만료되어 구글 보안 검증 필터에 의해 비활성화(403 차단)되었습니다. 개발 서버 재구동 및 복구를 위해 화면 상단의 [Secrets] 탭이나 우측 상단 기어(Settings) -> [Secrets] 팝업창에서 새로운 API Key를 발급/교체해 주세요.';
    }
    
    res.status(500).json({
      success: false,
      error: friendlyMessage
    });
  }
});

// For any other request, send the built react index.html (SPA Fallback)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'));
});

// Start listening
app.listen(port, () => {
  console.log(`Edufine assistant server is listening on port ${port}`);
});
