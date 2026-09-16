import express from "express";
import path from "path";
import dotenv from "dotenv";

// Load environment variables
dotenv.config();

const app = express();
const PORT = 3000;

// Enable JSON parsing
app.use(express.json({ limit: "10mb" }));

// Helper to extract and deduplicate multiple API keys (supports newline, comma, or semicolon)
function parseApiKeys(rawInput?: string): string[] {
  const combined = [rawInput, process.env.GEMINI_API_KEY, process.env.KOBOILLM_API_KEY, process.env.LITELLM_API_KEY]
    .filter(Boolean)
    .join('\n');
  const keys = Array.from(new Set(
    combined
      .split(/[\n,;]+/)
      .map(k => k.trim())
      .filter(k => k.length > 5)
  ));
  return keys;
}

// Helper to normalize Base URL for Google AI Studio & OpenAI compatibility
function normalizeBaseUrl(url?: string): string {
  let u = (url || '').trim().replace(/\/+$/, '');
  if (!u || u === 'https://api.koboillm.com/v1') {
    return 'https://generativelanguage.googleapis.com/v1beta/openai';
  }
  if (u.includes('generativelanguage.googleapis.com')) {
    return 'https://generativelanguage.googleapis.com/v1beta/openai';
  }
  return u;
}

// Helper to format AI errors nicely for client responses
function formatServerAiError(error: any): { statusCode: number; message: string } {
  const errorString = typeof error === 'string' ? error : (error?.message || JSON.stringify(error) || '');
  const isQuota = 
    error?.status === 429 || 
    error?.statusCode === 429 || 
    error?.status === 503 ||
    error?.statusCode === 503 ||
    error?.code === 429 ||
    error?.code === 503 ||
    /quota|limit|429|exhausted|503|demand|unavailable/i.test(errorString);

  if (isQuota) {
    return {
      statusCode: 429,
      message: `⚠️ Kuota / Rate Limit API Google AI Studio / Gemini Telah Terlampaui (Error 429 Exceeded Quota).

💡 SOLUSI CARA MENGATASINYA:
1. Buka Tab 1 ('1. Input Parameter & Prompt') > 'Pengaturan Koneksi AI (Google AI Studio & Gemini)'.
2. Masukkan Kunci API Google AI Studio Anda sendiri (dapatkan gratis di https://aistudio.google.com/app/apikey).
3. Anda dapat memasukkan beberapa API Key dipisah koma/baris baru untuk Rotasi Otomatis.`
    };
  }

  const isInvalidModel = /invalid model|model_not_found|not found|view available models|invalid model name/i.test(errorString);
  if (isInvalidModel) {
    return {
      statusCode: 400,
      message: `⚠️ Nama Model AI Tidak Valid / Tidak Didukung oleh Server API.

💡 SOLUSI:
1. Buka Tab '1. Input Parameter & Prompt' > 'Pengaturan Koneksi AI (Google AI Studio & Gemini)'.
2. Klik tombol '🔄 Ambil Daftar Model dari Server' untuk memuat daftar model resmi yang didukung oleh API Key Anda.
3. Pilih model aktif (seperti gemini-2.0-flash) dari menu dropdown.`
    };
  }

  return {
    statusCode: 500,
    message: errorString || "Terjadi kesalahan pada layanan API Google AI Studio / Gemini."
  };
}

// Helper to sanitize markdown JSON codeblocks and extract pure JSON strings
function cleanJsonOutput(text: string): string {
  if (!text) return "[]";
  let cleaned = text.trim();

  // 1. Extract content inside markdown ```json ... ``` or ``` ... ```
  const codeblockMatch = cleaned.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (codeblockMatch && codeblockMatch[1] && codeblockMatch[1].trim().length > 0) {
    cleaned = codeblockMatch[1].trim();
  } else {
    cleaned = cleaned.replace(/^```[a-zA-Z]*\n?/, "").replace(/\n?```$/, "").trim();
  }

  // 2. Extract JSON Array [...] if embedded inside conversational text
  const arrayMatch = cleaned.match(/\[\s*\{[\s\S]*\}\s*\]/);
  if (arrayMatch) {
    return arrayMatch[0].trim();
  }

  // 3. Extract JSON Object {...} if embedded inside conversational text
  const objectMatch = cleaned.match(/\{[\s\S]*\}/);
  if (objectMatch) {
    return objectMatch[0].trim();
  }

  return cleaned;
}

// Helper to detect STEM / Exact Science subjects (Matematika, Fisika, Kimia, Biologi, dll.)
function isExactScienceSubject(subject?: string): boolean {
  if (!subject) return false;
  const s = subject.toLowerCase();
  return (
    s.includes('matematika') ||
    s.includes('fisika') ||
    s.includes('kimia') ||
    s.includes('biologi') ||
    s.includes('ipa') ||
    s.includes('sains') ||
    s.includes('ekonomi') ||
    s.includes('akuntansi') ||
    s.includes('statistik') ||
    s.includes('kalkulus') ||
    s.includes('aljabar') ||
    s.includes('geometri') ||
    s.includes('matriks') ||
    s.includes('matrix')
  );
}

// Helper to detect matrix or advanced math topics specifically
function isMatrixOrMathTopic(subject?: string, elemen?: string, subElemen?: string, kompetensi?: string): boolean {
  const combined = `${subject || ''} ${elemen || ''} ${subElemen || ''} ${kompetensi || ''}`.toLowerCase();
  return (
    combined.includes('matriks') ||
    combined.includes('matrix') ||
    combined.includes('ordo') ||
    combined.includes('determinan') ||
    combined.includes('invers') ||
    combined.includes('transpose') ||
    combined.includes('aljabar') ||
    combined.includes('vektor') ||
    combined.includes('spldv') ||
    combined.includes('spltv') ||
    combined.includes('persamaan linear') ||
    combined.includes('matematika')
  );
}

// Ultra-robust JSON parser with multi-stage auto-repair for LLM output (supporting LaTeX math, matrices and tables)
function safeParseJson(jsonString: string): any {
  const cleaned = cleanJsonOutput(jsonString);

  // Pre-process unescaped LaTeX backslashes for all known LaTeX commands
  const sanitizeLatex = (str: string) => {
    // 1. Escape known LaTeX commands starting with letters that conflict with JSON escapes
    let s = str.replace(/\\(frac|forall|flat|bar|beta|binom|bold|begin|bmatrix|pmatrix|vmatrix|Vmatrix|Bmatrix|bullet|nabla|neq|not|nu|rho|right|rightarrow|rightleftharpoons|rangle|times|theta|text|tau|to|tan|top|tilde|det|pmatrix|bmatrix|vmatrix|cdot|approx|leq|geq|pm|div|degree|alpha|gamma|delta|epsilon|lambda|mu|pi|sigma|omega|sum|int|sqrt|array)\b/gi, '\\\\$1');
    s = s.replace(/\\end\{(pmatrix|bmatrix|vmatrix|Vmatrix|Bmatrix|matrix|smallmatrix|array)\}/gi, '\\\\end{$1}');
    s = s.replace(/\\begin\{(pmatrix|bmatrix|vmatrix|Vmatrix|Bmatrix|matrix|smallmatrix|array)\}/gi, '\\\\begin{$1}');
    
    // 2. Escape any backslash NOT followed by valid JSON escape character (" \ / b f n r t or uXXXX)
    s = s.replace(/\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})/g, '\\\\');
    return s;
  };

  try {
    return JSON.parse(cleaned);
  } catch (err1) {
    // Attempt 1: Fix trailing commas and sanitize LaTeX
    let repaired = sanitizeLatex(cleaned).replace(/,\s*([}\]])/g, '$1');
    try {
      return JSON.parse(repaired);
    } catch (err2) {
      // Attempt 2: Auto-repair unescaped LaTeX backslashes inside JSON strings and handle raw newlines
      repaired = repaired.replace(/"((?:\\.|[^"\\])*)"/g, (match, inner) => {
        const fixedInner = inner
          .replace(/\\([a-zA-Z0-9_{}\^]+|[^\s"\\/bfnrtu])/g, '\\\\$1')
          .replace(/[\u0000-\u001F]+/g, (m: string) => {
            if (m === '\n') return '\\n';
            if (m === '\r') return '\\r';
            if (m === '\t') return '\\t';
            return '';
          });
        return `"${fixedInner}"`;
      });

      try {
        return JSON.parse(repaired);
      } catch (err3) {
        // Attempt 3: Extract individual valid JSON objects using regex
        const objectMatches = repaired.match(/\{[^{}]*\}/g);
        if (objectMatches && objectMatches.length > 0) {
          const items: any[] = [];
          for (const m of objectMatches) {
            try {
              items.push(JSON.parse(m));
            } catch {}
          }
          if (items.length > 0) return items;
        }
        throw err1; // Throw original error if repair fails completely
      }
    }
  }
}

// Helper to clean model names for Google AI Studio / Gemini API
function cleanModelName(m: string): string {
  return (m || '')
    .trim()
    .replace(/^(models\/|google\/|gemini\/|publishers\/google\/models\/)/i, '')
    .trim();
}

// Native Google Gemini REST API caller (fallback when OpenAI-style endpoint gives 404 or fails)
async function callNativeGeminiApiServer(
  apiKey: string,
  modelName: string,
  contents: string,
  systemInstruction?: string,
  temperature?: number,
  responseMimeType?: string
): Promise<string> {
  const cleanModel = cleanModelName(modelName) || "gemini-2.0-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${cleanModel}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const requestBody: any = {
    contents: [
      {
        role: "user",
        parts: [{ text: contents }]
      }
    ],
    generationConfig: {
      temperature: temperature ?? 0.7
    }
  };

  if (systemInstruction) {
    requestBody.systemInstruction = {
      parts: [{ text: systemInstruction }]
    };
  }

  if (responseMimeType === "application/json") {
    requestBody.generationConfig.responseMimeType = "application/json";
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 90000);

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(requestBody),
    signal: controller.signal
  });

  clearTimeout(timeoutId);

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Native Gemini API Error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    throw new Error("Respon kosong dari Native Gemini API.");
  }
  return text;
}

// Helper to dynamically query available models from LiteLLM / Gemini server
async function fetchServerModelsList(baseUrl: string, apiKey: string): Promise<string[]> {
  try {
    const res = await fetch(`${baseUrl}/models`, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "x-api-key": apiKey
      }
    });
    if (!res.ok) return [];
    const data = await res.json();
    let rawList: any[] = [];
    if (Array.isArray(data.data)) rawList = data.data;
    else if (Array.isArray(data)) rawList = data;
    else if (Array.isArray(data.models)) rawList = data.models;

    const isGoogle = baseUrl.includes("generativelanguage.googleapis.com");
    let mapped = rawList
      .map((item: any) => typeof item === 'string' ? item : item.id || item.name || String(item))
      .filter(Boolean);

    if (isGoogle) {
      mapped = mapped.map(cleanModelName);
    }
    return Array.from(new Set(mapped));
  } catch (e) {
    return [];
  }
}

// Google AI Studio / Gemini API Chat Completions Caller
async function generateContentWithLiteLLM(
  params: {
    contents: string;
    config?: any;
    apiKeysRaw?: string;
    temperature?: number;
    baseUrl?: string;
  }
): Promise<{ text: string; meta: { keyIndex: number; totalKeys: number; rotated: boolean } }> {
  const keysToTry = parseApiKeys(params.apiKeysRaw);
  if (keysToTry.length === 0) {
    throw new Error(
      "Kunci API Google AI Studio / Gemini tidak ditemukan. Silakan masukkan Kunci API pada Pengaturan Koneksi AI (Langkah 1)."
    );
  }

  const baseUrl = normalizeBaseUrl(params.baseUrl || params.config?.baseUrl || process.env.GEMINI_BASE_URL || process.env.KOBOILLM_BASE_URL || process.env.LITELLM_BASE_URL);
  const isGoogle = baseUrl.includes("generativelanguage.googleapis.com");

  const rawModel = params.config?.model || "gemini-2.0-flash";
  let candidateModels: string[] = [];
  if (isGoogle) {
    const primary = cleanModelName(rawModel) || "gemini-2.0-flash";
    candidateModels = [
      primary,
      "gemini-2.0-flash",
      "gemini-1.5-flash",
      "gemini-1.5-pro",
      "gemini-2.0-flash-lite",
      "gemini-1.5-flash-8b"
    ];
  } else {
    candidateModels = [
      rawModel,
      cleanModelName(rawModel),
      `google/${cleanModelName(rawModel)}`,
      `gemini/${cleanModelName(rawModel)}`,
      "gemini-2.0-flash",
      "gemini-1.5-flash"
    ];
  }
  const modelsToTry = Array.from(new Set(candidateModels.filter(Boolean)));

  let lastError: any = null;

  // Key Rotation Loop
  for (let keyIdx = 0; keyIdx < keysToTry.length; keyIdx++) {
    const apiKey = keysToTry[keyIdx];
    console.log(`[Google AI Studio API Rotation] Trying Key #${keyIdx + 1} of ${keysToTry.length} at ${baseUrl}/chat/completions...`);

    let serverModelsFetched = false;

    for (let i = 0; i < modelsToTry.length; i++) {
      const model = modelsToTry[i];
      try {
        const messages: any[] = [];
        if (params.config?.systemInstruction) {
          messages.push({ role: "system", content: params.config.systemInstruction });
        }
        messages.push({ role: "user", content: params.contents });

        const requestBody: any = {
          model,
          messages,
          temperature: params.temperature ?? params.config?.temperature ?? 0.7,
        };

        if (params.config?.responseMimeType === "application/json") {
          requestBody.response_format = { type: "json_object" };
        }

        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 90000);

        const response = await fetch(`${baseUrl}/chat/completions`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${apiKey}`,
            "x-api-key": apiKey
          },
          body: JSON.stringify(requestBody),
          signal: controller.signal
        });

        clearTimeout(timeoutId);

        if (!response.ok) {
          const errText = await response.text();
          console.warn(`Gemini model ${model} OpenAI request failed (status ${response.status}): ${errText}`);
          
          if (response.status === 429 || /quota|rate limit|429/i.test(errText)) {
            lastError = new Error(`API Gemini Error (${response.status}): ${errText}`);
            break; // Exceeded quota for this key, move to next key
          }

          // If on Google, try native Gemini REST endpoint fallback when status is 404 or model not found
          if (isGoogle && (response.status === 404 || /not found|404|invalid model/i.test(errText))) {
            console.log(`[Server Fallback] Trying Native Gemini API for key #${keyIdx + 1} model ${model}...`);
            try {
              const nativeText = await callNativeGeminiApiServer(
                apiKey,
                model,
                params.contents,
                params.config?.systemInstruction,
                params.temperature ?? params.config?.temperature,
                params.config?.responseMimeType
              );
              if (nativeText) {
                return {
                  text: nativeText,
                  meta: {
                    keyIndex: keyIdx,
                    totalKeys: keysToTry.length,
                    rotated: keyIdx > 0
                  }
                };
              }
            } catch (nativeErr: any) {
              console.warn(`Native Gemini API call failed for model ${model}:`, nativeErr?.message);
              lastError = nativeErr;
            }
          }

          // If model name is invalid or not found, dynamically discover models from server if not done yet
          if (!serverModelsFetched && /invalid model|model_not_found|not found|404|view available models/i.test(errText)) {
            serverModelsFetched = true;
            console.log(`[Auto-Discovery] Fetching available models directly from ${baseUrl}/models...`);
            const remoteModels = await fetchServerModelsList(baseUrl, apiKey);
            if (remoteModels.length > 0) {
              console.log(`[Auto-Discovery] Found ${remoteModels.length} models:`, remoteModels);
              for (const rm of remoteModels) {
                const cleanedRm = isGoogle ? cleanModelName(rm) : rm;
                if (!modelsToTry.includes(cleanedRm)) {
                  modelsToTry.push(cleanedRm);
                }
              }
            }
          }

          lastError = new Error(`API Gemini Error (${response.status}): ${errText}`);
          continue; // Try next candidate model
        }

        const data = await response.json();
        let contentText = data.choices?.[0]?.message?.content || "";

        if (!contentText) {
          lastError = new Error("Respon kosong dari API Google AI Studio / Gemini.");
          continue;
        }

        return {
          text: contentText,
          meta: {
            keyIndex: keyIdx,
            totalKeys: keysToTry.length,
            rotated: keyIdx > 0
          }
        };
      } catch (err: any) {
        lastError = err;
        console.warn(`Key #${keyIdx + 1} model ${model} failed:`, err?.message || err);
      }
    }
  }

  throw lastError || new Error("Gagal memproses AI setelah merotasi seluruh API Key dan model.");
}

// Helper to attach rotation header metadata to response
function attachRotationHeader(res: express.Response, meta?: { keyIndex: number; totalKeys: number; rotated: boolean }) {
  if (meta) {
    res.setHeader('x-api-key-index', String(meta.keyIndex));
    res.setHeader('x-api-keys-total', String(meta.totalKeys));
    res.setHeader('x-api-key-rotated', meta.rotated ? 'true' : 'false');
  }
}

// API Routes
const apiRouter = express.Router();

apiRouter.get("/health", (req, res) => {
  res.json({ 
    status: "ok", 
    provider: "google-ai-studio",
    baseUrl: process.env.GEMINI_BASE_URL || "https://generativelanguage.googleapis.com/v1beta/openai",
    apiKeyPresent: parseApiKeys().length > 0 
  });
});

// Endpoint: Fetch Available Models from Base URL
apiRouter.post("/fetch-models", async (req, res) => {
  try {
    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;
    const keys = parseApiKeys(apiKeysRaw);
    const baseUrl = normalizeBaseUrl(req.body.baseUrl || process.env.GEMINI_BASE_URL);

    const apiKey = keys[0] || '';

    const headers: Record<string, string> = {
      "Content-Type": "application/json"
    };
    if (apiKey) {
      headers["Authorization"] = `Bearer ${apiKey}`;
      headers["x-api-key"] = apiKey;
    }

    const response = await fetch(`${baseUrl}/models`, {
      method: "GET",
      headers
    });

    if (!response.ok) {
      const errText = await response.text();
      return res.status(response.status).json({ 
        error: `Gagal mengambil daftar model dari ${baseUrl}/models: ${errText}` 
      });
    }

    const data = await response.json();
    let modelsList: string[] = [];

    if (Array.isArray(data.data)) {
      modelsList = data.data
        .map((item: any) => typeof item === 'string' ? item : item.id || item.name || String(item))
        .filter(Boolean);
    } else if (Array.isArray(data)) {
      modelsList = data
        .map((item: any) => typeof item === 'string' ? item : item.id || item.name || String(item))
        .filter(Boolean);
    } else if (data.models && Array.isArray(data.models)) {
      modelsList = data.models
        .map((item: any) => typeof item === 'string' ? item : item.id || item.name || String(item))
        .filter(Boolean);
    }

    if (modelsList.length === 0) {
      modelsList = [
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-2.5-flash",
        "gemini-2.0-flash-lite"
      ];
    }

    res.json({ baseUrl, totalModels: modelsList.length, models: modelsList });
  } catch (error: any) {
    console.error("Error fetching models:", error);
    res.status(500).json({ error: error?.message || "Gagal mengambil daftar model dari API." });
  }
});

// Endpoint 1: Generate Kisi-Kisi (Matriks Asesmen) via AI LiteLLM
apiRouter.post("/generate-kisi", async (req, res) => {
  try {
    const {
      mataPelajaran,
      definisi,
      muatan,
      kompetensi,
      elemenMateri,
      subElemenMateri,
      count = 3,
      baseUrl,
      model
    } = req.body;

    if (!mataPelajaran) {
      return res.status(400).json({ error: "Mata Pelajaran harus diisi" });
    }

    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;

    const systemInstruction = `Anda adalah ahli kurikulum pendidikan menengah SMA di Indonesia (khususnya untuk penyusunan Tes Kemampuan Akademik / TKA). 
Tugas Anda adalah membuat rancangan KISI-KISI SOAL dalam bentuk MATRIKS ASESMEN sesuai dengan Kurikulum Merdeka atau K-13 tingkat SMA kelas X, XI, atau XII.
Rancanglah kisi-kisi soal yang berbobot, mengandung stimulus yang kuat, valid, dan seimbang berdasarkan input dari pengguna.
Hasilkan output berupa JSON Array murni yang valid tanpa teks pembungkus lain.`;

    const prompt = `Buatkan ${count} baris matriks asesmen kisi-kisi soal untuk mata pelajaran berikut:
Mata Pelajaran: ${mataPelajaran}
${definisi ? `Definisi/Tujuan: ${definisi}` : ""}
${muatan ? `Muatan Kurikulum: ${muatan}` : ""}
${kompetensi ? `Kompetensi Umum: ${kompetensi}` : ""}
${elemenMateri ? `Elemen/Materi Utama: ${elemenMateri}` : ""}
${subElemenMateri ? `Sub-Elemen/Submateri: ${subElemenMateri}` : ""}

Aturan Penyusunan Matriks:
1. Setiap baris harus bervariasi jenis bentuk soalnya: 'pilihan_ganda_sederhana' (PG Sederhana), 'mcma' (PG Kompleks Multiple Choice Multiple Answers), atau 'kategori' (PG Kompleks kategori Benar/Salah atau Sesuai/Tidak Sesuai).
2. Tingkat kognitif harus bervariasi antara: 'level_1' (Pemahaman / Knowing: Mengenali, mengingat, dan memahami konsep dasar), 'level_2' (Penerapan / Applying: Menerapkan konsep pada fenomena nyata), atau 'level_3' (Penalaran / Reasoning: Berpikir kritis dan menalar secara logis).
3. Buat rincian elemen, sub-elemen, kompetensi yang diukur, serta batasan materi secara logis dan mendalam.
4. Distribusikan jumlah soal per kisi-kisi (misalnya antara 3-10 soal per baris).
5. Hasilkan juga 'konteksNusantara' (rencana integrasi konteks lokal Nusantara/Indonesia yang spesifik dan relevan dengan materi ini, misal adat daerah, keragaman etnis, geografi kepulauan, sejarah lokal, dsb) serta 'stimulusTambahan' (rencana bentuk stimulus seperti teks bacaan, studi kasus riil, berita, data tabel, atau peristiwa konkret khas Indonesia) untuk meningkatkan kualitas stimulus soal.

Hasilkan format JSON Array objek dengan properti:
[
  {
    "bentukSoal": "pilihan_ganda_sederhana" | "mcma" | "kategori",
    "levelKognitif": "level_1" | "level_2" | "level_3",
    "elemenMateri": "string",
    "subElemenMateri": "string",
    "kompetensi": "string",
    "batasanCatatan": "string",
    "jumlahSoal": 5,
    "konteksNusantara": "string",
    "stimulusTambahan": "string"
  }
]`;

    const result = await generateContentWithLiteLLM({
      contents: prompt,
      apiKeysRaw,
      baseUrl,
      config: {
        model,
        systemInstruction,
        responseMimeType: "application/json"
      }
    });

    attachRotationHeader(res, result.meta);
    let parsed = safeParseJson(result.text);

    if (!Array.isArray(parsed) && parsed && typeof parsed === 'object') {
      if (Array.isArray(parsed.kisi)) parsed = parsed.kisi;
      else if (Array.isArray(parsed.data)) parsed = parsed.data;
      else if (Array.isArray(parsed.items)) parsed = parsed.items;
      else parsed = [parsed];
    }

    res.json(parsed);
  } catch (error: any) {
    console.error("Error generating kisi-kisi:", error);
    const formatted = formatServerAiError(error);
    res.status(formatted.statusCode).json({ error: formatted.message });
  }
});

// Endpoint 2: Generate Soal (Pembuat Soal) dari Kisi-Kisi via AI LiteLLM
apiRouter.post("/generate-soal", async (req, res) => {
  try {
    const {
      kisi,
      mataPelajaran,
      definisi,
      muatan,
      jumlahOpsi = 5,
      jenisSoal = "tunggal",
      konteksLokal = [],
      stimulusKonten = [],
      kualitasChecklist = [],
      noSoalStart = 1,
      existingQuestions = [],
      baseUrl,
      model
    } = req.body;

    if (!kisi) {
      return res.status(400).json({ error: "Data Kisi-Kisi wajib dilampirkan" });
    }

    const currentSubject = mataPelajaran || kisi.mataPelajaran;
    const isExact = isExactScienceSubject(currentSubject);
    const isMatrix = isMatrixOrMathTopic(currentSubject, kisi.elemenMateri, kisi.subElemenMateri, kisi.kompetensi);
    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;

    const systemInstruction = `Anda adalah ahli pembuat soal ujian nasional dan TKA (Tes Kemampuan Akademik) SMA di Indonesia${isExact ? " bidang Sains, Matematika, dan Eksakta" : ""}.
Anda sangat terampil menyusun soal tingkat tinggi (HOTS - Higher Order Thinking Skills), bervariasi, mendalam, dan bebas dari bias.
${isExact ? "Untuk mata pelajaran eksakta, tuliskan persamaan matematika, reaksi kimia, simbol ilmiah, dan satuan fisik secara presisi menggunakan format LaTeX ($ ... $) atau Unicode ilmiah baku. Jika menyajikan tabel stimulus data/eksperimen, buatlah dalam format tabel Markdown standar yang rapi." : ""}
${isMatrix ? "STANDAR KHUSUS MATRIKS MATEMATIKA: WAJIB menuliskan matriks menggunakan format LaTeX standar dengan pemisah kolom '&' dan pemisah baris '\\\\' yang presisi (contoh: $\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$, $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$, $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$). DILARANG menuliskan elemen matriks rapat tanpa pemisah kolom '&' dan baris '\\\\'. Opsi pilihan jawaban yang memuat matriks wajib dituliskan dengan notasi LaTeX matriks utuh." : ""}
Patuhi instruksi bentuk soal dan parameter kognitif yang ditentukan pengguna secara presisi.
Hasilkan output berupa JSON Array murni yang valid tanpa teks pembungkus markdown tambahan.`;

    const activeKonteksLokal = (kisi.konteksLokal && kisi.konteksLokal.length > 0) ? kisi.konteksLokal : konteksLokal;
    const activeStimulusKonten = (kisi.stimulusKonten && kisi.stimulusKonten.length > 0) ? kisi.stimulusKonten : stimulusKonten;
    const activeKualitasChecklist = (kisi.kualitasChecklist && kisi.kualitasChecklist.length > 0) ? kisi.kualitasChecklist : kualitasChecklist;

    const konteksStr = activeKonteksLokal.length > 0 
      ? `Integrasikan KONTEKS LOKAL INDONESIA berikut ke dalam stimulus atau soal: ${activeKonteksLokal.join(", ")}.`
      : "";

    const stimulusStr = activeStimulusKonten.length > 0
      ? `Gunakan tipe STIMULUS DAN PENGEMBANGAN KONTEN berikut: ${activeStimulusKonten.join(", ")} (misal teks bacaan, data/tabel, berita, kasus nyata).`
      : "Gunakan stimulus yang relevan jika sesuai dengan kompetensi.";

    const checklistStr = activeKualitasChecklist.length > 0
      ? `Pastikan memenuhi KUALITAS SOAL berikut: ${activeKualitasChecklist.join(", ")}.`
      : "";

    const bentukSoalDesc = 
      kisi.bentukSoal === "pilihan_ganda_sederhana"
        ? "Pilihan ganda sederhana: Hanya ada satu jawaban yang benar. Sediakan pilihan A sampai " + (jumlahOpsi === 5 ? "E" : "D") + "."
        : kisi.bentukSoal === "mcma"
        ? "Pilihan ganda kompleks model multiple choice multiple answers (MCMA): Ada lebih dari satu jawaban yang benar. Peserta diminta memilih semua jawaban benar. Kunci jawaban harus menyebutkan semua pilihan yang benar (misal: 'A, C'). Sediakan pilihan A sampai " + (jumlahOpsi === 5 ? "E" : "D") + "."
        : "Pilihan Ganda Kompleks Kategori (PGK Kategori): WAJIB mengikuti struktur khusus: (1) SOAL memuat [Stimulus/Narasi Kasus], [Pertanyaan Utama spesifik], [Kalimat Perintah Kategori: **Pilihlah Sesuai atau Tidak Sesuai pada setiap pernyataan!**], dan [Tabel Pernyataan Markdown 4 kolom: | # | Pernyataan | Sesuai | Tidak Sesuai | dengan label angka 1, 2, 3, 4 (minimal 2, maksimal 4 pernyataan)], (2) Pernyataan 1 s.d 4 di array opsi, (3) Kunci Jawaban (misal Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, Pernyataan 3: Sesuai, Pernyataan 4: Sesuai atau 1. Sesuai, 2. Tidak Sesuai, 3. Sesuai, 4. Sesuai), (4) Pembahasan yang merinci poin nomor 1 s.d 4 satu per satu.";

    const countRequired = Number(req.body.count) || Number(kisi.jumlahSoal) || 1;

    let existingQuestionsConstraint = '';
    if (Array.isArray(existingQuestions) && existingQuestions.length > 0) {
      const sanitizedList = existingQuestions
        .filter(item => typeof item === 'string' && item.trim().length > 0)
        .slice(0, 30);
      if (sanitizedList.length > 0) {
        existingQuestionsConstraint = `

HINDARI PENGULANGAN SOAL (SANGAT PENTING):
Jangan membuat soal yang sama atau mirip dengan soal-soal berikut:
${sanitizedList.map((text: string, idx: number) => `- Soal ${idx + 1}: ${text.substring(0, 150)}...`).join('\n')}
Pastikan butir soal yang Anda hasilkan saat ini benar-benar segar, baru, unik secara naratif, bervariasi, dan tidak mengulangi pertanyaan di atas.`;
      }
    }

    const exactGuidelinesStr = (isExact || isMatrix) ? `
PANDUAN KHUSUS MATA PELAJARAN EKSAKTA / SAINS & MATEMATIKA:
- Tuliskan semua persamaan matematika, reaksi kimia, dan rumus fisika menggunakan notasi LaTeX ($ ... $) atau simbol Unicode baku (misal: $x^2 + 5x + 6 = 0$, $\\frac{a}{b}$, H₂SO₄, Fe³⁺, $v = \\frac{s}{t}$, dll.).
${isMatrix ? `- STANDAR KHUSUS SOAL MATRIKS: Seluruh matriks WAJIB disajikan dengan notasi LaTeX matriks (seperti $\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$, $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$, atau $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$) dengan pemisah kolom '&' dan pemisah baris '\\\\' yang presisi. DILARANG menuliskan elemen matriks secara rapat tanpa pemisah kolom dan baris.` : ""}
- Sertakan satuan fisika/kimia yang tepat dan konsisten (misal: m/s², kg, Joule, Watt, Hz, °C, Ω, dll.).
- Jika menyajikan tabel stimulus eksperimen/data numerik, gunakan format Markdown Table 1 baris per record dengan pemisah (|---|---|) yang valid.
- Pilihan ganda opsi A s.d E harus homogen dalam format angka, desimal, pecahan, atau satuan.` : "";

    const prompt = `Buatkan tepat sebanyak ${countRequired} butir soal ujian TKA SMA yang berbeda untuk Mata Pelajaran ${mataPelajaran}.
    
PENTING: Jumlah objek soal yang dihasilkan dalam array JSON HARUS tepat sebanyak ${countRequired} butir soal.

INFORMASI MATRIKS ASESMEN KISI-KISI:
- No Soal Mulai: ${noSoalStart}
- Bentuk Soal: ${kisi.bentukSoal} (${bentukSoalDesc})
- Tingkat Kognitif: ${kisi.levelKognitif}
- Elemen/Materi: ${kisi.elemenMateri}
- Sub-Elemen/Submateri: ${kisi.subElemenMateri}
- Kompetensi yang Diuji: ${kisi.kompetensi}
- Batasan/Catatan Khusus: ${kisi.batasanCatatan || "Tidak ada"}
- Konteks Nusantara: ${kisi.konteksNusantara || "Tidak ada khusus"}
- Stimulus Tambahan: ${kisi.stimulusTambahan || "Tidak ada khusus"}
- Jenis Soal: ${jenisSoal}
${existingQuestionsConstraint}

PANDUAN UTAMA:
1. ${konteksStr}
2. ${stimulusStr}
3. ${checklistStr}${exactGuidelinesStr}
4. Kunci jawaban harus sangat akurat dan pembahasan harus lengkap, ilmiah, edukatif, dan terstruktur dengan rapi.
5. SANGAT PENTING (MANDATORI): Gabungkan paragraf stimulus/pengantar/studi kasus (bila ada) langsung ke bagian awal field 'soal' (diikuti pertanyaan utama di bawahnya), dan kosongkan field 'stimulus' (isi dengan string kosong ""). Jangan memisahkannya.
6. SANGAT PENTING: JANGAN mencantumkan nomor soal (seperti '1.', 'Soal 1.') di dalam teks field 'soal'.
7. SANGAT PENTING: Untuk PGK Kategori, teks 'soal' WAJIB menggabungkan 4 komponen berurutan: (a) Stimulus/narasi/poin-poin, (b) Pertanyaan Utama spesifik, (c) Kalimat Perintah Kategori, dan (d) Tabel Pernyataan Markdown 4 kolom (| # | Pernyataan | Sesuai | Tidak Sesuai |) dengan nomor urut angka 1, 2, 3, 4 di kolom '#' (minimal 2, maksimal 4 pernyataan). Array 'opsi' memuat daftar pernyataan berlabel nomor ['1. ...', '2. ...', '3. ...', '4. ...'], 'kunciJawaban' merinci status per nomor (contoh: "Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, Pernyataan 3: Sesuai, Pernyataan 4: Sesuai" atau "1. Sesuai, 2. Tidak Sesuai, 3. Sesuai, 4. Sesuai"), dan 'pembahasan' merinci penjelasan ilmiah poin nomor 1 s.d 4 satu per satu.

Hasilkan format JSON Array:
[
  {
    "kompetensi": "${kisi.kompetensi}",
    "subKompetensi": "${kisi.subElemenMateri}",
    "bentukSoal": "${kisi.bentukSoal}",
    "stimulus": "",
    "soal": "Teks stimulus/pengantar dan pertanyaan utama...",
    "opsi": ["A. Opsi A", "B. Opsi B", "C. Opsi C", "D. Opsi D", "E. Opsi E"],
    "kunciJawaban": "A",
    "pembahasan": "Penjelasan...",
    "kataKunci": "Topik Utama",
    "gambarUrl": ""
  }
]`;

    const result = await generateContentWithLiteLLM({
      contents: prompt,
      apiKeysRaw,
      baseUrl,
      config: {
        model,
        systemInstruction,
        responseMimeType: "application/json"
      }
    });

    attachRotationHeader(res, result.meta);
    let parsed = safeParseJson(result.text);

    if (!Array.isArray(parsed) && parsed && typeof parsed === 'object') {
      if (Array.isArray(parsed.soal)) parsed = parsed.soal;
      else if (Array.isArray(parsed.questions)) parsed = parsed.questions;
      else if (Array.isArray(parsed.data)) parsed = parsed.data;
      else if (Array.isArray(parsed.items)) parsed = parsed.items;
      else parsed = [parsed];
    }

    if (Array.isArray(parsed)) {
      parsed = parsed.map((q: any) => {
        let cleanSoal = typeof q.soal === 'string' ? q.soal.trim() : (typeof q.question === 'string' ? q.question.trim() : '');
        cleanSoal = cleanSoal.replace(/^(soal\s*)?(no\.?\s*)?\d+[\.\)\:\-]\s*/i, '').trim();

        let cleanOpsi = Array.isArray(q.opsi) ? q.opsi : (Array.isArray(q.options) ? q.options : []);
        cleanOpsi = cleanOpsi.map((opt: string, idx: number) => {
          let text = String(opt || '').trim();
          text = text.replace(/^\*\*([^*]+)\*\*/, '$1').trim();
          text = text.replace(/^[\*\-\•\s]*\(?([A-Ea-e1-5])\)?[\.\)\:\-]\s*/, '').trim();
          text = text.replace(/^[A-Ea-e]\s+/, '').trim();
          const letter = String.fromCharCode(65 + idx);
          return `${letter}. ${text}`;
        });

        return {
          ...q,
          soal: cleanSoal,
          opsi: cleanOpsi
        };
      });
    }

    res.json(parsed);
  } catch (error: any) {
    console.error("Error generating soal:", error);
    const formatted = formatServerAiError(error);
    res.status(formatted.statusCode).json({ error: formatted.message });
  }
});

// Endpoint 3: Generate Custom SVG Illustration/Graphic via AI LiteLLM
apiRouter.post("/generate-illustration", async (req, res) => {
  try {
    const { prompt, context, baseUrl, model } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: "Permintaan ilustrasi (prompt) harus diisi." });
    }

    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;
    const systemInstruction = `Anda adalah desainer grafis dan ahli ilustrasi ilmiah/edukatif profesional untuk soal ujian SMA.
Tugas Anda adalah menghasilkan kode inline SVG (<svg> ... </svg>) yang valid, indah, bersih, modern, dan sangat responsif untuk mendukung pemahaman soal ujian.`;

    const userPrompt = `Buatlah kode SVG inline yang merepresentasikan ilustrasi berikut:
Permintaan Pengguna: "${prompt}"
${context ? `Konteks Soal yang berkaitan: "${context}"` : ""}

Ingat, hanya hasilkan kode SVG langsung tanpa penanda kode atau pembungkus markdown apapun. Dimulai dari '<svg' sampai '</svg>'.`;

    const result = await generateContentWithLiteLLM({
      contents: userPrompt,
      apiKeysRaw,
      baseUrl,
      config: {
        model,
        systemInstruction,
        temperature: 0.2,
      },
    });

    attachRotationHeader(res, result.meta);
    let svgCode = cleanJsonOutput(result.text);
    
    res.json({ svg: svgCode });
  } catch (error: any) {
    console.error("Error generating illustration:", error);
    const formatted = formatServerAiError(error);
    res.status(formatted.statusCode).json({ error: formatted.message });
  }
});

// Endpoint 4: Test Health and Validity of Multiple API Keys
apiRouter.post("/test-key-health", async (req, res) => {
  try {
    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;
    const keys = parseApiKeys(apiKeysRaw);
    const baseUrl = normalizeBaseUrl(req.body.baseUrl || process.env.GEMINI_BASE_URL);

    if (keys.length === 0) {
      return res.status(400).json({ error: "Tidak ada API Key yang diberikan." });
    }

    const results = [];
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      const snippet = key.substring(0, 7) + "..." + key.substring(key.length - 4);
      try {
        const response = await fetch(`${baseUrl}/models`, {
          method: "GET",
          headers: {
            "Authorization": `Bearer ${key}`,
            "x-api-key": key
          }
        });

        if (response.ok) {
          results.push({ keyIndex: i, snippet, status: "valid", message: "API Key Aktif & Normal" });
        } else {
          const errText = await response.text();
          const isQuota = response.status === 429 || /quota|limit|429/i.test(errText);
          results.push({ 
            keyIndex: i, 
            snippet, 
            status: isQuota ? "exhausted" : "invalid", 
            message: isQuota ? "Limit Kuota Terlampaui (429)" : `Error HTTP ${response.status}`
          });
        }
      } catch (err: any) {
        results.push({ 
          keyIndex: i, 
          snippet, 
          status: "invalid", 
          message: err?.message || "Gagal terhubung"
        });
      }
    }

    res.json({ totalKeys: keys.length, results });
  } catch (error: any) {
    res.status(500).json({ error: error?.message || "Gagal menguji kesehatan API Key" });
  }
});

// Endpoint 5: Optimize/Generate Professional AI Prompt for a Kisi-Kisi Row
apiRouter.post("/optimize-prompt", async (req, res) => {
  try {
    const { kisi, mataPelajaran, baseUrl, model } = req.body;
    if (!kisi) {
      return res.status(400).json({ error: "Data kisi-kisi harus disediakan." });
    }

    const currentMapel = mataPelajaran || kisi.mataPelajaran || "Mata Pelajaran Umum";
    const isExact = isExactScienceSubject(currentMapel);
    const isMatrix = isMatrixOrMathTopic(currentMapel, kisi.elemenMateri || kisi.materi, kisi.subElemenMateri || kisi.subMateri, kisi.kompetensi);
    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;
    
    const systemInstruction = `Anda adalah ahli Rekayasa Prompt (Prompt Engineer) profesional dan spesialis Kurikulum & Evaluasi Pendidikan Indonesia${isExact || isMatrix ? " (Spesialis Sains, Teknologi, Teknik & Matematika / STEM / Eksakta)" : ""}.
Tugas Anda adalah merumuskan Prompt AI yang sangat detail, spesifik, dan efektif (Megaprompt) agar guru atau akademisi dapat menyalin prompt tersebut ke LLM lain (seperti Gemini, ChatGPT, Claude) untuk menghasilkan butir soal HOTS yang luar biasa.

Buat prompt dalam bahasa Indonesia yang berwibawa, rapi, terstruktur menggunakan format markdown (gunakan list, tebal, kode blok untuk visualisasi jika perlu). Prompt tersebut harus menginstruksikan AI eksternal untuk membuat soal berkualitas tinggi sesuai dengan kisi-kisi yang dikirimkan.`;

    const exactPromptSection = (isExact || isMatrix) ? `
10. STANDAR PENULISAN RUMUS, SIMBOL & TABEL EKSAKTA (SANGAT PENTING):
   - **Format Persamaan & Rumus**: WAJIB menginstruksikan penggunaan notasi LaTeX standar ($ ... $ untuk inline equation dan $$ ... $$ untuk display math) ATAU format simbol Unicode ilmiah yang jelas. Pecahan wajib ditulis \\frac{a}{b} atau (a/b), akar \\sqrt{x}, dan pangkat x^2.
   - **Simbol Ilmiah & Satuan**: Wajib mencantumkan simbol dan besaran/satuan fisika/kimia/matematika baku (seperti m/s², kg, J, W, Hz, Ω, μF, °C, K, ±, ×, ÷, ≈, ≤, ≥, ≠, ∞, Δ, α, β, θ, λ, μ, π, ρ, σ, ω, Ω).
   - **Formula Kimia & Reaksi**: Tuliskan formula kimia dengan subskrip/supersprip yang tepat (contoh: H₂SO₄ atau \\text{H}_2\\text{SO}_4, Fe³⁺, SO₄²⁻), tanda panah reaksi (\\rightarrow atau →, \\rightleftharpoons), dan fasa zat ((s), (l), (g), (aq)).
${isMatrix ? `   - **STANDAR KHUSUS SOAL MATRIKS (MANDATORI UNTUK AI & LLM)**:
     * **Notasi Kolom & Baris Matriks LaTeX**: WAJIB menyajikan matriks dengan notasi LaTeX standar dengan pemisah kolom '&' dan pemisah baris '\\\\' (contoh: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$, $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$, $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$).
     * **Presisi Kolom**: DILARANG menuliskan elemen matriks secara rapat tanpa pemisah kolom '&' dan pemisah baris '\\\\'.
     * **Format Pilihan Jawaban**: Setiap pilihan jawaban yang memuat matriks WAJIB disajikan lengkap dengan format LaTeX matriks yang utuh.
     * **Keamanan Kolom Excel**: Pada tabel Markdown/Excel, gunakan format kurung bulat/siku LaTeX tanpa tanda pipa '|' tunggal di dalam sel determinan agar tidak merusak kolom tabel.` : ""}
   - **Format Tabel Data / Eksperimen / Stimulus Numerik**: WAJIB menyajikan tabel dalam format Markdown Table standar yang utuh 1 baris per record dengan header pemisah (|---|---|) yang presisi. DILARANG memecah 1 baris sel data menjadi beberapa baris (multiline) di dalam sel tabel.
   - **Homogenitas Opsi Pilihan Ganda**: Opsi A s.d E wajib homogen dalam format angka, desimal, pecahan, atau penyertaan satuan.` : "";

    const userPrompt = `Buatlah draf PROMPT AI (Megaprompt) yang siap disalin oleh guru. Prompt tersebut harus dioptimalkan untuk menghasilkan soal ujian yang sangat spesifik berdasarkan data matriks berikut:
- Mata Pelajaran: ${currentMapel}
- No Kisi-Kisi: ${kisi.no}
- Kompetensi Dasar / Lingkup: ${kisi.kompetensi}
- Materi Pokok: ${kisi.elemenMateri || kisi.materi || ""}
- Sub-materi / Indikator: ${kisi.subElemenMateri || kisi.subMateri || "-"}
- Level Kognitif: ${kisi.levelKognitif}
- Bentuk Soal: ${kisi.bentukSoal}
- Jumlah Soal yang Diminta: ${kisi.jumlahSoal} butir soal
${kisi.konteksNusantara ? `- Rencana Konteks Nusantara: ${kisi.konteksNusantara}` : ""}
${kisi.stimulusTambahan ? `- Rencana Stimulus Tambahan: ${kisi.stimulusTambahan}` : ""}

Draf Megaprompt yang Anda buat harus memuat:
1. Peran AI yang diinstruksikan (misal: "Anda adalah dosen/guru senior pembuat soal TKA SMA ${currentMapel}...").
2. Spesifikasi lengkap materi, tingkat kognitif (C1-C6/HOTS), serta integrasi Konteks Nusantara dan Stimulus Tambahan yang spesifik agar bernuansa ke-Indonesia-an yang otentik dan mendalam.
3. ATURAN PENULISAN SOAL MANDATORI: Wajib menginstruksikan AI eksternal untuk MENGGABUNGKAN stimulus (paragraf/data/tabel/studi kasus) dan pertanyaan utama secara langsung menyatu di dalam satu bagian 'Soal:', tanpa dipisah menjadi section/field stimulus tersendiri, serta tanpa nomor soal di dalam teks.
4. PENGEMBANGAN STIMULUS VISUAL: WAJIB menginstruksikan penggunaan STIMULUS VISUAL ATAU BERBASIS DATA KONKRET (seperti Tabel Data/Matriks Markdown, Diagram Alur/Flowchart ASCII, Grafik Tren Tekstual, Spesifikasi Infografis/Peta, atau Kutipan Dokumen/Dialog Berformat).
5. Aturan pengecoh pilihan ganda / daftar pernyataan yang homogen, diawali huruf A, B, C, D, E dan titik, serta tidak mengandung tanda asteris (*).
6. Format keluaran Word & Excel: Draf Megaprompt WAJIB MELETAKKAN BANNER INSTRUCTION DI BAGIAN PALING ATAS PROMPT: "⚡ PERSYARATAN MUTLAK KELUARAN AI (2 FORMAT FILE UTUH SIAP UNDUH / DIPAKAI): Anda WAJIB dan MUTLAK menyajikan SELURUH HASIL GENERASI SOAL dalam DUA FORMAT OUTPUT UTUH LENGKAP BERURUTAN DALAM SATU BALASAN: 1️⃣ [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY] dan 2️⃣ [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]. DILARANG KERAS HANYA MEMBUAT NASKAH WORD ATAU HANYA TABEL EXCEL!" dan Wajib menginstruksikan untuk menggabungkan stimulus dan pertanyaan utama secara langsung dalam bagian 'Soal'.
7. KHUSUS PILIHAN GANDA KOMPLEKS KATEGORI (PGK KATEGORI): WAJIB menginstruksikan AI eksternal untuk meminta peserta memberikan respon pada masing-masing pernyataan (opsi kategori fleksibel: Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Ya/Tidak, Fakta/Opini, Setuju/Tidak Setuju, dll.). Menyajikan naskah soal (Bagian 1) dengan struktur khusus: (a) SOAL memuat [Stimulus/Narasi Kasus], [Pertanyaan Utama Spesifik], dan [Kalimat Perintah Kategori], (b) TABEL PERNYATAAN berbentuk Tabel Markdown 4 Kolom (| # | Pernyataan | [Kategori A] | [Kategori B] |) dengan nomor urut angka 1, 2, 3, 4 di kolom '#' (minimal 2, maksimal 4 pernyataan), (c) KUNCI JAWABAN merinci status respon per nomor, (d) PEMBAHASAN merinci penjelasan ilmiah untuk poin nomor 1 s.d 4 satu per satu, dan (e) TABEL EXCEL (Bagian 2) diisi dengan kolom Pernyataan_1, Pernyataan_2, Pernyataan_3, Pernyataan_4 memuat teks pernyataan 1 s.d 4 secara terpisah. DILARANG KERAS menggabungkan seluruh daftar pernyataan ke dalam satu sel Excel!
8. Aturan Format Nama File Output: Wajib menginstruksikan AI eksternal untuk menyajikan Judul / Nama File Dokumentasi Baku: [Mata Pelajaran]_[Materi Pokok]_[Bentuk Soal] (misal: ${currentMapel.replace(/\s+/g, '_')}_${(kisi.elemenMateri || 'Materi').replace(/\s+/g, '_')}_MCMA).
9. Teknik melahirkan pertanyaan tingkat tinggi (HOTS) yang memicu daya analisis siswa.${exactPromptSection}

Tulis draf prompt tersebut langsung dalam format Markdown yang elegan, berwibawa, rapi, dan langsung bisa dicopy oleh pengguna. Jangan tambahkan penjelasan pembuka dari Anda sendiri seperti "Berikut adalah prompt yang Anda minta", melainkan langsung mulailah isi prompt tersebut dengan judul atau teks instruksi utama yang siap disalin.`;

    const { text, meta } = await generateContentWithLiteLLM({
      contents: userPrompt,
      apiKeysRaw,
      baseUrl,
      config: {
        model,
        systemInstruction,
        temperature: 0.7,
      },
    });

    attachRotationHeader(res, meta);
    res.json({ prompt: text });
  } catch (error: any) {
    console.error("Error optimizing prompt:", error);
    const formatted = formatServerAiError(error);
    res.status(formatted.statusCode).json({ error: formatted.message });
  }
});

// Endpoint 6: Generate Systematic Learning Material from Kisi-Kisi Row
apiRouter.post("/generate-materi", async (req, res) => {
  try {
    const { kisi, mataPelajaran, mode, baseUrl, model } = req.body;
    if (!kisi) {
      return res.status(400).json({ error: "Data kisi-kisi harus disediakan." });
    }

    const currentMapel = mataPelajaran || kisi.mataPelajaran || "Mata Pelajaran";
    const isExact = isExactScienceSubject(currentMapel);
    const apiKeysRaw = (req.headers['x-api-key'] as string) || req.body.apiKey || undefined;
    
    let systemInstruction = "";
    let userPrompt = "";

    if (mode === "materi") {
      systemInstruction = `Anda adalah ahli kurikulum ${currentMapel} dan pengajar senior kelas dunia yang sangat karismatik dan inspiratif.
Tugas Anda adalah menyusun RINGKASAN MATERI AJAR yang sangat detail, akademis, komprehensif, mendalam, dan sistematis berdasarkan parameter kisi-kisi matriks yang diberikan.

PEDOMAN GAYA PENULISAN & FORMATTING TERIKAT:
1. Tulis dengan gaya bahasa yang sangat MENARIK BACA, KOMUNIKATIF, INSPIRATIF, DAN MENGGUGAH MINAT BACA SISWA MAUPUN GURU.
2. ATURAN BOLD DAN ITALIC (CETAK TEBAL DAN MIRING):
   - WAJIB MENGGUNAKAN CETAK MIRING (Markdown *istilah*) KHUSUS UNTUK SEMUA ISTILAH BAHASA ASING, BAHASA LATIN, ATAU BAHASA INGGRIS.
   - WAJIB MENGGUNAKAN CETAK TEBAL (Markdown **Konsep**) KHUSUS UNTUK KATA ATAU KONSEP KHUSUS, NAMA TEORI, RUMUS KUNCI, NAMA TOKOH/AHLI ILMUWAN.
${isExact 
  ? `3. FORMAT PERSAMAAN, RUMUS & TABEL EKSAKTA:
   - Tuliskan rumus dan persamaan matematika/fisika/kimia dengan format notasi LaTeX ($ ... $ atau $$ ... $$) atau simbol Unicode ilmiah yang rapi.
   - Sajikan tabel perbandingan, ringkasan rumus, atau data eksperimen dalam format Markdown Table standar yang rapi jika diperlukan.`
  : `3. DILARANG KERAS MENGGUNAKAN TABEL. Sajikan semua perbandingan dalam bentuk narasi paragraf yang rapi dan daftar poin.`
}
4. DILARANG MENGGUNAKAN TANDA BINTANG / ASTERISK ('*') SECARA ACAK ATAU BERANTAKAN.

Materi ajar ini wajib terdiri dari 4 bagian utama yang diberi judul Markdown (#):
# 1. PENDAHULUAN & DEFINISI
# 2. KONSEP UTAMA, RUMUS & TEORI PENDEKATAN
# 3. STUDI KASUS KONKRIT & CONTOH PENERAPAN
# 4. ANALISIS KRITIS & REFLEKSI`;

      userPrompt = `Buatkan RINGKASAN MATERI AJAR yang komprehensif, inspiratif, dan sangat menarik untuk dibaca untuk tingkat SMA berdasarkan unit berikut:
Mata Pelajaran: ${currentMapel}
Topik / Elemen Materi: ${kisi.elemenMateri}
Sub-elemen / Sub-materi: ${kisi.subElemenMateri}
Target Kompetensi Siswa: ${kisi.kompetensi}
Level Kognitif: ${kisi.levelKognitif === 'level_1' ? 'Pemahaman & Pengetahuan (Knowing - C1/C2)' : kisi.levelKognitif === 'level_2' ? 'Penerapan/Aplikasi (Applying - C3)' : 'Penalaran/Analisis Tinggi (Reasoning/HOTS - C4/C5/C6)'}
Batasan & Catatan Kurikulum: ${kisi.batasanCatatan || "Tidak ada batasan khusus"}

Tulis secara panjang lebar, menarik, dan rapi. Gunakan cetak miring (*...*) untuk bahasa asing dan cetak tebal (**...**) untuk kata/konsep khusus.${isExact ? " Sajikan rumus dan persamaan dengan notasi ilmiah yang jelas." : " TANPA TABEL."}`;
    } else {
      systemInstruction = `Anda adalah ahli prompt engineering pendidikan dan desainer instruksional kelas dunia.
Tugas Anda adalah merumuskan sebuah MEGA-PROMPT yang sangat detail, komprehensif, terstruktur rapi, dan siap saji (copy-pasteable) untuk digunakan oleh guru/pengajar di NOTEBOOK LM atau GEMINI AI guna menghasilkan INFOGRAFIS PEMBELAJARAN atau SLIDE PRESENTASI INTERAKTIF yang berkualitas tinggi, estetis, dan mendalam.`;

      userPrompt = `Buatkan MEGA-PROMPT siap pakai untuk menyusun Slide Presentasi dan Infografis Pembelajaran tingkat SMA yang sangat detail dan kaya konten untuk unit berikut:
Mata Pelajaran: ${currentMapel}
Topik / Elemen Materi: ${kisi.elemenMateri}
Sub-elemen / Sub-materi: ${kisi.subElemenMateri}
Target Kompetensi Siswa: ${kisi.kompetensi}
Level Kognitif: ${kisi.levelKognitif === 'level_1' ? 'Pemahaman & Pengetahuan (Knowing - C1/C2)' : kisi.levelKognitif === 'level_2' ? 'Penerapan/Aplikasi (Applying - C3)' : 'Penalaran/Analisis Tinggi (Reasoning/HOTS - C4/C5/C6)'}
Batasan & Catatan Kurikulum: ${kisi.batasanCatatan || "Tidak ada batasan khusus"}`;
    }

    const { text, meta } = await generateContentWithLiteLLM({
      contents: userPrompt,
      apiKeysRaw,
      baseUrl,
      config: {
        model,
        systemInstruction,
        temperature: 0.7,
      },
    });

    attachRotationHeader(res, meta);
    res.json({ materi: text });
  } catch (error: any) {
    console.error("Error generating materi:", error);
    const formatted = formatServerAiError(error);
    res.status(formatted.statusCode).json({ error: formatted.message });
  }
});

// Mount the API Router on both /api and / to handle Vercel routing variations seamlessly
app.use("/api", apiRouter);
app.use("/", apiRouter);

// Global Error Handling Middleware
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error("Global error handler caught unexpected error:", err);
  res.status(500).json({ error: err.message || "Terjadi kesalahan internal pada server." });
});

// Serve frontend static assets in production, or let Vite handle it in dev
if (process.env.NODE_ENV !== "production") {
  const startDevServer = async () => {
    const { createServer } = await import("vite");
    const vite = await createServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
    
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`[DEV] Server running on http://localhost:${PORT}`);
    });
  };
  startDevServer();
} else {
  const distPath = path.join(process.cwd(), "dist");
  app.use(express.static(distPath));
  app.get("*", (req, res) => {
    res.sendFile(path.join(distPath, "index.html"));
  });

  if (!process.env.VERCEL) {
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`[PROD] Server running on port ${PORT}`);
    });
  }
}

export default app;
