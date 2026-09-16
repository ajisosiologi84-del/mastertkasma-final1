import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Sparkles, 
  Download, 
  Printer, 
  Copy, 
  Check, 
  Plus, 
  Trash2, 
  Edit,
  BookOpen, 
  FileText, 
  FileSpreadsheet, 
  RefreshCw, 
  Sliders, 
  Layers, 
  HelpCircle, 
  Globe, 
  CheckSquare, 
  ArrowRight,
  Info,
  AlertCircle,
  Settings,
  Eye,
  EyeOff,
  Layout,
  Type,
  Upload,
  ListOrdered,
  Image as ImageIcon,
  Lock,
  LogOut,
  Users,
  UserPlus,
  Shield,
  User,
  Calendar,
  Save,
  X,
  ZoomIn,
  ZoomOut,
  Maximize2,
  Minimize2,
  RotateCw,
  Code,
  Play,
  CheckCircle2,
  Clock,
  Award,
  Terminal,
  ExternalLink,
  Send,
  Zap,
  ShieldAlert,
  Key,
  UserCheck,
  FileUp,
  ClipboardPaste,
  FileCheck,
  Database,
  FileCode
} from 'lucide-react';
import mammoth from 'mammoth';
import { KisiKisiItem, Question, GeneratorConfig, BentukSoal, LevelKognitif, JumlahOpsi, JenisSoal, JadwalItem, isExactScienceSubject, isMatrixOrMathTopic } from './types';
import { auth, db } from './lib/firebase';
import { 
  subscribeToUsers,
  exportUsersToExcel,
  downloadUserExcelTemplate,
  importUsersBatchToCloud,
  fetchUsersFromCloud,
  addUserToCloud,
  updateUserInCloud,
  deleteUserFromCloud,
  loginUser,
  getStoredSupabaseConfig,
  saveStoredSupabaseConfig,
  testSupabaseConnection,
  SUPABASE_SQL_SETUP_CODE,
  UserProfile 
} from './lib/userManagement';
import { 
  onAuthStateChanged, 
  signOut,
  User as FirebaseUser 
} from 'firebase/auth';
import { 
  doc, 
  setDoc, 
  getDoc, 
  getDocs,
  collection, 
  onSnapshot, 
  writeBatch,
  deleteDoc,
  updateDoc,
  query,
  where
} from 'firebase/firestore';
import LoginScreen from './components/LoginScreen';
import { 
  exportKisiToExcel, 
  exportKisiToWord, 
  exportQuestionsToExcel, 
  exportQuestionsToWord,
  downloadTemplateExcelSoal,
  getBentukSoalLabel,
  getLevelKognitifLabel,
  getPgkCategories,
  getPgkCategoryIndex,
  isKategoriSoal,
  exportMateriToWord,
  exportAllMateriToWord,
  markdownToHtmlForWord,
  exportJadwalToExcel,
  exportJadwalToWord,
  exportQuestionsToJSON,
  exportQuestionsWithImagesZip,
  parseQuestionsFromJSONFile
} from './utils/exportUtils';
import * as XLSX from 'xlsx';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

// Helper to clean question text from duplicate numbering prefixes
export function cleanSoalText(soalStr: string): string {
  if (!soalStr) return '';
  let text = soalStr.trim();
  text = text.replace(/^\[?\s*(soal|no\.?\s*soal|butir\s*soal)?\s*[\#\d]+\]?\s*[\.\)\:\-]\s*/i, '').trim();
  text = text.replace(/^(soal\s*)?(no\.?\s*)?\d+[\.\)\:\-]\s*/i, '').trim();
  return text;
}

// Helper to clean option text from any option letter/number prefix
export function cleanOptionText(optStr: string): string {
  if (!optStr) return '';
  let text = optStr.trim();
  text = text.replace(/^pernyataan\s*[\d|a-e]\s*[\:\.\-]\s*/i, '').trim();
  text = text.replace(/^pernyataan\s*[\d|a-e]\s+/i, '').trim();
  text = text.replace(/^[\-\•\s]*\(?([A-Ea-e1-5])\)?[\.\)\:\-]\s*/, '').trim();
  text = text.replace(/^[A-Ea-e1-5]\s+/, '').trim();
  text = text.replace(/pilihan\s*jawaban\s*:\s*\(.*?\).*/i, '').trim();
  return text;
}

// Helper to format an option string with a clean canonical prefix "A. ", "B. ", etc.
export function formatOptionString(optStr: string, index: number): string {
  const letter = String.fromCharCode(65 + index);
  const clean = cleanOptionText(optStr);
  return `${letter}. ${clean}`;
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errStr = error instanceof Error ? error.message : String(error);
  // Log the detailed firestore notice silently for developers
  console.warn(`Firestore Notice (${operationType} on ${path}):`, errStr);
  // Do NOT dispatch any popup event or throw errors - system seamlessly uses local state
  return;
}

const PUSMENDIK_MATEMATIKA_PRESETS = [
  {
    elemenMateri: 'Bilangan',
    subElemenMateri: 'Bilangan Real',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Jenis dan sifat bilangan; Operasi bilangan (penjumlahan, pengurangan, perkalian, pembagian, dan gabungannya), beserta sifat-sifatnya antara lain komutatif, asosiatif, dan distributif.',
    batasanCatatan: 'Bilangan meliputi bilangan real, termasuk bilangan asli berpangkat bilangan bulat atau berpangkat bilangan pecahan.'
  },
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Persamaan dan Pertidaksaman Linear',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Sistem persamaan linear multivariabel; Sistem pertidaksamaan linear multivariabel; Program linear.',
    batasanCatatan: 'Maksimum banyaknya variabel yang digunakan tiga.'
  },
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Fungsi',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Domain, kodomain, daerah hasil (range), dan representasi fungsi linear, kuadrat, dan rasional dalam berbagai bentuk; Invers fungsi dan representasinya; Fungsi komposisi dan representasinya.',
    batasanCatatan: 'Identifikasi fungsi meliputi secara analitis dan grafis.'
  },
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Barisan dan Deret',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Barisan dan deret aritmetika; Barisan dan deret geometri.',
    batasanCatatan: 'Penerapan barisan dan deret termasuk dalam masalah pertumbuhan, peluruhan, bunga tunggal, dan bunga majemuk.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Objek Geometri',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Hubungan dua sudut, dua garis, dan dua bidang; Hubungan objek geometri pada bangun datar dan bangun ruang; Kesebangunan atau kekongruenan bangun datar; Teorema Pythagoras.',
    batasanCatatan: 'Bangun datar meliputi segitiga, segi empat, lingkaran, dan gabungannya. Bangun ruang meliputi bangun ruang beraturan sisi datar/lengkung. Jarak dua objek meliputi jarak titik/garis/bidang.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Transformasi Geometri',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: transformasi geometri (translasi, refleksi, rotasi, dan dilatasi, serta komposisinya) dari titik.',
    batasanCatatan: 'Transformasi geometri dari titik.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Pengukuran',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Keliling dan luas bangun datar; Volume dan luas permukaan bangun ruang; Jarak dua objek geometri.',
    batasanCatatan: 'Bangun datar dan bangun ruang.'
  },
  {
    elemenMateri: 'Trigonometri',
    subElemenMateri: 'Perbandingan Trigonometri',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: perbandingan trigonometri (sinus, kosinus, tangen, kotangen, sekan, kosekan).',
    batasanCatatan: 'Perbandingan trigonometri.'
  },
  {
    elemenMateri: 'Data dan Peluang',
    subElemenMateri: 'Data',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait: Penyajian data dalam bentuk diagram batang, diagram garis, diagram lingkaran, grafik, tabel, dan bentuk visual; Ukuran pemusatan dan penyebaran data tunggal dan data kelompok; Aturan pencacahan (aturan penjumlahan, aturan perkalian, permutasi, dan kombinasi); Peluang kejadian.',
    batasanCatatan: 'Peluang dan pencacahan.'
  }
];

const PUSMENDIK_BAHASA_INDONESIA_PRESETS = [
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Kosakata & Serapan',
    kompetensi: 'Mengidentifikasi penggunaan kata serapan dari bahasa daerah/asing dalam berbagai bidang.',
    batasanCatatan: 'Teks fiksi atau nonfiksi.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Kosakata & Latar/Karakter/Fenomena',
    kompetensi: 'Mengidentifikasi latar, karakter, dan/atau fenomena berdasarkan kosakata yang digunakan dalam teks fiksi atau nonfiksi.',
    batasanCatatan: 'Teks fiksi atau nonfiksi.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Struktur Teks (Kerangka/Bagan)',
    kompetensi: 'Menyusun kerangka atau bagan berdasarkan bagian-bagian penting dalam teks.',
    batasanCatatan: 'Kerangka teks atau bagan hubungan.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Ide Pokok & Unsur Teks',
    kompetensi: 'Menyimpulkan ide pokok, gagasan pendukung, tokoh, peristiwa, latar, konflik, atau nilai-nilai dalam teks.',
    batasanCatatan: 'Teks fiksi atau nonfiksi.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Hubungan Makna',
    kompetensi: 'Menjelaskan hubungan makna antarkalimat dan/atau antarparagraf dalam teks.',
    batasanCatatan: 'Hubungan sebab-akibat, kronologis, atau komparatif.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Prediksi Kelanjutan Cerita',
    kompetensi: 'Memprediksi lanjutan atau akhir uraian/cerita berdasarkan bagian tertentu dalam teks.',
    batasanCatatan: 'Teks naratif atau ekspositoris.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Relevansi Kehidupan Nyata',
    kompetensi: 'Menilai relevansi peristiwa dalam teks dengan kehidupan sehari-hari.',
    batasanCatatan: 'Menilai nilai moral atau relevansi kontekstual.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Keakuratan & Ketepatan Informasi',
    kompetensi: 'Menilai keakuratan, kesesuaian, kecukupan, atau ketepatan informasi dalam teks.',
    batasanCatatan: 'Evaluasi kredibilitas teks nonfiksi.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Penggunaan Bahasa',
    kompetensi: 'Menilai ketepatan dan kesesuaian penggunaan bahasa dalam teks.',
    batasanCatatan: 'Kesesuaian kaidah kebahasaan dan gaya penulisan.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Penggambaran Karakter/Latar',
    kompetensi: 'Menilai ketepatan bagian teks untuk menggambarkan karakter, peristiwa, atau latar dalam teks fiksi.',
    batasanCatatan: 'Teks fiksi/sastra.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Respons Emosional Karya Sastra',
    kompetensi: 'Menyimpulkan respons emosional terhadap unsur puisi, prosa, dan drama.',
    batasanCatatan: 'Apresiasi karya sastra.'
  }
];

const PUSMENDIK_BAHASA_INGGRIS_PRESETS = [
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Menemukan/mengidentifikasi Informasi',
    kompetensi: 'Menemukan atau mengidentifikasi informasi penting yang disebutkan secara eksplisit dalam teks.',
    batasanCatatan: 'Mampu memahami informasi eksplisit secara langsung.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Mengklasifikasi',
    kompetensi: 'Mengelompokkan orang, benda, tempat, atau peristiwa dalam teks berdasarkan kategori tertentu.',
    batasanCatatan: 'Klasifikasi data tekstual.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Membuat Kerangka',
    kompetensi: 'Menyusun poin-poin utama dari teks dalam bentuk kerangka atau daftar.',
    batasanCatatan: 'Outline/draft struktur informasi.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Meringkas',
    kompetensi: 'Menyajikan kembali isi teks secara ringkas dengan mengutip bagian penting.',
    batasanCatatan: 'Ringkasan esensi teks.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Mensintesis',
    kompetensi: 'Menggabungkan informasi dari sumber lain untuk mendapatkan pemahaman yang lebih komprehensif tentang suatu isu atau topik.',
    batasanCatatan: 'Sintesis multi-sumber/multi-teks.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Detail Pendukung',
    kompetensi: 'Menentukan fakta tambahan yang membuat teks lebih informatif, menarik, atau persuasif.',
    batasanCatatan: 'Inferensi fakta pendukung.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Topik & Gagasan Utama',
    kompetensi: 'Menyimpulkan topik, ide pokok/gagasan utama, makna, target pembaca, tujuan penulisan teks, atau pesan moral yang tidak secara eksplisit dinyatakan dalam teks.',
    batasanCatatan: 'Gagasan utama tersirat, target pembaca, dan moral value.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Urutan Kejadian',
    kompetensi: 'Memperkirakan urutan kejadian dan memperkirakan isi selanjutnya dari teks.',
    batasanCatatan: 'Kronologis dan kelanjutan teks.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Perbandingan',
    kompetensi: 'Menyimpulkan persamaan atau perbedaan antara tokoh, waktu, tempat, benda, atau gagasan dalam teks.',
    batasanCatatan: 'Komparasi dan kontras elemen teks.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Hubungan Sebab-Akibat',
    kompetensi: 'Menafsirkan hubungan/kaitan antara gagasan/tindakan satu dan lainnya yang dinyatakan dalam teks.',
    batasanCatatan: 'Hubungan kausalitas/sebab-akibat.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Karakter Tokoh',
    kompetensi: 'Menyimpulkan sifat atau kepribadian tokoh berdasarkan petunjuk eksplisit maupun implisit dalam teks.',
    batasanCatatan: 'Analisis karakter/tokoh.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Memprediksi Hasil Cerita',
    kompetensi: 'Memprediksi akhir cerita setelah membaca bagian awal atau bagian tertentu dari teks.',
    batasanCatatan: 'Prediksi konklusi cerita.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai Realitas atau Fantasi',
    kompetensi: 'Menganalisis peristiwa dalam teks dapat terjadi dalam kehidupan nyata berdasarkan pengalaman dan pengetahuan pribadi.',
    batasanCatatan: 'Analisis realitas vs fantasi.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai Fakta atau Opini',
    kompetensi: 'Menilai fakta/opini yang diberikan penulis untuk mendukung pendapatnya berdasarkan bukti atau sekadar berusaha mempengaruhi pembaca.',
    batasanCatatan: 'Fakta vs Opini.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai Kecukupan & Validitas Informasi',
    kompetensi: 'Menilai kesesuaian, kelengkapan, keakuratan informasi dalam teks (dengan membandingkannya dengan sumber lain).',
    batasanCatatan: 'Evaluasi validitas & kelengkapan data.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai Kesesuaian Penggambaran',
    kompetensi: 'Menentukan bagian teks yang paling sesuai untuk menggambarkan karakter utama atau aspek lain dari bacaan.',
    batasanCatatan: 'Kesesuaian representasi teks.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menanggapi Isi Teks (Respons Emosional)',
    kompetensi: 'Mengungkapkan perasaan/kesan/pendapat terhadap bacaan, seperti ketertarikan, kebosanan, kegembiraan, ketakutan, kebencian, atau kesenangan.',
    batasanCatatan: 'Respons subjektif, estetis, dan emosional.'
  }
];

const PUSMENDIK_MATEMATIKA_TL_PRESETS = [
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Matriks',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Determinan matriks; Invers matriks; Operasi matriks. Program linear.',
    batasanCatatan: 'Elemen matriks merupakan bilangan real. Determinan dan invers matriks berukuran 2 x 2 atau 3 x 3.'
  },
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Polinomial',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Operasi polinomial; Pemfaktoran polinomial; Suku sisa.',
    batasanCatatan: 'Orde dari polinomial maksimum 4 dan semua koefisien polinomial berupa bilangan real.'
  },
  {
    elemenMateri: 'Aljabar',
    subElemenMateri: 'Fungsi',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Domain, kodomain, daerah hasil (range), dan grafik fungsi polinom, rasional, akar, eksponensial, logaritma, mutlak, trigonometri.',
    batasanCatatan: 'Fungsi polinom maksimum berorde 4. Bilangan pokok fungsi eksponensial berupa bilangan asli.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Vektor',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Vektor pada bidang dan ruang; Panjang vektor; Operasi vektor.',
    batasanCatatan: 'Komponen vektor maksimum tiga.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Lingkaran',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Persamaan lingkaran dan persamaan garis singgung lingkaran; Luas dan keliling daerah lingkaran atau bagian daerah lingkaran.',
    batasanCatatan: 'Keliling, luas, dan persamaan garis singgung.'
  },
  {
    elemenMateri: 'Geometri dan Pengukuran',
    subElemenMateri: 'Transformasi Geometri',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Transformasi geometri (translasi, refleksi, rotasi, dilatasi, serta komposisinya) dari bentuk geometris dan matriks transformasinya.',
    batasanCatatan: 'Bentuk geometris yang ditransformasi meliputi titik, garis, dan bangun datar.'
  },
  {
    elemenMateri: 'Trigonometri',
    subElemenMateri: 'Limit',
    kompetensi: 'Memahami, mengaplikasikan, dan bernalar yang lebih tinggi untuk menyelesaikan permasalahan terkait cakupan sub-elemen berikut: Limit fungsi aljabar; Limit fungsi trigonometri.',
    batasanCatatan: 'Limit yang dapat diselesaikan tanpa menggunakan Teorema L’Hopital.'
  }
];

const PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS = [
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Teks Akademik',
    kompetensi: 'Mengidentifikasi informasi dalam teks akademik.',
    batasanCatatan: 'Teks akademik / ilmiah.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Penyampaian Tanggapan & Kritik',
    kompetensi: 'Mengidentifikasi kalimat yang tepat untuk menyampaikan tanggapan, respons, dan kritik sesuai norma sosial dan budaya.',
    batasanCatatan: 'Norma sosial dan budaya.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Pengajuan Usulan & Solusi',
    kompetensi: 'Mengidentifikasi kalimat yang tepat dalam pengajuan usulan, perumusan masalah, dan pemecahan masalah pada teks dalam bidang akademik dan/atau dunia kerja.',
    batasanCatatan: 'Bidang akademik dan/atau dunia kerja.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Sastra Indonesia & Terjemahan',
    kompetensi: 'Mengidentifikasi karakter, peristiwa, latar pada sastra Indonesia atau terjemahan.',
    batasanCatatan: 'Karya sastra atau terjemahannya.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Nilai Karya Sastra',
    kompetensi: 'Membandingkan nilai-nilai (budaya, sosial, moral, religius, dan/atau pendidikan) dalam karya sastra Indonesia dan/atau terjemahan.',
    batasanCatatan: 'Perbandingan nilai-nilai sastra.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Sastra Melayu Klasik',
    kompetensi: 'Mengungkapkan kembali isi sastra Melayu Klasik.',
    batasanCatatan: 'Sastra Melayu Klasik.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Konversi Uraian & Visual',
    kompetensi: 'Mengubah informasi dari tabel/grafik menjadi uraian atau uraian menjadi tabel/grafik dalam bidang akademik dan/atau dunia kerja.',
    batasanCatatan: 'Tabel, grafik, atau diagram.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Gaya Bahasa & Kiasan',
    kompetensi: 'Menjelaskan ketepatan penggunaan bahasa, kiasan, dan atau citraan dalam teks.',
    batasanCatatan: 'Majas, pencitraan, dan diksi.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Kaidah Kebahasaan (Sintaksis & Morfologi)',
    kompetensi: 'Menjelaskan ketepatan penggunaan afiks, konstruksi frasa, konstruksi klausa, dan/atau kalimat dalam teks.',
    batasanCatatan: 'Afiksasi, struktur frasa, klausa, kalimat.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Kohesi & Koherensi',
    kompetensi: 'Menjelaskan kohesi dan koherensi dalam teks ilmiah.',
    batasanCatatan: 'Kepaduan wacana/teks ilmiah.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Logika Berpikir',
    kompetensi: 'Menilai gagasan dan pandangan dalam berbagai teks (digital atau cetak) berdasarkan kaidah logika berpikir.',
    batasanCatatan: 'Logika berpikir, sesat pikir (fallacy).'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Kombinasi & Perbandingan Antarteks',
    kompetensi: 'Menilai ketepatan dan kesesuaian isi antarteks (digital atau cetak) dalam bidang sosial, akademik, dan dunia kerja.',
    batasanCatatan: 'Perbandingan multi-teks.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Evaluasi Tokoh & Norma',
    kompetensi: 'Menilai gagasan atau tindakan tokoh berdasarkan norma atau nilai individu dan sosial.',
    batasanCatatan: 'Norma individu/sosial.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Alih Wahana Karya',
    kompetensi: 'Mengalihwahanakan puisi (Indonesia dan/atau terjemahan) dalam bentuk prosa.',
    batasanCatatan: 'Parafrasa / alih wahana puisi ke prosa.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Respons Sastra',
    kompetensi: 'Menyimpulkan respons emosional terhadap unsur puisi, prosa, dan drama Indonesia atau terjemahan.',
    batasanCatatan: 'Apresiasi & respons subjektif/emosional.'
  }
];

const PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS = [
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Menemukan/mengidentifikasi informasi',
    kompetensi: 'Mampu menemukan atau mengidentifikasi gagasan utama serta informasi penting yang secara eksplisit disebutkan dalam teks.',
    batasanCatatan: 'Informasi eksplisit, gagasan utama literal.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Mengklasifikasi',
    kompetensi: 'Mampu mengelompokkan argumen, fakta, dan pendapat dalam teks berdasarkan kategori atau pola penyajian tertentu.',
    batasanCatatan: 'Kategorisasi argumen/opini.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Membuat kerangka',
    kompetensi: 'Mampu menyusun poin-poin utama dalam teks ke dalam bentuk peta konsep/diagram/diagram alir atau daftar untuk memahami struktur penyajian informasi.',
    batasanCatatan: 'Peta konsep, diagram alir, atau daftar terstruktur.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Meringkas',
    kompetensi: 'Mampu menyajikan kembali isi teks secara ringkas dengan tetap mempertahankan gagasan utama dan argumen kunci.',
    batasanCatatan: 'Ringkasan representatif.'
  },
  {
    elemenMateri: 'Pemahaman Tekstual',
    subElemenMateri: 'Mensintesis',
    kompetensi: 'Mampu menggabungkan informasi dari sumber lain untuk mendapatkan pemahaman yang lebih komprehensif tentang suatu isu atau topik.',
    batasanCatatan: 'Sintesis multi-teks/multi-sumber.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan detail pendukung',
    kompetensi: 'Mampu memperkirakan fakta tambahan yang mungkin dapat memperkuat atau memperjelas argumen dalam teks.',
    batasanCatatan: 'Fakta tambahan pendukung argumen.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan gagasan utama & tujuan',
    kompetensi: 'Mampu menyimpulkan topik, ide pokok/gagasan utama, makna, target pembaca, tujuan penulisan teks, atau pesan moral yang tidak secara eksplisit dinyatakan dalam teks.',
    batasanCatatan: 'Pesan moral, target pembaca, gagasan utama tersirat.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan hubungan antar-ide',
    kompetensi: 'Mampu menghubungkan berbagai argumen, alasan, dan bukti dalam teks untuk memahami logika penyajian informasi.',
    batasanCatatan: 'Logika penyajian & struktur argumen.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan hubungan sebab-akibat',
    kompetensi: 'Mampu menyimpulkan suatu peristiwa/kebijakan/fenomena dalam teks memengaruhi atau dipengaruhi oleh faktor lain.',
    batasanCatatan: 'Kausalitas dan pengaruh faktor luar.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan sudut pandang penulis',
    kompetensi: 'Mampu mengenali posisi atau sikap penulis terhadap suatu isu berdasarkan bahasa dan pilihan argumen yang digunakan.',
    batasanCatatan: 'Sikap, nada (tone), dan posisi penulis.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Memprediksi implikasi atau konsekuensi',
    kompetensi: 'Mampu memperkirakan dampak dari suatu gagasan atau argumen yang disampaikan dalam teks.',
    batasanCatatan: 'Prediksi dampak/konsekuensi.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai fakta atau opini',
    kompetensi: 'Mampu mengevaluasi fakta atau opini dalam teks berdasarkan bukti-bukti pendukung yang disajikan penulis.',
    batasanCatatan: 'Fakta vs Opini.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai keakuratan & kecukupan informasi',
    kompetensi: 'Mampu menilai kredibilitas informasi dan cakupan perspektif dalam teks.',
    batasanCatatan: 'Kredibilitas & cakupan perspektif.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai kecukupan & validitas informasi',
    kompetensi: 'Menilai kesesuaian, kelengkapan, keakuratan informasi dalam teks (dengan membandingkannya dengan sumber lain).',
    batasanCatatan: 'Validasi dengan sumber sekunder/eksternal.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menilai kekuatan argumen',
    kompetensi: 'Mampu mengevaluasi seberapa logis dan meyakinkan argumen yang disajikan dalam teks.',
    batasanCatatan: 'Kekuatan logika & persuasi argumen.'
  },
  {
    elemenMateri: 'Evaluasi dan Apresiasi',
    subElemenMateri: 'Menanggapi isi teks secara kritis',
    kompetensi: 'Mampu memberikan opini atau refleksi mengenai isi teks dari sudut pandang yang berbeda.',
    batasanCatatan: 'Refleksi kritis & opini personal.'
  }
];

const PUSMENDIK_FISIKA_PRESETS = [
  {
    elemenMateri: 'Kinematika',
    subElemenMateri: 'Pengukuran',
    kompetensi: 'Menelaah hasil pengukuran suatu besaran dengan alat ukur yang sesuai serta menyatakan hasil pengukuran sesuai aturan angka penting.',
    batasanCatatan: 'Alat ukur dapat meliputi jangka sorong, mikrometer sekrup, neraca tiga lengan. Pelaporan hasil pengukuran mengikuti kaidah angka penting.'
  },
  {
    elemenMateri: 'Kinematika',
    subElemenMateri: 'Gerak Lurus',
    kompetensi: 'Menganalisis keterkaitan beberapa besaran pada gerak lurus berdasarkan data yang ada untuk menyelesaikan masalah yang relevan.',
    batasanCatatan: 'Gerak lurus: meliputi konsep gerak (jarak, perpindahan, kecepatan dengan memperhatikan besar dan arahnya), GLB, dan GLBB.'
  },
  {
    elemenMateri: 'Kinematika',
    subElemenMateri: 'Gerak Lengkung',
    kompetensi: 'Mengaitkan hubungan antar variabel dalam gerak parabola dan gerak melingkar beraturan pada peristiwa dalam kehidupan sehari-hari.',
    batasanCatatan: 'Gerak parabola: dibatasi pada gerak parabola dengan referensi bidang horizontal. Gerak melingkar: dibatasi pada gerak melingkar beraturan.'
  },
  {
    elemenMateri: 'Dinamika',
    subElemenMateri: 'Hubungan Gaya dan Gerak (Hukum Newton)',
    kompetensi: 'Mengaitkan hubungan antara gaya dan gerak pada peristiwa yang terjadi dalam kehidupan sehari-hari.',
    batasanCatatan: 'Hukum Newton 1, 2, dan 3, jenis-jenis gaya (dengan memperhatikan besar dan arahnya).'
  },
  {
    elemenMateri: 'Dinamika',
    subElemenMateri: 'Momentum dan Impuls',
    kompetensi: 'Menerapkan konsep momentum dan impuls bagi penyelesaian masalah sehari-hari.',
    batasanCatatan: 'Konsep momentum & impuls, hukum kekekalan momentum.'
  },
  {
    elemenMateri: 'Dinamika',
    subElemenMateri: 'Dinamika Rotasi (Momen Gaya dan Momen Inersia)',
    kompetensi: 'Menerapkan konsep momen gaya dan momen inersia yang dimanfaatkan dalam kehidupan sehari-hari.',
    batasanCatatan: 'Momen gaya, momen inersia, Hukum Newton tentang rotasi.'
  },
  {
    elemenMateri: 'Fluida',
    subElemenMateri: 'Fluida Statis dan Dinamis',
    kompetensi: 'Menerapkan konsep yang berkaitan dengan fluida statis dan dinamis pada teknologi yang dimanfaatkan dalam kehidupan sehari-hari.',
    batasanCatatan: 'Tekanan fluida, Hukum Pascal, Hukum Archimedes, kontinuitas dan Hukum Bernoulli.'
  },
  {
    elemenMateri: 'Gelombang',
    subElemenMateri: 'Bunyi',
    kompetensi: 'Menganalisis keterkaitan sifat bunyi dengan parameter gelombangnya berdasarkan peristiwa dalam kehidupan sehari-hari.',
    batasanCatatan: 'Karakteristik gelombang mekanik dalam gelombang bunyi, sumber dan intensitas bunyi.'
  },
  {
    elemenMateri: 'Gelombang',
    subElemenMateri: 'Cahaya',
    kompetensi: 'Menganalisis keterkaitan sifat cahaya dengan peristiwa dalam kehidupan sehari-hari serta penerapannya.',
    batasanCatatan: 'Karakteristik gelombang elektromagnetik, sifat cahaya serta penerapannya dalam cermin, lensa, dan alat optik.'
  },
  {
    elemenMateri: 'Kalor dan Termodinamika',
    subElemenMateri: 'Kalor dan Perpindahannya',
    kompetensi: 'Menguraikan pengaruh kalor terhadap kenaikan suhu, perubahan wujud zat, serta faktor-faktor yang mempengaruhi perpindahan kalor dalam kehidupan sehari-hari.',
    batasanCatatan: 'Kalor untuk perubahan suhu, perubahan wujud, pemuaian, perpindahan kalor, serta peristiwa pencampuran zat.'
  },
  {
    elemenMateri: 'Kalor dan Termodinamika',
    subElemenMateri: 'Pemanasan Global',
    kompetensi: 'Menganalisis penyebab, proses, dan cara mengatasi permasalahan terkait dengan pemanasan global.',
    batasanCatatan: 'Efek rumah kaca: fokus pada proses, penyebab, cara mengatasinya.'
  },
  {
    elemenMateri: 'Kalor dan Termodinamika',
    subElemenMateri: 'Gas Ideal',
    kompetensi: 'Mengidentifikasi besaran tertentu pada perubahan suhu, tekanan, dan/atau volume gas ideal dalam ruang tertutup.',
    batasanCatatan: 'Hukum Boyle, Boyle-Gay Lussac, dan persamaan gas ideal.'
  },
  {
    elemenMateri: 'Kalor dan Termodinamika',
    subElemenMateri: 'Termodinamika',
    kompetensi: 'Menganalisis perubahan kalor, perubahan energi dalam, atau usaha dalam proses termodinamika.',
    batasanCatatan: 'Empat proses termodinamika, hukum Termodinamik, mesin kalor.'
  },
  {
    elemenMateri: 'Kelistrikan',
    subElemenMateri: 'Listrik Statis',
    kompetensi: 'Menganalisis keterkaitan besaran listrik statis dalam peristiwa yang terjadi dalam kehidupan sehari-hari.',
    batasanCatatan: 'Gejala listrik statis, hukum Coulomb, medan listrik, potensial listrik, dan penerapannya.'
  },
  {
    elemenMateri: 'Kelistrikan',
    subElemenMateri: 'Rangkaian Arus Searah',
    kompetensi: 'Menguraikan hubungan antara kuat arus, tegangan, dan hambatan dalam rangkaian listrik arus searah berdasarkan hukum-hukum kelistrikan.',
    batasanCatatan: 'Hukum Ohm dan Hukum Kirchoff dibatasi pada rangkaian listrik campuran (seri-paralel) dengan arus searah hingga perhitungan daya.'
  },
  {
    elemenMateri: 'Keterampilan Proses Sains',
    subElemenMateri: 'Penyelidikan Ilmiah, Mengamati, Mempertanyakan & Mengomunikasikan Data',
    kompetensi: 'Mengamati, mempertanyakan, memprediksi, merencanakan penyelidikan, menentukan variabel/alat/prosedur, serta menganalisis dan mengomunikasikan data eksperimen secara sistematis.',
    batasanCatatan: 'Keterampilan proses sains: mengamati, merumuskan hipotesis, merancang percobaan, mengorganisasi data tabel/grafik, dan mengomunikasikan hasil.'
  }
];

const PUSMENDIK_KIMIA_PRESETS = [
  {
    elemenMateri: 'Pemahaman (Knowing)',
    subElemenMateri: 'Mengenali (Recognize)',
    kompetensi: 'Mengidentifikasi atau menyatakan fakta, hubungan, dan konsep; mengidentifikasi karakteristik fisika/kimia serta peran setiap komponen yang terdapat dalam sistem, materi, dan proses tertentu.',
    batasanCatatan: 'Fakta, hubungan, konsep dasar, sifat fisis/kimia, dan peranan komponen dalam sistem.'
  },
  {
    elemenMateri: 'Pemahaman (Knowing)',
    subElemenMateri: 'Menjelaskan (Describe)',
    kompetensi: 'Memberikan informasi atau penjelasan secara rinci berdasarkan konsep kimia/mendeskripsikan sifat dan struktur suatu materi serta proses atau fenomena kimia.',
    batasanCatatan: 'Deskripsi rinci sifat, struktur materi, dan fenomena/proses kimia.'
  },
  {
    elemenMateri: 'Pemahaman (Knowing)',
    subElemenMateri: 'Memberikan Contoh (Provide Example)',
    kompetensi: 'Menuliskan contoh yang berkaitan dengan suatu fenomena, kegunaan, maupun kerugian suatu materi/proses kimia yang relevan dalam kehidupan sehari-hari.',
    batasanCatatan: 'Contoh fenomena, kegunaan, dan dampak negatif zat/reaksi kimia.'
  },
  {
    elemenMateri: 'Penerapan (Applying)',
    subElemenMateri: 'Membandingkan & Mengklasifikasikan',
    kompetensi: 'Menerapkan pengetahuan tentang fakta, hubungan, proses, konsep, dan metode ilmiah dalam menyelesaikan masalah sesuai konteks yang disajikan. Mengidentifikasi persamaan/perbedaan dari suatu zat/proses kimia. Mengelompokkan berbagai zat/proses berdasarkan sifat-sifatnya.',
    batasanCatatan: 'Komparasi zat/reaksi dan klasifikasi berdasarkan karakteristik fisik/kimia.'
  },
  {
    elemenMateri: 'Penerapan (Applying)',
    subElemenMateri: 'Menginterpretasikan Model',
    kompetensi: 'Menggunakan pengetahuan tentang konsep-konsep sains untuk menginterpretasikan proses, siklus, hubungan/sistem untuk menyelesaikan masalah kimia.',
    batasanCatatan: 'Analisis dan pemaknaan diagram proses, siklus, atau sistem kimia.'
  },
  {
    elemenMateri: 'Penerapan (Applying)',
    subElemenMateri: 'Menginterpretasikan Informasi',
    kompetensi: 'Menggunakan pengetahuan atau konsep untuk menjelaskan informasi tekstual, tabular, gambar, dan grafis yang relevan termasuk melakukan perhitungan kimia.',
    batasanCatatan: 'Penjelasan tabel, grafik, data numerik, dan perhitungan stoikiometri.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Menganalisis (Analyze)',
    kompetensi: 'Menganalisis hasil uraian atau perhitungan kimia menggunakan prinsip, konsep, rumus, dan hukum kimia untuk menarik suatu kesimpulan yang ilmiah. Menggunakan informasi yang relevan, konsep, hubungan antar parameter, dan data untuk menjawab pertanyaan.',
    batasanCatatan: 'Analisis korelasi data, hukum kimia dasar, dan perhitungan mendalam.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Memprediksi (Predict)',
    kompetensi: 'Memperkirakan hasil yang diperoleh dari penjelasan yang bersumber dari teori, konsep, hasil perhitungan, dan hasil analisis.',
    batasanCatatan: 'Prediksi hasil reaksi, arah kesetimbangan, atau parameter produk baru.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Mengevaluasi (Evaluate)',
    kompetensi: 'Mengevaluasi suatu hasil pemikiran atau penjelasan berdasarkan teori, konsep, hasil perhitungan, dan hasil analisis sehingga bisa diambil suatu kesimpulan dari masalah yang ingin diselesaikan. Mengevaluasi fenomena kimia sehari-hari.',
    batasanCatatan: 'Uji validitas argumentasi, optimasi resep kimia, atau evaluasi fenomena.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Merancang (Design)',
    kompetensi: 'Membuat rancangan eksperimen yang dapat digunakan untuk menjawab fenomena dalam kimia, dengan melibatkan variabel riset yang meliputi variabel bebas, variabel tergantung, dan variabel kontrol.',
    batasanCatatan: 'Desain eksperimen, penentuan metodologi, dan pengontrolan variabel.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Menguji Hipotesis',
    kompetensi: 'Menyelesaikan soal yang hasilnya sudah dihipotesiskan melalui langkah penerapan konsep kimia, perhitungan, analisis, dan pengambilan kesimpulan sehingga bisa mengambil kesimpulan yang sesuai dengan hipotesis atau tidak.',
    batasanCatatan: 'Pengujian hipotesis teoritis atau eksperimental dengan bukti ilmiah.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Menghubungkan Variabel',
    kompetensi: 'Menghubungkan karakteristik dengan karakteristik yang lain suatu material sehingga menghasilkan sifat material yang bisa dijelaskan secara ilmiah.',
    batasanCatatan: 'Hubungan struktur mikro-makro, fasa, maupun sifat fisis material.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Mengaplikasikan Prinsip Ilmiah',
    kompetensi: 'Menerapkan konsep kimia dan prinsip saintifik untuk membuat formulasi, mengevaluasi kualitas produk, dan menyelesaikan masalah yang dihadapi oleh masyarakat dan kalangan industri.',
    batasanCatatan: 'Formulasi produk, kontrol kualitas, dan pemecahan isu industri/lingkungan.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Menarik Kesimpulan (Draw Conclusion)',
    kompetensi: 'Membuat kesimpulan berdasarkan penerapan teori, konsep, data, dan perhitungan kimia serta mampu menghubungkan antara penyebab dan akibat dalam suatu proses kimia.',
    batasanCatatan: 'Konklusi deduktif/induktif dan penalaran sebab-akibat (kausalitas).'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Menggeneralisasi (Generalize)',
    kompetensi: 'Membuat kesimpulan dari suatu hasil penyelesaian masalah melalui pemahaman teori, konsep, perhitungan kimia, dan proses sintesis. Dan bisa menerapkan kesimpulan tersebut untuk kondisi yang baru.',
    batasanCatatan: 'Ekstrapolasi pola perilaku unsur periodik atau reaksi ke kasus baru.'
  },
  {
    elemenMateri: 'Penalaran (Reasoning)',
    subElemenMateri: 'Memberikan Penjelasan Disertai Bukti (Justify)',
    kompetensi: 'Menggunakan bukti dan pengetahuan sains untuk menjelaskan suatu fenomena untuk mendukung suatu penjelasan yang sulit dijelaskan alasannya atau penyelesaian masalah dan kesimpulan dari suatu investigasi.',
    batasanCatatan: 'Justifikasi ilmiah berbasis data empiris maupun postulat teori.'
  }
];

const PUSMENDIK_BIOLOGI_PRESETS = [
  {
    elemenMateri: 'Keanekaragaman Hayati',
    subElemenMateri: 'Klasifikasi & Keanekaragaman Makhluk Hidup',
    kompetensi: 'Menganalisis prinsip klasifikasi dan permasalahan keanekaragaman hayati.',
    batasanCatatan: 'Keanekaragaman hayati yang dimaksud adalah keanekaragaman hayati di Indonesia.'
  },
  {
    elemenMateri: 'Keanekaragaman Hayati',
    subElemenMateri: 'Bakteri',
    kompetensi: 'Menganalisis struktur bakteri (Gram positif dan negatif) dan peranannya terhadap manusia beserta resistensi bakteri.',
    batasanCatatan: 'Struktur bakteri, resistensi bakteri, dan peranan bakteri bagi manusia.'
  },
  {
    elemenMateri: 'Keanekaragaman Hayati',
    subElemenMateri: 'Ekosistem',
    kompetensi: 'Menganalisis komponen-komponen ekosistem, interaksi antarkomponen, dan solusi atas permasalahannya serta pelestarian ekosistem.',
    batasanCatatan: 'Ekosistem yang ada di Indonesia, komponen ekosistem, interaksi, pelestarian.'
  },
  {
    elemenMateri: 'Sel',
    subElemenMateri: 'Metabolisme Sel',
    kompetensi: 'Menganalisis proses metabolisme (katabolisme dan anabolisme) dan peran serta cara kerja enzim.',
    batasanCatatan: 'Katabolisme; anabolisme; sifat dan cara kerja enzim.'
  },
  {
    elemenMateri: 'Proses-proses pada Makhluk Hidup',
    subElemenMateri: 'Transport & Pertukaran Zat pada Manusia (Sirkulasi, Respirasi, Ekskresi)',
    kompetensi: 'Menganalisis keterkaitan struktur organ pada sistem sirkulasi, respirasi, dan ekskresi beserta fungsinya masing-masing.',
    batasanCatatan: 'Keterkaitan antarsistem sirkulasi, respirasi, dan ekskresi.'
  },
  {
    elemenMateri: 'Proses-proses pada Makhluk Hidup',
    subElemenMateri: 'Sistem Imun',
    kompetensi: 'Menganalisis peran sistem imun terhadap kekebalan tubuh dan mekanisme kerjanya.',
    batasanCatatan: 'Mekanisme kerja sistem imun.'
  },
  {
    elemenMateri: 'Proses-proses pada Makhluk Hidup',
    subElemenMateri: 'Sistem Koordinasi',
    kompetensi: 'Menganalisis sistem koordinasi dalam tubuh manusia meliputi kerja saraf dan hormon.',
    batasanCatatan: 'Mekanisme kerja sistem saraf dan sistem hormon.'
  },
  {
    elemenMateri: 'Proses-proses pada Makhluk Hidup',
    subElemenMateri: 'Sistem Reproduksi',
    kompetensi: 'Menganalisis keterkaitan struktur organ pada sistem reproduksi pria dan wanita serta fungsinya.',
    batasanCatatan: 'Struktur dan fungsi organ reproduksi pria dan wanita.'
  },
  {
    elemenMateri: 'Keterampilan Proses',
    subElemenMateri: 'Mempertanyakan dan memprediksi',
    kompetensi: 'Merumuskan pertanyaan yang dapat diselidiki secara ilmiah.',
    batasanCatatan: 'Konteks soal sesuai dengan konten keanekaragaman hayati, sel, dan proses pada makhluk hidup.'
  },
  {
    elemenMateri: 'Keterampilan Proses',
    subElemenMateri: 'Merencanakan dan melakukan penyelidikan',
    kompetensi: 'Merancang penyelidikan ilmiah.',
    batasanCatatan: 'Konteks soal sesuai dengan konten keanekaragaman hayati, sel, dan proses pada makhluk hidup.'
  },
  {
    elemenMateri: 'Keterampilan Proses',
    subElemenMateri: 'Memproses, menganalisis data dan informasi',
    kompetensi: 'Mengolah data dan menyimpulkan hasil penyelidikan.',
    batasanCatatan: 'Konteks soal sesuai dengan konten keanekaragaman hayati, sel, dan proses pada makhluk hidup.'
  }
];

const PUSMENDIK_PPKN_PRESETS = [
  {
    elemenMateri: 'Pancasila',
    subElemenMateri: 'Pancasila sebagai dasar negara, ideologi negara, identitas nasional, hak asasi manusia dan demokrasi Pancasila.',
    kompetensi: 'Menjelaskan, menerapkan dan menganalisis makna sila- sila Pancasila, sejarah perumusan Pancasila, dasar negara, ideologi, identitas nasional, pelaksanaan hak asasi manusia, demokrasi Pancasila, permasalahan dan solusi dalam penerapan nilai-nilai Pancasila.',
    batasanCatatan: 'Pancasila sebagai dasar negara, ideologi, identitas nasional, HAM, dan demokrasi.'
  },
  {
    elemenMateri: 'Undang-Undang Dasar Negara Republik Indonesia Tahun 1945',
    subElemenMateri: 'Penegakan hukum, perlindungan HAM, ketentuan UUD NRI Tahun 1945, demokrasi, hubungan pemerintah pusat dan daerah, kewenangan lembaga negara, hak dan kewajiban warga negara.',
    kompetensi: 'Menjelaskan, menerapkan, dan menganalisis perilaku taat hukum, sejarah dan perkembangan undang- undang dasar di Indonesia, kewenangan lembaga negara menurut UUD NRI Tahun 1945, hubungan pemerintah pusat dengan pemerintah daerah, demokrasi, hak dan kewajiban warga negara.',
    batasanCatatan: 'Ketentuan UUD NRI 1945, lembaga negara, otonomi daerah, serta hak & kewajiban warga negara.'
  },
  {
    elemenMateri: 'Bhinneka Tunggal Ika',
    subElemenMateri: 'Integrasi nasional, mengelola kebinekaan sebagai modal sosial, harmoni dalam keberagaman, implementasi prinsip gotong royong, dan ancaman terhadap kebinekaan.',
    kompetensi: 'Menjelaskan, menerapkan dan menganalisis kebersamaan dan keberagaman dalam Bhinneka Tunggal Ika, implementasi prinsip gotong royong, kebinekaan sebagai modal sosial, dan potensi ancaman terhadap keberagaman.',
    batasanCatatan: 'Integrasi nasional, pengelolaan kebinekaan, prinsip gotong royong, dan ancaman keberagaman.'
  },
  {
    elemenMateri: 'Negara Kesatuan Republik Indonesia',
    subElemenMateri: 'Perilaku warga negara yang baik, bentuk negara, bentuk dan sistem pemerintahan, pengaruh kemajuan IPTEKS terhadap NKRI, menjaga keutuhan NKRI dalam konteks Wawasan Nusantara, menjadi pelopor pemilih pemula dalam demokrasi Indonesia, menjaga keutuhan NKRI, sistem pertahanan Indonesia, peran Indonesia dalam perdamaian dunia, dan demokrasi Indonesia.',
    kompetensi: 'Menjelaskan, menerapkan dan menganalisis perilaku yang sesuai dengan hak dan kewajiban warga negara, menjaga keutuhan NKRI, peran Indonesia dalam perdamaian dunia, sistem pertahanan dan keamanan negara, praktik demokrasi, bentuk negara, bentuk pemerintahan, dan sistem pemerintahan.',
    batasanCatatan: 'NKRI, peran dunia, pertahanan keamanan, bentuk & sistem pemerintahan, dampak IPTEK.'
  }
];

const PUSMENDIK_EKONOMI_PRESETS = [
  {
    elemenMateri: 'Konsep Dasar Ilmu Ekonomi',
    subElemenMateri: 'Konsep dasar ilmu ekonomi: kelangkaan, biaya peluang, dan kegiatan ekonomi',
    kompetensi: 'Menganalisis konsep dasar ilmu ekonomi mencakup kelangkaan, biaya peluang, dan kegiatan ekonomi',
    batasanCatatan: 'Prinsip kelangkaan, penentuan biaya peluang, dan motif/kegiatan ekonomi sehari-hari.'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Permintaan, penawaran, dan keseimbangan pasar',
    kompetensi: 'Menganalisis permintaan, penawaran, dan keseimbangan pasar',
    batasanCatatan: 'Hukum dan kurva permintaan-penawaran, pergeseran kurva, serta titik keseimbangan pasar (ekuilibrium).'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Pendapatan nasional, pertumbuhan ekonomi, dan pembangunan ekonomi',
    kompetensi: 'Menganalisis konsep pendapatan nasional, pertumbuhan ekonomi, dan pembangunan ekonomi',
    batasanCatatan: 'Metode perhitungan GDP/GNP, indikator pertumbuhan ekonomi, dan strategi pembangunan nasional.'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Ketenagakerjaan',
    kompetensi: 'Menganalisis konsep ketenagakerjaan dan permasalahannya',
    batasanCatatan: 'Angkatan kerja, jenis-jenis pengangguran, serta upaya mengatasi masalah ketenagakerjaan di Indonesia.'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Indeks Harga dan Inflasi',
    kompetensi: 'Mengevaluasi indeks harga dan inflasi',
    batasanCatatan: 'Metode perhitungan indeks harga, penyebab/jenis inflasi, serta dampak inflasi bagi perekonomian.'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Bank sentral dan kebijakan moneter',
    kompetensi: 'Mengidentifikasi peran bank sentral dan menganalisis kebijakan moneter',
    batasanCatatan: 'Tugas dan wewenang Bank Indonesia sebagai bank sentral, instrumen kebijakan moneter (diskonto, pasar terbuka, cadangan wajib).'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Kebijakan fiskal dan perpajakan',
    kompetensi: 'Menerapkan konsep perpajakan dan menganalisis kebijakan fiskal',
    batasanCatatan: 'Fungsi pajak, tarif pajak, APBN/APBD, serta instrumen/tujuan kebijakan fiskal.'
  },
  {
    elemenMateri: 'Ekonomi Mikro dan Makro',
    subElemenMateri: 'Manajemen badan usaha dalam perekonomian Indonesia (BUMN, BUMD, BUMS, dan Koperasi)',
    kompetensi: 'Mendeskripsikan konsep manajemen badan usaha dalam perekonomian Indonesia',
    batasanCatatan: 'Peran dan prinsip pengelolaan BUMN, BUMD, BUMS, serta struktur kepengurusan dan SHU Koperasi.'
  },
  {
    elemenMateri: 'Ekonomi Internasional',
    subElemenMateri: 'Kerja sama ekonomi dan perdagangan internasional',
    kompetensi: 'Menganalisis kerja sama ekonomi dan perdagangan internasional',
    batasanCatatan: 'Teori perdagangan internasional (keunggulan mutlak/komparatif), neraca pembayaran, dan organisasi kerja sama ekonomi regional/global.'
  },
  {
    elemenMateri: 'Akuntansi Keuangan Dasar',
    subElemenMateri: 'Persamaan dasar akuntansi dan laporan keuangan',
    kompetensi: 'Menerapkan persamaan dasar akuntansi dan laporan keuangan',
    batasanCatatan: 'Analisis transaksi keuangan, pencatatan persamaan dasar akuntansi, serta penyusunan laporan laba rugi, perubahan modal, dan neraca.'
  }
];

const PUSMENDIK_GEOGRAFI_PRESETS = [
  {
    elemenMateri: 'Wilayah tempat tinggal dan lingkungan sekitar (karakteristik, keunikan, persamaan– perbedaan wilayah)',
    subElemenMateri: 'Karakteristik fisik dan sosial wilayah',
    kompetensi: 'Menjelaskan karakteristik wilayah berdasarkan informasi spasial.',
    batasanCatatan: 'Karakteristik fisik dan sosial wilayah berdasarkan peta/data spasial.'
  },
  {
    elemenMateri: 'Wilayah tempat tinggal dan lingkungan sekitar (karakteristik, keunikan, persamaan– perbedaan wilayah)',
    subElemenMateri: 'Konsep dan teori dasar mengenai dinamika kependudukan pada suatu wilayah tempat tinggal',
    kompetensi: 'Menganalisis konsep dan teori dasar mengenai dinamika kependudukan pada suatu wilayah tempat tinggal',
    batasanCatatan: 'Konsep dan teori dasar dinamika kependudukan di wilayah tempat tinggal.'
  },
  {
    elemenMateri: 'Wilayah tempat tinggal dan lingkungan sekitar (karakteristik, keunikan, persamaan– perbedaan wilayah)',
    subElemenMateri: 'Keterkaitan antara karakteristik wilayah fisik dan sosial (kependudukan)',
    kompetensi: 'Menganalisis keterkaitan antara karakteristik wilayah fisik dan sosial (kependudukan) terhadap daya dukung lingkungan;',
    batasanCatatan: 'Hubungan karakteristik wilayah fisik & sosial terhadap daya dukung lingkungan.'
  },
  {
    elemenMateri: 'Wilayah tempat tinggal dan lingkungan sekitar (karakteristik, keunikan, persamaan– perbedaan wilayah)',
    subElemenMateri: 'Permasalahan kewilayahan',
    kompetensi: 'Menerapkan penelitian geografi untuk memecahkan permasalahan kewilayahan',
    batasanCatatan: 'Penelitian geografi untuk pemecahan masalah kewilayahan.'
  },
  {
    elemenMateri: 'Proses yang memengaruhi lingkungan fisik dan sosial',
    subElemenMateri: 'Faktor yang berpengaruh dalam lingkungan sosial',
    kompetensi: 'Menerapkan indikator- indikator keberhasilan Pembangunan untuk pengembangan wilayah',
    batasanCatatan: 'Indikator keberhasilan pembangunan wilayah.'
  },
  {
    elemenMateri: 'Proses yang memengaruhi lingkungan fisik dan sosial',
    subElemenMateri: 'Faktor yang berpengaruh dalam lingkungan sosial',
    kompetensi: 'Menganalisis konsep dinamika kependudukan dan faktor-faktor yang mempengaruhinya',
    batasanCatatan: 'Dinamika kependudukan dan faktor-faktor pengaruhnya.'
  },
  {
    elemenMateri: 'Proses yang memengaruhi lingkungan fisik dan sosial',
    subElemenMateri: 'Proses/fenomena yang memengaruhi lingkungan fisik',
    kompetensi: 'Menggunakan informasi tentang proses alam untuk menjelaskan perubahan lingkungan fisik di wilayah tertentu.',
    batasanCatatan: 'Proses alam dan perubahan lingkungan fisik wilayah tertentu.'
  },
  {
    elemenMateri: 'Proses yang memengaruhi lingkungan fisik dan sosial',
    subElemenMateri: 'Proses/fenomena yang memengaruhi lingkungan fisik',
    kompetensi: 'Menganalisis peranan manusia dalam mengubah lingkungan fisik',
    batasanCatatan: 'Peran manusia dalam mengubah lingkungan fisik.'
  },
  {
    elemenMateri: 'Proses yang memengaruhi lingkungan fisik dan sosial',
    subElemenMateri: 'Proses/fenomena yang memengaruhi lingkungan fisik',
    kompetensi: 'Menganalisis persebaran bioma di dunia dan pengaruhnya terhadap manusia.',
    batasanCatatan: 'Persebaran bioma di dunia dan dampaknya terhadap kehidupan manusia.'
  },
  {
    elemenMateri: 'Interaksi antargejala fisik alam dan manusia dan pengaruhnya terhadap kehidupan',
    subElemenMateri: 'Posisi strategis Indonesia dan pengaruhnya bagi kehidupan ekonomi, sosial, budaya secara nasional maupun internasional',
    kompetensi: 'Menjelaskan posisi geografis Indonesia.',
    batasanCatatan: 'Letak geografis Indonesia dan dampaknya bagi kehidupan.'
  },
  {
    elemenMateri: 'Interaksi antargejala fisik alam dan manusia dan pengaruhnya terhadap kehidupan',
    subElemenMateri: 'Potensi Sumber Daya Alam Indonesia terhadap dinamika kehidupan',
    kompetensi: 'Menganalisis pemanfaatan SDA sesuai konteks wilayah dengan menggunakan informasi (peta, data)',
    batasanCatatan: 'Pemanfaatan SDA berdasarkan informasi peta atau data wilayah.'
  },
  {
    elemenMateri: 'Interaksi antargejala fisik alam dan manusia dan pengaruhnya terhadap kehidupan',
    subElemenMateri: 'Potensi Sumber Daya Alam Indonesia terhadap dinamika kehidupan',
    kompetensi: 'Menganalisis pengelolaan SDA secara berkelanjutan',
    batasanCatatan: 'Pengelolaan SDA berkelanjutan di Indonesia.'
  },
  {
    elemenMateri: 'Cara mitigas dan adaptasi terhadap bencana alam di lingkungan tempat tinggal dan negaranya.',
    subElemenMateri: 'Bencana geologis/ hidroklimatologis',
    kompetensi: 'Menganalisis ragam risiko dan faktor penyebab bencana alam.',
    batasanCatatan: 'Risiko dan faktor penyebab bencana geologis/hidroklimatologis.'
  },
  {
    elemenMateri: 'Cara mitigas dan adaptasi terhadap bencana alam di lingkungan tempat tinggal dan negaranya.',
    subElemenMateri: 'Mitigasi dan adaptasi manusia terhadap bencana geologis/ hidroklimatologis',
    kompetensi: 'Menggunakan data atau studi kasus untuk menjelaskan bentuk adaptasi masyarakat terhadap bencana geologis atau hidroklimatologis di suatu wilayah.',
    batasanCatatan: 'Studi kasus adaptasi masyarakat terhadap bencana.'
  },
  {
    elemenMateri: 'Cara mitigas dan adaptasi terhadap bencana alam di lingkungan tempat tinggal dan negaranya.',
    subElemenMateri: 'Mitigasi dan adaptasi manusia terhadap bencana geologis/ hidroklimatologis',
    kompetensi: 'Mengevaluasi upaya pengurangan risiko bencana geologis/ hidroklimatologis.',
    batasanCatatan: 'Evaluasi mitigasi/pengurangan risiko bencana alam.'
  },
  {
    elemenMateri: 'Fenomena geografi dalam kehidupan sehari-hari dan manfaatnya',
    subElemenMateri: 'Peta, penginderaan jauh dan SIG (Sistem Informasi Geografis)',
    kompetensi: 'Menjelaskan penggunaan informasi Geospasial dalam kehidupan sehari-hari',
    batasanCatatan: 'Penggunaan informasi geospasial dalam kehidupan sehari-hari.'
  },
  {
    elemenMateri: 'Fenomena geografi dalam kehidupan sehari-hari dan manfaatnya',
    subElemenMateri: 'Peta, penginderaan jauh dan SIG (Sistem Informasi Geografis)',
    kompetensi: 'Menganalisis fenomena geosfer dari peta/citra penginderaan jauh.',
    batasanCatatan: 'Analisis fenomena geosfer melalui peta/citra penginderaan jauh.'
  }
];

const PUSMENDIK_SOSIOLOGI_PRESETS = [
  {
    elemenMateri: 'Sosiologi sebagai Ilmu',
    subElemenMateri: 'Pengertian dan perkembangan sosiologi dan manfaat sosiologi dalam kehidupan masyarakat.',
    kompetensi: 'Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.',
    batasanCatatan: 'Sejarah sosiologi, objek kajian sosiologi, fungsi dan manfaat sosiologi bagi masyarakat.'
  },
  {
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Konsep dan bentuk hubungan sosial',
    kompetensi: 'Mengidentifikasi dan menganalisis konsep dan bentuk hubungan sosial yang terjadi di masyarakat',
    batasanCatatan: 'Interaksi sosial, syarat-syarat interaksi, dan bentuk interaksi (asosiatif dan disosiatif).'
  },
  {
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Pembentukan kepribadian, kelompok dan lembaga sosial.',
    kompetensi: 'Mengidentifikasi berbagai lembaga sosial dan perannya di masyarakat.',
    batasanCatatan: 'Proses sosialisasi, pembentukan kepribadian, jenis kelompok, serta peran lembaga sosial (keluarga, agama, ekonomi).'
  },
  {
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Ragam gejala sosial.',
    kompetensi: 'Menjelaskan ragam gejala sosial di lingkungan sekitar.',
    batasanCatatan: 'Perilaku menyimpang, masalah sosial, sosiologi perkotaan/pedesaan, dan dampaknya bagi keteraturan sosial.'
  },
  {
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Masyarakat multikultural.',
    kompetensi: 'Menganalisis dinamika masyarakat multikultural.',
    batasanCatatan: 'Keragaman ras, etnis, agama, serta toleransi dan integrasi sosial dalam kerangka multikultural.'
  },
  {
    elemenMateri: 'Penelitian Sosial',
    subElemenMateri: 'Langkah penelitian sosial dan metode penelitian.',
    kompetensi: 'Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.',
    batasanCatatan: 'Rancangan penelitian, jenis penelitian (kualitatif/kuantitatif), teknik sampling, pengumpulan data, dan penyusunan laporan.'
  },
  {
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Konsep Kelompok Sosial dan dinamika Kelompok Sosial.',
    kompetensi: 'Mengidentifikasi, menjelaskan, dan menganalisis berbagai kelompok sosial dengan dinamikanya.',
    batasanCatatan: 'Klasifikasi kelompok sosial (ingroup/outgroup, paguyuban/patembayan), dinamika dan perkembangan kelompok.'
  },
  {
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Ketidaksetaraan sosial dan upaya mewujudkan kesetaraan sosial.',
    kompetensi: 'Memahami faktor yang memengaruhi ketidaksetaraan sosial dan menganalisis upaya mewujudkan kesetaraan sosial.',
    batasanCatatan: 'Stratifikasi sosial, diferensiasi sosial, ketimpangan sosial, serta harmoni sosial.'
  },
  {
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Konflik sosial dan penanganan konflik.',
    kompetensi: 'Mendeskripsikan berbagai konsep konflik sosial dan menganalisis berbagai upaya penanganan konflik.',
    batasanCatatan: 'Penyebab konflik, bentuk konflik, kekerasan, serta resolusi konflik (akomodasi, negosiasi, mediasi, arbitrase).'
  },
  {
    elemenMateri: 'Perubahan Sosial dan Globalisasi',
    subElemenMateri: 'Bentuk-bentuk perubahan sosial dan dampaknya.',
    kompetensi: 'Mengidentifikasi bentuk-bentuk perubahan sosial dan menganalisis dampak perubahan sosial.',
    batasanCatatan: 'Faktor pendorong/penghambat perubahan, teori perubahan sosial, modernisasi, dan desintegrasi.'
  },
  {
    elemenMateri: 'Perubahan Sosial dan Globalisasi',
    subElemenMateri: 'Globalisasi dan dampak globalisasi.',
    kompetensi: 'Menjelaskan dan menganalisis pengaruh globalisasi dan dampaknya.',
    batasanCatatan: 'Globalisasi ekonomi, politik, budaya, serta ketimpangan global dan lokalisasi (glokalisasi).'
  },
  {
    elemenMateri: 'Perubahan Sosial dan Globalisasi',
    subElemenMateri: 'Sikap kritis globalisasi.',
    kompetensi: 'Menganalisis fenomena sosial yang dipengaruhi globalisasi secara kritis.',
    batasanCatatan: 'Respon terhadap tantangan globalisasi, penguatan kearifan lokal, dan pemberdayaan komunitas.'
  }
];

const PUSMENDIK_SEJARAH_TL_PRESETS = [
  {
    elemenMateri: 'Pengantar Ilmu Sejarah',
    subElemenMateri: 'Konsep Dasar Sejarah',
    kompetensi: 'Menjelaskan dan Menganalisis konsep perubahan, keberlanjutan, serta sebab-akibat untuk memahami pengaruh peristiwa sejarah terhadap fenomena sosial yang dialami murid.',
    batasanCatatan: 'Konsep perubahan, keberlanjutan, serta sebab-akibat peristiwa sejarah terhadap kehidupan sosial murid.'
  },
  {
    elemenMateri: 'Pengantar Ilmu Sejarah',
    subElemenMateri: 'Fenomena sejarah dalam kehidupan sehari-hari',
    kompetensi: 'Menganalisis keterkaitan antara peristiwa sejarah masa lalu dan dinamika sosial budaya masyarakat masa kini dengan menggunakan prinsip perubahan, keberlanjutan, dan sebab-akibat.',
    batasanCatatan: 'Keterkaitan peristiwa masa lalu dengan dinamika sosial budaya masyarakat masa kini.'
  },
  {
    elemenMateri: 'Pengantar Ilmu Sejarah',
    subElemenMateri: 'Sumber-sumber sejarah',
    kompetensi: 'Mengidentifikasi fungsi dan perbedaan antara sumber sejarah primer dan sekunder serta menganalisis penggunaannya untuk merekonstruksi peristiwa masa lalu secara kontekstual.',
    batasanCatatan: 'Fungsi dan perbedaan sumber sejarah primer dan sekunder untuk rekonstruksi sejarah.'
  },
  {
    elemenMateri: 'Periode Kerajaan Hindu-Buddha dan Islam',
    subElemenMateri: 'Kehidupan religi, budaya, politik, dan ekonomi masyarakat di Nusantara pada masa kerajaan Hindu-Buddha',
    kompetensi: 'Mengevaluasi teori masuknya agama dan kebudayaan Hindu-Buddha ke Nusantara berdasarkan konsep dasar sejarah, menganalisis kehidupan politik dan ekonomi kerajaan-kerajaan Hindu-Buddha, serta mengklasifikasikan peninggalan budaya yang dihasilkannya.',
    batasanCatatan: 'Teori masuknya Hindu-Buddha, kehidupan politik/ekonomi kerajaan, dan peninggalan budayanya.'
  },
  {
    elemenMateri: 'Periode Kerajaan Hindu-Buddha dan Islam',
    subElemenMateri: 'Kehidupan religi, budaya, politik, dan ekonomi masyarakat di Nusantara pada masa kerajaan Islam',
    kompetensi: 'Menganalisis hubungan antara masuknya agama dan kebudayaan Islam dengan perubahan dalam sistem politik, ekonomi, serta perkembangan budaya masyarakat Nusantara.',
    batasanCatatan: 'Masuknya Islam, hubungannya dengan perubahan sistem politik, ekonomi, dan budaya Nusantara.'
  },
  {
    elemenMateri: 'Perlawanan terhadap Bangsa Eropa',
    subElemenMateri: 'Proses kedatangan Bangsa Eropa ke Nusantara',
    kompetensi: 'Menjelaskan dan menganalisis keterkaitan antara motivasi kedatangan bangsa Eropa dan perubahan sosial, ekonomi, serta politik di Nusantara dengan menggunakan pendekatan diakronik dan sinkronik.',
    batasanCatatan: 'Motivasi kedatangan bangsa Eropa dan perubahan sosial, ekonomi, serta politik Nusantara.'
  },
  {
    elemenMateri: 'Perlawanan terhadap Bangsa Eropa',
    subElemenMateri: 'Perlawanan terhadap Bangsa Eropa sebelum Abad ke-20',
    kompetensi: 'Menganalisis keterkaitan antara kebijakan kolonial bangsa Eropa dan munculnya berbagai bentuk perlawanan rakyat Nusantara sebelum abad ke-20, serta mengevaluasi strategi perjuangan yang dilakukan.',
    batasanCatatan: 'Hubungan kebijakan kolonial dengan perlawanan rakyat Nusantara dan strategi perjuangannya.'
  },
  {
    elemenMateri: 'Pergerakan Nasional sampai Proklamasi Kemerdekaan',
    subElemenMateri: 'Pergerakan Nasional Indonesia',
    kompetensi: 'Mengidentifikasi dampak Politik Etis dan munculnya berbagai organisasi pada masa Pergerakan Nasional serta menganalisis strategi perlawanan yang dilakukan dalam berbagai bidang.',
    batasanCatatan: 'Politik Etis, kemunculan organisasi pergerakan nasional, serta strategi perlawanan.'
  },
  {
    elemenMateri: 'Pergerakan Nasional sampai Proklamasi Kemerdekaan',
    subElemenMateri: 'Relevansi semangat Pergerakan Nasional dengan masa kini',
    kompetensi: 'Menganalisis relevansi nilai dan semangat perjuangan tokoh-tokoh pergerakan nasional dalam menghadapi tantangan kehidupan berbangsa dan bernegara di masa kini.',
    batasanCatatan: 'Nilai perjuangan tokoh pergerakan nasional dan relevansinya di masa kini.'
  },
  {
    elemenMateri: 'Pergerakan Nasional sampai Proklamasi Kemerdekaan',
    subElemenMateri: 'Kehidupan Bangsa Indonesia di bawah penjajahan Jepang dan perlawanan Bangsa Indonesia',
    kompetensi: 'Mengevaluasi beberapa penjelasan tentang penyebab utama runtuhnya kekuasaan Belanda di Indonesia sebelum pendudukan Jepang, berdasarkan sumber sejarah yang relevan.',
    batasanCatatan: 'Penyebab runtuhnya kekuasaan Belanda di Indonesia sebelum Jepang.'
  },
  {
    elemenMateri: 'Pergerakan Nasional sampai Proklamasi Kemerdekaan',
    subElemenMateri: 'Kehidupan Bangsa Indonesia di bawah penjajahan Jepang dan perlawanan Bangsa Indonesia',
    kompetensi: 'Menganalisis dampak kebijakan pendudukan Jepang di bidang politik, ekonomi, sosial, dan budaya terhadap kehidupan masyarakat Indonesia.',
    batasanCatatan: 'Dampak kebijakan pendudukan Jepang dalam berbagai bidang kehidupan.'
  },
  {
    elemenMateri: 'Pergerakan Nasional sampai Proklamasi Kemerdekaan',
    subElemenMateri: 'Peristiwa dan Makna Proklamasi Kemerdekaan Indonesia',
    kompetensi: 'Menjelaskan dan menganalisis peristiwa pada masa proklamasi kemerdekaan Indonesia dan maknanya.',
    batasanCatatan: 'Peristiwa sekitar proklamasi kemerdekaan Indonesia dan makna historisnya.'
  },
  {
    elemenMateri: 'Revolusi Kemerdekaan Indonesia sampai Demokrasi Terpimpin',
    subElemenMateri: 'Perjuangan mempertahankan kemerdekaan',
    kompetensi: 'Mendeskripsikan proses pembentukan negara dan pemerintahan Republik Indonesia setelah Proklamasi Kemerdekaan, serta menganalisis upaya perjuangan bangsa Indonesia dalam mempertahankan kemerdekaan melalui jalur diplomasi dan perjuangan fisik.',
    batasanCatatan: 'Pembentukan negara dan perjuangan fisik serta diplomasi mempertahankan kemerdekaan.'
  },
  {
    elemenMateri: 'Revolusi Kemerdekaan Indonesia sampai Demokrasi Terpimpin',
    subElemenMateri: 'Kehidupan masyarakat Indonesia pada masa Demokrasi Liberal',
    kompetensi: 'Menggunakan konsep Kronologis untuk menjelaskan perkembangan politik dan ekonomi Indonesia pada masa Demokrasi Liberal berdasarkan peristiwa peristiwa penting.',
    batasanCatatan: 'Perkembangan politik dan ekonomi Indonesia masa Demokrasi Liberal secara kronologis.'
  },
  {
    elemenMateri: 'Revolusi Kemerdekaan Indonesia sampai Demokrasi Terpimpin',
    subElemenMateri: 'Kehidupan masyarakat pada masa Demokrasi Terpimpin',
    kompetensi: 'Mengidentifikasi perkembangan politik dan ekonomi serta menganalisis dampak kebijakan pemerintah pada masa Demokrasi Terpimpin.',
    batasanCatatan: 'Perkembangan politik-ekonomi dan dampak kebijakan pemerintah masa Demokrasi Terpimpin.'
  },
  {
    elemenMateri: 'Orde Baru sampai Reformasi',
    subElemenMateri: 'Kehidupan Masyarakat pada Masa Orde Baru',
    kompetensi: 'Menganalisis kronologi perubahan Demokrasi Terpimpin menjadi Orde Baru dengan menggunakan konsep sejarah serta mengevaluasi dampak kebijakan politik dan ekonomi Orde Baru terhadap kehidupan masyarakat Indonesia.',
    batasanCatatan: 'Peralihan kekuasaan, kebijakan politik-ekonomi Orde Baru, serta dampaknya.'
  },
  {
    elemenMateri: 'Orde Baru sampai Reformasi',
    subElemenMateri: 'Kehidupan Masyarakat pada Masa Reformasi',
    kompetensi: 'Menganalisis proses lahirnya Reformasi dan peran pelajar serta mahasiswa sebagai pelaku sejarah, dengan menggunakan konsep perubahan dan kronologi, serta mengevaluasi dampak kebijakan politik and ekonomi Reformasi terhadap kehidupan masyarakat Indonesia.',
    batasanCatatan: 'Lahirnya Reformasi, peran gerakan mahasiswa/pelajar, serta dampak kebijakan Reformasi.'
  }
];

const PUSMENDIK_ANTROPOLOGI_PRESETS = [
  {
    elemenMateri: 'Pengantar dan Ruang Lingkup Antropologi',
    subElemenMateri: 'Konsep Dasar dan Sejarah Perkembangan Antropologi',
    kompetensi: 'Mendeskripsikan dan menganalisis konsep dasar dan sejarah perkembangan Antropologi',
    batasanCatatan: 'Konsep dasar dan sejarah perkembangan ilmu Antropologi.'
  },
  {
    elemenMateri: 'Pengantar dan Ruang Lingkup Antropologi',
    subElemenMateri: 'Prinsip Dasar dan Pendekatan Antropologi',
    kompetensi: 'Mengidentifikasi, dan menganalisis prinsip dasar Antropologi untuk menjelaskan fenomena sosial budaya di masyarakat',
    batasanCatatan: 'Prinsip dasar dan pendekatan Antropologi dalam fenomena sosial budaya.'
  },
  {
    elemenMateri: 'Pengantar dan Ruang Lingkup Antropologi',
    subElemenMateri: 'Ruang Lingkup Antropologi (Antropologi Ragawi, Arkeologi, dan Etnologi Bahasa)',
    kompetensi: 'Menjelaskan dan menganalisis keterkaitan antara cabang-cabang antropologi untuk memahami keberagaman budaya dan dinamika masyarakat manusia.',
    batasanCatatan: 'Keterkaitan cabang-cabang antropologi (ragawi, arkeologi, etnologi bahasa).'
  },
  {
    elemenMateri: 'Etnografi',
    subElemenMateri: 'Pengertian Etnografi',
    kompetensi: 'Menjelaskan konsep etnografi sebagai metode dalam penelitian Antropologi.',
    batasanCatatan: 'Konsep etnografi sebagai metode khas penelitian Antropologi.'
  },
  {
    elemenMateri: 'Etnografi',
    subElemenMateri: 'Metode dan Proses Penelitian Etnografi',
    kompetensi: 'Menjelaskandan menganalisis berbagai jenis metode penelitian Etnografi',
    batasanCatatan: 'Metode, teknik pengumpulan data, dan proses penelitian Etnografi.'
  },
  {
    elemenMateri: 'Etnografi',
    subElemenMateri: 'Pemanfaatan Hasil Penelitian Etnografi secara Kritis',
    kompetensi: 'Menganalisis pemanfaatan hasil penelitian etnografi dalam memahami dinamika sosial budaya secaraskritis dan kontekstual.',
    batasanCatatan: 'Pemanfaatan hasil penelitian etnografi secara kritis dan kontekstual.'
  },
  {
    elemenMateri: 'Etnografi',
    subElemenMateri: 'Penerapan Etnografi dalam Kehidupan Sehari-Hari',
    kompetensi: 'Menjelaskan penerapan etnografi dalam kehidupan sehari-hari.',
    batasanCatatan: 'Penerapan konsep dan metode etnografi dalam konteks sehari-hari.'
  },
  {
    elemenMateri: 'Masyarakat Multikultural',
    subElemenMateri: 'Jenis-Jenis Multikulturalisme dalam Masyarakat',
    kompetensi: 'Mengidentifikasi dan menganalisis keberagaman masyarakat Indonesia sebagai masyarakat multikultural',
    batasanCatatan: 'Keberagaman Indonesia sebagai masyarakat multikultural.'
  },
  {
    elemenMateri: 'Masyarakat Multikultural',
    subElemenMateri: 'Jenis-Jenis Multikulturalisme dalam Masyarakat',
    kompetensi: 'Mendeskripsikan berbagai konsep masyarakat multikultural dan menjelaskan jenis-jenis multikulturalisme dalam masyarakat',
    batasanCatatan: 'Konsep dan jenis-jenis multikulturalisme.'
  },
  {
    elemenMateri: 'Masyarakat Multikultural',
    subElemenMateri: 'Masyarakat Multikultural di Indonesia dan Global serta Tantangan dan Peluangnya.',
    kompetensi: 'Mendeskripsikan berbagai konsep masyarakat multikultural dan menjelaskan jenis-jenis multikulturalisme dalam masyarakat',
    batasanCatatan: 'Kondisi masyarakat multikultural lokal/global serta peluang & tantangannya.'
  },
  {
    elemenMateri: 'Perubahan Sosial Budaya',
    subElemenMateri: 'Konsep, Bentuk, dan Faktor Perubahan Sosial dan Budaya',
    kompetensi: 'Mendeskripsikan serta menganalisis konsep, bentuk, dan faktor perubahan soisal dan budaya',
    batasanCatatan: 'Konsep, bentuk, dan faktor pendorong/penghambat perubahan sosial budaya.'
  },
  {
    elemenMateri: 'Perubahan Sosial Budaya',
    subElemenMateri: 'Dampak dan Respon Masyarakat terhadap Perubahan Sosial dan Budaya di Indonesia',
    kompetensi: 'Menjelaskan dan menganalisis dampak dan respon masyarakat terhadap perubahan sosial dan budaya di Indonesia',
    batasanCatatan: 'Dampak dan respon masyarakat Indonesia terhadap perubahan sosial budaya.'
  },
  {
    elemenMateri: 'Antropologi Sosial dan Antropologi Budaya',
    subElemenMateri: 'Pengertian dan Cakupan Antropologi Sosial dan Antropologi Budaya',
    kompetensi: 'Memahami ruang lingkup kajian antropologi sosial dan antropologi budaya',
    batasanCatatan: 'Ruang lingkup dan kajian antropologi sosial & budaya.'
  },
  {
    elemenMateri: 'Antropologi Sosial dan Antropologi Budaya',
    subElemenMateri: 'Antropologi Sosial dan Antropologi Budaya sebagai Antropologi Terapan serta Studi Kasusnya di Masyarakat',
    kompetensi: 'Menjelaskan dan Menganalisis berbagai kajian antropologi sosial budaya sebagai antropologi terapan dalam sistem sosial budaya masyarakat',
    batasanCatatan: 'Kajian antropologi sosial budaya sebagai antropologi terapan.'
  },
  {
    elemenMateri: 'Antropologi Sosial dan Antropologi Budaya',
    subElemenMateri: 'Antropologi Sosial dan Antropologi Budaya sebagai Antropologi Terapan serta Studi Kasusnya di Masyarakat',
    kompetensi: 'Menganalisis studi kasus antropologi terapan dalam sistem sosial budaya masyarakat untuk memahami dinamika atau solusi terhadap permasalahan budaya.',
    batasanCatatan: 'Studi kasus antropologi terapan untuk pemecahan masalah budaya.'
  },
  {
    elemenMateri: 'Kearifan Lokal dan Tradisi Lisan',
    subElemenMateri: 'Definisi dan Bentuk-Bentuk Kearifan Lokal dalam Antropologi',
    kompetensi: 'Memahami definisi dan bentuk-bentuk kearifan lokal dalam Antropologi',
    batasanCatatan: 'Definisi, karakteristik, dan ragam bentuk kearifan lokal.'
  },
  {
    elemenMateri: 'Kearifan Lokal dan Tradisi Lisan',
    subElemenMateri: 'Peran Kearifan Lokal dalam Kehidupan Masyarakat',
    kompetensi: 'Menjelaskan dan menganalisis peran kearifan lokal dalam kehidupan masyarakat',
    batasanCatatan: 'Peran dan nilai kearifan lokal dalam kelangsungan masyarakat.'
  },
  {
    elemenMateri: 'Kearifan Lokal dan Tradisi Lisan',
    subElemenMateri: 'Jenis-Jenis dan Fungsi Tradisi Lisan dalam Masyarakat',
    kompetensi: 'Menjelaskan dan menganalisis Jenis-jenis dan fungsi tradisi lisan dalam masyarakat',
    batasanCatatan: 'Jenis-jenis (mitos, legenda, dongeng, puisi rakyat) dan fungsi tradisi lisan.'
  },
  {
    elemenMateri: 'Kearifan Lokal dan Tradisi Lisan',
    subElemenMateri: 'Tantangan Kearifan Lokal dan Tradisi Lisan di Era Modern',
    kompetensi: 'Menjelaskan dan menganalisis tantangan kearifan lokal dan tradisi lisan di era modern',
    batasanCatatan: 'Tantangan eksistensi kearifan lokal dan tradisi lisan di era modern/digital.'
  }
];

const PUSMENDIK_BAHASA_JEPANG_PRESETS = [
  {
    elemenMateri: 'Pemahaman Literal',
    subElemenMateri: 'Menemukan Informasi Tersurat',
    kompetensi: 'Menemukan informasi tersurat dari gambar atau teks sederhana.',
    batasanCatatan: 'Mengidentifikasi informasi tertulis/gambar yang disajikan secara eksplisit dalam bahasa Jepang.'
  },
  {
    elemenMateri: 'Pemahaman Literal',
    subElemenMateri: 'Melengkapi Teks Sederhana',
    kompetensi: 'Melengkapi teks dengan kosakata dan ungkapan komunikatif sesuai topik dari isi teks sederhana.',
    batasanCatatan: 'Mengisi bagian rumpang dengan partikel dasar, kata kerja, kata sifat, atau ungkapan komunikatif harian.'
  },
  {
    elemenMateri: 'Reorganisasi',
    subElemenMateri: 'Struktur Kalimat Bahasa Jepang',
    kompetensi: 'Menyusun kata-kata menjadi kalimat utuh sesuai struktur Bahasa Jepang.',
    batasanCatatan: 'Menata pola kalimat dasar Jepang (tata kalimat subjek-objek-predikat, partikel wa, ga, o, ni, de, dll).'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Menyimpulkan Informasi Tersirat',
    kompetensi: 'Menyimpulkan informasi tersirat dari teks sederhana.',
    batasanCatatan: 'Memahami maksud tersirat, penokohan, latar belakang situasi, atau kesimpulan dari suatu teks.'
  },
  {
    elemenMateri: 'Pemahaman Inferensial',
    subElemenMateri: 'Aplikasi Tata Bahasa (Negasi, Lampau, Akan Datang)',
    kompetensi: 'Mengaplikasikan Pengunaan tata bahasa bentuk negasi  lampau, atau akan datang pada teks sederhana',
    batasanCatatan: 'Konjugasi kata kerja/sifat bentuk positif, negatif, lampau, maupun bentuk kamus/akan datang.'
  }
];

const PUSMENDIK_PKK_PRESETS = [
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pengembangan Desain Produk',
    kompetensi: 'Menganalisis desain dan prosedur pengembangan produk.',
    batasanCatatan: 'Analisis konsep desain, estetika, fungsi, dan prosedur sistematis dalam mengembangkan produk baru.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pengembangan Desain Kemasan Produk',
    kompetensi: 'Mengevaluasi desain kemasan dan label produk.',
    batasanCatatan: 'Evaluasi kesesuaian kemasan, daya tarik visual, informasi label, pelindungan produk, dan aspek regulasi/ramah lingkungan.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pengembangan Prototipe Produk',
    kompetensi: 'Menerapkan pengembangan prototipe produk.',
    batasanCatatan: 'Penerapan tahapan pembuatan, pengujian, dan penyempurnaan prototipe (model fisik awal) sebelum diproduksi massal.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Perencanaan Produksi',
    kompetensi: 'Menentukan perencanaan dan biaya produksi.',
    batasanCatatan: 'Penghitungan harga pokok produksi (HPP), break-even point (BEP), perencanaan kebutuhan bahan, tenaga kerja, dan penjadwalan.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Proses Produksi',
    kompetensi: 'Menerapkan proses produksi.',
    batasanCatatan: 'Penerapan tahapan, teknik, dan prosedur pembuatan produk secara efektif, efisien, dan aman sesuai standar.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pengemasan Produk',
    kompetensi: 'Menerapkan pengemasan produk.',
    batasanCatatan: 'Penerapan metode pengemasan yang tepat untuk mempertahankan kualitas, keamanan, serta estetika visual produk.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pengendalian Mutu Produk (Quality Assurance)',
    kompetensi: 'Menerapkan pengendalian mutu produk.',
    batasanCatatan: 'Penerapan standar kualitas, inspeksi kelayakan, pengujian produk, serta tindakan koreksi kegagalan mutu.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Pemasaran Produk',
    kompetensi: 'Mengevaluasi strategi dan pemasaran produk.',
    batasanCatatan: 'Evaluasi bauran pemasaran (marketing mix), segmentasi pasar, promosi digital/tradisional, serta pencapaian target penjualan.'
  },
  {
    elemenMateri: 'Kegiatan Produksi, Pemasaran, dan Distribusi',
    subElemenMateri: 'Distribusi Produk',
    kompetensi: 'Menerapkan distribusi produk.',
    batasanCatatan: 'Penerapan saluran distribusi (langsung/tidak langsung), logistik, serta rantai pasok agar produk sampai ke tangan konsumen tepat waktu.'
  },
  {
    elemenMateri: 'Pengelolaan Usaha',
    subElemenMateri: 'Analisis Peluang Usaha',
    kompetensi: 'Menentukan peluang usaha.',
    batasanCatatan: 'Identifikasi potensi pasar, analisis SWOT (Strengths, Weaknesses, Opportunities, Threats), serta kelayakan bisnis.'
  },
  {
    elemenMateri: 'Pengelolaan Usaha',
    subElemenMateri: 'Proposal Usaha',
    kompetensi: 'Menganalisis proposal usaha.',
    batasanCatatan: 'Analisis kerangka, komponen kelayakan usaha, strategi pendanaan, serta penyusunan proposal usaha yang profesional.'
  },
  {
    elemenMateri: 'Pengelolaan Usaha',
    subElemenMateri: 'Pelaporan Keuangan',
    kompetensi: 'Menganalisis laporan keuangan.',
    batasanCatatan: 'Analisis neraca, laporan laba rugi, laporan arus kas, serta interpretasi kinerja keuangan usaha mikro/menengah.'
  },
  {
    elemenMateri: 'Pengelolaan Usaha',
    subElemenMateri: 'HaKI',
    kompetensi: 'Mengevaluasi HaKI.',
    batasanCatatan: 'Evaluasi hak atas kekayaan intelektual (paten, merek, hak cipta, desain industri) untuk perlindungan hukum produk kreatif.'
  }
];

const DEFAULT_JADWAL_LIST: JadwalItem[] = [
  {
    id: 'jadwal-1',
    bulan: 'Juli',
    mingguKe: 1,
    elemenMateri: 'Sosiologi sebagai Ilmu',
    subElemenMateri: 'Pengertian dan perkembangan sosiologi dan manfaat sosiologi dalam kehidupan masyarakat.',
    kompetensi: 'Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.'
  },
  {
    id: 'jadwal-2',
    bulan: 'Juli',
    mingguKe: 2,
    elemenMateri: 'Sosiologi sebagai Ilmu',
    subElemenMateri: 'Objek kajian sosiologi, fungsi dan manfaat sosiologi bagi masyarakat.',
    kompetensi: 'Menganalisis objek kajian sosiologi serta fungsinya dalam memecahkan masalah sosial.'
  },
  {
    id: 'jadwal-3',
    bulan: 'Juli',
    mingguKe: 3,
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Konsep dan bentuk hubungan sosial',
    kompetensi: 'Mengidentifikasi dan menganalisis konsep dan bentuk hubungan sosial yang terjadi di masyarakat'
  },
  {
    id: 'jadwal-4',
    bulan: 'Juli',
    mingguKe: 4,
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Pembentukan kepribadian, kelompok dan lembaga sosial.',
    kompetensi: 'Mengidentifikasi berbagai lembaga sosial dan perannya di masyarakat.'
  },
  {
    id: 'jadwal-5',
    bulan: 'Agustus',
    mingguKe: 1,
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Ragam gejala sosial.',
    kompetensi: 'Menjelaskan ragam gejala sosial di lingkungan sekitar.'
  },
  {
    id: 'jadwal-6',
    bulan: 'Agustus',
    mingguKe: 2,
    elemenMateri: 'Hubungan dan Gejala Sosial',
    subElemenMateri: 'Masyarakat multikultural.',
    kompetensi: 'Menganalisis dinamika masyarakat multikultural.'
  },
  {
    id: 'jadwal-7',
    bulan: 'Agustus',
    mingguKe: 3,
    elemenMateri: 'Penelitian Sosial',
    subElemenMateri: 'Langkah penelitian sosial dan metode penelitian.',
    kompetensi: 'Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.'
  },
  {
    id: 'jadwal-8',
    bulan: 'Agustus',
    mingguKe: 4,
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Konsep Kelompok Sosial dan dinamika Kelompok Sosial.',
    kompetensi: 'Mengidentifikasi, menjelaskan, dan menganalisis berbagai kelompok sosial dengan dinamikanya.'
  },
  {
    id: 'jadwal-9',
    bulan: 'Oktober',
    mingguKe: 1,
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Ketidaksetaraan sosial dan upaya mewujudkan kesetaraan sosial.',
    kompetensi: 'Memahami faktor yang memengaruhi ketidaksetaraan sosial dan menganalisis upaya mewujudkan kesetaraan sosial.'
  },
  {
    id: 'jadwal-10',
    bulan: 'Oktober',
    mingguKe: 2,
    elemenMateri: 'Kelompok Sosial, Kesetaraan, and Konflik Sosial',
    subElemenMateri: 'Konflik sosial dan penanganan konflik.',
    kompetensi: 'Mendeskripsikan berbagai konsep konflik sosial dan menganalisis berbagai upaya penanganan konflik.'
  },
  {
    id: 'jadwal-11',
    bulan: 'Oktober',
    mingguKe: 3,
    elemenMateri: 'Perubahan Sosial dan Globalisasi',
    subElemenMateri: 'Bentuk-bentuk perubahan sosial dan dampaknya.',
    kompetensi: 'Mengidentifikasi bentuk-bentuk perubahan sosial dan menganalisis dampak perubahan sosial.'
  },
  {
    id: 'jadwal-12',
    bulan: 'Oktober',
    mingguKe: 4,
    elemenMateri: 'Perubahan Sosial dan Globalisasi',
    subElemenMateri: 'Globalisasi dan dampak globalisasi.',
    kompetensi: 'Menjelaskan dan menganalisis pengaruh globalisasi dan dampaknya.'
  }
];

export default function App() {
  // Navigation Tabs: 'config' (Generator & Prompt), 'kisi' (Matriks Asesmen), 'soal' (Pembuat Soal), 'jadwal' (Jadwal Pembelajaran), 'users' (Manajemen Pengguna)
  const [activeTab, setActiveTab] = useState<'config' | 'kisi' | 'soal' | 'jadwal' | 'users'>('config');

  // Interactive Quiz & Prompt Generator State (2 Wadah Prompt)
  const [cbtForm, setCbtForm] = useState({
    mataPelajaran: 'Sosiologi',
    materiPokok: 'Perubahan Sosial dan Globalisasi',
    lingkupMateri: 'Mengidentifikasi bentuk-bentuk perubahan sosial dan menganalisis dampak perubahan sosial.',
    subMateri: 'Bentuk-bentuk perubahan sosial dan dampaknya.',
    kompetensiYangDiuji: 'Peserta didik mampu menganalisis dampak globalisasi terhadap kearifan lokal masyarakat Nusantara.',
    batasanCatatan: 'Sajikan soal berorientasi HOTS (C4-C6) dengan stimulus studi kasus kontekstual.',
    levelKognitif: 'Penalaran HOTS (Reasoning) (level_3)',
    bentukSoal: 'Pilihan Ganda Sederhana (A-E)',
    jumlahSoal: 20, // Min 10, Max 50
    durasiMenit: 60,
    adminUsername: 'admin_cbt',
    adminPassword: 'admin_proktor2026',
    guruUsername: 'guru_sosiologi',
    guruPassword: 'guru_pass2026',
    usernameCbt: 'peserta_tka',
    passwordCbt: 'cbt2026_sosiologi',
    enableSecurity: true,
    randomizeOrder: true,
    antiTabSwitch: true,
    disableCopyPaste: true,
    konteksLokal: [
      '🎭 Budaya Nusantara',
      '🗺️ Geografis Indonesia',
      '👥 Kehidupan Sosial',
      '💰 Ekonomi Rakyat',
      '⚙️ Teknologi Tradisional',
      '🏛️ Kearifan Lokal',
      '🌈 Keragaman Etnis'
    ],
    stimulusKonten: [
      '📖 Teks Bacaan',
      '🖼️ Gambar/Ilustrasi',
      '📊 Data/Tabel',
      '📈 Grafik/Diagram',
      '🔍 Kasus Nyata',
      '📚 Cerita Pendek',
      '📰 Berita/Artikel'
    ],
    standarKualitas: [
      'Validasi Bahasa',
      'Konstruksi Soal',
      'Kesesuaian Materi',
      'Level Kognitif',
      'Konteks Relevan',
      'Tidak Bias',
      'Kejelasan Instruksi',
      'Kunci Jawaban Tepat',
      'Distractor Berkualitas',
      'Sesuai Kurikulum',
      'Waktu Pengerjaan',
      'Inklusivitas'
    ]
  });

  const [promptWadah1Text, setPromptWadah1Text] = useState<string>('');
  const [promptWadah2Text, setPromptWadah2Text] = useState<string>('');
  const [copiedWadah1, setCopiedWadah1] = useState(false);
  const [copiedWadah2, setCopiedWadah2] = useState(false);
  const [wadah2SuccessMsg, setWadah2SuccessMsg] = useState<string>('');
  const [quizActiveSubTab, setQuizActiveSubTab] = useState<'prompt' | 'embed'>('prompt');
  const [copiedEmbedCode, setCopiedEmbedCode] = useState(false);

  // User Management State (for Admin)
  const [usersList, setUsersList] = useState<any[]>([]);
  const [newUserName, setNewUserName] = useState('');
  const [newUserEmail, setNewUserEmail] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserRole, setNewUserRole] = useState<'admin' | 'user'>('user');
  const [newUserMataPelajaran, setNewUserMataPelajaran] = useState<string>('Sosiologi');
  const [isAddingUser, setIsAddingUser] = useState(false);
  const [userError, setUserError] = useState<string | null>(null);
  const [userSuccess, setUserSuccess] = useState<string | null>(null);

  // Supabase Database Connection State
  const [supabaseUrlInput, setSupabaseUrlInput] = useState<string>(() => getStoredSupabaseConfig().url);
  const [supabaseKeyInput, setSupabaseKeyInput] = useState<string>(() => getStoredSupabaseConfig().key);
  const [isTestingSupabase, setIsTestingSupabase] = useState<boolean>(false);
  const [supabaseTestMsg, setSupabaseTestMsg] = useState<{ success: boolean; text: string } | null>(null);
  const [showSupabaseSqlModal, setShowSupabaseSqlModal] = useState<boolean>(false);
  const [isCopyingSqlCode, setIsCopyingSqlCode] = useState<boolean>(false);


  const [config, setConfig] = useState<GeneratorConfig>({
    mataPelajaran: 'Sosiologi',
    definisi: 'Asesmen Akhir Semester (AAS) Ganjil',
    muatan: 'Kurikulum Merdeka - Fase F Kelas XII',
    kompetensi: 'Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.',
    bentukSoal: 'pilihan_ganda_sederhana',
    levelKognitif: 'level_2',
    elemenMateri: 'Sosiologi sebagai Ilmu',
    subElemenMateri: 'Pengertian dan perkembangan sosiologi dan manfaat sosiologi dalam kehidupan masyarakat.',
    batasanCatatan: 'Sejarah sosiologi, objek kajian sosiologi, fungsi dan manfaat sosiologi bagi masyarakat.',
    jumlahOpsi: 5,
    jenisSoal: 'tunggal',
    jumlahSoal: 1,
    konteksLokal: ['Kontekstual Indonesia', 'Kearifan Lokal'],
    stimulusKonten: ['Studi Kasus Konkrit', 'Data Statistik/Infografis', 'Wacana Ilmiah/Berita'],
    kualitasChecklist: [
      'Konstruksi Soal', 
      'Kesesuaian Materi', 
      'Level Kognitif', 
      'Kunci Jawaban Tepat', 
      'Distractor Berkualitas',
      'Sesuai Kurikulum'
    ]
  });

  const defaultKisiList: KisiKisiItem[] = [
    {
      id: "kisi-sosiologi-ref-1",
      no: 1,
      bentukSoal: "pilihan_ganda_sederhana",
      levelKognitif: "level_1",
      elemenMateri: "Sosiologi sebagai Ilmu",
      subElemenMateri: "Pengertian dan perkembangan sosiologi dan manfaat sosiologi dalam kehidupan masyarakat.",
      kompetensi: "Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.",
      batasanCatatan: "Sejarah sosiologi, objek kajian sosiologi, fungsi dan manfaat sosiologi bagi masyarakat.",
      jumlahSoal: 1
    },
    {
      id: "kisi-sosiologi-ref-2",
      no: 2,
      bentukSoal: "kategori",
      levelKognitif: "level_2",
      elemenMateri: "Hubungan dan Gejala Sosial",
      subElemenMateri: "Ragam gejala sosial.",
      kompetensi: "Menjelaskan ragam gejala sosial di lingkungan sekitar.",
      batasanCatatan: "Perilaku menyimpang, masalah sosial, sosiologi perkotaan/pedesaan, dan dampaknya bagi keteraturan sosial.",
      jumlahSoal: 1
    },
    {
      id: "kisi-sosiologi-ref-3",
      no: 3,
      bentukSoal: "mcma",
      levelKognitif: "level_3",
      elemenMateri: "Penelitian Sosial",
      subElemenMateri: "Langkah penelitian sosial dan metode penelitian.",
      kompetensi: "Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.",
      batasanCatatan: "Rancangan penelitian, jenis penelitian (kualitatif/kuantitatif), teknik sampling, pengumpulan data, dan penyusunan laporan.",
      jumlahSoal: 1
    }
  ];

  const defaultQuestions: Question[] = [
    {
      id: "question-sosiologi-ref-1",
      noSoal: 1,
      kisiKisiId: "kisi-sosiologi-ref-1",
      kompetensi: "Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.",
      subKompetensi: "Mengidentifikasi objek kajian sosiologi di era masyarakat digital.",
      bentukSoal: "pilihan_ganda_sederhana",
      stimulus: "Sosiologi merupakan ilmu pengetahuan murni yang membatasi diri pada apa yang nyata-nyata terjadi saat ini (das sein) dan bukan membicarakan apa yang seharusnya terjadi (das sollen). Di tengah era disrupsi teknologi digital saat ini, berbagai fenomena interaksi sosial baru bermunculan, mulai dari maraknya penggunaan media sosial, kecanduan gawai, hingga pola komunikasi virtual di kalangan remaja yang menggeser norma-norma konvensional di masyarakat.",
      soal: "Berdasarkan ilustrasi di atas, objek kajian sosiologi yang paling tepat ditunjukkan oleh pernyataan...",
      opsi: [
        "A. Dampak radiasi gelombang elektromagnetik gawai terhadap kesehatan mata dan fisik remaja secara klinis.",
        "B. Kecanduan teknologi yang mengubah pola interaksi, cara berpikir, dan perilaku sosial di kalangan remaja dalam kehidupan sehari-hari.",
        "C. Kerusakan jaringan infrastruktur internet nasional akibat dari maraknya serangan keamanan siber (cyber attack).",
        "D. Penurunan nilai tukar mata uang asing yang memengaruhi harga jual gawai di pasar lokal secara signifikan.",
        "E. Perancangan algoritma kecerdasan buatan pada aplikasi media sosial untuk meningkatkan efisiensi komputasi."
      ],
      kunciJawaban: "B",
      pembahasan: "Objek kajian sosiologi berpusat pada masyarakat and segala fenomena interaksi sosial serta gejala sosial yang terjadi di dalamnya. Dampak sosial kemajuan teknologi (kecanduan gawai yang mengubah pola interaksi, cara berpikir, dan perilaku sosial remaja) merupakan gejala sosial nyata yang menjadi objek kajian sosiologi (das sein). Pilihan lainnya berada di luar ranah kajian sosiologi, seperti kesehatan fisik/klinis (A), teknik informatika/cyber (C & E), dan ekonomi makro (D).",
      kataKunci: "Objek Kajian Sosiologi, Gejala Sosial, Disrupsi Teknologi"
    },
    {
      id: "question-sosiologi-ref-2",
      noSoal: 2,
      kisiKisiId: "kisi-sosiologi-ref-2",
      kompetensi: "Menjelaskan ragam gejala sosial di lingkungan sekitar.",
      subKompetensi: "Menerapkan konsep sosialisasi dan ragam gejala sosial terkait pembatasan screen-time pada anak.",
      bentukSoal: "kategori",
      stimulus: "Perhatikan anjuran durasi aman layar (screen-time) bagi anak-anak berikut ini!\nMenurut rekomendasi para ahli evaluasi perkembangan anak, anak usia 0 sampai 1,5 tahun disarankan sama sekali tidak terpapar layar gawai (0 jam). Anak usia 1,5 sampai 2 tahun diperbolehkan mengakses program yang berkualitas tinggi maksimal selama 1 jam dengan pendampingan ketat oleh orang tua. Anak usia 2 sampai 5 tahun juga dibatasi maksimal 1 jam per hari dengan pendampingan, sementara anak di atas 5 tahun harus memiliki batas waktu penggunaan gawai yang konsisten dan seimbang demi menjaga kesehatan fisik dan mental mereka.",
      soal: "Evaluasilah kesesuaian pernyataan terkait gejala penggunaan gawai pada anak berdasarkan anjuran tersebut! Tentukan SESUAI atau TIDAK SESUAI untuk setiap pernyataan berikut.",
      opsi: [
        "Pernyataan 1: Penggunaan gawai pada anak usia dini perlu diawasi ketat oleh orang tua agar proses sosialisasi primer anak tidak terhambat secara negatif.",
        "Pernyataan 2: Anak berusia 1 tahun diperbolehkan bermain gawai sendiri tanpa pendampingan asalkan kontennya edukatif dengan batas waktu maksimal 1 jam per hari.",
        "Pernyataan 3: Pola asuh yang terlalu longgar terhadap akses teknologi digital dapat mengganggu pembentukan karakter dan kepribadian sosial anak.",
        "Pernyataan 4: Pembatasan waktu layar secara konsisten bagi anak usia di atas 5 tahun dapat mengurangi risiko deviasi sosial berupa kecanduan gawai."
      ],
      kunciJawaban: "SESUAI, TIDAK SESUAI, SESUAI, SESUAI",
      pembahasan: "- Pernyataan 1 [SESUAI]: Sesuai dengan anjuran dalam stimulus bahwa pendampingan orang tua sangat krusial dalam masa sosialisasi primer anak usia dini.\n- Pernyataan 2 [TIDAK SESUAI]: Anak usia 1 tahun (berada di rentang 0-1,5 tahun) direkomendasikan sama sekali tidak terpapar layar (0 jam), serta tidak boleh dilepas bermain gawai sendiri.\n- Pernyataan 3 [SESUAI]: Sesuai dengan konsep sosiologi bahwa pengawasan/pendampingan penting untuk mencegah dampak negatif pembentukan kepribadian akibat paparan gawai yang bebas.\n- Pernyataan 4 [SESUAI]: Pembatasan waktu layar secara konsisten dan seimbang bagi anak di atas 5 tahun membantu mencegah risiko penyimpangan berupa kecanduan gawai.",
      kataKunci: "Sosialisasi, Pola Asuh, Gejala Sosial, Screen-Time"
    },
    {
      id: "question-sosiologi-ref-3",
      noSoal: 3,
      kisiKisiId: "kisi-sosiologi-ref-3",
      kompetensi: "Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.",
      subKompetensi: "Menganalisis rancangan metode penelitian sosial kuantitatif dan menyempurnakannya secara metodologis.",
      bentukSoal: "mcma",
      stimulus: "Seorang peneliti sosiologi SMA ingin meneliti pengaruh intensitas pergaulan kelompok teman sebaya (peer group) terhadap kelekatan hubungan antar-anggota keluarga di kalangan siswa kelas XII. Peneliti tersebut merumuskan masalah: 'Apakah terdapat hubungan antara pergaulan sebaya dengan kelekatan hubungan keluarga?' Peneliti menyusun instrumen pengumpulan data berupa daftar pertanyaan terbuka sebanyak 10 butir untuk wawancara mendalam. Namun, pada saat yang sama, ia juga berencana menganalisis kekuatan hubungan antar-variabel tersebut secara kuantitatif melalui uji korelasi statistik menggunakan aplikasi pengolah data.",
      soal: "Berdasarkan rancangan penelitian di atas, manakah rekomendasi metodologis yang paling tepat dan logis untuk menyempurnakan penelitian tersebut agar valid? (Pilihlah semua jawaban yang benar! Jawaban benar lebih dari satu)",
      opsi: [
        "A. Peneliti perlu mengubah daftar pertanyaan terbuka menjadi kuesioner tertutup berskala Likert agar data kuantitatif yang diperoleh dapat diolah dengan uji korelasi statistik secara valid.",
        "B. Peneliti harus menentukan teknik sampling (seperti simple random sampling atau stratified random sampling) dan ukuran sampel yang representatif terlebih dahulu sebelum menyebarkan instrumen.",
        "C. Peneliti sebaiknya menghapus rumusan masalah utama karena analisis statistik kuantitatif tidak memerlukan perumusan masalah yang detail terkait interaksi sosial.",
        "D. Peneliti wajib menggunakan metode observasi partisipatif penuh (peneliti tinggal bersama keluarga responden selama minimal satu tahun penuh) untuk mempercepat proses kuantifikasi.",
        "E. Peneliti perlu melakukan operasionalisasi konsep variabel bebas (intensitas pergaulan sebaya) dan variabel terikat (kelekatan hubungan keluarga) untuk mempermudah penyusunan indikator instrumen kuesioner."
      ],
      kunciJawaban: "A, B, E",
      pembahasan: "Penelitian ini memiliki kontradiksi metodologis: ingin menguji hubungan kuantitatif (korelasi statistik) tetapi instrumennya adalah pertanyaan terbuka (kualitatif). Maka rekomendasi penyempurnaan yang logis:\n1. [A BENAR] Pertanyaan terbuka harus diubah menjadi tertutup (seperti skala Likert) agar datanya berbentuk angka dan dapat diproses secara statistik.\n2. [B BENAR] Penentuan teknik sampling probabilitas dan jumlah sampel representatif sangat penting untuk penelitian kuantitatif agar hasil uji hubungan bisa digeneralisasi.\n3. [E BENAR] Operasionalisasi konsep variabel sangat krusial dalam kuantitatif untuk menerjemahkan teori sosiologi ke dalam indikator kuesioner yang valid.\nOpsi C salah karena rumusan masalah adalah fondasi utama penelitian. Opsi D tidak tepat karena observasi partisipatif penuh adalah metode khas kualitatif (etnografi) yang sangat lama dan bertolak belakang dengan kebutuhan pengujian korelasi kuantitatif cepat.",
      kataKunci: "Metodologi Penelitian, Penelitian Kuantitatif, Teknik Sampling, Validitas"
    }
  ];

  // User auth and role states
  const [currentUser, setCurrentUser] = useState<FirebaseUser | null>(null);
  const [userRole, setUserRole] = useState<'admin' | 'user' | null>(null);
  const [userName, setUserName] = useState<string>('');
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [showSignOutConfirm, setShowSignOutConfirm] = useState(false);
  const [quotaExceededInfo, setQuotaExceededInfo] = useState<FirestoreErrorInfo | null>(null);

  useEffect(() => {
    const handleQuotaExceeded = (_e: Event) => {
      // Intentionally do not show any quota exceeded banner to the user
      setQuotaExceededInfo(null);
    };
    window.addEventListener('firestore-quota-exceeded', handleQuotaExceeded);
    return () => window.removeEventListener('firestore-quota-exceeded', handleQuotaExceeded);
  }, []);

  // Synced collection states
  const [kisiList, setKisiList] = useState<KisiKisiItem[]>([]);
  const [questions, setQuestions] = useState<Question[]>([]);
  const [selectedBentukFilter, setSelectedBentukFilter] = useState<'all' | BentukSoal>('all');

  // Jadwal Rencana Pembelajaran State & Handlers
  const [jadwalList, setJadwalList] = useState<JadwalItem[]>([]);
  const [jadwalSortNotification, setJadwalSortNotification] = useState<string | null>(null);

  // Sync selectedJadwalPresetSubject with config.mataPelajaran
  useEffect(() => {
    if (!config.mataPelajaran) return;
    const mp = config.mataPelajaran;
    if (mp === 'Pendidikan Pancasila dan Kewarganegaraan') {
      setSelectedJadwalPresetSubject('PPKN');
    } else if (mp === 'Sejarah') {
      setSelectedJadwalPresetSubject('Sejarah Tingkat Lanjut');
    } else if (mp === 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK') {
      setSelectedJadwalPresetSubject('Produk Kreatif dan Kewirausahaan');
    } else {
      const validSubjects = [
        'Matematika', 'Bahasa Indonesia', 'Bahasa Inggris', 
        'Matematika Tingkat Lanjut', 'Bahasa Indonesia Tingkat Lanjut', 'Bahasa Inggris Tingkat Lanjut', 
        'Fisika', 'Kimia', 'Biologi', 'Ekonomi', 'Geografi', 'Sosiologi', 'Antropologi', 'Bahasa Jepang'
      ];
      if (validSubjects.includes(mp)) {
        setSelectedJadwalPresetSubject(mp as any);
      }
    }
  }, [config.mataPelajaran]);

  // Form State for Adding / Editing Jadwal
  const [isAddingJadwal, setIsAddingJadwal] = useState(false);
  const [newJadwal, setNewJadwal] = useState<Omit<JadwalItem, 'id'>>({
    bulan: 'Juli',
    mingguKe: 1,
    elemenMateri: '',
    subElemenMateri: '',
    kompetensi: ''
  });

  const [editingJadwalId, setEditingJadwalId] = useState<string | null>(null);
  const [editingJadwalData, setEditingJadwalData] = useState<JadwalItem | null>(null);

  // Custom confirmation modals states
  const [jadwalToDelete, setJadwalToDelete] = useState<JadwalItem | null>(null);
  const [showClearJadwalConfirm, setShowClearJadwalConfirm] = useState(false);
  const [showDeleteAllConfirm, setShowDeleteAllConfirm] = useState(false);
  const [showImportPresetsConfirm, setShowImportPresetsConfirm] = useState<{ count: number; subject: string; presets: any[] } | null>(null);
  const [showImportKisiPresetsConfirm, setShowImportKisiPresetsConfirm] = useState<{ count: number; subject: string; subjectMapped: string; presets: any[] } | null>(null);
  const [isSessionOnlyMode, setIsSessionOnlyMode] = useState<boolean>(false);

  const handleClearAllSessionData = async () => {
    if (!confirm("Apakah Anda yakin ingin MENGOSONGKAN SELURUH SESI (Matriks Kisi-Kisi dan Butir Soal)?\n\nTindakan ini akan menghapus semua kisi-kisi dan soal di layar serta membersihkan penyimpanan agar tidak membebani sistem.")) {
      return;
    }
    setKisiList([]);
    setQuestions([]);
    setIsEditingQuestion(false);
    setEditingQuestionId(null);

    if (currentUser?.uid) {
      try {
        const qKisi = query(collection(db, 'kisi_kisi'), where('userId', '==', currentUser.uid));
        const kisiSnap = await getDocs(qKisi);
        if (!kisiSnap.empty) {
          const batch = writeBatch(db);
          kisiSnap.docs.forEach(d => {
            batch.delete(d.ref);
            batch.delete(doc(db, 'materials', d.id));
          });
          await batch.commit().catch(e => console.warn("Clear session kisi error:", e));
        }

        const qQuest = query(collection(db, 'questions'), where('userId', '==', currentUser.uid));
        const questSnap = await getDocs(qQuest);
        if (!questSnap.empty) {
          const batch = writeBatch(db);
          questSnap.docs.forEach(d => batch.delete(d.ref));
          await batch.commit().catch(e => console.warn("Clear session questions error:", e));
        }
      } catch (err) {
        console.warn("Clear session firestore cleanup warning:", err);
      }
    }
    alert("Sesi berhasil dibersihkan! Tampilan kisi-kisi dan butir soal kini kosong dan bebas beban database.");
  };

  // State for preset subject selection in the Jadwal UI
  const [selectedJadwalPresetSubject, setSelectedJadwalPresetSubject] = useState<'Matematika' | 'Bahasa Indonesia' | 'Bahasa Inggris' | 'Matematika Tingkat Lanjut' | 'Bahasa Indonesia Tingkat Lanjut' | 'Bahasa Inggris Tingkat Lanjut' | 'Fisika' | 'Kimia' | 'Biologi' | 'PPKN' | 'Ekonomi' | 'Geografi' | 'Sosiologi' | 'Sejarah Tingkat Lanjut' | 'Antropologi' | 'Bahasa Jepang' | 'Produk Kreatif dan Kewirausahaan'>('Sosiologi');

  // Handle importing all presets into the schedule distributed sequentially across months
  const handleImportAllJadwalPresets = () => {
    const activePresets = selectedJadwalPresetSubject === 'Matematika' 
      ? PUSMENDIK_MATEMATIKA_PRESETS 
      : selectedJadwalPresetSubject === 'Bahasa Indonesia' 
      ? PUSMENDIK_BAHASA_INDONESIA_PRESETS 
      : selectedJadwalPresetSubject === 'Bahasa Inggris'
      ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
      : selectedJadwalPresetSubject === 'Matematika Tingkat Lanjut'
      ? PUSMENDIK_MATEMATIKA_TL_PRESETS
      : selectedJadwalPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
      ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
      : selectedJadwalPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
      ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
      : selectedJadwalPresetSubject === 'Fisika'
      ? PUSMENDIK_FISIKA_PRESETS
      : selectedJadwalPresetSubject === 'Kimia'
      ? PUSMENDIK_KIMIA_PRESETS
      : selectedJadwalPresetSubject === 'Biologi'
      ? PUSMENDIK_BIOLOGI_PRESETS
      : selectedJadwalPresetSubject === 'PPKN'
      ? PUSMENDIK_PPKN_PRESETS
      : selectedJadwalPresetSubject === 'Ekonomi'
      ? PUSMENDIK_EKONOMI_PRESETS
      : selectedJadwalPresetSubject === 'Geografi'
      ? PUSMENDIK_GEOGRAFI_PRESETS
      : selectedJadwalPresetSubject === 'Sosiologi'
      ? PUSMENDIK_SOSIOLOGI_PRESETS
      : selectedJadwalPresetSubject === 'Sejarah Tingkat Lanjut'
      ? PUSMENDIK_SEJARAH_TL_PRESETS
      : selectedJadwalPresetSubject === 'Antropologi'
      ? PUSMENDIK_ANTROPOLOGI_PRESETS
      : selectedJadwalPresetSubject === 'Bahasa Jepang'
      ? PUSMENDIK_BAHASA_JEPANG_PRESETS
      : PUSMENDIK_PKK_PRESETS;

    const count = activePresets.length;
    setShowImportPresetsConfirm({
      count,
      subject: selectedJadwalPresetSubject,
      presets: activePresets
    });
  };

  const executeImportAllJadwalPresets = async (presets: any[]) => {
    const months: ('Juli' | 'Agustus' | 'September' | 'Oktober')[] = ['Juli', 'Agustus', 'September', 'Oktober'];
    const newItems: JadwalItem[] = presets.map((preset, index) => {
      const slotIndex = index;
      const monthIndex = Math.min(3, Math.floor(slotIndex / 4));
      const weekNum = (slotIndex % 4) + 1;
      return {
        id: `jadwal-preset-${Date.now()}-${index}`,
        userId: currentUser?.uid,
        bulan: months[monthIndex],
        mingguKe: weekNum,
        elemenMateri: preset.elemenMateri,
        subElemenMateri: preset.subElemenMateri,
        kompetensi: preset.kompetensi
      };
    });

    setJadwalList(prev => {
      const updated = [...prev, ...newItems];
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_jadwal_local', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    try {
      const batch = writeBatch(db);
      newItems.forEach((item) => {
        batch.set(doc(db, 'jadwal_pembelajaran', item.id), item);
      });
      await batch.commit();
    } catch (err: any) {
      console.warn("Notice syncing jadwal presets to Cloud:", err);
    }

    setShowImportPresetsConfirm(null);
  };

  const handleAddJadwal = async (e: React.FormEvent) => {
    e.preventDefault();
    const item: JadwalItem = {
      ...newJadwal,
      id: `jadwal-${Date.now()}`,
      userId: currentUser?.uid
    };
    
    setJadwalList(prev => {
      const updated = [...prev, item];
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_jadwal_local', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    try {
      await setDoc(doc(db, 'jadwal_pembelajaran', item.id), item);
    } catch (err: any) {
      console.warn("Notice syncing new jadwal to Cloud:", err);
    }

    setNewJadwal({
      bulan: 'Juli',
      mingguKe: 1,
      elemenMateri: '',
      subElemenMateri: '',
      kompetensi: ''
    });
    setIsAddingJadwal(false);
  };

  const handleStartEditJadwal = (item: JadwalItem) => {
    setEditingJadwalId(item.id);
    setEditingJadwalData({ ...item });
  };

  const handleSaveEditJadwal = async () => {
    if (!editingJadwalData) return;

    setJadwalList(prev => {
      const updated = prev.map(item => item.id === editingJadwalData.id ? editingJadwalData : item);
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_jadwal_local', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    try {
      await setDoc(doc(db, 'jadwal_pembelajaran', editingJadwalData.id), editingJadwalData);
    } catch (err: any) {
      console.warn("Notice updating jadwal in Cloud:", err);
    }
    setEditingJadwalId(null);
    setEditingJadwalData(null);
  };

  const handleDeleteJadwal = (id: string) => {
    const itemToDelete = jadwalList.find(item => item.id === id);
    if (!itemToDelete) return;
    setJadwalToDelete(itemToDelete);
  };

  const executeDeleteJadwal = async () => {
    if (!jadwalToDelete) return;

    setJadwalList(prev => {
      const updated = prev.filter(item => item.id !== jadwalToDelete.id);
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_jadwal_local', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    try {
      await deleteDoc(doc(db, 'jadwal_pembelajaran', jadwalToDelete.id));
    } catch (err: any) {
      console.warn("Notice deleting jadwal from Cloud:", err);
    }
    setJadwalToDelete(null);
  };

  const handleResetJadwal = () => {
    setShowClearJadwalConfirm(true);
  };

  const handleSortJadwal = async () => {
    if (jadwalList.length === 0) return;

    const monthOrder: Record<string, number> = {
      'Juli': 1,
      'Agustus': 2,
      'September': 3,
      'Oktober': 4,
      'November': 5,
      'Desember': 6,
      'Januari': 7,
      'Februari': 8,
      'Maret': 9,
      'April': 10,
      'Mei': 11,
      'Juni': 12
    };

    const sorted = [...jadwalList].sort((a, b) => {
      const mA = monthOrder[a.bulan] || 99;
      const mB = monthOrder[b.bulan] || 99;
      if (mA !== mB) {
        return mA - mB;
      }
      return (Number(a.mingguKe) || 0) - (Number(b.mingguKe) || 0);
    });

    setJadwalList(sorted);
    try {
      if (currentUser?.uid) {
        localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify(sorted));
      }
      localStorage.setItem('tka_jadwal_local', JSON.stringify(sorted));
    } catch (e) {}

    try {
      const batch = writeBatch(db);
      sorted.forEach((item, index) => {
        batch.set(doc(db, 'jadwal_pembelajaran', item.id), { ...item, no: index + 1 });
      });
      await batch.commit();
    } catch (err: any) {
      console.warn("Notice sorting jadwal in Cloud:", err);
    }

    setJadwalSortNotification('Jadwal Rencana Pembelajaran TKA Kelas XII berhasil diurutkan berdasarkan Bulan dan Minggu ke-!');
    setTimeout(() => {
      setJadwalSortNotification(null);
    }, 4000);
  };

  const executeClearJadwal = async () => {
    setJadwalList([]);
    try {
      if (currentUser?.uid) {
        localStorage.setItem(`tka_jadwal_${currentUser.uid}`, JSON.stringify([]));
      }
      localStorage.setItem('tka_jadwal_local', JSON.stringify([]));
    } catch (e) {}

    try {
      const batch = writeBatch(db);
      jadwalList.forEach((item) => {
        batch.delete(doc(db, 'jadwal_pembelajaran', item.id));
      });
      await batch.commit();
    } catch (err: any) {
      console.warn("Notice clearing jadwal in Cloud:", err);
    }
    setShowClearJadwalConfirm(false);
  };

  const getPresetsForSubject = (subject: string) => {
    switch (subject) {
      case 'Matematika':
        return PUSMENDIK_MATEMATIKA_PRESETS;
      case 'Bahasa Indonesia':
        return PUSMENDIK_BAHASA_INDONESIA_PRESETS;
      case 'Bahasa Inggris':
        return PUSMENDIK_BAHASA_INGGRIS_PRESETS;
      case 'Matematika Tingkat Lanjut':
        return PUSMENDIK_MATEMATIKA_TL_PRESETS;
      case 'Bahasa Indonesia Tingkat Lanjut':
        return PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS;
      case 'Bahasa Inggris Tingkat Lanjut':
        return PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS;
      case 'Fisika':
        return PUSMENDIK_FISIKA_PRESETS;
      case 'Kimia':
        return PUSMENDIK_KIMIA_PRESETS;
      case 'Biologi':
        return PUSMENDIK_BIOLOGI_PRESETS;
      case 'PPKN':
      case 'Pendidikan Pancasila dan Kewarganegaraan':
        return PUSMENDIK_PPKN_PRESETS;
      case 'Ekonomi':
        return PUSMENDIK_EKONOMI_PRESETS;
      case 'Geografi':
        return PUSMENDIK_GEOGRAFI_PRESETS;
      case 'Sosiologi':
        return PUSMENDIK_SOSIOLOGI_PRESETS;
      case 'Sejarah':
      case 'Sejarah Tingkat Lanjut':
        return PUSMENDIK_SEJARAH_TL_PRESETS;
      case 'Antropologi':
        return PUSMENDIK_ANTROPOLOGI_PRESETS;
      case 'Bahasa Jepang':
        return PUSMENDIK_BAHASA_JEPANG_PRESETS;
      default:
        return PUSMENDIK_PKK_PRESETS;
    }
  };

  const seedDefaultData = async (userId: string, targetSubject: string = 'Sosiologi') => {
    try {
      const userDocRef = doc(db, 'users', userId);
      const userDocSnap = await getDoc(userDocRef);
      const isUserTeacher = userDocSnap.exists() && userDocSnap.data().role !== 'admin';

      // For non-admin teachers, do not auto-seed sample matrix rows so they start in an empty state ("posisi kosong blm ada yang dipilih")
      if (isUserTeacher) {
        await setDoc(userDocRef, { isSeeded: true }, { merge: true });
        return;
      }

      const kisiSnap = await getDocs(query(collection(db, 'kisi_kisi'), where('userId', '==', userId)));
      
      // If kisi_kisi is already seeded for this user, return
      if (!kisiSnap.empty) {
        await setDoc(doc(db, 'users', userId), { isSeeded: true }, { merge: true });
        return;
      }

      const batch = writeBatch(db);
      
      let defaultKisiList: KisiKisiItem[] = [];
      if (targetSubject === 'Sosiologi') {
        defaultKisiList = [
          {
            id: `kisi-sosiologi-ref-1-${userId}`,
            userId: userId,
            no: 1,
            bentukSoal: "pilihan_ganda_sederhana",
            levelKognitif: "level_1",
            elemenMateri: "Sosiologi sebagai Ilmu",
            subElemenMateri: "Pengertian dan perkembangan sosiologi dan manfaat sosiologi dalam kehidupan masyarakat.",
            kompetensi: "Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.",
            batasanCatatan: "Sejarah sosiologi, objek kajian sosiologi, fungsi dan manfaat sosiologi bagi masyarakat.",
            jumlahSoal: 1
          },
          {
            id: `kisi-sosiologi-ref-2-${userId}`,
            userId: userId,
            no: 2,
            bentukSoal: "kategori",
            levelKognitif: "level_2",
            elemenMateri: "Hubungan dan Gejala Sosial",
            subElemenMateri: "Ragam gejala sosial.",
            kompetensi: "Menjelaskan ragam gejala sosial di lingkungan sekitar.",
            batasanCatatan: "Perilaku menyimpang, masalah sosial, sosiologi perkotaan/pedesaan, dan dampaknya bagi keteraturan sosial.",
            jumlahSoal: 1
          },
          {
            id: `kisi-sosiologi-ref-3-${userId}`,
            userId: userId,
            no: 3,
            bentukSoal: "mcma",
            levelKognitif: "level_3",
            elemenMateri: "Penelitian Sosial",
            subElemenMateri: "Langkah penelitian sosial dan metode penelitian.",
            kompetensi: "Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.",
            batasanCatatan: "Rancangan penelitian, jenis penelitian (kualitatif/kuantitatif), teknik sampling, pengumpulan data, dan penyusunan laporan.",
            jumlahSoal: 1
          }
        ];
      } else {
        const presets = getPresetsForSubject(targetSubject).slice(0, 3);
        defaultKisiList = presets.map((preset, idx) => ({
          id: `kisi-${targetSubject.toLowerCase().replace(/[^a-z0-9]+/g, '-')}-ref-${idx + 1}-${userId}`,
          userId: userId,
          no: idx + 1,
          bentukSoal: idx === 0 ? "pilihan_ganda_sederhana" : idx === 1 ? "kategori" : "mcma",
          levelKognitif: idx === 0 ? "level_1" : idx === 1 ? "level_2" : "level_3",
          elemenMateri: preset.elemenMateri,
          subElemenMateri: preset.subElemenMateri,
          kompetensi: preset.kompetensi,
          batasanCatatan: preset.batasanCatatan,
          jumlahSoal: 1
        }));
      }

      defaultKisiList.forEach((item) => {
        batch.set(doc(db, 'kisi_kisi', item.id), item);
      });

      if (targetSubject === 'Sosiologi') {
        const defaultQuestions: Question[] = [
          {
            id: `question-sosiologi-ref-1-${userId}`,
            userId: userId,
            noSoal: 1,
            kisiKisiId: `kisi-sosiologi-ref-1-${userId}`,
            kompetensi: "Mendeskripsikan dan menganalisis pengertian dan perkembangan serta manfaat sosiologi sebagai ilmu pengetahuan.",
            subKompetensi: "Mengidentifikasi objek kajian sosiologi di era masyarakat digital.",
            bentukSoal: "pilihan_ganda_sederhana",
            stimulus: "Sosiologi merupakan ilmu pengetahuan murni yang membatasi diri pada apa yang nyata-nyata terjadi saat ini (das sein) dan bukan membicarakan apa yang seharusnya terjadi (das sollen). Di tengah era disrupsi teknologi digital saat ini, berbagai fenomena interaksi sosial baru bermunculan, mulai dari maraknya penggunaan media sosial, kecanduan gawai, hingga pola komunikasi virtual di kalangan remaja yang menggeser norma-norma konvensional di masyarakat.",
            soal: "Berdasarkan ilustrasi di atas, objek kajian sosiologi yang paling tepat ditunjukkan oleh pernyataan...",
            opsi: [
              "A. Dampak radiasi gelombang elektromagnetik gawai terhadap kesehatan mata dan fisik remaja secara klinis.",
              "B. Kecanduan teknologi yang mengubah pola interaksi, cara berpikir, dan perilaku sosial di kalangan remaja dalam kehidupan sehari-hari.",
              "C. Kerusakan jaringan infrastruktur internet nasional akibat dari maraknya serangan keamanan siber (cyber attack).",
              "D. Penurunan nilai tukar mata uang asing yang memengaruhi harga jual gawai di pasar lokal secara signifikan.",
              "E. Perancangan algoritma kecerdasan buatan pada aplikasi media sosial untuk meningkatkan efisiensi komputasi."
            ],
            kunciJawaban: "B",
            pembahasan: "Objek kajian sosiologi berpusat pada masyarakat and segala fenomena interaksi sosial serta gejala sosial yang terjadi di dalamnya. Dampak sosial kemajuan teknologi (kecanduan gawai yang mengubah pola interaksi, cara berpikir, dan perilaku sosial remaja) merupakan gejala sosial nyata yang menjadi objek kajian sosiologi (das sein). Pilihan lainnya berada di luar ranah kajian sosiologi, seperti kesehatan fisik/klinis (A), teknik informatika/cyber (C & E), dan ekonomi makro (D).",
            kataKunci: "Objek Kajian Sosiologi, Gejala Sosial, Disrupsi Teknologi"
          },
          {
            id: `question-sosiologi-ref-2-${userId}`,
            userId: userId,
            noSoal: 2,
            kisiKisiId: `kisi-sosiologi-ref-2-${userId}`,
            kompetensi: "Menjelaskan ragam gejala sosial di lingkungan sekitar.",
            subKompetensi: "Menerapkan konsep sosialisasi dan ragam gejala sosial terkait pembatasan screen-time pada anak.",
            bentukSoal: "kategori",
            stimulus: "Perhatikan anjuran durasi aman layar (screen-time) bagi anak-anak berikut ini!\nMenurut rekomendasi para ahli evaluasi perkembangan anak, anak usia 0 sampai 1,5 tahun disarankan sama sekali tidak terpapar layar gawai (0 jam). Anak usia 1,5 sampai 2 tahun diperbolehkan mengakses program yang berkualitas tinggi maksimal selama 1 jam dengan pendampingan ketat oleh orang tua. Anak usia 2 sampai 5 tahun juga dibatasi maksimal 1 jam per hari dengan pendampingan, sementara anak di atas 5 tahun harus memiliki batas waktu penggunaan gawai yang konsisten dan seimbang demi menjaga kesehatan fisik dan mental mereka.",
            soal: "Evaluasilah kesesuaian pernyataan terkait gejala penggunaan gawai pada anak berdasarkan anjuran tersebut! Tentukan SESUAI atau TIDAK SESUAI untuk setiap pernyataan berikut.",
            opsi: [
              "Pernyataan 1: Penggunaan gawai pada anak usia dini perlu diawasi ketat oleh orang tua agar proses sosialisasi primer anak tidak terhambat secara negatif.",
              "Pernyataan 2: Anak berusia 1 tahun diperbolehkan bermain gawai sendiri tanpa pendampingan asalkan kontennya edukatif dengan batas waktu maksimal 1 jam per hari.",
              "Pola asuh yang terlalu longgar terhadap akses teknologi digital dapat mengganggu pembentukan karakter dan kepribadian sosial anak.",
              "Pembatasan waktu layar secara konsisten bagi anak usia di atas 5 tahun dapat mengurangi risiko deviasi sosial berupa kecanduan gawai."
            ],
            kunciJawaban: "SESUAI, TIDAK SESUAI, SESUAI, SESUAI",
            pembahasan: "- Pernyataan 1 [SESUAI]: Sesuai dengan anjuran dalam stimulus bahwa pendampingan orang tua sangat krusial dalam masa sosialisasi primer anak usia dini.\n- Pernyataan 2 [TIDAK SESUAI]: Anak usia 1 tahun (berada di rentang 0-1,5 tahun) direkomendasikan sama sekali tidak terpapar layar (0 jam), serta tidak boleh dilepas bermain gawai sendiri.\n- Pernyataan 3 [SESUAI]: Sesuai dengan konsep sosiologi bahwa pengawasan/pendampingan penting untuk mencegah dampak negatif pembentukan kepribadian akibat paparan gawai yang bebas.\n- Pernyataan 4 [SESUAI]: Pembatasan waktu layar secara konsisten dan seimbang bagi anak di atas 5 tahun membantu mencegah risiko penyimpangan berupa kecanduan gawai.",
            kataKunci: "Sosialisasi, Pola Asuh, Gejala Sosial, Screen-Time"
          },
          {
            id: `question-sosiologi-ref-3-${userId}`,
            userId: userId,
            noSoal: 3,
            kisiKisiId: `kisi-sosiologi-ref-3-${userId}`,
            kompetensi: "Menjelaskan dan menganalisis berbagai langkah dan metode penelitian sosial.",
            subKompetensi: "Menganalisis rancangan metode penelitian sosial kuantitatif dan menyempurnakannya secara metodologis.",
            bentukSoal: "mcma",
            stimulus: "Seorang peneliti sosiologi SMA ingin meneliti pengaruh intensitas pergaulan kelompok teman sebaya (peer group) terhadap kelekatan hubungan antar-anggota keluarga di kalangan siswa kelas XII. Peneliti tersebut merumuskan masalah: 'Apakah terdapat hubungan antara pergaulan sebaya dengan kelekatan hubungan keluarga?' Peneliti menyusun instrumen pengumpulan data berupa daftar pertanyaan terbuka sebanyak 10 butir untuk wawancara mendalam. Namun, pada saat yang sama, ia juga berencana menganalisis kekuatan hubungan antar-variabel tersebut secara kuantitatif melalui uji korelasi statistik menggunakan aplikasi pengolah data.",
            soal: "Berdasarkan rancangan penelitian di atas, manakah rekomendasi metodologis yang paling tepat dan logis untuk menyempurnakan penelitian tersebut agar valid? (Pilihlah semua jawaban yang benar! Jawaban benar lebih dari satu)",
            opsi: [
              "A. Peneliti perlu mengubah daftar pertanyaan terbuka menjadi kuesioner tertutup berskala Likert agar data kuantitatif yang diperoleh dapat diolah dengan uji korelasi statistik secara valid.",
              "B. Peneliti harus menentukan teknik sampling (seperti simple random sampling atau stratified random sampling) dan ukuran sampel yang representatif terlebih dahulu sebelum menyebarkan instrumen.",
              "C. Peneliti sebaiknya menghapus rumusan masalah utama karena analisis statistik kuantitatif tidak memerlukan perumusan masalah yang detail terkait interaksi sosial.",
              "D. Peneliti wajib menggunakan metode observasi partisipatif penuh (peneliti tinggal bersama keluarga responden selama minimal satu tahun penuh) untuk mempercepat proses kuantifikasi.",
              "E. Peneliti perlu melakukan operasionalisasi konsep variabel bebas (intensitas pergaulan sebaya) dan variabel terikat (kelekatan hubungan keluarga) untuk mempermudah penyusunan indikator instrumen kuesioner."
            ],
            kunciJawaban: "A, B, E",
            pembahasan: "Penelitian ini memiliki kontradiksi metodologis: ingin menguji hubungan kuantitatif (korelasi statistik) tetapi instrumennya adalah pertanyaan terbuka (kualitatif). Maka rekomendasi penyempurnaan yang logis:\n1. [A BENAR] Pertanyaan terbuka harus diubah menjadi tertutup (seperti skala Likert) agar datanya berbentuk angka dan dapat diproses secara statistik.\n2. [B BENAR] Penentuan teknik sampling probabilitas dan jumlah sampel representatif sangat penting untuk penelitian kuantitatif agar hasil uji hubungan bisa digeneralisasi.\n3. [E BENAR] Operasionalisasi konsep variabel sangat krusial dalam kuantitatif untuk menerjemahkan teori sosiologi ke dalam indikator kuesioner yang valid.\nOpsi C salah karena rumusan masalah adalah fondasi utama penelitian. Opsi D tidak tepat karena observasi partisipatif penuh adalah metode khas kualitatif (etnografi) yang sangat lama dan bertolak belakang dengan kebutuhan pengujian korelasi kuantitatif cepat.",
            kataKunci: "Metodologi Penelitian, Penelitian Kuantitatif, Teknik Sampling, Validitas"
          }
        ];

        defaultQuestions.forEach((item) => {
          batch.set(doc(db, 'questions', item.id), item);
        });

        const defaultMaterials = {
          [`kisi-sosiologi-ref-1-${userId}`]: `# 1. PENDAHULUAN & DEFINISI\nSosiologi berasal dari bahasa Latin *socius* yang berarti teman atau kawan, dan bahasa Yunani *logos* yang berarti ilmu atau berbicara. Secara harfiah, sosiologi adalah ilmu tentang masyarakat. Auguste Comte, bapak sosiologi, mendefinisikan sosiologi sebagai ilmu positif tentang hukum-hukum dasar gejala sosial. Sosiologi merupakan ilmu pengetahuan murni (*pure science*) dan ilmu abstrak (*abstract science*) yang membatasi diri pada apa yang nyata terjadi (*das sein*) bukan apa yang seharusnya terjadi (*das sollen*).\n\n# 2. KONSEP UTAMA & TEORI PENDEKATAN\n* **Objek Kajian Sosiologi**: Objek material sosiologi adalah kehidupan sosial, gejala-gejala sosial, dan proses hubungan antarmanusia. Objek formal sosiologi adalah manusia sebagai makhluk sosial serta interaksi sosial antarmanusia dalam masyarakat.\n* **Paradigma Sosiologi**: Terdapat tiga paradigma utama menurut George Ritzer:\n  1. *Fakta Sosial* (Durkheim): Struktur dan institusi sosial yang memengaruhi individu secara eksternal dan memaksa.\n  2. *Definisi Sosial* (Weber): Tindakan sosial yang memiliki makna subjektif bagi pelakunya.\n  3. *Perilaku Sosial* (Skinner): Hubungan stimulus-respons dan pengulangan perilaku berdasarkan konsekuensi.\n\n# 3. STUDI KASUS KONKRIT (KONTEKSTUAL INDONESIA)\nDi era disrupsi digital Indonesia saat ini, muncul fenomena interaksi sosial virtual baru di kalangan remaja, seperti pembentukan komunitas daring di Discord dan penyebaran konten di TikTok. Interaksi ini tidak dibatasi oleh ruang fisik, namun memicu pergeseran nilai dan norma konvensional, seperti memudarnya sopan santun komunikasi langsung (tatap muka) karena terbiasa dengan anonimitas di dunia maya.\n\n# 4. ANALISIS KRITIS & REFLEKSI\n**Pertanyaan Reflektif**: Bagaimana kemunculan fenomena "flexing" (pamer kekayaan) di media sosial Indonesia dianalisis menggunakan paradigma definisi sosial Max Weber? Analisislah makna subjektif di balik tindakan pamer tersebut dan bagaimana masyarakat mengonstruksi status sosial di ruang digital!`,
          [`kisi-sosiologi-ref-2-${userId}`]: `# 1. PENDAHULUAN & DEFINISI\nSosialisasi adalah sebuah proses seumur hidup di mana individu mempelajari nilai, norma, peran, dan perilaku sosial yang berlaku di masyarakatnya untuk membentuk kepribadian yang utuh. Sosialisasi primer merupakan tahap awal yang berlangsung di lingkungan keluarga, yang menjadi landasan utama pembentukan karakter dasar anak sebelum ia berinteraksi dengan lingkungan luar (sosialisasi sekunder).\n\n# 2. KONSEP UTAMA & TEORI PENDEKATAN\n* **Tahapan Sosialisasi (George Herbert Mead)**:\n  1. *Preparatory Stage* (Persiapan): Bayi meniru tindakan orang dewasa tanpa memahami maknanya.\n  2. *Play Stage* (Meniru): Anak mulai meniru peran orang di sekitarnya secara sadar (misal bermain peran ibu/guru).\n  3. *Game Stage* (Siap Bertindak): Anak memahami perannya sendiri dan peran orang lain yang terlibat dalam permainan terstruktur.\n  4. *Generalized Other* (Penerimaan Norma): Anak mampu menginternalisasi nilai dan norma masyarakat secara luas serta bertindak sebagai warga masyarakat yang bertanggung jawab.\n* **Pola Asuh Sosialisasi**:\n  - *Sosialisasi Represif*: Berfokus pada kepatuhan ketat, hukuman fisik, dan komunikasi satu arah (dominasi orang tua).\n  - *Sosialisasi Partisipatoris*: Berfokus pada interaksi timbal balik, hadiah atas perilaku baik, dan komunikasi dua arah yang menempatkan anak sebagai pusat perhatian.\n\n# 3. STUDI KASUS KONKRIT (KONTEKSTUAL INDONESIA)\nBanyak keluarga perkotaan di Indonesia yang menerapkan pola asuh longgar atau menggunakan gawai sebagai "pengasuh elektronik" demi kepraktisan. Anak-anak dibiarkan mengakses layar (*screen-time*) di atas durasi aman tanpa pendampingan. Gejala sosial ini mengganggu tahap *play stage* anak karena interaksi konkret dengan manusia berkurang, yang berakibat pada hambatan emosional dan lambatnya pemahaman norma-norma sosial primer.\n\n# 4. ANALISIS KRITIS & REFLEKSI\n**Pertanyaan Reflektif**: Jika dikaitkan dengan pembentukan karakter Pancasila, apa dampak jangka panjang bagi ketahanan sosial nasional apabila sosialisasi primer dalam keluarga Indonesia digantikan sepenuhnya oleh algoritma media sosial global? Rincikan solusi taktis sosiologis bagi para orang tua modern!`,
          [`kisi-sosiologi-ref-3-${userId}`]: `# 1. PENDAHULUAN & DEFINISI\nPenelitian sosial adalah penyelidikan terencana, kritis, dan empiris untuk memecahkan masalah-masalah sosial atau menguji kebenaran teori sosiologi yang ada di masyarakat. Penelitian sosial bertumpu pada keobjektifan ilmiah, keteraturan metodologis, serta kejujuran data lapangan agar hasil kesimpulannya valid dan dapat dipertanggungjawabkan secara akademis.\n\n# 2. KONSEP UTAMA & TEORI PENDEKATAN\n* **Metode Penelitian Kuantitatif**: Berorientasi pada pembuktian teori, pengujian hubungan antar-variabel secara statistik, instrumen terstruktur (kuesioner tertutup/skala Likert), pengambilan sampel probabilitas (*random sampling*), serta analisis data objektif-numerik.\n* **Metode Penelitian Kualitatif**: Berorientasi pada pemahaman mendalam (*verstehen*), deskripsi interpretatif wacana atau makna sosial, instrumen fleksibel (wawancara mendalam, observasi partisipatif), serta teknik sampling non-probabilitas (*purposive/snowball sampling*).\n* **Operasionalisasi Variabel**: Proses menerjemahkan konsep teoretis yang abstrak (variabel bebas & terikat) menjadi indikator-indikator empiris terukur untuk memudahkan pembuatan instrumen kuesioner.\n\n# 3. STUDI KASUS KONKRIT (KONTEKSTUAL INDONESIA)\nSeorang peneliti sosiologi ingin meneliti pengaruh intensitas pergaulan kelompok teman sebaya (*peer group*) terhadap kelekatan hubungan antar-anggota keluarga siswa kelas XII di sebuah SMA di Jakarta. Agar riset kuantitatif ini valid, peneliti menerjemahkan konsep "intensitas pergaulan" menjadi indikator terukur (seperti frekuensi berkumpul dalam seminggu dan durasi interaksi) serta menggunakan skala Likert 1-5 untuk kuesioner tertutup.\n\n# 4. ANALISIS KRITIS & REFLEKSI\n**Pertanyaan Reflektif**: Mengapa pencampuran instrumen kualitatif (wawancara terbuka) ke dalam analisis korelasi statistik murni tanpa metodologi *Mixed Methods* yang jelas sering kali menghasilkan bias validitas? Jelaskan bagaimana integrasi triangulasi metode yang tepat dapat menyelesaikannya!`
        };

        Object.entries(defaultMaterials).forEach(([kId, content]) => {
          batch.set(doc(db, 'materials', kId), {
            content,
            userId,
            updatedAt: new Date()
          });
        });
      }

      batch.update(doc(db, 'users', userId), { isSeeded: true });

      await batch.commit();
      console.log("Seeding default data completed successfully.");
    } catch (err) {
      console.error("Gagal melakukan seeding data default:", err);
    }
  };

  // User Session Management (Apps Script & Local Storage Session)
  useEffect(() => {
    try {
      const activeSession = localStorage.getItem('tka_active_session');
      if (activeSession) {
        const userObj = JSON.parse(activeSession);
        if (userObj && userObj.email) {
          setCurrentUser({ uid: userObj.id || 'usr_demo', email: userObj.email, displayName: userObj.name || userObj.email.split('@')[0] });
          setUserRole(userObj.role === 'admin' ? 'admin' : 'user');
          setUserName(userObj.name || userObj.email.split('@')[0]);
          if (userObj.mataPelajaran) {
            setConfig(prev => ({ ...prev, mataPelajaran: userObj.mataPelajaran }));
          }
        }
      }
    } catch (e) {
      console.warn("Notice reading active session:", e);
    } finally {
      setIsAuthLoading(false);
    }
  }, []);

  // Real-time Cloud User Database Subscription (Multi-User & Multi-Device Sync)
  useEffect(() => {
    if (!currentUser) return;
    const unsub = subscribeToUsers((users) => {
      setUsersList(users);
    });
    return () => unsub();
  }, [currentUser]);

  // Fast local storage initialization and lightweight User Management listener
  useEffect(() => {
    if (!currentUser) return;

    // Load initial cached workspace data from localStorage for instant, offline-first loading
    try {
      const keys = [`tka_kisi_${currentUser.uid}`, 'tka_kisi_local', 'tka_kisi_guest'];
      for (const k of keys) {
        const cachedKisi = localStorage.getItem(k);
        if (cachedKisi) {
          const parsed = JSON.parse(cachedKisi);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setKisiList(parsed);
            break;
          }
        }
      }

      const qKeys = [`tka_questions_${currentUser.uid}`, 'tka_questions_local', 'tka_questions_guest'];
      for (const k of qKeys) {
        const cachedQ = localStorage.getItem(k);
        if (cachedQ) {
          const parsed = JSON.parse(cachedQ);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setQuestions(parsed);
            break;
          }
        }
      }

      const jKeys = [`tka_jadwal_${currentUser.uid}`, 'tka_jadwal_local'];
      let foundJadwal = false;
      for (const k of jKeys) {
        const cachedJ = localStorage.getItem(k);
        if (cachedJ) {
          const parsed = JSON.parse(cachedJ);
          if (Array.isArray(parsed) && parsed.length > 0) {
            setJadwalList(parsed);
            foundJadwal = true;
            break;
          }
        }
      }
      if (!foundJadwal) {
        setJadwalList(DEFAULT_JADWAL_LIST);
      }

      const cfgKey = `tka_config_${currentUser.uid}`;
      const cachedCfg = localStorage.getItem(cfgKey) || localStorage.getItem('tka_config_local');
      if (cachedCfg) {
        try { setConfig(JSON.parse(cachedCfg)); } catch (e) {}
      }

      const cbtKey = `tka_cbt_${currentUser.uid}`;
      const cachedCbt = localStorage.getItem(cbtKey) || localStorage.getItem('tka_cbt_local');
      if (cachedCbt) {
        try { setCbtForm(JSON.parse(cachedCbt)); } catch (e) {}
      }
    } catch (e) {
      console.warn("Notice reading initial cache:", e);
    }

    // Listen to current user profile changes in real time (User Management)
    const unsubscribeUserProfile = onSnapshot(doc(db, 'users', currentUser.uid), (snap) => {
      if (snap.exists()) {
        const data = snap.data();
        if (data.role) setUserRole(data.role);
        if (data.name) setUserName(data.name);
        if (data.mataPelajaran) {
          const activeSubject = data.mataPelajaran;
          setConfig(prev => {
            if (prev.mataPelajaran !== activeSubject) {
              return { ...prev, mataPelajaran: activeSubject };
            }
            return prev;
          });
        }
      }
    }, (error) => {
      console.warn("User profile listener notice:", error?.message || error);
    });

    // Single non-blocking background fetch for Cloud data if available
    const loadCloudDataOnce = async () => {
      try {
        const kisiSnap = await getDocs(query(collection(db, 'kisi_kisi'), where('userId', '==', currentUser.uid)));
        if (!kisiSnap.empty) {
          const list: KisiKisiItem[] = [];
          kisiSnap.forEach(d => list.push({ id: d.id, ...d.data() } as KisiKisiItem));
          list.sort((a, b) => (a.no || 0) - (b.no || 0));
          setKisiList(prev => prev.length > 0 ? prev : list);
        }
      } catch (e) {
        console.warn("Background kisi sync notice:", e);
      }

      try {
        const qSnap = await getDocs(query(collection(db, 'questions'), where('userId', '==', currentUser.uid)));
        if (!qSnap.empty) {
          const list: Question[] = [];
          qSnap.forEach(d => list.push({ id: d.id, ...d.data() } as Question));
          list.sort((a, b) => (a.noSoal || 0) - (b.noSoal || 0));
          setQuestions(prev => prev.length > 0 ? prev : list);
        }
      } catch (e) {
        console.warn("Background questions sync notice:", e);
      }

      try {
        const jSnap = await getDocs(query(collection(db, 'jadwal_pembelajaran'), where('userId', '==', currentUser.uid)));
        if (!jSnap.empty) {
          const list: JadwalItem[] = [];
          jSnap.forEach(d => list.push({ id: d.id, ...d.data() } as JadwalItem));
          setJadwalList(list);
        }
      } catch (e) {
        console.warn("Background jadwal sync notice:", e);
      }
    };

    loadCloudDataOnce();

    // Listen to Users (Admin only for User Management)
    let unsubscribeUsers = () => {};
    if (userRole === 'admin') {
      unsubscribeUsers = onSnapshot(collection(db, 'users'), (snapshot) => {
        const list: any[] = [];
        snapshot.forEach((doc) => {
          list.push({ id: doc.id, ...doc.data() });
        });
        list.sort((a, b) => (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase()));
        setUsersList(list);
      }, (error) => {
        console.warn("Admin users listener notice:", error?.message || error);
      });
    }

    return () => {
      unsubscribeUserProfile();
      unsubscribeUsers();
    };
  }, [currentUser, userRole]);

  // Auto-cache kisiList & questions to localStorage for offline / domain resiliency
  useEffect(() => {
    if (kisiList.length === 0) return;
    try {
      if (currentUser?.uid) {
        localStorage.setItem(`tka_kisi_${currentUser.uid}`, JSON.stringify(kisiList));
      }
      localStorage.setItem('tka_kisi_local', JSON.stringify(kisiList));
      localStorage.setItem('tka_kisi_guest', JSON.stringify(kisiList));
    } catch (e) {
      console.warn("Notice saving kisi cache:", e);
    }
  }, [kisiList, currentUser]);

  useEffect(() => {
    if (questions.length === 0) return;
    try {
      if (currentUser?.uid) {
        localStorage.setItem(`tka_questions_${currentUser.uid}`, JSON.stringify(questions));
      }
      localStorage.setItem('tka_questions_local', JSON.stringify(questions));
      localStorage.setItem('tka_questions_guest', JSON.stringify(questions));
    } catch (e) {
      console.warn("Notice saving questions cache:", e);
    }
  }, [questions, currentUser]);

  // Debounced auto-sync for Generator Config (Bagian 1) to user-specific Firestore document
  useEffect(() => {
    if (!currentUser) return;
    const timeout = setTimeout(() => {
      setDoc(doc(db, 'user_settings', `${currentUser.uid}_generator_config`), config, { merge: true }).catch((err) => {
        console.warn("Notice syncing generator_config to Cloud:", err);
      });
    }, 2000);
    return () => clearTimeout(timeout);
  }, [config, currentUser]);

  // Debounced auto-sync for CBT Config (Bagian 4) to user-specific Firestore document
  useEffect(() => {
    if (!currentUser) return;
    const timeout = setTimeout(() => {
      setDoc(doc(db, 'user_settings', `${currentUser.uid}_cbt_config`), cbtForm, { merge: true }).catch((err) => {
        console.warn("Notice syncing cbt_config to Cloud:", err);
      });
    }, 2000);
    return () => clearTimeout(timeout);
  }, [cbtForm, currentUser]);

  const handleSignOut = async () => {
    setShowSignOutConfirm(true);
  };

  const executeSignOut = async () => {
    setShowSignOutConfirm(false);
    try {
      localStorage.removeItem('tka_active_session');
      setCurrentUser(null);
      setUserRole(null);
      setUserName('');
      await signOut(auth).catch(() => {});
    } catch (err) {
      console.error("Error signing out:", err);
    }
  };

  const isAdmin = userRole === 'admin';

  // Loading States
  const [isGeneratingKisi, setIsGeneratingKisi] = useState(false);
  const [isGeneratingSoal, setIsGeneratingSoal] = useState(false);
  const [isSavingQuestion, setIsSavingQuestion] = useState(false);
  const [importingPresetIds, setImportingPresetIds] = useState<Record<string, boolean>>({});
  const [apiStatus, setApiStatus] = useState<string | null>(null);
  const [soalProgress, setSoalProgress] = useState({
    active: false,
    type: 'single' as 'single' | 'all',
    currentNo: 0,
    totalNo: 0,
    topic: '',
    countSuccess: 0,
    totalQuestions: 0,
    statusText: ''
  });

  // State for AI Custom Illustration Generator (Nana Banana)
  const [isAiIllustratorOpen, setIsAiIllustratorOpen] = useState(false);
  const [aiIllustratorPrompt, setAiIllustratorPrompt] = useState('');
  const [isGeneratingIllustration, setIsGeneratingIllustration] = useState(false);
  const [aiIllustratorStatus, setAiIllustratorStatus] = useState('');
  const [copiedPromptId, setCopiedPromptId] = useState<string | null>(null);
  const [activePromptTab, setActivePromptTab] = useState<'ilustrasi' | 'tabel' | 'grafik' | 'stimulus' | 'matriks'>('ilustrasi');
  
  // Menu 1 Sub-Tab State ('parameter' for Admin Input Parameter, 'prompt' for Promt Generator & AI)
  const [configSubTab, setConfigSubTab] = useState<'parameter' | 'prompt'>('parameter');
  const effectiveSubTab = userRole === 'admin' ? configSubTab : 'prompt';

  // Helper to normalize Base URL for Google AI Studio & OpenAI compatibility
  const normalizeBaseUrl = (url?: string): string => {
    let u = (url || '').trim().replace(/\/+$/, '');
    if (!u || u === 'https://api.koboillm.com/v1') {
      return 'https://generativelanguage.googleapis.com/v1beta/openai';
    }
    if (u.includes('generativelanguage.googleapis.com')) {
      return 'https://generativelanguage.googleapis.com/v1beta/openai';
    }
    return u;
  };

  // AI Config state (Support client-side direct bypass of Vercel 10s timeouts & Direct Gemini API)
  const [aiConfig, setAiConfig] = useState(() => {
    const savedKey = localStorage.getItem('gemini_api_key') || '';
    let savedMode = localStorage.getItem('gemini_api_mode') || 'server';
    if (savedMode === 'n8n') savedMode = 'client';
    let savedBaseUrl = localStorage.getItem('litellm_base_url') || 'https://generativelanguage.googleapis.com/v1beta/openai';
    if (savedBaseUrl === 'https://api.koboillm.com/v1' || savedBaseUrl.includes('generativelanguage.googleapis.com')) {
      savedBaseUrl = 'https://generativelanguage.googleapis.com/v1beta/openai';
      localStorage.setItem('litellm_base_url', 'https://generativelanguage.googleapis.com/v1beta/openai');
    }
    let savedModel = localStorage.getItem('gemini_api_model') || 'gemini-2.0-flash';
    if (
      !savedModel ||
      savedModel.includes('3.5') ||
      savedModel.includes('2.5')
    ) {
      savedModel = 'gemini-2.0-flash';
      localStorage.setItem('gemini_api_model', 'gemini-2.0-flash');
    }
    return {
      mode: savedMode as 'server' | 'client',
      apiKey: savedKey,
      baseUrl: savedBaseUrl,
      model: savedModel,
      temperature: parseFloat(localStorage.getItem('gemini_api_temperature') || '0.7'),
      requestDelayMs: parseInt(localStorage.getItem('gemini_api_delay') || '0', 10)
    };
  });
  const [availableModels, setAvailableModels] = useState<string[]>([
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash-8b"
  ]);
  const [fetchingModels, setFetchingModels] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [showApiKeySaved, setShowApiKeySaved] = useState(false);

  // States for API Key Health Check, Rotation Animations, and Quota Alerts
  const [apiKeyToast, setApiKeyToast] = useState<{
    id: string;
    type: 'rotation' | 'quota_exhausted' | 'health_check';
    title: string;
    message: string;
    keyIndex?: number;
    totalKeys?: number;
  } | null>(null);

  const [apiKeyHealthList, setApiKeyHealthList] = useState<Array<{
    keyIndex: number;
    snippet: string;
    status: 'valid' | 'exhausted' | 'invalid' | 'testing';
    message: string;
  }>>([]);
  const [testingKeyHealth, setTestingKeyHealth] = useState(false);

  const handleFetchModels = async () => {
    setFetchingModels(true);
    try {
      const res = await fetch('/api/fetch-models', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': getCleanApiKey(aiConfig.apiKey)
        },
        body: JSON.stringify({
          apiKey: aiConfig.apiKey,
          baseUrl: normalizeBaseUrl(aiConfig.baseUrl)
        })
      });

      if (res.ok) {
        const data = await res.json();
        if (Array.isArray(data.models) && data.models.length > 0) {
          setAvailableModels(data.models);
          if (!data.models.includes(aiConfig.model)) {
            const firstModel = data.models[0];
            setAiConfig(prev => ({ ...prev, model: firstModel }));
            localStorage.setItem('gemini_api_model', firstModel);
          }
          alert(`Berhasil memuat ${data.totalModels} model dari ${data.baseUrl}!`);
        } else {
          alert("Tidak ada model yang ditemukan dari Base URL.");
        }
      } else {
        const errText = await res.text();
        alert(`Gagal mengambil model: ${errText}`);
      }
    } catch (err: any) {
      alert(`Terjadi kesalahan saat mengambil daftar model: ${err.message}`);
    } finally {
      setFetchingModels(false);
    }
  };

  // Helper triggers for API Key Rotation Toast / Quota Exhaustion Alerts
  const triggerApiKeyRotationToast = (fromIdx: number, toIdx: number, totalKeys: number) => {
    setApiKeyToast({
      id: `rotation-${Date.now()}`,
      type: 'rotation',
      title: '🔄 Rotasi Otomatis Kunci API Google AI Studio',
      message: `Kunci API #${fromIdx + 1} mengalami limit kuota/rate limit. Sistem otomatis berpindah ke Kunci API #${toIdx + 1} dari total ${totalKeys} kunci tanpa menghentikan pembuatan soal.`,
      keyIndex: toIdx + 1,
      totalKeys
    });

    // Auto dismiss rotation toast after 7 seconds
    setTimeout(() => {
      setApiKeyToast(prev => (prev?.type === 'rotation' ? null : prev));
    }, 7000);
  };

  const triggerApiKeyQuotaExhaustedToast = (customMessage?: string) => {
    const totalKeys = (aiConfig.apiKey || '').split(/[\n,;]+/).map(k => k.trim()).filter(k => k.length > 5).length || 1;
    setApiKeyToast({
      id: `quota-${Date.now()}`,
      type: 'quota_exhausted',
      title: '⚠️ Kuota Kunci API Habis (Error 429)',
      message: customMessage || `Seluruh ${totalKeys} Kunci API Anda telah mencapai batas kuota harian/menit.`
    });
  };

  const handleTestApiKeyHealth = async () => {
    const keys = (aiConfig.apiKey || '').split(/[\n,;]+/).map(k => k.trim()).filter(k => k.length > 5);
    if (keys.length === 0) {
      alert("Masukkan minimal satu Kunci API terlebih dahulu di Langkah 1.");
      return;
    }

    setTestingKeyHealth(true);
    try {
      const res = await fetch('/api/test-key-health', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': getCleanApiKey(aiConfig.apiKey)
        },
        body: JSON.stringify({
          apiKey: aiConfig.apiKey,
          baseUrl: normalizeBaseUrl(aiConfig.baseUrl)
        })
      });

      if (res.ok) {
        const data = await res.json();
        setApiKeyHealthList(data.results || []);
        const validCount = (data.results || []).filter((r: any) => r.status === 'valid').length;
        setApiKeyToast({
          id: `health-${Date.now()}`,
          type: 'health_check',
          title: '✅ Hasil Uji Kesehatan Kunci API',
          message: `${validCount} dari ${keys.length} Kunci API aktif dan siap digunakan untuk rotasi otomatis.`
        });
        setTimeout(() => setApiKeyToast(prev => prev?.type === 'health_check' ? null : prev), 5000);
      } else {
        const errText = await res.text();
        alert(`Gagal menguji kesehatan key: ${errText}`);
      }
    } catch (err: any) {
      alert(`Terjadi kesalahan saat menguji kesehatan key: ${err.message}`);
    } finally {
      setTestingKeyHealth(false);
    }
  };

  const handleSetAiMode = (mode: 'server' | 'client') => {
    setAiConfig(prev => ({ ...prev, mode }));
    localStorage.setItem('gemini_api_mode', mode);
  };

  const getCleanApiKey = (keyString: string) => {
    return (keyString || '').replace(/[\r\n]+/g, ',').trim();
  };

  const callNativeGeminiClient = async (
    apiKey: string,
    modelName: string,
    systemInstruction: string,
    promptText: string,
    responseSchema?: any
  ): Promise<string> => {
    const cleanModel = (modelName || '').trim().replace(/^(models\/|google\/|gemini\/|publishers\/google\/models\/)/i, '').trim() || "gemini-2.0-flash";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${cleanModel}:generateContent?key=${encodeURIComponent(apiKey)}`;
    
    const requestBody: any = {
      contents: [
        {
          role: "user",
          parts: [{ text: promptText }]
        }
      ],
      generationConfig: {
        temperature: aiConfig.temperature || 0.7
      }
    };

    if (systemInstruction) {
      requestBody.systemInstruction = {
        parts: [{ text: systemInstruction }]
      };
    }

    if (responseSchema) {
      requestBody.generationConfig.responseMimeType = "application/json";
    }

    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(requestBody)
    });

    if (!res.ok) {
      let errText = '';
      try {
        const errObj = await res.json();
        errText = errObj.error?.message || JSON.stringify(errObj);
      } catch {
        errText = await res.text();
      }
      throw new Error(`Native Gemini API Error (${res.status}): ${errText}`);
    }

    const data = await res.json();
    const candidateText = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!candidateText) {
      throw new Error("Respon kosong dari Native Gemini API.");
    }
    return candidateText;
  };

  const callGeminiDirect = async (
    systemInstruction: string,
    promptText: string,
    responseSchema?: any
  ): Promise<string> => {
    if (!aiConfig.apiKey) {
      throw new Error("Kunci API Google AI Studio belum diatur! Silakan masukkan Kunci API Anda dari https://aistudio.google.com/app/apikey terlebih dahulu di Tab 1 (Pengaturan Koneksi AI).");
    }

    const apiKeys = (aiConfig.apiKey || '')
      .split(/[\n,;]+/)
      .map(k => k.trim())
      .filter(k => k.length > 5);

    if (apiKeys.length === 0) {
      throw new Error("Kunci API tidak valid! Silakan masukkan Kunci API terlebih dahulu di Langkah 1.");
    }

    const baseUrl = normalizeBaseUrl(aiConfig.baseUrl);
    const isGoogle = baseUrl.includes("generativelanguage.googleapis.com");
    const cleanModelName = (m: string) => (m || '').trim().replace(/^(models\/|google\/|gemini\/|publishers\/google\/models\/)/i, '').trim();

    const rawSelectedModel = aiConfig.model || "gemini-2.0-flash";
    let candidateModels: string[] = [];
    if (isGoogle) {
      const primary = cleanModelName(rawSelectedModel) || "gemini-2.0-flash";
      candidateModels = Array.from(new Set([
        primary,
        ...(availableModels || []).map(cleanModelName),
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-2.0-flash-lite",
        "gemini-1.5-flash-8b"
      ].filter(Boolean)));
    } else {
      candidateModels = Array.from(new Set([
        rawSelectedModel,
        cleanModelName(rawSelectedModel),
        `google/${cleanModelName(rawSelectedModel)}`,
        `gemini/${cleanModelName(rawSelectedModel)}`,
        ...(availableModels || []),
        "gemini-2.0-flash",
        "gemini-1.5-flash"
      ].filter(Boolean)));
    }

    let lastError: any = null;

    if (aiConfig.requestDelayMs && aiConfig.requestDelayMs > 0) {
      await new Promise(r => setTimeout(r, aiConfig.requestDelayMs));
    }

    for (let keyIdx = 0; keyIdx < apiKeys.length; keyIdx++) {
      const apiKey = apiKeys[keyIdx];
      console.log(`[Client API Key Rotation] Trying Key #${keyIdx + 1} of ${apiKeys.length} at ${baseUrl}...`);

      if (keyIdx > 0) {
        triggerApiKeyRotationToast(keyIdx - 1, keyIdx, apiKeys.length);
      }

      for (const modelCandidate of candidateModels) {
        try {
          const messages: any[] = [];
          if (systemInstruction) {
            messages.push({ role: "system", content: systemInstruction });
          }
          messages.push({ role: "user", content: promptText });

          const requestBody: any = {
            model: modelCandidate,
            messages,
            temperature: aiConfig.temperature || 0.7
          };

          if (responseSchema) {
            requestBody.response_format = { type: "json_object" };
          }

          const res = await fetch(`${baseUrl}/chat/completions`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${apiKey}`,
              'x-api-key': apiKey
            },
            body: JSON.stringify(requestBody)
          });

          if (res.ok) {
            const result = await res.json();
            const candidateText = result.choices?.[0]?.message?.content;
            if (candidateText) {
              return candidateText;
            }
          }

          let errorText = '';
          try {
            const errObj = await res.json();
            errorText = errObj.error?.message || JSON.stringify(errObj);
          } catch {
            errorText = await res.text();
          }

          // If 404 or model error on Google, try native Gemini REST fallback
          if (isGoogle && (res.status === 404 || /not found|404|invalid model/i.test(errorText))) {
            console.log(`[Client Fallback] Trying Native Gemini API for model ${modelCandidate}...`);
            try {
              const nativeText = await callNativeGeminiClient(
                apiKey,
                modelCandidate,
                systemInstruction,
                promptText,
                responseSchema
              );
              if (nativeText) {
                return nativeText;
              }
            } catch (nativeErr: any) {
              console.warn(`Native Gemini direct call failed for ${modelCandidate}:`, nativeErr?.message);
              lastError = nativeErr;
            }
          } else {
            lastError = new Error(`API Gemini Error (${res.status}): ${errorText || res.statusText}`);
          }
        } catch (err: any) {
          lastError = err;
          console.warn(`Direct call with Key #${keyIdx + 1} model ${modelCandidate} failed:`, err?.message || err);
        }
      }
    }

    const errorString = typeof lastError === 'string' ? lastError : (lastError?.message || JSON.stringify(lastError) || '');
    if (/quota|limit|429|exhausted|503/i.test(errorString)) {
      triggerApiKeyQuotaExhaustedToast('Seluruh Kunci API Direct Client Anda telah mencapai batas kuota (Error 429).');
      throw new Error(`⚠️ Kuota API Google AI Studio Direct (Client) Telah Terlampaui (Error 429 Exceeded Quota).`);
    }

    throw lastError || new Error("Gagal memanggil API Gemini setelah merotasi seluruh API Key.");
  };

  // Prompt Generator outputs
  const [generatedKisiPrompt, setGeneratedKisiPrompt] = useState('');
  const [generatedSoalPrompt, setGeneratedSoalPrompt] = useState('');
  const [copiedKisi, setCopiedKisi] = useState(false);
  const [copiedSoal, setCopiedSoal] = useState(false);

  // Form states for adding/editing a Kisi-Kisi Row manually
  const [isEditingKisi, setIsEditingKisi] = useState(false);
  const [editingKisiId, setEditingKisiId] = useState<string | null>(null);
  const [kisiForm, setKisiForm] = useState<Partial<KisiKisiItem>>({
    bentukSoal: 'pilihan_ganda_sederhana',
    levelKognitif: 'level_2',
    jenisSoal: 'tunggal',
    elemenMateri: '',
    subElemenMateri: '',
    kompetensi: '',
    batasanCatatan: '',
    jumlahSoal: 5,
    konteksNusantara: '',
    stimulusTambahan: '',
    konteksLokal: [],
    stimulusKonten: [],
    kualitasChecklist: []
  });

  // Form states for adding/editing a Question manually
  const [isEditingQuestion, setIsEditingQuestion] = useState(false);
  const [editingQuestionId, setEditingQuestionId] = useState<string | null>(null);
  
  // State for preset subject selection in the matrix UI
  const [selectedPresetSubject, setSelectedPresetSubject] = useState<'Matematika' | 'Bahasa Indonesia' | 'Bahasa Inggris' | 'Matematika Tingkat Lanjut' | 'Bahasa Indonesia Tingkat Lanjut' | 'Bahasa Inggris Tingkat Lanjut' | 'Fisika' | 'Kimia' | 'Biologi' | 'PPKN' | 'Ekonomi' | 'Geografi' | 'Sosiologi' | 'Sejarah Tingkat Lanjut' | 'Antropologi' | 'Bahasa Jepang' | 'Produk Kreatif dan Kewirausahaan'>('Sosiologi');

  // State for Print Settings (Menu Setting Cetak)
  const [printConfig, setPrintConfig] = useState({
    showHeader: true,
    kopDepartment: 'KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI',
    schoolName: 'SMA NEGERI NUSANTARA',
    examName: 'PENILAIAN AKHIR SEMESTER',
    academicYear: '2026/2027',
    semester: 'Ganjil',
    timeAllocation: '90 Menit',
    showStudentFields: true,
    showAnswerKey: false,
    fontSize: 'text-sm', // 'text-xs' | 'text-sm' | 'text-base'
    layoutColumns: '1', // '1' | '2'
    showStimulus: true,
    showIllustration: true,
    showCompetencyTag: false, // Default false: hide "Kompetensi" tag in actual exam view
    instructionText: 'Pilihlah salah satu jawaban yang paling tepat dengan memberi tanda silang (X) atau klik pada pilihan jawaban A, B, C, D, atau E!',
    schoolLogo: '', // Base64 or URL for left logo
    schoolLogoRight: '', // Base64 or URL for right logo
    pageSize: 'A4', // 'A4' | 'F4'
    subjectName: 'Sosiologi',
    schoolAddress: 'Jalan Pendidikan Raya No. 45 Nusantara - Telp/Fax: (021) 777-1234 - Website: www.sekolahkita.sch.id',
  });
  const [isPrintSettingsOpen, setIsPrintSettingsOpen] = useState(true);

  // Sync preset subject selection and subjectName with config mataPelajaran if applicable
  useEffect(() => {
    setPrintConfig(prev => ({ ...prev, subjectName: config.mataPelajaran }));
    if (config.mataPelajaran === 'Matematika Tingkat Lanjut') {
      setSelectedPresetSubject('Matematika Tingkat Lanjut');
    } else if (config.mataPelajaran === 'Matematika') {
      setSelectedPresetSubject('Matematika');
    } else if (config.mataPelajaran === 'Bahasa Indonesia Tingkat Lanjut') {
      setSelectedPresetSubject('Bahasa Indonesia Tingkat Lanjut');
    } else if (config.mataPelajaran === 'Bahasa Indonesia') {
      setSelectedPresetSubject('Bahasa Indonesia');
    } else if (config.mataPelajaran === 'Bahasa Inggris Tingkat Lanjut') {
      setSelectedPresetSubject('Bahasa Inggris Tingkat Lanjut');
    } else if (config.mataPelajaran === 'Bahasa Inggris') {
      setSelectedPresetSubject('Bahasa Inggris');
    } else if (config.mataPelajaran === 'Fisika') {
      setSelectedPresetSubject('Fisika');
    } else if (config.mataPelajaran === 'Kimia') {
      setSelectedPresetSubject('Kimia');
    } else if (config.mataPelajaran === 'Biologi') {
      setSelectedPresetSubject('Biologi');
    } else if (config.mataPelajaran === 'PPKN' || config.mataPelajaran === 'Pendidikan Pancasila dan Kewarganegaraan') {
      setSelectedPresetSubject('PPKN');
    } else if (config.mataPelajaran === 'Ekonomi') {
      setSelectedPresetSubject('Ekonomi');
    } else if (config.mataPelajaran === 'Geografi') {
      setSelectedPresetSubject('Geografi');
    } else if (config.mataPelajaran === 'Sosiologi') {
      setSelectedPresetSubject('Sosiologi');
    } else if (config.mataPelajaran === 'Sejarah' || config.mataPelajaran === 'Sejarah Tingkat Lanjut') {
      setSelectedPresetSubject('Sejarah Tingkat Lanjut');
    } else if (config.mataPelajaran === 'Antropologi') {
      setSelectedPresetSubject('Antropologi');
    } else if (config.mataPelajaran === 'Bahasa Jepang') {
      setSelectedPresetSubject('Bahasa Jepang');
    } else if (config.mataPelajaran === 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK') {
      setSelectedPresetSubject('Produk Kreatif dan Kewirausahaan');
    }

    if (currentUser && config.mataPelajaran) {
      setDoc(doc(db, 'users', currentUser.uid), {
        mataPelajaran: config.mataPelajaran
      }, { merge: true }).catch(err => console.warn("Notice saving subject selection to users:", err));
      setDoc(doc(db, 'user_settings', `${currentUser.uid}_generator_config`), {
        mataPelajaran: config.mataPelajaran
      }, { merge: true }).catch(err => console.warn("Notice saving subject selection to settings:", err));
    }
  }, [config.mataPelajaran, currentUser]);

  const handleSelectPresetSubject = (subject: typeof selectedPresetSubject) => {
    setSelectedPresetSubject(subject);
    const presetSubjectMapped = subject === 'PPKN' 
      ? 'Pendidikan Pancasila dan Kewarganegaraan'
      : subject === 'Sejarah Tingkat Lanjut'
      ? 'Sejarah'
      : subject === 'Produk Kreatif dan Kewirausahaan'
      ? 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK'
      : subject;
    setConfig(prev => ({
      ...prev,
      mataPelajaran: presetSubjectMapped
    }));
  };

  // Enforce Guru Mata Pelajaran lock for Matriks Asesmen & Preset Subject
  useEffect(() => {
    if (userRole !== 'admin' && config.mataPelajaran) {
      const mapel = config.mataPelajaran;
      if (mapel === 'Pendidikan Pancasila dan Kewarganegaraan') {
        setSelectedPresetSubject('PPKN');
      } else if (mapel === 'Sejarah') {
        setSelectedPresetSubject('Sejarah Tingkat Lanjut');
      } else if (mapel === 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK') {
        setSelectedPresetSubject('Produk Kreatif dan Kewirausahaan');
      } else if ([
        'Matematika', 'Bahasa Indonesia', 'Bahasa Inggris', 
        'Matematika Tingkat Lanjut', 'Bahasa Indonesia Tingkat Lanjut', 'Bahasa Inggris Tingkat Lanjut', 
        'Fisika', 'Kimia', 'Biologi', 'PPKN', 'Ekonomi', 'Geografi', 'Sosiologi', 
        'Sejarah Tingkat Lanjut', 'Antropologi', 'Bahasa Jepang', 'Produk Kreatif dan Kewirausahaan'
      ].includes(mapel as any)) {
        setSelectedPresetSubject(mapel as any);
      }
    }
  }, [userRole, config.mataPelajaran]);

  // Admin User Management CRUD & Edit States
  const [editingUser, setEditingUser] = useState<any | null>(null);
  const [userSearchQuery, setUserSearchQuery] = useState<string>('');
  const [isSavingUserEdit, setIsSavingUserEdit] = useState<boolean>(false);
  const [isBatchImportingUsers, setIsBatchImportingUsers] = useState<boolean>(false);
  const [showUserPasswords, setShowUserPasswords] = useState<boolean>(false);
  const [revealedPasswordUserIds, setRevealedPasswordUserIds] = useState<Record<string, boolean>>({});
  
  // Inline Deletion Confirmation States
  const [deletingKisiId, setDeletingKisiId] = useState<string | null>(null);
  const [deletingQuestionId, setDeletingQuestionId] = useState<string | null>(null);

  // Image Optimization & Zoom Lightbox States
  const [isCompressingImage, setIsCompressingImage] = useState<boolean>(false);
  const [imageCompressReport, setImageCompressReport] = useState<{ originalSizeKb: number; compressedSizeKb: number; savingPercent: number } | null>(null);
  const [activeZoomImage, setActiveZoomImage] = useState<{ url: string; caption?: string } | null>(null);
  const [zoomScale, setZoomScale] = useState<number>(1);
  const [zoomRotation, setZoomRotation] = useState<number>(0);

  // Client-side HTMLCanvasElement Image Compressor (Optimasi & Kompresi Performa)
  const compressImageFile = (
    file: File, 
    maxWidth = 1000, 
    maxHeight = 800, 
    quality = 0.8
  ): Promise<{ dataUrl: string; originalSizeKb: number; compressedSizeKb: number; savingPercent: number }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new window.Image();
        img.onload = () => {
          let width = img.width;
          let height = img.height;

          if (width > maxWidth) {
            height = Math.round((height * maxWidth) / width);
            width = maxWidth;
          }
          if (height > maxHeight) {
            width = Math.round((width * maxHeight) / height);
            height = maxHeight;
          }

          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          if (!ctx) {
            reject(new Error("Gagal menginisialisasi canvas context"));
            return;
          }

          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = 'high';
          ctx.drawImage(img, 0, 0, width, height);

          const dataUrl = canvas.toDataURL('image/jpeg', quality);
          
          const originalSizeKb = Math.round(file.size / 1024);
          const head = 'data:image/jpeg;base64,';
          const compressedSizeBytes = Math.round((dataUrl.length - head.length) * 3 / 4);
          const compressedSizeKb = Math.round(compressedSizeBytes / 1024);
          const savingPercent = originalSizeKb > 0 
            ? Math.max(0, Math.round(((originalSizeKb - compressedSizeKb) / originalSizeKb) * 100))
            : 0;

          resolve({ dataUrl, originalSizeKb, compressedSizeKb, savingPercent });
        };
        img.onerror = (err) => reject(err);
        img.src = e.target?.result as string;
      };
      reader.onerror = (err) => reject(err);
      reader.readAsDataURL(file);
    });
  };

  const [questionForm, setQuestionForm] = useState<Partial<Question>>({
    kisiKisiId: '',
    kompetensi: '',
    subKompetensi: '',
    bentukSoal: 'pilihan_ganda_sederhana',
    soal: '',
    stimulus: '',
    opsi: ['', '', '', '', ''],
    kunciJawaban: '',
    pembahasan: '',
    kataKunci: '',
    gambarUrl: '',
    gambarCaption: '',
    gambarPosisi: 'center',
    gambarUkuran: 'medium'
  });

  // Prompt Generator States
  const [isPromptModalOpen, setIsPromptModalOpen] = useState(false);
  const [selectedKisiForPrompt, setSelectedKisiForPrompt] = useState<KisiKisiItem | null>(null);
  const [generatedPromptText, setGeneratedPromptText] = useState('');
  const [isGeneratingPrompt, setIsGeneratingPrompt] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);

  // Master Megaprompt Generator States (For entire Kisi-Kisi matrix)
  const [isMasterMegapromptModalOpen, setIsMasterMegapromptModalOpen] = useState(false);
  const [masterMegapromptStyle, setMasterMegapromptStyle] = useState<'pusmendik' | 'hots' | 'snbt' | 'variasi'>('pusmendik');
  const [masterMegapromptText, setMasterMegapromptText] = useState('');
  const [copiedMasterMegaprompt, setCopiedMasterMegaprompt] = useState(false);
  const [syncedToWadah1, setSyncedToWadah1] = useState(false);

  const buildMasterMegaprompt = (
    items: KisiKisiItem[], 
    cfg: GeneratorConfig, 
    stylePreset: 'pusmendik' | 'hots' | 'snbt' | 'variasi' = 'pusmendik'
  ) => {
    if (!items || items.length === 0) return '';

    const totalBaris = items.length;
    const totalSoal = items.reduce((acc, k) => acc + (k.jumlahSoal || 1), 0);
    
    // Counts
    const countPG = items.filter(k => k.bentukSoal === 'pilihan_ganda_sederhana').reduce((acc, k) => acc + (k.jumlahSoal || 1), 0);
    const countMCMA = items.filter(k => k.bentukSoal === 'mcma').reduce((acc, k) => acc + (k.jumlahSoal || 1), 0);
    const countKategori = items.filter(k => k.bentukSoal === 'kategori').reduce((acc, k) => acc + (k.jumlahSoal || 1), 0);

    const countTunggal = items.filter(k => (k.jenisSoal || 'tunggal') === 'tunggal').length;
    const countGrup = items.filter(k => k.jenisSoal === 'grup').length;

    const countL1 = items.filter(k => k.levelKognitif === 'level_1').length;
    const countL2 = items.filter(k => k.levelKognitif === 'level_2').length;
    const countL3 = items.filter(k => k.levelKognitif === 'level_3').length;

    let styleInstruction = "";
    if (stylePreset === 'pusmendik') {
      styleInstruction = "Gunakan standar baku Pusmendik Kemendikbudristek: Stimulus edukatif, opsi pilihan ganda berimbang, dan bahasa Indonesia baku sesuai PUEBI.";
    } else if (stylePreset === 'hots') {
      styleInstruction = "Fokus pada Higher Order Thinking Skills (HOTS): Sajikan studi kasus empiris, grafik/data nyata, analisis kritis, serta pertanyaan yang menuntut penalaran mendalam.";
    } else if (stylePreset === 'snbt') {
      styleInstruction = "Standar UTBK-SNBT / TKA: Soal menuntut kemampuan literasi tinggi, penyelesaian masalah (problem solving), dan kejelian menganalisis fenomena.";
    } else if (stylePreset === 'variasi') {
      styleInstruction = "Kombinasi Variatif: Campuran tingkat kesulitan (Mudah, Sedang, HOTS) dengan gaya naratif kontekstual dan skenario kehidupan sehari-hari.";
    }

    let prompt = `⚡ PERSYARATAN MUTLAK KELUARAN AI (2 FORMAT FILE UTUH SIAP UNDUH / DIPAKAI):
================================================================================
Anda WAJIB dan MUTLAK menyajikan SELURUH HASIL GENERASI SOAL dalam DUA FORMAT OUTPUT UTUH LENGKAP BERURUTAN DALAM SATU BALASAN:
1️⃣ [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
2️⃣ [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]

DILARANG KERAS HANYA MEMBUAT NASKAH WORD ATAU HANYA TABEL EXCEL! KEDUANYA WAJIB DITULISKAN SELURUHNYA SEHINGGA PENGGUNA DAPAT LANGSUNG MEMAKAI DAN MENGAMBIL 2 FORMAT FILE SEKALIGUS (WORD & EXCEL)!
================================================================================

MEGAPROMPT UTAMA AI: PENYUSUNAN PAKET SOAL ASESMEN SUMATIF PUSMENDIK / TKA
================================================================================
MATA PELAJARAN : ${cfg.mataPelajaran || 'Sosiologi'}
FASE / KELAS   : SMA Kelas XI - XII (Fase F)
TOTAL KISI-KISI: ${totalBaris} Baris Spesifikasi
TOTAL SOAL     : ${totalSoal} Butir Soal Keseluruhan
GAYA PENULISAN : ${styleInstruction}
================================================================================

PERAN DAN INSTRUKSI UTAMA (SYSTEM PROMPT):
Anda adalah Tim Ahli Pembuat Soal Ujian Nasional, Tim Pengembang TKA (Tes Kemampuan Akademik) SMA, dan Pakar Evaluasi Pembelajaran Kemendikbudristek (Pusmendik).
Tugas Anda adalah merancang dan menyusun paket soal asesmen sumatif sebanyak ${totalSoal} butir soal secara LENGKAP, BEBAS HALLUCINATION, dan PRESISI berdasarkan ${totalBaris} baris Matriks Asesmen / Kisi-Kisi berikut.

⚡ PERSYARATAN MUTLAK FORMAT KELUARAN (2 FILE / HASIL DOKUMEN UTUH SIAP PAKAI):
Anda WAJIB DAN MUTLAK menyajikan SELURUH HASIL GENERASI SOAL dalam DUA FORMAT OUTPUT UTUH KETAT BERURUTAN DALAM SATU BALASAN:
1️⃣ [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
2️⃣ [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]

DILARANG KERAS HANYA MEMBUAT NASKAH WORD ATAU HANYA TABEL EXCEL. KEDUANYA WAJIB DITULISKAN SELURUHNYA SECARA LENGKAP, PRESISI, DAN TIDAK BOLEH TERPOTONG DI TENGAH JALAN!

IKHTISAR SPESIFIKASI ASESMEN:
- Total Baris Kisi-Kisi: ${totalBaris} Baris Spesifikasi
- Total Target Soal Keseluruhan: ${totalSoal} Butir
- Klasifikasi Jenis Struktur Soal:
  * Soal Tunggal (Berdiri Sendiri / Standalone) : ${countTunggal} Baris Kisi-Kisi
  * Soal Grup (Sekumpulan Soal 1 Stimulus Wacana Bersama) : ${countGrup} Baris Kisi-Kisi
- Rincian Bentuk Soal:
  * Pilihan Ganda Sederhana (PG 5 Opsi): ${countPG} Butir
  * Pilihan Ganda Kompleks (MCMA / Banyak Jawaban): ${countMCMA} Butir
  * Kategori / Menjodohkan / Pernyataan: ${countKategori} Butir
- Sebaran Level Kognitif:
  * Level 1 - Pemahaman (Knowing): ${countL1} Kisi-Kisi
  * Level 2 - Penerapan (Applying): ${countL2} Kisi-Kisi
  * Level 3 - Penalaran (Reasoning / HOTS): ${countL3} Kisi-Kisi

--------------------------------------------------------------------------------
MATRIKS ASESMEN & RINCIAN SPESIFIKASI SOAL (${totalBaris} BARIS):
--------------------------------------------------------------------------------
`;

    items.forEach((kisi, idx) => {
      const activeKonteks = (kisi.konteksLokal && kisi.konteksLokal.length > 0) ? kisi.konteksLokal : cfg.konteksLokal;
      const activeStimulus = (kisi.stimulusKonten && kisi.stimulusKonten.length > 0) ? kisi.stimulusKonten : cfg.stimulusKonten;
      const isGrup = kisi.jenisSoal === 'grup';

      prompt += `
[KISI-KISI NO. ${kisi.no || (idx + 1)}]
- Jenis Struktur Soal  : ${isGrup ? '📚 SOAL GRUP (Sekumpulan butir soal mengacu pada 1 STIMULUS BERSAMA/WACANA TERPADU)' : '📌 SOAL TUNGGAL (Soal mandiri berdiri sendiri dengan 1 stimulus khusus)'}
- Elemen / Materi Utama : ${kisi.elemenMateri || '-'}
- Sub-Elemen / Submateri: ${kisi.subElemenMateri || '-'}
- Bentuk Soal          : ${getBentukSoalLabel(kisi.bentukSoal)} (${kisi.jumlahSoal || 1} Soal)
- Level Kognitif       : ${getLevelKognitifLabel(kisi.levelKognitif)}
- Indikator / Diuji    : ${kisi.kompetensi || '-'}
- Batasan & Catatan    : ${kisi.batasanCatatan || 'Patuhi kurikulum resmi'}
- Konteks Lokal ID     : ${activeKonteks.length > 0 ? activeKonteks.join(", ") : 'Masyarakat & Budaya Indonesia'}
- Stimulus Konten      : ${activeStimulus.length > 0 ? activeStimulus.join(", ") : 'Teks Kasus / Data Relevan'}
`;
    });

    prompt += `
--------------------------------------------------------------------------------
PETUNJUK PENULISAN DAN FORMAT OUTPUT KETAT (DOKUMEN WORD & TABEL EXCEL):
--------------------------------------------------------------------------------
FORMAT PENAMAAN FILE DOWNLOAD / DOKUMEN HASIL GENERASI:
Wajib mencantumkan Judul / Nama File Dokumentasi Utama di bagian awal dengan format baku:
[Mata Pelajaran]_[Materi Pokok]_[Bentuk Soal]
(Contoh: ${cfg.mataPelajaran || 'Sosiologi'}_${cfg.elemenMateri || 'Materi_Pokok'}_MCMA, ${cfg.mataPelajaran || 'Sosiologi'}_${cfg.elemenMateri || 'Materi_Pokok'}_Kategori, atau ${cfg.mataPelajaran || 'Sosiologi'}_${cfg.elemenMateri || 'Materi_Pokok'}_Sederhana).

Sajikan seluruh hasil generasi soal secara LENGKAP, BEBAS POTONGAN, dan PRESISI dalam 2 FORMAT OUTPUT TERTATA KETAT:

[BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
Sajikan seluruh ${totalSoal} butir soal berurutan dari Kisi-Kisi No. 1 s.d. No. ${totalBaris} dalam bentuk naskah cetak bernomor rapi yang dapat langsung disalin-rekatkan (copy-paste) ke Microsoft Word:
1. SPESIFIKASI SOAL:
   - Nomor Soal, Materi Utama, Submateri, Level Kognitif (L1/L2/L3), dan Bentuk Soal.
2. STIMULUS LENGKAP & PENGEMBANGAN KONTEN VISUAL (MANDATORI):
   - Wajib menyajikan STIMULUS VISUAL ATAU BERBASIS DATA KONKRET kaya informasi analitis untuk setiap soal.
   - Sajikan stimulus dalam salah satu atau kombinasi bentuk visual berikut:
     a) TABEL DATA & MATRIKS STATISTIK (Markdown Table): Tabel empiris berkolom rapi memuat angka/fakta terukur.
     b) DIAGRAM ALUR / FLOWCHART / BAGAN KONSEP (ASCII Diagram): Diagram proses, siklus, atau hirarki dalam format code block ASCII (contoh: [Sebab/Input] ──► [Proses/Interaksi] ──► [Dampak/Output]).
     c) GRAFIK TREN TEKSTUAL (Visual Bar Chart): Diagram batang tekstual berskala (contoh: 2022: ██████████ 50% | 2023: ██████████████ 70%).
     d) SPESIFIKASI INFOGRAFIS / GAMBAR / PETA (Visual Specification): Blok terstruktur [DESKRIPSI INFOGRAFIS / SPESIFIKASI VISUAL: Judul] yang merinci elemen visual, tata letak, teks, serta data statistik secara lengkap.
     e) KUTIPAN DOKUMEN & DIALOG BERFORMAT (Blockquote / Dialog): Format kutipan berbingkai (> ...) atau percakapan berdialog untuk merepresentasikan naskah asli.
3. BUTIR PERTANYAAN & OPSI JAWABAN:
   - Pilihan Ganda Sederhana: Pertanyaan jelas + 5 Opsi (A, B, C, D, E). Pastikan distribusi kunci jawaban (A, B, C, D, E) tersebar proporsional.
   - Pilihan Ganda Kompleks (MCMA): Pertanyaan + 4-5 opsi/pernyataan terpisah (A, B, C, D, E) + instruksi memilih semua jawaban benar.
   - Pilihan Ganda Kompleks Kategori (PGK Kategori): 
     Peserta diminta untuk memberikan respon/pilihan pada masing-masing pernyataan (minimal 2, maksimal 4 pernyataan). Label/kategori respon TIDAK TERBATAS HANYA pada "Sesuai / Tidak Sesuai" atau "Benar / Salah", tetapi BISA MENGGUNAKAN PASANGAN KATEGORI LAIN yang relevan dengan materi/konteks (misalnya: "Tepat / Tidak Tepat", "Ya / Tidak", "Fakta / Opini", "Setuju / Tidak Setuju", "Masuk Akal / Tidak Masuk Akal", "Positif / Negatif", dll.).
     WAJIB MENGGUNAKAN STRUKTUR EKSPLISIT SESUAI FORMAT CBT:
     * SOAL (Stimulus + Pertanyaan Utama + Kalimat Perintah):
       [Narasi / Stimulus Kasus, contoh: Terdapat dua peneliti sosial yang ingin memahami dampak...]
       [Pertanyaan Utama, contoh: Berdasarkan pernyataan diatas, manakah pengolahan data yang sesuai dengan ilustrasi tersebut?]
       [Kalimat Perintah Kategori, contoh: **Berikan respon/pilihan Anda pada masing-masing pernyataan berikut!** (atau **Pilihlah Sesuai atau Tidak Sesuai / Fakta atau Opini pada setiap pernyataan!**)]

     * TABEL PERNYATAAN (Tabel Markdown 4 Kolom dengan opsi checklist/radio, MINIMAL 2 PERNYATAAN & MAKSIMAL 4 PERNYATAAN):
       | # | Pernyataan | Sesuai | Tidak Sesuai |
       |:---:|---|:---:|:---:|
       | 1 | [Teks Pernyataan 1] | | |
       | 2 | [Teks Pernyataan 2] | | |
       | 3 | [Teks Pernyataan 3] | | |
       | 4 | [Teks Pernyataan 4] | | |
       *(Urutan nomor wajib ANGKA 1, 2, 3, 4. Jumlah pernyataan minimal 2 dan maksimal 4. Peserta diminta memberikan respon untuk masing-masing pernyataan. Label kolom kategori disesuaikan secara fleksibel: Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Ya/Tidak, Fakta/Opini, Setuju/Tidak Setuju, dll.)*

     * KUNCI JAWABAN:
       Pernyataan 1: Sesuai
       Pernyataan 2: Tidak Sesuai
       Pernyataan 3: Sesuai
       Pernyataan 4: Sesuai

     * PEMBAHASAN:
       1. [Penjelasan analitis ilmiah mengapa Pernyataan 1 berstatus Sesuai]
       2. [Penjelasan analitis ilmiah mengapa Pernyataan 2 berstatus Tidak Sesuai]
       3. [Penjelasan analitis ilmiah mengapa Pernyataan 3 berstatus Sesuai]
       4. [Penjelasan analitis ilmiah mengapa Pernyataan 4 berstatus Sesuai]
4. KUNCI JAWABAN & PEMBAHASAN MENDALAM:
   - Kunci Jawaban Tepat + Pembahasan Komprehensif (Rationale ilmiah/analitis) serta penjelasan mengapa opsi lain tidak tepat.
5. PEDOMAN PENSKORAN & RUBRIK (SKALA 0-100):
   - Sertakan bobot nilai per bentuk soal (misal PG = 1 poin, MCMA = 2 poin, Kategori = 2 poin) dan rumus konversi skor total ke nilai 100.
${(isExactScienceSubject(cfg.mataPelajaran) || isMatrixOrMathTopic(cfg.mataPelajaran, cfg.elemenMateri, cfg.subElemenMateri, cfg.kompetensi) || (kisiList && kisiList.some(k => isMatrixOrMathTopic(cfg.mataPelajaran, k.elemenMateri, k.subElemenMateri, k.kompetensi)))) ? `
6. STANDAR KHUSUS PERSAMAAN, SIMBOL & RUMUS EKSAKTA / FISIKA / MATEMATIKA / KIMIA:
   - **Notasi LaTeX Baku**: WAJIB menuliskan semua persamaan matematika, rumus fisika, dan reaksi kimia dengan notasi LaTeX ($ ... $ untuk inline equation dan $$ ... $$ untuk display math) atau format Unicode ilmiah standar. Contoh: $v_t = v_0 + a \\cdot t$, $s = v_0 t + \\frac{1}{2} a t^2$, $F = m \\cdot a$, $E_k = \\frac{1}{2} m v^2$, $E_p = m g h$, $W = F \\cdot s \\cdot \\cos\\theta$, $P = \\frac{W}{t}$, $I = \\frac{V}{R}$, $\\rho = \\frac{m}{V}$, pecahan wajib \\frac{a}{b}, akar \\sqrt{x}, dan pangkat x^2 atau 10^{-3}.
   - **Simbol & Satuan Ilmiah Baku**: Wajib mencantumkan besaran dan satuan fisika/kimia yang tepat (seperti m/s², kg, J, W, Hz, Ω, μC, °C, K, ±, ×, ÷, ≈, ≤, ≥, ≠, ∞, Δ, α, β, θ, λ, μ, π, ρ, σ, ω).
   - **Formula & Reaksi Kimia**: Tuliskan formula kimia dengan subskrip/supersprip yang tepat (contoh: \\text{H}_2\\text{SO}_4, \\text{Fe}^{3+}, \\text{SO}_4^{2-}), tanda panah reaksi (\\rightarrow, \\rightleftharpoons), dan fasa zat ((s), (l), (g), (aq)).
   - **STANDAR KHUSUS SOAL MATRIKS & KOLOM MATEMATIKA (MANDATORI UNTUK AI & LLM)**:
     * **Notasi Kolom & Baris Matriks LaTeX**: WAJIB menuliskan seluruh matriks menggunakan notasi LaTeX matriks standar dengan pemisah kolom tanda ampersand (\`&\`) dan pemisah baris (\`\\\\\`):
       - Matriks Kurung Biasa: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$
       - Matriks Kurung Siku: $\\begin{bmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{bmatrix}$
       - Determinan Matriks: $\\begin{vmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{vmatrix}$ atau $\\det(A)$
       - Matriks Ordo 3x3: $\\begin{pmatrix} a_{11} & a_{12} & a_{13} \\\\ a_{21} & a_{22} & a_{23} \\\\ a_{31} & a_{32} & a_{33} \\end{pmatrix}$
       - Matriks Kolom / Vektor: $\\begin{pmatrix} x \\\\ y \\end{pmatrix}$
     * **Presisi Kolom & Operasi Matriks**: DILARANG KERAS menuliskan elemen matriks rapat tanpa pemisah kolom \`&\` atau tanpa baris \`\\\\\`. Operasi penjumlahan, perkalian, invers ($A^{-1}$), determinan, transpose ($A^T$), dan sistem persamaan linear matriks ($AX = B$) harus memuat setiap elemen baris dan kolom yang valid.
     * **Format Pilihan Jawaban (A s.d E)**: Setiap opsi jawaban yang memuat matriks WAJIB disajikan dalam format LaTeX matriks yang utuh (Contoh: A. $\\begin{pmatrix} 2 & 3 \\\\ -1 & 4 \\end{pmatrix}$).
     * **Keamanan Kolom Tabel Excel / Markdown (Bagian 2)**: Pada tabel rekapitulasi Markdown/Excel (Bagian 2), dilarang menggunakan karakter pipa \`|\` tunggal di dalam sel matriks (gunakan $\\begin{pmatrix}...\\end{pmatrix}$ atau $\\begin{bmatrix}...\\end{bmatrix}$) agar struktur kolom tabel Excel/Markdown tetap presisi.
   - **Format Tabel Data / Eksperimen / Stimulus Numerik**: WAJIB menyajikan data eksperimen dalam format Markdown Table standar yang utuh 1 baris per record dengan header pemisah (|---|---|) yang presisi.
   - **Homogenitas Opsi Pilihan Jawaban**: Opsi A s.d E wajib homogen dalam format angka, pecahan, desimal, atau penyertaan satuan.` : ''}

[BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]
Di bagian paling akhir (setelah seluruh Naskah Word selesai), Anda WAJIB menyajikan TABEL REKAPITULASI BENTUK MARKDOWN / TSV CODEBLOCK yang siap di-copy-paste langsung ke Microsoft Excel atau Google Sheets.
Tabel harus memuat seluruh ${totalSoal} soal dengan struktur kolom baku Excel sebagai berikut:

| No Soal | Elemen / Materi Utama | Submateri | Level Kognitif | Bentuk Soal | Ringkasan Stimulus | Pertanyaan / Teks Soal | Opsi_A / Pernyataan_1 | Opsi_B / Pernyataan_2 | Opsi_C / Pernyataan_3 | Opsi_D / Pernyataan_4 | Opsi_E | Kunci Jawaban | Pembahasan Singkat |

ATURAN MANDATORI PEMISAHAN OPSI EXCEL (KHUSUS MCMA & KATEGORI):
- Untuk PG Kompleks (MCMA), pilihan jawaban dipisahkan ke kolom Opsi_A s.d Opsi_E.
- Khusus untuk Pilihan Ganda Kompleks Kategori (PGK Kategori), daftar pernyataan (minimal 2, maksimal 4) WAJIB dipisahkan ke masing-masing kolom terpisah yaitu Pernyataan_1, Pernyataan_2, Pernyataan_3, dan Pernyataan_4 (atau Opsi_1, Opsi_2, Opsi_3, Opsi_4). Kunci Jawaban merinci status per nomor (contoh: Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, Pernyataan 3: Sesuai, Pernyataan 4: Sesuai). DILARANG KERAS menggabungkan seluruh daftar pernyataan ke dalam 1 sel Excel!

--------------------------------------------------------------------------------
PERINGATAN FINAL MANDATORI DUA HASIL KELUARAN (WORD & EXCEL):
- WAJIB DAN MUTLAK MENYAJIKAN DUA BAGIAN DOKUMEN DALAM BALASAN ANDA:
  1) [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
  2) [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]
- DILARANG BERHENTI ATAU MENGAKHIRI CHAT SEBELUM BAGIAN 2 (TABEL MATRIKS REKAPITULASI EXCEL) SELESAI DITULISKAN SELURUHNYA SEHINGGA PENGGUNA DAPAT MENGUNDUH/MENGAMBIL 2 FORMAT FILE SEKALIGUS (WORD DAN EXCEL)!
--------------------------------------------------------------------------------

Sila buatkan seluruh paket ${totalSoal} butir soal secara berurutan, rapi, dan komprehensif dari Kisi-Kisi No. 1 hingga No. ${totalBaris} memenuhi FORMAT WORD dan FORMAT EXCEL di atas!
`;

    return prompt.trim();
  };

  const downloadMasterMegapromptAsTxt = () => {
    const text = masterMegapromptText || buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
    const element = document.createElement("a");
    const file = new Blob([text], { type: 'text/plain;charset=utf-8' });
    element.href = URL.createObjectURL(file);
    element.download = `Megaprompt_Utama_${(config.mataPelajaran || 'Sosiologi').replace(/\s+/g, '_')}_${kisiList.length}_Baris.txt`;
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const handleSyncToWadah1 = () => {
    const text = masterMegapromptText || buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
    setPromptWadah1Text(text);
    setSyncedToWadah1(true);
    setTimeout(() => setSyncedToWadah1(false), 3000);
  };

  // --- STATE FOR SECTION 4: PEMBUATAN MATERI & PANDUAN ---
  const [uploadedPdf, setUploadedPdf] = useState<{ name: string; size: number } | null>(null);
  const [isUploadingPdf, setIsUploadingPdf] = useState(false);
  const [uploadedPdfStatus, setUploadedPdfStatus] = useState<string>('');
  const [guidanceContext, setGuidanceContext] = useState<string>('');
  const [activeMateriKisiId, setActiveMateriKisiId] = useState<string | null>(null);
  const [generatingMateriIds, setGeneratingMateriIds] = useState<Record<string, boolean>>({});
  const [generatedMaterials, setGeneratedMaterials] = useState<Record<string, { content?: string; prompt?: string }>>({});
  const [activeSubTab, setActiveSubTab] = useState<'materi' | 'prompt'>('materi');

  // Automatically select the first Kisi-Kisi on load or keep selected one valid
  useEffect(() => {
    if (kisiList.length > 0) {
      if (!activeMateriKisiId || !kisiList.some(k => k.id === activeMateriKisiId)) {
        setActiveMateriKisiId(kisiList[0].id);
      }
    } else {
      setActiveMateriKisiId(null);
    }
  }, [kisiList, activeMateriKisiId]);

  const [isEditingMateri, setIsEditingMateri] = useState<boolean>(false);
  const [editingMateriContent, setEditingMateriContent] = useState<string>('');

  const handlePdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.type !== 'application/pdf') {
      alert("Silakan unggah dokumen dalam format PDF saja.");
      return;
    }

    setIsUploadingPdf(true);
    setUploadedPdfStatus('Sedang membaca file PDF...');

    setTimeout(() => {
      setUploadedPdf({
        name: file.name,
        size: file.size
      });
      
      let extractedText = "";
      if (config.mataPelajaran.toLowerCase().includes('sosiologi')) {
        extractedText = "Rekomendasi kurikulum sosiologi menyarankan integrasi wacana empiris berbasis kearifan lokal di Indonesia, fokus pada pengenalan interaksi sosial digital, sosiologi perkotaan, serta pentingnya penguasaan dasar metodologi riset (kuantitatif, kualitatif, mixed methods). Evaluasi diorientasikan pada tingkat HOTS (C4-C6).";
      } else if (config.mataPelajaran.toLowerCase().includes('matematika')) {
        extractedText = "Panduan kurikulum matematika menekankan penguatan literasi numerasi, pemecahan masalah (problem solving), eksplorasi fungsi-fungsi kuadrat dan statistika deskriptif/inferensial, serta penerapan penalaran matematis dalam kehidupan sehari-hari secara rasional.";
      } else {
        extractedText = `Panduan Pembelajaran resmi untuk mata pelajaran ${config.mataPelajaran} menyarankan kesesuaian materi ajar dengan kompetensi dasar, penekanan pada literasi konten, stimulasi berpikir kritis melalui studi kasus nyata, dan kebebasan guru menyusun asesmen sesuai level kesiapan siswa.`;
      }
      
      setGuidanceContext(extractedText);
      setIsUploadingPdf(false);
      setUploadedPdfStatus('File PDF berhasil terunggah dan terintegrasi secara semantis dengan AI!');
    }, 1500);
  };

  const handleGenerateMateri = async (kisi: KisiKisiItem, modeOverride?: 'materi' | 'prompt') => {
    const mode = modeOverride || activeSubTab;
    setGeneratingMateriIds(prev => ({ ...prev, [kisi.id]: true }));
    try {
      const response = await fetch('/api/generate-materi', {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'x-api-key': getCleanApiKey(aiConfig.apiKey)
        },
        body: JSON.stringify({
          kisi,
          mataPelajaran: config.mataPelajaran,
          guidanceText: guidanceContext || undefined,
          mode
        })
      });
      if (!response.ok) {
        throw new Error('Gagal menghubungi AI Server.');
      }
      const data = await response.json();
      if (data.materi) {
        const updateField = mode === 'materi' ? 'content' : 'prompt';
        await setDoc(doc(db, 'materials', kisi.id), {
          [updateField]: data.materi,
          userId: currentUser?.uid,
          updatedAt: new Date()
        }, { merge: true });
        setActiveMateriKisiId(kisi.id);
      } else {
        alert("AI mengembalikan format yang tidak sesuai.");
      }
    } catch (err: any) {
      console.error(err);
      alert(`Gagal membuat ${mode === 'materi' ? 'materi pembelajaran' : 'prompt slide & infografis'} secara dinamis: ` + (err.message || err));
    } finally {
      setGeneratingMateriIds(prev => ({ ...prev, [kisi.id]: false }));
    }
  };

  const [copiedMateriKisiId, setCopiedMateriKisiId] = useState<string | null>(null);

  const copyToClipboard = async (text: string): Promise<boolean> => {
    if (!text) return false;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (err) {
      console.warn("navigator.clipboard failed:", err);
    }
    try {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      textArea.style.position = "fixed";
      textArea.style.left = "-999999px";
      textArea.style.top = "-999999px";
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      const successful = document.execCommand('copy');
      textArea.remove();
      return successful;
    } catch (fallbackErr) {
      console.error("Fallback copy failed:", fallbackErr);
      return false;
    }
  };

  const handleDeleteMateri = async (kisiId: string) => {
    const label = activeSubTab === 'materi' ? 'Ringkasan Materi' : 'Prompt Slide & Infografis';
    if (!confirm(`Apakah Anda yakin ingin menghapus ${label} untuk kisi-kisi ini?`)) return;
    try {
      const updateField = activeSubTab === 'materi' ? 'content' : 'prompt';
      
      // Update local state immediately so UI refreshes without delay
      setGeneratedMaterials(prev => {
        const existing = prev[kisiId] || {};
        return {
          ...prev,
          [kisiId]: {
            ...existing,
            [updateField]: ''
          }
        };
      });

      if (currentUser?.uid) {
        await setDoc(doc(db, 'materials', kisiId), {
          [updateField]: '',
          userId: currentUser.uid,
          updatedAt: new Date()
        }, { merge: true });
      }

      setIsEditingMateri(false);
      alert(`${label} berhasil dihapus!`);
    } catch (err: any) {
      console.error(err);
      alert(`Gagal menghapus ${label}: ` + (err.message || err));
    }
  };

  const handleSaveMateri = async (kisiId: string, text: string) => {
    const label = activeSubTab === 'materi' ? 'Ringkasan Materi' : 'Prompt Slide & Infografis';
    try {
      const updateField = activeSubTab === 'materi' ? 'content' : 'prompt';
      await setDoc(doc(db, 'materials', kisiId), {
        [updateField]: text,
        userId: currentUser?.uid,
        updatedAt: new Date()
      }, { merge: true });
      setIsEditingMateri(false);
      alert(`${label} berhasil disimpan!`);
    } catch (err: any) {
      console.error(err);
      alert(`Gagal menyimpan ${label}: ` + (err.message || err));
    }
  };

  const handleUploadPromptFile = (e: React.ChangeEvent<HTMLInputElement>, kisiId: string) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const label = activeSubTab === 'materi' ? 'Ringkasan Materi' : 'Prompt Slide & Infografis';
    const reader = new FileReader();
    reader.onload = async (event) => {
      const text = event.target?.result as string;
      if (text) {
        try {
          const updateField = activeSubTab === 'materi' ? 'content' : 'prompt';
          await setDoc(doc(db, 'materials', kisiId), {
            [updateField]: text,
            userId: currentUser?.uid,
            updatedAt: new Date()
          }, { merge: true });
          alert(`File ${label} berhasil diunggah!`);
        } catch (err: any) {
          console.error(err);
          alert(`Gagal mengunggah file ${label}: ` + (err.message || err));
        }
      }
    };
    reader.readAsText(file);
  };

  const handlePrintMateri = (kisi: KisiKisiItem, content: string) => {
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      alert("Popup blocker aktif. Harap izinkan popup untuk melakukan pencetakan.");
      return;
    }

    const parsedHtml = markdownToHtmlForWord(content);
    
    // Style the printed page to be extremely professional
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Modul_Ajar_${kisi.no}_${kisi.elemenMateri.replace(/[^a-zA-Z0-9]/g, '_')}</title>
        <style>
          @media print {
            body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .no-print { display: none; }
          }
          @page {
            size: A4;
            margin: 2cm;
          }
          body {
            font-family: 'Times New Roman', 'Georgia', Times, serif;
            color: #111827;
            line-height: 1.6;
            font-size: 12pt;
            background: white;
            margin: 0;
            padding: 0;
          }
          .header-kop {
            border-bottom: 4px double #111827;
            padding-bottom: 12px;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            gap: 20px;
          }
          .kop-logo {
            width: 70px;
            height: 70px;
            object-fit: contain;
          }
          .kop-logo-placeholder {
            width: 70px;
            height: 70px;
            border-radius: 50%;
            border: 2px solid #111827;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 11px;
            font-weight: bold;
            color: #111827;
          }
          .kop-text {
            flex: 1;
            text-align: center;
          }
          .kop-dept {
            font-size: 9.5pt;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin: 0 0 2px 0;
            color: #374151;
          }
          .kop-school {
            font-size: 15pt;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin: 0 0 4px 0;
            color: #111827;
          }
          .kop-info {
            font-size: 9pt;
            font-style: italic;
            color: #4b5563;
            margin: 0;
          }
          .title {
            color: #111827;
            font-family: 'Times New Roman', 'Georgia', Times, serif;
            font-size: 16pt;
            font-weight: normal;
            margin: 0 0 4px 0;
            text-align: center;
            text-transform: none;
          }
          .subtitle {
            font-size: 11pt;
            color: #374151;
            margin: 0;
            text-align: center;
            font-weight: normal;
            font-style: italic;
          }
          .meta-box {
            background-color: #f9fafb;
            border: 1.5px solid #111827;
            border-radius: 4px;
            padding: 14px 18px;
            margin-bottom: 24px;
          }
          .meta-table {
            width: 100%;
            border-collapse: collapse;
          }
          .meta-table td {
            padding: 5px 0;
            font-size: 11pt;
            color: #111827;
            vertical-align: top;
          }
          .meta-label {
            font-weight: bold;
            width: 25%;
            color: #111827;
          }
          .meta-separator {
            width: 3%;
            color: #111827;
          }
          .meta-value {
            width: 72%;
            font-weight: 500;
          }
          .content h1 {
            font-size: 14pt;
            color: #111827;
            border-bottom: 1.5pt solid #111827;
            padding-bottom: 4px;
            margin-top: 24pt;
            margin-bottom: 10pt;
            font-weight: normal;
            text-transform: none;
            page-break-after: avoid;
          }
          .content h2 {
            font-size: 13pt;
            color: #111827;
            margin-top: 18pt;
            margin-bottom: 8pt;
            font-weight: normal;
            border-left: 3pt solid #111827;
            padding-left: 8px;
            page-break-after: avoid;
          }
          .content h3 {
            font-size: 12pt;
            color: #111827;
            margin-top: 14pt;
            margin-bottom: 6pt;
            font-weight: normal;
            font-style: italic;
            page-break-after: avoid;
          }
          .content p {
            margin-top: 0;
            margin-bottom: 10pt;
            text-align: justify;
            text-justify: inter-word;
            text-indent: 0.5in;
          }
          .content ul, .content ol {
            margin-top: 0;
            margin-bottom: 10pt;
            padding-left: 24px;
          }
          .content li {
            margin-bottom: 6px;
            text-align: justify;
          }
          .content blockquote {
            border-left: 3.5pt solid #111827;
            background-color: #f9fafb;
            padding: 12px 18px;
            margin: 14pt 0;
            font-style: italic;
            color: #374151;
            border-radius: 0;
            text-align: justify;
          }
          .content table {
            width: 100%;
            border-collapse: collapse;
            margin: 16pt 0;
            border: 1.5pt solid #111827;
            page-break-inside: avoid;
          }
          .content th {
            background-color: #1e3a8a !important;
            color: #ffffff !important;
            font-weight: bold;
            padding: 8pt 10pt;
            border: 1pt solid #1e3a8a;
            text-align: left;
            font-size: 11pt;
            font-family: 'Times New Roman', serif;
          }
          .content td {
            padding: 8pt 10pt;
            border: 1pt solid #cbd5e1;
            font-size: 11pt;
            vertical-align: top;
            font-family: 'Times New Roman', serif;
          }
          .content tr:nth-child(even) {
            background-color: #f8fafc;
          }
          .print-btn-container {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: #4f46e5;
            color: white;
            padding: 10px 20px;
            border-radius: 9999px;
            font-weight: bold;
            box-shadow: 0 4px 10px rgba(0,0,0,0.15);
            cursor: pointer;
            z-index: 9999;
            font-family: inherit;
            border: none;
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
          }
          .print-btn-container:hover {
            background-color: #4338ca;
          }
        </style>
      </head>
      <body>
        <button class="print-btn-container no-print" onclick="window.print()">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"></path><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
          Cetak Dokumen / Simpan PDF
        </button>

        <div class="header-kop">
          ${printConfig.schoolLogo ? `<img src="${printConfig.schoolLogo}" class="kop-logo" />` : `<div class="kop-logo-placeholder">LOGO</div>`}
          <div class="kop-text">
            <p class="kop-dept">Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi</p>
            <h1 class="kop-school">${printConfig.schoolName || 'SEKOLAH MENENGAH ATAS'}</h1>
            <p class="kop-info">Tahun Pelajaran: ${printConfig.academicYear} | Semester: ${printConfig.semester.toUpperCase()}</p>
          </div>
          ${printConfig.schoolLogoRight ? `<img src="${printConfig.schoolLogoRight}" class="kop-logo" />` : `<div class="kop-logo-placeholder">SMA</div>`}
        </div>

        <div>
          <h1 class="title">Bahan Ajar / Modul Pembelajaran</h1>
          <p class="subtitle">Kurikulum Merdeka - Capaian & Kompetensi Mandiri</p>
        </div>

        <div class="meta-box" style="margin-top: 20px;">
          <table class="meta-table">
            <tr>
              <td class="meta-label">Mata Pelajaran</td>
              <td class="meta-separator">:</td>
              <td class="meta-value" style="color: #1e3a8a; font-weight: bold;">${config.mataPelajaran}</td>
            </tr>
            <tr>
              <td class="meta-label">Elemen Materi</td>
              <td class="meta-separator">:</td>
              <td class="meta-value">${kisi.elemenMateri}</td>
            </tr>
            <tr>
              <td class="meta-label">Sub-Materi</td>
              <td class="meta-separator">:</td>
              <td class="meta-value">${kisi.subElemenMateri}</td>
            </tr>
            <tr>
              <td class="meta-label">Kompetensi Inti</td>
              <td class="meta-separator">:</td>
              <td class="meta-value">${kisi.kompetensi}</td>
            </tr>
            <tr>
              <td class="meta-label">Tingkat Kognitif</td>
              <td class="meta-separator">:</td>
              <td class="meta-value">${getLevelKognitifLabel(kisi.levelKognitif)}</td>
            </tr>
          </table>
        </div>

        <div class="content">
          ${parsedHtml}
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 500);
          }
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();
  };

  // Check backend server status on mount
  useEffect(() => {
    fetch('/api/health')
      .then(res => res.json())
      .then(data => {
        if (data.status === 'ok') {
          setApiStatus('connected');
        } else {
          setApiStatus('error');
        }
      })
      .catch(() => {
        setApiStatus('disconnected');
      });
  }, []);

  // Sync prompts whenever config state changes
  useEffect(() => {
    const contextList = config.konteksLokal.length > 0 
      ? `\n- KONTEKS LOKAL INDONESIA: ${config.konteksLokal.join(', ')}`
      : '';
    const stimulusList = config.stimulusKonten.length > 0
      ? `\n- STIMULUS SOAL: ${config.stimulusKonten.join(', ')}`
      : '';
    const qualityList = config.kualitasChecklist.length > 0
      ? `\n- CHECKLIST STANDAR KUALITAS: ${config.kualitasChecklist.join(', ')}`
      : '';

    // 1. Prompt Kisi-kisi (Matriks Asesmen)
    const promptKisiText = `Anda adalah ahli kurikulum pendidikan menengah SMA di Indonesia dan spesialis penyusunan Tes Kemampuan Akademik (TKA) berstandar tinggi (HOTS).
Tugas Anda adalah merancang sebuah MATRIKS ASESMEN / KISI-KISI SOAL yang komprehensif, terstruktur, dan valid untuk mata pelajaran di bawah ini. Jadikan Hasil soal Outpot Format word dan excel.

INFORMASI MATA PELAJARAN & PARAMETER UTAMA:
- MATA PELAJARAN: ${config.mataPelajaran}
- DEFINISI/TUJUAN PEMBELAJARAN: ${config.definisi || 'Tidak ditentukan'}
- MUATAN/FASE KURIKULUM: ${config.muatan || 'Fase F (Kelas XI/XII)'}
- ELEMEN MATERI UTAMA: ${config.elemenMateri || 'Tidak ditentukan'}
- SUB-ELEMEN/SUBMATERI: ${config.subElemenMateri || 'Tidak ditentukan'}
- KOMPETENSI OPERASIONAL: ${config.kompetensi || 'Mengidentifikasi, menganalisis, dan memecahkan masalah'}
- BATASAN MATERI & CATATAN: ${config.batasanCatatan || 'Tidak ada batasan khusus'}
- TINGKAT KOGNITIF DEFAULT: ${getLevelKognitifLabel(config.levelKognitif)} (${config.levelKognitif})
- BENTUK SOAL DEFAULT: ${getBentukSoalLabel(config.bentukSoal)} (${config.bentukSoal})${contextList}${stimulusList}${qualityList}

INSTRUKSI PENYUSUNAN MATRIKS:
1. Buatlah minimal 3 sampai 5 baris variasi kisi-kisi soal yang seimbang, logis, dan mencakup kedalaman materi yang diminta.
2. Setiap baris kisi-kisi harus memvariasikan aspek:
   - bentukSoal: wajib memilih salah satu dari:
     * 'pilihan_ganda_sederhana' (Pilihan Ganda Tunggal / Satu Pilihan Benar)
     * 'mcma' (Pilihan Ganda Kompleks / Lebih dari Satu Jawaban Benar)
     * 'kategori' (Pilihan Ganda Kompleks / Klasifikasi Benar-Salah atau Ya-Tidak)
   - levelKognitif: wajib memilih salah satu dari:
     * 'level_1' (Pemahaman / Knowing: Mengenali, mengingat, mendefinisikan)
     * 'level_2' (Penerapan / Applying: Mengaplikasikan konsep pada kasus nyata)
     * 'level_3' (Penalaran / Reasoning: Menganalisis, mensintesis, berpikir kritis/HOTS)
3. Deskripsikan Kompetensi spesifik yang akan diukur serta batasan materi khusus untuk tiap baris secara ilmiah, jelas, dan berorientasi HOTS.
4. Tentukan jumlah soal per baris secara proporsional (rekomendasi total 5-10 soal per baris).

Sajikan output Anda ke dalam DUA format berikut:

1. TABEL RINGKASAN (Untuk Tampilan Visual):
Sajikan dalam bentuk tabel Markdown yang rapi dengan kolom: No, Elemen, Sub-Elemen, Kompetensi diukur, Level Kognitif, Bentuk Soal, Batasan/Catatan, Jumlah Soal.

2. BLOK KODE JSON ARRAY (Untuk Kebutuhan Impor/Integrasi):
Tuliskan blok kode JSON valid (di dalam format \`\`\`json) yang berisi array of objects dengan struktur persis seperti contoh berikut (pastikan kunci/key tidak diubah dan nilai bentukSoal & levelKognitif mengikuti enum di atas):
\`\`\`json
[
  {
    "no": 1,
    "bentukSoal": "${config.bentukSoal}",
    "levelKognitif": "${config.levelKognitif}",
    "elemenMateri": "${config.elemenMateri || '[Elemen]'} ",
    "subElemenMateri": "${config.subElemenMateri || '[Sub Elemen]'} ",
    "kompetensi": "[Kompetensi operasional spesifik yang diukur]",
    "batasanCatatan": "${config.batasanCatatan || '[Batasan khusus]'} ",
    "jumlahSoal": ${config.jumlahSoal || 5}
  }
]
\`\`\``;

    // 2. Prompt Pembuat Soal (Megaprompt)
    const promptSoalText = `⚡ PERSYARATAN MUTLAK KELUARAN AI (2 FORMAT FILE UTUH SIAP UNDUH / DIPAKAI):
================================================================================
Anda WAJIB dan MUTLAK menyajikan SELURUH HASIL GENERASI SOAL dalam DUA FORMAT OUTPUT UTUH LENGKAP BERURUTAN DALAM SATU BALASAN:
1️⃣ [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
2️⃣ [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]

DILARANG KERAS HANYA MEMBUAT NASKAH WORD ATAU HANYA TABEL EXCEL! KEDUANYA WAJIB DITULISKAN SELURUHNYA SEHINGGA PENGGUNA DAPAT LANGSUNG MEMAKAI DAN MENGAMBIL 2 FORMAT FILE SEKALIGUS (WORD & EXCEL)!
================================================================================

PROMPT PEMBUAT SOAL ASESMEN TKA SMA:
Anda adalah seorang ahli penyusun soal TKA (Tes Kemampuan Akademik) SMA tingkat nasional dan pakar evaluasi kurikulum pendidikan di Indonesia.
Tugas Anda adalah merancang ${config.jumlahSoal || 20} butir soal ${getBentukSoalLabel(config.bentukSoal)} berorientasi HOTS (Higher Order Thinking Skills) untuk mata pelajaran ${config.mataPelajaran || 'Sosiologi'} tingkat SMA, ${config.muatan || 'Kelas XII'}.

INFORMASI SPESIFIKASI SOAL:
- MATA PELAJARAN: ${config.mataPelajaran}
- MATERI UTAMA/ELEMEN: ${config.elemenMateri || 'Tidak ditentukan'}
- SUB-ELEMEN/SUBMATERI: ${config.subElemenMateri || 'Tidak ditentukan'}
- KOMPETENSI UTAMA YANG DIUJI: ${config.kompetensi || 'Menganalisis dan memecahkan masalah'}
- TINGKAT KOGNITIF: ${getLevelKognitifLabel(config.levelKognitif)} (${config.levelKognitif})
- BENTUK SOAL: ${getBentukSoalLabel(config.bentukSoal)}
- PILIHAN JAWABAN: ${config.jumlahOpsi} Pilihan (A s.d ${config.jumlahOpsi === 5 ? 'E' : 'D'})
- JENIS STRUKTUR: ${config.jenisSoal === 'grup' ? 'Soal Grup (Beberapa butir soal didasarkan pada satu stimulus terintegrasi)' : 'Soal Tunggal'}${contextList}${stimulusList}${qualityList}

PANDUAN PENYUSUNAN SOAL:
1. **Pendekatan HOTS**: Fokuskan pertanyaan pada keterampilan berpikir kritis, analisis mendalam, pemecahan masalah, atau evaluasi konsep. Hindari pertanyaan hafalan mentah.
2. **Kekuatan Stimulus & Pengembangan Konten Visual (Mandatori)**: WAJIB menyajikan STIMULUS VISUAL ATAU BERBASIS DATA KONKRET kaya informasi analitis untuk setiap soal. Gunakan salah satu atau kombinasi bentuk visual berikut: (a) TABEL DATA & MATRIKS STATISTIK (Markdown Table), (b) DIAGRAM ALUR / FLOWCHART / BAGAN KONSEP (Code Block ASCII Diagram), (c) GRAFIK TREN TEKSTUAL (Visual Bar Chart), (d) SPESIFIKASI INFOGRAFIS / GAMBAR / PETA (Visual Specification: [DESKRIPSI INFOGRAFIS / SPESIFIKASI VISUAL: Judul]), atau (e) KUTIPAN DOKUMEN & DIALOG BERFORMAT (Blockquote / Dialog). Pertanyaan harus bersandar kuat pada data visual stimulus tersebut.
3. **Kualitas Pilihan Pengecoh**: Pilihan jawaban (A s.d ${config.jumlahOpsi === 5 ? 'E' : 'D'}) harus homogen secara tata bahasa, logis, dan ilmiah. Distraktor tidak boleh terlalu mudah ditebak dan harus menuntut siswa untuk berpikir analitis sebelum memilih.
4. **Kunci & Pembahasan Komprehensif**: Berikan penjelasan analitis langkah-demi-langkah yang ilmiah, objektif, dan logis untuk membuktikan mengapa kunci jawaban tersebut benar dan mengapa opsi lainnya kurang tepat.${(isExactScienceSubject(config.mataPelajaran) || isMatrixOrMathTopic(config.mataPelajaran, config.elemenMateri, config.subElemenMateri, config.kompetensi)) ? `
5. **STANDAR PERSAMAAN, SIMBOL & RUMUS EKSAKTA / SAINS & MATEMATIKA (FISIKA, KIMIA, MATEMATIKA, MATRIKS)**:
   - **Notasi LaTeX Baku**: WAJIB menuliskan SEMUA rumus fisika, reaksi kimia, dan persamaan matematika dengan notasi LaTeX ($ ... $ untuk inline equation dan $$ ... $$ untuk display math) atau format Unicode ilmiah standar. Contoh: $v_t = v_0 + a \\cdot t$, $s = v_0 t + \\frac{1}{2} a t^2$, $F = m \\cdot a$, $E_k = \\frac{1}{2} m v^2$, $E_p = m g h$, $W = F \\cdot s \\cdot \\cos\\theta$, $P = \\frac{W}{t}$, $I = \\frac{V}{R}$, $\\rho = \\frac{m}{V}$, pecahan wajib \\frac{a}{b}, akar \\sqrt{x}, dan pangkat x^2 atau 10^{-3}.
   - **Simbol & Satuan Ilmiah Baku**: Wajib mencantumkan besaran dan satuan fisika/kimia yang tepat (seperti m/s², kg, J, W, Hz, Ω, μC, °C, K, ±, ×, ÷, ≈, ≤, ≥, ≠, ∞, Δ, α, β, θ, λ, μ, π, ρ, σ, ω).
   - **Formula & Reaksi Kimia**: Tuliskan formula kimia dengan subskrip/supersprip yang tepat (contoh: \\text{H}_2\\text{SO}_4, \\text{Fe}^{3+}, \\text{SO}_4^{2-}), tanda panah reaksi (\\rightarrow, \\rightleftharpoons), dan fasa zat ((s), (l), (g), (aq)).
${isMatrixOrMathTopic(config.mataPelajaran, config.elemenMateri, config.subElemenMateri, config.kompetensi) ? `   - **STANDAR KHUSUS SOAL MATRIKS & KOLOM MATEMATIKA (MANDATORI UNTUK AI & LLM)**:
     * **Notasi Kolom & Baris Matriks LaTeX**: WAJIB menuliskan seluruh matriks menggunakan notasi LaTeX matriks standar dengan pemisah kolom tanda ampersand (\`&\`) dan pemisah baris (\`\\\\\`):
       - Matriks Kurung Biasa: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$
       - Matriks Kurung Siku: $\\begin{bmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{bmatrix}$
       - Determinan Matriks: $\\begin{vmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{vmatrix}$ atau $\\det(A)$
       - Matriks Ordo 3x3: $\\begin{pmatrix} a_{11} & a_{12} & a_{13} \\\\ a_{21} & a_{22} & a_{23} \\\\ a_{31} & a_{32} & a_{33} \\end{pmatrix}$
       - Matriks Kolom / Vektor: $\\begin{pmatrix} x \\\\ y \\end{pmatrix}$
     * **Presisi Kolom & Operasi Matriks**: DILARANG KERAS menuliskan elemen matriks rapat tanpa pemisah kolom \`&\` atau tanpa baris \`\\\\\`.
     * **Format Pilihan Jawaban (A s.d E)**: Setiap opsi jawaban yang memuat matriks WAJIB disajikan dalam format LaTeX matriks yang utuh (Contoh: A. $\\begin{pmatrix} 2 & 3 \\\\ -1 & 4 \\end{pmatrix}$).
     * **Keamanan Kolom Tabel Excel / Markdown (Bagian 2)**: Pada tabel Markdown/Excel (Bagian 2), dilarang menggunakan karakter pipa \`|\` tunggal di dalam sel determinan/matriks agar tidak memecah kolom tabel Excel.` : ''}
   - **Format Tabel Data Eksperimen**: WAJIB menyajikan data eksperimen dalam format Markdown Table standar yang utuh 1 baris per record dengan header pemisah (|---|---|) yang presisi.
   - **Homogenitas Opsi Pilihan Jawaban**: Opsi A s.d E wajib homogen dalam format angka, pecahan, desimal, atau penyertaan satuan.` : ''}

==================================================
INSTRUKSI FORMAT OUTPUT UTAMA (SIAP DIPINDAHKAN KE MICROSOFT WORD DAN MICROSOFT EXCEL):
==================================================
FORMAT PENAMAAN FILE DOWNLOAD / DOKUMEN HASIL GENERASI:
Cantumkan Judul / Nama File Dokumentasi Utama pada bagian paling atas dengan format baku:
[Mata Pelajaran]_[Materi Pokok]_[Bentuk Soal]
(Contoh: ${config.mataPelajaran || 'Sosiologi'}_${config.elemenMateri || 'Materi_Pokok'}_${config.bentukSoal === 'mcma' ? 'MCMA' : config.bentukSoal === 'kategori' ? 'Kategori' : 'Sederhana'})

Wajib sajikan hasil pembuatan soal dalam DUA FORMAT OUTPUT LENGKAP berikut agar ketika di-copy dari Gemini, ChatGPT, Claude, dll. hasilnya langsung cocok dan dapat digunakan di MS Word dan MS Excel:

--------------------------------------------------
BAGIAN 1: FORMAT DOKUMEN MS WORD (NASKAH SOAL TERTULIS)
--------------------------------------------------
Sajikan setiap butir soal secara berurutan dengan format naskah dokumen yang rapi, ber-paragraf jelas, dan mudah dipindahkan langsung ke Microsoft Word:

No Soal : [Nomor Soal]
Kompetensi : [Kompetensi yang diuji]
Sub Kompetensi : [Sub kompetensi spesifik]
Bentuk Soal : [Jenis bentuk soal]

--- FORMAT PILIHAN GANDA SEDERHANA & MCMA ---
Soal : [Paragraf stimulus/kasus/data diikuti pertanyaan utama]
Pilihan Jawaban:
A. [Pilihan A]
B. [Pilihan B]
C. [Pilihan C]
D. [Pilihan D]
${config.jumlahOpsi === 5 ? 'E. [Pilihan E]\n' : ''}
Kunci Jawaban: [Kunci Jawaban, misal: B atau A, C]
Pembahasan: [Penjelasan analitis]

--- STRUKTUR KHUSUS PILIHAN GANDA KOMPLEKS KATEGORI (PGK KATEGORI) ---
SOAL:
[Narasi / Stimulus Kasus secara lengkap]
Contoh: Terdapat dua peneliti sosial yang ingin memahami dampak penggunaan media sosial terhadap interaksi sosial remaja di kota-kota besar. Masing-masing menggunakan pendekatan berikut:
• Andi menerapkan pendekatan kuantitatif dengan menyebar kuesioner/angket tertutup kepada 200 responden.
• Yuli menerapkan pendekatan kualitatif dengan melakukan wawancara mendalam pada 10 remaja aktif pengguna media sosial.

[Pertanyaan Utama yang Spesifik]
Contoh: Berdasarkan pernyataan diatas, manakah pengolahan data yang sesuai dengan ilustrasi tersebut?

[Kalimat Perintah Kategori]
Contoh: **Berikan respon/pilihan Anda pada masing-masing pernyataan berikut!** (Contoh label kategori: Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Fakta/Opini, Setuju/Tidak Setuju, dll.)

TABEL PERNYATAAN (Minimal 2 Pernyataan, Maksimal 4 Pernyataan):
| # | Pernyataan | Sesuai | Tidak Sesuai |
|:---:|---|:---:|:---:|
| 1 | Data dari kuesioner/angket yang dikumpulkan Andi dalam penelitian kuantitatif diolah menggunakan analisis statistik. | | |
| 2 | Hasil wawancara mendalam pada penelitian yang Yuli lakukan diolah dengan memberikan kode berupa angka-angka. | | |
| 3 | Andi perlu mengolah data menggunakan perangkat lunak statistik seperti SPSS atau Excel. | | |
| 4 | Dalam penelitian yang Yuli lakukan, makna dan konteks dari jawaban menjadi fokus utama data yang dianalisis. | | |

KUNCI JAWABAN:
Pernyataan 1: Sesuai
Pernyataan 2: Tidak Sesuai
Pernyataan 3: Sesuai
Pernyataan 4: Sesuai

PEMBAHASAN:
1. Data kuesioner kuantitatif diolah dengan analisis statistik.
2. Penelitian kualitatif diolah dengan pengodingan tematik (bukan sekadar angka-angka).
3. Penelitian kuantitatif Andi memerlukan software statistik seperti SPSS atau Excel.
4. Penelitian kualitatif Yuli berfokus pada analisis makna dan konteks jawaban.

--------------------------------------------------
BAGIAN 2: FORMAT TABEL MS EXCEL (MARKDOWN TABLE DATA TERSTRUKTUR)
--------------------------------------------------
Sajikan JUGA seluruh butir soal di atas dalam bentuk TABEL MARKDOWN SINGLE-LINE CELL agar saat pengguna meng-copy tabel ini dan me-paste langsung ke Microsoft Excel atau Google Sheets, data otomatis terbagi secara presisi ke dalam kolom-kolom sel Excel tanpa merusak baris atau tatanan tabel:

| No Soal | Elemen Materi | Sub Elemen | Level Kognitif | Bentuk Soal | Stimulus & Pertanyaan Soal | Opsi_A / Pernyataan_1 | Opsi_B / Pernyataan_2 | Opsi_C / Pernyataan_3 | Opsi_D / Pernyataan_4 | ${config.jumlahOpsi === 5 ? 'Opsi_E | ' : ''}Kunci Jawaban | Pembahasan |
|---|---|---|---|---|---|---|---|---|---|${config.jumlahOpsi === 5 ? '---|' : ''}---|---|
| 1 | ... | ... | ... | ... | ... | Opsi_A / Pernyataan 1 | Opsi_B / Pernyataan 2 | Opsi_C / Pernyataan 3 | Opsi_D / Pernyataan 4 | ${config.jumlahOpsi === 5 ? 'Opsi_E | ' : ''}... | ... |

PERINTAH KHUSUS INTEGRASI WORD & EXCEL:
1. Pastikan hasil soal ketika di-copy di AI (Gemini, ChatGPT, Claude, dll.) secara utuh memuat Format Word (Bagian 1) dan Format Excel (Bagian 2).
2. Pada bagian Tabel Excel (Bagian 2), jangan gunakan karakter baris baru (line-break / enter) di dalam sel tabel Markdown agar 1 nomor soal persis menduduki 1 baris sel di Microsoft Excel.
3. KHUSUS PILIHAN GANDA KOMPLEKS KATEGORI (PGK KATEGORI): Daftar pernyataan (minimal 2, maksimal 4) pada format Excel WAJIB dipisahkan ke masing-masing kolom terpisah yaitu Pernyataan_1, Pernyataan_2, Pernyataan_3, dan Pernyataan_4 (atau Opsi_1 s.d Opsi_4). Kunci Jawaban merinci status per nomor (contoh: Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, Pernyataan 3: Sesuai, Pernyataan 4: Sesuai). DILARANG KERAS menggabungkan seluruh daftar pernyataan ke dalam satu sel Excel!`;

    setGeneratedKisiPrompt(promptKisiText);
    setGeneratedSoalPrompt(promptSoalText);
  }, [config]);

  // Generator Prompt Wadah 2 (CBT Kompleks, Login User/Password, Anti-Contek, Konteks Lokal & Kualitas TKA)
  const buildPromptWadah2 = (form = cbtForm) => {
    const jumlahSoal = Math.min(Math.max(Number(form.jumlahSoal) || 20, 10), 50);
    const konteksText = form.konteksLokal && form.konteksLokal.length > 0
      ? form.konteksLokal.join(', ')
      : 'Budaya Nusantara, Geografis Indonesia, Kehidupan Sosial';
    const stimulusText = form.stimulusKonten && form.stimulusKonten.length > 0
      ? form.stimulusKonten.join(', ')
      : 'Teks Bacaan, Data/Tabel, Kasus Nyata';
    const kualitasText = form.standarKualitas && form.standarKualitas.length > 0
      ? form.standarKualitas.join(', ')
      : 'Validasi Bahasa, Konstruksi Soal, Kesesuaian Materi, Level Kognitif, Konteks Relevan, Tidak Bias, Kejelasan Instruksi, Kunci Jawaban Tepat, Distractor Berkualitas, Sesuai Kurikulum, Waktu Pengerjaan, Inklusivitas';

    return `[PROMPT UTUH APLIKASI QUIZ INTERAKTIF & CBT KOMPLEKS UNTUK CANVAS GEMINI AI]

Anda adalah Pakar Evaluasi Kurikulum Pendidikan SMA, Lead EdTech Software Engineer, dan Specialist Cyber Security Education.
Tugas Anda adalah merancang SATU FILE HTML UTUH TERINTEGRASI (HTML + CSS + JavaScript Interaktif Standalone) untuk Aplikasi Computer Based Test (CBT) Komprehensif yang siap dijalankan langsung di CANVAS Gemini AI.

==================================================
1. SPESIFIKASI DAN PARAMETER UJIAN CBT:
==================================================
- Mata Pelajaran          : ${form.mataPelajaran || config.mataPelajaran || 'Sosiologi'}
- Lingkup Materi          : ${form.lingkupMateri}
- Elemen / Materi Pokok   : ${form.materiPokok}
- Sub-Materi / Indikator  : ${form.subMateri}
- Kompetensi yang Diuji   : ${form.kompetensiYangDiuji || 'Peserta didik mampu menganalisis fenomena sosial secara kritis dan terstruktur.'}
- Batasan / Catatan Khusus: ${form.batasanCatatan || 'Sajikan soal berorientasi HOTS (C4-C6), hindari hafalan tekstual.'}
- Level Kognitif          : ${form.levelKognitif}
- Bentuk Soal             : ${form.bentukSoal}
- Target Jumlah Soal      : ${jumlahSoal} Butir Soal HOTS (C4–C6) Lengkap dengan Stimulus, Opsi/Pernyataan, Kunci & Pembahasan (Minimal 10 - Maksimal 50 Soal)
- Durasi Ujian            : ${form.durasiMenit} Menit

==================================================
2. 🎭 KONTEKS LOKAL INDONESIA & STIMULUS KONTEN:
==================================================
- Integrasi Konteks Lokal Nusantara:
  Wajib mengintegrasikan konteks: ${konteksText}
- Ragam Stimulus & Pengembangan Konten Visual (Mandatori):
  Menggunakan stimulus berupa: ${stimulusText}
  (Pastikan stimulus dikembangkan secara kaya informasi analitis, kontekstual, dan mengutamakan BENTUK VISUAL seperti: Tabel Data/Matriks Markdown, Diagram Alur/Flowchart ASCII, Grafik Tren Tekstual, Spesifikasi Infografis/Peta, atau Kutipan Dokumen/Dialog Berformat).

==================================================
3. 📋 STANDAR KUALITAS SOAL TKA (TES KEMAMPUAN AKADEMIK):
==================================================
Setiap butir soal wajib memenuhi 12 standar kualitas:
${kualitasText}

==================================================
4. SISTEM OTENTIKASI & KEAMANAN ROLE BERTINGKAT (ADMIN, GURU, PESERTA):
==================================================
- Sebelum masuk ke lembar soal, aplikasi WAJIB menampilkan Portal Login CBT Mandiri dengan Dukungan Multi-Role (Admin/Proktor, Guru/Pengampu, Peserta/Siswa).
- Kredensial Akses Ujian & Role Management:
  1. 🔑 ROLE ADMINISTRATOR / PROKTOR UJIAN:
     - Username Admin : ${form.adminUsername || 'admin_cbt'}
     - Password Admin : ${form.adminPassword || 'admin_proktor2026'}
     - Hak Akses Khusus: Panel Proktor untuk Reset Sesi/Token Siswa, Monitor Log Pelanggaran Tab-Switching Realtime, Pengaturan Durasi/Kunci Ujian, serta Ekspor Rekap Nilai Ujian (PDF/Excel).
  2. 👨‍🏫 ROLE GURU PENGAMPU / PEMBUAT SOAL:
     - Username Guru  : ${form.guruUsername || 'guru_sosiologi'}
     - Password Guru  : ${form.guruPassword || 'guru_pass2026'}
     - Hak Akses Khusus: Dashboard Guru untuk Preview & Validasi Butir Soal, Edit Pembahasan & Kunci Jawaban, serta Analisis Statistik Tingkat Kesukaran Soal.
  3. 🎓 ROLE PESERTA UJIAN / SISWA:
     - Username Peserta : ${form.usernameCbt || 'peserta_tka'}
     - Password Peserta : ${form.passwordCbt || 'cbt2026_sosiologi'}
     - Hak Akses Khusus: Sesi Pengerjaan Ujian CBT Mandiri Terlindungi Proteksi Anti-Contek, Anti Tab-Switching, dan Timer Auto-Submit.

- Proteksi Keamanan Kredensial & Validasi Otentikasi:
  * Validasi Ketat: Jika Username atau Password salah/tidak cocok dengan role yang dipilih, tampilkan notifikasi "Kredensial Login Tidak Valid atau Akses Ditolak!" dengan animasi shake.
  * Proteksi Anti Brute-Force: Maksimal 5 kali percobaan gagal berturut-turut, sistem akan mengunci sementara form login selama 60 detik.
  * Keamanan Sandi & Enkripsi Sesi: Masking input password dengan tombol Toggle Show/Hide, enkripsi token sesi lokal (Session Storage), serta perlindungan dari bypass console F12.
  * Kartu Profil Setelah Login: Tampilkan Kartu Profil (Nama User, Role Akses, Nomor Peserta/NIP, Mata Pelajaran, Jumlah Soal, Durasi) dan Tombol "Mulai Sesi Ujian".

==================================================
5. SISTEM KEAMANAN KOMPLEKS (ANTI-CONTEK & ANTI-CURANG):
==================================================
- Pengacakan Dinamis (Randomization Engine): Urutan ${jumlahSoal} butir soal dan pilihan jawaban/opsi diacak secara otomatis (Fisher-Yates Shuffle) untuk tiap sesi pengerjaan siswa.
- Focus Guard & Anti Tab-Switching: Pantau event 'visibilitychange' & 'blur'. Jika siswa berpindah tab/layar, tampilkan Peringatan Pelanggaran. Batas melanggar 3 kali; jika dilanggar 3x, sistem akan OTOMATIS MENGUMPULKAN (AUTO-SUBMIT) lembar jawaban secara paksa.
- Audit Trail Log Pelanggaran: Catat tanggal & waktu saat siswa terdeteksi berpindah tab (misal: "Pelanggaran #1: Pindah Tab/Layar jam 10:15:22").
- Proteksi Pemblokiran Konten (Anti Copy-Paste & Anti Inspect): Blokir Klik Kanan (contextmenu), fungsi Copy, Cut, Paste, dan Seleksi Teks. Blokir shortcut keyboard: F12, Ctrl+Shift+I, Ctrl+Shift+J, Ctrl+U, Ctrl+C, Ctrl+V, Alt+Tab, PrtScr.
- Fullscreen Lock Requirement: Wajibkan mode Layar Penuh (Fullscreen) sebelum soal terbuka. Ujian terkunci jika keluar dari mode Fullscreen.
- Timer Countdown Realtime & Auto-Submit: Hitung mundur durasi ${form.durasiMenit} menit. Ketika waktu habis (00:00), jawaban siswa otomatis terkirim.

==================================================
6. FITUR NAVIGASI, AKSESIBILITAS & DUKUNGAN PENGERJAAN:
==================================================
- Header Dashboard: Judul Mata Pelajaran, Sisa Waktu (Countdown Clock), Profil Siswa, dan Pengatur Ukuran Font Soal (A- / A / A+).
- Fitur "Ragu-Ragu" (Mark for Review): Siswa dapat menandai soal ragu-ragu sehingga indikator nomor soal di grid berubah warna menjadi Kuning.
- Auto-Save Progress (localStorage): Setiap jawaban dan penandaan ragu-ragu tersimpan otomatis di localStorage agar data tidak hilang jika halaman ter-refresh secara tidak sengaja.
- Grid Navigasi Nomor Soal (1–${jumlahSoal}): Indikator visual real-time:
  * Hijau : Soal Sudah Terjawab
  * Kuning: Soal Terjawab / Belum dengan Penanda Ragu-Ragu
  * Abu-abu: Soal Belum Terjawab

==================================================
7. EVALUASI AKHIR & LAPORAN HASIL UJIAN:
==================================================
- Modal Konfirmasi Selesai Ujian: Tampilkan ringkasan (Terjawab, Ragu-Ragu, Belum Terjawab) sebelum kirim final.
- Kartu Skor Akhir: Skor (0–100), Jumlah Benar, Salah, Kosong, Waktu Pengerjaan, dan Status Kelulusan KKM.
- Rekap Audit Pelanggaran: Menampilkan catatan total pelanggaran pindah tab jika ada.
- Fitur Cetak / Unduh Laporan: Tombol "Cetak Laporan / Unduh Hasil Ujian" untuk diarsipkan oleh Guru/Pengawas.
- Pembahasan Ilmiah HOTS Langkah Demi Langkah: Pembahasan analitis lengkap beserta kunci jawaban untuk seluruh ${jumlahSoal} butir soal yang dapat dibuka setelah ujian selesai.

==================================================
PETUNJUK EKSEKUSI PADA CANVAS GEMINI AI:
==================================================
Tolong buka fitur CANVAS di Gemini AI dan buatkan SATU KESATUAN KODE HTML, CSS, dan JAVASCRIPT UTUH TERINTEGRASI untuk Aplikasi CBT Interaktif sesuai spesifikasi di atas. Tampilan UI harus bersih, modern, kontras tinggi, dan nyaman untuk siswa SMA.

© 2026 @AJISOSIOLOGI - Assessment TKA SMA & CBT System`;
  };

  // Auto-sync Pembuat Prompt Otomatis AI (Megaprompt) ke Wadah 1 & Wadah 2 untuk CANVAS Gemini AI
  useEffect(() => {
    // Wadah 1: Synced from Matriks Asesmen (Kisi-Kisi)
    if (generatedSoalPrompt || generatedKisiPrompt) {
      const w1Formatted = `[INSTRUCTION FOR CANVAS GEMINI AI - HASIL PROMPT MATRIKS ASESMEN & KISI-KISI]

Anda adalah Pakar Evaluasi Kurikulum Pendidikan SMA & Developer Aplikasi EdTech Interaktif.
Gunakan Draf Megaprompt Terstruktur dari Matriks Asesmen Pintar TKA di bawah ini untuk menghasilkan Modul Asesmen / Quiz Interaktif pada CANVAS Gemini AI.

==================================================
DRAF MEGAPROMPT MATRIKS ASESMEN / KISI-KISI (${config.mataPelajaran || 'SMA'} - HOTS C4-C6):
==================================================
${generatedSoalPrompt || generatedKisiPrompt}

==================================================
PETUNJUK EKSEKUSI PADA CANVAS GEMINI AI:
==================================================
1. Tolong buka tampilan CANVAS di Gemini AI dan buatkan Modul Asesmen / Quiz Interaktif berbasis HTML, CSS, dan JS.
2. Tampilkan Soal HOTS dengan Stimulus Kontekstual, Pilihan Jawaban A–E, Skor Otomatis, dan Pembahasan Ilmiah.
3. Tampilan UI bersih, modern, kontras tinggi, dan nyaman untuk siswa SMA.

© 2026 @AJISOSIOLOGI - Assessment TKA SMA System`;
      setPromptWadah1Text(w1Formatted);
    }

    // Wadah 2: Auto-generate CBT Kompleks Prompt
    const w2Text = buildPromptWadah2(cbtForm);
    setPromptWadah2Text(w2Text);
  }, [generatedSoalPrompt, generatedKisiPrompt, config.mataPelajaran, cbtForm]);

  // Copy helper
  const handleCopy = (text: string, type: 'kisi' | 'soal') => {
    navigator.clipboard.writeText(text);
    if (type === 'kisi') {
      setCopiedKisi(true);
      setTimeout(() => setCopiedKisi(false), 2000);
    } else {
      setCopiedSoal(true);
      setTimeout(() => setCopiedSoal(false), 2000);
    }
  };

  // Toggle Context Checkboxes
  const handleToggleContext = (item: string) => {
    setConfig(prev => {
      const exists = prev.konteksLokal.includes(item);
      const updated = exists 
        ? prev.konteksLokal.filter(x => x !== item) 
        : [...prev.konteksLokal, item];
      return { ...prev, konteksLokal: updated };
    });
  };

  // Toggle Stimulus Checkboxes
  const handleToggleStimulus = (item: string) => {
    setConfig(prev => {
      const exists = prev.stimulusKonten.includes(item);
      const updated = exists 
        ? prev.stimulusKonten.filter(x => x !== item) 
        : [...prev.stimulusKonten, item];
      return { ...prev, stimulusKonten: updated };
    });
  };

  // Toggle Quality Checkboxes
  const handleToggleQuality = (item: string) => {
    setConfig(prev => {
      const exists = prev.kualitasChecklist.includes(item);
      const updated = exists 
        ? prev.kualitasChecklist.filter(x => x !== item) 
        : [...prev.kualitasChecklist, item];
      return { ...prev, kualitasChecklist: updated };
    });
  };

  // Toggle CBT Wadah 2 Context Checkboxes
  const handleToggleCbtContext = (item: string) => {
    setCbtForm(prev => {
      const current = prev.konteksLokal || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, konteksLokal: updated };
    });
  };

  // Toggle CBT Wadah 2 Stimulus Checkboxes
  const handleToggleCbtStimulus = (item: string) => {
    setCbtForm(prev => {
      const current = prev.stimulusKonten || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, stimulusKonten: updated };
    });
  };

  // Toggle CBT Wadah 2 Quality Checkboxes
  const handleToggleCbtQuality = (item: string) => {
    setCbtForm(prev => {
      const current = prev.standarKualitas || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, standarKualitas: updated };
    });
  };

  // Toggle Kisi Form Context Checkboxes
  const handleToggleKisiContext = (item: string) => {
    setKisiForm(prev => {
      const current = prev.konteksLokal || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, konteksLokal: updated };
    });
  };

  // Toggle Kisi Form Stimulus Checkboxes
  const handleToggleKisiStimulus = (item: string) => {
    setKisiForm(prev => {
      const current = prev.stimulusKonten || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, stimulusKonten: updated };
    });
  };

  // Toggle Kisi Form Quality Checkboxes
  const handleToggleKisiQuality = (item: string) => {
    setKisiForm(prev => {
      const current = prev.kualitasChecklist || [];
      const updated = current.includes(item)
        ? current.filter(x => x !== item)
        : [...current, item];
      return { ...prev, kualitasChecklist: updated };
    });
  };

  // Trigger server-side or client-side AI generation of Kisi-Kisi
  const handleGenerateKisiViaAI = async () => {
    if (!config.mataPelajaran) {
      alert('Sila pilih Mata Pelajaran terlebih dahulu di Tab 1!');
      return;
    }
    setIsGeneratingKisi(true);
    try {
      let data;
      if (aiConfig.mode === 'client') {
        const systemInstruction = `Anda adalah ahli kurikulum pendidikan Indonesia. Buatkan matriks kisi-kisi ujian TKA SMA tingkat tinggi (HOTS) berdasarkan masukan parameter mata pelajaran.`;
        const prompt = `Buatkan 3 baris matriks asesmen kisi-kisi baru yang bervariasi secara otomatis untuk Mata Pelajaran ${config.mataPelajaran} dengan parameter:
Definisi/Tujuan: ${config.definisi || ""}
Muatan Kurikulum: ${config.muatan || ""}
Kompetensi Umum: ${config.kompetensi || ""}
Elemen/Materi: ${config.elemenMateri || ""}
Sub-Elemen: ${config.subElemenMateri || ""}

Aturan Penyusunan Matriks:
1. Setiap baris harus bervariasi jenis bentuk soalnya: 'pilihan_ganda_sederhana' (PG Sederhana), 'mcma' (PG Kompleks Multiple Choice Multiple Answers), atau 'kategori' (PG Kompleks kategori Benar/Salah atau Sesuai/Tidak Sesuai).
2. Tingkat kognitif harus bervariasi antara: 'level_1' (Pemahaman), 'level_2' (Penerapan), atau 'level_3' (Penalaran).
3. Buat rincian elemen, sub-elemen, kompetensi yang diukur, serta batasan materi secara logis dan mendalam.
4. Distribusikan jumlah soal per kisi-kisi (misalnya antara 3-10 soal per baris).
5. Hasilkan juga 'konteksNusantara' serta 'stimulusTambahan' untuk meningkatkan kualitas stimulus soal.`;

        const kisiSchema = {
          type: "ARRAY",
          items: {
            type: "OBJECT",
            properties: {
              bentukSoal: { 
                type: "STRING", 
                description: "Nilai wajib berupa salah satu dari: 'pilihan_ganda_sederhana', 'mcma', atau 'kategori'" 
              },
              levelKognitif: { 
                type: "STRING", 
                description: "Nilai wajib berupa salah satu dari: 'level_1', 'level_2', atau 'level_3'" 
              },
              elemenMateri: { type: "STRING" },
              subElemenMateri: { type: "STRING" },
              kompetensi: { type: "STRING" },
              batasanCatatan: { type: "STRING" },
              jumlahSoal: { type: "INTEGER" },
              konteksNusantara: { type: "STRING" },
              stimulusTambahan: { type: "STRING" }
            },
            required: [
              "bentukSoal", "levelKognitif", "elemenMateri", "subElemenMateri",
              "kompetensi", "batasanCatatan", "jumlahSoal", "konteksNusantara", "stimulusTambahan"
            ]
          }
        };

        const responseText = await callGeminiDirect(systemInstruction, prompt, kisiSchema);
        data = JSON.parse(responseText);
      } else {
        const response = await fetch('/api/generate-kisi', {
          method: 'POST',
          headers: { 
            'Content-Type': 'application/json',
            'x-api-key': getCleanApiKey(aiConfig.apiKey)
          },
          body: JSON.stringify({
            mataPelajaran: config.mataPelajaran,
            definisi: config.definisi,
            muatan: config.muatan,
            kompetensi: config.kompetensi,
            elemenMateri: config.elemenMateri,
            subElemenMateri: config.subElemenMateri,
            count: 3
          })
        });

        if (!response.ok) {
          let errorMsg = 'Terjadi kesalahan pada server AI.';
          try {
            const textError = await response.text();
            if (textError.includes('<!doctype') || textError.includes('<html')) {
              errorMsg = 'Server sedang sibuk atau mengalami timeout (Gateway Timeout). Silakan coba lagi beberapa saat.';
            } else {
              try {
                const errorData = JSON.parse(textError);
                errorMsg = errorData.error || errorMsg;
              } catch {
                errorMsg = textError || errorMsg;
              }
            }
          } catch {}
          throw new Error(errorMsg);
        }

        const responseText = await response.text();
        data = JSON.parse(responseText);
      }

      if (Array.isArray(data)) {
        // Map to KisiKisiItem schema
        const mapped: KisiKisiItem[] = data.map((item: any, idx: number) => ({
          id: `kisi-ai-${Date.now()}-${idx}`,
          userId: currentUser?.uid,
          no: kisiList.length + idx + 1,
          bentukSoal: ['pilihan_ganda_sederhana', 'mcma', 'kategori'].includes(item.bentukSoal) 
            ? item.bentukSoal 
            : 'pilihan_ganda_sederhana',
          levelKognitif: ['level_1', 'level_2', 'level_3'].includes(item.levelKognitif) 
            ? item.levelKognitif 
            : 'level_2',
          elemenMateri: item.elemenMateri || config.elemenMateri,
          subElemenMateri: item.subElemenMateri || config.subElemenMateri,
          kompetensi: item.kompetensi || 'Menyelesaikan permasalahan',
          batasanCatatan: item.batasanCatatan || '',
          jumlahSoal: Number(item.jumlahSoal) || 5,
          konteksNusantara: item.konteksNusantara || '',
          stimulusTambahan: item.stimulusTambahan || '',
          konteksLokal: item.konteksLokal || [],
          stimulusKonten: item.stimulusKonten || [],
          kualitasChecklist: item.kualitasChecklist || []
        }));

        const batch = writeBatch(db);
        mapped.forEach((kItem) => {
          batch.set(doc(db, 'kisi_kisi', kItem.id), kItem);
        });
        await batch.commit();

        setActiveTab('kisi');
        alert(`Berhasil membuat ${mapped.length} baris Matriks Asesmen Kisi-Kisi secara otomatis via AI!`);
      } else {
        throw new Error('Format respon AI tidak valid');
      }
    } catch (err: any) {
      console.error(err);
      alert(`Gagal membuat otomatis: ${err.message}. Tenang, Anda masih bisa menambahkan baris kisi-kisi secara manual dengan sangat mudah!`);
    } finally {
      setIsGeneratingKisi(false);
    }
  };

  // Open Prompt Generator Modal for a specific Kisi-Kisi row
  const handleOpenPromptGenerator = (item: KisiKisiItem) => {
    setSelectedKisiForPrompt(item);
    setCopiedPrompt(false);
    
    const isMatrix = isMatrixOrMathTopic(config.mataPelajaran, item.elemenMateri, item.subElemenMateri, item.kompetensi);
    const isExact = isExactScienceSubject(config.mataPelajaran) || isMatrix;

    // Generate a beautiful, highly useful default prompt locally first (instant)
    const localPrompt = `⚡ PERSYARATAN MUTLAK KELUARAN AI (2 FORMAT FILE UTUH SIAP UNDUH / DIPAKAI):
================================================================================
Anda WAJIB dan MUTLAK menyajikan SELURUH HASIL GENERASI SOAL dalam DUA FORMAT OUTPUT UTUH LENGKAP BERURUTAN DALAM SATU BALASAN:
1️⃣ [BAGIAN 1: NASKAH SOAL LENGKAP SIAP CETAK - FORMAT WORD / DOCX READY]
2️⃣ [BAGIAN 2: TABEL MATRIKS REKAPITULASI SOAL - FORMAT EXCEL / SPREADSHEET READY]

DILARANG KERAS HANYA MEMBUAT NASKAH WORD ATAU HANYA TABEL EXCEL! KEDUANYA WAJIB DITULISKAN SELURUHNYA SEHINGGA PENGGUNA DAPAT LANGSUNG MEMAKAI DAN MENGAMBIL 2 FORMAT FILE SEKALIGUS (WORD & EXCEL)!
================================================================================

Anda adalah seorang ahli penyusun soal TKA (Tes Kemampuan Akademik) SMA tingkat nasional dan pakar evaluasi kurikulum pendidikan di Indonesia${isExact ? " (Spesialis Sains & Matematika Eksakta)" : ""}.
Tugas Anda adalah merancang ${item.jumlahSoal || 5} butir soal ${getBentukSoalLabel(item.bentukSoal)} berorientasi HOTS (Higher Order Thinking Skills) untuk mata pelajaran ${config.mataPelajaran || "Umum"} tingkat SMA, Kelas XII.

SPESIFIKASI BUTIR SOAL:
- Mata Pelajaran: ${config.mataPelajaran || "Umum"}
- Lingkup Materi / Kompetensi: ${item.kompetensi}
- Materi Pokok (Elemen): ${item.elemenMateri}
- Sub-materi (Sub-elemen) / Indikator Soal: ${item.subElemenMateri || '-'}
- Level Kognitif: ${getLevelKognitifLabel(item.levelKognitif)} (${item.levelKognitif})
- Bentuk Soal: ${getBentukSoalLabel(item.bentukSoal)}
${item.konteksNusantara ? `- Konteks Nusantara: ${item.konteksNusantara}` : ''}
${item.stimulusTambahan ? `- Stimulus Tambahan: ${item.stimulusTambahan}` : ''}
${item.batasanCatatan ? `- Catatan Khusus: ${item.batasanCatatan}` : ''}

PANDUAN PENYUSUNAN SOAL:
1. **Analisis HOTS (C4-C6)**: Pertanyaan harus mengukur kemampuan menganalisis, mengevaluasi, atau merancang/berpikir kritis siswa, bukan hafalan tekstual.
2. **PENGGABUNGAN STIMULUS VISUAL DAN PERTANYAAN (MANDATORI)**: Wajib menyajikan STIMULUS VISUAL ATAU BERBASIS DATA KONKRET (seperti Tabel Data/Matriks Markdown, Diagram Alur/Flowchart ASCII, Grafik Tren Tekstual, Spesifikasi Infografis/Peta, atau Kutipan Dokumen/Dialog Berformat) dan pertanyaan utama secara langsung menyatu di dalam satu bagian 'Soal:'. JANGAN memisahkan field stimulus dan soal, dan JANGAN mencantumkan nomor soal (misal '1.', 'Soal 1.') di dalam teks soal.
3. **Pengecoh Homogen & Ilmiah**: Seluruh pilihan jawaban (opsi/pernyataan) harus homogen secara sintaksis, setara panjangnya, logis, dan menantang siswa untuk mengeliminasi distraktor secara analitis. Setiap pilihan jawaban WAJIB diawali dengan huruf label A, B, C, D, E dan titik, contoh: A. ..., B. ..., C. ..., D. ..., E. ...
4. **PEMISAHAN OPSI EXCEL (KHUSUS MCMA & KATEGORI)**: Untuk bentuk Pilihan Ganda Kompleks (MCMA) dan Pilihan Ganda Kompleks Kategori, pilihan jawaban/pernyataan pada format Excel yang sebelumnya tergabung dalam satu sel WAJIB dipisahkan menjadi masing-masing kolom terpisah yaitu Opsi_A, Opsi_B, Opsi_C, Opsi_D, dan Opsi_E. DILARANG menggabungkan seluruh pilihan/pernyataan ke dalam 1 sel Excel.
5. **Pembahasan Ilmiah**: Sertakan pembahasan langkah demi langkah yang logis, mendalam, tanpa tanda asteris (*), serta membuktikan kebenaran kunci jawaban.${isMatrix ? `
6. **STANDAR KHUSUS SOAL MATRIKS & KOLOM MATEMATIKA (MANDATORI)**:
   - WAJIB menuliskan seluruh matriks menggunakan notasi LaTeX matriks standar:
     * Matriks kurung biasa: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$
     * Matriks kurung siku: $\\begin{bmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{bmatrix}$
     * Determinan: $\\begin{vmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{vmatrix}$
   - Setiap baris dipisahkan '\\\\' dan setiap kolom dipisahkan '&'. DILARANG menuliskan elemen matriks rapat tanpa pemisah kolom '&' dan baris '\\\\'.
   - Seluruh opsi pilihan jawaban (A s.d E) yang memuat matriks WAJIB disajikan dalam format LaTeX matriks yang utuh.` : ''}
7. **FORMAT PENAMAAN FILE DOWNLOAD DOKUMEN**:
   Wajib mencantumkan Judul / Nama File dokumen pada bagian awal dengan format baku: [Mata Pelajaran]_[Materi Pokok]_[Bentuk Soal]
   (Contoh: ${config.mataPelajaran || "Sosiologi"}_${item.elemenMateri || "Materi_Pokok"}_${item.bentukSoal === 'mcma' ? 'MCMA' : item.bentukSoal === 'kategori' ? 'Kategori' : 'Sederhana'}).

Sajikan output Anda dengan DUA FORMAT KELUARAN LENGKAP UTUH BERURUTAN (WAJIB MENAMPILKAN KEDUANYA SEKALIGUS):

[BAGIAN 1: FORMAT NASKAH WORD]
===========================================
No Soal : [Nomor Soal]
Kompetensi : [Kompetensi yang diuji]
Sub Kompetensi : [Sub kompetensi spesifik]
Bentuk Soal : [Jenis bentuk soal]

${item.bentukSoal === 'kategori' ? `SOAL:
[Narasi / Stimulus Kasus secara lengkap]
[Pertanyaan Utama yang Spesifik]
[Kalimat Perintah Kategori]
Contoh: **Berikan respon/pilihan Anda pada masing-masing pernyataan berikut!** (Contoh label kategori: Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Fakta/Opini, Setuju/Tidak Setuju, dll.)

TABEL PERNYATAAN (Minimal 2 Pernyataan, Maksimal 4 Pernyataan):
| # | Pernyataan | [Kategori A] | [Kategori B] |
|:---:|---|:---:|:---:|
| 1 | [Teks Pernyataan 1] | | |
| 2 | [Teks Pernyataan 2] | | |
| 3 | [Teks Pernyataan 3] | | |
| 4 | [Teks Pernyataan 4] | | |

KUNCI JAWABAN:
Pernyataan 1: [Kategori A / Status 1]
Pernyataan 2: [Kategori B / Status 2]
Pernyataan 3: [Kategori A / Status 3]
Pernyataan 4: [Kategori A / Status 4]

PEMBAHASAN:
1. [Penjelasan analitis ilmiah mengapa Pernyataan 1 berstatus Sesuai]
2. [Penjelasan analitis ilmiah mengapa Pernyataan 2 berstatus Tidak Sesuai]
3. [Penjelasan analitis ilmiah mengapa Pernyataan 3 berstatus Sesuai]
4. [Penjelasan analitis ilmiah mengapa Pernyataan 4 berstatus Sesuai]` : `Soal (Menggabungkan Stimulus dan Pertanyaan Utama):
[Tuliskan paragraf stimulus, data/tabel, atau situasi kontekstual, lalu diikuti langsung dengan pertanyaan utama atau instruksi pengerjaan secara menyatu dalam satu kesatuan teks soal tanpa dipisah]

Pilihan Jawaban / Pernyataan:
A. [Pilihan / Pernyataan A]
B. [Pilihan / Pernyataan B]
C. [Pilihan / Pernyataan C]
D. [Pilihan / Pernyataan D]
E. [Pilihan / Pernyataan E]

Kunci Jawaban: [Kunci Jawaban yang tepat]
Pembahasan: [Penjelasan analitis langkah demi langkah secara ilmiah dan terstruktur]`}
===========================================

[BAGIAN 2: FORMAT TABEL EXCEL]
| No Soal | Kompetensi | Bentuk Soal | Soal (Stimulus + Pertanyaan) | Opsi_A / Pernyataan_1 | Opsi_B / Pernyataan_2 | Opsi_C / Pernyataan_3 | Opsi_D / Pernyataan_4 | Opsi_E | Kunci Jawaban | Pembahasan |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | ... | ... | ... | [Pernyataan/Opsi A] | [Pernyataan/Opsi B] | [Pernyataan/Opsi C] | [Pernyataan/Opsi D] | [Pernyataan/Opsi E] | ... | ... |

PERINGATAN MANDATORI:
Wajib dan mutlak menuliskan DUA BAGIAN di atas: [BAGIAN 1: FORMAT NASKAH WORD] dan [BAGIAN 2: FORMAT TABEL EXCEL]. DILARANG memotong output sebelum Bagian 2 selesai dituliskan secara lengkap!`;

    setGeneratedPromptText(localPrompt);
    setIsPromptModalOpen(true);
  };

  // Optimize prompt using server-side or client-side Gemini AI
  const handleOptimizePromptWithAi = async () => {
    if (!selectedKisiForPrompt) return;
    setIsGeneratingPrompt(true);
    try {
      if (aiConfig.mode === 'client') {
        const systemInstruction = `Anda adalah pakar prompt engineering pendidikan.`;
        const isMatrix = isMatrixOrMathTopic(config.mataPelajaran, selectedKisiForPrompt.elemenMateri, selectedKisiForPrompt.subElemenMateri, selectedKisiForPrompt.kompetensi);
        const isExact = isExactScienceSubject(config.mataPelajaran) || isMatrix;
        const exactPromptRules = isExact ? `
9. STANDAR PERSAMAAN, RUMUS & TABEL EKSAKTA (FISIKA, KIMIA, MATEMATIKA):
   - Gunakan notasi LaTeX ($ ... $ inline, $$ ... $$ display) atau simbol Unicode ilmiah baku untuk semua rumus dan persamaan (misal: $v_t = v_0 + at$, $s = v_0 t + \\frac{1}{2}at^2$, $F = m \\cdot a$, $E_k = \\frac{1}{2}mv^2$, pecahan \\frac{a}{b}, akar \\sqrt{x}, pangkat x^2 atau 10^{-3}).
   - Cantumkan satuan baku (m/s², kg, J, W, Hz, Ω, μC, °C, K, ±, ×, ÷, ≈, ≤, ≥, ≠, ∞, Δ, α, β, θ, λ, μ, π, ρ, σ, ω).
   - Reaksi kimia dengan subskrip/supersprip yang tepat (\\text{H}_2\\text{SO}_4, \\text{Fe}^{3+}).
${isMatrix ? `   - STANDAR KHUSUS SOAL MATRIKS: WAJIB menyajikan matriks dengan notasi LaTeX matriks baku (contoh: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$, $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$, $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$). Elemen baris dipisahkan '\\\\' dan elemen kolom dipisahkan '&'. DILARANG menuliskan elemen matriks rapat tanpa pemisah kolom dan baris. Opsi pilihan jawaban A s.d E yang memuat matriks WAJIB disajikan lengkap dalam notasi LaTeX matriks.` : ''}
   - Tabel data disajikan dalam format Markdown Table standar yang utuh 1 baris per record.
   - Opsi jawaban A-E wajib homogen dalam format angka, pecahan, desimal, atau satuan.` : '';

        const prompt = `Anda adalah seorang ahli instruktur prompt (prompt engineer) yang berpengalaman membuat instruktur sistem tingkat lanjut (system instructions) dan prompt untuk AI LLM generasi terbaru.
Tugas Anda adalah memformulasikan prompt instruksi spesifik dan sangat mendalam (Super-Prompt) untuk menghasilkan butir-butir soal berkualitas tinggi (HOTS) berdasarkan kisi-kisi berikut.

ATURAN MANDATORI FORMAT OUTPUT:
1. Wajib menginstruksikan AI untuk MENGGABUNGKAN stimulus (paragraf/data/tabel/studi kasus) dan pertanyaan utama secara langsung menyatu di dalam satu bagian 'Soal:', tanpa dipisah menjadi section/field stimulus tersendiri.
2. WAJIB menginstruksikan penggunaan STIMULUS VISUAL ATAU BERBASIS DATA KONKRET (seperti Tabel Data/Matriks Markdown, Diagram Alur/Flowchart ASCII, Grafik Tren Tekstual, Spesifikasi Infografis/Peta, atau Kutipan Dokumen/Dialog Berformat).
3. JANGAN mencantumkan nomor soal (seperti '1.', 'Soal 1.') di dalam teks soal.
4. Setiap pilihan jawaban WAJIB diawali huruf A, B, C, D, E dan titik (misal 'A. ...').
5. Wajib menginstruksikan agar hasil soal disajikan dalam DUA format: Format Word dan Format Excel.
6. KHUSUS PILIHAN GANDA KOMPLEKS KATEGORI (PGK KATEGORI): Wajib menginstruksikan AI eksternal untuk meminta peserta memberikan respon pada masing-masing pernyataan (opsi kategori fleksibel: Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Ya/Tidak, Fakta/Opini, Setuju/Tidak Setuju, dll.). Menyajikan naskah soal (Bagian 1) dengan struktur khusus: (a) SOAL memuat [Stimulus/Narasi Kasus], [Pertanyaan Utama Spesifik], dan [Kalimat Perintah Kategori], (b) TABEL PERNYATAAN berbentuk Tabel Markdown 4 Kolom (| # | Pernyataan | [Kategori A] | [Kategori B] |) dengan nomor urut angka 1, 2, 3, 4 di kolom '#' (minimal 2, maksimal 4 pernyataan), (c) KUNCI JAWABAN merinci status respon per nomor, (d) PEMBAHASAN merinci penjelasan ilmiah untuk poin nomor 1 s.d 4 satu per satu, dan (e) TABEL EXCEL (Bagian 2) diisi dengan kolom Pernyataan_1, Pernyataan_2, Pernyataan_3, Pernyataan_4 memuat teks pernyataan 1 s.d 4 secara terpisah. DILARANG KERAS menggabungkan seluruh daftar pernyataan ke dalam 1 sel.
7. Wajib menginstruksikan AI untuk menyajikan nama/judul file dokumen download dengan format baku: [Mata Pelajaran]_[Materi Pokok]_[Bentuk Soal] (contoh: Sosiologi_Perubahan_Sosial_MCMA, Sosiologi_Perubahan_Sosial_Kategori, atau Sosiologi_Perubahan_Sosial_Sederhana).${exactPromptRules}

Masukan Parameter Kisi-Kisi:
- Mata Pelajaran: ${config.mataPelajaran}
- Bentuk Soal: ${selectedKisiForPrompt.bentukSoal}
- Tingkat Kognitif: ${selectedKisiForPrompt.levelKognitif}
- Elemen/Materi: ${selectedKisiForPrompt.elemenMateri}
- Sub-Elemen/Submateri: ${selectedKisiForPrompt.subElemenMateri}
- Kompetensi yang Diuji: ${selectedKisiForPrompt.kompetensi}
- Batasan/Catatan Khusus: ${selectedKisiForPrompt.batasanCatatan || 'Tidak ada'}
- Konteks Nusantara: ${selectedKisiForPrompt.konteksNusantara || 'Tidak ada'}
- Stimulus Tambahan: ${selectedKisiForPrompt.stimulusTambahan || 'Tidak ada'}

Hasilkan rancangan prompt instruksi lengkap, terstruktur, profesional, dan dalam bahasa Indonesia formal, tanpa mencantumkan kode JSON atau Markdown codeblock, melainkan langsung teks prompt siap pakai yang bisa disalin oleh guru.`;

        const optimizedPrompt = await callGeminiDirect(systemInstruction, prompt);
        setGeneratedPromptText(optimizedPrompt);
      } else {
        const response = await fetch('/api/optimize-prompt', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': getCleanApiKey(aiConfig.apiKey)
          },
          body: JSON.stringify({
            kisi: selectedKisiForPrompt,
            mataPelajaran: config.mataPelajaran,
          }),
        });

        if (!response.ok) {
          throw new Error('Gagal menghubungi server AI untuk optimasi.');
        }

        const data = await response.json();
        if (data.prompt) {
          setGeneratedPromptText(data.prompt);
        } else {
          throw new Error('Format respon tidak sesuai.');
        }
      }
    } catch (err: any) {
      alert(`Gagal optimasi prompt: ${err.message}. Menggunakan draf prompt lokal default.`);
    } finally {
      setIsGeneratingPrompt(false);
    }
  };

  // Trigger AI generation of questions for a specific Kisi-Kisi row (Server or Direct Browser)
  const handleGenerateQuestionsForKisi = async (kisi: KisiKisiItem) => {
    if (!config.mataPelajaran) {
      alert('Sila pilih Mata Pelajaran terlebih dahulu di Tab 1!');
      return;
    }

    setIsGeneratingSoal(true);
    const totalToGenerate = kisi.jumlahSoal || 1;
    const existingMaxNo = questions.reduce((max, q) => Math.max(max, Number(q.noSoal) || 0), 0);
    let currentNoSoal = existingMaxNo + 1;
    const allExistingSoalTexts = questions.map(q => q.soal);

    setSoalProgress({
      active: true,
      type: 'single',
      currentNo: kisi.no || 1,
      totalNo: 1,
      topic: kisi.subElemenMateri || 'Materi Single',
      countSuccess: 0,
      totalQuestions: totalToGenerate,
      statusText: `Menghubungkan ke AI (${aiConfig.mode === 'client' ? 'Direct Browser' : 'Server AI'}) untuk kisi-kisi No. ${kisi.no}...`
    });

    try {
      let data: any[] = [];
      if (aiConfig.mode === 'client') {
        const systemInstruction = `Anda adalah ahli pembuat soal ujian nasional dan TKA (Tes Kemampuan Akademik) SMA di Indonesia. Anda sangat terampil menyusun soal tingkat tinggi (HOTS), bervariasi, mendalam, dan bebas dari bias. Patuhi instruksi bentuk soal dan parameter kognitif secara presisi.`;
        const activeKonteksLokal = (kisi.konteksLokal && kisi.konteksLokal.length > 0) ? kisi.konteksLokal : config.konteksLokal;
        const activeStimulusKonten = (kisi.stimulusKonten && kisi.stimulusKonten.length > 0) ? kisi.stimulusKonten : config.stimulusKonten;
        const activeKualitasChecklist = (kisi.kualitasChecklist && kisi.kualitasChecklist.length > 0) ? kisi.kualitasChecklist : config.kualitasChecklist;

        const konteksStr = activeKonteksLokal.length > 0 ? `Integrasikan KONTEKS LOKAL INDONESIA berikut: ${activeKonteksLokal.join(", ")}.` : "";
        const stimulusStr = activeStimulusKonten.length > 0 ? `Gunakan tipe STIMULUS: ${activeStimulusKonten.join(", ")}.` : "";
        const checklistStr = activeKualitasChecklist.length > 0 ? `Pastikan KUALITAS SOAL: ${activeKualitasChecklist.join(", ")}.` : "";

        const bentukSoalDesc = kisi.bentukSoal === "pilihan_ganda_sederhana" 
          ? `Pilihan ganda sederhana: Satu jawaban benar (A-${config.jumlahOpsi === 5 ? "E" : "D"}).` 
          : kisi.bentukSoal === "mcma" 
          ? "PG Kompleks MCMA (pilih semua jawaban benar)." 
          : "PG Kompleks Kategori (peserta diminta memberikan respon untuk masing-masing pernyataan berpenomoran angka 1, 2, 3, 4; opsi/kategori respon fleksibel seperti Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Ya/Tidak, Fakta/Opini, Setuju/Tidak Setuju, dll.).";

        const isMatrix = isMatrixOrMathTopic(config.mataPelajaran, kisi.elemenMateri, kisi.subElemenMateri, kisi.kompetensi);
        const isExact = isExactScienceSubject(config.mataPelajaran) || isMatrix;
        const exactPromptRules = isExact ? `
5. KHUSUS MATA PELAJARAN EKSAKTA / FISIKA / MATEMATIKA / KIMIA:
   - WAJIB menuliskan persamaan, reaksi kimia, dan rumus fisika/matematika dengan notasi LaTeX ($...$ atau $$...$$) atau simbol Unicode ilmiah baku (misal: $v_t = v_0 + at$, $s = v_0 t + \\frac{1}{2}at^2$, $F = m \\cdot a$, $E_k = \\frac{1}{2}mv^2$, pecahan \\frac{a}{b}, akar \\sqrt{x}, pangkat x^2 atau 10^{-3}).
   - Cantumkan satuan baku (m/s², kg, J, W, Hz, Ω, μC, °C, K, ±, ×, ÷, ≈, ≤, ≥, ≠, ∞, Δ, α, β, θ, λ, μ, π, ρ, σ, ω).
   - Reaksi kimia dengan subskrip/supersprip yang tepat (\\text{H}_2\\text{SO}_4, \\text{Fe}^{3+}).
${isMatrix ? `   - STANDAR KHUSUS SOAL MATRIKS: Seluruh matriks WAJIB disajikan dengan notasi LaTeX matriks baku (seperti $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$, $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$, atau $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$) dengan pemisah kolom '&' dan pemisah baris '\\\\' yang presisi. Opsi pilihan jawaban A s.d E yang memuat matriks WAJIB disajikan lengkap dalam notasi LaTeX matriks.` : ''}
   - Tabel data eksperimen wajib disajikan dalam format Markdown Table standar yang utuh 1 baris per record.
   - Opsi jawaban A-E wajib homogen dalam format angka, pecahan, desimal, atau satuan.` : '';

        const prompt = `Buatkan tepat ${totalToGenerate} butir soal ujian TKA SMA untuk Mata Pelajaran ${config.mataPelajaran}.
INFORMASI MATRIKS:
- No Soal Mulai: ${currentNoSoal}
- Bentuk Soal: ${kisi.bentukSoal} (${bentukSoalDesc})
- Tingkat Kognitif: ${kisi.levelKognitif}
- Elemen/Materi: ${kisi.elemenMateri}
- Sub-Elemen: ${kisi.subElemenMateri}
- Kompetensi: ${kisi.kompetensi}
- Batasan/Catatan: ${kisi.batasanCatatan || "Tidak ada"}
- Konteks Nusantara: ${kisi.konteksNusantara || "Tidak ada"}
- Stimulus Tambahan: ${kisi.stimulusTambahan || "Tidak ada"}
- Jenis Soal: ${config.jenisSoal}
${konteksStr} ${stimulusStr} ${checklistStr}

PETUNJUK FORMAT SANGAT PENTING:
1. Field 'soal': Tulis langsung teks stimulus/soal secara utuh. Untuk PGK Kategori, bagian 'soal' WAJIB menggabungkan 4 komponen berurutan: (a) Stimulus/narasi/data, (b) Pertanyaan Utama spesifik, (c) Kalimat Perintah Kategori (contoh: "**Berikan respon/pilihan Anda pada masing-masing pernyataan berikut!**"), dan (d) Tabel Pernyataan berbentuk Markdown 4 kolom (| # | Pernyataan | [Kategori A] | [Kategori B] |) dengan nomor urut angka 1, 2, 3, 4 di kolom '#' (minimal 2, maksimal 4 pernyataan; peserta memberikan respon untuk masing-masing pernyataan dengan kategori fleksibel seperti Sesuai/Tidak Sesuai, Benar/Salah, Tepat/Tidak Tepat, Ya/Tidak, Fakta/Opini, Setuju/Tidak Setuju, dll.).
2. Field 'opsi': Untuk PGK Kategori, setiap item dalam array 'opsi' WAJIB diawali nomor angka dan titik, contoh: ['1. ...', '2. ...', '3. ...', '4. ...']. Untuk PG Sederhana/MCMA diawali huruf A, B, C, D, E.
3. Field 'kunciJawaban': Untuk PGK Kategori, rincikan status kategori per nomor (contoh: "Pernyataan 1: [Kategori A], Pernyataan 2: [Kategori B], Pernyataan 3: [Kategori A], Pernyataan 4: [Kategori A]" atau "1. Sesuai, 2. Tidak Sesuai...").
4. Field 'pembahasan': Rincikan penjelasan analitis ilmiah untuk setiap poin nomor 1, 2, 3, 4 satu per satu.${exactPromptRules}

Hasilkan array JSON tepat ${totalToGenerate} objek soal.
GABUNGKAN stimulus langsung ke awal field 'soal' dan kosongkan 'stimulus' (string kosong "").`;

        const soalSchema = {
          type: "ARRAY",
          items: {
            type: "OBJECT",
            properties: {
              kompetensi: { type: "STRING" },
              subKompetensi: { type: "STRING" },
              bentukSoal: { type: "STRING" },
              stimulus: { type: "STRING" },
              soal: { type: "STRING" },
              opsi: { type: "ARRAY", items: { type: "STRING" } },
              kunciJawaban: { type: "STRING" },
              pembahasan: { type: "STRING" },
              kataKunci: { type: "STRING" },
              gambarUrl: { type: "STRING" }
            },
            required: ["kompetensi", "subKompetensi", "bentukSoal", "soal", "opsi", "kunciJawaban", "pembahasan", "kataKunci", "gambarUrl"]
          }
        };

        const responseText = await callGeminiDirect(systemInstruction, prompt, soalSchema);
        data = JSON.parse(responseText);
      } else {
        const response = await fetch('/api/generate-soal', {
          method: 'POST',
          headers: { 
            'Content-Type': 'application/json',
            'x-api-key': getCleanApiKey(aiConfig.apiKey)
          },
          body: JSON.stringify({
            kisi,
            count: totalToGenerate,
            mataPelajaran: config.mataPelajaran,
            definisi: config.definisi,
            muatan: config.muatan,
            jumlahOpsi: config.jumlahOpsi,
            jenisSoal: config.jenisSoal,
            konteksLokal: config.konteksLokal,
            stimulusKonten: config.stimulusKonten,
            kualitasChecklist: config.kualitasChecklist,
            noSoalStart: currentNoSoal,
            existingQuestions: allExistingSoalTexts
          })
        });

        if (!response.ok) {
          let errorMsg = `Gagal menyusun soal untuk kisi-kisi No. ${kisi.no}`;
          try {
            const textError = await response.text();
            if (textError.includes('<!doctype') || textError.includes('<html')) {
              errorMsg = 'Server AI mengalami timeout atau sedang sibuk (Gateway Timeout). Silakan coba beberapa saat lagi.';
            } else {
              try {
                const errorData = JSON.parse(textError);
                errorMsg = errorData.error || errorMsg;
              } catch {
                errorMsg = textError || errorMsg;
              }
            }
          } catch {}
          throw new Error(errorMsg);
        }
        data = await response.json();
      }

      if (Array.isArray(data) && data.length > 0) {
        const mapped: Question[] = data.map((q: any, idx: number) => {
          const rawOpsi = Array.isArray(q.opsi) ? q.opsi : [];
          const normalizedOpsi = rawOpsi.map((opt: string, oIdx: number) => formatOptionString(opt, oIdx));
          return {
            id: `q-ai-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${idx}`,
            noSoal: currentNoSoal + idx,
            kisiKisiId: kisi.id,
            kompetensi: q.kompetensi || kisi.kompetensi,
            subKompetensi: q.subKompetensi || kisi.subElemenMateri,
            bentukSoal: kisi.bentukSoal,
            soal: cleanSoalText(q.soal || ''),
            stimulus: q.stimulus || '',
            opsi: normalizedOpsi,
            kunciJawaban: (q.kunciJawaban || 'A').trim().toUpperCase(),
            pembahasan: q.pembahasan || 'Pembahasan terstruktur.',
            kataKunci: q.kataKunci || '',
            gambarUrl: q.gambarUrl || ''
          };
        });

        setQuestions(prev => {
          const existingIds = new Set(prev.map(p => p.id));
          const filtered = mapped.filter(m => !existingIds.has(m.id));
          return [...prev, ...filtered];
        });

        if (currentUser?.uid) {
          for (const newQ of mapped) {
            try {
              await setDoc(doc(db, 'questions', newQ.id), newQ);
            } catch (err) {
              console.warn("Failed syncing to firestore:", err);
            }
          }
        }

        alert(`🎉 Berhasil membuat ${mapped.length} butir soal untuk kisi-kisi No. ${kisi.no}!`);
        setActiveTab('soal');
      } else {
        throw new Error("Respon yang dihasilkan kosong atau format tidak sesuai.");
      }
    } catch (err: any) {
      console.error('Error generating for kisi:', err);
      alert(`⚠️ Gagal menyusun soal: ${err.message || 'Error'}`);
    } finally {
      setIsGeneratingSoal(false);
      setSoalProgress(prev => ({ ...prev, active: false }));
    }
  };

  // Generate Custom SVG illustration via AI
  const handleGenerateCustomIllustration = async () => {
    if (!aiIllustratorPrompt.trim()) {
      alert('Sila tulis deskripsi ilustrasi atau grafik yang ingin Anda buat!');
      return;
    }

    setIsGeneratingIllustration(true);
    setAiIllustratorStatus('Menghubungkan ke Gemini AI...');

    const statuses = [
      'Menganalisis permintaan ilustrasi...',
      'Merancang kerangka koordinat dan objek geometri...',
      'Menggambar garis, kurva, dan bentuk presisi...',
      'Menambahkan pelabelan teks dan rumus...',
      'Mengoptimasi pewarnaan dan responsivitas kode SVG...',
      'Hampir selesai, memformat kode vektor...'
    ];

    let statusIndex = 0;
    const interval = setInterval(() => {
      if (statusIndex < statuses.length) {
        setAiIllustratorStatus(statuses[statusIndex]);
        statusIndex++;
      }
    }, 1500);

    try {
      let data;
      if (aiConfig.mode === 'client') {
        const systemInstruction = `Anda adalah desainer grafis vektor SVG profesional untuk konten edukasi sains, matematika, dan ilmu sosial. Hasilkan HANYA kode SVG inline lengkap yang valid, dimulai dengan '<svg' dan diakhiri dengan '</svg>' tanpa penjelasan markdown atau sapaan lainnya. Pastikan SVG menggunakan atribut viewBox agar responsif, berwarna elegan dengan skema modern, kontras tinggi yang jelas di latar belakang putih.`;
        const promptText = `Buatkan grafis vektor SVG profesional dan edukatif berdasarkan instruksi berikut:
"${aiIllustratorPrompt}"

Konteks soal/konten:
"${questionForm.soal || ''}"

Ingat: HANYA berikan kode SVG murni. Jika Anda membungkusnya dengan blok markdown seperti \`\`\`xml atau \`\`\`html, pastikan bagian luar dibersihkan. Namun lebih baik langsung string SVG murni dimulai dengan <svg.`;

        let responseText = await callGeminiDirect(systemInstruction, promptText);
        clearInterval(interval);
        
        // Clean markdown blocks if returned by any chance
        if (responseText.includes('```')) {
          responseText = responseText.replace(/```[a-z]*\n?/gi, '').trim();
        }
        
        data = { svg: responseText };
      } else {
        const response = await fetch('/api/generate-illustration', {
          method: 'POST',
          headers: { 
            'Content-Type': 'application/json',
            'x-api-key': getCleanApiKey(aiConfig.apiKey)
          },
          body: JSON.stringify({
            prompt: aiIllustratorPrompt,
            context: questionForm.soal || ''
          })
        });

        clearInterval(interval);

        if (!response.ok) {
          let errorMsg = 'Gagal menghasilkan gambar.';
          try {
            const textError = await response.text();
            if (textError.includes('<!doctype') || textError.includes('<html')) {
              errorMsg = 'Server sedang sibuk atau mengalami timeout (Gateway Timeout).';
            } else {
              try {
                const errorData = JSON.parse(textError);
                errorMsg = errorData.error || errorMsg;
              } catch {
                errorMsg = textError || errorMsg;
              }
            }
          } catch {}
          throw new Error(errorMsg);
        }

        try {
          const responseText = await response.text();
          data = JSON.parse(responseText);
        } catch {
          throw new Error('Respon dari server tidak valid.');
        }
      }

      if (data.svg) {
        setQuestionForm(prev => ({ ...prev, gambarUrl: data.svg }));
        setIsAiIllustratorOpen(false);
        setAiIllustratorPrompt('');
      } else {
        throw new Error('Kode gambar kosong atau tidak valid.');
      }
    } catch (err: any) {
      clearInterval(interval);
      console.error(err);
      alert(`Gagal merancang ilustrasi: ${err.message}`);
    } finally {
      setIsGeneratingIllustration(false);
    }
  };

  // States for Masukkan Soal AI Modal
  const [isImportAiModalOpen, setIsImportAiModalOpen] = useState(false);
  const [importAiRawText, setImportAiRawText] = useState('');
  const [importAiParsedQuestions, setImportAiParsedQuestions] = useState<Question[]>([]);
  const [importAiFileName, setImportAiFileName] = useState<string | null>(null);
  const [importAiError, setImportAiError] = useState<string | null>(null);
  const [isImportingAi, setIsImportingAi] = useState(false);

  // State to track detected parser strategy in UI
  const [importAiDetectedFormat, setImportAiDetectedFormat] = useState<string>('');

  // Helper function to enrich parsed questions with Matrix Kisi-Kisi if available
  const enrichWithMatrix = (qList: Question[]): Question[] => {
    if (!kisiList || kisiList.length === 0) return qList;

    return qList.map((q, idx) => {
      const qNum = q.noSoal || (idx + 1);

      // Calculate cumulative question range for each Kisi-Kisi row
      let matchedKisi: KisiKisiItem | undefined;
      let currStart = 1;
      for (const k of kisiList) {
        const qty = k.jumlahSoal || 1;
        const currEnd = currStart + qty - 1;
        if (qNum >= currStart && qNum <= currEnd) {
          matchedKisi = k;
          break;
        }
        currStart += qty;
      }

      if (!matchedKisi) {
        matchedKisi = kisiList.find(k => k.no === qNum) || kisiList[idx] || kisiList[kisiList.length - 1];
      }

      if (matchedKisi) {
        const mainKompetensi = (!q.kompetensi || q.kompetensi === 'Kompetensi Asesmen TKA SMA' || q.kompetensi === matchedKisi.subElemenMateri)
          ? matchedKisi.elemenMateri
          : q.kompetensi;

        const subKompetensi = (!q.subKompetensi || q.subKompetensi === 'Materi TKA SMA')
          ? matchedKisi.subElemenMateri
          : q.subKompetensi;

        return {
          ...q,
          noSoal: qNum,
          kisiKisiId: matchedKisi.id,
          kompetensi: mainKompetensi,
          subKompetensi: subKompetensi,
          bentukSoal: q.bentukSoal || matchedKisi.bentukSoal || 'pilihan_ganda_sederhana'
        };
      }
      return q;
    });
  };

  // Helper parser for raw text or JSON into Question objects
  const parseAiQuestionsText = (rawText: string): Question[] => {
    if (!rawText || !rawText.trim()) {
      setImportAiDetectedFormat('');
      return [];
    }

    const textToParse = rawText.trim();
    let parsedQuestions: Question[] = [];
    const startNo = (questions.reduce((max, q) => Math.max(max, Number(q.noSoal) || 0), 0)) + 1;

    // Helper to safely split markdown row cells while respecting math delimiters ($) and LaTeX matrices (\begin{pmatrix} etc.)
    const splitMarkdownCells = (rowLine: string): string[] => {
      let t = rowLine.trim();
      if (t.startsWith('|')) t = t.substring(1);
      if (t.endsWith('|')) t = t.substring(0, t.length - 1);

      const result: string[] = [];
      let current = '';
      let inMath = false;
      let matrixDepth = 0;

      for (let i = 0; i < t.length; i++) {
        const ch = t[i];
        const nextCh = t[i + 1] || '';

        // Check LaTeX math delimiter $
        if (ch === '$' && (i === 0 || t[i - 1] !== '\\')) {
          inMath = !inMath;
          current += ch;
          continue;
        }

        // Check \begin{...} / \end{...}
        if (t.startsWith('\\begin{', i)) {
          matrixDepth++;
        } else if (t.startsWith('\\end{', i) && matrixDepth > 0) {
          matrixDepth--;
        }

        // Check escaped pipe \|
        if (ch === '\\' && nextCh === '|') {
          current += '|';
          i++; // skip
          continue;
        }

        // Pipe separator when not inside math and not inside matrix
        if (ch === '|' && !inMath && matrixDepth === 0) {
          result.push(current.trim());
          current = '';
          continue;
        }

        current += ch;
      }
      result.push(current.trim());
      return result;
    };

    // ----------------------------------------------------
    // STRATEGY 1: JSON PARSER
    // ----------------------------------------------------
    let jsonMatch = textToParse.match(/\[\s*\{[\s\S]*\}\s*\]/);
    if (!jsonMatch && textToParse.startsWith('[') && textToParse.endsWith(']')) {
      jsonMatch = [textToParse];
    }

    if (jsonMatch) {
      try {
        let rawJson = jsonMatch[0];
        // Pre-sanitize LaTeX backslashes if standard JSON parse throws
        let parsedJson: any;
        try {
          parsedJson = JSON.parse(rawJson);
        } catch {
          const sanitized = rawJson
            .replace(/\\(frac|forall|flat|bar|beta|binom|bold|begin|bmatrix|pmatrix|vmatrix|Vmatrix|Bmatrix|bullet|nabla|neq|not|nu|rho|right|rightarrow|rightleftharpoons|rangle|times|theta|text|tau|to|tan|top|tilde|det|pmatrix|bmatrix|vmatrix|cdot|approx|leq|geq|pm|div|degree|alpha|gamma|delta|epsilon|lambda|mu|pi|sigma|omega|sum|int|sqrt|array)\b/gi, '\\\\$1')
            .replace(/\\end\{(pmatrix|bmatrix|vmatrix|Vmatrix|Bmatrix|matrix|smallmatrix|array)\}/gi, '\\\\end{$1}')
            .replace(/\\begin\{(pmatrix|bmatrix|vmatrix|Vmatrix|Bmatrix|matrix|smallmatrix|array)\}/gi, '\\\\begin{$1}')
            .replace(/\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})/g, '\\\\')
            .replace(/,\s*([}\]])/g, '$1');
          parsedJson = JSON.parse(sanitized);
        }

        if (Array.isArray(parsedJson) && parsedJson.length > 0) {
          parsedJson.forEach((item, idx) => {
            if (typeof item === 'object' && item !== null) {
              const questionBody = item.soal || item.pertanyaan || item.question || item.item || item.textSoal || '';
              if (questionBody) {
                let rawOptions = item.opsi || item.pilihan || item.options || [];
                if (typeof rawOptions === 'string') {
                  rawOptions = rawOptions.split(/\n+/).map((o: string) => o.trim()).filter(Boolean);
                } else if (!Array.isArray(rawOptions) && typeof item.opsiA === 'string') {
                  rawOptions = [item.opsiA, item.opsiB, item.opsiC, item.opsiD, item.opsiE].filter(Boolean);
                }

                const cleanOpts = (Array.isArray(rawOptions) ? rawOptions : []).map((optStr: string, oIdx: number) => {
                  return formatOptionString(optStr, oIdx);
                });

                let bSoal: BentukSoal = 'pilihan_ganda_sederhana';
                const rawBentuk = String(item.bentukSoal || item.bentuk || '').toLowerCase();
                if (rawBentuk.includes('mcma') || rawBentuk.includes('kompleks') || rawBentuk.includes('banyak')) {
                  bSoal = 'mcma';
                } else if (rawBentuk.includes('kategori') || rawBentuk.includes('jodoh') || rawBentuk.includes('pernyataan')) {
                  bSoal = 'kategori';
                }

                const uniqueId = `q-ai-import-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${idx + 1}`;
                parsedQuestions.push({
                  id: uniqueId,
                  userId: currentUser?.uid,
                  noSoal: item.noSoal || item.no || (startNo + idx),
                  kisiKisiId: item.kisiKisiId || '',
                  kompetensi: item.kompetensi || item.materi || item.elemen || config.kompetensi || 'Kompetensi Asesmen TKA SMA',
                  subKompetensi: item.subKompetensi || item.submateri || item.subElemen || config.subElemenMateri || 'Materi TKA SMA',
                  bentukSoal: bSoal,
                  soal: questionBody.trim(),
                  stimulus: (item.stimulus || item.wacana || item.ringkasanStimulus || '').replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
                  opsi: cleanOpts.length > 0 ? cleanOpts : ['A. Opsi A', 'B. Opsi B', 'C. Opsi C', 'D. Opsi D', 'E. Opsi E'],
                  kunciJawaban: String(item.kunciJawaban || item.kunci || item.jawaban || 'A').replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
                  pembahasan: (item.pembahasan || item.penjelasan || item.rationale || 'Pembahasan disesuaikan dengan naskah.').replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
                  kataKunci: item.kataKunci || `${config.mataPelajaran || 'TKA SMA'}, ${config.elemenMateri || ''}`
                });
              }
            }
          });

          if (parsedQuestions.length > 0) {
            setImportAiDetectedFormat('JSON Array');
            return enrichWithMatrix(parsedQuestions);
          }
        }
      } catch (e) {
        console.warn("JSON parsing skipped:", e);
      }
    }

    // ----------------------------------------------------
    // STRATEGY 2: MARKDOWN TABLE / TSV PARSER (BAGIAN 2 MEGAPROMPT)
    // ----------------------------------------------------
    const lines = textToParse.replace(/\r\n/g, '\n').split('\n');
    const tableHeaderIdx = lines.findIndex(l => {
      const lower = l.toLowerCase();
      return (lower.includes('no') || lower.includes('nomor')) &&
             (lower.includes('soal') || lower.includes('pertanyaan')) &&
             (lower.includes('opsi') || lower.includes('materi') || lower.includes('kunci'));
    });

    if (tableHeaderIdx !== -1) {
      const headerLine = lines[tableHeaderIdx];
      const isMarkdown = headerLine.includes('|');
      const delimiter = isMarkdown ? '|' : '\t';

      const rawHeaders = headerLine.split(delimiter).map(h => h.replace(/\*/g, '').trim()).filter(Boolean);

      if (rawHeaders.length >= 5) {
        const getColIdx = (keywords: string[]) => {
          return rawHeaders.findIndex(h => {
            const hl = h.toLowerCase();
            return keywords.some(kw => hl.includes(kw));
          });
        };

        const materiIdx = getColIdx(['elemen', 'materi utama', 'materi', 'kompetensi']);
        const submateriIdx = getColIdx(['submateri', 'sub-elemen', 'sub', 'sub kompetensi']);
        const bentukIdx = getColIdx(['bentuk']);
        const stimIdx = getColIdx(['stimulus', 'wacana', 'ringkasan']);
        const soalIdx = getColIdx(['pertanyaan', 'teks soal', 'soal']);
        const optAIdx = getColIdx(['opsi a', 'pilihan a']);
        const optBIdx = getColIdx(['opsi b', 'pilihan b']);
        const optCIdx = getColIdx(['opsi c', 'pilihan c']);
        const optDIdx = getColIdx(['opsi d', 'pilihan d']);
        const optEIdx = getColIdx(['opsi e', 'pilihan e']);
        const optCombIdx = getColIdx(['opsi', 'pilihan']);
        const kunciIdx = getColIdx(['kunci']);
        const pembIdx = getColIdx(['pembahasan', 'penjelasan', 'rationale']);

        let countRow = 0;
        for (let i = tableHeaderIdx + 1; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line || line.startsWith('|---') || line.startsWith('| :-') || line.startsWith('---')) continue;

          const cleanCells = isMarkdown 
            ? splitMarkdownCells(line)
            : line.split('\t').map(c => c.trim());

          if (cleanCells.length < 3) continue;

          const qSoalText = soalIdx !== -1 && cleanCells[soalIdx] ? cleanCells[soalIdx] : '';
          if (!qSoalText || qSoalText.length < 3) continue;

          countRow++;
          const qMateri = materiIdx !== -1 && cleanCells[materiIdx] ? cleanCells[materiIdx] : (config.elemenMateri || 'Materi Utama');
          const qSubMateri = submateriIdx !== -1 && cleanCells[submateriIdx] ? cleanCells[submateriIdx] : (config.subElemenMateri || 'Submateri');
          const qStim = stimIdx !== -1 && cleanCells[stimIdx] ? cleanCells[stimIdx] : '';

          let bSoal: BentukSoal = 'pilihan_ganda_sederhana';
          if (bentukIdx !== -1 && cleanCells[bentukIdx]) {
            const bt = cleanCells[bentukIdx].toLowerCase();
            if (bt.includes('mcma') || bt.includes('kompleks') || bt.includes('banyak')) bSoal = 'mcma';
            else if (bt.includes('kategori') || bt.includes('jodoh')) bSoal = 'kategori';
          }

          let opts: string[] = [];
          if (optAIdx !== -1 && cleanCells[optAIdx]) {
            const a = cleanCells[optAIdx].replace(/^[A-E]\.\s*/i, '');
            const b = optBIdx !== -1 && cleanCells[optBIdx] ? cleanCells[optBIdx].replace(/^[A-E]\.\s*/i, '') : '';
            const c = optCIdx !== -1 && cleanCells[optCIdx] ? cleanCells[optCIdx].replace(/^[A-E]\.\s*/i, '') : '';
            const d = optDIdx !== -1 && cleanCells[optDIdx] ? cleanCells[optDIdx].replace(/^[A-E]\.\s*/i, '') : '';
            const e = optEIdx !== -1 && cleanCells[optEIdx] ? cleanCells[optEIdx].replace(/^[A-E]\.\s*/i, '') : '';

            opts = [`A. ${a}`, `B. ${b}`, `C. ${c}`, `D. ${d}`];
            if (e) opts.push(`E. ${e}`);
          } else if (optCombIdx !== -1 && cleanCells[optCombIdx]) {
            const rawComb = cleanCells[optCombIdx];
            opts = rawComb.split(/[\n;]+/).map(s => s.trim()).filter(Boolean);
          }

          if (opts.length === 0) {
            opts = ['A. Opsi A', 'B. Opsi B', 'C. Opsi C', 'D. Opsi D', 'E. Opsi E'];
          }

          const qKunci = kunciIdx !== -1 && cleanCells[kunciIdx] ? cleanCells[kunciIdx] : 'A';
          const qPemb = pembIdx !== -1 && cleanCells[pembIdx] ? cleanCells[pembIdx] : 'Pembahasan sesuai naskah AI.';

          const uniqueId = `q-ai-import-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${countRow}`;
          parsedQuestions.push({
            id: uniqueId,
            userId: currentUser?.uid,
            noSoal: startNo + countRow - 1,
            kisiKisiId: '',
            kompetensi: qMateri,
            subKompetensi: qSubMateri,
            bentukSoal: bSoal,
            soal: qSoalText.trim(),
            stimulus: qStim.replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
            opsi: opts,
            kunciJawaban: qKunci.replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
            pembahasan: qPemb.replace(/\*\*([^*]+)\*\*/g, '$1').trim(),
            kataKunci: `${config.mataPelajaran || 'TKA SMA'}, ${qMateri}`
          });
        }

        if (parsedQuestions.length > 0) {
          setImportAiDetectedFormat('Tabel Matriks Excel/Markdown (Bagian 2)');
          return enrichWithMatrix(parsedQuestions);
        }
      }
    }

    // ----------------------------------------------------
    // STRATEGY 3: STRUCTURED TEXT / NASKAH WORD (.DOCX) PARSER
    // Matches Butir Soal TKA SMA format (Bagian 1 / Word Export)
    // ----------------------------------------------------
    let cleanedText = textToParse.replace(/\r\n/g, '\n');
    
    // Cut off [BAGIAN 2: ...] if present at the end of text so table rows don't append to last question
    const bagian2Idx = cleanedText.indexOf('[BAGIAN 2');
    if (bagian2Idx > 100) {
      cleanedText = cleanedText.substring(0, bagian2Idx);
    }

    // Split into blocks by question header patterns
    const questionHeaderRegex = /(?:\n+|^)(?=(?:#{1,6}\s*|\*{0,2}\s*)?(?:\[?\b(?:NO\.?\s*SOAL|SOAL\s*(?:NO\.?|NOMOR)?|BUTIR\s*SOAL|KISI-KISI\s*NO\.?|PERTANYAAN)\s*\d+\]?|\bNO\.?\s*\d+[\.\:\)]|\d+[\.\)]\s*(?:\*{0,2}\s*)?(?:\[?\bSOAL|\[?\bSPESIFIKASI|\bSTIMULUS|\bPERTANYAAN|\bKOMPETENSI)))/i;

    const rawBlocks = cleanedText.split(questionHeaderRegex).map(b => b.trim()).filter(b => b.length > 15);

    rawBlocks.forEach((blockStr, idx) => {
      if (idx === 0 && !/(?:SOAL|STIMULUS|PERTANYAAN|OPSI|KUNCI|PEMBAHASAN|KOMPETENSI)/i.test(blockStr)) {
        return;
      }

      const blockLines = blockStr.split('\n').map(l => l.trim()).filter(Boolean);

      let noSoalVal = startNo + idx;
      let kompetensiVal = config.kompetensi || 'Kompetensi Asesmen TKA SMA';
      let subKompetensiVal = config.subElemenMateri || 'Materi TKA SMA';
      let bentukSoalVal: BentukSoal = 'pilihan_ganda_sederhana';
      let stimulusVal = '';
      let kunciVal = 'A';
      let pembahasanVal = 'Pembahasan telah disesuaikan dengan naskah Gemini AI.';

      // Extract metadata with optional colons
      const noMatch = blockStr.match(/(?:NO\.?\s*SOAL|SOAL\s*(?:NO\.?|NOMOR)?|BUTIR\s*SOAL)\s*[\:\=]?\s*(\d+)/i) || blockStr.match(/^(\d+)[\.\)]/);
      if (noMatch) {
        noSoalVal = parseInt(noMatch[1], 10);
      }

      const komMatch = blockStr.match(/(?:Kompetensi|Materi\s*Utama|Elemen|Materi)\s*[\:\=]?\s*([^\n]+)/i);
      if (komMatch) {
        const raw = komMatch[1].replace(/\*/g, '').trim();
        if (raw && !raw.toLowerCase().startsWith('sub')) {
          kompetensiVal = raw;
        }
      }

      const subMatch = blockStr.match(/(?:Sub\s*Kompetensi|Submateri|Sub-Elemen|Sub\s*Materi)\s*[\:\=]?\s*([^\n]+)/i);
      if (subMatch) {
        subKompetensiVal = subMatch[1].replace(/\*/g, '').trim();
      }

      const bentukMatch = blockStr.match(/(?:Bentuk\s*Soal|Bentuk)\s*[\:\=]?\s*([^\n]+)/i);
      if (bentukMatch) {
        const bt = bentukMatch[1].toLowerCase();
        if (bt.includes('mcma') || bt.includes('kompleks') || bt.includes('banyak')) {
          bentukSoalVal = 'mcma';
        } else if (bt.includes('kategori') || bt.includes('jodoh') || bt.includes('pernyataan')) {
          bentukSoalVal = 'kategori';
        } else {
          bentukSoalVal = 'pilihan_ganda_sederhana';
        }
      } else {
        if (/MCMA|Pilihan Ganda Kompleks|Banyak Jawaban/i.test(blockStr)) bentukSoalVal = 'mcma';
        else if (/Kategori|Menjodohkan|Pernyataan|Pilihlah Sesuai|Sesuai atau Tidak Sesuai|Benar atau Salah|Tepat atau Tidak Tepat/i.test(blockStr)) bentukSoalVal = 'kategori';
      }

      if (bentukSoalVal === 'pilihan_ganda_sederhana' && /Pilihlah\s+(?:Sesuai|Benar|Tepat|Ya)|Sesuai\s*atau\s*Tidak\s*Sesuai|Benar\s*atau\s*Salah|Tepat\s*atau\s*Tidak\s*Tepat|Kategori|Tabel\s*Pernyataan|Pernyataan\s*1/i.test(blockStr)) {
        bentukSoalVal = 'kategori';
      }

      const kunciMatch = blockStr.match(/(?:KUNCI\s*JAWABAN|KUNCI|JAWABAN\s*BENAR|ANSWER)\s*[\:\=]?\s*([^\n]+)/i);
      if (kunciMatch) {
        kunciVal = kunciMatch[1].replace(/\*/g, '').trim();
      }

      const pemmMatch = blockStr.match(/(?:PEMBAHASAN|RATIONALE|PENJELASAN|ALASAN)\s*(?:MENDALAM)?\s*[\:\=]?\s*([\s\S]*?)(?=(?:\[?SOAL|\[?KISI-KISI|\[?BAGIAN|#{1,6}|\*{0,2}SOAL|NO\.?\s*SOAL|NO\.?\s*\d+|$))/i);
      if (pemmMatch) {
        pembahasanVal = pemmMatch[1].replace(/\*\*([^*]+)\*\*/g, '$1').trim();
      }

      const stimMatch = blockStr.match(/(?:STIMULUS|WACANA|KASUS|TEKS|NARASI)\s*(?:LENGKAP)?\s*[\:\=]?\s*([\s\S]*?)(?=(?:PERTANYAAN|SOAL|BUTIR|OPSI|A\s*[\.\)]|\*\*A\.\*\*|1\s*[\.\)]|KUNCI|PEMBAHASAN|$))/i);
      if (stimMatch) {
        stimulusVal = stimMatch[1].replace(/\*\*([^*]+)\*\*/g, '$1').trim();
      }

      // Line by line state parser for Soal & Options
      let section: 'metadata' | 'soal' | 'stimulus' | 'options' | 'footer' = 'metadata';
      const soalLinesArr: string[] = [];
      const optionLinesArr: string[] = [];

      for (let i = 0; i < blockLines.length; i++) {
        const line = blockLines[i];

        // Check if line is metadata key or footer key (Kunci/Pembahasan/Pedoman/Rubrik/Kompetensi/Sub Kompetensi/Bentuk)
        if (/^(?:NO\.?\s*SOAL|SOAL\s*(?:NO|NOMOR)?|BUTIR\s*SOAL|KOMPETENSI|SUB\s*KOMPETENSI|BENTUK\s*SOAL|KUNCI|JAWABAN|PEMBAHASAN|PENJELASAN|RATIONALE|PEDOMAN|RUBRIK|BOBOT|LEVEL)/i.test(line)) {
          if (/^(?:KUNCI|JAWABAN|PEMBAHASAN|PENJELASAN|RATIONALE|PEDOMAN|RUBRIK|BOBOT)/i.test(line)) {
            section = 'footer';
          }
          continue;
        }

        // If currently in footer section, do NOT append to options or soal
        if (section === 'footer') {
          continue;
        }

        // Check if line is Pertanyaan header
        if (/^(?:PERTANYAAN|SOAL|BUTIR\s*PERTANYAAN|TEKS\s*SOAL)\s*[\:\=]?/i.test(line)) {
          section = 'soal';
          const afterTag = line.replace(/^(?:PERTANYAAN|SOAL|BUTIR\s*PERTANYAAN|TEKS\s*SOAL)\s*[\:\=]?\s*/i, '').trim();
          if (afterTag) {
            soalLinesArr.push(afterTag);
          }
          continue;
        }

        // Check if line is Stimulus header
        if (/^(?:STIMULUS|WACANA|KASUS|NARASI)\s*[\:\=]?/i.test(line)) {
          section = 'stimulus';
          const afterTag = line.replace(/^(?:STIMULUS|WACANA|KASUS|NARASI)\s*[\:\=]?\s*/i, '').trim();
          if (afterTag && !stimulusVal) {
            stimulusVal = afterTag;
          }
          continue;
        }

        // Check if line is an Option start (A., B., C., D., E. or 1., 2., 3., 4., 5.)
        const optMatch = line.match(/^(?:[\*\-\•\s]*)?(?:\*{0,2})([A-E1-5])(?:\*{0,2})\s*[\.\)]\s*(.+)/i);
        if (optMatch) {
          section = 'options';
          const letter = optMatch[1].toUpperCase();
          const cleanOptText = optMatch[2].replace(/\*\*([^*]+)\*\*/g, '$1').trim();
          let normLetter = letter;
          if (bentukSoalVal === 'kategori') {
            const numIdx = ['1','2','3','4','5'].includes(letter)
              ? parseInt(letter, 10)
              : (letter.charCodeAt(0) - 64);
            normLetter = String(numIdx);
          } else {
            if (['1','2','3','4','5'].includes(letter)) {
              normLetter = String.fromCharCode(64 + parseInt(letter, 10));
            }
          }
          optionLinesArr.push(`${normLetter}. ${cleanOptText}`);
          continue;
        }

        // If in options section and line is a multi-line continuation of current option
        if (section === 'options' && optionLinesArr.length > 0) {
          optionLinesArr[optionLinesArr.length - 1] += ` ${line.replace(/\*\*([^*]+)\*\*/g, '$1').trim()}`;
          continue;
        }

        // If in soal/stimulus/metadata section and line contains question text
        if (section === 'soal' || section === 'metadata' || section === 'stimulus') {
          if (section === 'metadata' && line.length > 5) {
            section = 'soal';
          }
          if (section === 'soal') {
            soalLinesArr.push(line.replace(/\*\*([^*]+)\*\*/g, '$1').trim());
          }
        }
      }

      let soalStr = soalLinesArr.join(' ').trim();
      soalStr = soalStr.replace(/^\[?SOAL\s*(?:NO\.?|NOMOR)?\s*\d+\]?\s*[\:\.\-]?\s*/i, '').trim();
      soalStr = soalStr.replace(/^NO\.?\s*SOAL\s*\d+[\.\:\)]?\s*/i, '').trim();

      if (soalStr.length > 5 || optionLinesArr.length > 0) {
        const uniqueId = `q-ai-import-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${idx + 1}`;
        
        parsedQuestions.push({
          id: uniqueId,
          userId: currentUser?.uid,
          noSoal: noSoalVal,
          kisiKisiId: '',
          kompetensi: kompetensiVal,
          subKompetensi: subKompetensiVal,
          bentukSoal: bentukSoalVal,
          soal: (soalStr || `Butir Soal #${noSoalVal} dari Gemini AI`).trim(),
          stimulus: stimulusVal,
          opsi: optionLinesArr.length > 0 ? optionLinesArr : ['A. Opsi A', 'B. Opsi B', 'C. Opsi C', 'D. Opsi D', 'E. Opsi E'],
          kunciJawaban: kunciVal,
          pembahasan: pembahasanVal,
          kataKunci: `${config.mataPelajaran || 'TKA SMA'}, ${kompetensiVal}`
        });
      }
    });

    if (parsedQuestions.length > 0) {
      setImportAiDetectedFormat('Naskah Soal Word (Bagian 1)');
      return enrichWithMatrix(parsedQuestions);
    }

    setImportAiDetectedFormat('Format Bebas');
    return parsedQuestions;
  };

  // Helper to parse Excel ArrayBuffer (.xlsx, .xls, .csv) into Questions
  const parseExcelFileToQuestions = (arrayBuffer: ArrayBuffer, fileName: string = ''): Question[] => {
    try {
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      if (!sheetName) return [];

      const ws = wb.Sheets[sheetName];
      const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (!rawRows || rawRows.length === 0) return [];

      const startNo = (questions.reduce((max, q) => Math.max(max, Number(q.noSoal) || 0), 0)) + 1;
      const parsedQuestions: Question[] = [];

      rawRows.forEach((row, idx) => {
        const keys = Object.keys(row);

        const getVal = (keywords: string[]) => {
          // 1. Try exact match first
          for (const kw of keywords) {
            const exactKey = keys.find(k => k.toLowerCase().trim() === kw.toLowerCase().trim());
            if (exactKey && row[exactKey] !== undefined && row[exactKey] !== '') {
              return String(row[exactKey]).trim();
            }
          }
          // 2. Try substring match
          for (const kw of keywords) {
            const subKey = keys.find(k => k.toLowerCase().trim().includes(kw.toLowerCase().trim()));
            if (subKey && row[subKey] !== undefined && row[subKey] !== '') {
              return String(row[subKey]).trim();
            }
          }
          return '';
        };

        const qNo = parseInt(getVal(['no soal', 'no', 'nomor']), 10) || (startNo + idx);
        const qSoalTextRaw = getVal(['soal (stimulus + pertanyaan)', 'pertanyaan/soal', 'soal', 'pertanyaan', 'teks soal', 'question', 'item']);
        
        if (!qSoalTextRaw) return;

        let qSoalText = qSoalTextRaw.trim();

        const qBentukRaw = getVal(['bentuk soal', 'bentuk', 'tipe']).toLowerCase();
        const qKunci = getVal(['kunci jawaban', 'kunci_jawaban', 'kunci', 'jawaban']) || 'A';
        const qFullText = `${qSoalText} ${qKunci}`.toLowerCase();

        const lowerFn = (fileName || '').toLowerCase();
        let bSoal: BentukSoal = 'pilihan_ganda_sederhana';
        if (qBentukRaw.includes('mcma') || qBentukRaw.includes('kompleks') || qBentukRaw.includes('banyak')) {
          bSoal = 'mcma';
        } else if (qBentukRaw.includes('kategori') || qBentukRaw.includes('jodoh') || qBentukRaw.includes('pernyataan') || /pilihlah\s+(?:sesuai|benar|tepat|ya)/i.test(qFullText) || /sesuai\s*atau\s*tidak\s*sesuai/i.test(qFullText) || /benar\s*atau\s*salah/i.test(qFullText) || /tepat\s*atau\s*tidak\s*tepat/i.test(qFullText)) {
          bSoal = 'kategori';
        } else if (!qBentukRaw) {
          if (lowerFn.includes('pgk_kategori') || lowerFn.includes('kategori')) {
            bSoal = 'kategori';
          } else if (lowerFn.includes('pg_mcma') || lowerFn.includes('mcma')) {
            bSoal = 'mcma';
          } else if (lowerFn.includes('pg_sederhana') || lowerFn.includes('sederhana')) {
            bSoal = 'pilihan_ganda_sederhana';
          }
        }

        const qKompetensi = getVal(['kompetensi', 'elemen', 'materi utama', 'materi']) || config.kompetensi || 'Kompetensi Asesmen TKA SMA';
        const qSubKompetensi = getVal(['sub kompetensi', 'subkompetensi', 'submateri', 'sub-elemen', 'sub materi']) || config.subElemenMateri || 'Materi TKA SMA';
        
        // Stimulus logic
        let qStimulus = getVal(['stimulus', 'wacana', 'kasus', 'narasi']);
        // If qStimulus matched the same column as qSoalText, clear qStimulus to prevent duplication
        if (qStimulus === qSoalText) {
          qStimulus = '';
        }

        const optA = getVal(['opsi_a', 'opsi a', 'pilihan a', 'pilihan_a', 'a', 'pernyataan_1', 'pernyataan 1', 'pernyataan1', 'opsi_1', 'opsi 1']);
        const optB = getVal(['opsi_b', 'opsi b', 'pilihan b', 'pilihan_b', 'b', 'pernyataan_2', 'pernyataan 2', 'pernyataan2', 'opsi_2', 'opsi 2']);
        const optC = getVal(['opsi_c', 'opsi c', 'pilihan c', 'pilihan_c', 'c', 'pernyataan_3', 'pernyataan 3', 'pernyataan3', 'opsi_3', 'opsi 3']);
        const optD = getVal(['opsi_d', 'opsi d', 'pilihan d', 'pilihan_d', 'd', 'pernyataan_4', 'pernyataan 4', 'pernyataan4', 'opsi_4', 'opsi 4']);
        const optE = getVal(['opsi_e', 'opsi e', 'pilihan e', 'pilihan_e', 'e', 'pernyataan_5', 'pernyataan 5', 'pernyataan5', 'opsi_5', 'opsi 5']);

        let opts: string[] = [];
        if (optA || optB || optC || optD) {
          const rawOpts = [optA, optB, optC, optD];
          if (optE) rawOpts.push(optE);

          opts = rawOpts.map((opt, oIdx) => {
            if (!opt) return '';
            const cleaned = cleanOptionText(opt);
            if (!cleaned || cleaned === '-' || cleaned === '–') return '';
            return bSoal === 'kategori' ? `${oIdx + 1}. ${cleaned}` : formatOptionString(opt, oIdx);
          }).filter(Boolean);
        } else {
          const optComb = getVal(['opsi', 'pilihan']);
          if (optComb) {
            opts = optComb.split(/[\n;]+/).map((s, oIdx) => {
              const cleaned = cleanOptionText(s);
              if (!cleaned || cleaned === '-' || cleaned === '–') return '';
              return bSoal === 'kategori' ? `${oIdx + 1}. ${cleaned}` : formatOptionString(s, oIdx);
            }).filter(Boolean);
          }
        }

        if (opts.length === 0) {
          opts = ['A. Opsi A', 'B. Opsi B', 'C. Opsi C', 'D. Opsi D', 'E. Opsi E'];
        }

        const qKataKunci = getVal(['kata kunci', 'konsep', 'keywords']) || `${config.mataPelajaran || 'TKA SMA'}, ${qKompetensi}`;
        const qPembahasan = getVal(['pembahasan', 'penjelasan', 'rationale']) || 'Pembahasan telah disesuaikan dengan naskah Gemini AI.';

        const uniqueId = `q-excel-import-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${idx + 1}`;

        parsedQuestions.push({
          id: uniqueId,
          userId: currentUser?.uid,
          noSoal: qNo,
          kisiKisiId: '',
          kompetensi: qKompetensi,
          subKompetensi: qSubKompetensi,
          bentukSoal: bSoal,
          soal: qSoalText,
          stimulus: qStimulus,
          opsi: opts,
          kunciJawaban: qKunci,
          pembahasan: qPembahasan,
          kataKunci: qKataKunci
        });
      });

      return enrichWithMatrix(parsedQuestions);
    } catch (err) {
      console.error("Failed parsing Excel file:", err);
      return [];
    }
  };

  // Handler for uploading file (.xlsx, .xls, .csv, .docx, .txt, .json, .md)
  const handleFileUploadForImportAi = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setImportAiFileName(file.name);
    setImportAiError(null);

    try {
      const lowerName = file.name.toLowerCase();
      if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls') || lowerName.endsWith('.csv')) {
        const arrayBuffer = await file.arrayBuffer();
        const parsed = parseExcelFileToQuestions(arrayBuffer, file.name);
        if (parsed.length > 0) {
          setImportAiParsedQuestions(parsed);
          setImportAiDetectedFormat('File Excel (.xlsx / .csv)');
          setImportAiRawText(`[File Excel Berhasil Dimuat: ${file.name} - ${parsed.length} Butir Soal Terdeteksi]`);
        } else {
          const reader = new FileReader();
          reader.onload = (event) => {
            const text = event.target?.result as string || '';
            setImportAiRawText(text);
            const parsedTxt = parseAiQuestionsText(text);
            setImportAiParsedQuestions(parsedTxt);
          };
          reader.readAsText(file);
        }
      } else if (lowerName.endsWith('.docx') || lowerName.endsWith('.doc')) {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;
        setImportAiRawText(text);
        const parsed = parseAiQuestionsText(text);
        setImportAiParsedQuestions(parsed);
      } else {
        const reader = new FileReader();
        reader.onload = (event) => {
          const text = event.target?.result as string || '';
          setImportAiRawText(text);
          const parsed = parseAiQuestionsText(text);
          setImportAiParsedQuestions(parsed);
        };
        reader.readAsText(file);
      }
    } catch (err: any) {
      console.error(err);
      setImportAiError(`Gagal membaca file: ${err.message || 'Format file tidak didukung.'}`);
    }
  };

  // Save parsed questions to state and Firestore
  const handleSaveImportedAiQuestions = async () => {
    if (importAiParsedQuestions.length === 0) {
      alert("Tidak ada butir soal yang terdeteksi untuk diimpor. Silakan periksa kembali teks atau file yang Anda masukkan!");
      return;
    }

    setIsImportingAi(true);

    try {
      let combinedQuestions: Question[] = [];
      setQuestions(prev => {
        const existingIds = new Set(prev.map(p => p.id));
        const filtered = importAiParsedQuestions.filter(nq => !existingIds.has(nq.id));
        const merged = [...prev, ...filtered].map((q, idx) => ({
          ...q,
          noSoal: idx + 1
        }));
        combinedQuestions = merged;
        setConfig(c => ({ ...c, jumlahSoal: Math.max(c.jumlahSoal || 0, merged.length) }));
        return merged;
      });

      if (currentUser?.uid && combinedQuestions.length > 0) {
        for (const q of combinedQuestions) {
          try {
            await setDoc(doc(db, 'questions', q.id), q);
          } catch (err) {
            console.warn("Failed syncing imported question to firestore:", err);
          }
        }
      }

      alert(`🎉 Berhasil memasukkan ${importAiParsedQuestions.length} butir Soal ke dalam Daftar Butir Soal TKA SMA! Nomor urut otomatis berurutan (1 s.d. ${combinedQuestions.length}).`);
      setIsImportAiModalOpen(false);
      setImportAiRawText('');
      setImportAiParsedQuestions([]);
      setImportAiFileName(null);
      setActiveTab('soal');
    } catch (err: any) {
      console.error(err);
      alert(`Gagal menyimpan soal: ${err.message}`);
    } finally {
      setIsImportingAi(false);
    }
  };

  // Generate Soal for ALL kisi-kisi rows
  const handleGenerateAllQuestions = async () => {
    if (!config.mataPelajaran) {
      alert('Sila pilih Mata Pelajaran terlebih dahulu di Tab 1!');
      return;
    }
    if (kisiList.length === 0) {
      alert('Matriks Asesmen Kisi-Kisi masih kosong. Sila tambahkan atau generate kisi-kisi terlebih dahulu!');
      return;
    }
    
    setIsGeneratingSoal(true);
    let successCount = 0;
    
    // Accumulator for tracking all existing and newly generated questions to prevent duplication across iterations
    const allExistingSoalTexts = [...questions.map(q => q.soal)];
    
    const totalQuestionsTarget = kisiList.reduce((acc, k) => acc + (k.jumlahSoal || 1), 0);
    setSoalProgress({
      active: true,
      type: 'all',
      currentNo: 0,
      totalNo: kisiList.length,
      topic: 'Memulai sinkronisasi seluruh kisi-kisi...',
      countSuccess: 0,
      totalQuestions: totalQuestionsTarget,
      statusText: `Menghubungkan ke ${aiConfig.mode === 'client' ? 'Direct Browser API' : 'Server AI Gemini'}...`
    });
    
    for (let index = 0; index < kisiList.length; index++) {
      const kisi = kisiList[index];
      try {
        const totalToGenerate = kisi.jumlahSoal || 1;
        const chunkSize = 1; // Set to 1 for maximum stability and fast response per API call
        let currentNoSoal = questions.length + successCount + 1;

        setSoalProgress(prev => ({
          ...prev,
          currentNo: index + 1,
          topic: kisi.subElemenMateri || 'Materi Umum',
          statusText: `Menganalisis kisi-kisi No. ${kisi.no}...`
        }));

        for (let i = 0; i < totalToGenerate; i += chunkSize) {
          const countForThisChunk = Math.min(chunkSize, totalToGenerate - i);

          setSoalProgress(prev => ({
            ...prev,
            statusText: `Merancang butir soal #${i + 1} s.d #${i + countForThisChunk} untuk kisi-kisi No. ${kisi.no} via ${aiConfig.mode === 'client' ? 'Direct Browser' : 'Server AI'}...`
          }));

          let data;
          if (aiConfig.mode === 'client') {
            const systemInstruction = `Anda adalah ahli pembuat soal ujian nasional dan TKA (Tes Kemampuan Akademik) SMA di Indonesia. Anda sangat terampil menyusun soal tingkat tinggi (HOTS), bervariasi, mendalam, dan bebas dari bias. Patuhi instruksi bentuk soal dan parameter kognitif secara presisi.`;

            const activeKonteksLokal = (kisi.konteksLokal && kisi.konteksLokal.length > 0) ? kisi.konteksLokal : config.konteksLokal;
            const activeStimulusKonten = (kisi.stimulusKonten && kisi.stimulusKonten.length > 0) ? kisi.stimulusKonten : config.stimulusKonten;
            const activeKualitasChecklist = (kisi.kualitasChecklist && kisi.kualitasChecklist.length > 0) ? kisi.kualitasChecklist : config.kualitasChecklist;

            const konteksStr = activeKonteksLokal.length > 0 
              ? `Integrasikan KONTEKS LOKAL INDONESIA berikut ke dalam stimulus atau soal: ${activeKonteksLokal.join(", ")}.`
              : "";

            const stimulusStr = activeStimulusKonten.length > 0
              ? `Gunakan tipe STIMULUS DAN PENGEMBANGAN KONTEN VISUAL berikut: ${activeStimulusKonten.join(", ")} (utamakan format visual seperti Tabel Data/Matriks Markdown, Diagram Alur/Flowchart ASCII, Grafik Tren Tekstual, Spesifikasi Infografis/Peta, atau Kutipan Dokumen/Dialog Berformat).`
              : "Gunakan stimulus visual/data konkret (Tabel, Flowchart ASCII, Grafik, Spesifikasi Infografis, atau Kutipan Dokumen) yang relevan.";

            const checklistStr = activeKualitasChecklist.length > 0
              ? `Pastikan memenuhi KUALITAS SOAL berikut: ${activeKualitasChecklist.join(", ")}.`
              : "";

            const bentukSoalDesc = 
              kisi.bentukSoal === "pilihan_ganda_sederhana"
                ? `Pilihan ganda sederhana: Hanya ada satu jawaban yang benar. Sediakan pilihan A sampai ${config.jumlahOpsi === 5 ? "E" : "D"}.`
                : kisi.bentukSoal === "mcma"
                ? `Pilihan ganda kompleks model multiple choice multiple answers (MCMA): Ada lebih dari satu jawaban yang benar. Peserta diminta memilih semua jawaban benar. Kunci jawaban harus menyebutkan semua pilihan yang benar (misal: 'A, C'). Sediakan pilihan A sampai ${config.jumlahOpsi === 5 ? "E" : "D"}.`
                : "Pilihan ganda kompleks kategori: Teks 'soal' menggabungkan Stimulus dan Pertanyaan Utama. Field 'opsi' berisi daftar pernyataan berpenomoran angka (1, 2, 3, 4, 5) di mana di akhir kalimat tiap pernyataan diberikan opsi pilihan kategori dalam kurung (contoh: '1. [Teks Pernyataan 1] Pilihan Jawaban : (Tidak Sesuai), (Sesuai)'). Kunci jawaban merinci status per nomor (contoh: '1. Tidak Sesuai, 2. Sesuai, 3. Sesuai, 4. Tidak Sesuai, 5. Sesuai'). Pembahasan merinci poin nomor 1 s.d 5 satu per satu.";

            // Construct constraint for existing questions to avoid duplicates in client-side generator
            let clientExistingQuestionsConstraint = '';
            const activeSlices = allExistingSoalTexts.filter(Boolean).slice(0, 30);
            if (activeSlices.length > 0) {
              clientExistingQuestionsConstraint = `\n\nHINDARI PENGULANGAN SOAL (SANGAT PENTING):\nJangan membuat soal yang sama, memiliki konsep atau contoh kasus/studi yang mirip, atau menggunakan narasi stimulus yang mirip dengan soal-soal berikut:\n${activeSlices.map((text, idx) => `- Soal ${idx + 1}: ${text.substring(0, 150)}...`).join('\n')}\nPastikan butir soal yang Anda hasilkan saat ini benar-benar segar, baru, unik secara naratif, bervariasi, dan tidak mengulangi pertanyaan di atas.`;
            }

            let clientIndonesianLanguageCriteria = '';
            if (config.mataPelajaran && (config.mataPelajaran.toLowerCase().includes('bahasa indonesia') || config.mataPelajaran.toLowerCase().includes('indonesia'))) {
              clientIndonesianLanguageCriteria = `\n\nKAIDAH & KAIDAH MUATAN KHUSUS BAHASA INDONESIA (SANGAT PENTING):\n- Teks yang diujikan harus berupa Teks Informasi (Tunggal/Jamak yang berisi fakta, konsep, prosedur, metakognisi dari berbagai bidang pada skala lokal, nasional, global) ATAU Teks Fiksi (realisme/absurd dengan latar cerita konkret/abstrak, tokoh berkarakter bulat, konflik tunggal/jamak dengan penyelesaian terbuka, alur campuran, dan sudut pandang campuran).\n- Karakteristik Kosakata: Menggunakan kata khusus dan kata umum, kata berimbuhan kompleks, kata abstrak, makna denotatif, istilah teknis, atau konotatif konteks luas.\n- Karakteristik Kalimat: Setiap kalimat di dalam teks stimulus/soal harus berkisar antara 8-12 kata per kalimat, menggunakan kalimat kompleks berbagai pola serta kalimat inversi.\n- Karakteristik Wacana: Menggunakan konjungsi antarparagraf dengan makna 'pertentangan' dan 'sebab akibat', tanda baca pendukung makna yang tepat, dengan panjang teks berkisar antara 250-300 kata (kecuali jika bergenre puisi).`;
            }

            const prompt = `Buatkan tepat sebanyak ${countForThisChunk} butir soal ujian TKA SMA yang berbeda untuk Mata Pelajaran ${config.mataPelajaran}.${clientIndonesianLanguageCriteria}
            
PENTING: Jumlah objek soal yang dihasilkan dalam array JSON HARUS tepat sebanyak ${countForThisChunk} butir soal, tidak kurang dan tidak lebih.
Setiap butir soal harus unik, bervariasi, dan didasarkan pada kisi-kisi berikut.

INFORMASI MATRIKS ASESMEN KISI-KISI:
- No Soal Mulai: ${currentNoSoal}
- Bentuk Soal: ${kisi.bentukSoal} (${bentukSoalDesc})
- Tingkat Kognitif: ${kisi.levelKognitif} (${kisi.levelKognitif === 'level_1' ? 'Pemahaman (Knowing) - Mengenali, mengingat, dan memahami konsep dasar' : kisi.levelKognitif === 'level_2' ? 'Penerapan (Applying) - Menerapkan konsep pada fenomena nyata' : 'Penalaran (Reasoning) - Berpikir kritis dan menalar secara logis'})
- Elemen/Materi: ${kisi.elemenMateri}
- Sub-Elemen/Submateri: ${kisi.subElemenMateri}
- Kompetensi yang Diuji: ${kisi.kompetensi}
- Batasan/Catatan Khusus: ${kisi.batasanCatatan || "Tidak ada"}
- Konteks Nusantara: ${kisi.konteksNusantara || "Tidak ada khusus"}
- Stimulus Tambahan: ${kisi.stimulusTambahan || "Tidak ada khusus"}
- Jenis Soal: ${config.jenisSoal} (Soal Tunggal atau Soal Grup/Terhubung)
${clientExistingQuestionsConstraint}

PANDUAN EKSTRA:
1. ${konteksStr} ${kisi.konteksNusantara ? `Integrasikan juga secara mendalam target Konteks Nusantara berikut ke dalam stimulus atau pokok soal agar bernuansa ke-Indonesia-an yang otentik: "${kisi.konteksNusantara}".` : ""}
2. ${stimulusStr} ${kisi.stimulusTambahan ? `Gunakan secara aktif target Stimulus Tambahan berikut untuk merancang stimulus/skenario pendukung yang kaya dan berbobot: "${kisi.stimulusTambahan}".` : ""}
3. ${checklistStr}
4. Kunci jawaban harus sangat akurat dan pembahasan harus lengkap, ilmiah, edukatif, dan terstruktur dengan rapi agar mudah dipahami siswa SMA. Tambahkan juga field 'kataKunci' yang berisi kata kunci atau konsep penting/topik utama yang digunakan/diuji dalam soal ini.
5. JIKA soal membutuhkan visual pendukung (seperti grafik fungsi, diagram kartesius, bangun geometri, dsb.), Anda disarankan untuk membuat kode SVG inline yang valid (dimulai dengan '<svg' dan ditutup '</svg>' lengkap dengan viewBox, stroke, fill, teks label agar indah dan responsive) ATAU mencantumkan URL gambar Unsplash yang relevan pada field 'gambarUrl'. Jika tidak membutuhkan visual, isi 'gambarUrl' dengan string kosong "".
6. Harap sesuaikan bahasa agar baku, formal, sesuai EBI (Ejaan Bahasa Indonesia), namun mudah dimengerti.
7. Hasilkan tepat ${countForThisChunk} objek soal di dalam array hasil.
8. SANGAT PENTING (MANDATORI): Gabungkan paragraf stimulus/pengantar/studi kasus (bila ada) langsung ke bagian awal field 'soal' (diikuti pertanyaan utama di bawahnya), dan kosongkan field 'stimulus' (isi dengan string kosong ""). Jangan memisahkannya agar struktur soal konsisten dengan prompt.`;

            const soalSchema = {
              type: "ARRAY",
              items: {
                type: "OBJECT",
                properties: {
                  kompetensi: { type: "STRING" },
                  subKompetensi: { type: "STRING" },
                  bentukSoal: { type: "STRING" },
                  stimulus: { type: "STRING", description: "Sengaja dikosongkan karena stimulus digabungkan langsung ke dalam field 'soal' (isi dengan string kosong '')" },
                  soal: { type: "STRING", description: "Teks soal lengkap yang menggabungkan stimulus (paragraf stimulus/pengantar/teks bacaan/studi kasus jika ada) dan pertanyaan/pokok soal utama secara menyatu" },
                  opsi: { 
                    type: "ARRAY", 
                    items: { type: "STRING" }, 
                    description: "Array pilihan jawaban (misal ['A. ...', 'B. ...']) atau daftar pernyataan untuk tipe kategori" 
                  },
                  kunciJawaban: { type: "STRING" },
                  pembahasan: { type: "STRING" },
                  kataKunci: { type: "STRING" },
                  gambarUrl: { type: "STRING" }
                },
                required: ["kompetensi", "subKompetensi", "bentukSoal", "soal", "opsi", "kunciJawaban", "pembahasan", "kataKunci", "gambarUrl"]
              }
            };

            const responseText = await callGeminiDirect(systemInstruction, prompt, soalSchema);
            data = JSON.parse(responseText);
          } else {
            const response = await fetch('/api/generate-soal', {
              method: 'POST',
              headers: { 
                'Content-Type': 'application/json',
                'x-api-key': getCleanApiKey(aiConfig.apiKey)
              },
              body: JSON.stringify({
                kisi,
                count: countForThisChunk,
                mataPelajaran: config.mataPelajaran,
                definisi: config.definisi,
                muatan: config.muatan,
                jumlahOpsi: config.jumlahOpsi,
                jenisSoal: config.jenisSoal,
                konteksLokal: config.konteksLokal,
                stimulusKonten: config.stimulusKonten,
                kualitasChecklist: config.kualitasChecklist,
                noSoalStart: currentNoSoal,
                existingQuestions: allExistingSoalTexts
              })
            });

            if (!response.ok) {
              let errorMsg = `Gagal pada kisi-kisi No. ${kisi.no} (kumpulan ke-${Math.floor(i / chunkSize) + 1})`;
              try {
                const textError = await response.text();
                if (textError.includes('<!doctype') || textError.includes('<html')) {
                  errorMsg = 'Server sedang sibuk atau mengalami timeout (Gateway Timeout).';
                } else {
                  try {
                    const errorData = JSON.parse(textError);
                    errorMsg = errorData.error || errorMsg;
                  } catch {
                    errorMsg = textError || errorMsg;
                  }
                }
              } catch {}
              throw new Error(errorMsg);
            }

            try {
              const responseText = await response.text();
              data = JSON.parse(responseText);
            } catch {
              throw new Error('Respon dari server tidak valid (bukan JSON format).');
            }
          }

          if (Array.isArray(data)) {
            const mapped: Question[] = data.map((q: any, idx: number) => ({
              id: `q-ai-${Date.now()}-${Math.random().toString(36).substring(2, 7)}-${successCount}-${i}-${idx}`,
              noSoal: currentNoSoal + idx,
              kisiKisiId: kisi.id,
              kompetensi: q.kompetensi || kisi.kompetensi,
              subKompetensi: q.subKompetensi || kisi.subElemenMateri,
              bentukSoal: kisi.bentukSoal,
              soal: q.soal,
              stimulus: q.stimulus || '',
              opsi: q.opsi || [],
              kunciJawaban: q.kunciJawaban || 'A',
              pembahasan: q.pembahasan || 'Pembahasan terstruktur.',
              kataKunci: q.kataKunci || '',
              gambarUrl: q.gambarUrl || ''
            }));
            
            // Push newly generated questions' stems to accumulator to prevent any duplicates on subsequent iterations
            allExistingSoalTexts.push(...mapped.map(m => m.soal));

            setQuestions(prev => {
              const existingIds = new Set(prev.map(p => p.id));
              const filtered = mapped.filter(m => !existingIds.has(m.id));
              const combined = [...prev, ...filtered].map((q, idx) => ({
                ...q,
                noSoal: idx + 1
              }));
              setConfig(c => ({ ...c, jumlahSoal: Math.max(c.jumlahSoal || 0, combined.length) }));
              return combined;
            });
            successCount += mapped.length;
            currentNoSoal += mapped.length;

            setSoalProgress(prev => ({
              ...prev,
              countSuccess: successCount,
              statusText: `Berhasil menyusun ${successCount} dari ${totalQuestionsTarget} soal.`
            }));
          } else {
            throw new Error('Respon server tidak berbentuk array.');
          }
        }
      } catch (err: any) {
        console.error('Error generating for a row', err);
        setSoalProgress(prev => ({
          ...prev,
          statusText: `⚠️ Kisi-kisi No. ${kisi.no} gagal diproses: ${err.message || 'Error'}`
        }));
        // Pause briefly so the user can read which part failed/skipped
        await new Promise(resolve => setTimeout(resolve, 2500));
      }
    }
    
    setSoalProgress(prev => ({
      ...prev,
      statusText: 'Semua soal berhasil disusun! Memformat hasil akhir...'
    }));
    await new Promise(resolve => setTimeout(resolve, 800));

    setIsGeneratingSoal(false);
    setSoalProgress(prev => ({ ...prev, active: false }));
    setActiveTab('soal');
    alert(`Penyusunan Massal Selesai! Berhasil menyusun ${successCount} butir soal baru dari seluruh Matriks Kisi-Kisi.`);
  };

  // Fungsi Impor Preset Pusmendik
  const handleImportSinglePreset = async (preset: { elemenMateri: string, subElemenMateri: string, kompetensi: string, batasanCatatan: string }, idx: number) => {
    const presetId = `${preset.subElemenMateri}-${idx}`;
    setImportingPresetIds(prev => ({ ...prev, [presetId]: true }));

    const presetSubjectMapped = selectedPresetSubject === 'PPKN' 
      ? 'Pendidikan Pancasila dan Kewarganegaraan'
      : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
      ? 'Sejarah'
      : selectedPresetSubject === 'Produk Kreatif dan Kewirausahaan'
      ? 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK'
      : selectedPresetSubject;

    setConfig(prev => ({
      ...prev,
      mataPelajaran: presetSubjectMapped
    }));

    const newItem: KisiKisiItem = {
      id: `kisi-pusmendik-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
      userId: currentUser?.uid,
      no: kisiList.length + 1,
      bentukSoal: 'pilihan_ganda_sederhana',
      levelKognitif: 'level_2',
      elemenMateri: preset.elemenMateri,
      subElemenMateri: preset.subElemenMateri,
      kompetensi: preset.kompetensi,
      batasanCatatan: preset.batasanCatatan,
      jumlahSoal: 5
    };

    // Update local React state immediately (offline-first)
    setKisiList(prev => {
      const updated = [...prev, newItem];
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_kisi_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_kisi_local', JSON.stringify(updated));
        localStorage.setItem('tka_kisi_guest', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });

    if (currentUser?.uid) {
      try {
        await setDoc(doc(db, 'kisi_kisi', newItem.id), newItem);
      } catch (err: any) {
        console.warn("Notice syncing single preset to Firestore:", err);
      }
    }

    alert(`Berhasil menambahkan kisi-kisi: "${preset.subElemenMateri}" ke daftar!`);
    setImportingPresetIds(prev => ({ ...prev, [presetId]: false }));
  };

  const handleImportAllPresets = async () => {
    const activePresets = selectedPresetSubject === 'Matematika' 
      ? PUSMENDIK_MATEMATIKA_PRESETS 
      : selectedPresetSubject === 'Bahasa Indonesia'
      ? PUSMENDIK_BAHASA_INDONESIA_PRESETS
      : selectedPresetSubject === 'Bahasa Inggris'
      ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
      : selectedPresetSubject === 'Matematika Tingkat Lanjut'
      ? PUSMENDIK_MATEMATIKA_TL_PRESETS
      : selectedPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
      ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
      : selectedPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
      ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
      : selectedPresetSubject === 'Fisika'
      ? PUSMENDIK_FISIKA_PRESETS
      : selectedPresetSubject === 'Kimia'
      ? PUSMENDIK_KIMIA_PRESETS
      : selectedPresetSubject === 'Biologi'
      ? PUSMENDIK_BIOLOGI_PRESETS
      : selectedPresetSubject === 'PPKN'
      ? PUSMENDIK_PPKN_PRESETS
      : selectedPresetSubject === 'Ekonomi'
      ? PUSMENDIK_EKONOMI_PRESETS
      : selectedPresetSubject === 'Geografi'
      ? PUSMENDIK_GEOGRAFI_PRESETS
      : selectedPresetSubject === 'Sosiologi'
      ? PUSMENDIK_SOSIOLOGI_PRESETS
      : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
      ? PUSMENDIK_SEJARAH_TL_PRESETS
      : selectedPresetSubject === 'Antropologi'
      ? PUSMENDIK_ANTROPOLOGI_PRESETS
      : selectedPresetSubject === 'Bahasa Jepang'
      ? PUSMENDIK_BAHASA_JEPANG_PRESETS
      : PUSMENDIK_PKK_PRESETS;
    const count = activePresets.length;

    const presetSubjectMapped = selectedPresetSubject === 'PPKN' 
      ? 'Pendidikan Pancasila dan Kewarganegaraan'
      : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
      ? 'Sejarah'
      : selectedPresetSubject === 'Produk Kreatif dan Kewirausahaan'
      ? 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK'
      : selectedPresetSubject;

    setShowImportKisiPresetsConfirm({
      count,
      subject: selectedPresetSubject,
      subjectMapped: presetSubjectMapped,
      presets: activePresets
    });
  };

  const executeImportAllKisiPresets = async (presets: any[], subjectMapped: string, subjectName: string) => {
    setShowImportKisiPresetsConfirm(null);

    setConfig(prev => ({
      ...prev,
      mataPelajaran: subjectMapped
    }));

    const startNo = kisiList.length + 1;
    const newItems: KisiKisiItem[] = presets.map((preset, idx) => ({
      id: `kisi-pusmendik-all-${Date.now()}-${idx}-${Math.random().toString(36).substring(2, 6)}`,
      userId: currentUser?.uid,
      no: startNo + idx,
      bentukSoal: 'pilihan_ganda_sederhana',
      levelKognitif: 'level_2',
      elemenMateri: preset.elemenMateri,
      subElemenMateri: preset.subElemenMateri,
      kompetensi: preset.kompetensi,
      batasanCatatan: preset.batasanCatatan,
      jumlahSoal: 5
    }));

    // Update local React state immediately (offline-first)
    setKisiList(prev => {
      const updated = [...prev, ...newItems];
      try {
        if (currentUser?.uid) {
          localStorage.setItem(`tka_kisi_${currentUser.uid}`, JSON.stringify(updated));
        }
        localStorage.setItem('tka_kisi_local', JSON.stringify(updated));
        localStorage.setItem('tka_kisi_guest', JSON.stringify(updated));
      } catch (err) {
        console.warn("Notice saving kisi_kisi to localStorage:", err);
      }
      return updated;
    });

    if (currentUser?.uid) {
      try {
        const batch = writeBatch(db);
        newItems.forEach((item) => {
          batch.set(doc(db, 'kisi_kisi', item.id), item);
        });
        await batch.commit();
      } catch (err: any) {
        console.warn("Notice syncing all kisi presets to Firestore:", err);
      }
    }

    alert(`🎉 Berhasil mengimpor seluruh ${presets.length} matriks standar ${subjectName} ke daftar Matriks Asesmen!`);
  };

  const handleLoadPresetToForm = (preset: { elemenMateri: string, subElemenMateri: string, kompetensi: string, batasanCatatan: string }) => {
    const presetSubjectMapped = selectedPresetSubject === 'PPKN' 
      ? 'Pendidikan Pancasila dan Kewarganegaraan'
      : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
      ? 'Sejarah'
      : selectedPresetSubject === 'Produk Kreatif dan Kewirausahaan'
      ? 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK'
      : selectedPresetSubject;

    setConfig(prev => ({
      ...prev,
      mataPelajaran: presetSubjectMapped
    }));

    setKisiForm({
      bentukSoal: 'pilihan_ganda_sederhana',
      levelKognitif: 'level_2',
      elemenMateri: preset.elemenMateri,
      subElemenMateri: preset.subElemenMateri,
      kompetensi: preset.kompetensi,
      batasanCatatan: preset.batasanCatatan,
      jumlahSoal: 5,
      konteksNusantara: '',
      stimulusTambahan: '',
      konteksLokal: [],
      stimulusKonten: [],
      kualitasChecklist: []
    });
    alert(`Materi "${preset.subElemenMateri}" berhasil dimuat ke Form Tambah/Edit di bawah. Sila sesuaikan sebelum menyimpan.`);
  };

  // Kisi-Kisi Manual Actions
  const handleToggleJenisSoal = async (id: string, currentJenis?: JenisSoal) => {
    const newJenis: JenisSoal = currentJenis === 'grup' ? 'tunggal' : 'grup';
    setKisiList(prev => prev.map(item => item.id === id ? { ...item, jenisSoal: newJenis } : item));
    if (currentUser?.uid) {
      try {
        await setDoc(doc(db, 'kisi_kisi', id), { jenisSoal: newJenis }, { merge: true });
      } catch (err) {
        console.warn("Failed to update jenisSoal in firestore:", err);
      }
    }
  };

  const handleSaveKisiForm = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!kisiForm.elemenMateri || !kisiForm.subElemenMateri || !kisiForm.kompetensi) {
      alert('Sila isi kolom Materi, Sub-materi, dan Kompetensi terlebih dahulu!');
      return;
    }

    if (isEditingKisi && editingKisiId) {
      const updatedItem: KisiKisiItem = {
        id: editingKisiId,
        userId: currentUser?.uid,
        no: kisiList.find(item => item.id === editingKisiId)?.no || 1,
        bentukSoal: kisiForm.bentukSoal as BentukSoal,
        levelKognitif: kisiForm.levelKognitif as LevelKognitif,
        jenisSoal: (kisiForm.jenisSoal as JenisSoal) || 'tunggal',
        elemenMateri: kisiForm.elemenMateri || '',
        subElemenMateri: kisiForm.subElemenMateri || '',
        kompetensi: kisiForm.kompetensi || '',
        batasanCatatan: kisiForm.batasanCatatan || '',
        jumlahSoal: Number(kisiForm.jumlahSoal) || 5,
        konteksNusantara: kisiForm.konteksNusantara || '',
        stimulusTambahan: kisiForm.stimulusTambahan || '',
        konteksLokal: kisiForm.konteksLokal || [],
        stimulusKonten: kisiForm.stimulusKonten || [],
        kualitasChecklist: kisiForm.kualitasChecklist || []
      };

      setKisiList(prev => {
        const updated = prev.map(k => k.id === editingKisiId ? updatedItem : k);
        try {
          if (currentUser?.uid) localStorage.setItem(`tka_kisi_${currentUser.uid}`, JSON.stringify(updated));
          localStorage.setItem('tka_kisi_local', JSON.stringify(updated));
          localStorage.setItem('tka_kisi_guest', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });
      setIsEditingKisi(false);
      setEditingKisiId(null);

      if (currentUser?.uid) {
        try {
          await setDoc(doc(db, 'kisi_kisi', editingKisiId), updatedItem);
        } catch (err: any) {
          console.warn("Notice syncing edited kisi-kisi to Firestore:", err);
        }
      }
      alert(`Berhasil memperbarui data matriks/kisi-kisi No. ${updatedItem.no}!`);
    } else {
      const newItem: KisiKisiItem = {
        id: `kisi-manual-${Date.now()}`,
        userId: currentUser?.uid,
        no: kisiList.length + 1,
        bentukSoal: kisiForm.bentukSoal as BentukSoal,
        levelKognitif: kisiForm.levelKognitif as LevelKognitif,
        jenisSoal: (kisiForm.jenisSoal as JenisSoal) || 'tunggal',
        elemenMateri: kisiForm.elemenMateri || '',
        subElemenMateri: kisiForm.subElemenMateri || '',
        kompetensi: kisiForm.kompetensi || '',
        batasanCatatan: kisiForm.batasanCatatan || '',
        jumlahSoal: Number(kisiForm.jumlahSoal) || 5,
        konteksNusantara: kisiForm.konteksNusantara || '',
        stimulusTambahan: kisiForm.stimulusTambahan || '',
        konteksLokal: kisiForm.konteksLokal || [],
        stimulusKonten: kisiForm.stimulusKonten || [],
        kualitasChecklist: kisiForm.kualitasChecklist || []
      };

      setKisiList(prev => {
        const updated = [...prev, newItem];
        try {
          if (currentUser?.uid) localStorage.setItem(`tka_kisi_${currentUser.uid}`, JSON.stringify(updated));
          localStorage.setItem('tka_kisi_local', JSON.stringify(updated));
          localStorage.setItem('tka_kisi_guest', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });

      if (currentUser?.uid) {
        try {
          await setDoc(doc(db, 'kisi_kisi', newItem.id), newItem);
        } catch (err: any) {
          console.warn("Notice syncing new kisi-kisi to Firestore:", err);
        }
      }
      alert(`Berhasil menambahkan matriks/kisi-kisi baru No. ${newItem.no}!`);
    }

    // Reset Form
    setKisiForm({
      bentukSoal: 'pilihan_ganda_sederhana',
      levelKognitif: 'level_2',
      jenisSoal: 'tunggal',
      elemenMateri: '',
      subElemenMateri: '',
      kompetensi: '',
      batasanCatatan: '',
      jumlahSoal: 5,
      konteksNusantara: '',
      stimulusTambahan: '',
      konteksLokal: [],
      stimulusKonten: [],
      kualitasChecklist: []
    });
  };

  const handleEditKisi = (item: KisiKisiItem) => {
    setKisiForm({
      bentukSoal: item.bentukSoal,
      levelKognitif: item.levelKognitif,
      jenisSoal: item.jenisSoal || 'tunggal',
      elemenMateri: item.elemenMateri,
      subElemenMateri: item.subElemenMateri,
      kompetensi: item.kompetensi,
      batasanCatatan: item.batasanCatatan,
      jumlahSoal: item.jumlahSoal,
      konteksNusantara: item.konteksNusantara || '',
      stimulusTambahan: item.stimulusTambahan || '',
      konteksLokal: item.konteksLokal || [],
      stimulusKonten: item.stimulusKonten || [],
      kualitasChecklist: item.kualitasChecklist || []
    });
    setIsEditingKisi(true);
    setEditingKisiId(item.id);

    // Scroll to edit form section smoothly
    setTimeout(() => {
      const formElem = document.getElementById('kisi-form-section');
      if (formElem) {
        formElem.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }, 50);
  };

  const handleDeleteKisi = async (id: string) => {
    // 1. Immediately update local state so UI reflects deletion instantly
    setKisiList(prev => prev.filter(item => item.id !== id).map((item, idx) => ({ ...item, no: idx + 1 })));
    setQuestions(prev => prev.filter(q => q.kisiKisiId !== id));
    if (editingKisiId === id) {
      setEditingKisiId(null);
      setIsEditingKisi(false);
    }
    setDeletingKisiId(null);

    // 2. Perform background Cloud Firestore cleanup
    if (currentUser?.uid) {
      try {
        await deleteDoc(doc(db, 'kisi_kisi', id)).catch(err => {
          console.warn("Cloud delete kisi_kisi notice:", err);
        });
        await deleteDoc(doc(db, 'materials', id)).catch(err => {
          console.warn("Cloud delete materials notice:", err);
        });
        
        const qQuestions = query(collection(db, 'questions'), where('kisiKisiId', '==', id));
        const qSnapshot = await getDocs(qQuestions);
        if (!qSnapshot.empty) {
          const batch = writeBatch(db);
          qSnapshot.forEach((docSnap) => {
            batch.delete(docSnap.ref);
          });
          await batch.commit().catch(e => console.warn("Batch delete associated questions notice:", e));
        }
      } catch (err: any) {
        console.warn("Background cloud deletion notice:", err);
      }
    }
  };

  const handleDeleteUnusedKisi = async () => {
    if (kisiList.length === 0) {
      alert('Tabel Matriks Kisi-Kisi sudah dalam keadaan kosong.');
      return;
    }

    // Collect all referenced kisi-kisi IDs and row numbers from current questions
    const usedKisiKeys = new Set<string>();
    questions.forEach(q => {
      if (q.kisiKisiId) usedKisiKeys.add(String(q.kisiKisiId));
      if (q.noSoal) usedKisiKeys.add(String(q.noSoal));
    });

    const unusedKisi = kisiList.filter(item => {
      const isReferenced = usedKisiKeys.has(String(item.id)) || usedKisiKeys.has(String(item.no));
      const hasContent = Boolean(
        (item.elemenMateri && item.elemenMateri.trim()) ||
        (item.subElemenMateri && item.subElemenMateri.trim()) ||
        (item.indikatorSoal && item.indikatorSoal.trim()) ||
        (item.kompetensi && item.kompetensi.trim())
      );
      // Truly empty if NOT referenced by any question AND contains no filled text content
      return !isReferenced && !hasContent;
    });

    if (unusedKisi.length === 0) {
      alert('Tidak ditemukan baris Kisi-Kisi kosong. Semua baris yang ada sudah terhubung dengan butir soal dan memiliki isi materi.');
      return;
    }

    if (!confirm(`Apakah Anda yakin ingin menghapus ${unusedKisi.length} baris Kisi-Kisi kosong / tidak terpakai?`)) {
      return;
    }

    try {
      const unusedIds = new Set(unusedKisi.map(k => k.id));
      setKisiList(prev => prev.filter(k => !unusedIds.has(k.id)));

      if (currentUser?.uid) {
        const docsToDelete = unusedKisi.map(k => k.id);
        const chunkSize = 400;
        for (let i = 0; i < docsToDelete.length; i += chunkSize) {
          const chunk = docsToDelete.slice(i, i + chunkSize);
          const batch = writeBatch(db);
          chunk.forEach((id) => {
            batch.delete(doc(db, 'kisi_kisi', id));
            batch.delete(doc(db, 'materials', id));
          });
          await batch.commit().catch(e => console.warn("Batch delete unused kisi warning:", e));
        }
      }
      alert(`Berhasil menghapus ${unusedKisi.length} baris Kisi-Kisi kosong.`);
    } catch (err: any) {
      console.error("Gagal menghapus unused kisi-kisi:", err);
      const unusedIds = new Set(unusedKisi.map(k => k.id));
      setKisiList(prev => prev.filter(k => !unusedIds.has(k.id)));
      alert(`Berhasil membersihkan ${unusedKisi.length} baris Kisi-Kisi kosong dari tampilan.`);
    }
  };

  const handleDeleteAllKisi = async () => {
    if (kisiList.length === 0) {
      alert("Tidak ada baris Matriks Kisi-Kisi untuk dihapus.");
      return;
    }

    if (!confirm(`Apakah Anda yakin ingin menghapus SEMUA (${kisiList.length}) baris Matriks Kisi-Kisi? (Catatan: Ini juga akan membersihkan materi rujukan yang terkait)`)) {
      return;
    }

    try {
      const previousKisi = [...kisiList];
      setKisiList([]);

      if (currentUser?.uid) {
        const qKisi = query(collection(db, 'kisi_kisi'), where('userId', '==', currentUser.uid));
        const qSnapshot = await getDocs(qKisi);
        if (!qSnapshot.empty) {
          const docs = qSnapshot.docs;
          const chunkSize = 400;
          for (let i = 0; i < docs.length; i += chunkSize) {
            const chunk = docs.slice(i, i + chunkSize);
            const batch = writeBatch(db);
            chunk.forEach((docSnap) => {
              batch.delete(docSnap.ref);
              batch.delete(doc(db, 'materials', docSnap.id));
            });
            await batch.commit().catch(err => console.warn("Batch delete kisi warning:", err));
          }
        }

        const extraChunkSize = 400;
        for (let i = 0; i < previousKisi.length; i += extraChunkSize) {
          const chunk = previousKisi.slice(i, i + extraChunkSize);
          const batch = writeBatch(db);
          chunk.forEach((item) => {
            batch.delete(doc(db, 'kisi_kisi', item.id));
            batch.delete(doc(db, 'materials', item.id));
          });
          await batch.commit().catch(err => console.warn("Batch delete extra kisi warning:", err));
        }
      }

      alert("Semua baris Matriks Kisi-Kisi berhasil dihapus.");
    } catch (err: any) {
      console.error("Gagal menghapus semua kisi-kisi:", err);
      alert("Semua baris Matriks Kisi-Kisi telah dibersihkan.");
    }
  };

  // Questions Manual Actions
  const handleSaveQuestionForm = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!questionForm.soal || !questionForm.kunciJawaban) {
      alert('Teks Soal dan Kunci Jawaban wajib diisi!');
      return;
    }

    // Clean options (remove empty strings and format)
    const activeOptions = (questionForm.opsi || [])
      .filter(o => o.trim() !== '')
      .map((opt, i) => formatOptionString(opt, i));
    const cleanFormSoal = cleanSoalText(questionForm.soal || '');
    setIsSavingQuestion(true);

    try {
      if (isEditingQuestion && editingQuestionId) {
        const existingQ = questions.find(q => q.id === editingQuestionId);
        const updatedQ: Question = {
          id: editingQuestionId,
          userId: currentUser?.uid || existingQ?.userId,
          noSoal: existingQ?.noSoal || 1,
          kisiKisiId: questionForm.kisiKisiId || '',
          kompetensi: questionForm.kompetensi || '',
          subKompetensi: questionForm.subKompetensi || '',
          bentukSoal: questionForm.bentukSoal as BentukSoal,
          soal: cleanFormSoal,
          stimulus: questionForm.stimulus || '',
          opsi: activeOptions,
          kunciJawaban: questionForm.kunciJawaban || '',
          pembahasan: questionForm.pembahasan || '',
          kataKunci: questionForm.kataKunci || '',
          gambarUrl: questionForm.gambarUrl || '',
          gambarCaption: questionForm.gambarCaption || '',
          gambarPosisi: questionForm.gambarPosisi || 'center',
          gambarUkuran: questionForm.gambarUkuran || 'medium'
        };

        if (currentUser?.uid) {
          try {
            await setDoc(doc(db, 'questions', editingQuestionId), updatedQ);
          } catch (e) {
            console.warn("Gagal update Firestore, memperbarui state lokal:", e);
          }
        }
        setQuestions(prev => prev.map(q => q.id === editingQuestionId ? updatedQ : q));
        setIsEditingQuestion(false);
        setEditingQuestionId(null);
        alert(`Berhasil memperbarui butir soal No. ${updatedQ.noSoal}!`);
      } else {
        const newQ: Question = {
          id: `q-manual-${Date.now()}`,
          userId: currentUser?.uid,
          noSoal: (questions.reduce((max, q) => Math.max(max, Number(q.noSoal) || 0), 0)) + 1,
          kisiKisiId: questionForm.kisiKisiId || '',
          kompetensi: questionForm.kompetensi || 'Kompetensi Umum',
          subKompetensi: questionForm.subKompetensi || 'Sub-Materi',
          bentukSoal: questionForm.bentukSoal as BentukSoal,
          soal: cleanFormSoal,
          stimulus: questionForm.stimulus || '',
          opsi: activeOptions,
          kunciJawaban: questionForm.kunciJawaban || '',
          pembahasan: questionForm.pembahasan || 'Pembahasan terstruktur.',
          kataKunci: questionForm.kataKunci || '',
          gambarUrl: questionForm.gambarUrl || '',
          gambarCaption: questionForm.gambarCaption || '',
          gambarPosisi: questionForm.gambarPosisi || 'center',
          gambarUkuran: questionForm.gambarUkuran || 'medium'
        };

        if (currentUser?.uid) {
          try {
            await setDoc(doc(db, 'questions', newQ.id), newQ);
          } catch (e) {
            console.warn("Gagal simpan ke Firestore, menambahkan ke state lokal:", e);
          }
        }
        setQuestions(prev => {
          const updated = [...prev, newQ].map((q, idx) => ({ ...q, noSoal: idx + 1 }));
          setConfig(c => ({ ...c, jumlahSoal: Math.max(c.jumlahSoal || 0, updated.length) }));
          return updated;
        });
        alert(`Berhasil menyimpan butir soal baru No. ${questions.length + 1}!`);
      }

      // Reset Form
      setQuestionForm({
        kisiKisiId: '',
        kompetensi: '',
        subKompetensi: '',
        bentukSoal: 'pilihan_ganda_sederhana',
        soal: '',
        stimulus: '',
        opsi: ['', '', '', '', ''],
        kunciJawaban: '',
        pembahasan: '',
        kataKunci: '',
        gambarUrl: '',
        gambarCaption: '',
        gambarPosisi: 'center',
        gambarUkuran: 'medium'
      });
      setImageCompressReport(null);
      setIsEditingQuestion(false);
      setEditingQuestionId(null);
    } catch (err: any) {
      console.error("Gagal menyimpan soal:", err);
      alert(`Gagal menyimpan butir soal: ${err.message || err}`);
    } finally {
      setIsSavingQuestion(false);
    }
  };

  const handleResequenceQuestionNumbers = async () => {
    if (questions.length === 0) {
      alert("Belum ada butir soal untuk diurutkan.");
      return;
    }
    const resequenced = questions.map((q, idx) => ({ ...q, noSoal: idx + 1 }));
    setQuestions(resequenced);
    setConfig(c => ({ ...c, jumlahSoal: Math.max(c.jumlahSoal || 0, resequenced.length) }));
    if (currentUser?.uid) {
      for (const q of resequenced) {
        await setDoc(doc(db, 'questions', q.id), q).catch(() => {});
      }
    }
    alert(`🎉 Berhasil merapikan nomor urut ${resequenced.length} butir soal menjadi 1 s.d. ${resequenced.length}!`);
  };

  const handleEditQuestion = (q: Question) => {
    // Clean option letters if present to avoid duplicate "A. A. Jawaban"
    const cleanedOptions = (q.opsi || []).map(opt => cleanOptionText(opt));
    const padOptions = [...cleanedOptions, '', '', '', ''].slice(0, 5);

    setQuestionForm({
      kisiKisiId: q.kisiKisiId || '',
      kompetensi: q.kompetensi || '',
      subKompetensi: q.subKompetensi || '',
      bentukSoal: q.bentukSoal || 'pilihan_ganda_sederhana',
      soal: q.soal || '',
      stimulus: q.stimulus || '',
      opsi: padOptions,
      kunciJawaban: q.kunciJawaban || '',
      pembahasan: q.pembahasan || '',
      kataKunci: q.kataKunci || '',
      gambarUrl: q.gambarUrl || '',
      gambarCaption: q.gambarCaption || '',
      gambarPosisi: q.gambarPosisi || 'center',
      gambarUkuran: q.gambarUkuran || 'medium'
    });
    setImageCompressReport(null);
    setIsEditingQuestion(true);
    setEditingQuestionId(q.id);

    setTimeout(() => {
      const element = document.getElementById(`soal-card-${q.id}`);
      if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }, 100);
  };

  const handleDeleteQuestion = async (id: string) => {
    try {
      if (currentUser?.uid) {
        await deleteDoc(doc(db, 'questions', id));
      }
      setQuestions(prev => {
        const filtered = prev.filter(q => q.id !== id);
        const resequenced = filtered.map((q, idx) => ({ ...q, noSoal: idx + 1 }));
        if (currentUser?.uid) {
          resequenced.forEach(rq => setDoc(doc(db, 'questions', rq.id), rq).catch(() => {}));
        }
        return resequenced;
      });
    } catch (err: any) {
      console.error("Gagal menghapus soal:", err);
      setQuestions(prev => prev.filter(q => q.id !== id).map((q, idx) => ({ ...q, noSoal: idx + 1 })));
    }
  };

  const handleDeleteAllQuestions = async () => {
    if (questions.length === 0) {
      alert("Tidak ada soal untuk dihapus.");
      return;
    }

    try {
      const previousQuestions = [...questions];
      setQuestions([]);
      setShowDeleteAllConfirm(false);
      setIsEditingQuestion(false);
      setEditingQuestionId(null);

      if (currentUser?.uid) {
        const qQuestions = query(collection(db, 'questions'), where('userId', '==', currentUser.uid));
        const qSnapshot = await getDocs(qQuestions);
        if (!qSnapshot.empty) {
          const docs = qSnapshot.docs;
          const chunkSize = 400;
          for (let i = 0; i < docs.length; i += chunkSize) {
            const chunk = docs.slice(i, i + chunkSize);
            const batch = writeBatch(db);
            chunk.forEach((docSnap) => {
              batch.delete(docSnap.ref);
            });
            await batch.commit().catch(err => console.warn("Batch commit delete questions warning:", err));
          }
        }

        const extraChunkSize = 400;
        for (let i = 0; i < previousQuestions.length; i += extraChunkSize) {
          const chunk = previousQuestions.slice(i, i + extraChunkSize);
          const batch = writeBatch(db);
          chunk.forEach((q) => {
            batch.delete(doc(db, 'questions', q.id));
          });
          await batch.commit().catch(err => console.warn("Batch commit extra delete questions warning:", err));
        }
      }

      alert("Semua butir soal TKA SMA berhasil dihapus.");
    } catch (err: any) {
      console.error("Gagal menghapus semua soal:", err);
      alert("Semua butir soal di layar telah dibersihkan.");
    }
  };

  const handleOpsiChange = (index: number, val: string) => {
    setQuestionForm(prev => {
      const currentOpsi = [...(prev.opsi || ['', '', '', '', ''])];
      currentOpsi[index] = val;
      return { ...prev, opsi: currentOpsi };
    });
  };

  const handlePrint = () => {
    if (questions.length === 0) {
      alert("Belum ada butir soal yang tersusun! Silakan buat Kisi-Kisi terlebih dahulu, lalu susun/buat butir soal sebelum mencetak.");
      return;
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      alert("Popup blocker aktif. Harap izinkan popup untuk melakukan pencetakan.");
      return;
    }

    // Determine font size
    let fontSizeVal = '11pt';
    if (printConfig.fontSize === 'text-xs') fontSizeVal = '10pt';
    if (printConfig.fontSize === 'text-base') fontSizeVal = '12pt';

    // Generate questions html
    const questionsHtml = questions.map((q) => {
      // Parse options
      const isKategori = isKategoriSoal(q);
      let optionsHtml = '';

      if (isKategori) {
        const { cat1, cat2 } = getPgkCategories(q.soal, q.kunciJawaban);
        const validOpts = (q.opsi || []).filter(opt => {
          const txt = cleanOptionText(opt).trim();
          return txt !== '' && txt !== '-' && txt !== '–' && txt !== '—';
        });

        const rowsHtml = validOpts.map((opt, i) => {
          const optLetter = String.fromCharCode(65 + i);
          const optText = cleanOptionText(opt);
          const catSel = getPgkCategoryIndex(q.kunciJawaban, i, optLetter, cat1, cat2);
          const isCat1 = printConfig.showAnswerKey && catSel === 1;
          const isCat2 = printConfig.showAnswerKey && catSel === 2;

          return `
            <tr>
              <td style="border: 1px solid #111827; padding: 6px; text-align: center; font-weight: bold; width: 35px; font-size: 10pt;">${i + 1}.</td>
              <td style="border: 1px solid #111827; padding: 6px; font-size: 10pt; line-height: 1.4;">${optText}</td>
              <td style="border: 1px solid #111827; padding: 6px; text-align: center; width: 110px; font-size: 10pt; ${isCat1 ? 'background-color: #dcfce7; font-weight: bold;' : ''}">
                <span style="display: inline-block; width: 16px; height: 16px; border-radius: 50%; border: 1.5px solid #111827; ${isCat1 ? 'background-color: #059669;' : 'background-color: #ffffff;'}"></span>
              </td>
              <td style="border: 1px solid #111827; padding: 6px; text-align: center; width: 110px; font-size: 10pt; ${isCat2 ? 'background-color: #dcfce7; font-weight: bold;' : ''}">
                <span style="display: inline-block; width: 16px; height: 16px; border-radius: 50%; border: 1.5px solid #111827; ${isCat2 ? 'background-color: #059669;' : 'background-color: #ffffff;'}"></span>
              </td>
            </tr>
          `;
        }).join('');

        optionsHtml = `
          <div style="margin: 10px 0;">
            <table style="width: 100%; border-collapse: collapse; border: 1.5px solid #111827; font-size: 10pt;">
              <thead>
                <tr style="background-color: #f3f4f6;">
                  <th style="border: 1px solid #111827; padding: 6px; text-align: center; width: 35px; font-weight: bold;">#</th>
                  <th style="border: 1px solid #111827; padding: 6px; text-align: left; font-weight: bold;">Pernyataan</th>
                  <th style="border: 1px solid #111827; padding: 6px; text-align: center; width: 110px; background-color: #e5e7eb; font-weight: bold;">${cat1}</th>
                  <th style="border: 1px solid #111827; padding: 6px; text-align: center; width: 110px; background-color: #e5e7eb; font-weight: bold;">${cat2}</th>
                </tr>
              </thead>
              <tbody>
                ${rowsHtml}
              </tbody>
            </table>
          </div>
        `;
      } else {
        optionsHtml = `
          <div class="options-container ${printConfig.layoutColumns === '2' ? 'single-col' : 'grid-cols-2'}">
            ${(q.opsi || []).map((opt, i) => {
              let optLetter = '';
              let optText = opt;
              if (opt.trim().match(/^[A-E]\s*[\.\)]/i)) {
                optLetter = opt.trim().substring(0, 1).toUpperCase();
                const sepIdx = opt.indexOf('.') !== -1 ? opt.indexOf('.') : opt.indexOf(')');
                optText = opt.substring(sepIdx + 1).trim();
              } else {
                optLetter = String.fromCharCode(65 + i);
              }

              const isCorrectOption = q.kunciJawaban.trim().toUpperCase().includes(optLetter) && printConfig.showAnswerKey;

              return `
                <div class="option-item ${isCorrectOption ? 'correct-option' : ''}">
                  <span class="option-letter">${optLetter}</span>
                  <span class="option-text">${optText}</span>
                </div>
              `;
            }).join('')}
          </div>
        `;
      }

      // Parse illustration
      let illustrationHtml = '';
      if (q.gambarUrl && q.gambarUrl.trim() !== '' && printConfig.showIllustration) {
        if (q.gambarUrl.trim().toLowerCase().startsWith('<svg')) {
          illustrationHtml = `
            <div class="illustration-container">
              ${q.gambarUrl}
            </div>
          `;
        } else {
          illustrationHtml = `
            <div class="illustration-container">
              <img src="${q.gambarUrl}" alt="Ilustrasi" />
            </div>
          `;
        }
      }

      // Parse competency tag
      let competencyTagHtml = '';
      if (printConfig.showCompetencyTag) {
        competencyTagHtml = `
          <div class="competency-tag">
            <strong>No Soal:</strong> ${q.noSoal} | 
            <strong>Bentuk:</strong> ${getBentukSoalLabel(q.bentukSoal)} | 
            <strong>Kompetensi:</strong> ${q.kompetensi} | 
            <strong>Sub Kompetensi:</strong> ${q.subKompetensi}
          </div>
        `;
      }

      // Parse answer key & pembahasan
      let answerKeyHtml = '';
      if (printConfig.showAnswerKey) {
        answerKeyHtml = `
          <div class="answer-key-box">
            <div><strong>Kunci Jawaban:</strong> <span class="key-badge">${q.kunciJawaban}</span></div>
            ${q.kataKunci ? `<div><strong>Materi / Konsep:</strong> ${q.kataKunci}</div>` : ''}
            <div style="margin-top: 4px;"><strong>Pembahasan:</strong> ${q.pembahasan || '-'}</div>
          </div>
        `;
      }

      // Combined Stimulus & Question Statement
      const combinedQuestionTextHtml = `
        ${q.stimulus && printConfig.showStimulus ? `
          <div class="stimulus-combined-text" style="font-weight: normal; font-style: italic; margin-bottom: 8px; text-align: justify; line-height: 1.4;">
            ${q.stimulus}
          </div>
        ` : ''}
        <div class="question-statement-text" style="font-weight: bold;">
          ${q.soal}
        </div>
      `;

      return `
        <div class="question-item">
          ${competencyTagHtml}
          <div class="question-body">
            ${!printConfig.showCompetencyTag ? `<span class="question-number">${q.noSoal}.</span>` : ''}
            <div class="question-content">
              ${illustrationHtml}
              <div class="question-text">
                ${combinedQuestionTextHtml}
              </div>
              ${optionsHtml}
              ${answerKeyHtml}
            </div>
          </div>
        </div>
      `;
    }).join('');

    // Generate school logos
    const leftLogoHtml = printConfig.schoolLogo 
      ? `<img src="${printConfig.schoolLogo}" class="kop-logo" />` 
      : `<div class="kop-logo-placeholder">LOGO</div>`;
    const rightLogoHtml = printConfig.schoolLogoRight 
      ? `<img src="${printConfig.schoolLogoRight}" class="kop-logo" />` 
      : `<div class="kop-logo-placeholder">SMA</div>`;

    // Complete HTML structure
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Ujian_${config.mataPelajaran.replace(/\s+/g, '_')}</title>
        <style>
          @media print {
            body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .no-print { display: none !important; }
          }
          @page {
            size: ${printConfig.pageSize === 'F4' ? '215mm 330mm' : 'A4'};
            margin: 1.5cm 1.5cm 1.5cm 1.5cm;
          }
          body {
            font-family: 'Times New Roman', Times, serif;
            color: #000000;
            line-height: 1.4;
            font-size: ${fontSizeVal};
            background: white;
            margin: 0;
            padding: 0;
          }
          .print-btn-container {
            position: fixed;
            bottom: 24px;
            right: 24px;
            background-color: #2563eb;
            color: white;
            padding: 12px 24px;
            border-radius: 9999px;
            font-weight: bold;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            cursor: pointer;
            z-index: 9999;
            font-family: 'Times New Roman', Times, serif;
            border: none;
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
            transition: background-color 0.2s;
          }
          .print-btn-container:hover {
            background-color: #1d4ed8;
          }
          
          /* Kop Surat */
          .header-kop {
            border-bottom: 3px double #000000;
            padding-bottom: 8px;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 16px;
          }
          .kop-logo {
            width: 60px;
            height: 60px;
            object-fit: contain;
          }
          .kop-logo-placeholder {
            width: 60px;
            height: 60px;
            border-radius: 50%;
            border: 1px solid #000000;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            font-weight: bold;
          }
          .kop-text {
            flex: 1;
            text-align: center;
          }
          .kop-dept {
            font-size: 9pt;
            font-weight: bold;
            text-transform: uppercase;
            margin: 0 0 2px 0;
          }
          .kop-school {
            font-size: 13pt;
            font-weight: bold;
            text-transform: uppercase;
            margin: 0 0 2px 0;
          }
          .kop-address {
            font-size: 8pt;
            font-style: italic;
            margin: 0 0 4px 0;
          }
          .kop-info {
            font-size: 8.5pt;
            font-weight: bold;
            margin: 0;
            border-top: 1px solid #000000;
            padding-top: 2px;
            display: inline-block;
            word-spacing: 2px;
          }

          /* Simple Title if No Kop */
          .simple-title {
            border-bottom: 2px solid #000000;
            padding-bottom: 8px;
            text-align: center;
            margin-bottom: 16px;
          }
          .simple-title h2 {
            font-size: 14pt;
            margin: 0;
            font-weight: bold;
          }
          .simple-title p {
            font-size: 11pt;
            margin: 4px 0 0 0;
          }

          /* Exam Metadata */
          .exam-meta-title {
            text-align: center;
            margin-bottom: 16px;
          }
          .exam-meta-title h2 {
            font-size: 11pt;
            font-weight: bold;
            margin: 0 0 2px 0;
          }
          .exam-meta-title h1 {
            font-size: 12pt;
            font-weight: bold;
            margin: 0;
          }

          /* Student Fields Table */
          .student-fields {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 16px;
            font-size: 9.5pt;
          }
          .student-fields td {
            padding: 6px 10px;
            border: 1px solid #000000;
            font-weight: bold;
          }
          .dotted-line {
            display: inline-block;
            width: 80%;
            border-bottom: 1px dotted #000000;
            height: 12px;
          }

          /* Instructions */
          .instructions-box {
            border-left: 3px solid #000000;
            padding-left: 10px;
            margin-bottom: 20px;
            font-size: 9.5pt;
            font-style: italic;
          }

          /* Layout Columns */
          .questions-container {
            width: 100%;
          }
          .layout-columns-2 {
            -webkit-column-count: 2;
            -moz-column-count: 2;
            column-count: 2;
            -webkit-column-gap: 24px;
            -moz-column-gap: 24px;
            column-gap: 24px;
          }
          
          /* Question items styling */
          .question-item {
            display: block;
            width: 100%;
            margin-bottom: 16px;
            border-bottom: 1px solid #f3f4f6;
            padding-bottom: 12px;
            page-break-inside: auto; /* Allow natural page-break flow to prevent separating header and questions */
            break-inside: auto;
          }
          /* If two-column is active, keep questions self-contained within columns */
          .layout-columns-2 .question-item {
            display: inline-block;
            -webkit-column-break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid-column;
            break-inside: avoid;
          }
          .question-item:last-child {
            border-bottom: none;
          }
          .competency-tag {
            background-color: #f3f4f6;
            border: 1px solid #e5e7eb;
            padding: 4px 8px;
            font-size: 8pt;
            border-radius: 4px;
            margin-bottom: 8px;
          }
          .question-body {
            display: flex;
            align-items: start;
            gap: 8px;
          }
          .question-number {
            font-weight: bold;
            min-width: 20px;
          }
          .question-content {
            flex: 1;
          }
          .question-text {
            font-weight: bold;
            margin-bottom: 10px;
            text-align: justify;
          }
          
           /* Stimulus */
          .stimulus-box {
            background-color: #f9fafb;
            border-left: 3px solid #4f46e5;
            padding: 8px 12px;
            margin-bottom: 10px;
            font-size: 9pt;
            font-style: italic;
            border-radius: 0 4px 4px 0;
            text-align: justify;
            -webkit-column-break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          
          /* Illustration */
          .illustration-container {
            margin: 10px 0;
            text-align: center;
            -webkit-column-break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          .illustration-container img {
            max-height: 150px;
            max-width: 100%;
            object-fit: contain;
          }
          .illustration-container svg {
            max-width: 100%;
            height: auto;
          }

          /* Options */
          .options-container {
            display: grid;
            gap: 6px;
            margin-bottom: 8px;
            -webkit-column-break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          .options-container.grid-cols-2 {
            grid-template-columns: 1fr 1fr;
          }
          @media (max-width: 600px) {
            .options-container.grid-cols-2 {
              grid-template-columns: 1fr;
            }
          }
          .options-container.single-col {
            grid-template-columns: 1fr;
          }
          .option-item {
            display: flex;
            align-items: start;
            gap: 8px;
            padding: 4px 6px;
            border-radius: 4px;
            border: 1px solid transparent;
            font-size: 9.5pt;
          }
          .option-letter {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            border: 1px solid #000000;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 8pt;
            flex-shrink: 0;
            margin-top: 1px;
          }
          .option-text {
            flex: 1;
          }
          .correct-option {
            background-color: #f0fdf4;
            border-color: #bbf7d0;
            font-weight: bold;
          }
          .correct-option .option-letter {
            background-color: #22c55e;
            color: white;
            border-color: #22c55e;
          }

          /* Answer Key Box */
          .answer-key-box {
            background-color: #f0fdf4;
            border: 1px dashed #86efac;
            padding: 8px 12px;
            margin-top: 10px;
            font-size: 8.5pt;
            border-radius: 6px;
          }
          .key-badge {
            background-color: #22c55e;
            color: white;
            font-weight: bold;
            padding: 1px 6px;
            border-radius: 4px;
            font-family: monospace;
          }
        </style>
      </head>
      <body>
        <button class="print-btn-container no-print" onclick="window.print()">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"></path><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
          Cetak Dokumen / Simpan PDF
        </button>

        ${printConfig.showHeader ? `
          <div class="header-kop">
            ${leftLogoHtml}
            <div class="kop-text">
              ${(printConfig.kopDepartment || 'KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI')
                .split('\n')
                .map(line => `<p class="kop-dept">${line.trim()}</p>`)
                .join('')
              }
              <h1 class="kop-school">${printConfig.schoolName}</h1>
              <p class="kop-address">${printConfig.schoolAddress}</p>
              <p class="kop-info">TAHUN PELAJARAN: ${printConfig.academicYear} | SEMESTER: ${printConfig.semester.toUpperCase()}</p>
            </div>
            ${rightLogoHtml}
          </div>
        ` : `
          <div class="simple-title">
            <h2>LEMBAR SOAL UJIAN TKA SMA</h2>
            <p><strong>Mata Pelajaran:</strong> ${printConfig.subjectName || config.mataPelajaran || 'TES KEMAMPUAN AKADEMIK'} | <strong>Muatan:</strong> ${config.muatan || 'SMA'}</p>
          </div>
        `}

        ${printConfig.showHeader ? `
          <div class="exam-meta-title">
            <h2>${printConfig.examName}</h2>
            <h1>MATA PELAJARAN: ${printConfig.subjectName || config.mataPelajaran || 'TES KEMAMPUAN AKADEMIK'}</h1>
            <div style="font-size: 8.5pt; margin-top: 4px; font-weight: bold;">
              <span>Fase/Muatan: ${config.muatan || 'SMA'}</span> &nbsp;|&nbsp; 
              <span>Alokasi Waktu: ${printConfig.timeAllocation}</span>
            </div>
          </div>
        ` : ''}

        ${printConfig.showStudentFields ? `
          <table class="student-fields">
            <tr>
              <td style="width: 50%;">NAMA LENGKAP: <span class="dotted-line"></span></td>
              <td style="width: 50%;">KELAS / JURUSAN: <span class="dotted-line"></span></td>
            </tr>
            <tr>
              <td style="width: 50%;">NOMOR PESERTA: <span class="dotted-line"></span></td>
              <td style="width: 50%;">HARI / TANGGAL: ${new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</td>
            </tr>
          </table>
        ` : ''}

        ${printConfig.instructionText ? `
          <div class="instructions-box">
            <strong>PETUNJUK PENGERJAAN:</strong> ${printConfig.instructionText}
          </div>
        ` : ''}

        <div class="questions-container ${printConfig.layoutColumns === '2' ? 'layout-columns-2' : ''}">
          ${questionsHtml}
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 500);
          }
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();
  };

  if (isAuthLoading) {
    return (
      <div className="min-h-screen bg-slate-950 flex flex-col items-center justify-center text-white">
        <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-indigo-500 mb-4"></div>
        <p className="text-slate-400 font-medium animate-pulse">Memuat sistem autentikasi TKA SMA...</p>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <LoginScreen 
        onLoginSuccess={(role, name, userObj) => {
          const sessionUser = userObj || {
            id: role === 'admin' ? 'usr_admin' : 'usr_demo',
            email: role === 'admin' ? 'admin@tka.com' : 'user@tka.com',
            name: name || (role === 'admin' ? 'Admin TKA SMA' : 'Guru Sosiologi'),
            role: role,
            mataPelajaran: 'Sosiologi'
          };

          try {
            localStorage.setItem('tka_active_session', JSON.stringify(sessionUser));
          } catch (e) {
            console.warn("Gagal menyimpan sesi aktif:", e);
          }

          setCurrentUser({
            uid: sessionUser.id || 'usr_demo',
            email: sessionUser.email,
            displayName: sessionUser.name || name
          });
          setUserRole(role);
          setUserName(name || sessionUser.name);
          if (sessionUser.mataPelajaran) {
            setConfig(prev => ({ ...prev, mataPelajaran: sessionUser.mataPelajaran }));
          }
        }} 
      />
    );
  }

  return (
    <div id="app-root" className="min-h-screen bg-slate-50 text-slate-800 font-sans flex flex-col antialiased relative">
      {/* Animated API Key Rotation / Quota Toast Banner Overlay */}
      <AnimatePresence>
        {apiKeyToast && (
          <motion.div
            key={apiKeyToast.id}
            initial={{ opacity: 0, y: -60, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -30, scale: 0.95 }}
            transition={{ type: 'spring', stiffness: 400, damping: 25 }}
            className={`fixed top-4 left-1/2 -translate-x-1/2 z-[9999] w-[92%] max-w-lg p-4 rounded-2xl shadow-2xl border backdrop-blur-md text-left flex items-start gap-3.5 ${
              apiKeyToast.type === 'quota_exhausted'
                ? 'bg-rose-950/95 border-rose-500/80 text-rose-50 shadow-rose-950/50'
                : apiKeyToast.type === 'rotation'
                ? 'bg-slate-900/95 border-indigo-500/80 text-indigo-50 shadow-indigo-950/50'
                : 'bg-emerald-950/95 border-emerald-500/80 text-emerald-50 shadow-emerald-950/50'
            }`}
          >
            <div className="flex-shrink-0 mt-0.5">
              {apiKeyToast.type === 'quota_exhausted' ? (
                <div className="p-2 bg-rose-500/20 rounded-xl border border-rose-400/40 animate-pulse">
                  <ShieldAlert className="h-6 w-6 text-rose-400" />
                </div>
              ) : apiKeyToast.type === 'rotation' ? (
                <div className="p-2 bg-indigo-500/20 rounded-xl border border-indigo-400/40 relative">
                  <motion.div
                    animate={{ rotate: 360 }}
                    transition={{ repeat: Infinity, duration: 3, ease: 'linear' }}
                  >
                    <RefreshCw className="h-6 w-6 text-indigo-400" />
                  </motion.div>
                </div>
              ) : (
                <div className="p-2 bg-emerald-500/20 rounded-xl border border-emerald-400/40">
                  <CheckCircle2 className="h-6 w-6 text-emerald-400" />
                </div>
              )}
            </div>

            <div className="flex-1 space-y-1">
              <div className="flex items-center justify-between gap-2">
                <h4 className="text-xs font-extrabold uppercase tracking-wider flex items-center gap-1.5">
                  <span>{apiKeyToast.title}</span>
                  {apiKeyToast.keyIndex && (
                    <span className="bg-indigo-500/30 text-indigo-300 border border-indigo-400/30 text-[10px] px-2 py-0.5 rounded-full font-mono">
                      Key #{apiKeyToast.keyIndex} / {apiKeyToast.totalKeys}
                    </span>
                  )}
                </h4>
                <button
                  onClick={() => setApiKeyToast(null)}
                  className="text-slate-400 hover:text-white p-0.5 rounded-md hover:bg-white/10 transition cursor-pointer"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>

              <p className="text-xs leading-relaxed opacity-90 font-medium">
                {apiKeyToast.message}
              </p>

              {/* Quick action buttons if quota exhausted */}
              {apiKeyToast.type === 'quota_exhausted' && (
                <div className="flex items-center gap-2 pt-2 flex-wrap text-xs font-bold">
                  <button
                    onClick={() => {
                      setApiKeyToast(null);
                      setActiveTab('soal');
                      const el = document.getElementById('ai-connection-panel');
                      if (el) el.scrollIntoView({ behavior: 'smooth' });
                    }}
                    className="bg-rose-600 hover:bg-rose-500 text-white px-3 py-1.5 rounded-lg flex items-center gap-1 transition shadow-sm cursor-pointer"
                  >
                    <Key className="h-3.5 w-3.5" />
                    <span>+ Tambah API Key Baru</span>
                  </button>
                  <button
                    onClick={() => {
                      handleSetAiMode('client');
                      setApiKeyToast(null);
                    }}
                    className="bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-500/40 px-3 py-1.5 rounded-lg flex items-center gap-1 transition cursor-pointer"
                  >
                    <Zap className="h-3.5 w-3.5 text-emerald-400" />
                    <span>Mode Direct Browser</span>
                  </button>
                  <a
                    href="https://aistudio.google.com/app/apikey"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-indigo-300 hover:underline text-[11px] flex items-center gap-1 ml-auto"
                  >
                    <span>Buat Key Gratis ↗</span>
                  </a>
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header */}
      <header id="header-section" className="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white shadow-md py-6 px-4 sm:px-8 no-print">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 p-2.5 rounded-xl shadow-inner border border-indigo-400">
              <Sparkles className="h-7 w-7 text-yellow-300 animate-pulse" />
            </div>
            <div>
              <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight font-sans">
                TKA SMA <span className="text-yellow-400">Assessment Creator</span>
              </h1>
              <p className="text-xs sm:text-sm text-slate-300 font-medium">
                Generator Prompt, Matriks Asesmen Kisi-Kisi, dan Pembuat Soal Akademik SMA Terintegrasi AI
              </p>
            </div>
          </div>
          
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2 bg-slate-800/80 backdrop-blur border border-slate-700 rounded-full px-4 py-1.5 text-xs text-slate-300">
              <span className={`h-2 w-2 rounded-full ${apiStatus === 'connected' ? 'bg-emerald-500 animate-ping' : 'bg-amber-500'}`} />
              <span>AI Engine: <b>{apiStatus === 'connected' ? 'Aktif & Siap' : 'Offline / Tanpa Kunci'}</b></span>
            </div>

            {/* User Profile Indicator */}
            <div className="flex items-center gap-2 bg-slate-850/80 backdrop-blur border border-slate-700 rounded-lg px-3 py-1.5">
              <div className="h-2 w-2 rounded-full bg-indigo-400" />
              <span className="text-xs text-slate-200 font-medium">
                {userName} ({userRole === 'admin' ? '🔑 Admin' : '👤 Guru'})
              </span>
              <button 
                onClick={handleSignOut}
                className="ml-2 text-xs bg-red-500/20 hover:bg-red-500/30 border border-red-500/30 text-red-300 font-semibold px-2 py-0.5 rounded-md transition"
              >
                Keluar
              </button>
            </div>

            <button
              onClick={handlePrint}
              className="bg-white/10 hover:bg-white/20 border border-white/20 text-white px-4 py-1.5 rounded-lg text-sm font-semibold flex items-center gap-2 transition"
            >
              <Printer className="h-4 w-4" />
              <span>Cetak / PDF</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Tabs Navigation */}
      <div id="tabs-navigation" className="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-xs no-print">
        <div className="max-w-7xl mx-auto px-4 sm:px-8">
          <nav className="flex space-x-1 sm:space-x-8 py-3 overflow-x-auto">
            <button
              id="tab-btn-config"
              onClick={() => setActiveTab('config')}
              className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-sm font-bold transition-all whitespace-nowrap ${
                activeTab === 'config'
                  ? 'bg-blue-50 text-blue-800 shadow-sm border border-blue-100'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <Sliders className="h-4.5 w-4.5" />
              <span>1. Input Parameter & Promt</span>
            </button>
            <button
              id="tab-btn-kisi"
              onClick={() => setActiveTab('kisi')}
              className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-sm font-bold transition-all relative whitespace-nowrap ${
                activeTab === 'kisi'
                  ? 'bg-indigo-50 text-indigo-800 shadow-sm border border-indigo-100'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <Layers className="h-4.5 w-4.5" />
              <span>2. Matriks Asesmen (Kisi-Kisi)</span>
              {kisiList.length > 0 && (
                <span className="absolute -top-1.5 -right-1.5 bg-indigo-600 text-white text-[10px] h-5 px-1.5 rounded-full flex items-center justify-center font-bold">
                  {kisiList.length}
                </span>
              )}
            </button>
            <button
              id="tab-btn-soal"
              onClick={() => setActiveTab('soal')}
              className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-sm font-bold transition-all relative whitespace-nowrap ${
                activeTab === 'soal'
                  ? 'bg-emerald-50 text-emerald-800 shadow-sm border border-emerald-100'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <BookOpen className="h-4.5 w-4.5" />
              <span>3. Butir Soal TKA SMA</span>
              {questions.length > 0 && (
                <span className="absolute -top-1.5 -right-1.5 bg-emerald-600 text-white text-[10px] h-5 px-1.5 rounded-full flex items-center justify-center font-bold">
                  {questions.length}
                </span>
              )}
            </button>
            <button
              id="tab-btn-jadwal"
              onClick={() => setActiveTab('jadwal')}
              className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-sm font-bold transition-all relative whitespace-nowrap ${
                activeTab === 'jadwal'
                  ? 'bg-rose-50 text-rose-800 shadow-sm border border-rose-100'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <Calendar className="h-4.5 w-4.5 text-rose-600" />
              <span>4. Jadwal Pembelajaran XII</span>
            </button>
            {userRole === 'admin' && (
              <button
                id="tab-btn-users"
                onClick={() => setActiveTab('users')}
                className={`flex items-center gap-2.5 px-4 py-2.5 rounded-xl text-sm font-bold transition-all relative whitespace-nowrap ${
                  activeTab === 'users'
                    ? 'bg-amber-50 text-amber-800 shadow-sm border border-amber-100'
                    : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                }`}
              >
                <Users className="h-4.5 w-4.5 text-amber-600" />
                <span>7. Manajemen Pengguna (Admin)</span>
                {usersList.length > 0 && (
                  <span className="absolute -top-1.5 -right-1.5 bg-amber-600 text-white text-[10px] h-5 px-1.5 rounded-full flex items-center justify-center font-bold">
                    {usersList.length}
                  </span>
                )}
              </button>
            )}
          </nav>
        </div>
      </div>

      {/* Main Content Stage */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-8">
        
        {/* Tab 1: Parameter & Prompt Generator */}
        {activeTab === 'config' && (
          <div id="config-panel" className="space-y-6 animate-fadeIn no-print">
            {/* Info Banner */}
            <div className="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white rounded-2xl shadow-sm p-4 flex flex-col sm:flex-row items-center justify-between gap-3">
              <div className="flex items-center gap-3">
                <div className="p-2.5 bg-blue-600/50 rounded-xl border border-blue-400/40">
                  <Sliders className="h-5 w-5 text-blue-200" />
                </div>
                <div>
                  <h3 className="text-base font-extrabold text-white flex items-center gap-2">
                    <span>1. Input Parameter Asesmen & Promt Generator</span>
                    <span className="text-[10px] bg-emerald-400 text-slate-950 font-black px-2.5 py-0.5 rounded-full uppercase">
                      Akses Semua Pengguna
                    </span>
                  </h3>
                  <p className="text-xs text-blue-200 font-medium">
                    Lengkapi parameter asesmen dan salin prompt untuk menyusun kisi-kisi dan butir soal.
                  </p>
                </div>
              </div>

              <div className="text-xs font-medium text-slate-200 bg-white/10 border border-white/20 px-3 py-1.5 rounded-xl flex items-center gap-2">
                <span className="h-2 w-2 rounded-full bg-emerald-400 animate-pulse"></span>
                <span>Akses Semua Pengguna: Parameter Asesmen & Prompt Generator AI</span>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 animate-fadeIn">
              {/* Left Column: Form Parameter */}
              <div className="lg:col-span-5 space-y-6">
                <section className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6 space-y-6">
                <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <Sliders className="h-5 w-5 text-blue-600" />
                    <h2 className="text-lg font-bold text-slate-800">Isian Parameter Asesmen TKA</h2>
                  </div>
                  <span className="text-[10px] bg-emerald-50 text-emerald-700 font-extrabold px-2.5 py-1 rounded-full border border-emerald-200">
                    Dapat Digunakan Semua Guru
                  </span>
                </div>

                  <div className="space-y-4">
                    {/* 1. Mata Pelajaran */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        1. Mata Pelajaran TKA
                      </label>
                      <select
                        value={config.mataPelajaran}
                        onChange={(e) => setConfig({ ...config, mataPelajaran: e.target.value })}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm font-medium"
                      >
                        <option value="">-- Pilih Mata Pelajaran --</option>
                        <optgroup label="Mata Pelajaran Wajib">
                          <option value="Matematika">Matematika</option>
                          <option value="Bahasa Indonesia">Bahasa Indonesia</option>
                          <option value="Bahasa Inggris">Bahasa Inggris</option>
                        </optgroup>
                        <optgroup label="Mata Pelajaran Pilihan">
                          <option value="Matematika Tingkat Lanjut">Matematika Tingkat Lanjut</option>
                          <option value="Bahasa Indonesia Tingkat Lanjut">Bahasa Indonesia Tingkat Lanjut</option>
                          <option value="Bahasa Inggris Tingkat Lanjut">Bahasa Inggris Tingkat Lanjut</option>
                          <option value="Fisika">Fisika</option>
                          <option value="Kimia">Kimia</option>
                          <option value="Biologi">Biologi</option>
                          <option value="Pendidikan Pancasila dan Kewarganegaraan">Pendidikan Pancasila dan Kewarganegaraan</option>
                          <option value="Ekonomi">Ekonomi</option>
                          <option value="Geografi">Geografi</option>
                          <option value="Sosiologi">Sosiologi</option>
                          <option value="Sejarah">Sejarah</option>
                          <option value="Antropologi">Antropologi</option>
                          <option value="Bahasa Prancis">Bahasa Prancis</option>
                          <option value="Bahasa Jerman">Bahasa Jerman</option>
                          <option value="Bahasa Jepang">Bahasa Jepang</option>
                          <option value="Bahasa Mandarin">Bahasa Mandarin</option>
                          <option value="Bahasa Korea">Bahasa Korea</option>
                          <option value="Bahasa Arab">Bahasa Arab</option>
                          <option value="Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK">Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK</option>
                        </optgroup>
                      </select>
                    </div>

                    {/* 2. Definisi */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        2. Definisi / Tujuan Asesmen
                      </label>
                      <textarea
                        rows={2}
                        value={config.definisi}
                        onChange={(e) => setConfig({ ...config, definisi: e.target.value })}
                        placeholder="Tujuan spesifik asesmen ini..."
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2 text-sm"
                      />
                    </div>

                    {/* 3. Muatan */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        3. Muatan Kurikulum / Tingkat
                      </label>
                      <input
                        type="text"
                        value={config.muatan}
                        onChange={(e) => setConfig({ ...config, muatan: e.target.value })}
                        placeholder="Contoh: Kurikulum Merdeka - Fase F Kelas XII"
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm"
                      />
                    </div>

                    {/* 4. Kompetensi */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        4. Kompetensi Dasar / Capaian
                      </label>
                      <textarea
                        rows={2}
                        value={config.kompetensi}
                        onChange={(e) => setConfig({ ...config, kompetensi: e.target.value })}
                        placeholder="Kompetensi umum yang akan diuji..."
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2 text-sm"
                      />
                    </div>

                    {/* Two column layouts */}
                    <div className="grid grid-cols-2 gap-4">
                      {/* 5. Bentuk Soal */}
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          5. Bentuk Soal
                        </label>
                        <select
                          value={config.bentukSoal}
                          onChange={(e) => setConfig({ ...config, bentukSoal: e.target.value as BentukSoal })}
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm"
                        >
                          <option value="pilihan_ganda_sederhana">PG Sederhana (1 Jawaban)</option>
                          <option value="mcma">PG Kompleks (MCMA)</option>
                          <option value="kategori">PG Kompleks (Kategori)</option>
                        </select>
                      </div>

                      {/* 6. Level Kognitif */}
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          6. Level Kognitif
                        </label>
                        <select
                          value={config.levelKognitif}
                          onChange={(e) => setConfig({ ...config, levelKognitif: e.target.value as LevelKognitif })}
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm font-semibold text-slate-700"
                        >
                          <option value="level_1">Pemahaman (Knowing)</option>
                          <option value="level_2">Penerapan (Applying)</option>
                          <option value="level_3">Penalaran (Reasoning)</option>
                        </select>

                        {/* Interactive Level Kognitif Definitions */}
                        <div className="mt-2 p-3 bg-slate-50 rounded-xl border border-slate-200/60 text-xs leading-relaxed text-slate-600 space-y-1">
                          {config.levelKognitif === 'level_1' && (
                            <div>
                              <span className="inline-flex items-center gap-1 font-bold text-amber-700 uppercase text-[10px] bg-amber-50 px-1.5 py-0.5 rounded border border-amber-100 mb-0.5">🧠 Pemahaman (Knowing)</span>
                              <p className="text-slate-600 font-medium text-[11px]">Mengenali, mengingat, dan memahami konsep dasar secara teoretis.</p>
                            </div>
                          )}
                          {config.levelKognitif === 'level_2' && (
                            <div>
                              <span className="inline-flex items-center gap-1 font-bold text-sky-700 uppercase text-[10px] bg-sky-50 px-1.5 py-0.5 rounded border border-sky-100 mb-0.5">⚙️ Penerapan (Applying)</span>
                              <p className="text-slate-600 font-medium text-[11px]">Menerapkan konsep, rumus, atau prosedur ilmiah pada situasi nyata / konkret.</p>
                            </div>
                          )}
                          {config.levelKognitif === 'level_3' && (
                            <div>
                              <span className="inline-flex items-center gap-1 font-bold text-purple-700 uppercase text-[10px] bg-purple-50 px-1.5 py-0.5 rounded border border-purple-100 mb-0.5">🧩 Penalaran (Reasoning)</span>
                              <p className="text-slate-600 font-medium text-[11px]">Berpikir kritis, menganalisis hubungan sebab-akibat, memecahkan masalah non-rutin, dan menalar secara logis.</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* 7. Elemen & 8. Sub-Elemen */}
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          7. Elemen/Materi
                        </label>
                        <input
                          type="text"
                          value={config.elemenMateri}
                          onChange={(e) => setConfig({ ...config, elemenMateri: e.target.value })}
                          placeholder="Materi utama"
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          8. Sub-Elemen
                        </label>
                        <input
                          type="text"
                          value={config.subElemenMateri}
                          onChange={(e) => setConfig({ ...config, subElemenMateri: e.target.value })}
                          placeholder="Submateri"
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm"
                        />
                      </div>
                    </div>

                    {/* 9. Batasan / Catatan */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        9. Batasan / Catatan (Opsional)
                      </label>
                      <input
                        type="text"
                        value={config.batasanCatatan}
                        onChange={(e) => setConfig({ ...config, batasanCatatan: e.target.value })}
                        placeholder="Contoh: Maksimum variabel, jenis bilangan..."
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm"
                      />
                    </div>

                    {/* 10. Opsi & 11. Jenis Soal */}
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          10. Pilihan Jawaban
                        </label>
                        <select
                          value={config.jumlahOpsi}
                          onChange={(e) => setConfig({ ...config, jumlahOpsi: Number(e.target.value) as JumlahOpsi })}
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm"
                        >
                          <option value={4}>4 Opsi (A, B, C, D)</option>
                          <option value={5}>5 Opsi (A, B, C, D, E)</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                          11. Jenis Soal
                        </label>
                        <select
                          value={config.jenisSoal}
                          onChange={(e) => setConfig({ ...config, jenisSoal: e.target.value as JenisSoal })}
                          className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-3 py-2.5 text-sm"
                        >
                          <option value="tunggal">Soal Tunggal</option>
                          <option value="grup">Soal Grup / Bersama</option>
                        </select>
                      </div>
                    </div>

                    {/* Jumlah Soal */}
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        Distribusi Target Jumlah Soal
                      </label>
                      <input
                        type="number"
                        min={1}
                        max={50}
                        value={config.jumlahSoal}
                        onChange={(e) => setConfig({ ...config, jumlahSoal: Number(e.target.value) })}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm"
                      />
                    </div>
                  </div>
                </section>
              </div>

              {/* Right Column: Konteks, Prompt & Actions */}
              <div className="lg:col-span-7 space-y-6">
                <section className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6">
                  <h3 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-1.5 uppercase tracking-wide">
                    <Globe className="h-4 w-4 text-indigo-600" />
                    Konteks Nusantara & Stimulus Tambahan
                  </h3>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Local Context Indonesia */}
                    <div>
                      <span className="block text-xs font-bold text-slate-500 mb-2">KONTEKS LOKAL INDONESIA</span>
                      <div className="space-y-1.5 max-h-40 overflow-y-auto pr-2 border border-slate-100 p-2 rounded-xl bg-slate-50/50">
                        {[
                          { key: 'Budaya Nusantara', label: '🎭 Budaya Nusantara (Adat & Seni)' },
                          { key: 'Geografis Indonesia', label: '🗺️ Geografis & Kewilayahan ID' },
                          { key: 'Kehidupan Sosial', label: '👥 Kehidupan Sosial & Kemasyarakatan' },
                          { key: 'Ekonomi Rakyat', label: '💰 Ekonomi Rakyat, Pasar & UMKM' },
                          { key: 'Teknologi Tradisional', label: '⚙️ Etno-Sains & Teknologi Tradisional' },
                          { key: 'Kearifan Lokal', label: '🏛️ Kearifan Lokal & Ekologi' },
                          { key: 'Keragaman Etnis', label: '🌈 Keragaman Etnis & Inklusivitas' },
                          { key: 'Profil Pelajar Pancasila', label: '🇮🇩 Profil Pelajar Pancasila & Karakter' },
                          { key: 'Isu Kontemporer', label: '📢 Isu Kontemporer & Realitas Empiris' },
                          { key: 'Kebahasaan Diksi', label: '🗣️ Aksesibilitas Diksi & Bahasa ID' }
                        ].map((item) => (
                          <label key={item.key} className="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer hover:bg-slate-100 p-1 rounded">
                            <input
                              type="checkbox"
                              checked={config.konteksLokal.includes(item.key)}
                              onChange={() => handleToggleContext(item.key)}
                              className="rounded text-blue-600 focus:ring-blue-500 h-3.5 w-3.5"
                            />
                            <span>{item.label}</span>
                          </label>
                        ))}
                      </div>
                    </div>

                    {/* Stimulus Content */}
                    <div>
                      <span className="block text-xs font-bold text-slate-500 mb-2">STIMULUS & PENGEMBANGAN KONTEN</span>
                      <div className="space-y-1.5 max-h-40 overflow-y-auto pr-2 border border-slate-100 p-2 rounded-xl bg-slate-50/50">
                        {[
                          { key: 'Teks Bacaan', label: '📖 Teks Bacaan Naratif/Sains' },
                          { key: 'Gambar/Ilustrasi', label: '🖼️ Gambar / Infografis' },
                          { key: 'Data/Tabel', label: '📊 Data Statistik / Tabel Empiris' },
                          { key: 'Grafik/Diagram', label: '📈 Grafik / Diagram Tren' },
                          { key: 'Kasus Nyata', label: '🔍 Kasus Nyata / Skenario Masalah' },
                          { key: 'Cerita Pendek', label: '📚 Cerita Pendek / Anekdot' },
                          { key: 'Berita/Artikel', label: '📰 Artikel Berita / Opini Media' },
                          { key: 'Peta/Denah', label: '🗺️ Peta Geospasial / Denah' },
                          { key: 'Dokumen Resmi', label: '📜 Dokumen / Regulasi Kebijakan' },
                          { key: 'Pernyataan Tokoh', label: '🎙️ Wawancara / Pernyataan Tokoh' }
                        ].map((item) => (
                          <label key={item.key} className="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer hover:bg-slate-100 p-1 rounded">
                            <input
                              type="checkbox"
                              checked={config.stimulusKonten.includes(item.key)}
                              onChange={() => handleToggleStimulus(item.key)}
                              className="rounded text-blue-600 focus:ring-blue-500 h-3.5 w-3.5"
                            />
                            <span>{item.label}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Quality Standard Checklist */}
                  <div className="mt-4 pt-4 border-t border-slate-100">
                    <span className="block text-xs font-bold text-slate-500 mb-2">CHECKLIST STANDAR KUALITAS SOAL TKA</span>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      {[
                        'Validasi Bahasa', 'Konstruksi Soal', 'Kesesuaian Materi', 
                        'Level Kognitif', 'Konteks Relevan', 'Tidak Bias', 
                        'Kejelasan Instruksi', 'Kunci Jawaban Tepat', 'Distractor Berkualitas', 
                        'Sesuai Kurikulum', 'Waktu Pengerjaan', 'Inklusivitas'
                      ].map((item) => (
                        <label key={item} className="flex items-center gap-2 text-[11px] text-slate-700 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={config.kualitasChecklist.includes(item)}
                            onChange={() => handleToggleQuality(item)}
                            className="rounded text-emerald-600 focus:ring-emerald-500 h-3 w-3"
                          />
                          <span className="truncate">{item}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                </section>

                <section className="bg-slate-900 text-slate-100 rounded-2xl shadow-lg p-6 overflow-hidden relative">
                  <div className="absolute top-0 right-0 p-3 flex gap-2">
                    <span className="bg-indigo-500/20 text-indigo-300 text-[10px] uppercase font-mono px-2 py-0.5 rounded border border-indigo-500/30">
                      Prompt Engine v3.5
                    </span>
                  </div>

                  <h3 className="text-base font-bold text-white mb-4 flex items-center gap-2">
                    <Sparkles className="h-5 w-5 text-yellow-400" />
                    Salin Prompt Megaprompt AI
                  </h3>

                  <div className="space-y-4">
                    {/* Prompt Soal TKA SMA */}
                    <div className="space-y-1.5 pt-2 border-t border-slate-800">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-300 uppercase tracking-wide">
                          PROMPT: PEMBUAT SOAL TKA SMA
                        </span>
                        <button
                          onClick={() => handleCopy(generatedSoalPrompt, 'soal')}
                          className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-200 px-2.5 py-1 rounded flex items-center gap-1 transition-all cursor-pointer"
                        >
                          {copiedSoal ? <Check className="h-3 w-3 text-emerald-400" /> : <Copy className="h-3 w-3" />}
                          <span>{copiedSoal ? 'Tersalin!' : 'Salin Prompt'}</span>
                        </button>
                      </div>
                      <div className="bg-slate-950 p-3 rounded-xl border border-slate-800 max-h-36 overflow-y-auto text-xs font-mono text-slate-300 whitespace-pre-wrap">
                        {generatedSoalPrompt}
                      </div>
                    </div>
                  </div>
                </section>

                {/* Instructions / Pedoman */}
                <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 text-xs text-blue-900 space-y-2">
                  <div className="flex items-center gap-1.5 font-bold">
                    <Info className="h-4 w-4 text-blue-700 flex-shrink-0" />
                    <span>Petunjuk Kerja Aplikasi:</span>
                  </div>
                  <ul className="list-disc pl-4 space-y-1">
                    <li>Langkah 1: Tentukan parameter mata pelajaran serta muatan kurikulum (khusus Admin).</li>
                    <li>Langkah 2: Salin prompt rancangan AI untuk digunakan pada AI playground/pilihan Anda (Gemini, ChatGPT, AI Studio).</li>
                    <li>Langkah 3: Tinjau dan edit matriks asesmen kisi-kisi Anda di tab kedua.</li>
                    <li>Langkah 4: Jalankan penyusunan butir soal, kemudian cetak atau download dalam format MS Word (.doc) atau Excel (.xls).</li>
                  </ul>
                </div>

                {/* Next Step Action Button */}
                <div className="bg-gradient-to-r from-indigo-900 to-slate-900 text-white rounded-2xl p-6 shadow-md flex flex-col sm:flex-row items-center justify-between gap-4">
                  <div>
                    <h4 className="text-sm font-bold text-white flex items-center gap-1.5">
                      <Check className="h-4 w-4 text-emerald-400" />
                      <span>Prompt AI Siap Disalin</span>
                    </h4>
                    <p className="text-xs text-indigo-200 mt-0.5">
                      Lanjutkan ke Langkah 2 untuk melihat dan mengelola Matriks Asesmen (Kisi-Kisi).
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={() => setActiveTab('kisi')}
                    className="w-full sm:w-auto bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold px-6 py-3 rounded-xl text-xs sm:text-sm flex items-center justify-center gap-2 shadow-lg transition cursor-pointer shrink-0"
                  >
                    <span>Lanjutkan ke Step 2: Matriks Asesmen</span>
                    <ArrowRight className="h-4 w-4" />
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Tab 2: Matriks Asesmen (Kisi-Kisi) */}
        {activeTab === 'kisi' && (
          <div id="kisi-panel" className="space-y-6 animate-fadeIn no-print">
            
            {/* Quick Summary Widget */}
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6 flex flex-col md:flex-row items-center justify-between gap-4 no-print">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Matriks Asesmen Kisi-Kisi Soal TKA</h2>
                <p className="text-xs text-slate-500">
                  Berikut merupakan sebaran distribusi butir soal berdasarkan tingkat kognitif dan kompetensi dasar {config.mataPelajaran}.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-3">
                <button
                  onClick={handleDeleteUnusedKisi}
                  className="bg-amber-50 text-amber-800 border border-amber-200 hover:bg-amber-100 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition cursor-pointer"
                  title="Hapus semua baris kisi-kisi yang belum memiliki butir soal"
                >
                  <Trash2 className="h-4.5 w-4.5 text-amber-600" />
                  <span>Hapus Kisi-Kisi Kosong</span>
                </button>
                {kisiList.length > 0 && (
                  <button
                    onClick={handleDeleteAllKisi}
                    className="bg-rose-50 text-rose-800 border border-rose-200 hover:bg-rose-100 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition cursor-pointer"
                    title="Hapus seluruh baris matriks kisi-kisi"
                  >
                    <Trash2 className="h-4.5 w-4.5 text-rose-600" />
                    <span>Hapus Semua Kisi-Kisi</span>
                  </button>
                )}
                <button
                  onClick={() => exportKisiToExcel(kisiList, config.mataPelajaran)}
                  className="bg-emerald-50 text-emerald-800 border border-emerald-200 hover:bg-emerald-100 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition"
                >
                  <FileSpreadsheet className="h-4.5 w-4.5 text-emerald-600" />
                  <span>Download Excel (.xls)</span>
                </button>
                <button
                  onClick={() => exportKisiToWord(kisiList, config.mataPelajaran, printConfig.pageSize)}
                  className="bg-blue-50 text-blue-800 border border-blue-200 hover:bg-blue-100 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition"
                >
                  <FileText className="h-4.5 w-4.5 text-blue-600" />
                  <span>Download Word (.doc)</span>
                </button>
              </div>
            </div>

            {/* Rekomendasi Matriks Asesmen Pusmendik */}
            <div className="bg-gradient-to-r from-indigo-950 via-slate-900 to-slate-950 text-white border border-slate-800 rounded-2xl shadow-md p-6 no-print space-y-4">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h3 className="text-base font-bold text-white flex items-center gap-2">
                    <Sparkles className="h-5 w-5 text-yellow-400" />
                    <span>🎯 Rekomendasi Matriks Asesmen {selectedPresetSubject} (Pusmendik)</span>
                  </h3>
                  <p className="text-xs text-indigo-200 mt-1">
                    Berikut adalah standar matriks asesmen/kisi-kisi resmi Pusmendik untuk pelajaran {selectedPresetSubject}. Anda dapat mengimpor sekaligus atau memilih per materi.
                  </p>
                </div>
                <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                  {userRole === 'admin' ? (
                    <div className="bg-slate-900 border border-slate-800 p-1 rounded-xl flex flex-wrap gap-1">
                      <button
                        onClick={() => handleSelectPresetSubject('Matematika')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Matematika' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        📐 Matematika
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Bahasa Indonesia')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Bahasa Indonesia' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🇮🇩 Bahasa Indonesia
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Bahasa Inggris')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Bahasa Inggris' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🇬🇧 Bahasa Inggris
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Matematika Tingkat Lanjut')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Matematika Tingkat Lanjut' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🚀 Mat Lanjut
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Bahasa Indonesia Tingkat Lanjut')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Bahasa Indonesia Tingkat Lanjut' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        ✍️ Indo Lanjut
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Bahasa Inggris Tingkat Lanjut')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Bahasa Inggris Tingkat Lanjut' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🗣️ Inggris Lanjut
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Fisika')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Fisika' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        ⚛️ Fisika
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Kimia')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Kimia' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🧪 Kimia
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Biologi')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Biologi' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🧬 Biologi
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('PPKN')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'PPKN' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🗳️ PPKN
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Ekonomi')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Ekonomi' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        💰 Ekonomi
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Geografi')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Geografi' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🌍 Geografi
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Sosiologi')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Sosiologi' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        👥 Sosiologi
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Sejarah Tingkat Lanjut')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Sejarah Tingkat Lanjut' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        📜 Sejarah Tingkat Lanjut
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Antropologi')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Antropologi' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🗿 Antropologi
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Bahasa Jepang')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Bahasa Jepang' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        🎌 Bahasa Jepang
                      </button>
                      <button
                        onClick={() => handleSelectPresetSubject('Produk Kreatif dan Kewirausahaan')}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${selectedPresetSubject === 'Produk Kreatif dan Kewirausahaan' ? 'bg-indigo-600 text-white shadow' : 'text-slate-400 hover:text-slate-200'}`}
                      >
                        💼 Kewirausahaan (PKK)
                      </button>
                    </div>
                  ) : (
                    <div className="bg-slate-900 border border-slate-800 px-3.5 py-2 rounded-xl flex items-center gap-2 text-xs font-bold text-amber-300">
                      <Lock className="h-4 w-4 text-amber-400 flex-shrink-0" />
                      <span>Hak Akses Terkunci Khusus Mata Pelajaran Ampuan: <b className="text-white">{config.mataPelajaran || selectedPresetSubject}</b></span>
                    </div>
                  )}
                  <button
                    onClick={handleImportAllPresets}
                    className="bg-yellow-400 hover:bg-yellow-500 text-slate-950 font-bold px-4 py-2.5 rounded-xl text-xs flex items-center justify-center gap-2 transition shadow-lg"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Impor Semua {
                      selectedPresetSubject === 'Matematika' 
                        ? PUSMENDIK_MATEMATIKA_PRESETS.length 
                        : selectedPresetSubject === 'Bahasa Indonesia' 
                        ? PUSMENDIK_BAHASA_INDONESIA_PRESETS.length 
                        : selectedPresetSubject === 'Bahasa Inggris' 
                        ? PUSMENDIK_BAHASA_INGGRIS_PRESETS.length 
                        : selectedPresetSubject === 'Matematika Tingkat Lanjut'
                        ? PUSMENDIK_MATEMATIKA_TL_PRESETS.length
                        : selectedPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
                        ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS.length
                        : selectedPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
                        ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS.length
                        : selectedPresetSubject === 'Fisika'
                        ? PUSMENDIK_FISIKA_PRESETS.length
                        : selectedPresetSubject === 'Kimia'
                        ? PUSMENDIK_KIMIA_PRESETS.length
                        : selectedPresetSubject === 'Biologi'
                        ? PUSMENDIK_BIOLOGI_PRESETS.length
                        : selectedPresetSubject === 'PPKN'
                        ? PUSMENDIK_PPKN_PRESETS.length
                        : selectedPresetSubject === 'Ekonomi'
                        ? PUSMENDIK_EKONOMI_PRESETS.length
                        : selectedPresetSubject === 'Geografi'
                        ? PUSMENDIK_GEOGRAFI_PRESETS.length
                        : selectedPresetSubject === 'Sosiologi'
                        ? PUSMENDIK_SOSIOLOGI_PRESETS.length
                        : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
                        ? PUSMENDIK_SEJARAH_TL_PRESETS.length
                        : selectedPresetSubject === 'Antropologi'
                        ? PUSMENDIK_ANTROPOLOGI_PRESETS.length
                        : selectedPresetSubject === 'Bahasa Jepang'
                        ? PUSMENDIK_BAHASA_JEPANG_PRESETS.length
                        : PUSMENDIK_PKK_PRESETS.length
                    } Matriks</span>
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 max-h-96 overflow-y-auto pr-2">
                {(selectedPresetSubject === 'Matematika' 
                  ? PUSMENDIK_MATEMATIKA_PRESETS 
                  : selectedPresetSubject === 'Bahasa Indonesia' 
                  ? PUSMENDIK_BAHASA_INDONESIA_PRESETS 
                  : selectedPresetSubject === 'Bahasa Inggris'
                  ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
                  : selectedPresetSubject === 'Matematika Tingkat Lanjut'
                  ? PUSMENDIK_MATEMATIKA_TL_PRESETS
                  : selectedPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
                  ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
                  : selectedPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
                  ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
                  : selectedPresetSubject === 'Fisika'
                  ? PUSMENDIK_FISIKA_PRESETS
                  : selectedPresetSubject === 'Kimia'
                  ? PUSMENDIK_KIMIA_PRESETS
                  : selectedPresetSubject === 'Biologi'
                  ? PUSMENDIK_BIOLOGI_PRESETS
                  : selectedPresetSubject === 'PPKN'
                  ? PUSMENDIK_PPKN_PRESETS
                  : selectedPresetSubject === 'Ekonomi'
                  ? PUSMENDIK_EKONOMI_PRESETS
                  : selectedPresetSubject === 'Geografi'
                  ? PUSMENDIK_GEOGRAFI_PRESETS
                  : selectedPresetSubject === 'Sosiologi'
                  ? PUSMENDIK_SOSIOLOGI_PRESETS
                  : selectedPresetSubject === 'Sejarah Tingkat Lanjut'
                  ? PUSMENDIK_SEJARAH_TL_PRESETS
                  : selectedPresetSubject === 'Antropologi'
                  ? PUSMENDIK_ANTROPOLOGI_PRESETS
                  : selectedPresetSubject === 'Bahasa Jepang'
                  ? PUSMENDIK_BAHASA_JEPANG_PRESETS
                  : PUSMENDIK_PKK_PRESETS
                ).map((preset, idx) => (
                  <div key={idx} className="bg-slate-900/60 border border-slate-800 p-4 rounded-xl flex flex-col justify-between hover:border-indigo-500/50 transition">
                    <div>
                      <div className="flex items-center justify-between gap-2 mb-1.5">
                        <span className="text-[10px] bg-indigo-500/20 text-indigo-300 font-semibold px-2 py-0.5 rounded border border-indigo-500/30">
                          {preset.elemenMateri}
                        </span>
                        <span className="text-[10px] font-mono text-slate-400">
                          Materi #{idx + 1}
                        </span>
                      </div>
                      <h4 className="text-xs font-bold text-yellow-300">{preset.subElemenMateri}</h4>
                      <p className="text-[11px] text-slate-300 mt-1.5 line-clamp-3 hover:line-clamp-none transition-all duration-300 leading-relaxed" title={preset.kompetensi}>
                        <b>Kompetensi:</b> {preset.kompetensi}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-1 italic">
                        <b>Batasan:</b> {preset.batasanCatatan}
                      </p>
                    </div>
                    <div className="flex gap-2 mt-3.5 pt-3 border-t border-slate-800">
                      <button
                        onClick={() => handleImportSinglePreset(preset, idx)}
                        disabled={importingPresetIds[`${preset.subElemenMateri}-${idx}`]}
                        className="flex-1 bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-400 text-white font-bold py-1.5 px-2.5 rounded text-[10px] transition text-center flex items-center justify-center gap-1"
                      >
                        {importingPresetIds[`${preset.subElemenMateri}-${idx}`] ? (
                          <>
                            <RefreshCw className="h-3 w-3 animate-spin" />
                            <span>Menyimpan...</span>
                          </>
                        ) : (
                          <span>+ Tambah Langsung</span>
                        )}
                      </button>
                      <button
                        onClick={() => handleLoadPresetToForm(preset)}
                        className="bg-slate-850 hover:bg-slate-800 text-slate-200 font-semibold py-1 px-2 rounded text-[10px] transition"
                      >
                        Muat ke Form
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Manually Add / Edit Row Form */}
            <div 
              id="kisi-form-section" 
              className={`rounded-2xl shadow-md p-6 no-print transition-all duration-300 ${
                isEditingKisi 
                  ? 'bg-gradient-to-r from-amber-50/90 via-orange-50/50 to-white border-2 border-amber-400 ring-4 ring-amber-200/60' 
                  : 'bg-gradient-to-r from-blue-50/80 via-indigo-50/40 to-white border-2 border-blue-400 ring-4 ring-blue-100'
              }`}
            >
              <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-200/80">
                <h3 className="text-sm font-extrabold uppercase tracking-wider flex items-center gap-2.5">
                  {isEditingKisi ? (
                    <>
                      <span className="p-1.5 bg-amber-500 text-white rounded-lg shadow-sm shadow-amber-500/30 animate-pulse">
                        <Edit className="h-4 w-4" />
                      </span>
                      <span className="text-amber-900">Edit Baris Kisi-Kisi (No. {kisiList.find(k => k.id === editingKisiId)?.no || ''})</span>
                      <span className="text-[10px] bg-amber-500 text-white font-mono px-2.5 py-0.5 rounded-full font-bold shadow-sm">
                        Mode Edit Aktif
                      </span>
                    </>
                  ) : (
                    <>
                      <span className="p-1.5 bg-blue-600 text-white rounded-lg shadow-sm shadow-blue-500/30">
                        <Plus className="h-4 w-4" />
                      </span>
                      <span className="text-blue-950 font-black">Tambah Baris Kisi-Kisi Manual</span>
                      <span className="text-[10px] bg-blue-100 text-blue-800 border border-blue-300 font-extrabold px-2.5 py-0.5 rounded-full">
                        Form Input Manual
                      </span>
                    </>
                  )}
                </h3>
                {isEditingKisi && (
                  <span className="text-xs text-amber-800 bg-amber-100/90 px-3 py-1 rounded-lg border border-amber-300 font-bold hidden sm:inline-block">
                    ✏️ Silakan perbarui data di bawah lalu klik tombol <b>Update</b>
                  </span>
                )}
              </div>
              <form onSubmit={handleSaveKisiForm} className="grid grid-cols-1 md:grid-cols-5 gap-4">
                <div className="md:col-span-1">
                  <label className="block text-[11px] font-bold text-indigo-600 mb-1">Mata Pelajaran</label>
                  <select
                    value={config.mataPelajaran}
                    onChange={(e) => {
                      const value = e.target.value;
                      setConfig(prev => ({ ...prev, mataPelajaran: value }));
                      // Also sync preset subject selection for Pusmendik recommendations
                      if (value === 'Pendidikan Pancasila dan Kewarganegaraan') {
                        setSelectedPresetSubject('PPKN');
                      } else if (value === 'Sejarah') {
                        setSelectedPresetSubject('Sejarah Tingkat Lanjut');
                      } else if (value === 'Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK') {
                        setSelectedPresetSubject('Produk Kreatif dan Kewirausahaan');
                      } else if (value) {
                        setSelectedPresetSubject(value as any);
                      }
                    }}
                    className="w-full bg-indigo-50 border border-indigo-200 text-indigo-950 font-bold rounded-lg px-2 py-1.5 text-xs focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                  >
                    <option value="Matematika">Matematika</option>
                    <option value="Bahasa Indonesia">Bahasa Indonesia</option>
                    <option value="Bahasa Inggris">Bahasa Inggris</option>
                    <option value="Matematika Tingkat Lanjut">Matematika Tingkat Lanjut</option>
                    <option value="Bahasa Indonesia Tingkat Lanjut">Bahasa Indonesia Tingkat Lanjut</option>
                    <option value="Bahasa Inggris Tingkat Lanjut">Bahasa Inggris Tingkat Lanjut</option>
                    <option value="Fisika">Fisika</option>
                    <option value="Kimia">Kimia</option>
                    <option value="Biologi">Biologi</option>
                    <option value="Pendidikan Pancasila dan Kewarganegaraan">PPKN</option>
                    <option value="Ekonomi">Ekonomi</option>
                    <option value="Geografi">Geografi</option>
                    <option value="Sosiologi">Sosiologi</option>
                    <option value="Sejarah">Sejarah</option>
                    <option value="Antropologi">Antropologi</option>
                    <option value="Bahasa Prancis">Bahasa Prancis</option>
                    <option value="Bahasa Jerman">Bahasa Jerman</option>
                    <option value="Bahasa Jepang">Bahasa Jepang</option>
                    <option value="Bahasa Mandarin">Bahasa Mandarin</option>
                    <option value="Bahasa Korea">Bahasa Korea</option>
                    <option value="Bahasa Arab">Bahasa Arab</option>
                    <option value="Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK">Kewirausahaan (PKK)</option>
                  </select>
                </div>
                <div className="md:col-span-1">
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Materi / Elemen</label>
                  <input
                    type="text"
                    value={kisiForm.elemenMateri}
                    onChange={(e) => setKisiForm({ ...kisiForm, elemenMateri: e.target.value })}
                    placeholder="Materi"
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Sub-Materi</label>
                  <input
                    type="text"
                    value={kisiForm.subElemenMateri}
                    onChange={(e) => setKisiForm({ ...kisiForm, subElemenMateri: e.target.value })}
                    placeholder="Sub-materi"
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Jenis Structure Soal</label>
                  <select
                    value={kisiForm.jenisSoal || 'tunggal'}
                    onChange={(e) => setKisiForm({ ...kisiForm, jenisSoal: e.target.value as JenisSoal })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-bold text-slate-800 focus:border-indigo-500"
                  >
                    <option value="tunggal">📌 Soal Tunggal (Stand-alone)</option>
                    <option value="grup">📚 Soal Grup (1 Stimulus Bersama)</option>
                  </select>
                  <div className="mt-1 text-[10px] text-slate-500 leading-snug font-medium">
                    {kisiForm.jenisSoal === 'grup' 
                      ? "📚 1 Stimulus Wacana/Data terpadu dipakai bersama" 
                      : "📌 Setiap soal memiliki stimulus mandiri"}
                  </div>
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Bentuk Soal</label>
                  <select
                    value={kisiForm.bentukSoal}
                    onChange={(e) => setKisiForm({ ...kisiForm, bentukSoal: e.target.value as BentukSoal })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs"
                  >
                    <option value="pilihan_ganda_sederhana">PG Sederhana</option>
                    <option value="mcma">PG Kompleks (MCMA)</option>
                    <option value="kategori">PG Kompleks (Kategori)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Tingkat Kognitif</label>
                  <select
                    value={kisiForm.levelKognitif}
                    onChange={(e) => setKisiForm({ ...kisiForm, levelKognitif: e.target.value as LevelKognitif })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-semibold text-slate-700"
                  >
                    <option value="level_1">Pemahaman (Knowing)</option>
                    <option value="level_2">Penerapan (Applying)</option>
                    <option value="level_3">Penalaran (Reasoning)</option>
                  </select>

                  {/* Manual form context helper */}
                  <div className="mt-1 text-[10px] text-slate-500 leading-snug">
                    {kisiForm.levelKognitif === 'level_1' && "🧠 Konsep dasar & ingatan"}
                    {kisiForm.levelKognitif === 'level_2' && "⚙️ Aplikasi pada fenomena nyata"}
                    {kisiForm.levelKognitif === 'level_3' && "🧩 Berpikir kritis & penalaran logis"}
                  </div>
                </div>

                <div className="md:col-span-2">
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Kompetensi yang Diuji</label>
                  <input
                    type="text"
                    value={kisiForm.kompetensi}
                    onChange={(e) => setKisiForm({ ...kisiForm, kompetensi: e.target.value })}
                    placeholder="Contoh: Menganalisis sistem persamaan linear..."
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Batasan / Catatan</label>
                  <input
                    type="text"
                    value={kisiForm.batasanCatatan}
                    onChange={(e) => setKisiForm({ ...kisiForm, batasanCatatan: e.target.value })}
                    placeholder="Contoh: Bilangan real positif"
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1">Jumlah Soal</label>
                  <input
                    type="number"
                    min={1}
                    value={kisiForm.jumlahSoal}
                    onChange={(e) => setKisiForm({ ...kisiForm, jumlahSoal: Number(e.target.value) })}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5 text-xs text-center focus:border-indigo-500"
                  />
                </div>

                {/* Konteks Nusantara & Stimulus Tambahan Inputs */}
                <div className="md:col-span-2">
                  <label className="block text-[11px] font-bold text-indigo-600 mb-1">🇮🇩 Deskripsi Konteks Nusantara Khusus (Opsional)</label>
                  <input
                    type="text"
                    value={kisiForm.konteksNusantara || ''}
                    onChange={(e) => setKisiForm({ ...kisiForm, konteksNusantara: e.target.value })}
                    placeholder="Contoh: Tradisi Lompat Batu Nias, Suku Baduy, Isu Maritim Indonesia"
                    className="w-full bg-slate-50 border border-indigo-100 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>
                <div className="md:col-span-3">
                  <label className="block text-[11px] font-bold text-indigo-600 mb-1">📖 Deskripsi Stimulus Tambahan Khusus (Opsional)</label>
                  <input
                    type="text"
                    value={kisiForm.stimulusTambahan || ''}
                    onChange={(e) => setKisiForm({ ...kisiForm, stimulusTambahan: e.target.value })}
                    placeholder="Contoh: Kutipan studi kasus, data tabel demografi, narasi berita"
                    className="w-full bg-slate-50 border border-indigo-100 rounded-lg px-3 py-1.5 text-xs focus:border-indigo-500"
                  />
                </div>

                {/* Konteks Nusantara, Stimulus & Checklist Checkboxes */}
                <div className="md:col-span-3 grid grid-cols-1 md:grid-cols-3 gap-4 border-t border-slate-100 pt-3 mt-2">
                  {/* KONTEKS LOKAL INDONESIA */}
                  <div>
                    <label className="block text-[11px] font-bold text-slate-500 mb-1.5">🎭 KONTEKS LOKAL INDONESIA</label>
                    <div className="space-y-1 max-h-36 overflow-y-auto pr-1 border border-slate-200/60 p-2 rounded-lg bg-slate-50/50">
                      {[
                        { key: 'Budaya Nusantara', label: '🎭 Budaya Nusantara (Adat & Seni)' },
                        { key: 'Geografis Indonesia', label: '🗺️ Geografis & Kewilayahan ID' },
                        { key: 'Kehidupan Sosial', label: '👥 Kehidupan Sosial & Kemasyarakatan' },
                        { key: 'Ekonomi Rakyat', label: '💰 Ekonomi Rakyat, Pasar & UMKM' },
                        { key: 'Teknologi Tradisional', label: '⚙️ Etno-Sains & Teknologi Tradisional' },
                        { key: 'Kearifan Lokal', label: '🏛️ Kearifan Lokal & Ekologi' },
                        { key: 'Keragaman Etnis', label: '🌈 Keragaman Etnis & Inklusivitas' },
                        { key: 'Profil Pelajar Pancasila', label: '🇮🇩 Profil Pelajar Pancasila & Karakter' },
                        { key: 'Isu Kontemporer', label: '📢 Isu Kontemporer & Realitas Empiris' },
                        { key: 'Kebahasaan Diksi', label: '🗣️ Aksesibilitas Diksi & Bahasa ID' }
                      ].map((item) => (
                        <label key={item.key} className="flex items-center gap-1.5 text-[10.5px] font-medium text-slate-700 cursor-pointer hover:bg-slate-100 p-0.5 rounded">
                          <input
                            type="checkbox"
                            checked={(kisiForm.konteksLokal || []).includes(item.key)}
                            onChange={() => handleToggleKisiContext(item.key)}
                            className="rounded text-indigo-600 focus:ring-indigo-500 h-3 w-3"
                          />
                          <span>{item.label}</span>
                        </label>
                      ))}
                    </div>
                  </div>

                  {/* STIMULUS & PENGEMBANGAN KONTEN */}
                  <div>
                    <label className="block text-[11px] font-bold text-slate-500 mb-1.5">📖 STIMULUS & PENGEMBANGAN KONTEN</label>
                    <div className="space-y-1 max-h-36 overflow-y-auto pr-1 border border-slate-200/60 p-2 rounded-lg bg-slate-50/50">
                      {[
                        { key: 'Teks Bacaan', label: '📖 Teks Bacaan Naratif/Sains' },
                        { key: 'Gambar/Ilustrasi', label: '🖼️ Gambar / Infografis' },
                        { key: 'Data/Tabel', label: '📊 Data Statistik / Tabel Empiris' },
                        { key: 'Grafik/Diagram', label: '📈 Grafik / Diagram Tren' },
                        { key: 'Kasus Nyata', label: '🔍 Kasus Nyata / Skenario Masalah' },
                        { key: 'Cerita Pendek', label: '📚 Cerita Pendek / Anekdot' },
                        { key: 'Berita/Artikel', label: '📰 Artikel Berita / Opini Media' },
                        { key: 'Peta/Denah', label: '🗺️ Peta Geospasial / Denah' },
                        { key: 'Dokumen Resmi', label: '📜 Dokumen / Regulasi Kebijakan' },
                        { key: 'Pernyataan Tokoh', label: '🎙️ Wawancara / Pernyataan Tokoh' }
                      ].map((item) => (
                        <label key={item.key} className="flex items-center gap-1.5 text-[10.5px] font-medium text-slate-700 cursor-pointer hover:bg-slate-100 p-0.5 rounded">
                          <input
                            type="checkbox"
                            checked={(kisiForm.stimulusKonten || []).includes(item.key)}
                            onChange={() => handleToggleKisiStimulus(item.key)}
                            className="rounded text-indigo-600 focus:ring-indigo-500 h-3 w-3"
                          />
                          <span>{item.label}</span>
                        </label>
                      ))}
                    </div>
                  </div>

                  {/* CHECKLIST STANDAR KUALITAS SOAL TKA */}
                  <div>
                    <label className="block text-[11px] font-bold text-slate-500 mb-1.5">📋 STANDAR KUALITAS SOAL TKA</label>
                    <div className="space-y-1 max-h-36 overflow-y-auto pr-1 border border-slate-200/60 p-2 rounded-lg bg-slate-50/50">
                      {[
                        'Validasi Bahasa (PUEBI)', 'Konstruksi Soal Presisi', 'Kesesuaian Materi & Kurikulum', 
                        'Level Kognitif Terdistribusi', 'Konteks Relevan Nusantara', 'Tidak Bias & Bebas SARA', 
                        'Kejelasan Instruksi & Kunci', 'Kunci Jawaban Tepat Single/Multi', 'Distractor Berkualitas & Homogen', 
                        'Anti-Hallucination Data', 'Estimasi Waktu Pengerjaan Proposional', 'Inklusivitas & Kesetaraan Gender'
                      ].map((item) => (
                        <label key={item} className="flex items-center gap-1.5 text-[10.5px] font-medium text-slate-700 cursor-pointer hover:bg-slate-100 p-0.5 rounded">
                          <input
                            type="checkbox"
                            checked={(kisiForm.kualitasChecklist || []).includes(item)}
                            onChange={() => handleToggleKisiQuality(item)}
                            className="rounded text-emerald-600 focus:ring-emerald-500 h-3 w-3"
                          />
                          <span className="truncate">{item}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Form Buttons */}
                <div className="md:col-span-2 flex justify-end gap-2 border-t border-slate-100 pt-3 mt-2">
                  <button
                    type="submit"
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-lg text-xs px-5 py-2 transition whitespace-nowrap"
                  >
                    {isEditingKisi ? 'Update' : 'Simpan'}
                  </button>
                  {isEditingKisi && (
                    <button
                      type="button"
                      onClick={() => {
                        setIsEditingKisi(false);
                        setEditingKisiId(null);
                        setKisiForm({
                          bentukSoal: 'pilihan_ganda_sederhana',
                          levelKognitif: 'level_2',
                          elemenMateri: '',
                          subElemenMateri: '',
                          kompetensi: '',
                          batasanCatatan: '',
                          jumlahSoal: 5,
                          konteksNusantara: '',
                          stimulusTambahan: '',
                          konteksLokal: [],
                          stimulusKonten: [],
                          kualitasChecklist: []
                        });
                      }}
                      className="bg-slate-300 text-slate-700 font-semibold rounded-lg text-xs px-4 py-2 transition"
                    >
                      Batal
                    </button>
                  )}
                </div>
              </form>
            </div>

            {/* Megaprompt AI Master Banner Card when kisiList has rows */}
            {kisiList.length > 0 && (
              <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 border border-indigo-800/80 rounded-2xl p-5 text-white flex flex-col md:flex-row items-start md:items-center justify-between gap-4 shadow-xl relative overflow-hidden my-4">
                <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-500/10 rounded-full blur-3xl pointer-events-none" />
                <div className="space-y-1.5 relative z-10">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="px-2.5 py-0.5 rounded-full bg-amber-400 text-slate-950 text-[10px] font-black uppercase tracking-wider flex items-center gap-1 shadow-sm">
                      <Sparkles className="h-3 w-3" /> MEGAPROMPT AI MATRIKS
                    </span>
                    <span className="text-xs font-mono text-indigo-200 font-bold bg-indigo-900/60 border border-indigo-700/50 px-2 py-0.5 rounded-md">
                      {kisiList.length} Baris Spesifikasi ({kisiList.reduce((acc, k) => acc + (k.jumlahSoal || 1), 0)} Soal Total)
                    </span>
                  </div>
                  <h4 className="text-base font-black text-white tracking-tight flex items-center gap-2">
                    Pembuat Prompt Otomatis AI (Master Megaprompt)
                  </h4>
                  <p className="text-xs text-indigo-200/80 leading-relaxed max-w-2xl">
                    Gabungkan seluruh {kisiList.length} baris matriks kisi-kisi menjadi <b>1 Prompt Utama Komprehensif</b> siap pakai untuk Gemini, ChatGPT, Claude, atau sinkronkan ke Wadah 1 CBT Generator.
                  </p>
                </div>
                <div className="flex items-center gap-2 w-full md:w-auto relative z-10">
                  <button
                    onClick={() => {
                      const text = buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
                      setMasterMegapromptText(text);
                      setIsMasterMegapromptModalOpen(true);
                    }}
                    className="w-full md:w-auto px-5 py-3 bg-gradient-to-r from-amber-400 via-yellow-400 to-amber-500 hover:from-amber-500 hover:to-yellow-500 text-slate-950 font-black text-xs rounded-xl shadow-lg transition flex items-center justify-center gap-2 border border-amber-300"
                  >
                    <Zap className="h-4 w-4 fill-slate-950" />
                    <span>Buka Megaprompt AI Utama ({kisiList.length} Baris)</span>
                  </button>
                </div>
              </div>
            )}

            {/* Matrix Table */}
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
              <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">MATRIKS ASESMEN KISI-KISI SOAL</span>
                  <span className="text-xs bg-indigo-50 border border-indigo-200 text-indigo-800 font-mono px-2.5 py-0.5 rounded-full font-bold">
                    Total Kisi-Kisi: {kisiList.length} baris ({kisiList.reduce((acc, k) => acc + (k.jumlahSoal || 1), 0)} Soal Target)
                  </span>
                </div>
                {kisiList.length > 0 && (
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setIsImportAiModalOpen(true)}
                      className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm cursor-pointer"
                      title="Masukkan atau upload file teks/Word hasil Megaprompt Gemini AI"
                    >
                      <FileUp className="h-3.5 w-3.5 text-amber-300" />
                      <span>📥 Masukkan Soal AI</span>
                    </button>
                    <button
                      onClick={() => {
                        const text = buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
                        setMasterMegapromptText(text);
                        setIsMasterMegapromptModalOpen(true);
                      }}
                      className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm cursor-pointer"
                    >
                      <Sparkles className="h-3.5 w-3.5 text-amber-300" />
                      <span>⚡ Megaprompt AI ({kisiList.length} Baris)</span>
                    </button>
                    <button
                      onClick={handleDeleteUnusedKisi}
                      className="px-3 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 rounded-xl text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm cursor-pointer"
                      title="Hapus baris kisi-kisi yang belum memiliki butir soal"
                    >
                      <Trash2 className="h-3.5 w-3.5 text-amber-600" />
                      <span>Hapus Kisi-Kisi Kosong</span>
                    </button>
                    <button
                      onClick={handleDeleteAllKisi}
                      className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm cursor-pointer"
                      title="Hapus seluruh baris matriks kisi-kisi"
                    >
                      <Trash2 className="h-3.5 w-3.5 text-rose-600" />
                      <span>Hapus Semua Kisi-Kisi</span>
                    </button>
                  </div>
                )}
              </div>

              <div className="overflow-x-auto max-h-[75vh] overflow-y-auto rounded-xl border border-slate-200 shadow-xs relative">
                <table className="w-full text-left text-sm border-collapse relative">
                  <thead className="sticky top-0 z-20 bg-slate-100 border-b border-slate-200 shadow-xs">
                    <tr className="bg-slate-100 text-slate-700 font-bold">
                      <th className="py-3.5 px-4 text-center w-12 sticky top-0 bg-slate-100 shadow-xs">No</th>
                      <th className="py-3.5 px-4 min-w-[120px] text-center sticky top-0 bg-slate-100 shadow-xs">Jenis Soal</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Bentuk Soal</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Tingkat Kognitif</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Elemen / Materi</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Sub-elemen / Submateri</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Kompetensi yang Diuji</th>
                      <th className="py-3.5 px-4 sticky top-0 bg-slate-100 shadow-xs">Batasan / Catatan</th>
                      <th className="py-3.5 px-4 w-[220px] sticky top-0 bg-slate-100 shadow-xs">🇮🇩 Konteks & Stimulus</th>
                      <th className="py-3.5 px-4 text-center w-24 sticky top-0 bg-slate-100 shadow-xs">Jumlah Soal</th>
                      <th className="py-3.5 px-4 text-center w-32 no-print sticky top-0 bg-slate-100 shadow-xs">Tindakan</th>
                    </tr>
                  </thead>
                  <tbody>
                    {kisiList.length === 0 ? (
                      <tr>
                        <td colSpan={11} className="text-center py-12 px-6">
                          <div className="max-w-md mx-auto flex flex-col items-center gap-3">
                            <div className="p-3 bg-amber-50 text-amber-600 rounded-2xl border border-amber-200">
                              <Sparkles className="h-6 w-6 animate-pulse" />
                            </div>
                            <h4 className="text-sm font-bold text-slate-800">Belum Ada Data Matriks Asesmen</h4>
                            <p className="text-xs text-slate-500 leading-relaxed mb-2">
                              Tabel matriks masih kosong. Anda dapat mengimpor seluruh paket matriks standar Pusmendik untuk mata pelajaran <strong className="text-slate-700">{selectedPresetSubject}</strong> secara otomatis, atau menambahkan sendiri secara manual.
                            </p>
                            <div className="flex flex-wrap items-center justify-center gap-2.5">
                              <button
                                type="button"
                                onClick={handleImportAllPresets}
                                className="bg-amber-500 hover:bg-amber-600 text-slate-950 font-extrabold px-4 py-2 rounded-xl text-xs flex items-center gap-1.5 shadow-md transition-all cursor-pointer"
                              >
                                <Sparkles className="h-4 w-4" />
                                <span>⚡ Impor Semua Matriks ({selectedPresetSubject})</span>
                              </button>
                            </div>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      kisiList.map((item) => (
                        <tr 
                          key={item.id} 
                          className={`border-b border-slate-100 transition-all duration-300 text-xs ${
                            editingKisiId === item.id 
                              ? 'bg-amber-50/90 ring-2 ring-amber-400 font-medium' 
                              : 'hover:bg-slate-50'
                          }`}
                        >
                          <td className="py-4 px-4 text-center font-bold text-slate-700">{item.no}</td>
                          <td className="py-4 px-4 text-center min-w-[120px]">
                            <button
                              type="button"
                              onClick={() => handleToggleJenisSoal(item.id, item.jenisSoal)}
                              className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10.5px] font-bold border transition-all cursor-pointer shadow-sm ${
                                item.jenisSoal === 'grup'
                                  ? 'bg-purple-100 text-purple-900 border-purple-300 hover:bg-purple-200'
                                  : 'bg-sky-100 text-sky-900 border-sky-300 hover:bg-sky-200'
                              }`}
                              title="Klik untuk mengubah jenis (Soal Tunggal ↔ Soal Grup)"
                            >
                              {item.jenisSoal === 'grup' ? (
                                <>
                                  <Layers className="h-3 w-3 text-purple-600" />
                                  <span>📚 Soal Grup</span>
                                </>
                              ) : (
                                <>
                                  <FileText className="h-3 w-3 text-sky-600" />
                                  <span>📌 Soal Tunggal</span>
                                </>
                              )}
                            </button>
                          </td>
                          <td className="py-4 px-4">
                            <span className={`inline-block px-2.5 py-1 rounded-full font-semibold text-[10px] ${
                              item.bentukSoal === 'pilihan_ganda_sederhana' ? 'bg-blue-100 text-blue-800' :
                              item.bentukSoal === 'mcma' ? 'bg-indigo-100 text-indigo-800' : 'bg-emerald-100 text-emerald-800'
                            }`}>
                              {getBentukSoalLabel(item.bentukSoal)}
                            </span>
                          </td>
                          <td className="py-4 px-4 min-w-[220px]">
                            {item.levelKognitif === 'level_1' && (
                              <div className="space-y-1">
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-200">
                                  🧠 Pemahaman (Knowing)
                                </span>
                                <p className="text-[10.5px] text-slate-500 leading-snug font-medium">
                                  Mengenali, mengingat, dan memahami konsep dasar secara teoretis.
                                </p>
                              </div>
                            )}
                            {item.levelKognitif === 'level_2' && (
                              <div className="space-y-1">
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-sky-100 text-sky-800 border border-sky-200">
                                  ⚙️ Penerapan (Applying)
                                </span>
                                <p className="text-[10.5px] text-slate-500 leading-snug font-medium">
                                  Menerapkan konsep pada fenomena nyata / konkret.
                                </p>
                              </div>
                            )}
                            {item.levelKognitif === 'level_3' && (
                              <div className="space-y-1">
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-purple-100 text-purple-800 border border-purple-200">
                                  🧩 Penalaran (Reasoning)
                                </span>
                                <p className="text-[10.5px] text-slate-500 leading-snug font-medium">
                                  Berpikir kritis, menganalisis hubungan sebab-akibat, memecahkan masalah non-rutin, dan menalar secara logis.
                                </p>
                              </div>
                            )}
                          </td>
                          <td className="py-4 px-4 font-bold text-slate-800">{item.elemenMateri}</td>
                          <td className="py-4 px-4 text-slate-600">{item.subElemenMateri}</td>
                          <td className="py-4 px-4 text-slate-700 leading-relaxed font-medium">{item.kompetensi}</td>
                          <td className="py-4 px-4 text-slate-500 italic">{item.batasanCatatan || '-'}</td>
                          <td className="py-4 px-4 min-w-[240px]">
                            {/* Konteks Nusantara */}
                            {((item.konteksLokal && item.konteksLokal.length > 0) || item.konteksNusantara) ? (
                              <div className="mb-2">
                                <span className="inline-block font-bold text-indigo-700 bg-indigo-50 px-1 rounded text-[10px] mr-1">🇮🇩 Konteks:</span>
                                {item.konteksLokal && item.konteksLokal.length > 0 && (
                                  <div className="flex flex-wrap gap-1 mt-1 mb-1">
                                    {item.konteksLokal.map(k => (
                                      <span key={k} className="bg-slate-100 text-slate-800 text-[9px] px-1.5 py-0.5 rounded font-medium">
                                        {k}
                                      </span>
                                    ))}
                                  </div>
                                )}
                                {item.konteksNusantara && (
                                  <span className="text-slate-700 block text-[11px] leading-relaxed">{item.konteksNusantara}</span>
                                )}
                              </div>
                            ) : null}

                            {/* Stimulus Tambahan */}
                            {((item.stimulusKonten && item.stimulusKonten.length > 0) || item.stimulusTambahan) ? (
                              <div className="mb-2">
                                <span className="inline-block font-bold text-purple-700 bg-purple-50 px-1 rounded text-[10px] mr-1">📖 Stimulus:</span>
                                {item.stimulusKonten && item.stimulusKonten.length > 0 && (
                                  <div className="flex flex-wrap gap-1 mt-1 mb-1">
                                    {item.stimulusKonten.map(s => (
                                      <span key={s} className="bg-purple-100/50 text-purple-800 text-[9px] px-1.5 py-0.5 rounded font-medium">
                                        {s}
                                      </span>
                                    ))}
                                  </div>
                                )}
                                {item.stimulusTambahan && (
                                  <span className="text-slate-700 block text-[11px] leading-relaxed">{item.stimulusTambahan}</span>
                                )}
                              </div>
                            ) : null}

                            {/* Standar Mutu */}
                            {item.kualitasChecklist && item.kualitasChecklist.length > 0 ? (
                              <div>
                                <span className="inline-block font-bold text-emerald-700 bg-emerald-50 px-1 rounded text-[10px] mr-1">📋 Standar Mutu:</span>
                                <div className="flex flex-wrap gap-1 mt-1">
                                  {item.kualitasChecklist.map(c => (
                                    <span key={c} className="bg-emerald-100/50 text-emerald-800 text-[9px] px-1.5 py-0.5 rounded font-medium">
                                      {c}
                                    </span>
                                  ))}
                                </div>
                              </div>
                            ) : null}

                            {!item.konteksNusantara && !item.stimulusTambahan && (!item.konteksLokal || item.konteksLokal.length === 0) && (!item.stimulusKonten || item.stimulusKonten.length === 0) && (!item.kualitasChecklist || item.kualitasChecklist.length === 0) ? (
                              <span className="text-slate-400 italic">-</span>
                            ) : null}
                          </td>
                          <td className="py-4 px-4 text-center">
                            <div className="font-bold text-slate-800 text-sm">
                              {item.jumlahSoal}
                            </div>
                            <div className="mt-1">
                              {(() => {
                                const count = questions.filter(q => q.kisiKisiId === item.id).length;
                                return count > 0 ? (
                                  <span className="inline-block bg-emerald-100 text-emerald-800 font-mono text-[9px] font-bold px-1.5 py-0.5 rounded-full">
                                    {count} Soal Terbuat
                                  </span>
                                ) : (
                                  <span className="inline-block bg-slate-100 text-slate-400 font-mono text-[9px] font-bold px-1.5 py-0.5 rounded-full">
                                    Belum Ada Soal
                                  </span>
                                );
                              })()}
                            </div>
                          </td>
                          <td className="py-4 px-4 text-center space-y-1.5 no-print">
                            {deletingKisiId === item.id ? (
                              <div className="bg-rose-50 border border-rose-200 p-1.5 rounded-lg space-y-1">
                                <span className="text-[10px] font-bold text-rose-700 block text-center">Yakin Hapus?</span>
                                <div className="flex gap-1 justify-center">
                                  <button
                                    onClick={() => {
                                      handleDeleteKisi(item.id);
                                      setDeletingKisiId(null);
                                    }}
                                    className="bg-red-600 hover:bg-red-700 text-white font-extrabold px-1.5 py-0.5 rounded text-[10px] transition"
                                  >
                                    Ya
                                  </button>
                                  <button
                                    onClick={() => setDeletingKisiId(null)}
                                    className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold px-1.5 py-0.5 rounded text-[10px] transition"
                                  >
                                    Batal
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <>
                                <div className="flex justify-center gap-1.5 items-center">
                                  <button
                                    onClick={() => handleEditKisi(item)}
                                    className={`group relative px-2.5 py-1.5 rounded-lg text-[11px] font-extrabold flex items-center justify-center gap-1.5 transition-all duration-300 shadow-sm cursor-pointer ${
                                      editingKisiId === item.id
                                        ? 'bg-amber-500 text-white ring-2 ring-amber-300 shadow-amber-500/40 animate-pulse'
                                        : 'bg-gradient-to-r from-amber-500 to-orange-500 text-white hover:from-amber-600 hover:to-orange-600 border border-amber-300 shadow-amber-500/20 hover:shadow-md hover:scale-105 active:scale-95'
                                    }`}
                                    title="Edit baris kisi-kisi ini"
                                  >
                                    <Edit className={`h-3.5 w-3.5 transition-transform duration-300 ${
                                      editingKisiId === item.id 
                                        ? 'animate-spin' 
                                        : 'group-hover:-rotate-12 group-hover:scale-125'
                                    }`} />
                                    <span>{editingKisiId === item.id ? 'Di-edit' : 'Edit'}</span>
                                    {editingKisiId !== item.id && (
                                      <span className="absolute -top-1 -right-1 flex h-2.5 w-2.5">
                                        <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-300 opacity-75"></span>
                                        <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-400"></span>
                                      </span>
                                    )}
                                  </button>
                                  <button
                                    onClick={() => setDeletingKisiId(item.id)}
                                    className="text-slate-400 hover:text-red-600 p-1.5 hover:bg-red-50 rounded-lg transition-all duration-200"
                                    title="Hapus baris"
                                  >
                                    <Trash2 className="h-4 w-4" />
                                  </button>
                                </div>
                                <div className="flex flex-col gap-1.5 mt-1.5">
                                  <button
                                    onClick={() => handleOpenPromptGenerator(item)}
                                    className="w-full bg-indigo-50 text-indigo-800 hover:bg-indigo-100 text-[10px] font-bold py-1.5 px-1.5 rounded-lg border border-indigo-100 flex items-center justify-center gap-1 transition"
                                    title="Buat prompt otomatis untuk disalin ke AI eksternal"
                                  >
                                    <FileText className="h-3 w-3 text-indigo-600" />
                                    <span>Buat Prompt</span>
                                  </button>
                                </div>
                              </>
                            )}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Tab 3: Pembuat Soal TKA SMA */}
        {activeTab === 'soal' && (
          <div id="soal-panel" className="space-y-6 animate-fadeIn no-print">
            
            {/* Step-by-Step Megaprompt & Excel Import Banner */}
            <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl shadow-md p-5 border border-indigo-900/50 relative overflow-hidden">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3.5 border-b border-indigo-800/60 pb-3">
                <div className="flex items-center gap-2.5">
                  <div className="w-8 h-8 rounded-xl bg-amber-400 text-slate-950 flex items-center justify-center font-black text-sm shadow">
                    ⚡
                  </div>
                  <div>
                    <h3 className="text-sm font-extrabold text-white tracking-tight">
                      Alur Pembuatan Soal dengan "Pembuat Prompt Otomatis AI (Megaprompt)"
                    </h3>
                    <p className="text-[11px] text-indigo-200">
                      5 langkah praktis dari Matriks Asesmen ke Gemini AI hingga menghasilkan file Excel & Word rapi.
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setIsImportAiModalOpen(true)}
                    className="bg-amber-400 hover:bg-amber-500 text-slate-950 font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 transition shadow"
                  >
                    <Upload className="h-3.5 w-3.5" />
                    <span>Upload Excel Soal</span>
                  </button>
                  <button
                    onClick={() => downloadTemplateExcelSoal(config.mataPelajaran)}
                    className="bg-emerald-500 hover:bg-emerald-600 text-slate-950 font-black text-xs px-3.5 py-1.5 rounded-xl flex items-center gap-1.5 transition shadow"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>Download Template Excel</span>
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3 text-xs">
                <div className="bg-white/5 border border-white/10 rounded-xl p-3 flex flex-col justify-between">
                  <div>
                    <span className="bg-amber-400 text-slate-950 font-black text-[10px] px-2 py-0.5 rounded-md inline-block mb-1.5">
                      Langkah 1
                    </span>
                    <h4 className="font-bold text-slate-100 text-[11px]">Salin Megaprompt AI</h4>
                    <p className="text-[10.5px] text-slate-300 mt-1 leading-snug">
                      Buka <strong>1. Kisi-Kisi TKA</strong>, lalu klik <strong>"⚡ Megaprompt AI"</strong> pada baris kisi-kisi pilihan Anda.
                    </p>
                  </div>
                </div>

                <div className="bg-white/5 border border-white/10 rounded-xl p-3 flex flex-col justify-between">
                  <div>
                    <span className="bg-indigo-400 text-slate-950 font-black text-[10px] px-2 py-0.5 rounded-md inline-block mb-1.5">
                      Langkah 2
                    </span>
                    <h4 className="font-bold text-slate-100 text-[11px]">Tempel di Gemini AI</h4>
                    <p className="text-[10.5px] text-slate-300 mt-1 leading-snug">
                      Buka <a href="https://gemini.google.com" target="_blank" rel="noreferrer" className="underline text-amber-300 font-bold">gemini.google.com</a> atau Google AI Studio, lalu tempel Megaprompt.
                    </p>
                  </div>
                </div>

                <div className="bg-white/5 border border-white/10 rounded-xl p-3 flex flex-col justify-between">
                  <div>
                    <span className="bg-cyan-400 text-slate-950 font-black text-[10px] px-2 py-0.5 rounded-md inline-block mb-1.5">
                      Langkah 3
                    </span>
                    <h4 className="font-bold text-slate-100 text-[11px]">Salin Hasil ke Excel / Word</h4>
                    <p className="text-[10.5px] text-slate-300 mt-1 leading-snug">
                      Gemini AI membuat tabel & naskah soal HOTS. Salin tabel ke Excel atau simpan sebagai file Excel/Word/Text.
                    </p>
                  </div>
                </div>

                <div className="bg-white/5 border border-white/10 rounded-xl p-3 flex flex-col justify-between">
                  <div>
                    <span className="bg-emerald-400 text-slate-950 font-black text-[10px] px-2 py-0.5 rounded-md inline-block mb-1.5">
                      Langkah 4
                    </span>
                    <h4 className="font-bold text-slate-100 text-[11px]">Upload Excel Ke Sini</h4>
                    <p className="text-[10.5px] text-slate-300 mt-1 leading-snug">
                      Klik tombol <strong>"Upload Excel Soal"</strong>. Pilih file Excel (`.xlsx`), Word, atau file teks dari Gemini AI.
                    </p>
                  </div>
                </div>

                <div className="bg-white/5 border border-white/10 rounded-xl p-3 flex flex-col justify-between">
                  <div>
                    <span className="bg-purple-400 text-slate-950 font-black text-[10px] px-2 py-0.5 rounded-md inline-block mb-1.5">
                      Langkah 5
                    </span>
                    <h4 className="font-bold text-slate-100 text-[11px]">Output Terisi Rapi</h4>
                    <p className="text-[10.5px] text-slate-300 mt-1 leading-snug">
                      Seluruh butir soal langsung tersusun rapi di tabel di bawah. Siap diedit, dicetak, atau didownload kembali!
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {/* Upper Action Panel */}
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6 space-y-4 no-print">
              <div className="flex flex-col md:flex-row items-center justify-between gap-4">
                <div>
                  <h2 className="text-xl font-bold text-slate-800">Daftar Soal TKA SMA</h2>
                  <p className="text-xs text-slate-500">
                    Sistem mendukung penggabungan 3 bentuk soal TKA SMA: PG Sederhana, PG Kompleks (MCMA), dan PGK Kategori (Respon Pernyataan Multi-Kategori).
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => setIsImportAiModalOpen(true)}
                    className="bg-amber-500 hover:bg-amber-600 text-slate-950 px-3.5 py-2 rounded-xl text-xs font-black flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                    title="Upload file Excel (.xlsx), Word, atau tempel teks dari Megaprompt Gemini AI"
                  >
                    <Upload className="h-4 w-4" />
                    <span>Upload Excel / Hasil AI</span>
                  </button>
                  <button
                    onClick={() => downloadTemplateExcelSoal(config.mataPelajaran, selectedBentukFilter)}
                    className="bg-teal-600 hover:bg-teal-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                    title={`Download template Excel untuk ${selectedBentukFilter === 'all' ? 'Semua Bentuk Soal' : getBentukSoalLabel(selectedBentukFilter)}`}
                  >
                    <Download className="h-4 w-4" />
                    <span>Template Excel</span>
                  </button>
                  <button
                    onClick={() => {
                      const targetQuestions = selectedBentukFilter === 'all'
                        ? questions
                        : questions.filter(q => {
                            if (selectedBentukFilter === 'kategori') return isKategoriSoal(q);
                            if (selectedBentukFilter === 'mcma') return q.bentukSoal === 'mcma';
                            if (selectedBentukFilter === 'pilihan_ganda_sederhana') return q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q));
                            return true;
                          });
                      exportQuestionsToExcel(targetQuestions, printConfig.subjectName || config.mataPelajaran, printConfig.examName, printConfig.showAnswerKey);
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                  >
                    <FileSpreadsheet className="h-4 w-4" />
                    <span>Download Excel</span>
                  </button>
                  <button
                    onClick={() => {
                      const targetQuestions = selectedBentukFilter === 'all'
                        ? questions
                        : questions.filter(q => {
                            if (selectedBentukFilter === 'kategori') return isKategoriSoal(q);
                            if (selectedBentukFilter === 'mcma') return q.bentukSoal === 'mcma';
                            if (selectedBentukFilter === 'pilihan_ganda_sederhana') return q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q));
                            return true;
                          });
                      exportQuestionsToWord(targetQuestions, printConfig.subjectName || config.mataPelajaran, printConfig.pageSize, printConfig.examName, printConfig.showAnswerKey);
                    }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                  >
                    <FileText className="h-4 w-4" />
                    <span>Download Word</span>
                  </button>
                  <button
                    onClick={() => {
                      const targetQuestions = selectedBentukFilter === 'all'
                        ? questions
                        : questions.filter(q => {
                            if (selectedBentukFilter === 'kategori') return isKategoriSoal(q);
                            if (selectedBentukFilter === 'mcma') return q.bentukSoal === 'mcma';
                            if (selectedBentukFilter === 'pilihan_ganda_sederhana') return q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q));
                            return true;
                          });
                      exportQuestionsWithImagesZip(targetQuestions, printConfig.subjectName || config.mataPelajaran, printConfig.examName);
                    }}
                    className="bg-purple-600 hover:bg-purple-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                    title="Download file JSON lengkap beserta semua gambar terpisah dalam format ZIP"
                  >
                    <FileCode className="h-4 w-4" />
                    <span>Download JSON + Gambar (.zip)</span>
                  </button>
                  <button
                    onClick={() => {
                      const targetQuestions = selectedBentukFilter === 'all'
                        ? questions
                        : questions.filter(q => {
                            if (selectedBentukFilter === 'kategori') return isKategoriSoal(q);
                            if (selectedBentukFilter === 'mcma') return q.bentukSoal === 'mcma';
                            if (selectedBentukFilter === 'pilihan_ganda_sederhana') return q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q));
                            return true;
                          });
                      exportQuestionsToJSON(targetQuestions, printConfig.subjectName || config.mataPelajaran, printConfig.examName);
                    }}
                    className="bg-amber-600 hover:bg-amber-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                    title="Download file JSON lengkap beserta data gambar base64"
                  >
                    <Code className="h-4 w-4" />
                    <span>Download JSON</span>
                  </button>
                  {questions.length > 0 && (
                    <button
                      onClick={handleResequenceQuestionNumbers}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                      title="Urutkan kembali nomor urut seluruh soal dari 1 s.d. N"
                    >
                      <ListOrdered className="h-4 w-4" />
                      <span>Rapikan Nomor (1-{questions.length})</span>
                    </button>
                  )}
                  {questions.length > 0 && (
                    <button
                      onClick={() => {
                        if (confirm(`Apakah Anda yakin ingin menghapus SEMUA (${questions.length}) butir soal yang tersimpan?`)) {
                          handleDeleteAllQuestions();
                        }
                      }}
                      className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm cursor-pointer"
                      title="Hapus seluruh butir soal yang ada"
                    >
                      <Trash2 className="h-4 w-4" />
                      <span>Hapus Semua</span>
                    </button>
                  )}
                  <button
                    onClick={() => {
                      setEditingQuestionId(null);
                      setIsEditingQuestion(true);
                      setQuestionForm({
                        kisiKisiId: '',
                        kompetensi: '',
                        subKompetensi: '',
                        bentukSoal: 'pilihan_ganda_sederhana',
                        soal: '',
                        stimulus: '',
                        opsi: ['', '', '', '', ''],
                        kunciJawaban: '',
                        pembahasan: '',
                        kataKunci: '',
                        gambarUrl: '',
                        gambarCaption: '',
                        gambarPosisi: 'center',
                        gambarUkuran: 'medium'
                      });
                      setTimeout(() => {
                        document.getElementById('manual-question-form')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
                      }, 100);
                    }}
                    className="bg-slate-900 hover:bg-black text-white px-4 py-2 rounded-xl text-xs font-extrabold flex items-center gap-1.5 transition shadow-md cursor-pointer ring-2 ring-slate-400/50"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Tambah Soal Baru (No. #{questions.length + 1})</span>
                  </button>
                </div>
              </div>

              {/* Summary Stats & Filter Toolbar */}
              <div className="pt-3 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="font-bold text-slate-500">Filter Bentuk Soal:</span>
                  {(() => {
                    const countPG = questions.filter(q => q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q))).length;
                    const countMCMA = questions.filter(q => q.bentukSoal === 'mcma').length;
                    const countKategori = questions.filter(q => isKategoriSoal(q)).length;

                    return (
                      <div className="flex flex-wrap items-center gap-1.5">
                        <button
                          onClick={() => setSelectedBentukFilter('all')}
                          className={`px-3 py-1 rounded-lg font-extrabold transition text-xs cursor-pointer ${
                            selectedBentukFilter === 'all'
                              ? 'bg-slate-900 text-white shadow-xs'
                              : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                          }`}
                        >
                          Semua ({questions.length})
                        </button>
                        <button
                          onClick={() => setSelectedBentukFilter('pilihan_ganda_sederhana')}
                          className={`px-3 py-1 rounded-lg font-bold transition text-xs cursor-pointer ${
                            selectedBentukFilter === 'pilihan_ganda_sederhana'
                              ? 'bg-blue-600 text-white shadow-xs'
                              : 'bg-blue-50 text-blue-800 hover:bg-blue-100 border border-blue-200'
                          }`}
                        >
                          PG Sederhana ({countPG})
                        </button>
                        <button
                          onClick={() => setSelectedBentukFilter('mcma')}
                          className={`px-3 py-1 rounded-lg font-bold transition text-xs cursor-pointer ${
                            selectedBentukFilter === 'mcma'
                              ? 'bg-purple-600 text-white shadow-xs'
                              : 'bg-purple-50 text-purple-800 hover:bg-purple-100 border border-purple-200'
                          }`}
                        >
                          PG Kompleks MCMA ({countMCMA})
                        </button>
                        <button
                          onClick={() => setSelectedBentukFilter('kategori')}
                          className={`px-3 py-1 rounded-lg font-bold transition text-xs cursor-pointer ${
                            selectedBentukFilter === 'kategori'
                              ? 'bg-amber-600 text-white shadow-xs'
                              : 'bg-amber-50 text-amber-900 hover:bg-amber-100 border border-amber-200'
                          }`}
                        >
                          PGK Kategori ({countKategori})
                        </button>
                      </div>
                    );
                  })()}
                </div>

                <div className="text-[11px] font-semibold text-slate-500 bg-slate-50 px-3 py-1 rounded-lg border border-slate-200">
                  Total Target Soal: <span className="font-extrabold text-slate-900">{questions.length} / {config.jumlahSoal || questions.length}</span>
                </div>
              </div>
            </div>

            {/* Manual Question Form */}
            {isEditingQuestion && (
              <div 
                id="manual-question-form" 
                className={`bg-white rounded-2xl p-6 no-print transition-all duration-300 shadow-lg ${
                  editingQuestionId 
                    ? 'border-2 border-indigo-600 ring-4 ring-indigo-200/80 bg-gradient-to-b from-indigo-50/40 via-white to-white shadow-indigo-100' 
                    : 'border-2 border-slate-300'
                }`}
              >
                <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-200">
                  <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <Sliders className={`h-4.5 w-4.5 ${editingQuestionId ? 'text-indigo-600 animate-bounce' : 'text-blue-600'}`} />
                    <span>{editingQuestionId ? 'Ubah Butir Soal' : 'Form Manual Pembuatan Soal'}</span>
                  </h3>
                  {editingQuestionId ? (
                    <div className="flex items-center gap-2">
                      <span className="bg-indigo-600 text-white text-[11px] font-extrabold px-3 py-1 rounded-full animate-pulse shadow-sm flex items-center gap-1.5">
                        <Edit className="h-3.5 w-3.5" />
                        <span>Mode Ubah Soal Aktif</span>
                      </span>
                      <button
                        type="button"
                        onClick={() => {
                          setIsEditingQuestion(false);
                          setEditingQuestionId(null);
                        }}
                        className="text-xs text-slate-500 hover:text-rose-600 font-bold underline px-2 py-1 transition cursor-pointer"
                      >
                        Batal
                      </button>
                    </div>
                  ) : (
                    <span className="bg-slate-100 text-slate-600 text-[11px] font-semibold px-2.5 py-0.5 rounded-full border border-slate-200">
                      Mode Soal Baru
                    </span>
                  )}
                </div>
                <form onSubmit={handleSaveQuestionForm} className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-500 mb-1">Hubungkan ke Kisi-Kisi</label>
                      <select
                        value={questionForm.kisiKisiId}
                        onChange={(e) => {
                          const selectedKisi = kisiList.find(k => k.id === e.target.value);
                          setQuestionForm({
                            ...questionForm,
                            kisiKisiId: e.target.value,
                            kompetensi: selectedKisi?.kompetensi || '',
                            subKompetensi: selectedKisi?.subElemenMateri || '',
                            bentukSoal: selectedKisi?.bentukSoal || 'pilihan_ganda_sederhana'
                          });
                        }}
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs"
                      >
                        <option value="">-- Hubungkan Kisi-Kisi (Opsional) --</option>
                        {kisiList.map(k => (
                          <option key={k.id} value={k.id}>No {k.no}: {k.elemenMateri} - {k.subElemenMateri}</option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-500 mb-1">Kompetensi Spesifik</label>
                      <input
                        type="text"
                        value={questionForm.kompetensi}
                        onChange={(e) => setQuestionForm({ ...questionForm, kompetensi: e.target.value })}
                        placeholder="Kompetensi yang diuji"
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Bentuk Soal TKA SMA</label>
                      <select
                        value={questionForm.bentukSoal}
                        onChange={(e) => setQuestionForm({ ...questionForm, bentukSoal: e.target.value as BentukSoal })}
                        className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-900 focus:bg-white focus:ring-2 focus:ring-indigo-500"
                      >
                        <option value="pilihan_ganda_sederhana">Pilihan Ganda Sederhana (1 Jawaban Benar)</option>
                        <option value="mcma">Pilihan Ganda Kompleks - MCMA (Banyak Jawaban)</option>
                        <option value="kategori">Pilihan Ganda Kompleks - Kategori (Tabel Pernyataan Multi-Kategori)</option>
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Teks Stimulus (Paragraf Pengantar/Data/Kasus/Grafik)</label>
                    <textarea
                      rows={2}
                      value={questionForm.stimulus}
                      onChange={(e) => setQuestionForm({ ...questionForm, stimulus: e.target.value })}
                      placeholder="Masukkan stimulus soal (jika ada)..."
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Pertanyaan / Pokok Soal (Wajib)</label>
                    <textarea
                      rows={3}
                      value={questionForm.soal}
                      onChange={(e) => setQuestionForm({ ...questionForm, soal: e.target.value })}
                      placeholder={
                        questionForm.bentukSoal === 'kategori'
                          ? "Tuliskan narasi pertanyaan utama diikuti kalimat instruksi tabel (Contoh: Evaluasilah kesesuaian pernyataan berikut! Tentukan SESUAI atau TIDAK SESUAI)..."
                          : "Masukkan pertanyaan utama..."
                      }
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500 font-medium"
                      required
                    />
                  </div>

                  {/* Options management */}
                  <div className="space-y-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                    <div className="flex justify-between items-center mb-1">
                      <span className="block text-xs font-extrabold text-slate-800">
                        {questionForm.bentukSoal === 'kategori'
                          ? 'Daftar Pernyataan (Pernyataan 1 s.d. 5 yang Direspon Siswa)'
                          : 'Pilihan Jawaban (Opsi A s.d. E)'}
                      </span>
                      <span className="text-[10.5px] text-indigo-700 font-bold">
                        {questionForm.bentukSoal === 'kategori'
                          ? 'Gunakan 2 s.d. 5 butir pernyataan'
                          : 'Isi teks opsi jawaban lengkap'}
                      </span>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {['A', 'B', 'C', 'D', 'E'].map((letter, idx) => (
                        <div key={letter} className="flex items-center gap-2">
                          <span className={`text-xs font-black w-7 h-7 rounded-lg flex items-center justify-center font-mono flex-shrink-0 ${
                            questionForm.bentukSoal === 'kategori' 
                              ? 'bg-amber-100 text-amber-900 border border-amber-300' 
                              : 'bg-slate-200 text-slate-800'
                          }`}>
                            {questionForm.bentukSoal === 'kategori' ? `P${idx + 1}` : letter}
                          </span>
                          <input
                            type="text"
                            value={(questionForm.opsi || [])[idx] || ''}
                            onChange={(e) => handleOpsiChange(idx, e.target.value)}
                            placeholder={
                              questionForm.bentukSoal === 'kategori'
                                ? `Teks Pernyataan ke-${idx + 1} (Contoh: Sosiologi mengkaji fenomena sosial secara empiris)`
                                : `Pilihan ${letter}`
                            }
                            className="flex-1 bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs focus:ring-2 focus:ring-indigo-500"
                          />
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Kunci Jawaban {questionForm.bentukSoal === 'kategori' ? 'Per Pernyataan' : 'Tepat'} (Wajib)
                      </label>
                      <input
                        type="text"
                        value={questionForm.kunciJawaban}
                        onChange={(e) => setQuestionForm({ ...questionForm, kunciJawaban: e.target.value })}
                        placeholder={
                          questionForm.bentukSoal === 'kategori'
                            ? "Contoh: Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, Pernyataan 3: Sesuai"
                            : questionForm.bentukSoal === 'mcma'
                            ? "Contoh: A, C, E"
                            : "Contoh: A"
                        }
                        className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-indigo-900 focus:bg-white focus:ring-2 focus:ring-indigo-500"
                        required
                      />
                      <p className="text-[10px] text-slate-500 mt-1 leading-snug">
                        {questionForm.bentukSoal === 'kategori'
                          ? '💡 Format Kategori: Tuliskan status respon/kategori untuk tiap nomor (misal: Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai, atau Fakta/Opini, dll.).'
                          : questionForm.bentukSoal === 'mcma'
                          ? '💡 Format MCMA: Tuliskan huruf opsi benar dipisahkan koma (misal: A, C, D).'
                          : '💡 Format PG Sederhana: Tuliskan 1 huruf opsi benar (misal: A).'}
                      </p>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-500 mb-1">Kata Kunci / Konsep Utama</label>
                      <input
                        type="text"
                        value={questionForm.kataKunci || ''}
                        onChange={(e) => setQuestionForm({ ...questionForm, kataKunci: e.target.value })}
                        placeholder="Contoh: Persamaan Kuadrat, Fotosintesis"
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-500 mb-1">Pembahasan Terstruktur</label>
                      <textarea
                        rows={2}
                        value={questionForm.pembahasan}
                        onChange={(e) => setQuestionForm({ ...questionForm, pembahasan: e.target.value })}
                        placeholder="Cara kerja penyelesaian ilmiah..."
                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-500 mb-1 flex justify-between items-center">
                      <span>Ilustrasi / Grafik / Infografis Pendukung (Opsional)</span>
                      <span className="text-[10px] text-indigo-600 font-semibold font-sans">Mendukung Link Gambar (http://), Kode SVG Lengkap (&lt;svg&gt;), atau Upload File</span>
                    </label>
                    <textarea
                      rows={2}
                      value={questionForm.gambarUrl || ''}
                      onChange={(e) => setQuestionForm({ ...questionForm, gambarUrl: e.target.value })}
                      placeholder="Contoh: https://images.unsplash.com/photo-1543269865-cbf427effbad?w=500  ATAU kode <svg viewBox='0 0 400 150'>...</svg>"
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-mono"
                    />

                    {/* Direct Image File Uploader & Clear Button */}
                    <div className="mt-2 flex items-center gap-2 flex-wrap">
                      <label className="cursor-pointer bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition border border-slate-250 select-none">
                        <Upload className="h-3.5 w-3.5 text-slate-500" />
                        <span>{isCompressingImage ? "⚡ Mengompresi Gambar..." : "📁 Pilih/Upload Gambar (Auto-Compress Canvas)"}</span>
                        <input 
                          type="file" 
                          accept="image/*" 
                          className="hidden" 
                          disabled={isCompressingImage}
                          onChange={async (e) => {
                            const file = e.target.files?.[0];
                            if (file) {
                              setIsCompressingImage(true);
                              try {
                                const result = await compressImageFile(file);
                                setQuestionForm(prev => ({
                                  ...prev,
                                  gambarUrl: result.dataUrl
                                }));
                                setImageCompressReport({
                                  originalSizeKb: result.originalSizeKb,
                                  compressedSizeKb: result.compressedSizeKb,
                                  savingPercent: result.savingPercent
                                });
                              } catch (err) {
                                console.error("Compression fallback to FileReader:", err);
                                const reader = new FileReader();
                                reader.onloadend = () => {
                                  setQuestionForm(prev => ({ ...prev, gambarUrl: reader.result as string }));
                                };
                                reader.readAsDataURL(file);
                              } finally {
                                setIsCompressingImage(false);
                              }
                            }
                          }} 
                        />
                      </label>
                      {questionForm.gambarUrl && (
                        <button
                          type="button"
                          onClick={() => {
                            setQuestionForm(prev => ({ ...prev, gambarUrl: '', gambarCaption: '' }));
                            setImageCompressReport(null);
                          }}
                          className="bg-rose-50 hover:bg-rose-100 text-rose-600 px-3 py-1.5 rounded-xl text-xs font-semibold flex items-center gap-1 transition border border-rose-100 cursor-pointer"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                          Hapus Gambar
                        </button>
                      )}
                      {questionForm.gambarUrl && !questionForm.gambarUrl.trim().toLowerCase().startsWith('<svg') && (
                        <div className="h-8 w-8 rounded border border-slate-200 overflow-hidden flex items-center justify-center bg-slate-50">
                          <img src={questionForm.gambarUrl} alt="Preview" className="h-full w-full object-cover" />
                        </div>
                      )}
                    </div>

                    {/* Gambar Metadata: Caption, Posisi, Ukuran */}
                    {questionForm.gambarUrl && (
                      <div className="mt-3 bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-2.5">
                        <div className="flex items-center justify-between text-[11px] font-bold text-slate-700">
                          <span className="flex items-center gap-1">
                            <ImageIcon className="h-3.5 w-3.5 text-indigo-600" />
                            Pengaturan Tampilan & Caption Gambar (Optimasi Responsive)
                          </span>
                          {imageCompressReport && (
                            <span className="bg-emerald-100 text-emerald-800 text-[9px] font-bold px-2 py-0.5 rounded-full">
                              ⚡ Terkompresi Canvas: {imageCompressReport.originalSizeKb}KB ➔ {imageCompressReport.compressedSizeKb}KB (-{imageCompressReport.savingPercent}%)
                            </span>
                          )}
                        </div>
                        
                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs">
                          {/* Caption */}
                          <div className="sm:col-span-3">
                            <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Judul / Keterangan Gambar (Caption)</label>
                            <input 
                              type="text"
                              value={questionForm.gambarCaption || ''}
                              onChange={(e) => setQuestionForm({ ...questionForm, gambarCaption: e.target.value })}
                              placeholder="Contoh: Gambar 1.1 Struktur Organisasi Kemdikbudristek"
                              className="w-full bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs font-medium"
                            />
                          </div>

                          {/* Posisi Alignment */}
                          <div>
                            <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Posisi Alignment</label>
                            <select
                              value={questionForm.gambarPosisi || 'center'}
                              onChange={(e) => setQuestionForm({ ...questionForm, gambarPosisi: e.target.value as any })}
                              className="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium"
                            >
                              <option value="left">Rata Kiri (Left)</option>
                              <option value="center">Rata Tengah (Center)</option>
                              <option value="right">Rata Kanan (Right)</option>
                            </select>
                          </div>

                          {/* Ukuran Display */}
                          <div>
                            <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Ukuran Gambar</label>
                            <select
                              value={questionForm.gambarUkuran || 'medium'}
                              onChange={(e) => setQuestionForm({ ...questionForm, gambarUkuran: e.target.value as any })}
                              className="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium"
                            >
                              <option value="small">Kecil (25% / ~160px)</option>
                              <option value="medium">Sedang (50% / ~320px)</option>
                              <option value="large">Besar (75% / ~480px)</option>
                              <option value="full">Lebar Penuh (100%)</option>
                            </select>
                          </div>

                          {/* Quick Preview & Lightbox Trigger */}
                          <div className="flex items-end">
                            <button
                              type="button"
                              onClick={() => {
                                setZoomScale(1);
                                setZoomRotation(0);
                                setActiveZoomImage({ url: questionForm.gambarUrl!, caption: questionForm.gambarCaption });
                              }}
                              className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold px-3 py-1.5 rounded-lg border border-indigo-200 text-xs transition flex items-center justify-center gap-1 cursor-pointer"
                            >
                              <ZoomIn className="h-3.5 w-3.5" />
                              <span>Pratinjau Zoom</span>
                            </button>
                          </div>
                        </div>
                      </div>
                    )}

                    {/* Bank Prompt Super-AI Section */}
                    <div className="mt-4 bg-slate-50 border border-slate-200 rounded-2xl p-4">
                      <div className="flex items-center gap-2 mb-2">
                        <Sparkles className="h-4 w-4 text-indigo-600" />
                        <span className="text-xs font-bold text-slate-700 uppercase tracking-wide">📋 Bank Prompt Siap Pakai (Stimulus, Gambar & Tabel)</span>
                      </div>
                      <p className="text-[11px] text-slate-500 mb-3">
                        Gunakan rekomendasi prompt super-efektif berikut untuk dimasukkan ke AI (seperti Nano Banana, Gemini, atau AI Studio) demi menghasilkan stimulus berkelas HOTS SMA.
                      </p>

                      {/* Tab buttons */}
                      <div className="flex flex-wrap gap-1 border-b border-slate-200 pb-2 mb-3">
                        {[
                          { id: 'ilustrasi', label: '🖼️ Gambar/Ilustrasi', icon: ImageIcon },
                          { id: 'tabel', label: '📊 Data/Tabel', icon: FileSpreadsheet },
                          { id: 'matriks', label: '🔢 Matriks & Rumus', icon: Layers },
                          { id: 'grafik', label: '📈 Grafik/Diagram', icon: Sliders },
                          { id: 'stimulus', label: '📝 Kasus Stimulus', icon: FileText }
                        ].map((tab) => {
                          const Icon = tab.icon;
                          const isActive = activePromptTab === tab.id;
                          return (
                            <button
                              key={tab.id}
                              type="button"
                              onClick={() => setActivePromptTab(tab.id as any)}
                              className={`px-3 py-1 text-xs font-bold rounded-lg transition flex items-center gap-1 ${
                                isActive 
                                  ? 'bg-indigo-600 text-white shadow-sm' 
                                  : 'bg-white hover:bg-slate-100 text-slate-600 border border-slate-200'
                              }`}
                            >
                              <Icon className="h-3.5 w-3.5" />
                              <span>{tab.label}</span>
                            </button>
                          );
                        })}
                      </div>

                      {/* Active Tab Content */}
                      <div className="space-y-3">
                        {activePromptTab === 'ilustrasi' && (
                          <div className="bg-white border border-slate-200 rounded-lg p-3">
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-xs font-bold text-indigo-700">Prompt Vektor Ilustrasi SVG Kreatif</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const text = `Buatkan sebuah ilustrasi ikonik atau diagram vektor sederhana dengan SVG mentah (raw SVG) bertema [Isi Topik, misal: pembelahan sel / simbol gotong royong sosiologis / piramida kasta sosial]. Gunakan paduan warna kekinian (indigo, teal, slate). Desain harus bersih, modern, rata tengah, dengan viewBox='0 0 400 150'. Hanya berikan kode SVG tanpa teks intro atau penjelasan apapun.`;
                                  navigator.clipboard.writeText(text);
                                  setCopiedPromptId('ilustrasi');
                                  setTimeout(() => setCopiedPromptId(null), 2000);
                                }}
                                className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-bold flex items-center gap-1 border border-slate-200"
                              >
                                {copiedPromptId === 'ilustrasi' ? (
                                  <>
                                    <Check className="h-3 w-3 text-emerald-600" />
                                    <span className="text-emerald-600">Tersalin!</span>
                                  </>
                                ) : (
                                  <>
                                    <Copy className="h-3 w-3 text-slate-500" />
                                    <span>Salin Prompt</span>
                                  </>
                                )}
                              </button>
                            </div>
                            <p className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-lg border border-slate-150 font-mono select-all">
                              Buatkan sebuah ilustrasi ikonik atau diagram vektor sederhana dengan SVG mentah (raw SVG) bertema <span className="bg-yellow-100 px-1 font-bold text-slate-800 rounded">[Isi Topik, misal: pembelahan sel / simbol gotong royong sosiologis]</span>. Gunakan paduan warna kekinian (indigo, teal, slate). Desain harus bersih, modern, rata tengah, dengan viewBox='0 0 400 150'. Hanya berikan kode SVG tanpa teks intro atau penjelasan apapun.
                            </p>
                            <span className="text-[10px] text-slate-400 mt-2 block italic">💡 Cara pakai: Salin prompt di atas, ganti bagian kuning dengan topik Anda, paste ke AI, dan tempelkan kode SVG hasilnya langsung ke kolom input gambar di atas!</span>
                          </div>
                        )}

                        {activePromptTab === 'tabel' && (
                          <div className="bg-white border border-slate-200 rounded-lg p-3">
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-xs font-bold text-indigo-700">Prompt Pembuatan Tabel Data HTML</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const text = `Buatkan tabel data statistik/matriks dalam format HTML sederhana yang berisi data perbandingan [Isi Topik, misal: laju inflasi 5 negara / persentase kesenjangan sosial antar wilayah]. Tabel harus memiliki class Tailwind yang minimalis dan elegan, atau styling inline border abu-abu tipis (border-collapse: collapse). Baris header harus kontras dengan latar belakang soft-slate. Tuliskan HANYA kode HTML tabel di dalam blok kode \`\`\`html agar siap saya tempelkan sebagai stimulus.`;
                                  navigator.clipboard.writeText(text);
                                  setCopiedPromptId('tabel');
                                  setTimeout(() => setCopiedPromptId(null), 2000);
                                }}
                                className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-bold flex items-center gap-1 border border-slate-200"
                              >
                                {copiedPromptId === 'tabel' ? (
                                  <>
                                    <Check className="h-3 w-3 text-emerald-600" />
                                    <span className="text-emerald-600">Tersalin!</span>
                                  </>
                                ) : (
                                  <>
                                    <Copy className="h-3 w-3 text-slate-500" />
                                    <span>Salin Prompt</span>
                                  </>
                                )}
                              </button>
                            </div>
                            <p className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-lg border border-slate-150 font-mono select-all">
                              Buatkan tabel data statistik/matriks dalam format HTML sederhana yang berisi data perbandingan <span className="bg-yellow-100 px-1 font-bold text-slate-800 rounded">[Isi Topik, misal: laju inflasi 5 negara / kesenjangan sosial]</span>. Tabel harus memiliki class Tailwind yang minimalis dan elegan, atau styling inline border abu-abu tipis (border-collapse: collapse). Baris header harus kontras dengan latar belakang soft-slate. Tuliskan HANYA kode HTML tabel di dalam blok kode ```html agar siap saya tempelkan sebagai stimulus.
                            </p>
                            <span className="text-[10px] text-slate-400 mt-2 block italic">💡 Cara pakai: Hasil dari AI berupa tabel HTML/Tailwind bisa langsung diletakkan di input Teks Stimulus di atas agar tampil rapi dan presisi di lembar soal!</span>
                          </div>
                        )}

                        {activePromptTab === 'matriks' && (
                          <div className="bg-white border border-slate-200 rounded-lg p-3">
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-xs font-bold text-indigo-700">Prompt Khusus Matriks Matematika & Eksakta</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const text = `Buatkan butir soal HOTS Matematika SMA materi [Isi Topik, misal: Perkalian & Invers Matriks Ordo 2x2 / 3x3 / Determinan / Sistem Persamaan Linear]. WAJIB menyajikan setiap matriks menggunakan format LaTeX standar: $\\begin{pmatrix} a_{11} & a_{12} \\\\ a_{21} & a_{22} \\end{pmatrix}$ atau $\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$ atau determinan $\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$. Pastikan setiap baris dipisahkan '\\\\' dan setiap kolom dipisahkan '&'. Seluruh opsi jawaban A sampai E WAJIB disajikan dalam notasi LaTeX matriks yang lengkap dan terstruktur.`;
                                  navigator.clipboard.writeText(text);
                                  setCopiedPromptId('matriks');
                                  setTimeout(() => setCopiedPromptId(null), 2000);
                                }}
                                className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-bold flex items-center gap-1 border border-slate-200"
                              >
                                {copiedPromptId === 'matriks' ? (
                                  <>
                                    <Check className="h-3 w-3 text-emerald-600" />
                                    <span className="text-emerald-600">Tersalin!</span>
                                  </>
                                ) : (
                                  <>
                                    <Copy className="h-3 w-3 text-slate-500" />
                                    <span>Salin Prompt</span>
                                  </>
                                )}
                              </button>
                            </div>
                            <p className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-lg border border-slate-150 font-mono select-all">
                              Buatkan butir soal HOTS Matematika SMA materi <span className="bg-yellow-100 px-1 font-bold text-slate-800 rounded">[Isi Topik, misal: Perkalian & Invers Matriks Ordo 2x2 / 3x3]</span>. WAJIB menyajikan setiap matriks menggunakan format LaTeX standar: <code className="bg-indigo-50 text-indigo-800 px-1 rounded font-bold">$\begin&#123;pmatrix&#125; a_{11} & a_{12} \\ a_{21} & a_{22} \end&#123;pmatrix&#125;$</code> atau <code className="bg-indigo-50 text-indigo-800 px-1 rounded font-bold">$\begin&#123;bmatrix&#125; a & b \\ c & d \end&#123;bmatrix&#125;$</code>. Pastikan setiap baris dipisahkan '\\\\' dan setiap kolom dipisahkan '&'. Seluruh opsi jawaban A sampai E WAJIB disajikan dalam notasi LaTeX matriks.
                            </p>
                            <span className="text-[10px] text-slate-400 mt-2 block italic">💡 Cara pakai: Prompt ini memaksa AI (Gemini / ChatGPT) menuliskan kolom dan baris matriks secara presisi dengan pemisah kolom '&' dan baris '\\\\', sehingga kolom matriks terbaca sempurna di sistem!</span>
                          </div>
                        )}

                        {activePromptTab === 'grafik' && (
                          <div className="bg-white border border-slate-200 rounded-lg p-3">
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-xs font-bold text-indigo-700">Prompt Grafik / Kurva SVG Presisi</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const text = `Saya ingin membuat butir soal HOTS SMA. Tolong buatkan kode SVG mentah (raw code) untuk grafik/diagram [Isi Topik, misal: kurva permintaan ekonomi / grafik sosiogram / diagram jaring-jaring makanan]. SVG harus: 1. Berwarna modern, bersih, latar belakang transparan. 2. Memiliki sumbu X dan Y dengan label teks yang jelas. 3. Elemen garis/kurva dengan stroke tebal yang estetik. 4. Ukuran viewBox='0 0 500 200'. Tuliskan HANYA kode SVG-nya saja di dalam blok kode \`\`\`xml tanpa penjelasan tambahan.`;
                                  navigator.clipboard.writeText(text);
                                  setCopiedPromptId('grafik');
                                  setTimeout(() => setCopiedPromptId(null), 2000);
                                }}
                                className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-bold flex items-center gap-1 border border-slate-200"
                              >
                                {copiedPromptId === 'grafik' ? (
                                  <>
                                    <Check className="h-3 w-3 text-emerald-600" />
                                    <span className="text-emerald-600">Tersalin!</span>
                                  </>
                                ) : (
                                  <>
                                    <Copy className="h-3 w-3 text-slate-500" />
                                    <span>Salin Prompt</span>
                                  </>
                                )}
                              </button>
                            </div>
                            <p className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-lg border border-slate-150 font-mono select-all">
                              Saya ingin membuat butir soal HOTS SMA. Tolong buatkan kode SVG mentah (raw code) untuk grafik/diagram <span className="bg-yellow-100 px-1 font-bold text-slate-800 rounded">[Isi Topik, misal: kurva permintaan ekonomi / diagram jaring makanan]</span>. SVG harus: 1. Berwarna modern, bersih, latar belakang transparan. 2. Memiliki sumbu X dan Y dengan label teks yang jelas. 3. Elemen garis/kurva dengan stroke tebal yang estetik. 4. Ukuran viewBox='0 0 500 200'. Tuliskan HANYA kode SVG-nya saja di dalam blok kode ```xml tanpa penjelasan tambahan.
                            </p>
                          </div>
                        )}

                        {activePromptTab === 'stimulus' && (
                          <div className="bg-white border border-slate-200 rounded-lg p-3">
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-xs font-bold text-indigo-700">Prompt Penulisan Stimulus HOTS Berkualitas</span>
                              <button
                                type="button"
                                onClick={() => {
                                  const text = `Tuliskan sebuah teks stimulus berkualitas tinggi (HOTS) bertema [Isi Topik, misal: fenomena gentrifikasi perkotaan / perubahan iklim global]. Teks harus berupa studi kasus pendek atau berita ilmiah (150-200 kata), menyajikan konflik/dilema nyata, objektif, ilmiah, dan memicu kemampuan berpikir kritis siswa SMA. Akhiri dengan satu pertanyaan analisis mendalam.`;
                                  navigator.clipboard.writeText(text);
                                  setCopiedPromptId('stimulus');
                                  setTimeout(() => setCopiedPromptId(null), 2000);
                                }}
                                className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-md transition font-bold flex items-center gap-1 border border-slate-200"
                              >
                                {copiedPromptId === 'stimulus' ? (
                                  <>
                                    <Check className="h-3 w-3 text-emerald-600" />
                                    <span className="text-emerald-600">Tersalin!</span>
                                  </>
                                ) : (
                                  <>
                                    <Copy className="h-3 w-3 text-slate-500" />
                                    <span>Salin Prompt</span>
                                  </>
                                )}
                              </button>
                            </div>
                            <p className="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-lg border border-slate-150 font-mono select-all">
                              Tuliskan sebuah teks stimulus berkualitas tinggi (HOTS) bertema <span className="bg-yellow-100 px-1 font-bold text-slate-800 rounded">[Isi Topik, misal: fenomena gentrifikasi perkotaan]</span>. Teks harus berupa studi kasus pendek atau berita ilmiah (150-200 kata), menyajikan konflik/dilema nyata, objektif, ilmiah, dan memicu kemampuan berpikir kritis siswa SMA. Akhiri dengan satu pertanyaan analisis mendalam.
                            </p>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* AI Illustration / Graphic Generator Widget (Nana Banana) */}
                    <div className="mt-2.5 space-y-2">
                      {!isGeneratingIllustration && (
                        <button
                          type="button"
                          onClick={() => {
                            setIsAiIllustratorOpen(!isAiIllustratorOpen);
                            // Auto-fill prompt if empty based on question content
                            if (!aiIllustratorPrompt && questionForm.soal) {
                              const cleanSoal = questionForm.soal.length > 60 ? questionForm.soal.substring(0, 60) + '...' : questionForm.soal;
                              setAiIllustratorPrompt(`Buat ilustrasi diagram/grafik matematika/sains yang sesuai untuk soal: ${cleanSoal}`);
                            }
                          }}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-gradient-to-r from-indigo-50 to-purple-50 hover:from-indigo-100 hover:to-purple-100 text-indigo-700 border border-indigo-100/60 text-[11px] font-bold transition shadow-sm"
                        >
                          <Sparkles className="h-3.5 w-3.5 text-indigo-600 animate-pulse" />
                          {isAiIllustratorOpen ? "Tutup Perancang Gambar AI" : "✨ Rancang Ilustrasi / Grafik via AI Gemini"}
                        </button>
                      )}

                      {isAiIllustratorOpen && !isGeneratingIllustration && (
                        <motion.div
                          initial={{ opacity: 0, y: -8 }}
                          animate={{ opacity: 1, y: 0 }}
                          className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3"
                        >
                          <div className="text-[11px] font-bold text-slate-700 flex items-center justify-between">
                            <span>Instruksi Gambar / Grafik (Bisa Diatur Sesuai Keinginan):</span>
                            <span className="text-[10px] text-indigo-600 font-semibold font-mono">Output Vektor SVG</span>
                          </div>
                          
                          <textarea
                            rows={3}
                            value={aiIllustratorPrompt}
                            onChange={(e) => setAiIllustratorPrompt(e.target.value)}
                            placeholder="Contoh: Grafik fungsi kuadrat y = x^2 - 4x + 3 lengkap dengan label sumbu X, Y dan titik puncak (2,-1)"
                            className="w-full bg-white border border-slate-250 rounded-xl px-3 py-2 text-xs focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                          />

                          {/* Quick Preset Chips */}
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">Inspirasi Cepat:</span>
                            <div className="flex flex-wrap gap-1.5">
                              {[
                                "Grafik Parabola Kuadrat",
                                "Sirkuit Listrik Seri Paralel",
                                "Diagram Venn Himpunan",
                                "Diagram Alir (Flowchart)",
                                "Siklus Biologi (Air/Karbon)",
                                "Bangun Geometri Kubus 3D"
                              ].map((preset) => (
                                <button
                                  key={preset}
                                  type="button"
                                  onClick={() => setAiIllustratorPrompt(`Buatlah ilustrasi ${preset.toLowerCase()} yang indah, presisi, berwarna modern, lengkap dengan label keterangannya.`)}
                                  className="px-2 py-1 text-[10px] bg-white border border-slate-200 text-slate-600 rounded-lg hover:border-indigo-400 hover:text-indigo-600 transition font-medium"
                                >
                                  + {preset}
                                </button>
                              ))}
                            </div>
                          </div>

                          <div className="flex justify-end gap-1.5 pt-1 border-t border-slate-150">
                            <button
                              type="button"
                              onClick={() => setIsAiIllustratorOpen(false)}
                              className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold px-3 py-1.5 rounded-xl text-[11px] transition"
                            >
                              Batal
                            </button>
                            <button
                              type="button"
                              onClick={handleGenerateCustomIllustration}
                              className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-4 py-1.5 rounded-xl text-[11px] transition shadow-sm flex items-center gap-1"
                            >
                              <Sparkles className="h-3 w-3" />
                              Mulai Desain
                            </button>
                          </div>
                        </motion.div>
                      )}

                      {/* Animated Shimmer Loading Area */}
                      {isGeneratingIllustration && (
                        <motion.div
                          initial={{ opacity: 0 }}
                          animate={{ opacity: 1 }}
                          className="bg-slate-900 text-white border border-slate-800 rounded-2xl p-4 space-y-4 shadow-md overflow-hidden relative"
                        >
                          {/* Top shining progress bar line */}
                          <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 animate-pulse" />
                          
                          <div className="flex items-start gap-3">
                            <div className="relative w-10 h-10 flex-shrink-0 flex items-center justify-center">
                              <div className="absolute inset-0 rounded-full border-2 border-slate-800 border-t-indigo-400 animate-spin" />
                              <Sparkles className="h-4 w-4 text-indigo-400 animate-bounce" />
                            </div>
                            
                            <div className="space-y-1 flex-1 min-w-0">
                              <div className="text-[10px] font-extrabold text-indigo-300 uppercase tracking-widest">
                                AI Desainer sedang Bekerja...
                              </div>
                              <div className="text-xs font-bold text-slate-100 flex items-center gap-2">
                                <span className="relative flex h-2 w-2">
                                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                                  <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
                                </span>
                                <span className="italic text-slate-200 font-sans">{aiIllustratorStatus}</span>
                              </div>
                            </div>
                          </div>

                          {/* Animated progress track */}
                          <div className="space-y-1">
                            <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden relative">
                              <motion.div
                                className="bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 h-full rounded-full absolute left-0 top-0"
                                initial={{ width: "3%" }}
                                animate={{ width: ["15%", "45%", "75%", "92%"] }}
                                transition={{ duration: 15, ease: "easeInOut" }}
                              />
                            </div>
                            <div className="text-[10px] text-slate-400 flex justify-between font-medium">
                              <span>Menggambar elemen grafik presisi tinggi</span>
                              <span className="animate-pulse">Mohon tunggu...</span>
                            </div>
                          </div>
                        </motion.div>
                      )}
                    </div>
                  </div>

                  <div className="flex gap-2 justify-end pt-2">
                    <button
                      type="button"
                      onClick={() => {
                        setIsEditingQuestion(false);
                        setEditingQuestionId(null);
                      }}
                      className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold px-4 py-2 rounded-xl text-xs transition"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      disabled={isSavingQuestion}
                      className="bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-400 text-white font-bold px-5 py-2 rounded-xl text-xs transition flex items-center gap-1.5"
                    >
                      {isSavingQuestion ? (
                        <>
                          <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                          <span>Menyimpan...</span>
                        </>
                      ) : (
                        <span>Simpan Butir Soal</span>
                      )}
                    </button>
                  </div>
                </form>
              </div>
            )}

            {/* MENU SETTING CETAK (PRINT SETTINGS) - NEW REQUESTED FEATURE */}
            <div className="bg-slate-900 text-slate-100 border border-slate-800 rounded-2xl shadow-xl p-5 sm:p-6 no-print space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                <div className="flex items-center gap-2.5">
                  <div className="bg-blue-950 p-2 rounded-xl border border-blue-500/30">
                    <Settings className="h-5 w-5 text-blue-400 animate-pulse" />
                  </div>
                  <div>
                    <h3 className="text-sm sm:text-base font-bold text-white flex items-center gap-2">
                      Menu Pengaturan MasterPrint TKA SMA
                    </h3>
                    <p className="text-[11px] text-slate-400">
                      Atur Kop Surat, identitas siswa, tata letak dua kolom, ukuran tulisan, dan cetak lembar kosong siswa atau lembar kunci jawaban.
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsPrintSettingsOpen(!isPrintSettingsOpen)}
                  className="text-xs font-bold px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition"
                >
                  {isPrintSettingsOpen ? "Sembunyikan Menu" : "Tampilkan Menu"}
                </button>
              </div>

              {isPrintSettingsOpen && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  className="grid grid-cols-1 lg:grid-cols-12 gap-5 pt-1"
                >
                  {/* Left Column - School Header info */}
                  <div className="lg:col-span-6 space-y-3">
                    <span className="text-[11px] font-bold text-blue-400 uppercase tracking-widest block border-l-2 border-blue-500 pl-2">
                      Identitas Akademik & Kop Surat
                    </span>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">
                        Dinas / Kementerian (Kop Atas - Bisa beberapa baris)
                      </label>
                      <textarea
                        rows={2}
                        value={printConfig.kopDepartment}
                        onChange={(e) => setPrintConfig({ ...printConfig, kopDepartment: e.target.value })}
                        className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none focus:border-blue-500"
                        placeholder="Contoh:&#10;PEMERINTAH PROVINSI JAWA TIMUR&#10;DINAS PENDIDIKAN"
                      />
                    </div>
                    
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Nama Lembaga / Sekolah</label>
                        <input
                          type="text"
                          value={printConfig.schoolName}
                          onChange={(e) => setPrintConfig({ ...printConfig, schoolName: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none focus:border-blue-500"
                          placeholder="SMA Negeri Nusantara"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Alamat Sekolah</label>
                        <input
                          type="text"
                          value={printConfig.schoolAddress}
                          onChange={(e) => setPrintConfig({ ...printConfig, schoolAddress: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none focus:border-blue-500"
                          placeholder="Jalan Pendidikan Raya No. 45 Nusantara - Telp/Fax: (021) 777-1234 - Website: www.sekolahkita.sch.id"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Nama Ujian / Asesmen</label>
                        <input
                          type="text"
                          value={printConfig.examName}
                          onChange={(e) => setPrintConfig({ ...printConfig, examName: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none focus:border-blue-500"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Mata Pelajaran</label>
                        <input
                          type="text"
                          value={printConfig.subjectName || ''}
                          onChange={(e) => setPrintConfig({ ...printConfig, subjectName: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none focus:border-blue-500"
                          placeholder="Mata Pelajaran"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-3 gap-2">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Tahun Ajaran</label>
                        <input
                          type="text"
                          value={printConfig.academicYear}
                          onChange={(e) => setPrintConfig({ ...printConfig, academicYear: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Semester</label>
                        <select
                          value={printConfig.semester}
                          onChange={(e) => setPrintConfig({ ...printConfig, semester: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        >
                          <option value="Ganjil">Ganjil</option>
                          <option value="Genap">Genap</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Alokasi Waktu</label>
                        <input
                          type="text"
                          value={printConfig.timeAllocation}
                          onChange={(e) => setPrintConfig({ ...printConfig, timeAllocation: e.target.value })}
                          className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Petunjuk Pengerjaan Soal</label>
                      <textarea
                        rows={2}
                        value={printConfig.instructionText}
                        onChange={(e) => setPrintConfig({ ...printConfig, instructionText: e.target.value })}
                        className="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-white focus:outline-none"
                      />
                    </div>

                    {/* LOGO UPLOAD AREA FOR KOP SURAT RESMI */}
                    <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 space-y-3.5">
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block border-b border-slate-800 pb-1.5">
                        Logo Kop Surat Resmi Sekolah
                      </span>
                      
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                        {/* Left Logo */}
                        <div className="space-y-1.5">
                          <label className="block text-[9px] font-bold text-slate-300 uppercase">
                            Logo Kiri (Logo Sekolah / Daerah)
                          </label>
                          
                          {printConfig.schoolLogo ? (
                            <div className="relative border border-slate-800 bg-slate-900 rounded-xl p-2.5 flex flex-col items-center justify-center min-h-[90px]">
                              <img 
                                src={printConfig.schoolLogo} 
                                alt="Logo Kiri" 
                                className="h-14 w-auto object-contain mb-1.5" 
                              />
                              <button
                                type="button"
                                onClick={() => setPrintConfig({ ...printConfig, schoolLogo: '' })}
                                className="text-[9px] font-bold text-rose-400 hover:text-rose-300 bg-slate-950/40 px-2 py-0.5 rounded-md transition"
                              >
                                Hapus Logo
                              </button>
                            </div>
                          ) : (
                            <label className="flex flex-col items-center justify-center border border-dashed border-slate-700 hover:border-slate-500 bg-slate-900/60 rounded-xl p-4 cursor-pointer transition min-h-[90px] text-center">
                              <Upload className="h-4.5 w-4.5 text-slate-500 mb-1" />
                              <span className="text-[9px] font-bold text-slate-300">Pilih Logo Kiri</span>
                              <input 
                                type="file" 
                                accept="image/*" 
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) {
                                    const reader = new FileReader();
                                    reader.onloadend = () => {
                                      setPrintConfig({ ...printConfig, schoolLogo: reader.result as string });
                                    };
                                    reader.readAsDataURL(file);
                                  }
                                }}
                                className="hidden" 
                              />
                            </label>
                          )}
                          <p className="text-[8px] text-slate-400 leading-normal">
                            <b>Rekomendasi ukuran:</b> Dimensi persegi (1:1), minimal 200×200 pixel, format PNG transparan atau JPG.
                          </p>
                        </div>

                        {/* Right Logo */}
                        <div className="space-y-1.5">
                          <label className="block text-[9px] font-bold text-slate-300 uppercase">
                            Logo Kanan (Tut Wuri / Kemenag / Opsional)
                          </label>
                          
                          {printConfig.schoolLogoRight ? (
                            <div className="relative border border-slate-800 bg-slate-900 rounded-xl p-2.5 flex flex-col items-center justify-center min-h-[90px]">
                              <img 
                                src={printConfig.schoolLogoRight} 
                                alt="Logo Kanan" 
                                className="h-14 w-auto object-contain mb-1.5" 
                              />
                              <button
                                type="button"
                                onClick={() => setPrintConfig({ ...printConfig, schoolLogoRight: '' })}
                                className="text-[9px] font-bold text-rose-400 hover:text-rose-300 bg-slate-950/40 px-2 py-0.5 rounded-md transition"
                              >
                                Hapus Logo
                              </button>
                            </div>
                          ) : (
                            <label className="flex flex-col items-center justify-center border border-dashed border-slate-700 hover:border-slate-500 bg-slate-900/60 rounded-xl p-4 cursor-pointer transition min-h-[90px] text-center">
                              <Upload className="h-4.5 w-4.5 text-slate-500 mb-1" />
                              <span className="text-[9px] font-bold text-slate-300">Pilih Logo Kanan</span>
                              <input 
                                type="file" 
                                accept="image/*" 
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) {
                                    const reader = new FileReader();
                                    reader.onloadend = () => {
                                      setPrintConfig({ ...printConfig, schoolLogoRight: reader.result as string });
                                    };
                                    reader.readAsDataURL(file);
                                  }
                                }}
                                className="hidden" 
                              />
                            </label>
                          )}
                          <p className="text-[8px] text-slate-400 leading-normal">
                            <b>Rekomendasi ukuran:</b> Dimensi persegi (1:1), minimal 200×200 pixel, format PNG transparan atau JPG.
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Right Column - Options & Visibility Switches */}
                  <div className="lg:col-span-6 space-y-4">
                    <span className="text-[11px] font-bold text-indigo-400 uppercase tracking-widest block border-l-2 border-indigo-500 pl-2">
                      Pengaturan Tampilan & Format Kertas
                    </span>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2.5 text-xs">
                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-slate-300 hover:text-white transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showHeader}
                          onChange={(e) => setPrintConfig({ ...printConfig, showHeader: e.target.checked })}
                          className="rounded border-slate-750 bg-slate-800 text-blue-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Aktifkan Kop Surat Resmi</span>
                      </label>

                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-slate-300 hover:text-white transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showStudentFields}
                          onChange={(e) => setPrintConfig({ ...printConfig, showStudentFields: e.target.checked })}
                          className="rounded border-slate-750 bg-slate-800 text-blue-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Kolom Identitas Siswa</span>
                      </label>

                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-slate-300 hover:text-white transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showStimulus}
                          onChange={(e) => setPrintConfig({ ...printConfig, showStimulus: e.target.checked })}
                          className="rounded border-slate-750 bg-slate-800 text-blue-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Tampilkan Stimulus Wacana</span>
                      </label>

                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-slate-300 hover:text-white transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showIllustration}
                          onChange={(e) => setPrintConfig({ ...printConfig, showIllustration: e.target.checked })}
                          className="rounded border-slate-750 bg-slate-800 text-blue-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Tampilkan Grafik/Gambar AI</span>
                      </label>

                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-slate-300 hover:text-white transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showCompetencyTag}
                          onChange={(e) => setPrintConfig({ ...printConfig, showCompetencyTag: e.target.checked })}
                          className="rounded border-slate-750 bg-slate-800 text-blue-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Tampilkan Metadata (Kompetensi)</span>
                      </label>

                      <label className="flex items-center gap-2.5 cursor-pointer select-none text-emerald-400 hover:text-emerald-300 font-bold transition">
                        <input
                          type="checkbox"
                          checked={printConfig.showAnswerKey}
                          onChange={(e) => setPrintConfig({ ...printConfig, showAnswerKey: e.target.checked })}
                          className="rounded border-emerald-700 bg-slate-850 text-emerald-500 focus:ring-0 focus:ring-offset-0 h-4 w-4"
                        />
                        <span>Cetak Kunci & Pembahasan</span>
                      </label>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-slate-800/80">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1 flex items-center gap-1">
                          <Type className="h-3 w-3" /> Ukuran Huruf (Font)
                        </label>
                        <select
                          value={printConfig.fontSize}
                          onChange={(e) => setPrintConfig({ ...printConfig, fontSize: e.target.value })}
                          className="w-full bg-slate-800/85 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        >
                          <option value="text-xs">Kecil (Kertas Padat)</option>
                          <option value="text-sm">Sedang (Standar Nasional)</option>
                          <option value="text-base">Besar (Sangat Jelas)</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1 flex items-center gap-1">
                          <Layout className="h-3 w-3" /> Format Tata Letak
                        </label>
                        <select
                          value={printConfig.layoutColumns}
                          onChange={(e) => setPrintConfig({ ...printConfig, layoutColumns: e.target.value })}
                          className="w-full bg-slate-800/85 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        >
                          <option value="1">1 Kolom Penuh</option>
                          <option value="2">2 Kolom (Hemat Kertas)</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1 flex items-center gap-1">
                          <FileText className="h-3 w-3" /> Ukuran Kertas (Page)
                        </label>
                        <select
                          value={printConfig.pageSize}
                          onChange={(e) => setPrintConfig({ ...printConfig, pageSize: e.target.value })}
                          className="w-full bg-slate-800/85 border border-slate-700 rounded-xl px-2 py-1.5 text-xs text-white focus:outline-none"
                        >
                          <option value="A4">A4 (210 x 297 mm)</option>
                          <option value="F4">F4 / Folio (215 x 330 mm)</option>
                        </select>
                      </div>
                    </div>

                    <div className="pt-1 flex flex-col items-end gap-2">
                      <button
                        onClick={handlePrint}
                        className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-2 transition shadow-lg w-full sm:w-auto justify-center"
                      >
                        <Printer className="h-4 w-4" />
                        <span>Mulai Cetak / Simpan ke PDF</span>
                      </button>
                      <button
                        onClick={() => exportQuestionsToWord(questions, printConfig.subjectName || config.mataPelajaran, printConfig.pageSize, printConfig.examName, printConfig.showAnswerKey)}
                        className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-2 transition shadow-lg w-full sm:w-auto justify-center"
                      >
                        <FileText className="h-4 w-4" />
                        <span>Unduh Versi Word (.docx)</span>
                      </button>
                      <button
                        onClick={() => exportQuestionsToExcel(questions, printConfig.subjectName || config.mataPelajaran, printConfig.examName, printConfig.showAnswerKey)}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-2 transition shadow-lg w-full sm:w-auto justify-center"
                      >
                        <FileSpreadsheet className="h-4 w-4" />
                        <span>Unduh Versi Excel (.xlsx)</span>
                      </button>
                      <button
                        onClick={() => exportQuestionsWithImagesZip(questions, printConfig.subjectName || config.mataPelajaran, printConfig.examName)}
                        className="bg-purple-600 hover:bg-purple-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-2 transition shadow-lg w-full sm:w-auto justify-center"
                        title="Unduh file JSON beserta semua folder gambar yang diunggah dalam format ZIP"
                      >
                        <FileCode className="h-4 w-4" />
                        <span>Unduh Paket JSON + Gambar (.zip)</span>
                      </button>
                      <button
                        onClick={() => exportQuestionsToJSON(questions, printConfig.subjectName || config.mataPelajaran, printConfig.examName)}
                        className="bg-amber-600 hover:bg-amber-700 text-white font-bold px-4 py-2 rounded-xl text-xs flex items-center gap-2 transition shadow-lg w-full sm:w-auto justify-center"
                        title="Unduh file data JSON lengkap beserta data gambar base64 embedded"
                      >
                        <Code className="h-4 w-4" />
                        <span>Unduh Data JSON (.json)</span>
                      </button>
                    </div>
                  </div>

                  {/* Full-width System/Reset Data section */}
                  <div className="lg:col-span-12 border-t border-slate-800/80 pt-4 mt-2 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                    <div className="space-y-1">
                      <span className="text-[11px] font-bold text-rose-400 uppercase tracking-widest block border-l-2 border-rose-500 pl-2">
                        Utilitas & Manajemen Data
                      </span>
                      <p className="text-[11px] text-slate-400">
                        Hapus semua butir soal TKA SMA yang tersimpan untuk memudahkan Anda melakukan request pembuatan paket soal yang baru dari awal.
                      </p>
                    </div>
                    <div className="flex-shrink-0 flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={handleClearAllSessionData}
                        className="bg-amber-950/40 hover:bg-amber-900/40 text-amber-300 hover:text-amber-200 border border-amber-800/60 font-bold px-4 py-2.5 rounded-xl text-xs flex items-center gap-2 transition cursor-pointer"
                        title="Kosongkan tampilan kisi-kisi dan soal sekaligus untuk melepaskan beban memori/database"
                      >
                        <Trash2 className="h-4 w-4 text-amber-400" />
                        <span>⚡ Sapu Bersih Sesi (Kosongkan Kisi & Soal)</span>
                      </button>
                      {!showDeleteAllConfirm ? (
                        <button
                          type="button"
                          onClick={() => setShowDeleteAllConfirm(true)}
                          className="bg-rose-950/40 hover:bg-rose-900/40 text-rose-400 hover:text-rose-300 border border-rose-800/60 font-bold px-4 py-2.5 rounded-xl text-xs flex items-center gap-2 transition cursor-pointer"
                        >
                          <Trash2 className="h-4 w-4" />
                          <span>Hapus Semua Soal</span>
                        </button>
                      ) : (
                        <div className="flex flex-col sm:flex-row items-center gap-3 bg-rose-950/20 border border-rose-900/50 p-2.5 rounded-xl">
                          <span className="text-xs font-medium text-rose-300 px-2 text-center sm:text-left">
                            Yakin ingin menghapus {questions.length} soal?
                          </span>
                          <div className="flex items-center gap-2 w-full sm:w-auto">
                            <button
                              type="button"
                              onClick={handleDeleteAllQuestions}
                              className="bg-rose-600 hover:bg-rose-700 text-white font-bold px-3 py-1.5 rounded-lg text-xs transition flex-1 sm:flex-none cursor-pointer"
                            >
                              Ya, Hapus Semua
                            </button>
                            <button
                              type="button"
                              onClick={() => setShowDeleteAllConfirm(false)}
                              className="bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold px-3 py-1.5 rounded-lg text-xs transition flex-1 sm:flex-none cursor-pointer"
                            >
                              Batal
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </motion.div>
              )}
            </div>
          </div>
        )}

        {/* Tab 4 (Disabled/Hidden): Prompt QUIZ INTERAKTIF & Generator App CBT */}
        {false && (
          <div id="quiz-panel" className="space-y-6 animate-fadeIn no-print text-left">
            
            {/* Header Banner */}
            <div className="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 rounded-3xl p-6 sm:p-8 text-white shadow-xl relative overflow-hidden border border-indigo-500/30">
              <div className="absolute top-0 right-0 -mt-10 -mr-10 w-64 h-64 bg-indigo-500/20 rounded-full blur-3xl pointer-events-none" />
              <div className="absolute bottom-0 left-1/3 -mb-10 w-48 h-48 bg-purple-500/20 rounded-full blur-2xl pointer-events-none" />

              <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div className="space-y-2 max-w-3xl">
                  <div className="inline-flex items-center gap-2 bg-indigo-500/20 border border-indigo-400/30 text-indigo-200 text-xs font-bold px-3 py-1 rounded-full backdrop-blur-md">
                    <Sparkles className="h-3.5 w-3.5 text-amber-300" />
                    <span>Interactive Quiz & CBT Prompt Generator EdTech</span>
                  </div>
                  <h2 className="text-2xl sm:text-3xl font-black text-white tracking-tight">
                    Promt QUIZ INTERAKTIF & CBT Generator
                  </h2>
                  <p className="text-sm text-slate-300 leading-relaxed">
                    Sediakan 2 Wadah Prompt AI Khusus: (1) Hasil Draf Prompt Matriks Asesmen / Kisi-Kisi & (2) Prompt Utuh Aplikasi QUIZ INTERAKTIF & CBT Kompleks (10–50 Soal, Login CBT & Anti-Contek).
                  </p>
                </div>

                {/* Sub tab selector */}
                <div className="flex bg-slate-800/90 border border-slate-700/80 p-1.5 rounded-2xl flex-shrink-0 self-start md:self-center shadow-inner">
                  <button
                    onClick={() => setQuizActiveSubTab('prompt')}
                    className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                      quizActiveSubTab === 'prompt'
                        ? 'bg-gradient-to-r from-indigo-600 to-blue-600 text-white shadow-md'
                        : 'text-slate-300 hover:text-white hover:bg-slate-700/50'
                    }`}
                  >
                    <Sparkles className="h-3.5 w-3.5 text-amber-300" />
                    <span>Dua Wadah Prompt AI (Canvas Gemini)</span>
                  </button>
                  <button
                    onClick={() => setQuizActiveSubTab('embed')}
                    className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                      quizActiveSubTab === 'embed'
                        ? 'bg-gradient-to-r from-indigo-600 to-blue-600 text-white shadow-md'
                        : 'text-slate-300 hover:text-white hover:bg-slate-700/50'
                    }`}
                  >
                    <Code className="h-3.5 w-3.5 text-purple-300" />
                    <span>Kode HTML Embed Standalone</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Sub-Tab 1: Dua Wadah Prompt AI */}
            {quizActiveSubTab === 'prompt' && (
              <div className="space-y-8">
                
                {/* WADAH 1: HASIL PROMPT MATRIKS ASESMEN (KISI-KISI) */}
                <div className="bg-white border-2 border-indigo-200 rounded-3xl p-6 sm:p-7 shadow-md space-y-4 relative overflow-hidden">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-indigo-100 pb-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-indigo-600 text-white rounded-2xl shadow-md">
                        <FileText className="h-6 w-6" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="bg-indigo-100 text-indigo-900 font-black text-[11px] px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                            Wadah 1
                          </span>
                          <span className="bg-emerald-100 text-emerald-800 text-[10.5px] font-bold px-2.5 py-0.5 rounded-full flex items-center gap-1">
                            <Sparkles className="h-3 w-3 text-emerald-600" />
                            <span>Terhubung Otomatis Matriks Asesmen</span>
                          </span>
                        </div>
                        <h3 className="font-extrabold text-slate-900 text-lg sm:text-xl tracking-tight mt-0.5">
                          Hasil Prompt Matriks Asesmen & Kisi-Kisi (Megaprompt AI)
                        </h3>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      <button
                        onClick={async () => {
                          const activeText = promptWadah1Text || generatedSoalPrompt || generatedKisiPrompt;
                          if (!activeText) return;
                          await copyToClipboard(activeText);
                          setCopiedWadah1(true);
                          setTimeout(() => setCopiedWadah1(false), 2500);
                          window.open("https://aistudio.google.com/app/prompts/new_chat", "_blank");
                        }}
                        className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white px-4 py-2 rounded-xl text-xs font-black transition shadow-md flex items-center gap-1.5 cursor-pointer"
                        title="Salin Prompt Wadah 1 dan Buka CANVAS Gemini AI"
                      >
                        <Sparkles className="h-3.5 w-3.5 text-amber-300" />
                        <span>⚡ Terapkan & Buka CANVAS Gemini AI ↗</span>
                      </button>

                      <button
                        onClick={async () => {
                          const activeText = promptWadah1Text || generatedSoalPrompt || generatedKisiPrompt;
                          if (!activeText) return;
                          await copyToClipboard(activeText);
                          setCopiedWadah1(true);
                          setTimeout(() => setCopiedWadah1(false), 2500);
                        }}
                        className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                          copiedWadah1
                            ? 'bg-emerald-600 text-white'
                            : 'bg-indigo-600 hover:bg-indigo-700 text-white shadow'
                        }`}
                      >
                        {copiedWadah1 ? <Check className="h-4 w-4" /> : <Copy className="h-4 w-4" />}
                        <span>{copiedWadah1 ? 'Tersalin!' : 'Salin Prompt Wadah 1'}</span>
                      </button>
                    </div>
                  </div>

                  {/* Terminal Box Wadah 1 */}
                  <div className="relative">
                    <textarea
                      readOnly
                      value={
                        promptWadah1Text ||
                        generatedSoalPrompt ||
                        generatedKisiPrompt ||
                        `[BELUM ADA PROMPT MATRIKS ASESMEN]

Draf Megaprompt Matriks Asesmen belum dibuat. Silakan buka menu "Matriks Asesmen (Kisi-Kisi)" untuk menggenerasikan prompt kisi-kisi atau butir soal HOTS secara otomatis.`
                      }
                      className="w-full h-[220px] bg-slate-950 text-indigo-200 p-4 rounded-2xl font-mono text-xs leading-relaxed focus:outline-none border border-slate-800 resize-y selection:bg-indigo-500 selection:text-white"
                    />
                    {!(promptWadah1Text || generatedSoalPrompt || generatedKisiPrompt) && (
                      <div className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm rounded-2xl flex flex-col items-center justify-center p-4 text-center space-y-3 border border-slate-800">
                        <FileText className="h-8 w-8 text-indigo-400" />
                        <p className="text-xs text-slate-300 max-w-md font-medium">
                          Belum ada draf prompt dari Matriks Asesmen. Buka menu Matriks Asesmen untuk menggenerasikan prompt kisi-kisi secara otomatis.
                        </p>
                        <button
                          onClick={() => setActiveTab('kisi')}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow flex items-center gap-1.5 cursor-pointer"
                        >
                          <Zap className="h-4 w-4 text-amber-300" />
                          <span>⚡ Buka Menu Matriks Asesmen (Kisi-Kisi)</span>
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="flex items-center justify-between text-xs text-slate-500 pt-1">
                    <span className="flex items-center gap-1.5 font-medium">
                      <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                      <span>Format Draf Asesmen Standar Pusmendik & TKA HOTS</span>
                    </span>
                    <span className="text-[11px] text-indigo-600 font-bold">
                      Wadah 1: Matriks Asesmen & Kisi-Kisi
                    </span>
                  </div>
                </div>


                {/* WADAH 2: PROMPT UTUH QUIZ INTERAKTIF & CBT KOMPLEKS */}
                <div className="bg-white border-2 border-purple-300 rounded-3xl p-6 sm:p-7 shadow-lg space-y-6 relative overflow-hidden">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-purple-100 pb-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-gradient-to-r from-purple-600 to-indigo-600 text-white rounded-2xl shadow-md">
                        <ShieldAlert className="h-6 w-6" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="bg-purple-100 text-purple-900 font-black text-[11px] px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                            Wadah 2
                          </span>
                          <span className="bg-amber-100 text-amber-900 text-[10.5px] font-extrabold px-2.5 py-0.5 rounded-full flex items-center gap-1">
                            <Lock className="h-3 w-3 text-amber-700" />
                            <span>10–50 Soal + Anti-Contek & Login CBT</span>
                          </span>
                        </div>
                        <h3 className="font-extrabold text-slate-900 text-lg sm:text-xl tracking-tight mt-0.5">
                          Prompt Utuh QUIZ INTERAKTIF & CBT Kompleks (Anti-Contek System)
                        </h3>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      <button
                        onClick={async () => {
                          const activeText = promptWadah2Text || buildPromptWadah2(cbtForm);
                          await copyToClipboard(activeText);
                          setCopiedWadah2(true);
                          setTimeout(() => setCopiedWadah2(false), 2500);
                          window.open("https://aistudio.google.com/app/prompts/new_chat", "_blank");
                        }}
                        className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white px-4 py-2 rounded-xl text-xs font-black transition shadow-md flex items-center gap-1.5 cursor-pointer"
                        title="Salin Prompt Wadah 2 dan Buka CANVAS Gemini AI"
                      >
                        <Sparkles className="h-3.5 w-3.5 text-amber-300" />
                        <span>⚡ Terapkan & Buka CANVAS Gemini AI ↗</span>
                      </button>

                      <button
                        onClick={async () => {
                          const activeText = promptWadah2Text || buildPromptWadah2(cbtForm);
                          await copyToClipboard(activeText);
                          setCopiedWadah2(true);
                          setTimeout(() => setCopiedWadah2(false), 2500);
                        }}
                        className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                          copiedWadah2
                            ? 'bg-emerald-600 text-white'
                            : 'bg-purple-600 hover:bg-purple-700 text-white shadow'
                        }`}
                      >
                        {copiedWadah2 ? <Check className="h-4 w-4" /> : <Copy className="h-4 w-4" />}
                        <span>{copiedWadah2 ? 'Tersalin!' : 'Salin Prompt Wadah 2'}</span>
                      </button>
                    </div>
                  </div>

                  {/* Form Controls for Wadah 2 */}
                  <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-5">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between border-b border-slate-200/80 pb-3 gap-2">
                      <div>
                        <h4 className="font-extrabold text-slate-800 text-sm flex items-center gap-2">
                          <Sliders className="h-4 w-4 text-purple-600" />
                          <span>Pengaturan Parameter Prompt CBT Wadah 2</span>
                        </h4>
                        <span className="text-[11px] text-slate-500 font-medium">
                          Konfigurasi Bentuk Soal, Jumlah Soal (10–50), Konteks Lokal, Stimulus & Kualitas TKA
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          setCbtForm(prev => ({
                            ...prev,
                            mataPelajaran: config.mataPelajaran || prev.mataPelajaran,
                            materiPokok: config.materiPokok || prev.materiPokok,
                            lingkupMateri: config.definisi || prev.lingkupMateri,
                            subMateri: config.subElemenMateri || prev.subMateri,
                            kompetensiYangDiuji: config.kompetensi || prev.kompetensiYangDiuji,
                            batasanCatatan: config.batasanCatatan || prev.batasanCatatan,
                            levelKognitif: config.levelKognitif || prev.levelKognitif,
                            jumlahSoal: config.jumlahSoal || prev.jumlahSoal,
                            konteksLokal: config.konteksLokal && config.konteksLokal.length > 0 ? config.konteksLokal : prev.konteksLokal,
                            stimulusKonten: config.stimulusKonten && config.stimulusKonten.length > 0 ? config.stimulusKonten : prev.stimulusKonten,
                            standarKualitas: config.kualitasChecklist && config.kualitasChecklist.length > 0 ? config.kualitasChecklist : prev.standarKualitas
                          }));
                          setWadah2SuccessMsg('✅ Parameter CBT berhasil disinkronkan dari Generator Utama!');
                          setTimeout(() => setWadah2SuccessMsg(''), 3500);
                        }}
                        className="text-[11px] font-bold text-purple-700 bg-purple-100 hover:bg-purple-200 border border-purple-300 px-3 py-1.5 rounded-xl transition flex items-center gap-1.5 cursor-pointer self-start sm:self-auto"
                      >
                        <RefreshCw className="h-3.5 w-3.5 text-purple-600" />
                        <span>Sinkronkan Parameter dari Config</span>
                      </button>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
                      
                      {/* Mata Pelajaran */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Mata Pelajaran</label>
                        <input
                          type="text"
                          value={cbtForm.mataPelajaran}
                          onChange={(e) => setCbtForm({ ...cbtForm, mataPelajaran: e.target.value })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      {/* Materi Pokok */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Materi Pokok (Elemen)</label>
                        <input
                          type="text"
                          value={cbtForm.materiPokok}
                          onChange={(e) => setCbtForm({ ...cbtForm, materiPokok: e.target.value })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      {/* Sub-Materi */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Sub-Materi / Indikator</label>
                        <input
                          type="text"
                          value={cbtForm.subMateri}
                          onChange={(e) => setCbtForm({ ...cbtForm, subMateri: e.target.value })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      {/* Bentuk Soal */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Bentuk Soal CBT</label>
                        <select
                          value={cbtForm.bentukSoal}
                          onChange={(e) => setCbtForm({ ...cbtForm, bentukSoal: e.target.value })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-slate-800 font-medium focus:outline-none focus:ring-2 focus:ring-purple-500"
                        >
                          <option value="Pilihan Ganda Sederhana (A-E)">PG Sederhana (Pilihan Ganda A-E)</option>
                          <option value="Pilihan Ganda Kompleks (MCMA - Jawaban Lebih Dari Satu)">PG Kompleks (MCMA - Jawaban Lebih Dari Satu)</option>
                          <option value="Pilihan Ganda Kompleks (Kategori / Multi-Kategori Respon Pernyataan)">PG Kompleks (Kategori / Multi-Kategori Respon Pernyataan)</option>
                        </select>
                      </div>

                      {/* Level Kognitif */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Level Kognitif</label>
                        <select
                          value={cbtForm.levelKognitif}
                          onChange={(e) => setCbtForm({ ...cbtForm, levelKognitif: e.target.value })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        >
                          <option value="Penalaran HOTS (Reasoning) (level_3)">L3 - Penalaran HOTS (Reasoning)</option>
                          <option value="Penerapan (Applying) (level_2)">L2 - Penerapan (Applying)</option>
                          <option value="Pemahaman (Knowing) (level_1)">L1 - Pemahaman (Knowing)</option>
                        </select>
                      </div>

                      {/* Durasi Ujian */}
                      <div>
                        <label className="font-bold text-slate-700 block mb-1">Durasi Ujian (Menit)</label>
                        <input
                          type="number"
                          min={10}
                          max={180}
                          value={cbtForm.durasiMenit}
                          onChange={(e) => setCbtForm({ ...cbtForm, durasiMenit: parseInt(e.target.value) || 60 })}
                          className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-slate-800 font-bold focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      {/* Jumlah Soal (10 - 50 Slider & Buttons) */}
                      <div className="md:col-span-2 lg:col-span-3 bg-purple-50/70 border border-purple-200/80 p-3.5 rounded-xl space-y-2">
                        <div className="flex items-center justify-between">
                          <label className="font-extrabold text-purple-950 text-xs flex items-center gap-2">
                            <span>Target Jumlah Soal Ujian (Min 10 - Max 50)</span>
                          </label>
                          <span className="bg-purple-600 text-white font-mono font-black text-xs px-2.5 py-0.5 rounded-lg shadow">
                            {cbtForm.jumlahSoal} Soal
                          </span>
                        </div>
                        <input
                          type="range"
                          min={10}
                          max={50}
                          step={5}
                          value={cbtForm.jumlahSoal}
                          onChange={(e) => setCbtForm({ ...cbtForm, jumlahSoal: parseInt(e.target.value) || 10 })}
                          className="w-full accent-purple-600 cursor-pointer"
                        />
                        <div className="flex items-center justify-between gap-1">
                          {[10, 20, 30, 40, 50].map((num) => (
                            <button
                              key={num}
                              type="button"
                              onClick={() => setCbtForm({ ...cbtForm, jumlahSoal: num })}
                              className={`px-3 py-1 rounded-lg text-[11px] font-bold cursor-pointer transition ${
                                cbtForm.jumlahSoal === num
                                  ? 'bg-purple-600 text-white shadow'
                                  : 'bg-white text-purple-900 border border-purple-200 hover:bg-purple-100'
                              }`}
                            >
                              {num} Soal
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* Kompetensi yang Diuji */}
                      <div className="md:col-span-2 lg:col-span-3">
                        <label className="font-bold text-slate-700 block mb-1">Kompetensi yang Diuji</label>
                        <textarea
                          rows={2}
                          value={cbtForm.kompetensiYangDiuji}
                          onChange={(e) => setCbtForm({ ...cbtForm, kompetensiYangDiuji: e.target.value })}
                          placeholder="Contoh: Peserta didik mampu menganalisis dampak perubahan sosial budaya terhadap kearifan lokal..."
                          className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      {/* Batasan / Catatan Khusus */}
                      <div className="md:col-span-2 lg:col-span-3">
                        <label className="font-bold text-slate-700 block mb-1">Batasan / Catatan Khusus Ujian</label>
                        <textarea
                          rows={2}
                          value={cbtForm.batasanCatatan}
                          onChange={(e) => setCbtForm({ ...cbtForm, batasanCatatan: e.target.value })}
                          placeholder="Contoh: Sajikan stimulus berbasis data aktual dan studi kasus kontekstual di Indonesia..."
                          className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                    </div>

                    {/* 🎭 KONTEKS LOKAL INDONESIA */}
                    <div className="bg-emerald-50/70 border border-emerald-200/90 rounded-2xl p-4 space-y-3">
                      <div className="flex items-center justify-between border-b border-emerald-200/80 pb-2">
                        <h5 className="font-extrabold text-emerald-950 text-xs flex items-center gap-2 uppercase tracking-wider">
                          <span>🎭 KONTEKS LOKAL INDONESIA</span>
                        </h5>
                        <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-md">
                          {cbtForm.konteksLokal?.length || 0} Terpilih
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-2 text-xs">
                        {[
                          '🎭 Budaya Nusantara',
                          '🗺️ Geografis Indonesia',
                          '👥 Kehidupan Sosial',
                          '💰 Ekonomi Rakyat',
                          '⚙️ Teknologi Tradisional',
                          '🏛️ Kearifan Lokal',
                          '🌈 Keragaman Etnis'
                        ].map((item) => {
                          const isSelected = (cbtForm.konteksLokal || []).includes(item);
                          return (
                            <button
                              key={item}
                              type="button"
                              onClick={() => handleToggleCbtContext(item)}
                              className={`px-3 py-1.5 rounded-xl font-bold transition flex items-center gap-1.5 cursor-pointer text-[11px] ${
                                isSelected
                                  ? 'bg-emerald-600 text-white shadow-sm ring-2 ring-emerald-300'
                                  : 'bg-white text-emerald-900 border border-emerald-200 hover:bg-emerald-100'
                              }`}
                            >
                              <span>{item}</span>
                              {isSelected && <Check className="h-3 w-3" />}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* 📖 STIMULUS & PENGEMBANGAN KONTEN */}
                    <div className="bg-sky-50/70 border border-sky-200/90 rounded-2xl p-4 space-y-3">
                      <div className="flex items-center justify-between border-b border-sky-200/80 pb-2">
                        <h5 className="font-extrabold text-sky-950 text-xs flex items-center gap-2 uppercase tracking-wider">
                          <span>📖 STIMULUS & PENGEMBANGAN KONTEN</span>
                        </h5>
                        <span className="text-[10px] font-bold text-sky-800 bg-sky-100 px-2 py-0.5 rounded-md">
                          {cbtForm.stimulusKonten?.length || 0} Terpilih
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-2 text-xs">
                        {[
                          '📖 Teks Bacaan',
                          '🖼️ Gambar/Ilustrasi',
                          '📊 Data/Tabel',
                          '📈 Grafik/Diagram',
                          '🔍 Kasus Nyata',
                          '📚 Cerita Pendek',
                          '📰 Berita/Artikel'
                        ].map((item) => {
                          const isSelected = (cbtForm.stimulusKonten || []).includes(item);
                          return (
                            <button
                              key={item}
                              type="button"
                              onClick={() => handleToggleCbtStimulus(item)}
                              className={`px-3 py-1.5 rounded-xl font-bold transition flex items-center gap-1.5 cursor-pointer text-[11px] ${
                                isSelected
                                  ? 'bg-sky-600 text-white shadow-sm ring-2 ring-sky-300'
                                  : 'bg-white text-sky-900 border border-sky-200 hover:bg-sky-100'
                              }`}
                            >
                              <span>{item}</span>
                              {isSelected && <Check className="h-3 w-3" />}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* 📋 STANDAR KUALITAS SOAL TKA */}
                    <div className="bg-indigo-50/70 border border-indigo-200/90 rounded-2xl p-4 space-y-3">
                      <div className="flex items-center justify-between border-b border-indigo-200/80 pb-2">
                        <h5 className="font-extrabold text-indigo-950 text-xs flex items-center gap-2 uppercase tracking-wider">
                          <span>📋 STANDAR KUALITAS SOAL TKA</span>
                        </h5>
                        <span className="text-[10px] font-bold text-indigo-800 bg-indigo-100 px-2 py-0.5 rounded-md">
                          {cbtForm.standarKualitas?.length || 0} / 12 Terpenuhi
                        </span>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2 text-xs">
                        {[
                          'Validasi Bahasa',
                          'Konstruksi Soal',
                          'Kesesuaian Materi',
                          'Level Kognitif',
                          'Konteks Relevan',
                          'Tidak Bias',
                          'Kejelasan Instruksi',
                          'Kunci Jawaban Tepat',
                          'Distractor Berkualitas',
                          'Sesuai Kurikulum',
                          'Waktu Pengerjaan',
                          'Inklusivitas'
                        ].map((item) => {
                          const isSelected = (cbtForm.standarKualitas || []).includes(item);
                          return (
                            <button
                              key={item}
                              type="button"
                              onClick={() => handleToggleCbtQuality(item)}
                              className={`px-2.5 py-1.5 rounded-xl font-bold transition flex items-center justify-between gap-1 cursor-pointer text-[11px] ${
                                isSelected
                                  ? 'bg-indigo-600 text-white shadow-sm ring-1 ring-indigo-400'
                                  : 'bg-white text-indigo-900 border border-indigo-200 hover:bg-indigo-100'
                              }`}
                            >
                              <span className="truncate">{item}</span>
                              {isSelected ? <Check className="h-3 w-3 flex-shrink-0" /> : <div className="h-3 w-3 rounded-full border border-indigo-300 flex-shrink-0" />}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Authentication CBT Section */}
                    <div className="bg-amber-50/80 border border-amber-200 rounded-xl p-4 space-y-4">
                      <div className="flex items-center justify-between border-b border-amber-200/80 pb-2">
                        <div className="flex items-center gap-2">
                          <Lock className="h-4 w-4 text-amber-700" />
                          <h5 className="font-extrabold text-amber-950 text-xs uppercase tracking-wider">
                            Sistem Otentikasi & Keamanan Role Bertingkat (Admin, Guru, Peserta)
                          </h5>
                        </div>
                        <span className="text-[10px] font-extrabold text-amber-800 bg-amber-100 px-2 py-0.5 rounded-md border border-amber-200">
                          🔐 Multi-Role Secured
                        </span>
                      </div>

                      {/* Role 1: Admin / Proktor */}
                      <div className="bg-white/80 p-3 rounded-xl border border-amber-200 space-y-2">
                        <div className="flex items-center gap-1.5 font-extrabold text-amber-950 text-xs">
                          <Key className="h-3.5 w-3.5 text-purple-600" />
                          <span>1. Kredensial Administrator / Proktor Ujian</span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Username Admin</label>
                            <input
                              type="text"
                              value={cbtForm.adminUsername || 'admin_cbt'}
                              onChange={(e) => setCbtForm({ ...cbtForm, adminUsername: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Password Admin</label>
                            <input
                              type="text"
                              value={cbtForm.adminPassword || 'admin_proktor2026'}
                              onChange={(e) => setCbtForm({ ...cbtForm, adminPassword: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                        </div>
                      </div>

                      {/* Role 2: Guru / Pengampu */}
                      <div className="bg-white/80 p-3 rounded-xl border border-amber-200 space-y-2">
                        <div className="flex items-center gap-1.5 font-extrabold text-amber-950 text-xs">
                          <UserCheck className="h-3.5 w-3.5 text-indigo-600" />
                          <span>2. Kredensial Guru Pengampu / Pembuat Soal</span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Username Guru</label>
                            <input
                              type="text"
                              value={cbtForm.guruUsername || 'guru_sosiologi'}
                              onChange={(e) => setCbtForm({ ...cbtForm, guruUsername: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Password Guru</label>
                            <input
                              type="text"
                              value={cbtForm.guruPassword || 'guru_pass2026'}
                              onChange={(e) => setCbtForm({ ...cbtForm, guruPassword: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                        </div>
                      </div>

                      {/* Role 3: Peserta / Siswa */}
                      <div className="bg-white/80 p-3 rounded-xl border border-amber-200 space-y-2">
                        <div className="flex items-center gap-1.5 font-extrabold text-amber-950 text-xs">
                          <Users className="h-3.5 w-3.5 text-emerald-600" />
                          <span>3. Kredensial Peserta Ujian / Siswa</span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Username Peserta</label>
                            <input
                              type="text"
                              value={cbtForm.usernameCbt}
                              onChange={(e) => setCbtForm({ ...cbtForm, usernameCbt: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                          <div>
                            <label className="font-bold text-amber-900 block mb-1 text-[11px]">Password Peserta</label>
                            <input
                              type="text"
                              value={cbtForm.passwordCbt}
                              onChange={(e) => setCbtForm({ ...cbtForm, passwordCbt: e.target.value })}
                              className="w-full bg-white border border-amber-300 rounded-lg px-3 py-1.5 font-mono text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Anti-Cheat Parameters Badges */}
                    <div className="p-3.5 bg-slate-900 text-slate-200 rounded-xl space-y-2 text-xs border border-slate-800">
                      <div className="flex items-center gap-2">
                        <ShieldAlert className="h-4 w-4 text-rose-400" />
                        <span className="font-bold text-rose-300 uppercase tracking-wider text-[11px]">
                          Proteksi Keamanan Anti-Contek & Anti-Curang Aktif
                        </span>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 text-[11px]">
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Acak Soal & Jawaban A–E</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Focus Guard (Anti Tab-Switch Max 3x)</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Block Copy, Paste, Cut, Right-Click & F12</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Fullscreen Mode Lock Requirement</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Timer Realtime & Auto-Submit Paksa</span>
                        </div>
                        <div className="flex items-center gap-1.5 text-emerald-300 bg-slate-950 px-2.5 py-1 rounded-lg border border-slate-800">
                          <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400 flex-shrink-0" />
                          <span>Portal Login User/Pass Mandiri</span>
                        </div>
                      </div>
                    </div>

                    {/* Prominent Button "Buatkan / Sinkronkan Prompt Quiz CBT Wadah 2" */}
                    <button
                      type="button"
                      onClick={async () => {
                        const promptText = buildPromptWadah2(cbtForm);
                        setPromptWadah2Text(promptText);
                        await copyToClipboard(promptText);
                        setCopiedWadah2(true);
                        setWadah2SuccessMsg(`✨ Prompt Quiz CBT Wadah 2 Berhasil Digenerasikan (${cbtForm.jumlahSoal} Soal) & Tersalin ke Clipboard!`);
                        setTimeout(() => setCopiedWadah2(false), 3000);
                        setTimeout(() => setWadah2SuccessMsg(''), 5000);
                      }}
                      className="w-full bg-gradient-to-r from-purple-600 via-indigo-600 to-blue-600 hover:from-purple-700 hover:via-indigo-700 hover:to-blue-700 text-white font-black text-sm sm:text-base px-5 py-3.5 rounded-xl shadow-lg hover:shadow-purple-500/30 transition-all flex items-center justify-center gap-2.5 cursor-pointer transform hover:-translate-y-0.5 active:translate-y-0"
                    >
                      <Sparkles className="h-5 w-5 text-amber-300 animate-spin" style={{ animationDuration: '4s' }} />
                      <span>Buatkan / Sinkronkan Prompt Quiz CBT Wadah 2</span>
                    </button>

                  </div>

                  {/* Success Toast Banner Wadah 2 */}
                  {wadah2SuccessMsg && (
                    <div className="bg-emerald-600 text-white p-3.5 rounded-2xl flex items-center justify-between text-xs font-bold shadow-md animate-fadeIn border border-emerald-500">
                      <div className="flex items-center gap-2">
                        <CheckCircle2 className="h-5 w-5 text-amber-300 flex-shrink-0" />
                        <span>{wadah2SuccessMsg}</span>
                      </div>
                      <span className="bg-emerald-800/80 px-2.5 py-1 rounded-lg text-[10px] uppercase tracking-wider font-mono">
                        Canvas Ready
                      </span>
                    </div>
                  )}

                  {/* Terminal Box Wadah 2 */}
                  <div className="relative">
                    <textarea
                      readOnly
                      value={promptWadah2Text || buildPromptWadah2(cbtForm)}
                      className="w-full h-[320px] bg-slate-950 text-indigo-200 p-4 rounded-2xl font-mono text-xs leading-relaxed focus:outline-none border border-slate-800 resize-y selection:bg-purple-500 selection:text-white"
                    />
                  </div>

                  <div className="flex items-center justify-between text-xs text-slate-500 pt-1">
                    <span className="flex items-center gap-1.5 font-medium">
                      <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                      <span>Lengkap dengan Instruksi Pembahasan HOTS & Anti-Contek Canvas AI</span>
                    </span>
                    <span className="text-[11px] text-purple-600 font-bold">
                      Wadah 2: Prompt Utuh CBT Kompleks
                    </span>
                  </div>

                </div>

              </div>
            )}

            {/* Sub-Tab 2: Source Code Single File HTML Standalone */}
            {quizActiveSubTab === 'embed' && (
              <div className="bg-white border border-slate-200 rounded-3xl p-6 shadow-sm space-y-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-3">
                  <div>
                    <h3 className="font-extrabold text-slate-900 text-base flex items-center gap-2">
                      <Code className="h-5 w-5 text-purple-600" />
                      <span>Source Code Single File HTML (Canvas / iFrame Embed Ready)</span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      Satu file utuh terintegrasi (HTML + CSS + JS) siap dijalankan langsung di browser atau LMS Canvas
                    </p>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={async () => {
                        const htmlCode = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Interactive Quiz & CBT Generator - @AJISOSIOLOGI</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
    body { background: #f8fafc; color: #1e293b; padding: 20px; min-height: 100vh; display: flex; flex-direction: column; }
    .container { max-width: 900px; margin: 0 auto; width: 100%; }
    .header { background: linear-gradient(135deg, #1e1b4b, #312e81); color: white; padding: 24px; border-radius: 16px; margin-bottom: 24px; text-align: center; }
    .header h1 { font-size: 24px; margin-bottom: 6px; }
    .header p { font-size: 13px; color: #c7d2fe; }
    .card { background: white; border: 1px solid #e2e8f0; border-radius: 16px; padding: 24px; margin-bottom: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .btn-main { background: linear-gradient(135deg, #2563eb, #4f46e5); color: white; border: none; padding: 14px 24px; border-radius: 12px; font-weight: bold; font-size: 15px; cursor: pointer; width: 100%; margin-bottom: 16px; }
    .prompt-box { background: #0f172a; color: #a5b4fc; font-family: monospace; padding: 16px; border-radius: 12px; font-size: 12px; line-height: 1.6; white-space: pre-wrap; width: 100%; min-height: 250px; border: 1px solid #1e293b; }
    .footer { text-align: center; padding: 20px 0; font-size: 12px; color: #64748b; border-top: 1px solid #e2e8f0; margin-top: 40px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Interactive Quiz & CBT Prompt Generator</h1>
      <p>Aplikasi CBT & Pembuat Prompt TKA Sosiologi HOTS</p>
    </div>

    <div class="card">
      <button class="btn-main" onclick="copyPrompt()">✨ Salin Prompt CBT Wadah 2</button>
      <textarea id="promptArea" class="prompt-box" readonly>${(promptWadah2Text || buildPromptWadah2(cbtForm)).replace(/</g, '&lt;').replace(/>/g, '&gt;')}</textarea>
    </div>
  </div>

  <div class="footer">
    <p><strong>© 2026 @AJISOSIOLOGI</strong> - All Rights Reserved.</p>
  </div>

  <script>
    function copyPrompt() {
      var copyText = document.getElementById("promptArea");
      copyText.select();
      navigator.clipboard.writeText(copyText.value);
      alert("Prompt Quiz CBT Berhasil Disalin ke Clipboard!");
    }
  </script>
</body>
</html>`;
                        await copyToClipboard(htmlCode);
                        setCopiedEmbedCode(true);
                        setTimeout(() => setCopiedEmbedCode(false), 2500);
                      }}
                      className={`px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                        copiedEmbedCode ? 'bg-emerald-600 text-white' : 'bg-purple-600 hover:bg-purple-700 text-white shadow'
                      }`}
                    >
                      {copiedEmbedCode ? <Check className="h-4 w-4" /> : <Copy className="h-4 w-4" />}
                      <span>{copiedEmbedCode ? 'Tersalin!' : 'Salin Semua Kode HTML'}</span>
                    </button>
                  </div>
                </div>

                <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4">
                  <p className="text-xs text-indigo-300 font-mono mb-2">
                    // Silakan salin kode HTML di bawah ini untuk dimasukkan langsung ke Canvas / LMS iFrame:
                  </p>
                  <pre className="bg-slate-950 text-slate-300 p-4 rounded-xl text-xs font-mono overflow-x-auto max-h-[350px]">
{`<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Interactive Quiz & CBT Prompt Generator - @AJISOSIOLOGI</title>
</head>
<body>
  ...
  <footer>
    <p>© 2026 @AJISOSIOLOGI</p>
  </footer>
</body>
</html>`}
                  </pre>
                </div>
              </div>
            )}

            {/* Mandatory Copyright Footer */}
            <div className="pt-4 border-t border-slate-200 text-center text-xs text-slate-500 space-y-1">
              <p className="font-extrabold text-slate-700">© 2026 @AJISOSIOLOGI</p>
              <p className="text-[11px] text-slate-400">Interactive Quiz & Prompt Generator Sosiologi / TKA SMA</p>
            </div>

          </div>
        )}

        {/* Tab 5 (Disabled/Hidden): Pembuatan Materi & Panduan */}
        {false && (
          <div id="materi-panel" className="space-y-6 animate-fadeIn no-print">
            
            {/* Main Content Layout for Materi */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              
              {/* Left Column: List of Kisi-Kisi items for materi mapel */}
              <div className="lg:col-span-5 space-y-4">
                <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 space-y-3">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
                    <h4 className="font-bold text-slate-800 text-sm">Pilih Kisi-Kisi Matriks</h4>
                    <div className="flex items-center gap-1.5">
                      {Object.keys(generatedMaterials).length > 0 && (
                        <button
                          onClick={() => exportAllMateriToWord(kisiList, generatedMaterials, config.mataPelajaran, printConfig.pageSize)}
                          className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-100 px-2 px-2.5 py-1 rounded-xl text-[10px] font-bold flex items-center gap-1 transition"
                          title="Unduh semua materi yang telah disusun dalam satu file Word"
                        >
                          <Download className="h-3 w-3" />
                          <span>Unduh Semua Word</span>
                        </button>
                      )}
                      <span className="bg-slate-100 text-slate-600 font-mono text-xs font-bold px-2 py-0.5 rounded-full">
                        {kisiList.length} Baris
                      </span>
                    </div>
                  </div>

                  {kisiList.length === 0 ? (
                    <div className="text-center py-8 text-slate-400 space-y-2">
                      <Layers className="h-8 w-8 mx-auto text-slate-300" />
                      <p className="text-xs">Belum ada kisi-kisi terdaftar. Silakan buat di Tab 1 atau Tab 2 terlebih dahulu!</p>
                    </div>
                  ) : (
                    <div className="space-y-2.5 max-h-[500px] overflow-y-auto pr-1">
                      {kisiList.map((kisi) => {
                        const isSelected = activeMateriKisiId === kisi.id;
                        const isMateriReady = !!generatedMaterials[kisi.id]?.content;
                        const isPromptReady = !!generatedMaterials[kisi.id]?.prompt;
                        const isGenerating = !!generatingMateriIds[kisi.id];

                        return (
                          <div 
                            key={kisi.id}
                            onClick={() => {
                              setActiveMateriKisiId(kisi.id);
                              setIsEditingMateri(false);
                            }}
                            className={`p-3.5 rounded-xl border text-left cursor-pointer transition ${
                              isSelected 
                                ? 'bg-purple-50 border-purple-300 shadow-sm' 
                                : 'bg-slate-50 hover:bg-slate-100/70 border-slate-200'
                            }`}
                          >
                            <div className="flex items-start justify-between gap-2">
                              <span className="font-mono text-[10px] font-extrabold bg-slate-200 text-slate-700 h-5 w-5 rounded-full flex items-center justify-center flex-shrink-0">
                                {kisi.no}
                              </span>
                              <div className="flex-1 min-w-0">
                                <span className="text-xs font-extrabold text-slate-900 block truncate">
                                  {kisi.elemenMateri}
                                </span>
                                <span className="text-[10px] text-slate-500 block truncate mt-0.5">
                                  {kisi.subElemenMateri}
                                </span>
                              </div>
                              
                              <div className="flex-shrink-0 flex gap-1 flex-wrap justify-end max-w-[100px]">
                                {isMateriReady && (
                                  <span className="inline-flex items-center gap-0.5 bg-emerald-100 text-emerald-800 text-[8px] font-bold px-1.5 py-0.5 rounded-full">
                                    Materi
                                  </span>
                                )}
                                {isPromptReady && (
                                  <span className="inline-flex items-center gap-0.5 bg-indigo-100 text-indigo-800 text-[8px] font-bold px-1.5 py-0.5 rounded-full">
                                    Prompt
                                  </span>
                                )}
                                {!isMateriReady && !isPromptReady && (
                                  <span className="inline-flex items-center gap-1 bg-amber-100 text-amber-800 text-[8px] font-bold px-2 py-0.5 rounded-full">
                                    Belum
                                  </span>
                                )}
                              </div>
                            </div>

                            <p className="text-[11px] text-slate-600 mt-2 line-clamp-2 leading-relaxed">
                              {kisi.kompetensi}
                            </p>

                            <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between gap-1.5 flex-wrap">
                              <div className="flex items-center gap-1 flex-wrap">
                                <span className="bg-slate-200 text-slate-700 text-[9px] font-semibold px-1.5 py-0.5 rounded">
                                  {getLevelKognitifLabel(kisi.levelKognitif)}
                                </span>
                                {((activeSubTab === 'materi' && isMateriReady) || (activeSubTab === 'prompt' && isPromptReady)) && (
                                  <>
                                    <button
                                      onClick={async (e) => {
                                        e.stopPropagation();
                                        const text = activeSubTab === 'materi' ? generatedMaterials[kisi.id]?.content : generatedMaterials[kisi.id]?.prompt;
                                        if (text) {
                                          const success = await copyToClipboard(text);
                                          if (success) {
                                            setCopiedMateriKisiId(kisi.id);
                                            setTimeout(() => setCopiedMateriKisiId(null), 2500);
                                          }
                                        }
                                      }}
                                      className="text-[9px] font-bold px-1.5 py-0.5 bg-slate-200/80 hover:bg-slate-300 text-slate-700 rounded transition flex items-center gap-0.5 cursor-pointer"
                                      title="Salin Teks"
                                    >
                                      {copiedMateriKisiId === kisi.id ? <Check className="h-3 w-3 text-emerald-600" /> : <Copy className="h-3 w-3" />}
                                      <span>{copiedMateriKisiId === kisi.id ? 'Tersalin' : 'Salin'}</span>
                                    </button>

                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        handleDeleteMateri(kisi.id);
                                      }}
                                      className="text-[9px] font-bold px-1.5 py-0.5 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded transition flex items-center gap-0.5 cursor-pointer"
                                      title="Hapus"
                                    >
                                      <Trash2 className="h-3 w-3" />
                                      <span>Hapus</span>
                                    </button>
                                  </>
                                )}
                              </div>

                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleGenerateMateri(kisi, activeSubTab);
                                }}
                                disabled={isGenerating}
                                className={`text-[10px] font-black px-2 py-1 rounded transition-all flex items-center gap-1 ${
                                  (activeSubTab === 'materi' ? isMateriReady : isPromptReady)
                                    ? 'bg-slate-200 hover:bg-slate-300 text-slate-700' 
                                    : 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-sm hover:from-purple-700 hover:to-indigo-700'
                                }`}
                              >
                                {isGenerating ? (
                                  <>
                                    <RefreshCw className="h-3 w-3 animate-spin" />
                                    <span>Memproses...</span>
                                  </>
                                ) : (
                                  <>
                                    <Sparkles className="h-3 w-3" />
                                    <span>
                                      {(activeSubTab === 'materi' ? isMateriReady : isPromptReady)
                                        ? `Buat Ulang`
                                        : `Buat ${activeSubTab === 'materi' ? 'Materi' : 'Prompt'} AI`
                                      }
                                    </span>
                                  </>
                                )}
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>

              {/* Right Column: Material Viewer / Workspace */}
              <div className="lg:col-span-7">
                <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-0 overflow-hidden min-h-[500px] flex flex-col justify-between">
                  {(() => {
                    const activeMateri = activeMateriKisiId ? generatedMaterials[activeMateriKisiId] : null;
                    const activeKisi = kisiList.find(k => k.id === activeMateriKisiId);
                    const activeMateriContent = activeMateri ? (activeSubTab === 'materi' ? activeMateri.content : activeMateri.prompt) : '';

                    if (!activeKisi) {
                      return (
                        <div className="flex-1 flex flex-col items-center justify-center text-center p-8 text-slate-400 space-y-3">
                          <FileText className="h-12 w-12 text-slate-300 animate-pulse" />
                          <div>
                            <h5 className="font-bold text-slate-700 text-sm">Pilih atau Buat Prompt Slide & Infografis</h5>
                            <p className="text-xs max-w-sm mt-1">
                              Silakan klik salah satu <b>Kisi-Kisi Matriks</b> di sebelah kiri untuk melihat, merumuskan, mengedit, atau menghapus materi dan prompt NotebookLM.
                            </p>
                          </div>
                        </div>
                      );
                    }

                    return (
                      <div className="flex-1 flex flex-col justify-between">
                        {/* Sub-Tabs Selector */}
                        <div className="flex border-b border-slate-200 bg-slate-50/55">
                          <button
                            onClick={() => {
                              setActiveSubTab('materi');
                              setIsEditingMateri(false);
                            }}
                            className={`flex-1 py-3 px-4 text-center font-bold text-xs flex items-center justify-center gap-2 border-b-2 transition-all ${
                              activeSubTab === 'materi'
                                ? 'border-purple-600 text-purple-700 bg-white shadow-sm font-black'
                                : 'border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-100/40'
                            }`}
                          >
                            <BookOpen className="h-4 w-4 text-purple-500" />
                            <span>1. Ringkasan Materi Ajar</span>
                          </button>
                          <button
                            onClick={() => {
                              setActiveSubTab('prompt');
                              setIsEditingMateri(false);
                            }}
                            className={`flex-1 py-3 px-4 text-center font-bold text-xs flex items-center justify-center gap-2 border-b-2 transition-all ${
                              activeSubTab === 'prompt'
                                ? 'border-purple-600 text-purple-700 bg-white shadow-sm font-black'
                                : 'border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-100/40'
                            }`}
                          >
                            <Sparkles className="h-4 w-4 text-purple-500" />
                            <span>2. Prompt Slide & Infografis (NotebookLM)</span>
                          </button>
                        </div>

                        {isEditingMateri ? (
                          <div className="p-5 space-y-4 flex-1 flex flex-col justify-between">
                            <div className="flex-1 flex flex-col space-y-3">
                              <div className="border-b border-slate-100 pb-3 flex items-center justify-between">
                                <div>
                                  <span className="bg-purple-100 text-purple-800 text-[9px] font-extrabold px-2 py-0.5 rounded-full uppercase">
                                    {activeSubTab === 'materi' ? 'MODE EDIT RINGKASAN MATERI' : 'MODE EDIT PROMPT NOTEBOOKLM'} {activeKisi.no}
                                  </span>
                                  <h3 className="font-extrabold text-slate-900 text-sm mt-1">
                                    {activeKisi.elemenMateri}
                                  </h3>
                                </div>
                                <button
                                  onClick={() => setIsEditingMateri(false)}
                                  className="text-slate-400 hover:text-slate-600 p-1 rounded-full hover:bg-slate-100 transition"
                                >
                                  <X className="h-5 w-5" />
                                </button>
                              </div>

                              <div className="flex-1 flex flex-col">
                                <label className="text-xs font-bold text-slate-700 mb-1 block">
                                  {activeSubTab === 'materi' ? 'Isi Materi Pembelajaran:' : 'Isi Prompt Slide & Infografis:'}
                                </label>
                                <textarea
                                  value={editingMateriContent}
                                  onChange={(e) => setEditingMateriContent(e.target.value)}
                                  placeholder={
                                    activeSubTab === 'materi'
                                      ? "Tulis ringkasan materi pembelajaran sosiologi lengkap di sini..."
                                      : "Tulis mega-prompt untuk NotebookLM & Gemini AI di sini secara lengkap..."
                                  }
                                  className="w-full flex-1 min-h-[350px] bg-slate-50 border border-slate-200 rounded-xl p-4 font-mono text-xs text-slate-800 focus:ring-2 focus:ring-purple-400 focus:border-purple-400 focus:outline-none shadow-inner resize-y"
                                />
                              </div>
                            </div>

                            <div className="pt-3 border-t border-slate-100 flex items-center justify-end gap-2.5">
                              <button
                                onClick={() => setIsEditingMateri(false)}
                                className="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition"
                              >
                                <X className="h-3.5 w-3.5" />
                                <span>Batal</span>
                              </button>

                              <button
                                onClick={() => handleSaveMateri(activeKisi.id, editingMateriContent)}
                                className="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm"
                              >
                                <Save className="h-3.5 w-3.5" />
                                <span>Simpan {activeSubTab === 'materi' ? 'Materi' : 'Prompt'}</span>
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div className="p-5 flex-1 flex flex-col justify-between space-y-4">
                            <div>
                              {/* Header Detail */}
                              <div className="border-b border-slate-100 pb-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                <div>
                                  <div className="flex items-center gap-2">
                                    <span className="bg-purple-100 text-purple-800 text-[9px] font-extrabold px-2 py-0.5 rounded-full uppercase">
                                      {activeSubTab === 'materi' ? 'Ringkasan Materi' : 'Prompt NotebookLM'} {activeKisi.no}
                                    </span>
                                    <span className="text-[10px] text-slate-400 font-mono">
                                      {getLevelKognitifLabel(activeKisi.levelKognitif)}
                                    </span>
                                  </div>
                                  <h3 className="font-extrabold text-slate-900 text-base mt-1">
                                    {activeKisi.elemenMateri}
                                  </h3>
                                  <p className="text-xs text-slate-500 font-medium mt-0.5">
                                    Sub-materi: {activeKisi.subElemenMateri}
                                  </p>
                                </div>

                                <div className="flex flex-wrap gap-2">
                                  {activeMateriContent ? (
                                    <>
                                      <button
                                        onClick={() => {
                                          setIsEditingMateri(true);
                                          setEditingMateriContent(activeMateriContent);
                                        }}
                                        className="bg-purple-50 hover:bg-purple-100 text-purple-700 border border-purple-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                                        title={`Edit ${activeSubTab === 'materi' ? 'materi' : 'prompt'} secara manual`}
                                      >
                                        <Edit className="h-3.5 w-3.5" />
                                        <span>Edit</span>
                                      </button>

                                      <button
                                        onClick={() => handleDeleteMateri(activeKisi.id)}
                                        className="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                                        title={`Hapus ${activeSubTab === 'materi' ? 'materi' : 'prompt'}`}
                                      >
                                        <Trash2 className="h-3.5 w-3.5" />
                                        <span>Hapus</span>
                                      </button>

                                      <button
                                        onClick={async () => {
                                          const success = await copyToClipboard(activeMateriContent);
                                          if (success) {
                                            setCopiedMateriKisiId(activeKisi.id);
                                            setTimeout(() => setCopiedMateriKisiId(null), 2500);
                                            const successMsg = activeSubTab === 'materi'
                                              ? "Ringkasan materi berhasil disalin ke clipboard!"
                                              : "Mega-Prompt berhasil disalin ke clipboard!";
                                            alert(successMsg);
                                          } else {
                                            alert("Gagal menyalin teks secara otomatis. Silakan pilih dan salin teks secara manual.");
                                          }
                                        }}
                                        className="bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer"
                                        title={`Salin ${activeSubTab === 'materi' ? 'ringkasan materi' : 'prompt'}`}
                                      >
                                        {copiedMateriKisiId === activeKisi.id ? (
                                          <>
                                            <Check className="h-3.5 w-3.5 text-emerald-600" />
                                            <span className="text-emerald-700">Tersalin!</span>
                                          </>
                                        ) : (
                                          <>
                                            <Copy className="h-3.5 w-3.5 text-slate-600" />
                                            <span>Salin</span>
                                          </>
                                        )}
                                      </button>

                                      <button
                                        onClick={() => exportMateriToWord(activeKisi, activeMateriContent, config.mataPelajaran, printConfig.pageSize)}
                                        className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                                        title={`Unduh ${activeSubTab === 'materi' ? 'Materi' : 'Prompt'} sebagai file Word (.doc)`}
                                      >
                                        <Download className="h-3.5 w-3.5" />
                                        <span>Word</span>
                                      </button>

                                      <button
                                        onClick={() => handlePrintMateri(activeKisi, activeMateriContent)}
                                        className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                                        title={`Cetak ${activeSubTab === 'materi' ? 'Materi' : 'Prompt'}`}
                                      >
                                        <Printer className="h-3.5 w-3.5" />
                                        <span>Cetak</span>
                                      </button>
                                    </>
                                  ) : (
                                    <div className="flex items-center gap-2">
                                      <input
                                        type="file"
                                        id={`prompt-file-${activeKisi.id}`}
                                        accept=".txt,.md"
                                        onChange={(e) => handleUploadPromptFile(e, activeKisi.id)}
                                        className="hidden"
                                      />
                                      <label
                                        htmlFor={`prompt-file-${activeKisi.id}`}
                                        className="bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition cursor-pointer"
                                        title="Unggah file teks (.txt / .md)"
                                      >
                                        <Upload className="h-3.5 w-3.5" />
                                        <span>Unggah File</span>
                                      </label>
                                    </div>
                                  )}
                                </div>
                              </div>

                              {/* Content Render Area */}
                              <div className="mt-4 bg-slate-50/65 border border-slate-100 rounded-xl p-5 max-h-[500px] overflow-y-auto shadow-inner">
                                {activeMateriContent ? (
                                  <SimpleMarkdown content={activeMateriContent} />
                                ) : (
                                  <div className="text-center py-12 px-4 space-y-4">
                                    <FileText className="h-10 w-10 mx-auto text-slate-300" />
                                    <div className="max-w-md mx-auto space-y-1">
                                      <p className="text-sm font-bold text-slate-700">
                                        {activeSubTab === 'materi' ? 'Ringkasan Materi Belum Tersedia' : 'Prompt Belum Tersedia'}
                                      </p>
                                      <p className="text-xs text-slate-500 leading-relaxed">
                                        {activeSubTab === 'materi'
                                          ? 'Silakan pilih opsi di bawah untuk menyusun Ringkasan Materi Ajar komprehensif bagi sosiologi kelas XII berdasarkan parameter kisi-kisi ini:'
                                          : 'Silakan pilih opsi di bawah untuk membuat Mega-Prompt siap saji yang dioptimalkan untuk menyusun slide & infografis di NotebookLM:'
                                        }
                                      </p>
                                    </div>
                                    <div className="flex flex-col sm:flex-row justify-center items-center gap-2.5 pt-2">
                                      <button
                                        onClick={() => handleGenerateMateri(activeKisi, activeSubTab)}
                                        className="w-full sm:w-auto bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white px-4 py-2 rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 transition shadow-sm"
                                      >
                                        <Sparkles className="h-3.5 w-3.5" />
                                        <span>
                                          {activeSubTab === 'materi' ? 'Buat Materi dengan AI' : 'Buat Prompt dengan AI'}
                                        </span>
                                      </button>

                                      <button
                                        onClick={() => {
                                          setIsEditingMateri(true);
                                          setEditingMateriContent('');
                                        }}
                                        className="w-full sm:w-auto bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 transition"
                                      >
                                        <Plus className="h-3.5 w-3.5" />
                                        <span>Tulis Manual</span>
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>
                            </div>

                            {/* Bottom Actions inside Card */}
                            <div className="pt-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
                              <span className="flex items-center gap-1">
                                <Info className="h-3.5 w-3.5 text-indigo-500" />
                                <span>
                                  {activeSubTab === 'materi' 
                                    ? 'Materi ajar dioptimalkan dengan teori mendalam & studi kasus nyata Indonesia.' 
                                    : 'Prompt dirancang khusus untuk memandu NotebookLM & Gemini AI agar konten interaktif.'
                                  }
                                </span>
                              </span>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              </div>

            </div>

          </div>
        )}

        {/* Tab 5: Jadwal Rencana Pembelajaran TKA XII */}
        {activeTab === 'jadwal' && (
          <div id="jadwal-panel" className="grid grid-cols-1 lg:grid-cols-12 gap-8 animate-fadeIn">
            
            {/* Left Main Section: Schedule Table and Form */}
            <div className="lg:col-span-8 space-y-6">
              
              <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6 space-y-6">
                
                {/* Title & Actions */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
                  <div>
                    <h2 className="text-xl font-bold text-slate-900 flex items-center gap-2">
                      <Calendar className="h-5 w-5 text-indigo-600" />
                      <span>Jadwal Rencana Pembelajaran TKA Kelas XII</span>
                    </h2>
                    <p className="text-xs text-slate-500 mt-1">
                      Distribusi materi mingguan khusus bulan <b>Juli, Agustus, September dan Oktober</b>.
                    </p>
                  </div>
                  
                  {/* Export and Action buttons */}
                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      onClick={handleResetJadwal}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 border border-slate-200 text-slate-600 hover:text-red-600 hover:bg-red-50 hover:border-red-100 rounded-xl text-xs font-bold transition"
                      title="Kosongkan seluruh rencana pembelajaran"
                    >
                      <Trash2 className="h-3.5 w-3.5 text-slate-400 group-hover:text-red-500" />
                      <span>Kosongkan Jadwal</span>
                    </button>
                    
                    <button
                      onClick={() => exportJadwalToExcel(jadwalList, selectedJadwalPresetSubject)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200 rounded-xl text-xs font-bold transition"
                    >
                      <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Excel</span>
                    </button>
                    
                    <button
                      onClick={() => exportJadwalToWord(jadwalList, selectedJadwalPresetSubject, printConfig.pageSize)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 rounded-xl text-xs font-bold transition"
                    >
                      <FileText className="h-3.5 w-3.5 text-blue-600" />
                      <span>Word</span>
                    </button>

                    <button
                      onClick={handleSortJadwal}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
                      title="Urutkan Jadwal Rencana Pembelajaran TKA Kelas XII berdasarkan bulan dan Minggu ke-"
                    >
                      <Sliders className="h-3.5 w-3.5 text-indigo-600" />
                      <span>Urutkan Jadwal Rencana Pembelajaran TKA Kelas XII</span>
                    </button>

                    <button
                      onClick={() => {
                        const printWindow = window.open('', '_blank');
                        if (!printWindow) return;
                        
                        const tableRowsHtml = jadwalList.map(item => `
                          <tr>
                            <td style="border: 1px solid black; padding: 8px; text-align: center; font-weight: bold;">${item.bulan}</td>
                            <td style="border: 1px solid black; padding: 8px; text-align: center;">Minggu ke-${item.mingguKe}</td>
                            <td style="border: 1px solid black; padding: 8px; font-weight: bold;">${item.elemenMateri}</td>
                            <td style="border: 1px solid black; padding: 8px;">${item.subElemenMateri}</td>
                            <td style="border: 1px solid black; padding: 8px; font-style: italic;">${item.kompetensi}</td>
                          </tr>
                        `).join('');

                        const html = `
                          <html>
                          <head>
                            <title>Cetak Jadwal Pembelajaran TKA XII</title>
                            <style>
                              body { font-family: Arial, sans-serif; margin: 30px; }
                              h2 { text-align: center; color: #1e3a8a; }
                              table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                              th { border: 1px solid black; padding: 10px; background-color: #f1f5f9; }
                              td { border: 1px solid black; padding: 8px; }
                            </style>
                          </head>
                          <body>
                            <h2>TABEL JADWAL RENCANA PEMBELAJARAN TKA KELAS XII (Juli, Agustus, September, Oktober)</h2>
                            <p><b>Mata Pelajaran:</b> ${selectedJadwalPresetSubject}</p>
                            <p><b>Periode Pembelajaran:</b> Juli, Agustus, September dan Oktober</p>
                            <p><b>Tanggal Cetak:</b> ${new Date().toLocaleDateString('id-ID')}</p>
                            <table>
                              <thead>
                                <tr>
                                  <th>Bulan</th>
                                  <th>Minggu Ke-</th>
                                  <th>Elemen / Materi</th>
                                  <th>Sub-elemen / Submateri</th>
                                  <th>Kompetensi yang Diuji</th>
                                </tr>
                              </thead>
                              <tbody>
                                ${tableRowsHtml}
                              </tbody>
                            </table>
                            <script>
                              window.onload = function() { window.print(); }
                            </script>
                          </body>
                          </html>
                        `;
                        printWindow.document.write(html);
                        printWindow.document.close();
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200 rounded-xl text-xs font-bold transition"
                    >
                      <Printer className="h-3.5 w-3.5" />
                      <span>Cetak</span>
                    </button>
                  </div>
                </div>

                <AnimatePresence>
                  {jadwalSortNotification && (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between gap-2 text-xs text-emerald-800 font-medium overflow-hidden"
                    >
                      <div className="flex items-center gap-2">
                        <CheckSquare className="h-4 w-4 text-emerald-600 shrink-0" />
                        <span>{jadwalSortNotification}</span>
                      </div>
                      <button 
                        onClick={() => setJadwalSortNotification(null)}
                        className="p-1 hover:bg-emerald-100 rounded-lg text-emerald-600 transition"
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    </motion.div>
                  )}
                </AnimatePresence>

                {/* Info Box */}
                <div className="bg-amber-50 border border-amber-200 rounded-xl p-3.5 text-xs text-amber-800 flex gap-2.5">
                  <Info className="h-4 w-4 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-bold block">💡 Petunjuk Penyusunan:</span>
                    <span className="leading-relaxed text-slate-750">
                      Anda bisa menambah, mengedit secara langsung di baris tabel, atau menghapus rencana pembelajaran mingguan. Gunakan panel <b>"Rujukan Pusmendik Sosiologi"</b> di sebelah kanan untuk langsung menyalin elemen, submateri, dan kompetensi rujukan ke dalam baris tabel dengan sekali klik.
                    </span>
                  </div>
                </div>

                {/* Inline form to Add New Row */}
                {isAddingJadwal ? (
                  <form onSubmit={handleAddJadwal} className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-4">
                    <h3 className="text-xs font-extrabold text-slate-700">Tambah Baris Rencana Pembelajaran Baru</h3>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Bulan</label>
                        <select
                          value={newJadwal.bulan}
                          onChange={(e) => setNewJadwal(prev => ({ ...prev, bulan: e.target.value as any }))}
                          className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-xs text-slate-700 outline-none"
                        >
                          <option value="Juli">Juli</option>
                          <option value="Agustus">Agustus</option>
                          <option value="September">September</option>
                          <option value="Oktober">Oktober</option>
                        </select>
                      </div>
                      
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Minggu Ke-</label>
                        <input
                          type="number"
                          min="1"
                          max="5"
                          value={newJadwal.mingguKe}
                          onChange={(e) => setNewJadwal(prev => ({ ...prev, mingguKe: parseInt(e.target.value) || 1 }))}
                          className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-xs text-slate-700 outline-none"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Elemen / Materi Pokok</label>
                        <input
                          type="text"
                          value={newJadwal.elemenMateri}
                          onChange={(e) => setNewJadwal(prev => ({ ...prev, elemenMateri: e.target.value }))}
                          placeholder="Contoh: Sosiologi sebagai Ilmu"
                          className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-xs text-slate-700 outline-none"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Sub-elemen / Submateri</label>
                        <textarea
                          value={newJadwal.subElemenMateri}
                          onChange={(e) => setNewJadwal(prev => ({ ...prev, subElemenMateri: e.target.value }))}
                          placeholder="Masukkan rincian materi"
                          className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-xs text-slate-700 outline-none h-16 resize-none"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Kompetensi yang Diuji</label>
                        <textarea
                          value={newJadwal.kompetensi}
                          onChange={(e) => setNewJadwal(prev => ({ ...prev, kompetensi: e.target.value }))}
                          placeholder="Masukkan kompetensi rujukan"
                          className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-xs text-slate-700 outline-none h-16 resize-none"
                          required
                        />
                      </div>
                    </div>

                    <div className="flex justify-end gap-2 text-xs">
                      <button
                        type="button"
                        onClick={() => setIsAddingJadwal(false)}
                        className="px-3.5 py-1.5 border border-slate-200 text-slate-600 hover:bg-slate-100 rounded-lg font-bold"
                      >
                        Batal
                      </button>
                      <button
                        type="submit"
                        className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg font-bold"
                      >
                        Simpan Rencana
                      </button>
                    </div>
                  </form>
                ) : (
                  <button
                    onClick={() => setIsAddingJadwal(true)}
                    className="w-full py-2.5 border-2 border-dashed border-slate-200 hover:border-indigo-400 text-slate-500 hover:text-indigo-600 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Tambah Baris Rencana Pembelajaran Baru</span>
                  </button>
                )}

                {/* Main Table */}
                <div className="overflow-x-auto border border-slate-150 rounded-xl">
                  <table className="w-full border-collapse text-left text-xs text-slate-600">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-800 font-extrabold">
                        <th className="px-4 py-3 font-bold text-slate-700 w-28">Bulan</th>
                        <th className="px-4 py-3 font-bold text-slate-700 w-24">Minggu Ke-</th>
                        <th className="px-4 py-3 font-bold text-slate-700 w-44">Elemen / Materi</th>
                        <th className="px-4 py-3 font-bold text-slate-700">Sub-elemen / Submateri</th>
                        <th className="px-4 py-3 font-bold text-slate-700">Kompetensi yang Diuji</th>
                        <th className="px-4 py-3 font-bold text-slate-700 text-center w-28">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {jadwalList.length === 0 ? (
                        <tr>
                          <td colSpan={6} className="text-center py-10 px-6 bg-slate-50/50">
                            <div className="max-w-md mx-auto space-y-3 py-4">
                              <div className="inline-flex items-center justify-center h-12 w-12 rounded-full bg-indigo-50 text-indigo-600 mb-1">
                                <Sparkles className="h-6 w-6 animate-pulse" />
                              </div>
                              <h3 className="text-xs font-extrabold text-slate-800 uppercase tracking-wide">Jadwal Rencana Pembelajaran Kosong</h3>
                              <p className="text-[11px] text-slate-500 leading-relaxed">
                                Anda wajib memilih <strong>Rekomendasi Matriks Asesmen</strong> sesuai dengan Mata Pelajaran Anda pada panel rujukan kanan, lalu gunakan tombol <strong>"Impor Semua Rencana"</strong> atau klik tombol <strong>"⚡ Impor"</strong> untuk menyusun jadwal secara otomatis.
                              </p>
                              <div className="pt-1.5 text-indigo-600 font-extrabold text-[10px] uppercase tracking-wider flex items-center justify-center gap-1">
                                <span>👉 Silakan Pilih Pelajaran & Impor di Panel Sebelah Kanan!</span>
                              </div>
                            </div>
                          </td>
                        </tr>
                      ) : (
                        jadwalList.map((item) => {
                          const isEditing = editingJadwalId === item.id;
                          
                          if (isEditing && editingJadwalData) {
                            return (
                              <tr key={item.id} className="bg-indigo-50/40">
                                <td className="px-3 py-2">
                                  <select
                                    value={editingJadwalData.bulan}
                                    onChange={(e) => setEditingJadwalData(prev => prev ? ({ ...prev, bulan: e.target.value as any }) : null)}
                                    className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded p-1 text-xs outline-none font-bold"
                                  >
                                    <option value="Juli">Juli</option>
                                    <option value="Agustus">Agustus</option>
                                    <option value="September">September</option>
                                    <option value="Oktober">Oktober</option>
                                  </select>
                                </td>
                                <td className="px-3 py-2">
                                  <input
                                    type="number"
                                    min="1"
                                    max="5"
                                    value={editingJadwalData.mingguKe}
                                    onChange={(e) => setEditingJadwalData(prev => prev ? ({ ...prev, mingguKe: parseInt(e.target.value) || 1 }) : null)}
                                    className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded p-1 text-xs outline-none text-center font-bold"
                                  />
                                </td>
                                <td className="px-3 py-2">
                                  <input
                                    type="text"
                                    value={editingJadwalData.elemenMateri}
                                    onChange={(e) => setEditingJadwalData(prev => prev ? ({ ...prev, elemenMateri: e.target.value }) : null)}
                                    className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded p-1 text-xs outline-none font-bold text-indigo-900"
                                  />
                                </td>
                                <td className="px-3 py-2">
                                  <textarea
                                    value={editingJadwalData.subElemenMateri}
                                    onChange={(e) => setEditingJadwalData(prev => prev ? ({ ...prev, subElemenMateri: e.target.value }) : null)}
                                    className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded p-1 text-xs outline-none h-16 resize-none text-slate-700"
                                  />
                                </td>
                                <td className="px-3 py-2">
                                  <textarea
                                    value={editingJadwalData.kompetensi}
                                    onChange={(e) => setEditingJadwalData(prev => prev ? ({ ...prev, kompetensi: e.target.value }) : null)}
                                    className="w-full bg-white border border-slate-200 focus:ring-1 focus:ring-indigo-500 rounded p-1 text-xs outline-none h-16 resize-none italic text-slate-600"
                                  />
                                </td>
                                <td className="px-3 py-2 text-center">
                                  <div className="flex flex-col sm:flex-row items-center justify-center gap-1">
                                    <button
                                      onClick={handleSaveEditJadwal}
                                      className="w-full px-2 py-1 bg-indigo-600 text-white rounded text-[10px] font-bold hover:bg-indigo-700 transition"
                                    >
                                      Simpan
                                    </button>
                                    <button
                                      onClick={() => {
                                        setEditingJadwalId(null);
                                        setEditingJadwalData(null);
                                      }}
                                      className="w-full px-2 py-1 border border-slate-200 text-slate-500 bg-white rounded text-[10px] font-bold hover:bg-slate-100 transition"
                                    >
                                      Batal
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          }

                          return (
                            <tr key={item.id} className="hover:bg-slate-50/50 transition">
                              <td className="px-4 py-3.5 font-bold text-slate-900">{item.bulan}</td>
                              <td className="px-4 py-3.5 text-center">
                                <span className="bg-slate-100 text-slate-700 px-2.5 py-1 rounded-full text-[10px] font-extrabold">
                                  Minggu {item.mingguKe}
                                </span>
                              </td>
                              <td className="px-4 py-3.5 font-bold text-indigo-900">{item.elemenMateri}</td>
                              <td className="px-4 py-3.5 text-slate-700 leading-relaxed max-w-xs">{item.subElemenMateri}</td>
                              <td className="px-4 py-3.5 text-slate-600 leading-relaxed italic max-w-sm">{item.kompetensi}</td>
                              <td className="px-4 py-3.5 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button
                                    onClick={() => handleStartEditJadwal(item)}
                                    className="text-indigo-600 hover:text-indigo-900 font-bold text-[11px] hover:underline"
                                  >
                                    Edit
                                  </button>
                                  <span className="text-slate-200">|</span>
                                  <button
                                    onClick={() => handleDeleteJadwal(item.id)}
                                    className="text-red-500 hover:text-red-700 font-bold text-[11px] hover:underline"
                                  >
                                    Hapus
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

              </div>

            </div>

            {/* Right Side Panel: Rekomendasi Matriks Asesmen (Pusmendik) */}
            <div className="lg:col-span-4 space-y-6">
              
              <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-5 space-y-4 sticky top-16">
                
                <div className="pb-3 border-b border-slate-100 flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-1.5">
                      <Sparkles className="h-4 w-4 text-indigo-600 animate-pulse" />
                      <span>Rekomendasi Matriks Asesmen</span>
                    </h3>
                    <p className="text-[10px] text-slate-500 mt-0.5">
                      Pusmendik standard for {selectedJadwalPresetSubject}
                    </p>
                  </div>
                  <span className="bg-indigo-100 text-indigo-700 text-[10px] font-extrabold px-2 py-0.5 rounded-full">
                    {(selectedJadwalPresetSubject === 'Matematika' 
                      ? PUSMENDIK_MATEMATIKA_PRESETS 
                      : selectedJadwalPresetSubject === 'Bahasa Indonesia' 
                      ? PUSMENDIK_BAHASA_INDONESIA_PRESETS 
                      : selectedJadwalPresetSubject === 'Bahasa Inggris'
                      ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
                      : selectedJadwalPresetSubject === 'Matematika Tingkat Lanjut'
                      ? PUSMENDIK_MATEMATIKA_TL_PRESETS
                      : selectedJadwalPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
                      ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
                      : selectedJadwalPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
                      ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
                      : selectedJadwalPresetSubject === 'Fisika'
                      ? PUSMENDIK_FISIKA_PRESETS
                      : selectedJadwalPresetSubject === 'Kimia'
                      ? PUSMENDIK_KIMIA_PRESETS
                      : selectedJadwalPresetSubject === 'Biologi'
                      ? PUSMENDIK_BIOLOGI_PRESETS
                      : selectedJadwalPresetSubject === 'PPKN'
                      ? PUSMENDIK_PPKN_PRESETS
                      : selectedJadwalPresetSubject === 'Ekonomi'
                      ? PUSMENDIK_EKONOMI_PRESETS
                      : selectedJadwalPresetSubject === 'Geografi'
                      ? PUSMENDIK_GEOGRAFI_PRESETS
                      : selectedJadwalPresetSubject === 'Sosiologi'
                      ? PUSMENDIK_SOSIOLOGI_PRESETS
                      : selectedJadwalPresetSubject === 'Sejarah Tingkat Lanjut'
                      ? PUSMENDIK_SEJARAH_TL_PRESETS
                      : selectedJadwalPresetSubject === 'Antropologi'
                      ? PUSMENDIK_ANTROPOLOGI_PRESETS
                      : selectedJadwalPresetSubject === 'Bahasa Jepang'
                      ? PUSMENDIK_BAHASA_JEPANG_PRESETS
                      : PUSMENDIK_PKK_PRESETS
                    ).length} Rujukan
                  </span>
                </div>

                <p className="text-[11px] text-slate-500 leading-relaxed">
                  Pilih mata pelajaran di bawah ini, lalu salin per materi atau impor sekaligus untuk menyusun jadwal secara otomatis.
                </p>

                {/* Subject Buttons Grid */}
                <div className="space-y-1.5">
                  <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider">Mata Pelajaran:</label>
                  {userRole === 'admin' ? (
                    <div className="grid grid-cols-2 gap-1.5 max-h-44 overflow-y-auto p-1.5 bg-slate-50 border border-slate-100 rounded-xl">
                      {[
                        { id: 'Matematika', label: '📐 Matematika' },
                        { id: 'Bahasa Indonesia', label: '🇮🇩 B. Indonesia' },
                        { id: 'Bahasa Inggris', label: '🇬🇧 B. Inggris' },
                        { id: 'Matematika Tingkat Lanjut', label: '🚀 Mat Lanjut' },
                        { id: 'Bahasa Indonesia Tingkat Lanjut', label: '✍️ Indo Lanjut' },
                        { id: 'Bahasa Inggris Tingkat Lanjut', label: '🗣️ Inggris Lanjut' },
                        { id: 'Fisika', label: '⚛️ Fisika' },
                        { id: 'Kimia', label: '🧪 Kimia' },
                        { id: 'Biologi', label: '🧬 Biologi' },
                        { id: 'PPKN', label: '🗳️ PPKN' },
                        { id: 'Ekonomi', label: '💰 Ekonomi' },
                        { id: 'Geografi', label: '🌍 Geografi' },
                        { id: 'Sosiologi', label: '👥 Sosiologi' },
                        { id: 'Sejarah Tingkat Lanjut', label: '📜 Sejarah Lanjut' },
                        { id: 'Antropologi', label: '🗿 Antropologi' },
                        { id: 'Bahasa Jepang', label: '🎌 B. Jepang' },
                        { id: 'Produk Kreatif dan Kewirausahaan', label: '💼 Kewirausahaan' }
                      ].map((subj) => (
                        <button
                          key={subj.id}
                          onClick={() => setSelectedJadwalPresetSubject(subj.id as any)}
                          className={`px-2 py-1.5 rounded-lg text-[10px] font-bold text-left transition truncate ${
                            selectedJadwalPresetSubject === subj.id
                              ? 'bg-indigo-600 text-white shadow-sm'
                              : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-800'
                          }`}
                        >
                          {subj.label}
                        </button>
                      ))}
                    </div>
                  ) : (
                    <div className="p-2 bg-amber-50 border border-amber-200 rounded-xl flex items-center gap-2 text-xs font-bold text-amber-800">
                      <Lock className="h-4 w-4 text-amber-600 flex-shrink-0" />
                      <span>Ampuan: <b>{config.mataPelajaran || selectedJadwalPresetSubject}</b></span>
                    </div>
                  )}
                </div>

                {/* Import All Button */}
                <button
                  onClick={handleImportAllJadwalPresets}
                  className="w-full bg-gradient-to-r from-indigo-600 to-blue-600 hover:from-indigo-700 hover:to-blue-700 text-white font-bold py-2 px-3 rounded-xl text-xs flex items-center justify-center gap-1.5 transition shadow-sm"
                >
                  <Plus className="h-3.5 w-3.5" />
                  <span>Impor Semua {(selectedJadwalPresetSubject === 'Matematika' 
                    ? PUSMENDIK_MATEMATIKA_PRESETS 
                    : selectedJadwalPresetSubject === 'Bahasa Indonesia' 
                    ? PUSMENDIK_BAHASA_INDONESIA_PRESETS 
                    : selectedJadwalPresetSubject === 'Bahasa Inggris'
                    ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
                    : selectedJadwalPresetSubject === 'Matematika Tingkat Lanjut'
                    ? PUSMENDIK_MATEMATIKA_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
                    ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
                    ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Fisika'
                    ? PUSMENDIK_FISIKA_PRESETS
                    : selectedJadwalPresetSubject === 'Kimia'
                    ? PUSMENDIK_KIMIA_PRESETS
                    : selectedJadwalPresetSubject === 'Biologi'
                    ? PUSMENDIK_BIOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'PPKN'
                    ? PUSMENDIK_PPKN_PRESETS
                    : selectedJadwalPresetSubject === 'Ekonomi'
                    ? PUSMENDIK_EKONOMI_PRESETS
                    : selectedJadwalPresetSubject === 'Geografi'
                    ? PUSMENDIK_GEOGRAFI_PRESETS
                    : selectedJadwalPresetSubject === 'Sosiologi'
                    ? PUSMENDIK_SOSIOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'Sejarah Tingkat Lanjut'
                    ? PUSMENDIK_SEJARAH_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Antropologi'
                    ? PUSMENDIK_ANTROPOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Jepang'
                    ? PUSMENDIK_BAHASA_JEPANG_PRESETS
                    : PUSMENDIK_PKK_PRESETS
                  ).length} Rencana</span>
                </button>

                {/* List of references */}
                <div className="space-y-2.5 max-h-[380px] overflow-y-auto pr-1">
                  {(selectedJadwalPresetSubject === 'Matematika' 
                    ? PUSMENDIK_MATEMATIKA_PRESETS 
                    : selectedJadwalPresetSubject === 'Bahasa Indonesia' 
                    ? PUSMENDIK_BAHASA_INDONESIA_PRESETS 
                    : selectedJadwalPresetSubject === 'Bahasa Inggris'
                    ? PUSMENDIK_BAHASA_INGGRIS_PRESETS
                    : selectedJadwalPresetSubject === 'Matematika Tingkat Lanjut'
                    ? PUSMENDIK_MATEMATIKA_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Indonesia Tingkat Lanjut'
                    ? PUSMENDIK_BAHASA_INDONESIA_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Inggris Tingkat Lanjut'
                    ? PUSMENDIK_BAHASA_INGGRIS_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Fisika'
                    ? PUSMENDIK_FISIKA_PRESETS
                    : selectedJadwalPresetSubject === 'Kimia'
                    ? PUSMENDIK_KIMIA_PRESETS
                    : selectedJadwalPresetSubject === 'Biologi'
                    ? PUSMENDIK_BIOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'PPKN'
                    ? PUSMENDIK_PPKN_PRESETS
                    : selectedJadwalPresetSubject === 'Ekonomi'
                    ? PUSMENDIK_EKONOMI_PRESETS
                    : selectedJadwalPresetSubject === 'Geografi'
                    ? PUSMENDIK_GEOGRAFI_PRESETS
                    : selectedJadwalPresetSubject === 'Sosiologi'
                    ? PUSMENDIK_SOSIOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'Sejarah Tingkat Lanjut'
                    ? PUSMENDIK_SEJARAH_TL_PRESETS
                    : selectedJadwalPresetSubject === 'Antropologi'
                    ? PUSMENDIK_ANTROPOLOGI_PRESETS
                    : selectedJadwalPresetSubject === 'Bahasa Jepang'
                    ? PUSMENDIK_BAHASA_JEPANG_PRESETS
                    : PUSMENDIK_PKK_PRESETS
                  ).map((preset, idx) => (
                    <div 
                      key={idx}
                      className="p-3 bg-slate-50 hover:bg-indigo-50/50 border border-slate-200 hover:border-indigo-200 rounded-xl transition text-[11px] space-y-2"
                    >
                      <div className="flex justify-between items-start gap-1">
                        <span className="font-extrabold text-indigo-950 block">
                          {preset.elemenMateri}
                        </span>
                        
                        <div className="flex gap-1">
                          <button
                            onClick={() => {
                              // Auto-fill into new row form
                              setNewJadwal({
                                bulan: 'Juli',
                                mingguKe: 1,
                                elemenMateri: preset.elemenMateri,
                                subElemenMateri: preset.subElemenMateri,
                                kompetensi: preset.kompetensi
                              });
                              setIsAddingJadwal(true);
                              
                              // Scroll smoothly to form
                              const element = document.getElementById('jadwal-panel');
                              if (element) {
                                element.scrollIntoView({ behavior: 'smooth' });
                              }
                            }}
                            className="bg-slate-200 text-slate-800 font-bold px-1.5 py-0.5 rounded text-[10px] hover:bg-slate-300 transition shrink-0"
                            title="Salin ke Form Tambah Baru"
                          >
                            + Form
                          </button>
                          
                          <button
                            onClick={() => {
                              // Direct append to list
                              const item: JadwalItem = {
                                id: `jadwal-preset-${Date.now()}-${idx}`,
                                bulan: 'Juli',
                                mingguKe: 1,
                                elemenMateri: preset.elemenMateri,
                                subElemenMateri: preset.subElemenMateri,
                                kompetensi: preset.kompetensi
                              };
                              setJadwalList(prev => [...prev, item]);
                            }}
                            className="bg-indigo-600 text-white font-bold px-1.5 py-0.5 rounded text-[10px] hover:bg-indigo-700 transition shrink-0"
                            title="Impor langsung ke baris jadwal"
                          >
                            ⚡ Impor
                          </button>
                        </div>
                      </div>

                      <div className="space-y-1 text-slate-600 leading-relaxed">
                        <p><b>Sub-materi:</b> {preset.subElemenMateri}</p>
                        <p className="italic text-slate-500"><b>Kompetensi:</b> {preset.kompetensi}</p>
                      </div>
                    </div>
                  ))}
                </div>
              
              </div>

            </div>

          </div>
        )}

        {/* Tab 5: Manajemen Pengguna */}
        {activeTab === 'users' && userRole === 'admin' && (
          <div id="users-panel" className="space-y-8 animate-fadeIn no-print">
            {/* Top Banner: Supabase Cloud Database Configuration */}
            <section className="bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 text-white rounded-3xl p-6 sm:p-8 shadow-xl border border-emerald-500/30 space-y-5">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-emerald-800/40 pb-5">
                <div className="flex items-center gap-3">
                  <div className="p-3 bg-emerald-500/20 border border-emerald-400/30 rounded-2xl text-emerald-300 shrink-0">
                    <Database className="h-6 w-6" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <h2 className="text-lg sm:text-xl font-extrabold tracking-tight">Database Pengguna: Supabase Cloud (PostgreSQL)</h2>
                      <span className={`text-[10px] font-black px-2.5 py-0.5 rounded-full border uppercase tracking-wider ${
                        supabaseUrlInput.trim() && supabaseKeyInput.trim()
                          ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40'
                          : 'bg-indigo-500/20 text-indigo-300 border-indigo-500/40'
                      }`}>
                        {supabaseUrlInput.trim() && supabaseKeyInput.trim() ? '🟢 Supabase PostgreSQL Terhubung' : '🔵 Cloud Database Aktif'}
                      </span>
                    </div>
                    <p className="text-xs text-slate-300 mt-1">
                      Data akun guru dan password tersimpan aman dengan Row Level Security (RLS) & enkripsi cloud di <a href="https://supabase.com" target="_blank" rel="noreferrer" className="text-emerald-300 font-bold underline hover:text-emerald-200">Supabase.com</a> untuk domain <code className="text-emerald-300 font-mono">mastertkasma.my.id</code>.
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2 flex-wrap shrink-0">
                  <button
                    type="button"
                    onClick={() => setShowSupabaseSqlModal(true)}
                    className="bg-emerald-500 hover:bg-emerald-400 text-emerald-950 font-black text-xs px-3.5 py-2.5 rounded-xl shadow transition flex items-center gap-1.5"
                    title="Buka panduan & salin kode SQL untuk membuat tabel users di Supabase"
                  >
                    <FileCode className="h-4 w-4" />
                    <span>📖 Setup SQL Supabase</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => downloadUserExcelTemplate()}
                    className="bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-bold text-xs px-3.5 py-2.5 rounded-xl shadow transition flex items-center gap-1.5"
                    title="Unduh contoh template file Excel untuk penambahan massal pengguna"
                  >
                    <Download className="h-4 w-4 text-slate-400" />
                    <span>Template Excel</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => exportUsersToExcel(usersList)}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white font-extrabold text-xs px-4 py-2.5 rounded-xl shadow transition flex items-center gap-1.5"
                    title="Ekspor seluruh daftar pengguna yang terdaftar ke file Excel"
                  >
                    <Download className="h-4 w-4" />
                    <span>Ekspor (.xlsx)</span>
                  </button>
                </div>
              </div>

              {/* Supabase Configuration Inputs */}
              <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-end pt-1">
                <div className="md:col-span-5 space-y-1">
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-emerald-200">
                    Project URL Supabase
                  </label>
                  <input
                    type="url"
                    placeholder="https://xyzcompany.supabase.co"
                    value={supabaseUrlInput}
                    onChange={(e) => {
                      setSupabaseUrlInput(e.target.value);
                      setSupabaseTestMsg(null);
                    }}
                    className="w-full bg-slate-900/90 border border-emerald-700/60 focus:border-emerald-400 focus:ring-1 focus:ring-emerald-400 rounded-xl px-3.5 py-2 text-xs text-emerald-100 placeholder-slate-500 font-mono"
                  />
                </div>

                <div className="md:col-span-4 space-y-1">
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-emerald-200">
                    Anon / Public Key Supabase
                  </label>
                  <input
                    type="password"
                    placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
                    value={supabaseKeyInput}
                    onChange={(e) => {
                      setSupabaseKeyInput(e.target.value);
                      setSupabaseTestMsg(null);
                    }}
                    className="w-full bg-slate-900/90 border border-emerald-700/60 focus:border-emerald-400 focus:ring-1 focus:ring-emerald-400 rounded-xl px-3.5 py-2 text-xs text-emerald-100 placeholder-slate-500 font-mono"
                  />
                </div>

                <div className="md:col-span-3 flex items-center gap-2">
                  <button
                    type="button"
                    disabled={isTestingSupabase || !supabaseUrlInput.trim() || !supabaseKeyInput.trim()}
                    onClick={async () => {
                      setIsTestingSupabase(true);
                      setSupabaseTestMsg(null);
                      const res = await testSupabaseConnection(supabaseUrlInput.trim(), supabaseKeyInput.trim());
                      setSupabaseTestMsg({
                        success: res.success,
                        text: res.message
                      });
                      setIsTestingSupabase(false);
                    }}
                    className="flex-1 bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-slate-200 font-bold text-xs py-2 px-2.5 rounded-xl border border-slate-700 transition flex items-center justify-center gap-1.5"
                  >
                    {isTestingSupabase ? <RefreshCw className="h-3.5 w-3.5 animate-spin" /> : <Zap className="h-3.5 w-3.5 text-amber-400" />}
                    <span>Tes Koneksi</span>
                  </button>

                  <button
                    type="button"
                    onClick={async () => {
                      await saveStoredSupabaseConfig(supabaseUrlInput.trim(), supabaseKeyInput.trim());
                      setSupabaseTestMsg({
                        success: true,
                        text: supabaseUrlInput.trim() && supabaseKeyInput.trim()
                          ? "Kredensial Supabase berhasil disimpan dan disinkronkan ke semua perangkat! Memuat data pengguna dari Supabase..."
                          : "Kredensial Supabase dikosongkan. Sistem menggunakan Cloud Database bawaan."
                      });
                      const users = await fetchUsersFromCloud();
                      setUsersList(users);
                    }}
                    className="flex-1 bg-emerald-600 hover:bg-emerald-500 text-white font-extrabold text-xs py-2 px-2.5 rounded-xl shadow transition flex items-center justify-center gap-1.5"
                  >
                    <Save className="h-3.5 w-3.5" />
                    <span>Simpan</span>
                  </button>
                </div>
              </div>

              {supabaseTestMsg && (
                <div className={`p-3 rounded-xl text-xs flex items-center gap-2 font-medium ${
                  supabaseTestMsg.success
                    ? 'bg-emerald-500/20 border border-emerald-500/40 text-emerald-200'
                    : 'bg-rose-500/20 border border-rose-500/40 text-rose-200'
                }`}>
                  {supabaseTestMsg.success ? <Check className="h-4 w-4 shrink-0 text-emerald-400" /> : <AlertCircle className="h-4 w-4 shrink-0 text-rose-400" />}
                  <span>{supabaseTestMsg.text}</span>
                </div>
              )}
            </section>

            {/* Grid 2 Column: Add Form & Users List */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
              {/* Left Col: Add User Form */}
              <section className="lg:col-span-5 bg-white border border-slate-200 rounded-2xl shadow-sm p-6 space-y-5">
                <div className="flex items-center gap-2 pb-4 border-b border-slate-100">
                  <UserPlus className="h-5 w-5 text-amber-500" />
                  <h2 className="text-lg font-bold text-slate-800">Tambah Akun Guru Baru</h2>
                </div>

                {userError && (
                  <div className="bg-rose-500/10 border border-rose-500/20 text-rose-700 p-3 rounded-xl text-xs flex items-start gap-2">
                    <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                    <span>{userError}</span>
                  </div>
                )}

                {userSuccess && (
                  <div className="bg-emerald-500/10 border border-emerald-500/20 text-emerald-700 p-3 rounded-xl text-xs flex items-center gap-2">
                    <Sparkles className="h-4 w-4 shrink-0" />
                    <span>{userSuccess}</span>
                  </div>
                )}

                <form onSubmit={async (e) => {
                  e.preventDefault();
                  if (!newUserName || !newUserEmail || !newUserPassword) {
                    setUserError("Silakan lengkapi semua isian.");
                    return;
                  }
                  setUserError(null);
                  setUserSuccess(null);
                  setIsAddingUser(true);
                  try {
                    const result = await addUserToCloud({
                      name: newUserName.trim(),
                      email: newUserEmail.trim(),
                      password: newUserPassword,
                      role: newUserRole,
                      mataPelajaran: newUserRole === 'user' ? newUserMataPelajaran : 'Sosiologi'
                    });

                    if (result.success) {
                      setUserSuccess(`Akun ${newUserName} (${newUserRole === 'admin' ? 'Admin' : 'Guru'}) berhasil didaftarkan ke Cloud Database!`);
                      setNewUserName('');
                      setNewUserEmail('');
                      setNewUserPassword('');
                      setNewUserRole('user');
                      setNewUserMataPelajaran('Sosiologi');
                    } else {
                      setUserError(result.message || "Gagal mendaftarkan pengguna baru.");
                    }
                  } catch (err: any) {
                    console.error(err);
                    setUserError(`Gagal membuat akun: ${err.message || err}`);
                  } finally {
                    setIsAddingUser(false);
                  }
                }} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Nama Lengkap Guru</label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: Budi Santoso, S.Pd."
                      value={newUserName}
                      onChange={(e) => setNewUserName(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-xs text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Alamat Email Resmi</label>
                    <input
                      type="email"
                      required
                      placeholder="nama.guru@sekolah.sch.id"
                      value={newUserEmail}
                      onChange={(e) => setNewUserEmail(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-xs text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Password Baru</label>
                    <input
                      type="password"
                      required
                      placeholder="Minimal 6 karakter"
                      value={newUserPassword}
                      onChange={(e) => setNewUserPassword(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-xs text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Peran Hak Akses (Role)</label>
                    <div className="grid grid-cols-2 gap-3 mt-1.5">
                      <button
                        type="button"
                        onClick={() => setNewUserRole('user')}
                        className={`py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-center gap-1.5 transition ${
                          newUserRole === 'user'
                            ? 'bg-indigo-50 border-indigo-500 text-indigo-700'
                            : 'bg-slate-50 border-slate-200 text-slate-500 hover:text-slate-800'
                        }`}
                      >
                        <User className="h-4 w-4" />
                        Guru Mapel
                      </button>
                      <button
                        type="button"
                        onClick={() => setNewUserRole('admin')}
                        className={`py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-center gap-1.5 transition ${
                          newUserRole === 'admin'
                            ? 'bg-indigo-50 border-indigo-500 text-indigo-700'
                            : 'bg-slate-50 border-slate-200 text-slate-500 hover:text-slate-800'
                        }`}
                      >
                        <Shield className="h-4 w-4" />
                        Administrator
                      </button>
                    </div>
                  </div>

                  {newUserRole === 'user' && (
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Mata Pelajaran Ampuan</label>
                      <select
                        value={newUserMataPelajaran}
                        onChange={(e) => setNewUserMataPelajaran(e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-xs text-slate-800 font-bold"
                      >
                        <option value="Sosiologi">👥 Sosiologi</option>
                        <option value="Matematika">📐 Matematika</option>
                        <option value="Bahasa Indonesia">🇮🇩 Bahasa Indonesia</option>
                        <option value="Bahasa Inggris">🇬🇧 Bahasa Inggris</option>
                        <option value="Matematika Tingkat Lanjut">🚀 Matematika Tingkat Lanjut</option>
                        <option value="Bahasa Indonesia Tingkat Lanjut">✍️ Bahasa Indonesia Tingkat Lanjut</option>
                        <option value="Bahasa Inggris Tingkat Lanjut">🗣️ Bahasa Inggris Tingkat Lanjut</option>
                        <option value="Fisika">⚛️ Fisika</option>
                        <option value="Kimia">🧪 Kimia</option>
                        <option value="Biologi">🧬 Biologi</option>
                        <option value="PPKN">🏛️ PPKN</option>
                        <option value="Ekonomi">📈 Ekonomi</option>
                        <option value="Geografi">🗺️ Geografi</option>
                        <option value="Sejarah Tingkat Lanjut">📜 Sejarah Tingkat Lanjut</option>
                        <option value="Antropologi">🗿 Antropologi</option>
                        <option value="Bahasa Jepang">🇯🇵 Bahasa Jepang</option>
                        <option value="Produk Kreatif dan Kewirausahaan">🛠️ Produk Kreatif dan Kewirausahaan</option>
                      </select>
                    </div>
                  )}

                  <button
                    type="submit"
                    disabled={isAddingUser}
                    className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white font-extrabold py-3 px-4 rounded-xl shadow-md transition flex items-center justify-center gap-2 text-xs"
                  >
                    {isAddingUser ? (
                      <>
                        <RefreshCw className="h-4 w-4 animate-spin" />
                        <span>Mendaftarkan Guru ke Cloud...</span>
                      </>
                    ) : (
                      <>
                        <UserPlus className="h-4 w-4" />
                        <span>Daftarkan Guru Baru</span>
                      </>
                    )}
                  </button>
                </form>
              </section>

              {/* Right Col: Current Users List */}
              <section className="lg:col-span-7 bg-white border border-slate-200 rounded-2xl shadow-sm p-6 flex flex-col space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <Users className="h-5 w-5 text-indigo-600" />
                    <h2 className="text-lg font-bold text-slate-800">Daftar Pengguna Sistem</h2>
                  </div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="bg-indigo-50 text-indigo-700 text-[10px] font-black px-2.5 py-1 rounded-full uppercase tracking-wider">
                      {usersList.length} Pengguna
                    </span>

                    {/* Batch Impor Excel / CSV / JSON Button */}
                    <label className="cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white shadow-sm border border-emerald-600 px-3 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1.5 transition">
                      {isBatchImportingUsers ? (
                        <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                      ) : (
                        <Upload className="h-3.5 w-3.5" />
                      )}
                      <span>📊 IMPOR EXCEL / CSV</span>
                      <input 
                        type="file" 
                        accept=".xlsx,.xls,.csv,.json"
                        className="hidden"
                        disabled={isBatchImportingUsers}
                        onChange={async (e) => {
                          const file = e.target.files?.[0];
                          if (!file) return;
                          setIsBatchImportingUsers(true);
                          setUserError(null);
                          setUserSuccess(null);
                          try {
                            let userItems: Array<{ name: string; email: string; password?: string; role?: string; mataPelajaran?: string }> = [];

                            if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                              const buffer = await file.arrayBuffer();
                              const workbook = XLSX.read(buffer, { type: 'array' });
                              const firstSheetName = workbook.SheetNames[0];
                              const worksheet = workbook.Sheets[firstSheetName];
                              const rawData = XLSX.utils.sheet_to_json<any>(worksheet);

                              userItems = rawData.map(row => ({
                                name: String(row['Nama'] || row['nama'] || row['Nama Lengkap'] || row['Name'] || row['name'] || 'Guru').trim(),
                                email: String(row['Email'] || row['email'] || row['Alamat Email'] || '').trim(),
                                password: String(row['Password'] || row['password'] || row['Pass'] || row['pass'] || 'guru123').trim(),
                                role: String(row['Role'] || row['role'] || row['Peran'] || 'user').toLowerCase().includes('admin') ? 'admin' : 'user',
                                mataPelajaran: String(row['Mata Pelajaran'] || row['mataPelajaran'] || row['Mapel'] || row['mapel'] || 'Sosiologi').trim()
                              }));
                            } else if (file.name.endsWith('.json')) {
                              const text = await file.text();
                              const data = JSON.parse(text);
                              const userArray = Array.isArray(data) ? data : [data];
                              userItems = userArray.map(item => ({
                                name: item.name || item.Nama || 'Guru',
                                email: item.email || item.Email || '',
                                password: String(item.password || item.Password || 'guru123'),
                                role: (item.role || item.Role || 'user').toString().toLowerCase().includes('admin') ? 'admin' : 'user',
                                mataPelajaran: item.mataPelajaran || item.Mapel || 'Sosiologi'
                              }));
                            } else {
                              // CSV Parsing
                              const text = await file.text();
                              const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
                              for (let i = 0; i < lines.length; i++) {
                                if (i === 0 && (lines[i].toLowerCase().includes('email') || lines[i].toLowerCase().includes('nama'))) {
                                  continue;
                                }
                                const cols = lines[i].split(',').map(c => c.trim().replace(/^"|"$/g, ''));
                                if (cols.length >= 2) {
                                  const [name, email, pass, roleVal, mapelVal] = cols;
                                  userItems.push({
                                    name: name || 'Guru',
                                    email: email || '',
                                    password: pass || 'guru123',
                                    role: roleVal === 'admin' ? 'admin' : 'user',
                                    mataPelajaran: mapelVal || 'Sosiologi'
                                  });
                                }
                              }
                            }

                            const result = await importUsersBatchToCloud(userItems);
                            if (result.success) {
                              setUserSuccess(`Berhasil mengimpor ${result.count} akun pengguna ke Cloud Database!`);
                            } else {
                              setUserError(result.message || "Gagal mengimpor data akun.");
                            }
                          } catch (err: any) {
                            console.error(err);
                            setUserError(`Gagal impor file Excel/CSV: ${err.message || err}`);
                          } finally {
                            setIsBatchImportingUsers(false);
                            e.target.value = '';
                          }
                        }}
                      />
                    </label>

                    {/* Export Excel with Password Button */}
                    <button
                      type="button"
                      onClick={() => {
                        if (usersList.length === 0) {
                          alert("Belum ada data pengguna untuk diekspor.");
                          return;
                        }

                        try {
                          const exportData = usersList.map((u, idx) => ({
                            'No': idx + 1,
                            'Nama Lengkap / Username': u.name || 'Guru Sosiologi',
                            'Alamat Email': u.email || '',
                            'Password': u.password || (u.email === 'admin@tka.com' ? 'admin123' : u.email === 'user@tka.com' ? 'user123' : 'guru123'),
                            'Peran (Role)': u.role === 'admin' ? 'Administrator' : 'Guru Mapel',
                            'Mata Pelajaran Ampuan': u.mataPelajaran || 'Sosiologi',
                            'Tanggal Dibuat': u.createdAt ? new Date(u.createdAt).toLocaleDateString('id-ID') : '-'
                          }));

                          const ws = XLSX.utils.json_to_sheet(exportData);
                          const wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, ws, "Daftar Akun Pengguna");
                          XLSX.writeFile(wb, `Data_Pengguna_MasterTKA_${new Date().toISOString().split('T')[0]}.xlsx`);
                        } catch (err: any) {
                          console.error("Gagal mengekspor data akun:", err);
                          alert("Gagal mengekspor file data pengguna.");
                        }
                      }}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white shadow-sm border border-indigo-600 px-3 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1.5 transition cursor-pointer"
                      title="Unduh seluruh data pengguna beserta password ke Excel"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>📥 EKSPOR EXCEL</span>
                    </button>
                  </div>
                </div>

                {/* Filter and Password Reveal Toggle Toolbar */}
                <div className="flex flex-col sm:flex-row items-center justify-between gap-3">
                  <div className="flex-1 w-full">
                    <input
                      type="text"
                      placeholder="🔍 Cari berdasarkan nama, email, password, atau mapel..."
                      value={userSearchQuery}
                      onChange={(e) => setUserSearchQuery(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 rounded-xl px-3.5 py-2 text-xs text-slate-800"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowUserPasswords(!showUserPasswords)}
                    className={`px-3 py-2 rounded-xl text-xs font-bold border transition flex items-center gap-1.5 shrink-0 cursor-pointer ${
                      showUserPasswords
                        ? 'bg-amber-500 text-white border-amber-500 shadow-sm'
                        : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-200'
                    }`}
                    title={showUserPasswords ? "Sembunyikan semua password" : "Tampilkan semua password akun"}
                  >
                    {showUserPasswords ? <EyeOff className="h-3.5 w-3.5" /> : <Eye className="h-3.5 w-3.5 text-indigo-600" />}
                    <span>{showUserPasswords ? "Tutup Password" : "Lihat Semua Password"}</span>
                  </button>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="border-b border-slate-150 text-slate-500 font-bold uppercase tracking-wider text-[10px]">
                        <th className="py-3 px-2">Nama Pengguna</th>
                        <th className="py-3 px-2">Email</th>
                        <th className="py-3 px-2">Password Akun</th>
                        <th className="py-3 px-2">Role</th>
                        <th className="py-3 px-2 text-right">Aksi CRUD</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {usersList
                        .filter((usr) => {
                          if (!userSearchQuery) return true;
                          const q = userSearchQuery.toLowerCase();
                          return (usr.name || '').toLowerCase().includes(q) ||
                                 (usr.email || '').toLowerCase().includes(q) ||
                                 (usr.password || '').toLowerCase().includes(q) ||
                                 (usr.mataPelajaran || '').toLowerCase().includes(q);
                        })
                        .map((usr) => {
                          const displayPassword = usr.password || (usr.email === 'admin@tka.com' ? 'admin123' : usr.email === 'user@tka.com' ? 'user123' : 'guru123');
                          const isRevealed = showUserPasswords || !!revealedPasswordUserIds[usr.id];

                          return (
                          <tr key={usr.id} className="hover:bg-slate-50/50 transition">
                            <td className="py-3.5 px-2 font-bold text-slate-800">
                              <div>{usr.name || 'Guru Sosiologi'}</div>
                              {usr.role !== 'admin' && (
                                <div className="text-[10px] text-indigo-600 font-medium mt-0.5 flex items-center gap-1">
                                  <BookOpen className="h-3 w-3 text-indigo-500 shrink-0" />
                                  <span>Mapel Ampuan: <b>{usr.mataPelajaran || 'Sosiologi'}</b></span>
                                </div>
                              )}
                            </td>
                            <td className="py-3.5 px-2 text-slate-600">{usr.email}</td>
                            
                            {/* Password Column with Reveal and Copy */}
                            <td className="py-3.5 px-2">
                              <div className="flex items-center gap-1.5">
                                <span className={`font-mono text-xs px-2 py-0.5 rounded-lg border ${
                                  isRevealed 
                                    ? 'bg-amber-50 text-amber-900 border-amber-200 font-bold' 
                                    : 'bg-slate-100 text-slate-500 border-slate-200 tracking-widest'
                                }`}>
                                  {isRevealed ? displayPassword : '••••••••'}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setRevealedPasswordUserIds(prev => ({
                                      ...prev,
                                      [usr.id]: !prev[usr.id]
                                    }));
                                  }}
                                  className="p-1 rounded text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition cursor-pointer"
                                  title={isRevealed ? "Sembunyikan password" : "Lihat password"}
                                >
                                  {isRevealed ? <EyeOff className="h-3.5 w-3.5 text-amber-600" /> : <Eye className="h-3.5 w-3.5" />}
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    navigator.clipboard.writeText(displayPassword);
                                    setUserSuccess(`Password untuk ${usr.name || usr.email} (${displayPassword}) berhasil disalin!`);
                                    setTimeout(() => setUserSuccess(null), 3000);
                                  }}
                                  className="p-1 rounded text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 transition cursor-pointer"
                                  title="Salin password ke clipboard"
                                >
                                  <Copy className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </td>

                            <td className="py-3.5 px-2">
                              <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full font-bold text-[9px] uppercase tracking-wide ${
                                usr.role === 'admin'
                                  ? 'bg-amber-50 text-amber-700 border border-amber-200'
                                  : 'bg-indigo-50 text-indigo-700 border border-indigo-200'
                              }`}>
                                {usr.role === 'admin' ? <Shield className="h-2.5 w-2.5" /> : <User className="h-2.5 w-2.5" />}
                                {usr.role === 'admin' ? 'Admin' : 'Guru'}
                              </span>
                            </td>
                            <td className="py-3.5 px-2 text-right">
                              <div className="flex items-center justify-end gap-1.5">
                                {/* Edit Button */}
                                <button
                                  onClick={() => setEditingUser({
                                    id: usr.id,
                                    name: usr.name || '',
                                    email: usr.email || '',
                                    role: usr.role || 'user',
                                    mataPelajaran: usr.mataPelajaran || 'Sosiologi',
                                    newPassword: ''
                                  })}
                                  className="p-1.5 rounded-lg border bg-indigo-50 hover:bg-indigo-100 text-indigo-600 border-indigo-200 hover:border-indigo-300 transition cursor-pointer"
                                  title="Ubah Nama, Password, Peran & Mapel Ampuan"
                                >
                                  <Edit className="h-3.5 w-3.5" />
                                </button>

                                {/* Delete Button */}
                                <button
                                  disabled={usr.email === currentUser?.email}
                                  onClick={async () => {
                                    if (confirm(`Apakah Anda yakin ingin menghapus akses pengguna ${usr.name || usr.email}?`)) {
                                      try {
                                        const res = await deleteUserFromCloud(usr.id);
                                        if (res.success) {
                                          setUserSuccess(`Pengguna ${usr.name || usr.email} berhasil dihapus.`);
                                        } else {
                                          alert(`Gagal menghapus pengguna: ${res.message}`);
                                        }
                                      } catch (err) {
                                        console.error("Gagal menghapus pengguna:", err);
                                        alert("Gagal menghapus pengguna dari database.");
                                      }
                                    }
                                  }}
                                  className={`p-1.5 rounded-lg border transition ${
                                    usr.email === currentUser?.email
                                      ? 'opacity-40 cursor-not-allowed bg-slate-100 border-slate-200 text-slate-400'
                                      : 'bg-rose-50 hover:bg-rose-100 text-rose-600 border-rose-200 hover:border-rose-300 cursor-pointer'
                                  }`}
                                  title={usr.email === currentUser?.email ? "Anda tidak dapat menghapus akun Anda sendiri" : "Hapus Pengguna"}
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                          );
                        })}
                      {usersList.length === 0 && (
                        <tr>
                          <td colSpan={5} className="text-center py-8 text-slate-400">
                            Tidak ada pengguna terdaftar di Cloud Database.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>

            {/* Modal Edit User CRUD */}
            {editingUser && (
              <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
                <div className="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden p-6 space-y-5">
                  <div className="flex justify-between items-center pb-4 border-b border-slate-100">
                    <div className="flex items-center gap-2">
                      <Edit className="h-5 w-5 text-indigo-600" />
                      <h3 className="text-base font-extrabold text-slate-800">Ubah Profil & Mapel Ampuan</h3>
                    </div>
                    <button 
                      onClick={() => setEditingUser(null)}
                      className="text-slate-400 hover:text-slate-600 p-1 rounded-full hover:bg-slate-100 transition"
                    >
                      <X className="h-5 w-5" />
                    </button>
                  </div>

                  <div className="space-y-4 text-xs">
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Email Pengguna (Readonly)</label>
                      <input type="text" disabled value={editingUser.email} className="w-full bg-slate-100 border border-slate-200 rounded-xl px-4 py-2.5 text-slate-500 font-medium cursor-not-allowed" />
                      <p className="text-[10px] text-slate-400 mt-1">Email adalah identitas unik login dan akun sistem.</p>
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Nama Lengkap / Username</label>
                      <input 
                        type="text" 
                        value={editingUser.name} 
                        onChange={(e) => setEditingUser({ ...editingUser, name: e.target.value })}
                        placeholder="Contoh: Budi Santoso, S.Pd. atau budisosiologi"
                        className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-slate-800 font-bold text-xs"
                      />
                      <p className="text-[10px] text-slate-400 mt-1">Guru juga dapat menggunakan Nama ini saat login di layar masuk.</p>
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">
                        Password Baru <span className="text-slate-400 font-normal">(Kosongkan jika tidak ingin mengubah password)</span>
                      </label>
                      <div className="relative">
                        <input 
                          type={editingUser.showPassword ? "text" : "password"} 
                          value={editingUser.newPassword || ''} 
                          onChange={(e) => setEditingUser({ ...editingUser, newPassword: e.target.value })}
                          placeholder="Masukkan password baru (minimal 6 karakter)"
                          className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl pl-4 pr-10 py-2.5 text-slate-800 font-bold text-xs"
                        />
                        <button
                          type="button"
                          onClick={() => setEditingUser({ ...editingUser, showPassword: !editingUser.showPassword })}
                          className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-400 hover:text-slate-600 transition cursor-pointer"
                          title={editingUser.showPassword ? "Sembunyikan Password" : "Lihat Password"}
                        >
                          {editingUser.showPassword ? <EyeOff className="h-4 w-4 text-indigo-500" /> : <Eye className="h-4 w-4" />}
                        </button>
                      </div>
                      <p className="text-[10px] text-slate-500 mt-1">
                        🔑 Tulis password baru jika guru lupa password atau ingin mengganti kredensial.
                      </p>
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Peran Hak Akses (Role)</label>
                      <div className="grid grid-cols-2 gap-3 mt-1">
                        <button
                          type="button"
                          onClick={() => setEditingUser({ ...editingUser, role: 'user' })}
                          className={`py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-center gap-1.5 transition ${
                            editingUser.role === 'user'
                              ? 'bg-indigo-50 border-indigo-500 text-indigo-700'
                              : 'bg-slate-50 border-slate-200 text-slate-500'
                          }`}
                        >
                          <User className="h-4 w-4" />
                          Guru Mapel
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditingUser({ ...editingUser, role: 'admin' })}
                          className={`py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-center gap-1.5 transition ${
                            editingUser.role === 'admin'
                              ? 'bg-indigo-50 border-indigo-500 text-indigo-700'
                              : 'bg-slate-50 border-slate-200 text-slate-500'
                          }`}
                        >
                          <Shield className="h-4 w-4" />
                          Administrator
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Mata Pelajaran Ampuan</label>
                      <select
                        value={editingUser.mataPelajaran || 'Sosiologi'}
                        onChange={(e) => setEditingUser({ ...editingUser, mataPelajaran: e.target.value })}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-xl px-4 py-2.5 text-xs text-slate-800 font-bold"
                      >
                        <option value="Sosiologi">👥 Sosiologi</option>
                        <option value="Matematika">📐 Matematika</option>
                        <option value="Bahasa Indonesia">🇮🇩 Bahasa Indonesia</option>
                        <option value="Bahasa Inggris">🇬🇧 Bahasa Inggris</option>
                        <option value="Matematika Tingkat Lanjut">🚀 Matematika Tingkat Lanjut</option>
                        <option value="Bahasa Indonesia Tingkat Lanjut">✍️ Bahasa Indonesia Tingkat Lanjut</option>
                        <option value="Bahasa Inggris Tingkat Lanjut">🗣️ Bahasa Inggris Tingkat Lanjut</option>
                        <option value="Fisika">⚛️ Fisika</option>
                        <option value="Kimia">🧪 Kimia</option>
                        <option value="Biologi">🧬 Biologi</option>
                        <option value="Pendidikan Pancasila dan Kewarganegaraan">🏛️ Pendidikan Pancasila dan Kewarganegaraan (PPKN)</option>
                        <option value="Ekonomi">📈 Ekonomi</option>
                        <option value="Geografi">🗺️ Geografi</option>
                        <option value="Sejarah">📜 Sejarah</option>
                        <option value="Antropologi">🗿 Antropologi</option>
                        <option value="Bahasa Jepang">🇯🇵 Bahasa Jepang</option>
                        <option value="Produk atau Projek Kreatif dan Kewirausahaan SMK dan MAK">🛠️ Produk / Projek Kreatif dan Kewirausahaan</option>
                      </select>
                    </div>
                  </div>

                  <div className="flex items-center justify-end gap-2 pt-4 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setEditingUser(null)}
                      className="px-4 py-2.5 text-xs font-bold text-slate-600 hover:bg-slate-100 rounded-xl transition"
                    >
                      Batal
                    </button>
                    <button
                      type="button"
                      disabled={isSavingUserEdit}
                      onClick={async () => {
                        if (!editingUser.name || !editingUser.name.trim()) {
                          alert('Nama pengguna tidak boleh kosong.');
                          return;
                        }
                        if (editingUser.newPassword && editingUser.newPassword.trim().length > 0 && editingUser.newPassword.trim().length < 6) {
                          alert('Password baru minimal harus 6 karakter.');
                          return;
                        }

                        setIsSavingUserEdit(true);
                        try {
                          const updatedSubject = editingUser.mataPelajaran || 'Sosiologi';
                          const updatePayload: any = {
                            id: editingUser.id,
                            email: editingUser.email,
                            name: editingUser.name.trim(),
                            role: editingUser.role,
                            mataPelajaran: updatedSubject
                          };

                          if (editingUser.newPassword && editingUser.newPassword.trim().length >= 6) {
                            updatePayload.password = editingUser.newPassword.trim();
                          }

                          const res = await updateUserInCloud(updatePayload);

                          if (res.success) {
                            if (editingUser.id === currentUser?.uid || editingUser.email === currentUser?.email) {
                              setUserName(editingUser.name);
                              setUserRole(editingUser.role);
                              setConfig(prev => ({ ...prev, mataPelajaran: updatedSubject }));
                            }
                            const passNote = editingUser.newPassword ? ' dan Password' : '';
                            setUserSuccess(`Profil, Nama${passNote}, & Mata Pelajaran "${editingUser.name}" berhasil diperbarui!`);
                            setEditingUser(null);
                          } else {
                            alert(`Gagal menyimpan perubahan: ${res.message}`);
                          }
                        } catch (err: any) {
                          console.error(err);
                          alert(`Gagal menyimpan perubahan: ${err.message}`);
                        } finally {
                          setIsSavingUserEdit(false);
                        }
                      }}
                      className="px-5 py-2.5 text-xs font-extrabold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl shadow transition flex items-center gap-1.5 cursor-pointer"
                    >
                      {isSavingUserEdit ? <RefreshCw className="h-4 w-4 animate-spin" /> : <Save className="h-4 w-4" />}
                      <span>Simpan Perubahan</span>
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Modal Panduan & Salin Kode SQL Supabase */}
            {showSupabaseSqlModal && (
              <div className="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-md flex items-center justify-center p-4 overflow-y-auto animate-fadeIn">
                <div className="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-3xl w-full my-8 overflow-hidden p-6 sm:p-8 space-y-6">
                  <div className="flex justify-between items-start pb-4 border-b border-slate-100">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-emerald-100 text-emerald-700 rounded-2xl">
                        <FileCode className="h-6 w-6" />
                      </div>
                      <div>
                        <h3 className="text-lg font-black text-slate-800">Panduan Setup Database Supabase</h3>
                        <p className="text-xs text-slate-500">PostgreSQL Cloud dengan enkripsi keamanan tingkat tinggi & Row Level Security.</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => setShowSupabaseSqlModal(false)}
                      className="text-slate-400 hover:text-slate-600 p-1.5 rounded-full hover:bg-slate-100 transition"
                    >
                      <X className="h-5 w-5" />
                    </button>
                  </div>

                  <div className="space-y-4 text-xs text-slate-600 leading-relaxed max-h-[60vh] overflow-y-auto pr-2">
                    <div className="bg-emerald-50 border border-emerald-200 p-4 rounded-2xl text-emerald-900 space-y-1">
                      <p className="font-bold flex items-center gap-1.5 text-emerald-800">
                        <Sparkles className="h-4 w-4 text-emerald-600" />
                        Keuntungan Database Supabase
                      </p>
                      <p className="text-[11px] text-emerald-700">
                        Supabase adalah database PostgreSQL berbasis cloud. Password guru tersimpan aman, performa super cepat, dan multi-perangkat tersinkronisasi otomatis untuk website <b>mastertkasma.my.id</b>.
                      </p>
                    </div>

                    <ol className="space-y-4 list-decimal list-inside font-medium text-slate-700">
                      <li className="space-y-1">
                        <b>Buka Dashboard Supabase:</b>
                        <p className="text-slate-500 pl-5">Buka <a href="https://supabase.com/dashboard" target="_blank" rel="noreferrer" className="text-emerald-600 font-bold underline">supabase.com/dashboard</a>, login, dan buat Project baru (contoh: <code>tka-sma-db</code>).</p>
                      </li>

                      <li className="space-y-2">
                        <b>Buka SQL Editor & Jalankan Kode Berikut:</b>
                        <p className="text-slate-500 pl-5">Di menu sebelah kiri Supabase, klik <b>SQL Editor</b> &gt; klik <b>New query</b> &gt; Tempel kode SQL di bawah ini &gt; klik tombol hijau <b>Run</b>.</p>
                        
                        <div className="relative bg-slate-900 text-emerald-300 p-4 rounded-2xl font-mono text-[11px] overflow-x-auto border border-slate-800">
                          <button
                            type="button"
                            onClick={() => {
                              navigator.clipboard.writeText(SUPABASE_SQL_SETUP_CODE);
                              setIsCopyingSqlCode(true);
                              setTimeout(() => setIsCopyingSqlCode(false), 2000);
                            }}
                            className="absolute top-3 right-3 bg-emerald-500 hover:bg-emerald-400 text-emerald-950 font-bold text-[10px] px-3 py-1.5 rounded-lg shadow transition flex items-center gap-1"
                          >
                            {isCopyingSqlCode ? <Check className="h-3.5 w-3.5" /> : <Copy className="h-3.5 w-3.5" />}
                            <span>{isCopyingSqlCode ? 'Tersalin!' : 'Salin Kode SQL'}</span>
                          </button>
                          <pre className="max-h-56 overflow-y-auto whitespace-pre-wrap">{SUPABASE_SQL_SETUP_CODE}</pre>
                        </div>
                      </li>

                      <li className="space-y-1">
                        <b>Ambil URL dan API Key:</b>
                        <p className="text-slate-500 pl-5">Buka menu <b>Project Settings</b> (ikon roda gigi) &gt; pilih tab <b>API</b>.</p>
                        <p className="text-slate-500 pl-5">Salin <b>Project URL</b> (contoh: <code>https://abcdefgh.supabase.co</code>) dan <b>anon public key</b>.</p>
                      </li>

                      <li className="space-y-1">
                        <b>Tempel di Panel Ini:</b>
                        <p className="text-slate-500 pl-5">Tempelkan Project URL dan Anon Key pada input di atas, klik <b>Tes Koneksi</b> lalu klik <b>Simpan</b>.</p>
                      </li>
                    </ol>
                  </div>

                  <div className="flex justify-end pt-4 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setShowSupabaseSqlModal(false)}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold px-6 py-2.5 rounded-xl shadow text-xs transition"
                    >
                      Tutup & Mulai Gunakan
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Questions Printable List */}
        <div 
          id="questions-printable-container" 
          className={`space-y-6 bg-white border border-slate-200 rounded-2xl shadow-sm p-6 sm:p-10 print:border-none print:p-0 print:shadow-none ${
            activeTab === 'soal' ? 'block' : 'hidden print:block'
          }`}
        >
              
              {/* Kop Surat Resmi */}
              {printConfig.showHeader && (
                <div className="border-b-[3px] border-double border-slate-900 pb-3 mb-5 flex items-center gap-4 text-center">
                  {printConfig.schoolLogo ? (
                    <img 
                      src={printConfig.schoolLogo} 
                      className="w-14 h-14 object-contain flex-shrink-0" 
                      alt="Logo Kiri" 
                    />
                  ) : (
                    <div className="w-14 h-14 border border-slate-800 rounded-full flex-shrink-0 flex items-center justify-center font-sans font-bold text-[9px] text-slate-800">
                      KOP
                    </div>
                  )}
                  <div className="flex-1">
                    {(printConfig.kopDepartment || 'KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI')
                      .split('\n')
                      .map((line, idx) => (
                        <h4 key={idx} className="text-xs font-bold uppercase tracking-wide">{line.trim()}</h4>
                      ))
                    }
                    <h3 className="text-sm sm:text-base font-black uppercase tracking-wider">{printConfig.schoolName}</h3>
                    <p className="text-[9px] text-slate-600 italic">{printConfig.schoolAddress}</p>
                    <div className="border-t border-slate-400 mt-1 pt-1 flex justify-center gap-4 text-[9px] font-bold text-slate-700">
                      <span>TAHUN PELAJARAN: {printConfig.academicYear}</span>
                      <span>SEMESTER: {printConfig.semester.toUpperCase()}</span>
                    </div>
                  </div>
                  {printConfig.schoolLogoRight ? (
                    <img 
                      src={printConfig.schoolLogoRight} 
                      className="w-14 h-14 object-contain flex-shrink-0" 
                      alt="Logo Kanan" 
                    />
                  ) : (
                    <div className="w-14 h-14 border border-slate-800 rounded-full flex-shrink-0 flex items-center justify-center font-sans font-bold text-[9px] text-slate-800">
                      SMA
                    </div>
                  )}
                </div>
              )}

              {/* Title Section inside paper */}
              {!printConfig.showHeader && (
                <div className="border-b-2 border-slate-900 pb-4 text-center mb-4">
                  <h3 className="text-lg font-bold uppercase tracking-wide">LEMBAR SOAL UJIAN TKA SMA</h3>
                  <p className="text-xl font-extrabold text-slate-900">{printConfig.subjectName || config.mataPelajaran || 'TES KEMAMPUAN AKADEMIK'}</p>
                  <div className="mt-2 text-xs text-slate-600 flex justify-center gap-6">
                    <span><b>Tingkat/Kurikulum:</b> {config.muatan}</span>
                    <span><b>Tanggal:</b> {new Date().toLocaleDateString('id-ID')}</span>
                  </div>
                </div>
              )}

              {/* Exam Metadata Title (under Kop Surat) */}
              {printConfig.showHeader && (
                <div className="text-center space-y-0.5 mb-5">
                  <h2 className="text-sm font-extrabold tracking-wide uppercase">{printConfig.examName}</h2>
                  <h1 className="text-base font-black text-slate-900 uppercase">MATA PELAJARAN: {printConfig.subjectName || config.mataPelajaran || 'TES KEMAMPUAN AKADEMIK'}</h1>
                  <div className="text-[10px] text-slate-600 flex justify-center gap-4 font-semibold">
                    <span>Fase/Muatan: {config.muatan || 'SMA'}</span>
                    <span>Alokasi Waktu: {printConfig.timeAllocation}</span>
                  </div>
                </div>
              )}

              {/* Student Identity Section */}
              {printConfig.showStudentFields && (
                <div className="grid grid-cols-2 gap-4 border border-slate-400 p-3 rounded-lg text-[11px] font-semibold mb-5">
                  <div className="space-y-1">
                    <div className="flex"><span className="w-24">NAMA LENGKAP</span><span className="mr-2">:</span><span className="flex-1 border-b border-dashed border-slate-400"></span></div>
                    <div className="flex"><span className="w-24">NOMOR PESERTA</span><span className="mr-2">:</span><span className="flex-1 border-b border-dashed border-slate-400"></span></div>
                  </div>
                  <div className="space-y-1">
                    <div className="flex"><span className="w-24">KELAS / JURUSAN</span><span className="mr-2">:</span><span className="flex-1 border-b border-dashed border-slate-400"></span></div>
                    <div className="flex"><span className="w-24">HARI / TANGGAL</span><span className="mr-2">:</span><span className="flex-1 text-slate-700">{new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</span></div>
                  </div>
                </div>
              )}

              {/* Instructions text */}
              {printConfig.instructionText && (
                <div className="border-l-2 border-slate-900 pl-3 py-0.5 text-xs text-slate-700 italic mb-5 leading-relaxed">
                  <b>PETUNJUK PENGERJAAN:</b> {printConfig.instructionText}
                </div>
              )}

              {(() => {
                const displayQuestions = questions.filter(q => {
                  if (selectedBentukFilter === 'all') return true;
                  if (selectedBentukFilter === 'kategori') return isKategoriSoal(q);
                  if (selectedBentukFilter === 'mcma') return q.bentukSoal === 'mcma';
                  if (selectedBentukFilter === 'pilihan_ganda_sederhana') return q.bentukSoal === 'pilihan_ganda_sederhana' || (!q.bentukSoal && !isKategoriSoal(q));
                  return true;
                });

                if (displayQuestions.length === 0) {
                  return (
                    <div className="text-center py-16 text-slate-400 font-medium bg-slate-50 border border-dashed border-slate-200 rounded-2xl">
                      {questions.length === 0 
                        ? 'Belum ada butir soal yang tersusun. Sila buat Kisi-Kisi terlebih dahulu, lalu tekan tombol AI untuk menyusun atau tambahkan secara manual.'
                        : `Tidak ada butir soal yang sesuai dengan filter '${getBentukSoalLabel(selectedBentukFilter)}'.`}
                    </div>
                  );
                }

                return (
                  <div className={`${printConfig.layoutColumns === '2' ? 'columns-1 md:columns-2 gap-x-8 gap-y-6 print:columns-2 print:gap-x-8' : 'space-y-8'} ${printConfig.fontSize}`}>
                    {displayQuestions.map((q, idx) => (
                    <div 
                      id={`soal-card-${q.id}`}
                      key={q.id ? `${q.id}-${idx}` : `q-${idx}`} 
                      className={`inline-block w-full break-inside-avoid page-break-inside-avoid transition-all duration-300 space-y-3 ${
                        editingQuestionId === q.id
                          ? 'border-2 border-indigo-600 bg-gradient-to-b from-indigo-50/70 via-indigo-50/20 to-white p-4 sm:p-5 rounded-2xl ring-4 ring-indigo-400/50 shadow-2xl my-6 animate-active-card-glow'
                          : 'pb-6 mb-6 border-b border-slate-100 last:border-b-0'
                      }`}
                    >
                      {editingQuestionId === q.id ? (
                        /* INLINE EDIT FORM DIRECTLY AT THE QUESTION CARD LOCATION */
                        <div className="space-y-4 no-print">
                          {/* Active Animated Header Banner */}
                          <div className="flex items-center justify-between bg-gradient-to-r from-indigo-600 via-purple-600 to-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-black uppercase tracking-wider shadow-md animate-pulse">
                            <span className="flex items-center gap-2">
                              <Edit className="h-4 w-4 text-amber-300 animate-bounce" />
                              ✏️ SEDANG EDIT LANGSUNG DI TEMPAT: SOAL NO. {q.noSoal}
                            </span>
                            <div className="flex items-center gap-2">
                              <span className="bg-white/20 text-indigo-100 px-2.5 py-0.5 rounded-full text-[10px] font-extrabold border border-white/30">
                                Mode Edit Aktif
                              </span>
                              <button
                                type="button"
                                onClick={() => {
                                  setIsEditingQuestion(false);
                                  setEditingQuestionId(null);
                                }}
                                className="bg-white/20 hover:bg-white/30 text-white px-2.5 py-1 rounded-lg text-[10px] font-extrabold transition cursor-pointer"
                              >
                                ✖ Batal
                              </button>
                            </div>
                          </div>

                          {/* Inline Edit Form Body */}
                          <form onSubmit={handleSaveQuestionForm} className="bg-white p-4 sm:p-5 rounded-xl border border-indigo-200 shadow-sm space-y-4 text-xs">
                            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                              <div>
                                <label className="block font-bold text-slate-700 mb-1">Kompetensi</label>
                                <input
                                  type="text"
                                  value={questionForm.kompetensi}
                                  onChange={(e) => setQuestionForm({ ...questionForm, kompetensi: e.target.value })}
                                  className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 focus:bg-white focus:ring-2 focus:ring-indigo-500 font-medium text-xs"
                                />
                              </div>
                              <div>
                                <label className="block font-bold text-slate-700 mb-1">Sub-Kompetensi / Materi</label>
                                <input
                                  type="text"
                                  value={questionForm.subKompetensi}
                                  onChange={(e) => setQuestionForm({ ...questionForm, subKompetensi: e.target.value })}
                                  className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 focus:bg-white focus:ring-2 focus:ring-indigo-500 font-medium text-xs"
                                />
                              </div>
                              <div>
                                <label className="block font-bold text-slate-700 mb-1">Bentuk Soal</label>
                                <select
                                  value={questionForm.bentukSoal}
                                  onChange={(e) => setQuestionForm({ ...questionForm, bentukSoal: e.target.value as BentukSoal })}
                                  className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 font-bold text-slate-900 focus:bg-white focus:ring-2 focus:ring-indigo-500 text-xs"
                                >
                                  <option value="pilihan_ganda_sederhana">Pilihan Ganda Sederhana (PG)</option>
                                  <option value="mcma">PG Kompleks - MCMA (Multi-Jawaban)</option>
                                  <option value="kategori">PG Kompleks - Kategori (Tabel Pernyataan)</option>
                                </select>
                              </div>
                            </div>

                            {/* Stimulus */}
                            <div>
                              <label className="block font-bold text-slate-700 mb-1">Teks Stimulus / Wacana (Opsional)</label>
                              <textarea
                                rows={2}
                                value={questionForm.stimulus}
                                onChange={(e) => setQuestionForm({ ...questionForm, stimulus: e.target.value })}
                                placeholder="Masukkan stimulus/paragraf pengantar..."
                                className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500"
                              />
                            </div>

                            {/* Soal / Pertanyaan */}
                            <div>
                              <div className="mb-1">
                                <label className="block font-bold text-slate-800">Pertanyaan / Pokok Soal (Wajib)</label>
                              </div>
                              <textarea
                                rows={3}
                                value={questionForm.soal}
                                onChange={(e) => setQuestionForm({ ...questionForm, soal: e.target.value })}
                                placeholder="Masukkan pertanyaan utama..."
                                className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 font-semibold text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500"
                                required
                              />
                            </div>

                            {/* Opsi Jawaban */}
                            <div className="space-y-2 bg-slate-50/80 p-3 rounded-xl border border-slate-200">
                              <div className="flex justify-between items-center mb-1">
                                <span className="font-extrabold text-slate-800">
                                  {questionForm.bentukSoal === 'kategori' ? 'Daftar Pernyataan (P1 s.d. P5)' : 'Pilihan Jawaban (Opsi A s.d. E)'}
                                </span>
                                <span className="text-[10px] text-indigo-700 font-bold">
                                  {questionForm.bentukSoal === 'kategori' ? 'Pernyataan Tabel' : 'Opsi Jawaban'}
                                </span>
                              </div>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                {['A', 'B', 'C', 'D', 'E'].map((letter, i) => (
                                  <div key={letter} className="flex items-center gap-1.5">
                                    <span className={`font-black w-6 h-6 rounded-md flex items-center justify-center font-mono text-[10px] flex-shrink-0 ${
                                      questionForm.bentukSoal === 'kategori' ? 'bg-amber-100 text-amber-900 border border-amber-300' : 'bg-slate-200 text-slate-800'
                                    }`}>
                                      {questionForm.bentukSoal === 'kategori' ? `P${i + 1}` : letter}
                                    </span>
                                    <input
                                      type="text"
                                      value={(questionForm.opsi || [])[i] || ''}
                                      onChange={(e) => handleOpsiChange(i, e.target.value)}
                                      placeholder={questionForm.bentukSoal === 'kategori' ? `Pernyataan ke-${i + 1}` : `Opsi ${letter}`}
                                      className="flex-1 bg-white border border-slate-300 rounded-lg px-2.5 py-1 text-xs focus:ring-2 focus:ring-indigo-500"
                                    />
                                  </div>
                                ))}
                              </div>
                            </div>

                            {/* Kunci Jawaban & Pembahasan */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                              <div>
                                <label className="block font-bold text-slate-800 mb-1">
                                  Kunci Jawaban {questionForm.bentukSoal === 'kategori' ? 'Per Pernyataan' : ''} (Wajib)
                                </label>
                                <input
                                  type="text"
                                  value={questionForm.kunciJawaban}
                                  onChange={(e) => setQuestionForm({ ...questionForm, kunciJawaban: e.target.value })}
                                  placeholder={questionForm.bentukSoal === 'kategori' ? "Pernyataan 1: Sesuai, Pernyataan 2: Tidak Sesuai" : "A"}
                                  className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 font-bold text-indigo-900 text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500"
                                  required
                                />
                              </div>
                              <div>
                                <label className="block font-bold text-slate-700 mb-1">Pembahasan / Solusi</label>
                                <input
                                  type="text"
                                  value={questionForm.pembahasan}
                                  onChange={(e) => setQuestionForm({ ...questionForm, pembahasan: e.target.value })}
                                  placeholder="Pembahasan singkat..."
                                  className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs focus:bg-white focus:ring-2 focus:ring-indigo-500"
                                />
                              </div>
                            </div>

                            {/* Gambar / Ilustrasi */}
                            <div className="bg-slate-50/80 p-2.5 rounded-xl border border-slate-200 space-y-2">
                              <div className="flex justify-between items-center">
                                <label className="font-bold text-slate-700">Ilustrasi / Gambar Soal (Opsional)</label>
                                {questionForm.gambarUrl && (
                                  <span className="text-[10px] text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                                    ✓ Terpasang
                                  </span>
                                )}
                              </div>
                              <div className="flex items-center gap-2 flex-wrap">
                                <input
                                  type="text"
                                  value={questionForm.gambarUrl || ''}
                                  onChange={(e) => setQuestionForm({ ...questionForm, gambarUrl: e.target.value })}
                                  placeholder="https://... atau <svg>...</svg>"
                                  className="flex-1 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs font-mono"
                                />
                                <label className="cursor-pointer bg-slate-200 hover:bg-slate-300 text-slate-800 px-2.5 py-1 rounded-lg text-xs font-bold flex items-center gap-1 transition select-none">
                                  <Upload className="h-3.5 w-3.5" />
                                  <span>{isCompressingImage ? "⚡..." : "Upload"}</span>
                                  <input 
                                    type="file" 
                                    accept="image/*" 
                                    className="hidden" 
                                    disabled={isCompressingImage}
                                    onChange={async (e) => {
                                      const file = e.target.files?.[0];
                                      if (file) {
                                        setIsCompressingImage(true);
                                        try {
                                          const result = await compressImageFile(file);
                                          setQuestionForm(prev => ({ ...prev, gambarUrl: result.dataUrl }));
                                        } catch (err) {
                                          const reader = new FileReader();
                                          reader.onloadend = () => {
                                            setQuestionForm(prev => ({ ...prev, gambarUrl: reader.result as string }));
                                          };
                                          reader.readAsDataURL(file);
                                        } finally {
                                          setIsCompressingImage(false);
                                        }
                                      }
                                    }} 
                                  />
                                </label>
                                {questionForm.gambarUrl && (
                                  <button
                                    type="button"
                                    onClick={() => setQuestionForm(prev => ({ ...prev, gambarUrl: '', gambarCaption: '' }))}
                                    className="bg-rose-50 hover:bg-rose-100 text-rose-700 px-2 py-1 rounded-lg text-xs font-bold border border-rose-200 cursor-pointer"
                                  >
                                    Hapus
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Action Buttons: Save & Cancel */}
                            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-200">
                              <button
                                type="button"
                                onClick={() => {
                                  setIsEditingQuestion(false);
                                  setEditingQuestionId(null);
                                }}
                                className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-4 py-2 rounded-xl text-xs transition cursor-pointer"
                              >
                                Batal
                              </button>
                              <button
                                type="submit"
                                disabled={isSavingQuestion}
                                className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-black px-5 py-2 rounded-xl text-xs flex items-center gap-2 shadow-md hover:scale-105 active:scale-95 transition cursor-pointer"
                              >
                                {isSavingQuestion ? (
                                  <>
                                    <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                                    <span>Menyimpan...</span>
                                  </>
                                ) : (
                                  <>
                                    <Save className="h-4 w-4" />
                                    <span>Simpan Perubahan Soal No. {q.noSoal}</span>
                                  </>
                                )}
                              </button>
                            </div>
                          </form>
                        </div>
                      ) : (
                        /* REGULAR QUESTION CARD VIEW MODE */
                        <>
                          {/* Question Header */}
                          {printConfig.showCompetencyTag ? (
                            <div className="flex flex-col sm:flex-row justify-between items-start gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200 text-[10px]">
                              <div className="space-y-0.5">
                                <div><span className="font-bold">No Soal:</span> <span className="font-mono bg-slate-200 px-1.5 py-0.5 rounded font-bold">{q.noSoal}</span></div>
                                <div><span className="font-bold">Kompetensi:</span> {q.kompetensi}</div>
                                <div><span className="font-bold">Sub Kompetensi:</span> {q.subKompetensi}</div>
                              </div>
                              <div className="flex items-center gap-2">
                                <span className="bg-indigo-100 text-indigo-800 font-bold px-2 py-0.5 rounded">
                                  {getBentukSoalLabel(q.shapes || q.bentukSoal)}
                                </span>
                                
                                {/* Control buttons inside card (hidden on print) */}
                                <div className="flex gap-1 no-print items-center">
                                  {deletingQuestionId === q.id ? (
                                    <div className="flex gap-1 items-center bg-rose-50 border border-rose-200 px-1.5 py-0.5 rounded-lg">
                                      <span className="text-[9px] font-bold text-rose-700">Hapus?</span>
                                      <button
                                        onClick={() => {
                                          handleDeleteQuestion(q.id);
                                          setDeletingQuestionId(null);
                                        }}
                                        className="bg-red-600 text-white font-extrabold px-1 py-0.5 rounded text-[9px] transition"
                                      >
                                        Ya
                                      </button>
                                      <button
                                        onClick={() => setDeletingQuestionId(null)}
                                        className="bg-slate-200 text-slate-700 font-bold px-1 py-0.5 rounded text-[9px] transition"
                                      >
                                        Batal
                                      </button>
                                    </div>
                                  ) : (
                                    <div className="flex gap-1.5 items-center">
                                      <button
                                        onClick={() => handleEditQuestion(q)}
                                        className="group flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-extrabold transition-all duration-200 shadow-xs hover:shadow-md cursor-pointer bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white border border-indigo-200 hover:scale-105 active:scale-95"
                                        title="Ubah Butir Soal Langsung di Tempat"
                                      >
                                        <Edit className="h-3 w-3 group-hover:rotate-12 transition-transform duration-200" />
                                        <span>Ubah Butir Soal</span>
                                      </button>
                                      <button
                                        onClick={() => setDeletingQuestionId(q.id)}
                                        className="flex items-center gap-1 bg-rose-50 text-rose-700 hover:bg-rose-100 border border-rose-200 px-2 py-1 rounded-lg text-[10px] font-bold transition shadow-xs cursor-pointer hover:scale-105 active:scale-95"
                                        title="Hapus Soal"
                                      >
                                        <Trash2 className="h-3 w-3" />
                                        <span>Hapus</span>
                                      </button>
                                    </div>
                                  )}
                                </div>
                              </div>
                            </div>
                          ) : (
                            /* Traditional exam number format if technical tags are hidden */
                            <div className="flex justify-between items-center no-print pb-1 border-b border-slate-100 text-[10px] text-slate-400">
                              <span>Metadata Soal #{q.noSoal} ({getBentukSoalLabel(q.bentukSoal)})</span>
                              <div className="flex gap-1 items-center">
                                {deletingQuestionId === q.id ? (
                                  <div className="flex gap-1 items-center bg-rose-50 border border-rose-200 px-1.5 py-0.5 rounded-lg">
                                    <span className="text-[9px] font-bold text-rose-700">Hapus?</span>
                                    <button
                                      onClick={() => {
                                        handleDeleteQuestion(q.id);
                                        setDeletingQuestionId(null);
                                      }}
                                      className="bg-red-600 text-white font-extrabold px-1 py-0.5 rounded text-[9px] transition"
                                    >
                                      Ya
                                    </button>
                                    <button
                                      onClick={() => setDeletingQuestionId(null)}
                                      className="bg-slate-200 text-slate-700 font-bold px-1 py-0.5 rounded text-[9px] transition"
                                    >
                                      Batal
                                    </button>
                                  </div>
                                ) : (
                                  <div className="flex gap-1.5 items-center">
                                    <button
                                      onClick={() => handleEditQuestion(q)}
                                      className="group flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[10px] font-extrabold transition-all duration-200 shadow-xs hover:shadow-md cursor-pointer bg-indigo-50 text-indigo-700 hover:bg-indigo-600 hover:text-white border border-indigo-200 hover:scale-105 active:scale-95"
                                      title="Ubah Butir Soal Langsung di Tempat"
                                    >
                                      <Edit className="h-3 w-3 group-hover:rotate-12 transition-transform duration-200" />
                                      <span>Ubah Butir Soal</span>
                                    </button>
                                    <button
                                      onClick={() => setDeletingQuestionId(q.id)}
                                      className="flex items-center gap-1 bg-rose-50 text-rose-700 hover:bg-rose-100 border border-rose-200 px-2 py-1 rounded-lg text-[10px] font-bold transition shadow-xs cursor-pointer hover:scale-105 active:scale-95"
                                      title="Hapus Soal"
                                    >
                                      <Trash2 className="h-3 w-3" />
                                      <span>Hapus</span>
                                    </button>
                                  </div>
                                )}
                              </div>
                            </div>
                          )}

                          {/* Question Content Block */}
                          <div className="flex gap-2.5 items-start">
                            {!printConfig.showCompetencyTag && (
                              <span className="font-bold font-mono text-sm text-slate-900 min-w-[20px]">{q.noSoal}.</span>
                            )}
                            
                            <div className="flex-1 space-y-3.5">
                              {/* Stimulus Box */}
                              {q.stimulus && printConfig.showStimulus && (
                                <div className="text-slate-700 leading-relaxed font-normal italic text-xs sm:text-sm text-justify mb-2">
                                  {q.stimulus}
                                </div>
                              )}

                              {/* Illustration / Image Box */}
                              {q.gambarUrl && q.gambarUrl.trim() !== '' && printConfig.showIllustration && (() => {
                                const alignClass = q.gambarPosisi === 'left' ? 'items-start text-left' : q.gambarPosisi === 'right' ? 'items-end text-right' : 'items-center text-center';
                                const sizeClass = q.gambarUkuran === 'small' ? 'max-w-[160px]' : q.gambarUkuran === 'large' ? 'max-w-md' : q.gambarUkuran === 'full' ? 'max-w-full' : 'max-w-xs';

                                return (
                                  <div className={`my-2 bg-slate-50/80 border border-slate-200/60 rounded-2xl p-3 flex flex-col ${alignClass} space-y-1.5 break-inside-avoid transition`}>
                                    {q.gambarUrl.trim().toLowerCase().startsWith('<svg') ? (
                                      <div className={`w-full ${sizeClass} overflow-x-auto flex justify-center py-2 px-3 bg-white rounded-xl border border-slate-100 shadow-xs relative group`}>
                                        <div dangerouslySetInnerHTML={{ __html: q.gambarUrl }} />
                                      </div>
                                    ) : (
                                      <div className={`relative group ${sizeClass} w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xs`}>
                                        <img 
                                          src={q.gambarUrl} 
                                          alt={`Ilustrasi Soal ${q.noSoal}`}
                                          referrerPolicy="no-referrer"
                                          className="w-full h-auto object-contain max-h-[260px] mx-auto transition-transform duration-300 group-hover:scale-102 cursor-pointer"
                                          onClick={() => {
                                            setZoomScale(1);
                                            setZoomRotation(0);
                                            setActiveZoomImage({ url: q.gambarUrl!, caption: q.gambarCaption || `Gambar Ilustrasi Soal #${q.noSoal}` });
                                          }}
                                        />
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setZoomScale(1);
                                            setZoomRotation(0);
                                            setActiveZoomImage({ url: q.gambarUrl!, caption: q.gambarCaption || `Gambar Ilustrasi Soal #${q.noSoal}` });
                                          }}
                                          className="no-print absolute top-2 right-2 bg-slate-900/80 hover:bg-slate-900 text-white p-1.5 rounded-lg opacity-0 group-hover:opacity-100 transition-opacity duration-200 flex items-center gap-1 text-[10px] font-bold shadow-md cursor-pointer backdrop-blur-xs"
                                          title="Klik untuk Perbesar Gambar (Zoom Lightbox)"
                                        >
                                          <ZoomIn className="h-3.5 w-3.5" />
                                          <span>Zoom</span>
                                        </button>
                                      </div>
                                    )}

                                    {q.gambarCaption && (
                                      <p className="text-[11px] font-semibold text-slate-600 italic tracking-tight max-w-md">
                                        {q.gambarCaption}
                                      </p>
                                    )}
                                  </div>
                                );
                              })()}

                              {/* Question Statement */}
                              <div className="text-slate-900 leading-relaxed font-semibold text-xs sm:text-sm">
                                {printConfig.showCompetencyTag && <span className="font-bold mr-1">{q.noSoal}.</span>}
                                <SimpleMarkdown content={cleanSoalText(q.soal)} />
                              </div>

                              {/* Options */}
                              {q.opsi && q.opsi.length > 0 && (() => {
                                const isKategori = isKategoriSoal(q);

                                if (isKategori) {
                                  const { cat1, cat2 } = getPgkCategories(q.soal, q.kunciJawaban);
                                  const validOpts = q.opsi.filter(opt => {
                                    const txt = cleanOptionText(opt).trim();
                                    return txt !== '' && txt !== '-' && txt !== '–' && txt !== '—';
                                  });

                                  return (
                                    <div className="my-3 overflow-x-auto rounded-xl border border-slate-300 shadow-xs bg-white">
                                      <table className="w-full text-xs text-left border-collapse min-w-[500px]">
                                        <thead>
                                          <tr className="bg-slate-100/90 text-slate-900 border-b-2 border-slate-900 font-extrabold text-[11px] sm:text-xs">
                                            <th className="py-2.5 px-3 w-10 text-center border-r border-slate-300 font-black">#</th>
                                            <th className="py-2.5 px-3 border-r border-slate-300 font-black">Pernyataan</th>
                                            <th className="py-2.5 px-3 w-28 sm:w-36 text-center border-r border-slate-300 bg-slate-200/70 font-black text-slate-900">{cat1}</th>
                                            <th className="py-2.5 px-3 w-28 sm:w-36 text-center bg-slate-200/70 font-black text-slate-900">{cat2}</th>
                                          </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-200">
                                          {validOpts.map((opt, i) => {
                                            const optLetter = String.fromCharCode(65 + i);
                                            const optText = cleanOptionText(opt);
                                            const catSelected = getPgkCategoryIndex(q.kunciJawaban, i, optLetter, cat1, cat2);
                                            const showKey = printConfig.showAnswerKey;
                                            const isCat1Selected = showKey && catSelected === 1;
                                            const isCat2Selected = showKey && catSelected === 2;

                                            return (
                                              <tr key={i} className={`hover:bg-slate-50 transition-colors ${showKey && (isCat1Selected || isCat2Selected) ? 'bg-emerald-50/50' : ''}`}>
                                                <td className="py-2.5 px-3 text-center font-bold font-mono text-slate-800 border-r border-slate-200 align-middle">
                                                  {i + 1}.
                                                </td>
                                                <td className="py-2.5 px-3 text-slate-900 leading-relaxed font-normal border-r border-slate-200 align-middle text-xs sm:text-[13px]">
                                                  {renderInlines(optText)}
                                                </td>
                                                <td className={`py-2.5 px-3 text-center border-r border-slate-200 align-middle ${isCat1Selected ? 'bg-emerald-100/80 font-bold' : ''}`}>
                                                  <div className="flex justify-center items-center">
                                                    <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                                                      isCat1Selected 
                                                        ? 'border-emerald-700 bg-emerald-600 text-white ring-2 ring-emerald-200 shadow-xs' 
                                                        : 'border-slate-400 bg-white'
                                                    }`}>
                                                      {isCat1Selected && <div className="w-2 h-2 rounded-full bg-white" />}
                                                    </div>
                                                  </div>
                                                </td>
                                                <td className={`py-2.5 px-3 text-center align-middle ${isCat2Selected ? 'bg-emerald-100/80 font-bold' : ''}`}>
                                                  <div className="flex justify-center items-center">
                                                    <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                                                      isCat2Selected 
                                                        ? 'border-emerald-700 bg-emerald-600 text-white ring-2 ring-emerald-200 shadow-xs' 
                                                        : 'border-slate-400 bg-white'
                                                    }`}>
                                                      {isCat2Selected && <div className="w-2 h-2 rounded-full bg-white" />}
                                                    </div>
                                                  </div>
                                                </td>
                                              </tr>
                                            );
                                          })}
                                        </tbody>
                                      </table>
                                    </div>
                                  );
                                }

                                return (
                                  <div className={`grid gap-2 pl-2 ${printConfig.layoutColumns === '2' ? 'grid-cols-1' : 'grid-cols-1 md:grid-cols-2'}`}>
                                    {q.opsi.map((opt, i) => {
                                      const optLetter = String.fromCharCode(65 + i);
                                      const optText = cleanOptionText(opt);
                                      const isCorrectOption = q.kunciJawaban.trim().toUpperCase().includes(optLetter) && printConfig.showAnswerKey;
                                      return (
                                        <div 
                                          key={i} 
                                          className={`flex items-start gap-2 p-1.5 rounded-lg border text-xs transition-all ${
                                            isCorrectOption 
                                              ? 'bg-emerald-50 border-emerald-300 text-emerald-950 font-semibold' 
                                              : 'border-slate-100 bg-slate-50/40 text-slate-800'
                                          }`}
                                        >
                                          <span className={`w-4.5 h-4.5 rounded-full flex items-center justify-center font-mono font-bold text-[9px] flex-shrink-0 ${
                                            isCorrectOption
                                              ? 'bg-emerald-500 text-white shadow-xs'
                                              : 'bg-slate-200 text-slate-700'
                                          }`}>
                                            {optLetter}
                                          </span>
                                          <span className="font-sans leading-relaxed text-[11px] sm:text-xs">{renderInlines(optText)}</span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                );
                              })()}

                              {/* Answer Key & Explanation */}
                              {printConfig.showAnswerKey && (
                                <motion.div 
                                  initial={{ opacity: 0, y: 5 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  className="bg-emerald-50/40 p-3 rounded-xl border border-dashed border-emerald-300 space-y-1.5 text-xs"
                                >
                                  <div className="flex items-center gap-2">
                                    <span className="text-[10px] font-bold text-emerald-800 uppercase tracking-wider">Kunci Jawaban:</span>
                                    <span className="font-mono bg-emerald-100 text-emerald-900 border border-emerald-200 text-xs font-extrabold px-2.5 py-0.5 rounded-full">
                                      {q.kunciJawaban}
                                    </span>
                                  </div>
                                  
                                  {q.kataKunci && (
                                    <div className="flex flex-wrap items-center gap-2">
                                      <span className="text-[10px] font-bold text-indigo-800 uppercase tracking-wider">Materi / Konsep:</span>
                                      <span className="bg-indigo-100 text-indigo-950 border border-indigo-200 text-[10px] font-semibold px-2 py-0.5 rounded">
                                        {q.kataKunci}
                                      </span>
                                    </div>
                                  )}
                                  
                                  <div className="text-[11px] text-slate-700 leading-relaxed">
                                    <span className="font-bold text-slate-800">Pembahasan Ilmiah:</span>
                                    <div className="whitespace-pre-wrap mt-0.5">{renderInlines(cleanSoalText(q.pembahasan))}</div>
                                  </div>
                                </motion.div>
                              )}

                            </div>
                          </div>
                        </>
                      )}
                    </div>
                    ))}
                  </div>
                );
              })()}
            </div>
      </main>

      {/* Footer */}
      <footer id="footer-section" className="bg-slate-900 text-slate-400 text-xs py-8 border-t border-slate-800 no-print mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-8 flex flex-col md:flex-row items-center justify-between gap-4">
          <div>
            <p className="font-bold text-slate-200 mb-1">Generator Kisi-Kisi & Pembuat Soal TKA SMA</p>
            <p>Sistem Asesmen Pintar untuk Guru, Dosen, dan Pengajar Seluruh Indonesia.</p>
          </div>
          <div className="text-right flex flex-col items-end gap-1">
            <p className="font-extrabold text-slate-300">© 2026 @AJISOSIOLOGI</p>
            <p className="text-[11px] text-slate-500">Assessment TKA SMA. All Rights Reserved.</p>
            <p>Dikembangkan dengan menggunakan Gemini Flash & React.</p>
            <a 
              href="https://lynk.id/ajisosiologi" 
              target="_blank" 
              rel="noopener noreferrer"
              className="inline-flex items-center gap-1 px-3 py-1 bg-slate-800 hover:bg-indigo-950 text-[10px] font-extrabold text-indigo-400 hover:text-indigo-300 rounded-lg border border-slate-700 hover:border-indigo-500/50 transition duration-200 mt-1"
            >
              <span>Create @ajisosiologi</span>
            </a>
          </div>
        </div>
      </footer>

      {/* Dynamic Animated AI Question Generation Progress Modal */}
      <AnimatePresence>
        {soalProgress.active && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/85 backdrop-blur-md z-[9999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.92, y: 15, opacity: 0 }}
              animate={{ scale: 1, y: 0, opacity: 1 }}
              exit={{ scale: 0.92, y: 15, opacity: 0 }}
              transition={{ type: "spring", duration: 0.4 }}
              className="bg-white rounded-3xl shadow-2xl p-6 md:p-8 max-w-md w-full text-center space-y-6 border border-slate-100 relative overflow-hidden"
            >
              {/* Top accent glow line */}
              <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500" />
              <div className="absolute -top-12 -left-12 w-32 h-32 bg-indigo-100 rounded-full blur-3xl opacity-60" />
              <div className="absolute -bottom-12 -right-12 w-32 h-32 bg-emerald-100 rounded-full blur-3xl opacity-60" />

              {/* Loader Animation */}
              <div className="relative w-24 h-24 mx-auto flex items-center justify-center">
                <div className="absolute inset-0 rounded-full border-4 border-slate-100 border-t-indigo-600 animate-spin" />
                <div className="absolute inset-2 rounded-full border-4 border-slate-100 border-b-purple-500 animate-spin" style={{ animationDirection: 'reverse', animationDuration: '1.2s' }} />
                <div className="w-14 h-14 rounded-full bg-indigo-50 flex items-center justify-center animate-pulse">
                  <Sparkles className="h-7 w-7 text-indigo-600 animate-bounce" style={{ animationDuration: '2s' }} />
                </div>
              </div>

              {/* Titles */}
              <div className="space-y-2">
                <h3 className="text-lg font-extrabold text-slate-800 tracking-tight">
                  Penyusunan Soal Otomatis oleh AI Gemini
                </h3>
                <p className="text-xs text-slate-500 leading-relaxed px-2">
                  Asisten AI sedang menyelaraskan materi asesmen dan merumuskan butir soal berstandar tinggi (HOTS, Literasi & Numerasi).
                </p>
              </div>

              {/* Progress Indicator */}
              {soalProgress.totalQuestions > 0 && (
                <div className="space-y-2.5">
                  <div className="flex items-center justify-between text-xs font-bold text-slate-500">
                    <span className="uppercase tracking-wider text-[10px]">Progress Penyusunan</span>
                    <span className="text-indigo-600 font-mono text-sm font-bold">
                      {Math.min(100, Math.round((soalProgress.countSuccess / soalProgress.totalQuestions) * 100))}%
                    </span>
                  </div>
                  
                  {/* Progress Track */}
                  <div className="w-full bg-slate-100 h-3 rounded-full overflow-hidden p-0.5 border border-slate-200/50">
                    <motion.div 
                      className="bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 h-full rounded-full"
                      initial={{ width: 0 }}
                      animate={{ width: `${Math.min(100, Math.round((soalProgress.countSuccess / soalProgress.totalQuestions) * 100))}%` }}
                      transition={{ duration: 0.3, ease: "easeOut" }}
                    />
                  </div>
                  
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>Target: {soalProgress.totalQuestions} Soal</span>
                    <span className="text-emerald-600">Selesai: {soalProgress.countSuccess} Soal</span>
                  </div>
                </div>
              )}

              {/* Current Status Box */}
              <div className="bg-slate-50 border border-slate-150 rounded-2xl p-4 text-left space-y-2 relative">
                {soalProgress.totalNo > 1 && (
                  <div className="text-[10px] font-extrabold text-indigo-600 uppercase tracking-widest flex items-center gap-1">
                    <Layers className="h-3.5 w-3.5 text-indigo-500" />
                    <span>Kisi-Kisi ke-{soalProgress.currentNo} dari {soalProgress.totalNo}</span>
                  </div>
                )}
                
                {soalProgress.topic && (
                  <div className="text-xs font-extrabold text-slate-800 flex items-center gap-1.5 leading-snug">
                    <BookOpen className="h-4 w-4 text-indigo-600 flex-shrink-0" />
                    <span className="truncate">{soalProgress.topic}</span>
                  </div>
                )}

                <div className="text-xs text-slate-600 font-semibold flex items-center gap-2 mt-1">
                  <span className="relative flex h-2.5 w-2.5">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
                  </span>
                  <span className="italic text-[11px] text-slate-700">{soalProgress.statusText}</span>
                </div>
              </div>

              {/* Helpful standard tip */}
              <div className="text-[10px] text-slate-400 bg-slate-50/50 p-2.5 rounded-xl border border-slate-100 flex items-center justify-center gap-1.5">
                <CheckSquare className="h-3.5 w-3.5 text-emerald-500 flex-shrink-0" />
                <span>Mengecek distractor & merumuskan kunci pembahasan ilmiah.</span>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Auto Prompt Generator Modal */}
      <AnimatePresence>
        {isPromptModalOpen && selectedKisiForPrompt && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/85 backdrop-blur-md z-[9999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.92, y: 15, opacity: 0 }}
              animate={{ scale: 1, y: 0, opacity: 1 }}
              exit={{ scale: 0.92, y: 15, opacity: 0 }}
              transition={{ type: "spring", duration: 0.4 }}
              className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full border border-slate-100 relative overflow-hidden flex flex-col max-h-[90vh]"
            >
              {/* Top accent glow line */}
              <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500" />
              
              {/* Header */}
              <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="w-9 h-9 rounded-xl bg-indigo-50 flex items-center justify-center">
                    <FileText className="h-5 w-5 text-indigo-600" />
                  </div>
                  <div>
                    <h3 className="text-base font-extrabold text-slate-800 tracking-tight">
                      Pembuat Prompt Otomatis AI (Megaprompt)
                    </h3>
                    <p className="text-[11px] text-slate-500 font-medium">
                      Salin prompt terstruktur di bawah ini untuk digunakan di Gemini, ChatGPT, Claude, dll.
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsPromptModalOpen(false)}
                  className="text-slate-400 hover:text-slate-600 font-extrabold text-lg p-1"
                >
                  ✕
                </button>
              </div>

              {/* Content body - Scrollable */}
              <div className="p-6 overflow-y-auto space-y-4 flex-1">
                {/* Info Badges of current Kisi row */}
                <div className="bg-slate-50 border border-slate-150 rounded-2xl p-4 space-y-2.5 text-left">
                  <span className="text-[10px] font-extrabold text-indigo-600 uppercase tracking-widest block">
                    Spesifikasi Kisi-Kisi No. {selectedKisiForPrompt.no}
                  </span>
                  <div className="grid grid-cols-2 gap-2 text-xs">
                    <div className="bg-white px-3 py-2 rounded-xl border border-slate-100">
                      <span className="text-[10px] text-slate-400 font-bold block uppercase">Mata Pelajaran</span>
                      <span className="font-bold text-slate-700 truncate block">{config.mataPelajaran}</span>
                    </div>
                    <div className="bg-white px-3 py-2 rounded-xl border border-slate-100">
                      <span className="text-[10px] text-slate-400 font-bold block uppercase">Materi Pokok</span>
                      <span className="font-bold text-slate-700 truncate block">{selectedKisiForPrompt.elemenMateri}</span>
                    </div>
                    <div className="bg-white px-3 py-2 rounded-xl border border-slate-100 col-span-2">
                      <span className="text-[10px] text-slate-400 font-bold block uppercase">Sub-Materi / Indikator</span>
                      <span className="font-bold text-slate-700 block">{selectedKisiForPrompt.subElemenMateri || '-'}</span>
                    </div>
                    <div className="bg-white px-3 py-2 rounded-xl border border-slate-100">
                      <span className="text-[10px] text-slate-400 font-bold block uppercase">Level Kognitif</span>
                      <span className="font-bold text-slate-700 block truncate">
                        {getLevelKognitifLabel(selectedKisiForPrompt.levelKognitif)}
                      </span>
                    </div>
                    <div className="bg-white px-3 py-2 rounded-xl border border-slate-100">
                      <span className="text-[10px] text-slate-400 font-bold block uppercase">Bentuk & Jumlah</span>
                      <span className="font-bold text-slate-700 block truncate">
                        {getBentukSoalLabel(selectedKisiForPrompt.bentukSoal)} ({selectedKisiForPrompt.jumlahSoal || 5} Soal)
                      </span>
                    </div>
                  </div>
                </div>

                {/* Prompt Text Box */}
                <div className="space-y-1.5 text-left">
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] font-extrabold text-slate-500 uppercase tracking-wider">
                      Draf Prompt AI Anda
                    </span>
                    <button
                      onClick={handleOptimizePromptWithAi}
                      disabled={isGeneratingPrompt}
                      className="text-[11px] font-extrabold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded-lg transition"
                    >
                      {isGeneratingPrompt ? (
                        <>
                          <RefreshCw className="h-3 w-3 animate-spin" />
                          <span>Menganalisis & Mengoptimasi...</span>
                        </>
                      ) : (
                        <>
                          <Sparkles className="h-3 w-3" />
                          <span>Optimasi via AI (Megaprompt)</span>
                        </>
                      )}
                    </button>
                  </div>

                  <div className="relative group">
                    <textarea
                      readOnly
                      value={generatedPromptText}
                      className="w-full h-64 bg-slate-900 text-slate-100 p-4 rounded-2xl font-mono text-xs leading-relaxed focus:outline-none border border-slate-800 resize-none"
                    />
                    <div className="absolute top-3 right-3 opacity-90 group-hover:opacity-100 transition">
                      <button
                        onClick={() => {
                          navigator.clipboard.writeText(generatedPromptText);
                          setCopiedPrompt(true);
                          setTimeout(() => setCopiedPrompt(false), 2000);
                        }}
                        className={`p-2 rounded-xl flex items-center gap-1 text-xs font-bold transition shadow ${
                          copiedPrompt 
                            ? 'bg-emerald-600 text-white' 
                            : 'bg-slate-800 text-slate-200 hover:bg-slate-750'
                        }`}
                      >
                        {copiedPrompt ? (
                          <>
                            <Check className="h-3.5 w-3.5" />
                            <span>Tersalin!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="h-3.5 w-3.5" />
                            <span>Salin Prompt</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Explanation Card */}
                <div className="bg-indigo-50/50 border border-indigo-100 rounded-2xl p-4 flex gap-3 text-left">
                  <Info className="h-5 w-5 text-indigo-600 flex-shrink-0 mt-0.5" />
                  <div className="space-y-1 text-xs text-indigo-900 leading-normal font-medium">
                    <p className="font-bold">💡 Apa keuntungan menggunakan Prompt ini?</p>
                    <p className="text-slate-600">
                      Anda bisa menempelkan prompt ini pada platform AI luar untuk memperoleh materi penunjang pembelajaran lainnya, merancang bank soal alternatif yang sinkron dengan kurikulum, atau melatih pemahaman Anda secara mandiri di browser Anda sendiri.
                    </p>
                  </div>
                </div>
              </div>

              {/* Footer */}
              <div className="p-4 bg-slate-50 border-t border-slate-100 flex justify-end gap-2.5">
                <button
                  onClick={() => setIsPromptModalOpen(false)}
                  className="px-4 py-2 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition"
                >
                  Tutup
                </button>
                <button
                  onClick={() => {
                    navigator.clipboard.writeText(generatedPromptText);
                    setCopiedPrompt(true);
                    setTimeout(() => setCopiedPrompt(false), 2000);
                  }}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5"
                >
                  <Copy className="h-3.5 w-3.5" />
                  <span>{copiedPrompt ? 'Berhasil Disalin!' : 'Salin & Mulai'}</span>
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
        {/* Master Megaprompt AI Modal (Entire Kisi-Kisi Table Combined) */}
        {isMasterMegapromptModalOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/85 backdrop-blur-md z-[9999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.92, y: 15, opacity: 0 }}
              animate={{ scale: 1, y: 0, opacity: 1 }}
              exit={{ scale: 0.92, y: 15, opacity: 0 }}
              transition={{ type: "spring", duration: 0.4 }}
              className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full border border-slate-100 relative overflow-hidden flex flex-col max-h-[92vh]"
            >
              {/* Top accent glow line */}
              <div className="absolute top-0 left-0 right-0 h-2 bg-gradient-to-r from-amber-400 via-indigo-600 to-emerald-500" />

              {/* Modal Header */}
              <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-amber-100 border border-amber-200 flex items-center justify-center shadow-sm">
                    <Sparkles className="h-5 w-5 text-amber-700" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <h3 className="text-base font-extrabold text-slate-900 tracking-tight">
                        Pembuat Prompt Otomatis AI (Master Megaprompt)
                      </h3>
                      <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-amber-400 text-slate-950 uppercase tracking-wider">
                        {kisiList.length} Baris Kisi-Kisi
                      </span>
                      <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                        ✨ Output 2 Format: Word (.docx) & Excel (.xlsx)
                      </span>
                    </div>
                    <p className="text-xs text-slate-500 font-medium">
                      Menggabungkan seluruh matriks asesmen menjadi 1 Prompt Utama terstruktur presisi untuk AI
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsMasterMegapromptModalOpen(false)}
                  className="text-slate-400 hover:text-slate-600 font-extrabold text-xl p-1.5 rounded-xl hover:bg-slate-200/60 transition"
                >
                  ✕
                </button>
              </div>

              {/* Modal Body */}
              <div className="p-6 overflow-y-auto space-y-5 flex-1">
                {/* Stats Summary Bar */}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                  <div className="bg-slate-50 border border-slate-200/80 p-3 rounded-2xl">
                    <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider block">Mata Pelajaran</span>
                    <span className="text-xs font-black text-slate-800 truncate block mt-0.5">{config.mataPelajaran || 'Sosiologi'}</span>
                  </div>
                  <div className="bg-indigo-50/60 border border-indigo-100 p-3 rounded-2xl">
                    <span className="text-[10px] font-extrabold text-indigo-500 uppercase tracking-wider block">Total Kisi-Kisi</span>
                    <span className="text-sm font-black text-indigo-900 mt-0.5 block">{kisiList.length} Baris Spesifikasi</span>
                  </div>
                  <div className="bg-amber-50/60 border border-amber-100 p-3 rounded-2xl">
                    <span className="text-[10px] font-extrabold text-amber-600 uppercase tracking-wider block">Target Jumlah Soal</span>
                    <span className="text-sm font-black text-amber-900 mt-0.5 block">
                      {kisiList.reduce((acc, k) => acc + (k.jumlahSoal || 1), 0)} Butir Soal Total
                    </span>
                  </div>
                  <div className="bg-emerald-50/60 border border-emerald-100 p-3 rounded-2xl">
                    <span className="text-[10px] font-extrabold text-emerald-600 uppercase tracking-wider block">Level Kognitif</span>
                    <span className="text-xs font-black text-emerald-900 mt-0.5 block truncate">
                      L1: {kisiList.filter(k => k.levelKognitif === 'level_1').length} | L2: {kisiList.filter(k => k.levelKognitif === 'level_2').length} | L3: {kisiList.filter(k => k.levelKognitif === 'level_3').length}
                    </span>
                  </div>
                </div>

                {/* Preset Style Selector */}
                <div className="space-y-2">
                  <label className="text-[11px] font-extrabold text-slate-600 uppercase tracking-wider block">
                    Pilih Gaya & Karakter Megaprompt AI:
                  </label>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                    {[
                      { id: 'pusmendik', label: '🎓 Standard Pusmendik', desc: 'Baku & Seimbang' },
                      { id: 'hots', label: '🔥 HOTS & Kasus', desc: 'Penalaran Tinggi' },
                      { id: 'snbt', label: '🎯 UTBK-SNBT / TKA', desc: 'Literasi & Solutif' },
                      { id: 'variasi', label: '📝 Variasi Multi-Level', desc: 'Mudah, Sedang, HOTS' }
                    ].map((preset) => (
                      <button
                        key={preset.id}
                        onClick={() => {
                          const newStyle = preset.id as any;
                          setMasterMegapromptStyle(newStyle);
                          setMasterMegapromptText(buildMasterMegaprompt(kisiList, config, newStyle));
                        }}
                        className={`p-2.5 rounded-2xl border text-left transition-all ${
                          masterMegapromptStyle === preset.id
                            ? 'bg-indigo-600 text-white border-indigo-600 shadow-md ring-2 ring-indigo-200'
                            : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                        }`}
                      >
                        <span className="text-xs font-extrabold block">{preset.label}</span>
                        <span className={`text-[10px] block mt-0.5 font-medium ${masterMegapromptStyle === preset.id ? 'text-indigo-100' : 'text-slate-400'}`}>
                          {preset.desc}
                        </span>
                      </button>
                    ))}
                  </div>
                </div>

                {/* Text Area Container */}
                <div className="space-y-2 text-left">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <FileText className="h-4 w-4 text-indigo-600" />
                      Draf Master Megaprompt AI ({kisiList.length} Baris Total)
                    </span>
                    <span className="text-[11px] text-slate-400 font-mono">
                      {masterMegapromptText.length} Karakter | {masterMegapromptText.split(/\s+/).length} Kata
                    </span>
                  </div>

                  <div className="relative group">
                    <textarea
                      readOnly
                      value={masterMegapromptText || buildMasterMegaprompt(kisiList, config, masterMegapromptStyle)}
                      className="w-full h-80 bg-slate-950 text-slate-100 p-4 rounded-2xl font-mono text-xs leading-relaxed focus:outline-none border border-slate-800 resize-none shadow-inner"
                    />
                    <div className="absolute top-3 right-3 flex items-center gap-2">
                      <button
                        onClick={() => {
                          const text = masterMegapromptText || buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
                          navigator.clipboard.writeText(text);
                          setCopiedMasterMegaprompt(true);
                          setTimeout(() => setCopiedMasterMegaprompt(false), 2000);
                        }}
                        className={`px-3 py-1.5 rounded-xl flex items-center gap-1.5 text-xs font-bold transition shadow ${
                          copiedMasterMegaprompt 
                            ? 'bg-emerald-600 text-white' 
                            : 'bg-slate-800 text-slate-200 hover:bg-slate-700'
                        }`}
                      >
                        {copiedMasterMegaprompt ? (
                          <>
                            <Check className="h-3.5 w-3.5" />
                            <span>Tersalin!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="h-3.5 w-3.5" />
                            <span>Salin Prompt</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Action Toolbar Inside Modal */}
                <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-4 flex flex-col sm:flex-row items-center justify-between gap-3">
                  <div className="flex items-center gap-2 w-full sm:w-auto">
                    <button
                      onClick={downloadMasterMegapromptAsTxt}
                      className="flex-1 sm:flex-none px-3.5 py-2 bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1.5 shadow-sm"
                    >
                      <Download className="h-3.5 w-3.5 text-indigo-600" />
                      <span>Unduh File (.txt)</span>
                    </button>
                    <button
                      onClick={handleSyncToWadah1}
                      className="flex-1 sm:flex-none px-3.5 py-2 bg-indigo-50 border border-indigo-200 text-indigo-900 hover:bg-indigo-100 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1.5"
                    >
                      <Layers className="h-3.5 w-3.5 text-indigo-600" />
                      <span>{syncedToWadah1 ? 'Tersinkron!' : 'Sinkronkan ke Wadah 1 (CBT)'}</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Modal Footer */}
              <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
                <span className="text-[11px] text-slate-400 font-medium">
                  💡 Prompt ini menggabungkan seluruh {kisiList.length} baris matriks secara sistematis.
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setIsMasterMegapromptModalOpen(false)}
                    className="px-4 py-2 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-200 transition"
                  >
                    Tutup
                  </button>
                  <button
                    onClick={() => {
                      const text = masterMegapromptText || buildMasterMegaprompt(kisiList, config, masterMegapromptStyle);
                      navigator.clipboard.writeText(text);
                      setCopiedMasterMegaprompt(true);
                      setTimeout(() => setCopiedMasterMegaprompt(false), 2000);
                    }}
                    className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-extrabold transition flex items-center gap-2 shadow"
                  >
                    <Copy className="h-3.5 w-3.5" />
                    <span>{copiedMasterMegaprompt ? 'Berhasil Disalin!' : 'Salin Megaprompt Utama'}</span>
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Upload Excel / Import Soal AI Modal */}
        {isImportAiModalOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/85 backdrop-blur-md z-[9999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.92, y: 15, opacity: 0 }}
              animate={{ scale: 1, y: 0, opacity: 1 }}
              exit={{ scale: 0.92, y: 15, opacity: 0 }}
              transition={{ type: "spring", duration: 0.4 }}
              className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full border border-slate-100 relative overflow-hidden flex flex-col max-h-[92vh]"
            >
              {/* Header glow */}
              <div className="absolute top-0 left-0 right-0 h-2 bg-gradient-to-r from-emerald-500 via-teal-500 to-indigo-600" />

              {/* Modal Header */}
              <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-100 border border-emerald-200 flex items-center justify-center shadow-sm">
                    <FileSpreadsheet className="h-5 w-5 text-emerald-700" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2">
                      <h3 className="text-base font-extrabold text-slate-900 tracking-tight">
                        Upload File Excel / Hasil Megaprompt Gemini AI
                      </h3>
                      <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-200 uppercase tracking-wider">
                        Menu 3. Butir Soal TKA SMA
                      </span>
                    </div>
                    <p className="text-xs text-slate-500 font-medium">
                      Unggah file Excel (.xlsx / .csv), Word (.docx), atau tempelkan teks hasil salinan dari Gemini AI
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setIsImportAiModalOpen(false)}
                  className="text-slate-400 hover:text-slate-600 font-extrabold text-xl p-1.5 rounded-xl hover:bg-slate-200/60 transition"
                >
                  ✕
                </button>
              </div>

              {/* Modal Body */}
              <div className="p-6 overflow-y-auto space-y-5 flex-1">
                
                {/* File Upload Zone + Download Template Row */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="md:col-span-2 border-2 border-dashed border-emerald-300 hover:border-emerald-500 bg-emerald-50/30 hover:bg-emerald-50/60 rounded-2xl p-5 transition flex flex-col items-center justify-center text-center cursor-pointer relative group">
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv,.docx,.doc,.txt,.json,.md"
                      onChange={handleFileUploadForImportAi}
                      className="absolute inset-0 opacity-0 cursor-pointer w-full h-full z-10"
                    />
                    <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center mb-2 shadow-sm group-hover:scale-105 transition">
                      <Upload className="h-6 w-6" />
                    </div>
                    <p className="text-xs font-black text-slate-800">
                      Pilih atau Seret File Excel (.xlsx / .csv) / Word / Text di Sini
                    </p>
                    <p className="text-[10.5px] text-slate-500 mt-1">
                      Mendukung format <strong>.xlsx, .xls, .csv, .docx, .txt</strong> hasil olahan Gemini AI
                    </p>
                  </div>

                  <div className="bg-gradient-to-br from-indigo-50 via-slate-50 to-emerald-50 border border-slate-200/80 rounded-2xl p-4 flex flex-col justify-between">
                    <div>
                      <span className="text-[10px] font-extrabold text-indigo-600 uppercase tracking-widest block mb-1">
                        Format Standar Excel
                      </span>
                      <h4 className="text-xs font-bold text-slate-800">Belum punya template Excel?</h4>
                      <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">
                        Unduh template Excel dengan 14 kolom standar TKA SMA yang sudah dilengkapi contoh isian.
                      </p>
                    </div>
                    <button
                      onClick={() => downloadTemplateExcelSoal(config.mataPelajaran)}
                      className="mt-3 w-full bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs py-2 px-3 rounded-xl flex items-center justify-center gap-1.5 transition shadow-sm cursor-pointer"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>Download Template Excel</span>
                    </button>
                  </div>
                </div>

                {importAiFileName && (
                  <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-3 flex items-center justify-between text-xs">
                    <div className="flex items-center gap-2 text-emerald-900 font-bold">
                      <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                      <span>File Terpilih: {importAiFileName}</span>
                    </div>
                    <span className="bg-emerald-600 text-white text-[10px] font-black px-2.5 py-0.5 rounded-full">
                      Format: {importAiDetectedFormat || 'Terdeteksi'}
                    </span>
                  </div>
                )}

                {importAiError && (
                  <div className="bg-rose-50 border border-rose-200 rounded-2xl p-3 text-xs text-rose-700 font-medium">
                    ⚠️ {importAiError}
                  </div>
                )}

                {/* Direct Paste Text Area */}
                <div className="space-y-1.5 text-left">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <FileText className="h-4 w-4 text-indigo-600" />
                      Atau Tempelkan Teks Hasil Megaprompt Gemini AI Langsung:
                    </label>
                    {importAiDetectedFormat && (
                      <span className="text-[10px] font-bold bg-indigo-50 text-indigo-700 px-2.5 py-0.5 rounded-full border border-indigo-100">
                        Format Terdeteksi: {importAiDetectedFormat}
                      </span>
                    )}
                  </div>
                  <textarea
                    value={importAiRawText}
                    onChange={(e) => {
                      const text = e.target.value;
                      setImportAiRawText(text);
                      const parsed = parseAiQuestionsText(text);
                      setImportAiParsedQuestions(parsed);
                    }}
                    placeholder="Tempelkan naskah soal atau tabel markdown/tsv hasil jawaban Gemini AI di sini..."
                    className="w-full h-36 bg-slate-900 text-slate-100 p-3.5 rounded-2xl font-mono text-xs leading-relaxed focus:outline-none border border-slate-800 resize-none shadow-inner"
                  />
                </div>

                {/* Preview Table of Parsed Questions */}
                {importAiParsedQuestions.length > 0 ? (
                  <div className="space-y-2">
                    <div className="flex items-center justify-between bg-emerald-50/80 border border-emerald-200 p-3 rounded-2xl">
                      <div className="flex items-center gap-2">
                        <Sparkles className="h-4 w-4 text-emerald-600" />
                        <span className="text-xs font-black text-emerald-950">
                          🎉 Terdeteksi {importAiParsedQuestions.length} Butir Soal TKA SMA Siap Diimpor
                        </span>
                      </div>
                      <span className="text-[10px] font-bold text-slate-500">
                        Preview Butir Soal
                      </span>
                    </div>

                    <div className="border border-slate-200 rounded-2xl overflow-hidden max-h-56 overflow-y-auto shadow-inner bg-white">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead className="bg-slate-100 text-slate-700 sticky top-0 font-extrabold border-b border-slate-200">
                          <tr>
                            <th className="p-2.5 w-12 text-center">No</th>
                            <th className="p-2.5 w-28">Bentuk</th>
                            <th className="p-2.5 w-40">Materi / Kompetensi</th>
                            <th className="p-2.5">Teks Soal / Pertanyaan</th>
                            <th className="p-2.5 w-20 text-center">Kunci</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-slate-800">
                          {importAiParsedQuestions.slice(0, 10).map((pq, pIdx) => (
                            <tr key={pq.id || pIdx} className="hover:bg-slate-50/80">
                              <td className="p-2.5 text-center font-bold text-slate-500">{pq.noSoal}</td>
                              <td className="p-2.5 font-semibold text-indigo-700">{getBentukSoalLabel(pq.bentukSoal)}</td>
                              <td className="p-2.5 font-medium text-slate-600 truncate max-w-[150px]">{pq.kompetensi}</td>
                              <td className="p-2.5 font-medium leading-normal line-clamp-2">{pq.soal}</td>
                              <td className="p-2.5 text-center font-black text-emerald-700">{pq.kunciJawaban}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : importAiRawText.trim().length > 0 ? (
                  <div className="bg-amber-50 border border-amber-200 rounded-2xl p-3 text-xs text-amber-800 font-medium text-center">
                    ⚠️ Belum ada butir soal yang terdeteksi dengan tepat. Pastikan teks atau file Excel Anda memiliki kolom "Soal / Pertanyaan" dan "Opsi/Pilihan".
                  </div>
                ) : null}

              </div>

              {/* Modal Footer */}
              <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
                <span className="text-[11px] text-slate-500 font-medium">
                  💡 Soal yang diimpor akan langsung tersusun rapi di Menu 3. Butir Soal TKA SMA.
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setIsImportAiModalOpen(false)}
                    className="px-4 py-2 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-200 transition"
                  >
                    Batal
                  </button>
                  <button
                    onClick={handleSaveImportedAiQuestions}
                    disabled={isImportingAi || importAiParsedQuestions.length === 0}
                    className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-xl text-xs font-extrabold transition flex items-center gap-2 shadow"
                  >
                    {isImportingAi ? (
                      <>
                        <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                        <span>Menyimpan...</span>
                      </>
                    ) : (
                      <>
                        <Check className="h-3.5 w-3.5" />
                        <span>Simpan & Masukkan {importAiParsedQuestions.length} Soal</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}

        {showSignOutConfirm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-[999] flex items-center justify-center p-4"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-slate-900 border border-slate-800 rounded-3xl max-w-md w-full overflow-hidden shadow-2xl p-6 relative"
            >
              <div className="flex items-center gap-3.5 mb-4 text-amber-500">
                <div className="h-10 w-10 rounded-2xl bg-amber-500/10 flex items-center justify-center">
                  <LogOut className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Konfirmasi Keluar</h3>
                  <p className="text-xs text-slate-400">Yakin ingin meninggalkan sistem?</p>
                </div>
              </div>

              <p className="text-xs text-slate-300 leading-relaxed mb-6">
                Anda akan keluar dari sesi saat ini. Pastikan semua perubahan data parameter atau draf Anda telah tersimpan dengan benar di sistem.
              </p>

              <div className="flex items-center justify-end gap-3">
                <button
                  onClick={() => setShowSignOutConfirm(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition"
                >
                  Batal
                </button>
                <button
                  onClick={executeSignOut}
                  className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold transition"
                >
                  Ya, Keluar
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Custom Confirmation Modal: Delete Single Jadwal */}
        {jadwalToDelete && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-[999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-white border border-slate-200 rounded-3xl max-w-md w-full overflow-hidden shadow-2xl p-6 relative"
            >
              <div className="flex items-center gap-3.5 mb-4 text-red-600">
                <div className="h-10 w-10 rounded-2xl bg-red-50 flex items-center justify-center">
                  <Trash2 className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-950">Hapus Rencana Belajar?</h3>
                  <p className="text-xs text-slate-500">Konfirmasi pembatalan baris jadwal</p>
                </div>
              </div>

              <div className="bg-slate-50 border border-slate-100 rounded-2xl p-4 text-xs space-y-2 mb-6">
                <div className="flex justify-between border-b border-slate-100 pb-2">
                  <span className="text-slate-500 font-medium">Bulan / Periode:</span>
                  <span className="text-slate-900 font-extrabold">{jadwalToDelete.bulan}</span>
                </div>
                <div className="flex justify-between border-b border-slate-100 pb-2">
                  <span className="text-slate-500 font-medium">Minggu Ke:</span>
                  <span className="text-slate-900 font-extrabold">Minggu {jadwalToDelete.mingguKe}</span>
                </div>
                <div className="space-y-1">
                  <span className="text-slate-500 font-medium block">Elemen / Materi:</span>
                  <p className="text-indigo-950 font-bold leading-relaxed">{jadwalToDelete.elemenMateri}</p>
                </div>
              </div>

              <div className="flex items-center justify-end gap-3">
                <button
                  onClick={() => setJadwalToDelete(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
                >
                  Batal
                </button>
                <button
                  onClick={executeDeleteJadwal}
                  className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold transition shadow-sm"
                >
                  Ya, Hapus Rencana
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Custom Confirmation Modal: Clear All Jadwal */}
        {showClearJadwalConfirm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-[999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-white border border-slate-200 rounded-3xl max-w-md w-full overflow-hidden shadow-2xl p-6 relative"
            >
              <div className="flex items-center gap-3.5 mb-4 text-red-600">
                <div className="h-10 w-10 rounded-2xl bg-red-50 flex items-center justify-center">
                  <Trash2 className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-950">Kosongkan Semua Jadwal?</h3>
                  <p className="text-xs text-slate-500">Konfirmasi pembersihan penuh tabel</p>
                </div>
              </div>

              <p className="text-xs text-slate-600 leading-relaxed mb-6">
                Apakah Anda yakin ingin menghapus <strong>seluruh rencana pembelajaran</strong> yang ada pada tabel? 
                Seluruh baris jadwal Anda saat ini akan dibersihkan secara permanen. Anda dapat memulihkannya lagi dengan mengimpor rekomendasi matriks asesmen yang sesuai di panel sebelah kanan.
              </p>

              <div className="flex items-center justify-end gap-3">
                <button
                  onClick={() => setShowClearJadwalConfirm(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
                >
                  Batal
                </button>
                <button
                  onClick={executeClearJadwal}
                  className="px-4 py-2 bg-red-600 hover:bg-red-750 text-white rounded-xl text-xs font-bold transition shadow-sm"
                >
                  Ya, Kosongkan Semua
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Custom Confirmation Modal: Import All Presets */}
        {showImportPresetsConfirm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-[999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-white border border-slate-200 rounded-3xl max-w-md w-full overflow-hidden shadow-2xl p-6 relative"
            >
              <div className="flex items-center gap-3.5 mb-4 text-indigo-600">
                <div className="h-10 w-10 rounded-2xl bg-indigo-50 flex items-center justify-center">
                  <Sparkles className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-950">Impor Semua Rencana?</h3>
                  <p className="text-xs text-slate-500">Mata Pelajaran: {showImportPresetsConfirm.subject}</p>
                </div>
              </div>

              <p className="text-xs text-slate-600 leading-relaxed mb-6">
                Apakah Anda yakin ingin mengimpor sekaligus seluruh <strong>{showImportPresetsConfirm.count} rencana pembelajaran</strong> standar Pusmendik <strong>{showImportPresetsConfirm.subject}</strong> ke tabel jadwal Anda?
                <br /><br />
                Sistem akan secara otomatis mendistribusikannya secara merata ke minggu-minggu pada bulan <strong>Juli, Agustus, September, dan Oktober</strong>.
              </p>

              <div className="flex items-center justify-end gap-3">
                <button
                  onClick={() => setShowImportPresetsConfirm(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
                >
                  Batal
                </button>
                <button
                  onClick={() => executeImportAllJadwalPresets(showImportPresetsConfirm.presets)}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition shadow-sm"
                >
                  Ya, Impor Semua
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Custom Confirmation Modal: Import All Kisi-Kisi Presets */}
        {showImportKisiPresetsConfirm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-[999] flex items-center justify-center p-4 no-print"
          >
            <motion.div
              initial={{ scale: 0.95, y: 15 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 15 }}
              className="bg-white border border-slate-200 rounded-3xl max-w-md w-full overflow-hidden shadow-2xl p-6 relative"
            >
              <div className="flex items-center gap-3.5 mb-4 text-amber-600">
                <div className="h-10 w-10 rounded-2xl bg-amber-50 flex items-center justify-center">
                  <Sparkles className="h-5 w-5 text-amber-600" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-950">Impor Semua Matriks?</h3>
                  <p className="text-xs text-slate-500">Mata Pelajaran: {showImportKisiPresetsConfirm.subject}</p>
                </div>
              </div>

              <p className="text-xs text-slate-600 leading-relaxed mb-6">
                Apakah Anda yakin ingin mengimpor sekaligus seluruh <strong>{showImportKisiPresetsConfirm.count} matriks standar</strong> Pusmendik <strong>{showImportKisiPresetsConfirm.subject}</strong> ke daftar Matriks Asesmen Anda?
                <br /><br />
                Matriks baru akan langsung ditambahkan secara berurutan dan siap digunakan untuk merumuskan butir soal AI.
              </p>

              <div className="flex items-center justify-end gap-3">
                <button
                  onClick={() => setShowImportKisiPresetsConfirm(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition"
                >
                  Batal
                </button>
                <button
                  onClick={() => executeImportAllKisiPresets(
                    showImportKisiPresetsConfirm.presets, 
                    showImportKisiPresetsConfirm.subjectMapped, 
                    showImportKisiPresetsConfirm.subject
                  )}
                  className="px-4 py-2 bg-amber-500 hover:bg-amber-600 text-slate-950 rounded-xl text-xs font-bold transition shadow-sm"
                >
                  ⚡ Ya, Impor Semua Matriks
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {/* Full-Screen Image Zoom Lightbox Modal */}
        {activeZoomImage && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-950/90 backdrop-blur-lg z-[99999] flex flex-col justify-between p-4 sm:p-6 no-print select-none"
          >
            {/* Lightbox Top Header Toolbar */}
            <div className="flex items-center justify-between gap-4 bg-slate-900/80 border border-slate-800 backdrop-blur-md px-4 py-3 rounded-2xl shadow-xl">
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="p-2 bg-indigo-500/20 text-indigo-400 rounded-xl border border-indigo-500/30 flex-shrink-0">
                  <Maximize2 className="h-4 w-4" />
                </div>
                <div className="min-w-0">
                  <h3 className="text-sm font-bold text-white truncate">
                    {activeZoomImage.caption || 'Pratinjau Gambar Soal (Detail Lightbox)'}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-mono">
                    Skala: {Math.round(zoomScale * 100)}% | Rotasi: {zoomRotation}°
                  </p>
                </div>
              </div>

              {/* Control Buttons */}
              <div className="flex items-center gap-1.5 flex-wrap justify-end">
                {/* Zoom Out */}
                <button
                  onClick={() => setZoomScale(prev => Math.max(0.5, prev - 0.25))}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold transition flex items-center gap-1 border border-slate-700 cursor-pointer"
                  title="Perkecil (-25%)"
                >
                  <ZoomOut className="h-4 w-4" />
                  <span className="hidden sm:inline">Perkecil</span>
                </button>

                {/* Zoom In */}
                <button
                  onClick={() => setZoomScale(prev => Math.min(3.5, prev + 0.25))}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold transition flex items-center gap-1 border border-slate-700 cursor-pointer"
                  title="Perbesar (+25%)"
                >
                  <ZoomIn className="h-4 w-4" />
                  <span className="hidden sm:inline">Perbesar</span>
                </button>

                {/* Rotate */}
                <button
                  onClick={() => setZoomRotation(prev => (prev + 90) % 360)}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold transition flex items-center gap-1 border border-slate-700 cursor-pointer"
                  title="Putar Gambar (90°)"
                >
                  <RotateCw className="h-4 w-4" />
                  <span className="hidden sm:inline">Putar</span>
                </button>

                {/* Reset */}
                <button
                  onClick={() => {
                    setZoomScale(1);
                    setZoomRotation(0);
                  }}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition flex items-center gap-1 border border-slate-700 cursor-pointer"
                  title="Reset Ukuran Normal"
                >
                  <Minimize2 className="h-4 w-4" />
                  <span className="hidden sm:inline">Reset</span>
                </button>

                {/* Close Button */}
                <button
                  onClick={() => {
                    setActiveZoomImage(null);
                    setZoomScale(1);
                    setZoomRotation(0);
                  }}
                  className="p-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition ml-2 shadow-lg flex items-center gap-1 cursor-pointer"
                >
                  <span className="font-extrabold text-sm px-1">✕</span>
                  <span className="hidden sm:inline">Tutup</span>
                </button>
              </div>
            </div>

            {/* Center Image Container */}
            <div className="flex-1 flex items-center justify-center overflow-auto my-4 p-2 relative">
              {activeZoomImage.url.trim().toLowerCase().startsWith('<svg') ? (
                <div 
                  className="bg-white p-6 rounded-2xl shadow-2xl overflow-auto max-w-4xl max-h-[75vh]"
                  style={{
                    transform: `scale(${zoomScale}) rotate(${zoomRotation}deg)`,
                    transition: 'transform 0.2s ease-out'
                  }}
                  dangerouslySetInnerHTML={{ __html: activeZoomImage.url }}
                />
              ) : (
                <motion.img
                  src={activeZoomImage.url}
                  alt={activeZoomImage.caption || 'Detail Gambar'}
                  className="max-h-[78vh] max-w-full object-contain rounded-2xl shadow-2xl border border-slate-800 cursor-grab active:cursor-grabbing"
                  style={{
                    transform: `scale(${zoomScale}) rotate(${zoomRotation}deg)`,
                    transition: 'transform 0.2s ease-out'
                  }}
                />
              )}
            </div>

            {/* Bottom Caption Bar */}
            {activeZoomImage.caption && (
              <div className="bg-slate-900/80 border border-slate-800 backdrop-blur-md px-6 py-3 rounded-2xl text-center shadow-xl max-w-3xl mx-auto w-full">
                <p className="text-xs sm:text-sm font-semibold text-slate-200 tracking-tight">
                  {activeZoomImage.caption}
                </p>
              </div>
            )}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function renderMathFormula(latex: string, key?: number | string) {
  let s = latex.trim();
  // Clean \text{...}, \mathrm{...}, \mathbf{...}
  s = s.replace(/\\text\{([^}]+)\}/g, '$1');
  s = s.replace(/\\mathrm\{([^}]+)\}/g, '$1');
  s = s.replace(/\\mathbf\{([^}]+)\}/g, '$1');

  // Fractions: \frac{a}{b} -> (a / b)
  s = s.replace(/\\frac\{([^}]+)\}\{([^}]+)\}/g, '($1 / $2)');
  s = s.replace(/\\sqrt\{([^}]+)\}/g, '√($1)');
  s = s.replace(/\\sqrt\[([^\]]+)\]\{([^}]+)\}/g, '^{$1}√($2)');

  // Greek and math symbols mapping
  const symbolMap: Record<string, string> = {
    '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\Gamma': 'Γ',
    '\\delta': 'δ', '\\Delta': 'Δ', '\\epsilon': 'ε', '\\varepsilon': 'ε',
    '\\theta': 'θ', '\\Theta': 'Θ', '\\lambda': 'λ', '\\Lambda': 'Λ',
    '\\mu': 'μ', '\\pi': 'π', '\\Pi': 'Π', '\\rho': 'ρ',
    '\\sigma': 'σ', '\\Sigma': 'Σ', '\\tau': 'τ', '\\phi': 'φ',
    '\\Phi': 'Φ', '\\omega': 'ω', '\\Omega': 'Ω',
    '\\times': ' × ', '\\cdot': ' · ', '\\div': ' ÷ ',
    '\\pm': '±', '\\mp': '∓', '\\approx': ' ≈ ', '\\neq': ' ≠ ',
    '\\leq': ' ≤ ', '\\geq': ' ≥ ', '\\ll': ' ≪ ', '\\gg': ' ≫ ',
    '\\infty': '∞', '\\rightarrow': ' → ', '\\leftarrow': ' ← ',
    '\\rightleftharpoons': ' ⇌ ', '\\degree': '°', '^{\\circ}': '°',
    '\\int': '∫', '\\sum': '∑', '\\prod': '∏',
    '\\partial': '∂', '\\nabla': '∇', '\\in': ' ∈ '
  };

  for (const [sym, rep] of Object.entries(symbolMap)) {
    s = s.replaceAll(sym, rep);
  }

  // Subscripts & Superscripts parsing
  const parts = s.split(/(\^\{[^}]+\}|\^[0-9a-zA-Z+-]+|_\{[^}]+\}|_[0-9a-zA-Z+-]+)/g);

  return (
    <span key={key} className="inline-flex items-baseline font-serif font-medium tracking-tight text-slate-900 mx-0.5 px-1.5 py-0.5 bg-indigo-50/60 rounded border border-indigo-100/80 shadow-[0_1px_2px_rgba(0,0,0,0.03)] text-[0.95em]">
      {parts.map((p, pIdx) => {
        if (!p) return null;
        if (p.startsWith('^{') && p.endsWith('}')) {
          return <sup key={pIdx} className="text-[0.75em] font-bold text-indigo-950 font-sans">{p.slice(2, -1)}</sup>;
        }
        if (p.startsWith('^')) {
          return <sup key={pIdx} className="text-[0.75em] font-bold text-indigo-950 font-sans">{p.slice(1)}</sup>;
        }
        if (p.startsWith('_{') && p.endsWith('}')) {
          return <sub key={pIdx} className="text-[0.75em] font-semibold text-slate-800 font-sans">{p.slice(2, -1)}</sub>;
        }
        if (p.startsWith('_')) {
          return <sub key={pIdx} className="text-[0.75em] font-semibold text-slate-800 font-sans">{p.slice(1)}</sub>;
        }
        return <span key={pIdx}>{p}</span>;
      })}
    </span>
  );
}

function SimpleMarkdown({ content }: { content: string }) {
  const blocks = content.split(/\n\n+/);

  return (
    <div className="space-y-4 text-slate-700 leading-relaxed text-sm">
      {blocks.map((block, bIdx) => {
        const trimmed = block.trim();
        if (!trimmed) return null;

        // Code block or ASCII diagram
        if (trimmed.startsWith('```') && trimmed.endsWith('```')) {
          const codeLines = trimmed.split('\n');
          const codeContent = codeLines.slice(1, -1).join('\n');
          return (
            <pre key={bIdx} className="bg-slate-900 text-emerald-400 font-mono text-xs p-4 rounded-xl overflow-x-auto shadow-inner border border-slate-800 my-3 leading-relaxed">
              <code>{codeContent}</code>
            </pre>
          );
        }

        // Standalone display equation $$...$$
        if (trimmed.startsWith('$$') && trimmed.endsWith('$$') && trimmed.length >= 4) {
          const math = trimmed.slice(2, -2).trim();
          return (
            <div key={bIdx} className="my-3 p-3 bg-indigo-50/40 border border-indigo-100 rounded-xl overflow-x-auto text-center shadow-sm">
              {renderMathFormula(math)}
            </div>
          );
        }

        // Table
        if (trimmed.includes('|') && trimmed.split('\n').some(l => l.includes('|-') || l.includes('| -') || l.includes('|:'))) {
          const lines = trimmed.split('\n').map(l => l.trim()).filter(Boolean);
          if (lines.length >= 2) {
            const isSeparator = (line: string) => /^\|?\s*:?-+:?\s*(\||\s*$)/.test(line) && line.includes('-');
            const headerLine = lines[0];
            const dataLines = lines.slice(1).filter(l => !isSeparator(l));

            const parseRow = (line: string) => {
              let t = line;
              if (t.startsWith('|')) t = t.slice(1);
              if (t.endsWith('|')) t = t.slice(0, -1);
              return t.split('|').map(c => c.trim());
            };

            const headers = parseRow(headerLine);

            return (
              <div key={bIdx} className="my-5 overflow-x-auto rounded-xl border border-indigo-100 shadow-sm bg-white">
                <table className="w-full text-left border-collapse text-xs sm:text-sm">
                  <thead>
                    <tr className="bg-indigo-900 text-white font-bold divide-x divide-indigo-800">
                      {headers.map((h, i) => (
                        <th key={i} className="px-3.5 py-2.5 border-b border-indigo-800 text-xs font-bold uppercase tracking-wider">
                          {renderInlines(h)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-slate-700">
                    {dataLines.map((rowLine, rIdx) => {
                      const cells = parseRow(rowLine);
                      return (
                        <tr key={rIdx} className={rIdx % 2 === 0 ? 'bg-white hover:bg-indigo-50/40 transition' : 'bg-slate-50/60 hover:bg-indigo-50/40 transition'}>
                          {cells.map((cell, cIdx) => (
                            <td key={cIdx} className="px-3.5 py-2.5 leading-relaxed border-r border-slate-100 last:border-r-0">
                              {renderInlines(cell)}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          }
        }

        // Header 1
        if (trimmed.startsWith('# ')) {
          return <h1 key={bIdx} className="text-xl font-extrabold text-indigo-950 mt-6 mb-2 pb-1 border-b border-indigo-200">{renderInlines(trimmed.slice(2))}</h1>;
        }
        // Header 2
        if (trimmed.startsWith('## ')) {
          return <h2 key={bIdx} className="text-lg font-bold text-indigo-900 mt-5 mb-2 pl-2.5 border-l-4 border-indigo-600">{renderInlines(trimmed.slice(3))}</h2>;
        }
        // Header 3
        if (trimmed.startsWith('### ')) {
          return <h3 key={bIdx} className="text-base font-bold text-slate-800 mt-4 mb-2 italic">{renderInlines(trimmed.slice(4))}</h3>;
        }
        // Bullet points
        if (trimmed.startsWith('* ') || trimmed.startsWith('- ') || trimmed.startsWith('+ ')) {
          const lines = trimmed.split('\n');
          return (
            <ul key={bIdx} className="list-disc pl-5 space-y-1.5 my-2">
              {lines.map((line, lIdx) => {
                const itemText = line.replace(/^[*+\-\s]+/, '');
                return <li key={lIdx}>{renderInlines(itemText)}</li>;
              })}
            </ul>
          );
        }
        // Numbered list
        if (/^\d+\.\s/.test(trimmed)) {
          const lines = trimmed.split('\n');
          return (
            <ol key={bIdx} className="list-decimal pl-5 space-y-1.5 my-2">
              {lines.map((line, lIdx) => {
                const itemText = line.replace(/^\d+\.\s+/, '');
                return <li key={lIdx}>{renderInlines(itemText)}</li>;
              })}
            </ol>
          );
        }
        // Blockquote
        if (trimmed.startsWith('> ')) {
          const text = trimmed.slice(2).replace(/\n>\s/g, '\n');
          return (
            <blockquote key={bIdx} className="border-l-4 border-indigo-500 bg-indigo-50/50 p-3.5 rounded-r-xl italic text-slate-700 my-3">
              {renderInlines(text)}
            </blockquote>
          );
        }

        // Standard paragraph
        return (
          <p key={bIdx} className="whitespace-pre-line text-slate-700 leading-relaxed text-justify">
            {renderInlines(trimmed)}
          </p>
        );
      })}
    </div>
  );
}

function renderInlines(text: string): React.ReactNode {
  if (!text) return null;
  // Regex matches $$...$$, $...$, ***...***, **...**, *...*, `...`
  const regex = /(\$\$[\s\S]+?\$\$|\$[^\$\n]+?\$|\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|`.*?`)/g;
  const parts = text.split(regex);

  return (
    <>
      {parts.map((part, index) => {
        if (!part) return null;
        // Display math $$...$$
        if (part.startsWith('$$') && part.endsWith('$$') && part.length >= 4) {
          const math = part.slice(2, -2);
          return (
            <span key={index} className="block my-2 text-center overflow-x-auto py-1.5 px-3 bg-indigo-50/30 border border-indigo-100 rounded-lg">
              {renderMathFormula(math, index)}
            </span>
          );
        }
        // Inline math $...$
        if (part.startsWith('$') && part.endsWith('$') && part.length >= 2) {
          const math = part.slice(1, -1);
          return renderMathFormula(math, index);
        }
        // Bold Italic
        if (part.startsWith('***') && part.endsWith('***') && part.length >= 6) {
          const inner = part.slice(3, -3);
          return <strong key={index} className="font-extrabold italic text-slate-900">{renderInlines(inner)}</strong>;
        }
        // Bold
        if (part.startsWith('**') && part.endsWith('**') && part.length >= 4) {
          const inner = part.slice(2, -2);
          return <strong key={index} className="font-extrabold text-slate-900">{renderInlines(inner)}</strong>;
        }
        // Italic
        if (part.startsWith('*') && part.endsWith('*') && part.length >= 2) {
          const inner = part.slice(1, -1);
          return <em key={index} className="italic text-slate-800">{renderInlines(inner)}</em>;
        }
        // Inline Code
        if (part.startsWith('`') && part.endsWith('`') && part.length >= 2) {
          const inner = part.slice(1, -1);
          return <code key={index} className="bg-slate-100 text-rose-700 px-1.5 py-0.5 rounded font-mono text-xs border border-slate-200">{inner}</code>;
        }

        // Convert common LaTeX scientific symbols in plain text
        let cleanText = part
          .replace(/\\degree/g, '°')
          .replace(/\\times/g, '×')
          .replace(/\\div/g, '÷')
          .replace(/\\pm/g, '±')
          .replace(/\\approx/g, '≈')
          .replace(/\\neq/g, '≠')
          .replace(/\\leq/g, '≤')
          .replace(/\\geq/g, '≥')
          .replace(/\\rightarrow/g, '→')
          .replace(/\\Delta/g, 'Δ')
          .replace(/\\Omega/g, 'Ω')
          .replace(/\\alpha/g, 'α')
          .replace(/\\beta/g, 'β')
          .replace(/\\gamma/g, 'γ')
          .replace(/\\theta/g, 'θ')
          .replace(/\\lambda/g, 'λ')
          .replace(/\\mu/g, 'μ')
          .replace(/\\pi/g, 'π')
          .replace(/\\rho/g, 'ρ')
          .replace(/\\sigma/g, 'σ')
          .replace(/\\omega/g, 'ω');

        // Check for superscripts/subscripts like ^2, ^3, _1, _2
        if (/(\^\{[^}]+\}|\^[0-9a-zA-Z+-]+|_\{[^}]+\}|_[0-9a-zA-Z+-]+)/.test(cleanText)) {
          const subParts = cleanText.split(/(\^\{[^}]+\}|\^[0-9a-zA-Z+-]+|_\{[^}]+\}|_[0-9a-zA-Z+-]+)/g);
          return (
            <span key={index}>
              {subParts.map((sp, sIdx) => {
                if (!sp) return null;
                if (sp.startsWith('^{') && sp.endsWith('}')) return <sup key={sIdx} className="text-[0.75em] font-bold text-indigo-950 font-sans">{sp.slice(2, -1)}</sup>;
                if (sp.startsWith('^')) return <sup key={sIdx} className="text-[0.75em] font-bold text-indigo-950 font-sans">{sp.slice(1)}</sup>;
                if (sp.startsWith('_{') && sp.endsWith('}')) return <sub key={sIdx} className="text-[0.75em] font-semibold text-slate-800 font-sans">{sp.slice(2, -1)}</sub>;
                if (sp.startsWith('_')) return <sub key={sIdx} className="text-[0.75em] font-semibold text-slate-800 font-sans">{sp.slice(1)}</sub>;
                return <span key={sIdx}>{sp}</span>;
              })}
            </span>
          );
        }

        return <span key={index}>{cleanText}</span>;
      })}
    </>
  );
}
