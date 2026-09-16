import React, { useState, useEffect } from 'react';
import { motion } from 'motion/react';
import { loginUser } from '../lib/userManagement';
import { syncGlobalSupabaseConfig } from '../lib/supabase';
import { 
  Lock, 
  User, 
  Sparkles, 
  AlertCircle, 
  ArrowRight,
  BookOpen,
  ShieldCheck,
  Eye,
  EyeOff
} from 'lucide-react';

interface LoginScreenProps {
  onLoginSuccess: (role: 'admin' | 'user', name: string, userObj?: any) => void;
}

export default function LoginScreen({ onLoginSuccess }: LoginScreenProps) {
  const [identifier, setIdentifier] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  // Sync Supabase config from Cloud on load so all devices & browsers have connection
  useEffect(() => {
    syncGlobalSupabaseConfig().catch(() => {});
  }, []);

  const handleSignIn = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!identifier.trim() || !password.trim()) {
      setError("Silakan masukkan Email / Username dan Password.");
      return;
    }
    setLoading(true);
    setError(null);

    try {
      const res = await loginUser(identifier.trim(), password.trim());
      if (res.success && res.user) {
        onLoginSuccess(res.user.role, res.user.name, res.user);
      } else {
        setError(res.message || "Email/Username atau Password yang Anda masukkan salah.");
      }
    } catch (err: any) {
      console.warn("Login attempt error:", err);
      setError(`Gagal Masuk: ${err.message || err}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col justify-center items-center px-4 relative overflow-hidden">
      {/* Absolute Decorative Circles */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-blue-600/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-1/4 right-1/4 w-96 h-96 bg-indigo-600/10 rounded-full blur-3xl pointer-events-none" />

      {/* Main Container */}
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
        className="w-full max-w-md bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl p-6 relative z-10"
      >
        {/* TKA SMA Official-Style Header & Emblem Banner */}
        <div className="mb-5 space-y-3 bg-slate-950/80 border border-slate-800 p-4 rounded-2xl text-center relative overflow-hidden shadow-inner">
          {/* Subtle Ambient Glow Effects */}
          <div className="absolute -top-12 -left-12 w-32 h-32 bg-amber-500/10 rounded-full blur-2xl pointer-events-none" />
          <div className="absolute -bottom-12 -right-12 w-32 h-32 bg-sky-500/10 rounded-full blur-2xl pointer-events-none" />

          {/* TKA Visual Banner Logo */}
          <div className="pt-1 space-y-1.5">
            <div className="inline-flex items-center justify-center gap-2">
              {/* Custom Styled 3D Letter TKA */}
              <div className="flex items-center gap-1">
                <span className="text-2xl font-black tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-amber-400 via-yellow-300 to-amber-500 drop-shadow-[0_2px_8px_rgba(245,158,11,0.3)]">
                  T
                </span>
                {/* Book integrated in K */}
                <div className="relative inline-flex items-center justify-center">
                  <span className="text-2xl font-black tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-amber-400 via-yellow-300 to-amber-500">
                    K
                  </span>
                  <BookOpen className="h-3.5 w-3.5 text-amber-300 absolute -top-0.5 -right-1 transform rotate-12 drop-shadow-sm" />
                </div>
                <span className="text-2xl font-black tracking-tight text-transparent bg-clip-text bg-gradient-to-r from-amber-400 via-yellow-300 to-amber-500 drop-shadow-[0_2px_8px_rgba(245,158,11,0.3)]">
                  A
                </span>
              </div>
              <span className="text-xs font-black bg-amber-400/10 text-amber-300 border border-amber-400/30 px-2 py-0.5 rounded-md uppercase tracking-wider">
                SMA
              </span>
            </div>

            <p className="text-[10px] font-bold text-slate-300 tracking-wide uppercase">
              Tes Kemampuan Akademik SMA
            </p>

            {/* Slogan Banner Pill */}
            <div className="inline-flex items-center gap-1.5 bg-sky-950/60 border border-sky-500/30 px-3 py-1 rounded-full text-[10px] font-extrabold text-sky-300 shadow-sm mt-1">
              <span className="text-amber-400">#JUJUR</span>
              <span className="text-sky-500">•</span>
              <span className="text-sky-300">#PRESTATIF</span>
              <span className="text-sky-500">•</span>
              <span className="text-emerald-400">#GEMBIRA</span>
            </div>
          </div>
        </div>

        {/* App Title Sub-Header */}
        <div className="text-center mb-5">
          <h2 className="text-base font-extrabold text-white tracking-tight">Master TKA SMA</h2>
          <p className="text-xs text-slate-400 mt-0.5">Sistem Manajemen Kisi-Kisi & Soal TKA SMA berbasis AI</p>
        </div>

        {/* Alert Messages */}
        {error && (
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="mb-4 bg-rose-500/10 border border-rose-500/20 text-rose-300 p-3 rounded-xl text-xs flex items-start gap-2"
          >
            <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
            <span>{error}</span>
          </motion.div>
        )}

        {successMsg && (
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="mb-4 bg-emerald-500/10 border border-emerald-500/20 text-emerald-300 p-3 rounded-xl text-xs flex items-center gap-2"
          >
            <Sparkles className="h-4 w-4 shrink-0" />
            <span>{successMsg}</span>
          </motion.div>
        )}

        {/* Form Inputs (Sign In only) */}
        <form onSubmit={handleSignIn} className="space-y-4">
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">
              Email / Nama Pengguna / Username
            </label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center text-slate-500">
                <User className="h-4 w-4" />
              </span>
              <input
                type="text"
                placeholder="nama@sekolah.sch.id atau Nama Guru"
                value={identifier}
                onChange={(e) => setIdentifier(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-10 pr-4 py-2.5 text-xs text-white placeholder-slate-600 focus:outline-none focus:border-indigo-500 transition"
                disabled={loading}
              />
            </div>
            <p className="text-[10px] text-slate-500 mt-1">
              *Bisa masuk menggunakan <b>Email</b> atau <b>Nama Lengkap</b> yang terdaftar.
            </p>
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Password</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center text-slate-500">
                <Lock className="h-4 w-4" />
              </span>
              <input
                type={showPassword ? "text" : "password"}
                placeholder="******"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-10 pr-10 py-2.5 text-xs text-white placeholder-slate-600 focus:outline-none focus:border-indigo-500 transition"
                disabled={loading}
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-500 hover:text-slate-300 transition cursor-pointer"
                title={showPassword ? "Sembunyikan Password" : "Tampilkan Password"}
              >
                {showPassword ? <EyeOff className="h-4 w-4 text-indigo-400" /> : <Eye className="h-4 w-4" />}
              </button>
            </div>
          </div>

          {/* Action Button */}
          <button
            type="submit"
            className="w-full bg-gradient-to-r from-indigo-600 to-blue-600 hover:from-indigo-700 hover:to-blue-700 text-white font-extrabold py-3 px-4 rounded-xl shadow-lg transition-all duration-200 flex items-center justify-center gap-2 text-xs mt-2 disabled:opacity-50 cursor-pointer"
            disabled={loading}
          >
            {loading ? "Memproses..." : "Masuk ke Aplikasi"}
            {!loading && <ArrowRight className="h-4 w-4" />}
          </button>

          <div className="text-center pt-2 text-[11px] text-slate-500">
            Penambahan akun pengguna baru hanya dapat dilakukan oleh <span className="text-indigo-400 font-semibold">Administrator</span>.
          </div>
        </form>

        {/* Pusmendik & Perkaban Regulation Compliance Banner */}
        <div className="mt-5 p-3.5 bg-slate-950/90 border border-slate-800 rounded-2xl text-[10px] text-slate-300 text-center leading-relaxed shadow-sm">
          <div className="flex items-center justify-center gap-1.5 font-bold text-sky-400 mb-1">
            <ShieldCheck className="h-3.5 w-3.5 text-sky-400" />
            <span className="uppercase tracking-wider">Standar & Kerangka Resmi</span>
          </div>
          <p className="text-slate-400 font-medium leading-normal">
            Sesuai Kerangka Asesmen Pusmendik TKA (Pusat Asesmen & Standar Pendidikan SMA) dan Perkaban Nomor 45 Tahun 2025 tentang Kerangka Asesmen TKA SMA-MA dan SMK-MAK.
          </p>
        </div>

        {/* Developer Credit Link */}
        <div className="mt-4 text-center">
          <a
            href="https://lynk.id/ajisosiologi"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 px-3 py-2 bg-indigo-950/40 hover:bg-indigo-950 text-[10px] font-extrabold text-indigo-400 rounded-xl border border-indigo-900/60 hover:border-indigo-500/55 transition duration-200"
          >
            <span>Create @ajisosiologi</span>
          </a>
        </div>
      </motion.div>
    </div>
  );
}
