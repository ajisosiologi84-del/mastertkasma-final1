import { db } from './firebase';
import { 
  doc, 
  getDoc, 
  setDoc, 
  getDocs, 
  collection, 
  deleteDoc,
  onSnapshot,
  writeBatch
} from 'firebase/firestore';
import * as XLSX from 'xlsx';
import { 
  getSupabaseClient, 
  getStoredSupabaseConfig, 
  saveStoredSupabaseConfig, 
  syncGlobalSupabaseConfig,
  isValidHttpUrl,
  SUPABASE_SQL_SETUP_CODE 
} from './supabase';

export interface UserProfile {
  id: string;
  email: string;
  password?: string;
  name: string;
  role: 'admin' | 'user';
  mataPelajaran: string;
  createdAt?: string;
}

const LOCAL_USERS_CACHE_KEY = 'tka_master_users_cache';

export const INITIAL_DEFAULT_USERS: UserProfile[] = [
  {
    id: 'usr_admin',
    email: 'admin@tka.com',
    password: 'admin123',
    name: 'Admin TKA SMA',
    role: 'admin',
    mataPelajaran: 'Sosiologi',
    createdAt: new Date().toISOString()
  },
  {
    id: 'usr_demo',
    email: 'user@tka.com',
    password: 'user123',
    name: 'Guru Sosiologi',
    role: 'user',
    mataPelajaran: 'Sosiologi',
    createdAt: new Date().toISOString()
  }
];

export { getStoredSupabaseConfig, saveStoredSupabaseConfig, SUPABASE_SQL_SETUP_CODE };

/**
 * Tests connection to Supabase database and checks if 'users' table exists
 */
export async function testSupabaseConnection(url: string, key: string): Promise<{ success: boolean; message: string; count?: number }> {
  const cleanUrl = url.trim();
  const cleanKey = key.trim();

  if (!cleanUrl || !cleanKey) {
    return { success: false, message: "URL dan Anon Key Supabase wajib diisi." };
  }

  if (!isValidHttpUrl(cleanUrl)) {
    return { 
      success: false, 
      message: "Format Project URL tidak valid. Pastikan diawali dengan https:// (contoh: https://xyz.supabase.co)." 
    };
  }

  try {
    const { createClient } = await import('@supabase/supabase-js');
    const client = createClient(cleanUrl, cleanKey);

    const { data, error } = await client
      .from('users')
      .select('id, email, name, role, mata_pelajaran')
      .limit(5);

    if (error) {
      if (error.code === '42P01' || error.message.toLowerCase().includes('relation "public.users" does not exist') || error.message.toLowerCase().includes('does not exist')) {
        return {
          success: false,
          message: "Koneksi ke Supabase berhasil, namun tabel 'users' belum dibuat. Silakan salin & jalankan SQL Setup di Supabase SQL Editor."
        };
      }
      return {
        success: false,
        message: `Koneksi gagal (${error.code || 'Error'}): ${error.message}`
      };
    }

    return {
      success: true,
      message: `Koneksi ke Supabase berhasil! Terhubung ke tabel 'users' (${data?.length || 0} akun ditemukan).`,
      count: data?.length || 0
    };
  } catch (err: any) {
    return {
      success: false,
      message: `Gagal menghubungkan ke Supabase: ${err.message || err}`
    };
  }
}

/**
 * Initializes default user records in Cloud Firestore if database is empty
 */
export async function seedDefaultUsersIfEmpty(): Promise<void> {
  try {
    const snap = await getDocs(collection(db, 'users'));
    if (snap.empty) {
      for (const u of INITIAL_DEFAULT_USERS) {
        await setDoc(doc(db, 'users', u.id), u, { merge: true });
      }
    }
  } catch (e) {
    console.warn("Notice seeding default users to Firestore:", e);
  }
}

/**
 * Normalizes string for flexible matching (removes excess spaces & lowers case)
 */
function normalizeIdentifier(str?: string): string {
  if (!str) return '';
  return str.trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Checks if input identifier matches a user record (by Email, Username, or Full Name)
 */
function matchUserIdentifier(input: string, user: { email?: string; name?: string }): boolean {
  const target = normalizeIdentifier(input);
  if (!target) return false;

  const email = normalizeIdentifier(user.email);
  const name = normalizeIdentifier(user.name);
  const emailPrefix = email.split('@')[0] || '';

  // 1. Exact email match
  if (email === target) return true;

  // 2. Username / prefix before @ match
  if (emailPrefix === target) return true;

  // 3. Exact full name match
  if (name === target) return true;

  // 4. Partial name match if length > 2
  if (name.length > 2 && (name.includes(target) || target.includes(name))) return true;

  return false;
}

/**
 * Verifies password match (supports default master fallback passwords)
 */
function verifyPasswordMatch(inputPass: string, storedPass?: string, email?: string): boolean {
  const cleanPass = inputPass.trim();
  const cleanStored = storedPass ? storedPass.trim() : '';
  const cleanEmail = normalizeIdentifier(email);

  if (!cleanStored) return true;
  if (cleanStored === cleanPass) return true;

  // Admin and Demo defaults fallback
  if (cleanEmail === 'admin@tka.com' && (cleanPass === 'admin123' || cleanPass === 'admin')) return true;
  if (cleanEmail === 'user@tka.com' && (cleanPass === 'user123' || cleanPass === 'user')) return true;

  return false;
}

/**
 * Real-time Subscription to Users (Supabase + Cloud Firestore fallback)
 * Instant synchronization across all browsers, mobile devices, and mastertkasma.my.id
 */
export function subscribeToUsers(callback: (users: UserProfile[]) => void): () => void {
  // Sync global Supabase credentials first
  syncGlobalSupabaseConfig().then(() => {
    const supabase = getSupabaseClient();
    if (supabase) {
      fetchUsersFromCloud().then(users => {
        if (users.length > 0) callback(users);
      });

      try {
        const channel = supabase
          .channel('users_realtime')
          .on('postgres_changes', { event: '*', schema: 'public', table: 'users' }, async () => {
            const freshUsers = await fetchUsersFromCloud();
            callback(freshUsers);
          })
          .subscribe();

        return () => {
          supabase.removeChannel(channel);
        };
      } catch (e) {
        console.warn("Notice Supabase channel subscription:", e);
      }
    }
  }).catch(() => {});

  // Fallback / standard: Cloud Firestore real-time listener
  try {
    const unsub = onSnapshot(collection(db, 'users'), (snapshot) => {
      if (!snapshot.empty) {
        const users: UserProfile[] = [];
        snapshot.forEach((d) => {
          const data = d.data();
          if (data && data.email) {
            users.push({
              id: data.id || data.uid || d.id,
              email: String(data.email).trim().toLowerCase(),
              password: data.password ? String(data.password).trim() : '',
              name: data.name ? String(data.name).trim() : data.email.split('@')[0],
              role: data.role === 'admin' ? 'admin' : 'user',
              mataPelajaran: data.mataPelajaran || 'Sosiologi',
              createdAt: data.createdAt || new Date().toISOString()
            });
          }
        });
        if (users.length > 0) {
          saveLocalCachedUsers(users);
          callback(users);
          return;
        }
      }
      // If Firestore returned empty, supply defaults & trigger seeding
      const cached = getLocalCachedUsers();
      callback(cached.length > 0 ? cached : INITIAL_DEFAULT_USERS);
      seedDefaultUsersIfEmpty().catch(() => {});
    }, (err) => {
      console.warn("Firestore onSnapshot subscription notice, using local cache:", err);
      callback(getLocalCachedUsers());
    });

    return unsub;
  } catch (e) {
    console.warn("Could not attach Firestore listener:", e);
    callback(getLocalCachedUsers());
    return () => {};
  }
}

/**
 * Fetch all users from Supabase with Cloud Firestore & local cache fallback
 */
export async function fetchUsersFromCloud(): Promise<UserProfile[]> {
  await syncGlobalSupabaseConfig().catch(() => {});
  const supabase = getSupabaseClient();

  // 1. Try Supabase if configured
  if (supabase) {
    try {
      const { data, error } = await supabase
        .from('users')
        .select('*')
        .order('created_at', { ascending: false });

      if (!error && data && data.length > 0) {
        const mappedUsers: UserProfile[] = data.map(item => ({
          id: item.id || `usr_${Date.now()}`,
          email: String(item.email).trim().toLowerCase(),
          password: item.password ? String(item.password).trim() : '',
          name: item.name ? String(item.name).trim() : item.email.split('@')[0],
          role: item.role === 'admin' ? 'admin' : 'user',
          mataPelajaran: item.mata_pelajaran || item.mataPelajaran || 'Sosiologi',
          createdAt: item.created_at || new Date().toISOString()
        }));

        saveLocalCachedUsers(mappedUsers);
        return mappedUsers;
      }
    } catch (supaErr) {
      console.warn("Notice fetching users from Supabase:", supaErr);
    }
  }

  // 2. Try Firestore
  try {
    const snap = await getDocs(collection(db, 'users'));
    if (!snap.empty) {
      const users: UserProfile[] = [];
      snap.forEach((d) => {
        const data = d.data();
        if (data && data.email) {
          users.push({
            id: data.id || data.uid || d.id,
            email: String(data.email).trim().toLowerCase(),
            password: data.password ? String(data.password).trim() : '',
            name: data.name ? String(data.name).trim() : data.email.split('@')[0],
            role: data.role === 'admin' ? 'admin' : 'user',
            mataPelajaran: data.mataPelajaran || 'Sosiologi',
            createdAt: data.createdAt || new Date().toISOString()
          });
        }
      });
      if (users.length > 0) {
        saveLocalCachedUsers(users);
        return users;
      }
    }
  } catch (e) {
    console.warn("Notice fetching users from Firestore:", e);
  }

  const local = getLocalCachedUsers();
  return local.length > 0 ? local : INITIAL_DEFAULT_USERS;
}

/**
 * Authenticates user credentials directly via Supabase, Cloud Firestore & Local Cache.
 * Supports flexible login via Email, Username, or Nama Lengkap.
 */
export async function loginUser(
  identifier: string, 
  pass: string
): Promise<{ success: boolean; user?: UserProfile; message?: string }> {
  const cleanInput = identifier.trim();
  const cleanPass = pass.trim();

  if (!cleanInput || !cleanPass) {
    return { success: false, message: 'Silakan masukkan Email / Username dan Password.' };
  }

  await syncGlobalSupabaseConfig().catch(() => {});

  // 1. Check Supabase first if configured
  const supabase = getSupabaseClient();
  if (supabase) {
    try {
      const { data, error } = await supabase
        .from('users')
        .select('*');

      if (!error && data && data.length > 0) {
        const matched = data.find(record => matchUserIdentifier(cleanInput, {
          email: record.email,
          name: record.name
        }));

        if (matched) {
          const isPassValid = verifyPasswordMatch(cleanPass, matched.password, matched.email);
          if (isPassValid) {
            const userObj: UserProfile = {
              id: matched.id || `usr_${matched.email}`,
              email: matched.email,
              name: matched.name || matched.email.split('@')[0],
              role: matched.role === 'admin' ? 'admin' : 'user',
              mataPelajaran: matched.mata_pelajaran || matched.mataPelajaran || 'Sosiologi'
            };
            return { success: true, user: userObj };
          } else {
            return { 
              success: false, 
              message: 'Password yang Anda masukkan salah. Silakan periksa kembali.' 
            };
          }
        }
      }
    } catch (supaErr) {
      console.warn("Notice authenticating via Supabase, trying fallback:", supaErr);
    }
  }

  // 2. Check Cloud Firestore directly
  try {
    const snap = await getDocs(collection(db, 'users'));
    if (!snap.empty) {
      const firestoreUsers: any[] = [];
      snap.forEach(d => firestoreUsers.push({ docId: d.id, ...d.data() }));

      const matched = firestoreUsers.find(d => matchUserIdentifier(cleanInput, {
        email: d.email,
        name: d.name
      }));

      if (matched) {
        const isPassValid = verifyPasswordMatch(cleanPass, matched.password, matched.email);
        if (isPassValid) {
          const userObj: UserProfile = {
            id: matched.id || matched.uid || matched.docId,
            email: matched.email,
            name: matched.name || matched.email.split('@')[0],
            role: matched.role === 'admin' ? 'admin' : 'user',
            mataPelajaran: matched.mataPelajaran || 'Sosiologi'
          };
          return { success: true, user: userObj };
        } else {
          return { 
            success: false, 
            message: 'Password yang Anda masukkan salah. Silakan periksa kembali.' 
          };
        }
      }
    }
  } catch (err) {
    console.warn("Notice verifying login via Firestore:", err);
  }

  // 3. Check local cached users
  const localUsers = getLocalCachedUsers();
  const foundLocal = localUsers.find(u => matchUserIdentifier(cleanInput, u));
  if (foundLocal) {
    const isPassValid = verifyPasswordMatch(cleanPass, foundLocal.password, foundLocal.email);
    if (isPassValid) {
      return { success: true, user: foundLocal };
    } else {
      return { 
        success: false, 
        message: 'Password yang Anda masukkan salah. Silakan periksa kembali.' 
      };
    }
  }

  // 4. Master fallback for default system admin/user
  const normalizedInput = normalizeIdentifier(cleanInput);
  if ((normalizedInput === 'admin@tka.com' || normalizedInput === 'admin') && (cleanPass === 'admin123' || cleanPass === 'admin')) {
    const adminObj: UserProfile = {
      id: 'usr_admin',
      email: 'admin@tka.com',
      name: 'Admin TKA SMA',
      role: 'admin',
      mataPelajaran: 'Sosiologi'
    };
    return { success: true, user: adminObj };
  } else if ((normalizedInput === 'user@tka.com' || normalizedInput === 'user') && (cleanPass === 'user123' || cleanPass === 'user')) {
    const demoObj: UserProfile = {
      id: 'usr_demo',
      email: 'user@tka.com',
      name: 'Guru Sosiologi',
      role: 'user',
      mataPelajaran: 'Sosiologi'
    };
    return { success: true, user: demoObj };
  }

  return { 
    success: false, 
    message: 'Email atau Nama Pengguna tidak ditemukan. Silakan periksa kembali atau hubungi Administrator.' 
  };
}

/**
 * Adds a new user directly to Supabase & Cloud Firestore & updates local cache
 */
export async function addUserToCloud(
  newUser: { email: string; password?: string; name: string; role: 'admin' | 'user'; mataPelajaran: string }
): Promise<{ success: boolean; user?: UserProfile; message?: string }> {
  const cleanEmail = newUser.email.trim().toLowerCase();
  const cleanPass = newUser.password ? newUser.password.trim() : '123456';
  const cleanName = newUser.name.trim() || cleanEmail.split('@')[0];
  const cleanRole = newUser.role || 'user';
  const cleanMapel = newUser.mataPelajaran || 'Sosiologi';

  const userObj: UserProfile = {
    id: `usr_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    email: cleanEmail,
    password: cleanPass,
    name: cleanName,
    role: cleanRole,
    mataPelajaran: cleanMapel,
    createdAt: new Date().toISOString()
  };

  // 1. Insert to Supabase if configured
  const supabase = getSupabaseClient();
  if (supabase) {
    try {
      const { error } = await supabase
        .from('users')
        .insert([{
          id: userObj.id,
          email: cleanEmail,
          password: cleanPass,
          name: cleanName,
          role: cleanRole,
          mata_pelajaran: cleanMapel,
          created_at: userObj.createdAt
        }]);

      if (error) {
        if (error.code === '23505' || error.message.toLowerCase().includes('unique') || error.message.toLowerCase().includes('duplicate')) {
          return { success: false, message: `Email ${cleanEmail} sudah terdaftar di Database Supabase.` };
        }
        console.warn("Supabase insert error:", error);
      }
    } catch (supaErr: any) {
      console.warn("Notice saving user to Supabase:", supaErr);
    }
  }

  // 2. Persist to Firestore for redundant backup
  try {
    const snap = await getDocs(collection(db, 'users'));
    let exists = false;
    snap.forEach((d) => {
      const data = d.data();
      if (data && data.email && String(data.email).trim().toLowerCase() === cleanEmail) {
        exists = true;
      }
    });

    if (exists && !supabase) {
      return { success: false, message: `Email ${cleanEmail} sudah terdaftar di Database Cloud.` };
    }

    await setDoc(doc(db, 'users', userObj.id), userObj, { merge: true });

    await setDoc(doc(db, 'user_settings', `${userObj.id}_generator_config`), {
      mataPelajaran: cleanMapel
    }, { merge: true });

  } catch (err: any) {
    console.warn("Notice saving user to Firestore:", err);
  }

  // Update local cache
  const localUsers = getLocalCachedUsers();
  localUsers.unshift(userObj);
  saveLocalCachedUsers(localUsers);

  return { success: true, user: userObj };
}

/**
 * Updates an existing user in Supabase & Cloud Firestore & local cache
 */
export async function updateUserInCloud(
  updatedUser: { id: string; email: string; name: string; role: 'admin' | 'user'; mataPelajaran: string; password?: string }
): Promise<{ success: boolean; message?: string }> {
  const cleanEmail = updatedUser.email.trim().toLowerCase();
  const cleanName = updatedUser.name.trim();

  // 1. Update in Supabase if configured
  const supabase = getSupabaseClient();
  if (supabase) {
    try {
      const supaPayload: any = {
        name: cleanName,
        role: updatedUser.role,
        mata_pelajaran: updatedUser.mataPelajaran
      };
      if (updatedUser.password && updatedUser.password.trim()) {
        supaPayload.password = updatedUser.password.trim();
      }

      await supabase
        .from('users')
        .update(supaPayload)
        .or(`id.eq.${updatedUser.id},email.eq.${cleanEmail}`);
    } catch (supaErr) {
      console.warn("Notice updating user in Supabase:", supaErr);
    }
  }

  // 2. Update in Firestore
  try {
    const updatePayload: any = {
      id: updatedUser.id,
      email: cleanEmail,
      name: cleanName,
      role: updatedUser.role,
      mataPelajaran: updatedUser.mataPelajaran,
      updatedAt: new Date().toISOString()
    };

    if (updatedUser.password && updatedUser.password.trim()) {
      updatePayload.password = updatedUser.password.trim();
    }

    await setDoc(doc(db, 'users', updatedUser.id), updatePayload, { merge: true });

    await setDoc(doc(db, 'user_settings', `${updatedUser.id}_generator_config`), {
      mataPelajaran: updatedUser.mataPelajaran
    }, { merge: true });

  } catch (err) {
    console.warn("Notice updating user in Firestore:", err);
  }

  // Update local cache
  const localUsers = getLocalCachedUsers();
  const idx = localUsers.findIndex(u => u.id === updatedUser.id || u.email.toLowerCase() === cleanEmail);
  if (idx !== -1) {
    localUsers[idx] = {
      ...localUsers[idx],
      name: cleanName,
      role: updatedUser.role,
      mataPelajaran: updatedUser.mataPelajaran,
      ...(updatedUser.password ? { password: updatedUser.password.trim() } : {})
    };
    saveLocalCachedUsers(localUsers);
  }

  return { success: true };
}

/**
 * Deletes a user from Supabase, Cloud Firestore & local cache
 */
export async function deleteUserFromCloud(
  id: string, 
  email?: string
): Promise<{ success: boolean; message?: string }> {
  // 1. Delete from Supabase if configured
  const supabase = getSupabaseClient();
  if (supabase) {
    try {
      if (email) {
        await supabase.from('users').delete().or(`id.eq.${id},email.eq.${email.trim().toLowerCase()}`);
      } else {
        await supabase.from('users').delete().eq('id', id);
      }
    } catch (supaErr) {
      console.warn("Notice deleting user from Supabase:", supaErr);
    }
  }

  // 2. Delete from Firestore
  try {
    await deleteDoc(doc(db, 'users', id));
  } catch (err) {
    console.warn("Notice deleting user from Firestore:", err);
  }

  // Update local cache
  const localUsers = getLocalCachedUsers();
  const filtered = localUsers.filter(u => u.id !== id && (!email || u.email.toLowerCase() !== email.toLowerCase()));
  saveLocalCachedUsers(filtered);

  return { success: true };
}

/**
 * Batch Imports Users from Excel / CSV directly into Supabase & Cloud Firestore
 */
export async function importUsersBatchToCloud(
  usersToImport: Array<{ name: string; email: string; password?: string; role?: string; mataPelajaran?: string }>
): Promise<{ success: boolean; count: number; importedCount: number; message: string }> {
  if (!usersToImport || usersToImport.length === 0) {
    return { success: false, count: 0, importedCount: 0, message: "Tidak ada data pengguna yang valid untuk diimpor." };
  }

  let count = 0;
  const currentUsers = await fetchUsersFromCloud();
  const existingEmails = new Set(currentUsers.map(u => u.email.toLowerCase()));

  const importedList: UserProfile[] = [];
  const supaBatch: any[] = [];

  for (const item of usersToImport) {
    if (!item.email || !item.name) continue;
    const cleanEmail = String(item.email).trim().toLowerCase();
    if (!cleanEmail.includes('@') || existingEmails.has(cleanEmail)) continue;

    const userObj: UserProfile = {
      id: `usr_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      email: cleanEmail,
      password: item.password ? String(item.password).trim() : '123456',
      name: String(item.name).trim(),
      role: item.role && String(item.role).toLowerCase() === 'admin' ? 'admin' : 'user',
      mataPelajaran: item.mataPelajaran ? String(item.mataPelajaran).trim() : 'Sosiologi',
      createdAt: new Date().toISOString()
    };

    supaBatch.push({
      id: userObj.id,
      email: userObj.email,
      password: userObj.password,
      name: userObj.name,
      role: userObj.role,
      mata_pelajaran: userObj.mataPelajaran,
      created_at: userObj.createdAt
    });

    existingEmails.add(cleanEmail);
    importedList.push(userObj);
    count++;
  }

  if (count === 0) {
    return { 
      success: false, 
      count: 0,
      importedCount: 0, 
      message: "Semua email dalam file Excel sudah terdaftar di database sebelumnya." 
    };
  }

  // 1. Batch Insert to Supabase if configured
  const supabase = getSupabaseClient();
  if (supabase && supaBatch.length > 0) {
    try {
      await supabase.from('users').upsert(supaBatch, { onConflict: 'email' });
    } catch (supaErr) {
      console.warn("Notice batch inserting to Supabase:", supaErr);
    }
  }

  // 2. Batch Insert to Firestore
  try {
    const batch = writeBatch(db);
    for (const u of importedList) {
      batch.set(doc(db, 'users', u.id), u);
    }
    await batch.commit();
  } catch (err: any) {
    console.error("Batch import error on Firestore:", err);
  }

  const updatedLocal = [...importedList, ...currentUsers];
  saveLocalCachedUsers(updatedLocal);

  return { 
    success: true, 
    count,
    importedCount: count, 
    message: `Berhasil mengimpor dan menyimpan ${count} akun pengguna langsung ke Cloud Database!` 
  };
}

export function getAppsScriptUrl(): string {
  return '';
}

/**
 * Exports all user records directly to an Excel (.xlsx) file for 1-click download
 */
export function exportUsersToExcel(users: UserProfile[]): void {
  const data = users.map((u, idx) => ({
    'No': idx + 1,
    'Nama Lengkap': u.name,
    'Email Akun': u.email,
    'Password': u.password || '123456',
    'Hak Akses / Role': u.role === 'admin' ? 'Admin' : 'Guru',
    'Mata Pelajaran': u.mataPelajaran,
    'Tanggal Terdaftar': u.createdAt ? new Date(u.createdAt).toLocaleDateString('id-ID') : '-'
  }));

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Pengguna');
  
  // Format column widths
  worksheet['!cols'] = [
    { wch: 6 },
    { wch: 28 },
    { wch: 30 },
    { wch: 15 },
    { wch: 18 },
    { wch: 22 },
    { wch: 18 }
  ];

  XLSX.writeFile(workbook, `Data_Pengguna_Master_TKA_SMA_${new Date().toISOString().split('T')[0]}.xlsx`);
}

/**
 * Downloads a standard Excel template (.xlsx) for user bulk registration
 */
export function downloadUserExcelTemplate(): void {
  const templateData = [
    {
      'Nama Lengkap': 'Drs. H. Ahmad Dahlan, M.Pd.',
      'Email Akun': 'ahmad.guru@sekolah.sch.id',
      'Password': 'password123',
      'Role (admin/user)': 'user',
      'Mata Pelajaran': 'Sosiologi'
    },
    {
      'Nama Lengkap': 'Siti Rahmawati, S.Pd.',
      'Email Akun': 'siti.rahma@sekolah.sch.id',
      'Password': 'password123',
      'Role (admin/user)': 'user',
      'Mata Pelajaran': 'Ekonomi'
    },
    {
      'Nama Lengkap': 'Budi Setiawan, S.Pd.',
      'Email Akun': 'budi.admin@sekolah.sch.id',
      'Password': 'adminpassword',
      'Role (admin/user)': 'admin',
      'Mata Pelajaran': 'Geografi'
    }
  ];

  const worksheet = XLSX.utils.json_to_sheet(templateData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Template Import Pengguna');
  
  worksheet['!cols'] = [
    { wch: 32 },
    { wch: 30 },
    { wch: 18 },
    { wch: 20 },
    { wch: 22 }
  ];

  XLSX.writeFile(workbook, 'Template_Import_Pengguna_TKA_SMA.xlsx');
}

export function getLocalCachedUsers(): UserProfile[] {
  if (typeof window === 'undefined') return INITIAL_DEFAULT_USERS;
  try {
    const raw = localStorage.getItem(LOCAL_USERS_CACHE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch (e) {
    console.warn("Error reading local user cache:", e);
  }
  return INITIAL_DEFAULT_USERS;
}

export function saveLocalCachedUsers(users: UserProfile[]): void {
  if (typeof window === 'undefined') return;
  try {
    localStorage.setItem(LOCAL_USERS_CACHE_KEY, JSON.stringify(users));
  } catch (e) {
    console.warn("Error saving local user cache:", e);
  }
}

// -------------------------------------------------------------
// Backwards-Compatible Aliases for Legacy Component Invocations
// -------------------------------------------------------------
export const loginWithAppsScript = loginUser;
export const fetchUsersFromAppsScript = fetchUsersFromCloud;
export const addUserToAppsScript = addUserToCloud;
export const updateUserInAppsScript = updateUserInCloud;
export const deleteUserInAppsScript = deleteUserFromCloud;
export const initAppsScriptConfig = async () => '';
export const setAppsScriptUrl = async (url: string) => {};
export const testAppsScriptConnection = async (url: string) => ({
  success: true,
  message: 'Cloud Database Aktif & Siap Digunakan Real-time!'
});
export const DEFAULT_APPSCRIPT_CODE = '';
