import { initializeApp, deleteApp } from "firebase/app";
import { getAuth, createUserWithEmailAndPassword, signOut } from "firebase/auth";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD9o5QpiKyfpRBt2BC-quCG07C4SNAcFNA",
  authDomain: "gen-lang-client-0692785103.firebaseapp.com",
  projectId: "gen-lang-client-0692785103",
  storageBucket: "gen-lang-client-0692785103.firebasestorage.app",
  messagingSenderId: "615148222754",
  appId: "1:615148222754:web:a8a8c4b919aa6eb38acd7b"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
}, "ai-studio-copyofremixgener-b07a4ffe-e968-483e-a44f-cb6e8eb2dc8c");

export async function createNewUserByAdmin(email: string, pass: string, name: string, role: 'admin' | 'user', mataPelajaran?: string) {
  // Use a secondary app instance to register the new user without signing out the current admin session
  const secondaryAppName = `SecondaryApp_${Date.now()}`;
  const secondaryApp = initializeApp(firebaseConfig, secondaryAppName);
  const secondaryAuth = getAuth(secondaryApp);
  try {
    const userCred = await createUserWithEmailAndPassword(secondaryAuth, email, pass);
    const uid = userCred.user.uid;
    
    // Save details to the primary db (Firestore)
    await setDoc(doc(db, 'users', uid), {
      uid,
      email: email,
      name: name,
      role: role,
      mataPelajaran: mataPelajaran || 'Sosiologi',
      createdAt: new Date()
    });

    // Initialize user settings with assigned subject
    await setDoc(doc(db, 'user_settings', `${uid}_generator_config`), {
      mataPelajaran: mataPelajaran || 'Sosiologi'
    }, { merge: true });
    
    // Sign out from the temporary secondary session
    await signOut(secondaryAuth);
    // Delete the temporary app
    await deleteApp(secondaryApp);
    
    return { success: true, uid };
  } catch (error: any) {
    // Delete the secondary app to prevent memory leaks or duplicate app errors
    try {
      await deleteApp(secondaryApp);
    } catch (e) {}
    
    if (error.code === 'auth/email-already-in-use' || error.message?.includes('email-already-in-use')) {
      const customErr: any = new Error(`Email '${email}' sudah terdaftar dalam sistem.`);
      customErr.code = 'auth/email-already-in-use';
      throw customErr;
    }
    throw error;
  }
}

export { app, auth, db };
