import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDCqZRb0jHxwhLHVo6EBPAtWNeFGZdb4xQ",
  authDomain: "matriz-sena.firebaseapp.com",
  projectId: "matriz-sena",
  storageBucket: "matriz-sena.firebasestorage.app",
  messagingSenderId: "752214353831",
  appId: "1:752214353831:web:ad81daabfabfae8d317882",
};

const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(app, {
  experimentalAutoDetectLongPolling: true,
});