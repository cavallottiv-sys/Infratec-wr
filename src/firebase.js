import { initializeApp } from "firebase/app";
import { getDatabase } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyAR_bBIcQJgaFNPKV7T0yx6AnlB7CCtxII",
  authDomain: "infratec-wrmanager.firebaseapp.com",
  databaseURL: "https://infratec-wrmanager-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "infratec-wrmanager",
  storageBucket: "infratec-wrmanager.firebasestorage.app",
  messagingSenderId: "144091227908",
  appId: "1:144091227908:web:6c81a0c80f9b0f52339251",
  measurementId: "G-K7PXPPSE7Q"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

export { app, db };