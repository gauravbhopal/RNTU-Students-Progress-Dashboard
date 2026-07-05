// Firebase Configuration — EduEval Pro Multi-Tenant Portal
// Project: English Bol (english-bol-eb) — used as host for EduEval Pro

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, doc, getDoc, setDoc, collection, getDocs, addDoc, updateDoc, deleteDoc, onSnapshot, query, serverTimestamp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyACzFHyHhDnLWhGFcvKNVrur8jkxOmCJZk",
    authDomain: "english-bol-eb.firebaseapp.com",
    projectId: "english-bol-eb",
    storageBucket: "english-bol-eb.firebasestorage.app",
    messagingSenderId: "871555354966",
    appId: "1:871555354966:web:a065c8b50a9bf45c4806f7"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

export { app, auth, db, onAuthStateChanged, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword, doc, getDoc, setDoc, collection, getDocs, addDoc, updateDoc, deleteDoc, onSnapshot, query, serverTimestamp };
