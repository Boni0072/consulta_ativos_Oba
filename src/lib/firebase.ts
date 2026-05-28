import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

// Configuração do Firebase extraída do seu console
const firebaseConfig = {
  apiKey: "AIzaSyDY0Zly50vPeXr6WPF8QLRzeqW17DMwaOU",
  authDomain: "consulta-oba-ativos.firebaseapp.com",
  projectId: "consulta-oba-ativos",
  storageBucket: "consulta-oba-ativos.firebasestorage.app",
  messagingSenderId: "762941577322",
  appId: "1:762941577322:web:d7f656f1f02520dcec9213"
};

// Inicializa o Firebase e exporta a instância do Firestore
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);