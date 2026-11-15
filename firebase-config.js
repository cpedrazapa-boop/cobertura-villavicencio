// Configuración de la app web de Firebase (proyecto inventarios-villao)
export const firebaseConfig = {
  apiKey: "AIzaSyCVVs00Q068pD7vm4B_H2puAng351o1jFo",
  authDomain: "inventarios-villao.firebaseapp.com",
  projectId: "inventarios-villao",
  storageBucket: "inventarios-villao.firebasestorage.app",
  messagingSenderId: "297836888259",
  appId: "1:297836888259:web:53683a83a16ad8106f6276"
};

// Correos (usuarios creados en Firebase Authentication) que pueden publicar datos.
// Deben ser los mismos que están en las reglas de Firestore.
export const EDITORES = [
  "cpedrazapa@gmail.com"
];
