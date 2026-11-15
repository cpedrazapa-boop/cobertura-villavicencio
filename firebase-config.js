// 1) Pega aquí la configuración de tu app web de Firebase
//    (Consola de Firebase > Configuración del proyecto > Tus apps > SDK setup and configuration > Config)
export const firebaseConfig = {
  apiKey: "PEGA_AQUI_TU_API_KEY",
  authDomain: "TU_PROYECTO.firebaseapp.com",
  projectId: "TU_PROYECTO",
  storageBucket: "TU_PROYECTO.appspot.com",
  messagingSenderId: "000000000000",
  appId: "1:000000000000:web:xxxxxxxxxxxx"
};

// 2) Correos de Google que pueden cargar y publicar datos.
//    Deben ser los mismos que pongas en las reglas de Firestore (firestore.rules).
export const EDITORES = [
  "tu-correo@gmail.com"
];
