# Cobertura de Inventarios · Villavicencio

Página web que muestra la pestaña **Data** del Excel de cobertura. Tiene un cargador donde subes **ZEXI_01**, **TMS** y **Preventa**. La página recalcula Data con las mismas fórmulas del Excel y, al publicar, todos ven la nueva versión.

- **GitHub Pages** aloja la página (gratis).
- **Firebase (Firestore)** guarda los datos publicados (plan gratuito Spark).
- **Firebase Authentication (Google)** controla quién puede publicar.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La página |
| `engine.js` | Las fórmulas de la pestaña Data |
| `master.json` | Maestros fijos (BD PT, ABC, Cubo Fc, SL, MCSK, BS Producción, Stock Disp) |
| `data.json` | Datos de respaldo, se muestran mientras no se haya publicado nada |
| `firebase-config.js` | **Aquí pegas tu configuración de Firebase y los correos autorizados** |
| `firestore.rules` | Reglas de seguridad para pegar en Firebase |
| `herramientas/actualizar_maestros.py` | Regenera `master.json` cuando cambien los maestros |

## Configuración (una sola vez, unos 15 minutos)

### 1. Crear el proyecto en Firebase
1. Entra a https://console.firebase.google.com y crea un proyecto. Analytics no es necesario.
2. **Firestore Database** → *Crear base de datos* → modo producción → región `southamerica-east1` (o la que prefieras).
3. En Firestore → pestaña **Reglas**: pega el contenido de `firestore.rules`. Cambia `tu-correo@gmail.com` por los correos que van a cargar datos y publica las reglas.
4. **Authentication** → *Comenzar* → *Método de acceso* → habilita **Google**.
5. Ve a **Configuración del proyecto** (engranaje) → *Tus apps* → agrega una app **Web** (`</>`). Copia el objeto `firebaseConfig` que te muestra.

### 2. Configurar la página
Abre `firebase-config.js` y haz dos cambios:
- Reemplaza el objeto `firebaseConfig` por el que copiaste.
- En `EDITORES` pon los mismos correos que pusiste en las reglas.

> La `apiKey` de Firebase no es secreta: está hecha para ir en la página. Lo que protege los datos son las reglas de Firestore.

### 3. Subir a GitHub Pages
1. Crea un repositorio en GitHub (por ejemplo `cobertura-villavicencio`) y sube **todos** los archivos de esta carpeta.
2. En el repositorio entra a **Settings** → **Pages** → *Source: Deploy from a branch* → rama `main`, carpeta `/ (root)` → **Save**.
3. En 1 o 2 minutos la página queda en `https://TU_USUARIO.github.io/cobertura-villavicencio/`.
4. Vuelve a Firebase → **Authentication** → **Settings** → **Dominios autorizados** → agrega `TU_USUARIO.github.io`.

> Si el repositorio es público, cualquiera puede ver el código. Los datos solo los puede cambiar quien esté en las reglas.

## Uso semanal
1. Abre la página y pulsa **Actualizar datos**. La primera vez te pide entrar con Google.
2. Sube **ZEXI_01**, **TMS** y **Preventa**. También puedes subir el libro completo y se toman esas tres pestañas.
3. Revisa la **vista previa** y pulsa **Publicar para todos**.
4. Quien tenga la página abierta ve el cambio al instante, y no hace falta tocar GitHub.

## Cuándo sí tocar GitHub
Solo si cambian los maestros (BD PT, Cubo Fc, ABC, SL, MCSK, BS Producción, Stock Disp). En ese caso:
```
python herramientas/actualizar_maestros.py "Cobertura Inventarios Villavicencio.xlsx" master.json
```
Después sube el nuevo `master.json` al repositorio.

## Notas
- **Stock de seguridad (HL y DC):** queda en 0 porque en el Excel la fórmula da `#REF!`.
- **TMS:** si no subes el archivo, TMS queda en 0.
- **Archivo TMS:** se lee con las columnas `LLAVE`, `HL` y `Fecha Arribo`, como en *Libro1*.
