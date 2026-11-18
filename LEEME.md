# Cobertura de Inventarios · Villavicencio

Son dos enlaces:
- **`index.html`**, la página pública. Muestra la pestaña **Data**, no pide inicio de sesión y es la que compartes con todos.
- **`cargador.html`**, para quien actualiza. Se entra con correo y contraseña (no necesita cuenta de Google). Ahí subes **ZEXI_01**, **TMS** y **Preventa**. El cargador recalcula Data con las mismas fórmulas del Excel y, al publicar, la página pública se actualiza para todos.

- **GitHub Pages** aloja la página (gratis).
- **Firebase (Firestore)** guarda los datos publicados (plan gratuito Spark).
- **Firebase Authentication (correo y contraseña)** controla quién puede publicar.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Página pública (solo ver) |
| `cargador.html` | Página para cargar y publicar los datos |
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
4. **Authentication** → *Comenzar* → *Método de acceso* → habilita **Correo electrónico/contraseña**. Luego, en la pestaña **Usuarios** → *Agregar usuario*, crea el correo y la contraseña con los que vas a entrar al cargador. El correo puede ser cualquiera, por ejemplo el de la empresa: no se verifica ni recibe mensajes.
5. Ve a **Configuración del proyecto** (engranaje) → *Tus apps* → agrega una app **Web** (`</>`). Copia el objeto `firebaseConfig` que te muestra.

### 2. Configurar la página
Abre `firebase-config.js` y haz dos cambios:
- Reemplaza el objeto `firebaseConfig` por el que copiaste.
- En `EDITORES` pon los mismos correos que pusiste en las reglas.

> La `apiKey` de Firebase no es secreta: está hecha para ir en la página. Lo que protege los datos son las reglas de Firestore.

### 3. Subir a GitHub Pages
1. Crea un repositorio en GitHub (por ejemplo `cobertura-villavicencio`) y sube **todos** los archivos de esta carpeta.
2. En el repositorio entra a **Settings** → **Pages** → *Source: Deploy from a branch* → rama `main`, carpeta `/ (root)` → **Save**.
3. En 1 o 2 minutos quedan los dos enlaces:
   - Página pública: `https://TU_USUARIO.github.io/cobertura-villavicencio/`
   - Cargador: `https://TU_USUARIO.github.io/cobertura-villavicencio/cargador.html`
4. Vuelve a Firebase → **Authentication** → **Settings** → **Dominios autorizados** → agrega `TU_USUARIO.github.io`.

> Si el repositorio es público, cualquiera puede ver el código. Los datos solo los puede cambiar quien esté en las reglas.

## Uso diario (carga única para los dos CD)
1. Abre `cargador.html` y entra con tu correo y contraseña.
2. Arrastra los archivos tal como salen, todos juntos:
   - **ZEXI** de SAP (el .XLS, sin limpiar los espacios)
   - **TMS Shipment Plan** (.xlsb completo; se lee la hoja *Data TMS*, solo FGOODS con destino AV19 o AV24)
   - **Preventa / surtido** (.csv): D+1 = el día siguiente que trae el archivo, D+3 = todos los días del archivo. TAT = ZCP1, KA = ZCP6
   - **SL base 1** (solicitado/planeado/entregado por centro) y **SL base 2** (Codigo Centro, SKU, Clase Documento, Volumen Entregado HL)
   - **Una vez al mes: el pronóstico** por semana (Regional, CD, Sku_Demanda, semanas en hl). Se suman las semanas del mes por CD y se calcula el ABC (A = 80 % del FC, B = siguiente 15 %, C = resto). Queda guardado en Firebase y se usa hasta que subas el del mes siguiente.
3. Revisa la lista: cada archivo sale con ✓ y lo que se leyó de cada centro. Pon la capacidad de cada bodega.
4. Pulsa **Publicar para todos**: se publican Villavicencio y Yopal a la vez.

## Uso semanal (formato anterior)
1. Abre **cargador.html** y entra con tu correo y contraseña.
2. Sube **ZEXI_01**, **TMS** y **Preventa**. También puedes subir el libro completo y se toman esas tres pestañas.
3. Revisa la **vista previa** y pulsa **Publicar para todos**.
4. Quien tenga abierta la página pública ve el cambio al instante, y no hace falta tocar GitHub.

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
- **Equipos con Google bloqueado:** el ingreso no usa cuentas de Google, pero Firebase sí carga desde servidores de Google (`gstatic.com`, `googleapis.com`). Si la red del equipo bloquea esos dominios, el cargador no podrá publicar; en ese caso hay que pedir a sistemas que los habiliten.

## Varios centros (Villavicencio y Yopal)
- Los centros se configuran en `centros.js`. Cada uno tiene su nombre, su colección en Firebase, su `master` y su capacidad.
- Enlaces:
  - Villavicencio: `index.html` (o `index.html?centro=AV19`) y `cargador.html?centro=AV19`
  - Yopal: `index.html?centro=AV24` y `cargador.html?centro=AV24`
- Arriba de la página hay un selector para cambiar de centro.
- Los almacenes salen del código del centro: AV19 → 1901 (bodega), 1940 (tránsito), 1941 (calle); AV24 → 2401, 2440, 2441.
- Las reglas de Firestore (`firestore.rules`) deben permitir las colecciones `cobertura` y `cobertura_AV24`.
- El maestro de productos (`master.json`) es el mismo para los dos centros.

## Enlace con una sola vista
Agrega `&vista=` al enlace público para mostrar solo esa pestaña, sin las demás:
- `index.html?centro=AV19&vista=canal` (volumen entregado por canal, Villavicencio)
- Otras vistas: `ocu` (ocupación), `d1`, `d3`, `canal`, `rot` (rotación), `sku` (detalle), `data` (tabla completa)
Es un enlace de presentación: los datos siguen siendo los mismos de la página pública.

## TSO (stock out teórico)
- Stock out del SKU = FCST día − HL en piso, cuando el piso no alcanza. TSO día = stock out ÷ FCST del día. TSO mes = acumulado de los días con foto.
- Cada publicación guarda en Firebase la "foto" del día (`tso_AAAA-MM-DD`: HL en piso y FCST día por SKU). Vale la primera publicación del día (la de la mañana, antes de las 7); las siguientes no la cambian. Si la de la mañana quedó mal, marca "Reemplazar la foto TSO de hoy" al publicar.
- La meta del gráfico está en `TSO_META` (3,5 %) dentro de index.html y cargador.html.
