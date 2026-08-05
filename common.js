/* ============================================================
   COMMON.JS — Padrón Nominal Ilo, sistema de tránsito/bajas-altas
   ============================================================
   Compartido por todas las secciones (transito, bajas-altas, y las
   que vengan). Extraído tal cual de index.html, sin cambiar nada de
   su lógica ya probada — solo se movió aquí para no duplicarla.

   Cada página que lo use debe cargarlo ANTES de su propio <script>:
     <script src="common.js"></script>
   ============================================================ */

const msalConfig = {
  auth: {
    clientId: "962a40fd-acc8-450f-b5c8-7a294610d1fb",
    authority: "https://login.microsoftonline.com/pba01.onmicrosoft.com",
    // OJO: tiene que ser EXACTAMENTE la misma URL que ya está registrada en
    // Azure AD (la de transito.html). Azure rechaza cualquier redirectUri
    // que no esté registrada de antemano — por eso NO se arma dinámicamente
    // según la página actual, aunque estemos en bajas-altas.html u otra.
    redirectUri: "https://claudiozzh.github.io/transito-ilo/"
  },
  cache: { cacheLocation: "localStorage" }
};
const graphScopes = ["User.Read", "Files.ReadWrite"];
// OJO: tiene que llevar el mismo prefijo que las otras carpetas — la ruta
// real en OneDrive es AUTOMATIZACION_PADRON_NOMINAL/01_REPORTES_DIARIOS/...,
// no 01_REPORTES_DIARIOS/... suelto. Confirmado directamente contra la
// carpeta real (bug encontrado el 04/08/2026: todos los archivos daban
// "no existe" porque se buscaban en la ruta equivocada).
const REPORTES_BASE = "AUTOMATIZACION_PADRON_NOMINAL/01_REPORTES_DIARIOS";
const CARPETA_TRANSITOS = "AUTOMATIZACION_PADRON_NOMINAL/04_TRANSITOS";
const CARPETA_BAJAS_ALTAS = "AUTOMATIZACION_PADRON_NOMINAL/05_NINOS_ALTA_BAJA";

let msalInstance = null;
let msalLoadError = null;
try {
  if (typeof msal === "undefined") throw new Error("La libreria MSAL no cargo desde el CDN. Revisa tu conexion o vuelve a intentar en unos segundos.");
  msalInstance = new msal.PublicClientApplication(msalConfig);
} catch (e) {
  msalLoadError = e.message;
  console.error(e);
}
let account = null;

/* ================== LOGIN ================== */
async function login() {
  if (msalLoadError) {
    alert("No se pudo iniciar el sistema de login: " + msalLoadError + "\n\nRecarga la pagina (Ctrl+F5) e intenta de nuevo.");
    return;
  }
  try {
    const res = await msalInstance.loginPopup({ scopes: graphScopes });
    account = res.account;
    if (typeof onSignedIn === "function") await onSignedIn();
  } catch (e) {
    alert("Error al iniciar sesión: " + e.message);
    console.error(e);
  }
}
function logout() { msalInstance.logoutPopup(); }
async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error("No hay sesión activa");
  account = accounts[0];
  try {
    const res = await msalInstance.acquireTokenSilent({ scopes: graphScopes, account });
    return res.accessToken;
  } catch (e) {
    const res = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return res.accessToken;
  }
}
async function inicializarLogin() {
  if (!msalInstance) return false;
  await msalInstance.handleRedirectPromise();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) { account = accounts[0]; return true; }
  return false;
}

/* ================== GRAPH HELPERS (OneDrive) ================== */
function encodeURIPath(path) { return path.split("/").map(encodeURIComponent).join("/"); }

async function graphGet(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIPath(path)}:/content`, { headers: { Authorization: "Bearer " + token } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error("Error leyendo " + path + ": " + res.status);
  return await res.json();
}
async function graphGetRaw(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIPath(path)}:/content`, { headers: { Authorization: "Bearer " + token } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error("Error leyendo " + path + ": " + res.status);
  return await res.arrayBuffer();
}
async function graphPutJson(path, obj) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIPath(path)}:/content`, {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body: JSON.stringify(obj, null, 2)
  });
  if (!res.ok) throw new Error("Error guardando " + path + ": " + res.status);
  return await res.json();
}
async function graphPutFile(path, file) {
  const token = await getToken();
  if (file.size > 4 * 1024 * 1024) throw new Error(`El archivo ${file.name} pesa más de 4MB. Comprime la imagen.`);
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIPath(path)}:/content`, {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": file.type || "application/octet-stream" }, body: file
  });
  if (!res.ok) throw new Error("Error subiendo " + file.name + ": " + res.status);
  return await res.json();
}
async function graphListChildren(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIPath(path)}:/children`, { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) return [];
  const data = await res.json();
  return data.value || [];
}

/* ================== FECHAS ================== */
function fechaHoyPeru() {
  const now = new Date();
  const utcMs = now.getTime() + now.getTimezoneOffset() * 60000;
  return new Date(utcMs - 5 * 3600000);
}
function mesNombre(m) {
  const nombres = ["01_Enero","02_Febrero","03_Marzo","04_Abril","05_Mayo","06_Junio","07_Julio","08_Agosto","09_Septiembre","10_Octubre","11_Noviembre","12_Diciembre"];
  return nombres[m];
}
/* Calcula edad exacta en "00a 00m 00d" respetando el calendario real
   (con acarreo correcto de mes/año) — mismo criterio que usa Python
   en el resto del proyecto. */
function calcularEdad(fechaNacISO, ref) {
  const n = new Date(fechaNacISO + "T00:00:00");
  let anios = ref.getFullYear() - n.getFullYear();
  let meses = ref.getMonth() - n.getMonth();
  let dias = ref.getDate() - n.getDate();
  if (dias < 0) {
    meses -= 1;
    const mesAnt = new Date(ref.getFullYear(), ref.getMonth(), 0);
    dias += mesAnt.getDate();
  }
  if (meses < 0) { anios -= 1; meses += 12; }
  return `${anios}a ${meses}m ${dias}d`;
}

/* ================== LECTURA DEL PADRÓN (Excel en el navegador) ==================
   Mismo criterio que encontrarInfoCabeceraCodigo() del sistema de tránsito:
   nunca asumir una fila/columna fija — buscar el texto del encabezado, porque
   el gobierno cambia el formato entre exportaciones sin avisar. */
/* El "!ref" que declara un archivo Excel (hasta dónde dice que llegan sus
   datos) a veces queda desactualizado — un archivo real de Ilo del
   04/08/2026 declaraba "A1:BN5" (5 filas) cuando en realidad tenía 4,481.
   Nunca confiar en "!ref" a ciegas: se recalcula el rango real revisando
   qué celdas existen de verdad. */
function calcularRangoReal(ws) {
  let maxR = 0, maxC = 0, minR = 0, minC = 0, hayDatos = false;
  for (const key in ws) {
    if (key[0] === "!") continue;
    const dec = XLSX.utils.decode_cell(key);
    if (!hayDatos) { minR = maxR = dec.r; minC = maxC = dec.c; hayDatos = true; }
    if (dec.r > maxR) maxR = dec.r;
    if (dec.r < minR) minR = dec.r;
    if (dec.c > maxC) maxC = dec.c;
    if (dec.c < minC) minC = dec.c;
  }
  return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
}

function normalizarTexto(s) {
  return String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").toUpperCase().trim();
}

function encontrarInfoCabeceraCodigo(ws) {
  const range = calcularRangoReal(ws);
  const filaLimite = Math.min(range.s.r + 10, range.e.r);
  let candidatoDebil = null;
  for (let r = range.s.r; r <= filaLimite; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const t = normalizarTexto(cell.v);
      // Coincidencia fuerte: el encabezado dice "PADRON NOMINAL" completo,
      // no solo fragmentos sueltos como "COD" y "PAD" en cualquier orden
      // (eso hacía que una nota explicativa tipo "DNI=1 CUI=2 COD.PAD=4"
      // se confundiera con la columna real del código — bug real detectado
      // el 04/08/2026 contra un archivo ACTIVOS.xlsx real).
      if (t.includes("CODIGO DEL PADRON NOMINAL") || t.includes("CÓDIGO DEL PADRON NOMINAL")) {
        return { fila: r, col: c };
      }
      if (!candidatoDebil && t.includes("COD") && t.includes("PAD")) {
        candidatoDebil = { fila: r, col: c };
      }
    }
  }
  return candidatoDebil;
}

/* Verificación extra: un código de padrón real son 8 dígitos. Si lo que se
   leyó no cumple eso, es señal de que se agarró la columna equivocada —
   mejor avisar que devolver un dato que parece código pero no lo es. */
function pareceCodigoPadronValido(v) {
  return /^\d{6,9}$/.test(String(v || "").trim());
}

function extraerCodigosDeHoja(ws) {
  const info = encontrarInfoCabeceraCodigo(ws);
  if (!info) return [];
  const range = calcularRangoReal(ws);
  const codigos = [];
  for (let r = info.fila + 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: info.col })];
    if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
      const val = String(cell.v).trim();
      if (pareceCodigoPadronValido(val)) codigos.push(val);
    }
  }
  return codigos;
}

/* Mapa de qué palabras buscar en el encabezado para cada campo que
   necesitamos autocompletar. Como el gobierno no siempre nombra las
   columnas igual, cada campo acepta varias variantes. */
const CAMPOS_PADRON = {
  cnv: ["NUMERO CNV", "N CNV"],
  cui: ["CODIGO UNICO DE IDENTI"],
  dni: ["NUMERO DE DOCUMENTO DE IDENTIDAD", "DNI DEL NI"],
  apellidoPaterno: ["APELLIDO PATERNO DEL NI"],
  apellidoMaterno: ["APELLIDO MATERNO DEL NI"],
  nombres: ["NOMBRES DEL NI"],
  fechaNacimiento: ["FECHA DE NACIMIENTO"],
  edad: ["EDAD DEL NI"],
  ejeVial: ["EJE VIAL"],
  descripcionDireccion: ["DESCRIPCION"],
  referenciaDireccion: ["REFERENCIA DE DIRECCION"],
  tipoSeguro: ["TIPO DE SEGURO DEL BENEFICIARIO"],
  eessAdscripcion: ["NOMBRE DEL EESS ADSCRIPCION"],
  vinculoJefe: ["VINCULO", "PARENTESCO"],
  dniJefe: ["NUMERO DE DOCUMENTO DEL JEFE"],
  apellidoPaternoJefe: ["APELLIDO PATERNO DEL JEFE"],
  apellidoMaternoJefe: ["APELLIDO MATERNO DEL JEFE"],
  nombresJefe: ["NOMBRES DEL JEFE"],
  dniMadre: ["NUMERO DE DOCUMENTO DE LA MADRE"],
  apellidoPaternoMadre: ["APELLIDO PATERNO DE LA MADRE"],
  apellidoMaternoMadre: ["APELLIDO MATERNO DE LA MADRE"],
  nombresMadre: ["NOMBRES DE LA MADRE"],
  celularMadre: ["NUMERO DE CELULAR DE LA MADRE"],
  correoMadre: ["CORREO ELECTRONICO DE LA MADRE"],
};
/* Encuentra, en la fila de encabezados ya detectada, la columna de
   cada campo de CAMPOS_PADRON. Devuelve {campo: columna|null, ...} */
function mapearColumnas(ws, filaEncabezado) {
  const range = calcularRangoReal(ws);
  const mapa = {};
  for (const campo in CAMPOS_PADRON) mapa[campo] = null;
  mapa.eessAtencion = null;
  const columnasVinculo = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: filaEncabezado, c })];
    if (!cell || cell.v == null) continue;
    const t = normalizarTexto(cell.v);
    for (const campo in CAMPOS_PADRON) {
      if (mapa[campo] !== null) continue;
      if (CAMPOS_PADRON[campo].some(variante => t.includes(variante))) mapa[campo] = c;
    }
    // "NOMBRE DEL EESS" (Último establecimiento de atención) es EXACTAMENTE
    // esas 3 palabras — pero también es un fragmento literal de "NOMBRE DEL
    // EESS NACIMIENTO" y "NOMBRE DEL EESS ADSCRIPCIÓN", así que necesita
    // coincidencia exacta, no "incluye", o agarraría la columna equivocada.
    if (mapa.eessAtencion === null && t === "NOMBRE DEL EESS") mapa.eessAtencion = c;
    // La columna de vínculo (1=MADRE 2=PADRE 3=HERMANO...) no tiene título
    // propio, solo la leyenda — se detecta por el patrón, no por un nombre.
    if (t.includes("1=MADRE") && t.includes("2=PADRE")) columnasVinculo.push(c);
  }
  // Hay dos columnas así en el padrón: la del bloque "madre" y la del
  // bloque "jefe de familia". La que nos interesa (si el jefe es el padre)
  // es la última de las dos.
  if (columnasVinculo.length) mapa.vinculoJefe = columnasVinculo[columnasVinculo.length - 1];
  return mapa;
}

/* Busca un documento (DNI/CUI/CNV) en una hoja ya cargada y devuelve
   todos los campos de CAMPOS_PADRON que haya podido encontrar, o null
   si el documento no está en esa hoja. */
function buscarDocumentoEnHoja(ws, documento) {
  const info = encontrarInfoCabeceraCodigo(ws);
  if (!info) return null;
  const mapa = mapearColumnas(ws, info.fila);
  const range = calcularRangoReal(ws);
  const val = (r, c) => {
    if (c === null) return "";
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell && cell.v != null ? String(cell.v).trim() : "";
  };
  const docNorm = String(documento).trim();
  for (let r = info.fila + 1; r <= range.e.r; r++) {
    const cnv = val(r, mapa.cnv), cui = val(r, mapa.cui), dni = val(r, mapa.dni);
    if (docNorm !== cnv && docNorm !== cui && docNorm !== dni) continue;
    const cod = val(r, info.col);
    const out = { codigoPadron: cod };
    for (const campo in mapa) out[campo] = val(r, mapa[campo]);
    return out;
  }
  return null;
}

/* Busca un documento en los 3 archivos de un día (ACTIVOS, OBSERVADOS,
   TRANSITOS) dentro de 01_REPORTES_DIARIOS/{fecha}/. Devuelve
   {datos, archivo} o null si no está en ninguno. */
async function buscarEnPadronDelDia(fecha, documento) {
  const archivos = [
    { nombre: "ACTIVOS.xlsx", tag: "ACTIVOS" },
    { nombre: "OBSERVADOS.xlsx", tag: "OBSERVADOS" },
    { nombre: "TRANSITOS.xlsx", tag: "TRANSITO" },
  ];
  const diagnostico = [];
  for (const a of archivos) {
    const path = `${REPORTES_BASE}/${fecha}/${a.nombre}`;
    let buffer;
    try {
      buffer = await graphGetRaw(path);
    } catch (e) {
      diagnostico.push(`${a.nombre}: ERROR al leer — ${e.message}`);
      continue;
    }
    if (!buffer) {
      diagnostico.push(`${a.nombre}: no existe en esa carpeta (${path})`);
      continue;
    }
    let wb, ws;
    try {
      wb = XLSX.read(buffer, { type: "array" });
      ws = wb.Sheets[wb.SheetNames[0]];
    } catch (e) {
      diagnostico.push(`${a.nombre}: se descargó (${buffer.byteLength} bytes) pero no se pudo abrir como Excel — ${e.message}`);
      continue;
    }
    const rango = calcularRangoReal(ws);
    const totalFilas = rango.e.r - rango.s.r;
    const datos = buscarDocumentoEnHoja(ws, documento);
    if (datos) return { datos, archivo: a.tag, diagnostico };
    diagnostico.push(`${a.nombre}: leído bien (${buffer.byteLength} bytes, ~${totalFilas} filas de datos) — el documento no está ahí`);
  }
  return { datos: null, archivo: null, diagnostico };
}

/* Igual que extraerCodigosDeHoja, pero para los 3 archivos de un día
   de una sola vez — para saber si un código YA NO aparece en ninguno
   (confirmar bajas) o SI aparece (confirmar altas/reubicaciones). */
async function codigosDelDia(fecha) {
  const archivos = ["ACTIVOS.xlsx", "OBSERVADOS.xlsx", "TRANSITOS.xlsx"];
  const encontrados = new Set();
  const leidos = [], faltantes = [];
  for (const nombreArchivo of archivos) {
    const path = `${REPORTES_BASE}/${fecha}/${nombreArchivo}`;
    try {
      const buffer = await graphGetRaw(path);
      if (!buffer) { faltantes.push(nombreArchivo); continue; }
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      extraerCodigosDeHoja(ws).forEach(c => encontrados.add(c));
      leidos.push(nombreArchivo);
    } catch (e) {
      faltantes.push(nombreArchivo + " (error: " + e.message + ")");
    }
  }
  return { encontrados, leidos, faltantes };
}
