# Código de Google Apps Script para Google Sheets (Versión Ultra-Compatible & Segura)

Esta guía y código solucionan los problemas de visualización e incompatibilidad cuando despliegas la aplicación como **Aplicación Web** independiente para los Padres de Familia.

### 🔍 ¿Por qué ocurrían los errores?
1. **Error de Hoja Nula (`getActiveSpreadsheet()` es null)**: Al acceder un padre mediante el enlace independiente, no hay una "hoja abierta activa" (active spreadsheet). Ahora el script guarda automáticamente el ID del documento en las propiedades del servidor para resolverlo autónomamente.
2. **Falla de serialización de Fechas**: Google Apps Script arroja un error crítico y no muestra nada (pantalla vacía) cuando intenta enviarle un objeto de tipo `Date` de Javascript al navegador. Ahora filtramos y formateamos cada valor a texto plano (`cleanValue()`) garantizando una visualización al 100%.

---

## 1. Código de Servidor (`Code.gs`)

Crea o reemplaza por completo el archivo `Code.gs` en tu editor de Google Apps Script con este bloque de código corregido:

```javascript
// =========================================================================
// CONFIGURACIÓN OPCIONAL
// =========================================================================
// Si continúas teniendo problemas de permisos o inicialización, puedes colocar
// aquí el ID de tu hoja de cálculo (lo encuentras en la barra de direcciones de tu navegador).
// Ejemplo: "1A2b3C4d5E6f7G8h9I0j..."
var SPREADSHEET_ID_MANUAL = ""; 

/**
 * Función requerida para cargar la Aplicación Web de Google Apps Script.
 * Resuelve el inicio del HTML y permite visualización externa.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SADOSA 2026 - Primaria')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Agrega el menú personalizado a la barra superior de Sheets al abrir el Excel.
 * De paso, guarda el ID de la hoja actual en las propiedades del servidor.
 */
function onOpen() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      PropertiesService.getScriptProperties().setProperty('SAVED_SPREADSHEET_ID', ss.getId());
    }
  } catch(e) {
    Logger.log("Error al guardar ID en onOpen: " + e.toString());
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SADOSA 2026')
    .addItem('🔑 Abrir Panel de Maestros (Barra Lateral)', 'showSidebar')
    .addToUi();
}

/**
 * Muestra el panel interactivo en la barra lateral derecha de Google Sheets.
 */
function showSidebar() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      PropertiesService.getScriptProperties().setProperty('SAVED_SPREADSHEET_ID', ss.getId());
    }
  } catch(e) {}
  
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SADOSA 2026 - Primaria')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Resuelve de forma súper segura el Spreadsheet activo o guardado por ID.
 * Evita que el programa falle cuando los padres acceden externamente.
 */
function getSpreadsheet() {
  // 1. Intentar con ID manual si se configuró arriba
  if (typeof SPREADSHEET_ID_MANUAL !== 'undefined' && SPREADSHEET_ID_MANUAL !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID_MANUAL);
  }
  
  // 2. Intentar capturar la hoja activa si se ejecuta como sidebar o editor directo
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      PropertiesService.getScriptProperties().setProperty('SAVED_SPREADSHEET_ID', ss.getId());
      return ss;
    }
  } catch(e) {}
  
  // 3. Recuperar ID persistido desde propiedades
  var savedId = PropertiesService.getScriptProperties().getProperty('SAVED_SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch(e) {
      throw new Error("No pudimos conectar con el documento de Google Sheets (ID guardado inválido). Abre tu documento de Google Sheets y presiona el menú 'SADOSA 2026 > Abrir Panel de Maestros' para reactivar la sincronización.");
    }
  }
  
  throw new Error("Base de datos desinicializada. Por favor, abre la hoja de cálculo y presiona en el menú superior: 'SADOSA 2026 > Abrir Panel de Maestros' para vincular este formulario.");
}

/**
 * Inicializa u obtiene las hojas "Entrevistas" y "Evaluaciones" automáticamente.
 */
function getOrCreateSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === "Entrevistas") {
      sheet.appendRow(["ID", "Día", "Horario", "Curso", "Profesor", "Asunto", "Estado", "Estudiante", "Padre/Madre"]);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#dbeafe");
    } else if (sheetName === "Evaluaciones") {
      sheet.appendRow(["ID", "Fecha", "Curso", "Trimestre", "Bloque", "Área", "Tema", "Docente"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3e8ff");
    }
  }
  return sheet;
}

/**
 * Función CRÍTICA de conversión avanzada. Evita crasheos de JSON con fechas.
 */
function cleanValue(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    // Formatear fechas como YYYY-MM-DD
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT-4", "yyyy-MM-dd");
  }
  return String(val).trim();
}

/**
 * Nómina oficial de maestros Sadosa 2026
 */
function getTeachersList() {
  return [
    "AGUILAR GOMEZ MIRTZA DENIS",
    "NINA MAMANI LOURDES EULALIA",
    "SALVATIERRA CARTAGENA MARÍA ROSA",
    "TORRICO PRADO MIRIAM PATRICIA",
    "SILES SOTO DEISY ESTHER",
    "CÉSPEDES BERNAL RUTH GLADYS",
    "ZUNA BERNAL JM JOSELIN",
    "MEJIA CESPEDES VIRGINIA AMALIA",
    "HUANCA GUTIERREZ MARTHA",
    "PACO GRANIER GROVER",
    "ARIAS ALBORTA LIZBETH",
    "PIMENTEL FLORES NORKA ZULEMA",
    "ALCON LLUSCO BETZA",
    "QUIÑONES FUENTES EDELFRIDA",
    "CÉSPEDES BERNAL GIOVANA",
    "GUTIERREZ TORRES CLAUDER",
    "SANDIVAR ANTURIANO OMERSICERI MARIBI",
    "ZURITA PAREDES SHANNON",
    "RODRIGUEZ NOGALES JOSE LUIS",
    "EVELIN CHOQUECALLATA",
    "RAMOS PACHECO ERICK",
    "ROJAS MOSCOSO SANDRA",
    "PACO TORRES ALICIA ANABEL",
    "AYZACAYO CONDE JOSE LUIS",
    "CUELLAR UGARTE ELIZABET",
    "MATIAS CAHUAYA MARIA LUISA"
  ];
}

/**
 * Retorna de forma segura toda la información limpia de tipos raros para evitar pantallas vacías.
 */
function getSadosaData() {
  try {
    var interviewsSheet = getOrCreateSheet("Entrevistas");
    var evaluationsSheet = getOrCreateSheet("Evaluaciones");
    
    var intData = interviewsSheet.getDataRange().getValues();
    var evData = evaluationsSheet.getDataRange().getValues();
    
    var interviews = [];
    for (var i = 1; i < intData.length; i++) {
      if (!cleanValue(intData[i][0])) continue; // Saltar filas vacías
      interviews.push({
        id: cleanValue(intData[i][0]),
        day: cleanValue(intData[i][1]),
        timeSlot: cleanValue(intData[i][2]),
        course: cleanValue(intData[i][3]),
        teacherName: cleanValue(intData[i][4]),
        parentReason: cleanValue(intData[i][5]),
        parentStatus: cleanValue(intData[i][6]) || "Pendiente",
        studentName: cleanValue(intData[i][7]),
        parentName: cleanValue(intData[i][8])
      });
    }
    
    var evaluations = [];
    for (var j = 1; j < evData.length; j++) {
      if (!cleanValue(evData[j][0])) continue; // Saltar filas vacías
      evaluations.push({
        id: cleanValue(evData[j][0]),
        date: cleanValue(evData[j][1]),
        course: cleanValue(evData[j][2]),
        trimester: cleanValue(evData[j][3]),
        blockIndex: Number(evData[j][4]) || 0,
        subject: cleanValue(evData[j][5]),
        topic: cleanValue(evData[j][6]),
        teacherName: cleanValue(evData[j][7])
      });
    }
    
    return {
      interviews: interviews,
      evaluations: evaluations,
      teachers: getTeachersList()
    };
  } catch(error) {
    throw new Error("No se pudo extraer la información: " + error.toString());
  }
}

/**
 * Valida la clave maestra del panel docente
 */
function checkMasterPassword(pass) {
  return pass === "sadosaprimaria2026";
}

/**
 * Guarda o actualiza una asignación de entrevista
 */
function saveInterviewFromUI(data) {
  var sheet = getOrCreateSheet("Entrevistas");
  var rows = sheet.getDataRange().getValues();
  var id = data.id || "int-" + Utilities.getUuid().substring(0, 8);
  var foundIndex = -1;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) {
      foundIndex = i + 1;
      break;
    }
  }
  
  var rowData = [
    id,
    data.day,
    data.timeSlot,
    data.course,
    data.teacherName,
    data.parentReason,
    data.parentStatus || "Pendiente",
    data.studentName,
    data.parentName || "Familiar"
  ];
  
  if (foundIndex > -1) {
    sheet.getRange(foundIndex, 1, 1, 9).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return id;
}

/**
 * Guarda el estado de asistencia de la entrevista
 */
function updateInterviewStatus(id, status) {
  var sheet = getOrCreateSheet("Entrevistas");
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) {
      sheet.getRange(i + 1, 7).setValue(status); // Columna G es Estado (7)
      return true;
    }
  }
  return false;
}

/**
 * Guarda una evaluación coordinando que no se crucen las evaluaciones
 */
function saveEvaluationFromUI(data) {
  var sheet = getOrCreateSheet("Evaluaciones");
  var rows = sheet.getDataRange().getValues();
  
  // Validar cruces de horarios en el mismo bloque, curso y fecha de examen
  for (var i = 1; i < rows.length; i++) {
    var rawDate = cleanValue(rows[i][1]);
    var rawCourse = cleanValue(rows[i][2]);
    var rawBlock = Number(rows[i][4]);
    
    if (rawDate == data.date && rawCourse == data.course && rawBlock == data.blockIndex) {
      throw new Error("Este bloque de examen ya está ocupado por la materia de '" + rows[i][5] + "' reservado por " + rows[i][7] + ".");
    }
  }
  
  var id = "eval-" + Utilities.getUuid().substring(0, 8);
  sheet.appendRow([
    id,
    data.date,
    data.course,
    data.trimester,
    data.blockIndex,
    data.subject,
    data.topic,
    data.teacherName
  ]);
  
  return id;
}

/**
 * Elimina registros seleccionados por ID
 */
function deleteRecord(sheetName, id) {
  var sheet = getOrCreateSheet(sheetName);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
```

---

## 2. Interfaz del Cliente (`Index.html`)

Reemplaza todo el contenido del archivo de tipo HTML llamado **`Index`** en tu Apps Script con la siguiente plantilla mejorada. Cuenta con diagnóstico inteligente visible para el usuario si algo falla:

```html
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <base target="_top">
    <!-- Tailwind CSS para diseño fluido y elegante -->
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Outfit:wght@600;800&display=swap');
      body {
        font-family: 'Plus Jakarta Sans', sans-serif;
        background-color: #f8fafc;
      }
      .font-display {
        font-family: 'Outfit', sans-serif;
      }
    </style>
  </head>
  <body class="text-slate-800 p-3 md:p-6 max-w-4xl mx-auto">

    <!-- Encabezado de la Escuela Sadosa 2026 -->
    <div class="bg-gradient-to-r from-blue-800 via-blue-900 to-indigo-950 text-white rounded-2xl p-5 shadow-xl mb-6 relative overflow-hidden">
      <div class="flex flex-col sm:flex-row items-center justify-between gap-4 z-10 relative">
        <div class="text-center sm:text-left">
          <div class="flex items-center gap-2 justify-center sm:justify-start">
            <span class="bg-amber-400 text-slate-900 text-[10px] font-extrabold px-2.5 py-0.5 rounded-full uppercase tracking-wider shadow">
              SADOSA 2026
            </span>
            <span class="bg-blue-700/50 text-white text-[10px] font-bold px-2 py-0.5 rounded-full uppercase">
              Primaria
            </span>
          </div>
          <h1 class="text-lg md:text-xl font-display font-extrabold tracking-tight mt-1">
            Unidad Educativa Santo Domingo Sabio
          </h1>
          <p class="text-[11px] text-blue-200">Plataforma de Entrevistas y Exámenes Sincronizada con Google Sheets</p>
        </div>

        <!-- Botón de Conexión Maestro -->
        <div id="auth-header-container" class="text-center sm:text-right">
          <button id="btn-login-trigger" onclick="openLoginModal()" class="bg-white/10 hover:bg-white/15 border border-white/20 text-white text-xs font-bold px-3 py-2 rounded-xl transition flex items-center gap-1.5 mx-auto">
            🔒 Modificar como Maestro
          </button>
          <div id="badge-teacher-active" class="hidden bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 text-xs font-semibold px-3 py-1.5 rounded-xl flex items-center justify-center gap-1.5">
            <span>🔑 Maestro Activo</span>
            <button onclick="logoutTeacher()" class="text-[10px] text-white underline hover:text-rose-200 ml-1 font-bold">Salir</button>
          </div>
        </div>
      </div>
    </div>




    <!-- Pestañas de Navegación -->
    <div class="flex bg-white p-1 rounded-xl border border-slate-200 shadow-sm mb-6 max-w-md mx-auto">
      <button onclick="switchTab('interviews')" id="tab-int" class="flex-1 py-2.5 text-xs font-bold rounded-lg text-white bg-blue-600 transition">
        Horarios de Entrevista
      </button>
      <button onclick="switchTab('evaluations')" id="tab-ev" class="flex-1 py-2.5 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-50 transition">
        Rol de Evaluaciones
      </button>
    </div>




    <!-- SECCIÓN DE ENTREVISTAS -->
    <div id="section-interviews" class="space-y-6">
      
      <!-- Formulario para Agendar (Solo visible si es Maestro) -->
      <div id="panel-add-interview" class="hidden bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 class="text-sm font-bold text-slate-800 flex items-center gap-1.5 mb-1 text-blue-800">
          ➕ Agendar Cita en el Rol Primario
        </h3>
        <p class="text-xs text-slate-400 mb-4 font-normal">Los padres de familia verán esta reserva de 40 minutos en tiempo real.</p>
        
        <form onsubmit="handleSaveInterview(event)" class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Día de la semana</label>
            <select id="int-day" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
              <option>Lunes</option><option>Martes</option><option>Miércoles</option><option>Jueves</option><option>Viernes</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Horario (40 Minutos)</label>
            <select id="int-time" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs">
              <option>08:00 - 08:40</option>
              <option>08:40 - 09:20</option>
              <option>09:20 - 10:00</option>
              <option>10:00 - 10:40</option>
              <option>10:40 - 11:20</option>
              <option>11:20 - 12:00</option>
              <option>12:00 - 12:40</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Curso</label>
            <select id="int-course" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs">
              <option>Primero A</option><option>Primero B</option><option>Primero C</option>
              <option>Segundo A</option><option>Segundo B</option><option>Segundo C</option>
              <option>Tercero A</option><option>Tercero B</option><option>Tercero C</option>
              <option>Cuarto A</option><option>Cuarto B</option><option>Cuarto C</option>
              <option>Quinto A</option><option>Quinto B</option><option>Quinto C</option>
              <option>Sexto A</option><option>Sexto B</option><option>Sexto C</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Docente Responsable</label>
            <select id="int-teacher" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs"></select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Nombre Completo Alumno</label>
            <input type="text" id="int-student" required placeholder="Ej. Mateo Rojas" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Familiar citados (Papá/Mamá)</label>
            <input type="text" id="int-parent" placeholder="Ej. Clara de Rojas" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
          </div>

          <div class="md:col-span-2">
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Asunto de citación</label>
            <select id="int-reason" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs mb-2">
              <option>Rendimiento académico y notas bajas</option>
              <option>Faltas de disciplina y mala conducta en aula</option>
              <option>Incumplimiento frecuente de tareas y deberes escolares</option>
              <option>Falta de atención, desinterés o desmotivación constante</option>
              <option>Coordinación pedagógica y de conducta integral</option>
            </select>
          </div>

          <div class="md:col-span-2 flex justify-end">
            <button type="submit" class="bg-blue-700 hover:bg-blue-800 text-white font-bold px-5 py-2.5 rounded-xl text-xs transition">
              Programar Entrevista
            </button>
          </div>
        </form>
      </div>

      <!-- Filtros Rápidos -->
      <div class="bg-slate-100 rounded-xl p-3 border border-slate-250 flex flex-wrap gap-2 items-center justify-between">
        <span class="text-xs font-bold text-slate-600">🔍 Buscar por:</span>
        <div class="flex gap-2">
          <select id="filter-course" onchange="renderInterviews()" class="bg-white border rounded-lg px-2 py-1 text-xs outline-none">
            <option value="">Todos los Cursos</option>
            <option>Primero A</option><option>Primero B</option><option>Primero C</option>
            <option>Segundo A</option><option>Segundo B</option><option>Segundo C</option>
            <option>Tercero A</option><option>Tercero B</option><option>Tercero C</option>
            <option>Cuarto A</option><option>Cuarto B</option><option>Cuarto C</option>
            <option>Quinto A</option><option>Quinto B</option><option>Quinto C</option>
            <option>Sexto A</option><option>Sexto B</option><option>Sexto C</option>
          </select>
          <select id="filter-day" onchange="renderInterviews()" class="bg-white border border-slate-200 rounded-lg px-2 py-1 text-xs outline-none">
            <option value="">Todos los Días</option>
            <option>Lunes</option><option>Martes</option><option>Miércoles</option><option>Jueves</option><option>Viernes</option>
          </select>
        </div>
      </div>

      <!-- Lista de entrevistas -->
      <div class="space-y-3">
        <h3 class="text-xs font-extrabold text-slate-500 uppercase tracking-widest">Cronología de Entrevistas:</h3>
        <div id="list-interviews" class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <!-- Render dinámico -->
        </div>
      </div>
    </div>




    <!-- SECCIÓN DE EVALUACIONES -->
    <div id="section-evaluations" class="space-y-6 hidden">
      
      <!-- Formulario para Agendar Exámenes (Solo visible si es Maestro) -->
      <div id="panel-add-evaluation" class="hidden bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 class="text-sm font-bold text-slate-800 flex items-center gap-1.5 mb-1 text-indigo-800">
          📝 Programar Examen
        </h3>
        <p class="text-xs text-slate-400 mb-4 ">Se inhabilitará el bloque de horario para evitar cualquier cruce entre maestros.</p>
        
        <form onsubmit="handleSaveEvaluation(event)" class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Fecha de Examen</label>
            <input type="date" id="ev-date" required class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Trimestre Escolar</label>
            <select id="ev-trimester" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
              <option>Segundo Trimestre</option>
              <option>Tercer Trimestre</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Curso de Primaria</label>
            <select id="ev-course" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
              <option>Primero A</option><option>Primero B</option><option>Primero C</option>
              <option>Segundo A</option><option>Segundo B</option><option>Segundo C</option>
              <option>Tercero A</option><option>Tercero B</option><option>Tercero C</option>
              <option>Cuarto A</option><option>Cuarto B</option><option>Cuarto C</option>
              <option>Quinto A</option><option>Quinto B</option><option>Quinto C</option>
              <option>Sexto A</option><option>Sexto B</option><option>Sexto C</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Área / Materia</label>
            <select id="ev-subject" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
              <option>Lenguaje y comunicación</option>
              <option>Matemática</option>
              <option>Ciencias naturales</option>
              <option>Ciencias sociales</option>
              <option>Técnica, tecnología</option>
              <option>Artes plásticas</option>
              <option>Educación física</option>
              <option>Educación musical</option>
              <option>Valores y espiritualidad</option>
              <option>Gramática</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Bloque de Periodos</label>
            <select id="ev-block" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
              <option value="0">Bloque 1 (1° y 2° Periodo: 08:00 - 09:20)</option>
              <option value="1">Bloque 2 (3° y 4° Periodo: 09:40 - 11:00)</option>
              <option value="2">Bloque 3 (5° y 6° Periodo: 11:00 - 12:20)</option>
            </select>
          </div>

          <div>
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Maestro Evaluador</label>
            <select id="ev-teacher" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none"></select>
          </div>

          <div class="md:col-span-2">
            <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1">Contenido / Temario a examinar</label>
            <input type="text" id="ev-topic" required placeholder="Ej. Tema 3: Divisiones exactas y problemas de aplicación" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none">
          </div>

          <div class="md:col-span-2 flex justify-end">
            <button type="submit" class="bg-indigo-700 hover:bg-indigo-800 text-white font-bold px-5 py-2.5 rounded-xl text-xs transition shadow-md">
              Asignar y Bloquear Espacio
            </button>
          </div>
        </form>
      </div>

      <!-- Filtros Rápidos de evaluaciones -->
      <div class="bg-slate-100 rounded-xl p-3 border border-slate-200 flex flex-wrap gap-2 items-center justify-between">
        <span class="text-xs font-bold text-slate-600">🔍 Filtrar Rol Exámenes:</span>
        <div class="flex gap-2">
          <select id="filter-ev-course" onchange="renderEvaluations()" class="bg-white border rounded-lg px-2 py-1 text-xs outline-none">
            <option value="">Todos los Cursos</option>
            <option>Primero A</option><option>Primero B</option><option>Primero C</option>
            <option>Segundo A</option><option>Segundo B</option><option>Segundo C</option>
            <option>Tercero A</option><option>Tercero B</option><option>Tercero C</option>
            <option>Cuarto A</option><option>Cuarto B</option><option>Cuarto C</option>
            <option>Quinto A</option><option>Quinto B</option><option>Quinto C</option>
            <option>Sexto A</option><option>Sexto B</option><option>Sexto C</option>
          </select>
        </div>
      </div>

      <!-- Tabla / Calendario de Exámenes -->
      <div>
        <h3 class="text-xs font-extrabold text-slate-500 uppercase tracking-widest mb-3">Roles de Evaluación Registrados:</h3>
        <div id="list-evaluations" class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <!-- Render dinámico -->
        </div>
      </div>
    </div>




    <!-- MODAL DE INICIO DE SESIÓN DE MAESTRO -->
    <div id="login-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
      <div class="bg-white rounded-2xl max-w-sm w-full border border-slate-200 shadow-2xl p-5 space-y-4">
        <div>
          <h3 class="text-base font-bold text-slate-800">🔑 Validación de Seguridad</h3>
          <p class="text-xs text-slate-500 mt-1">Los profesores deben ingresar la contraseña de maestro para poder sincronizar cambios en Sheets.</p>
        </div>

        <div>
          <label class="block text-[10px] font-bold text-slate-600 uppercase mb-1.5">Contraseña Maestra</label>
          <input type="password" id="input-password" placeholder="••••••••••••••" autocomplete="current-password" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs outline-none focus:ring-2 focus:ring-blue-500">
        </div>

        <div id="login-error-msg" class="hidden bg-rose-50 text-rose-700 p-2.5 rounded-xl text-xs">
          ⚠️ Contraseña no válida. Intente nuevamente.
        </div>

        <div class="flex items-center justify-end gap-2 border-t pt-3 mt-2">
          <button onclick="closeLoginModal()" class="px-3 py-1.5 border border-slate-200 rounded-xl text-xs hover:bg-slate-50">Cancelar</button>
          <button onclick="submitTeacherLogin()" class="px-4 py-1.5 bg-blue-700 hover:bg-blue-800 text-white rounded-xl text-xs font-bold shadow">Acceder</button>
        </div>
      </div>
    </div>




    <!-- CONTROL DE SCRIPT FRONT-END -->
    <script>
      let isTeacherMode = false;
      let localData = { interviews: [], evaluations: [], teachers: [] };

      // Colores por Materia
      const SUBJECT_COLORS = {
        "Lenguaje y comunicación": "bg-rose-50 border-rose-300 text-rose-800",
        "Matemática": "bg-blue-50 border-blue-300 text-blue-800",
        "Ciencias naturales": "bg-green-50 border-green-300 text-green-800",
        "Ciencias sociales": "bg-amber-50 border-amber-300 text-amber-800",
        "Técnica, tecnología": "bg-teal-50 border-teal-300 text-teal-800",
        "Artes plásticas": "bg-purple-50 border-purple-300 text-purple-800",
        "Educación física": "bg-emerald-50 border-emerald-300 text-emerald-800",
        "Educación musical": "bg-pink-50 border-pink-300 text-pink-800",
        "Valores y espiritualidad": "bg-indigo-50 border-indigo-300 text-indigo-800",
        "Gramática": "bg-violet-50 border-violet-300 text-violet-800"
      };

      window.onload = function() {
        // Recordar sesión del maestro autónomamente
        const remembered = localStorage.getItem("sadosa_teacher_gs_auth");
        if (remembered === "true") {
          isTeacherMode = true;
          updateAuthUI();
        }
        loadSheetData();
      };

      function openLoginModal() {
        document.getElementById("login-modal").classList.remove("hidden");
        document.getElementById("input-password").value = "";
        document.getElementById("login-error-msg").classList.add("hidden");
      }

      function closeLoginModal() {
        document.getElementById("login-modal").classList.add("hidden");
      }

      function submitTeacherLogin() {
        const pass = document.getElementById("input-password").value;
        google.script.run
          .withSuccessHandler(function(isValid) {
            if (isValid) {
              isTeacherMode = true;
              localStorage.setItem("sadosa_teacher_gs_auth", "true");
              updateAuthUI();
              closeLoginModal();
            } else {
              document.getElementById("login-error-msg").classList.remove("hidden");
            }
          })
          .withFailureHandler(function(err) {
            alert("Error de validación: " + err.message);
          })
          .checkMasterPassword(pass);
      }

      function logoutTeacher() {
        isTeacherMode = false;
        localStorage.removeItem("sadosa_teacher_gs_auth");
        updateAuthUI();
      }

      function updateAuthUI() {
        const btnTrigger = document.getElementById("btn-login-trigger");
        const badgeActive = document.getElementById("badge-teacher-active");
        const panelAddInt = document.getElementById("panel-add-interview");
        const panelAddEv = document.getElementById("panel-add-evaluation");

        if (isTeacherMode) {
          btnTrigger.classList.add("hidden");
          badgeActive.classList.remove("hidden");
          panelAddInt.classList.remove("hidden");
          panelAddEv.classList.remove("hidden");
        } else {
          btnTrigger.classList.remove("hidden");
          badgeActive.classList.add("hidden");
          panelAddInt.classList.add("hidden");
          panelAddEv.classList.add("hidden");
        }
        
        renderInterviews();
        renderEvaluations();
      }

      function switchTab(tab) {
        document.getElementById("section-interviews").classList.toggle("hidden", tab !== "interviews");
        document.getElementById("section-evaluations").classList.toggle("hidden", tab !== "evaluations");
        
        const btnInt = document.getElementById("tab-int");
        const btnEv = document.getElementById("tab-ev");
        if (tab === "interviews") {
          btnInt.className = "flex-1 py-2.5 text-xs font-bold rounded-lg text-white bg-blue-600 transition";
          btnEv.className = "flex-1 py-2.5 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-50 transition";
        } else {
          btnInt.className = "flex-1 py-2.5 text-xs font-semibold rounded-lg text-slate-600 hover:bg-slate-50 transition";
          btnEv.className = "flex-1 py-2.5 text-xs font-bold rounded-lg text-white bg-indigo-600 transition";
        }
      }

      function loadSheetData() {
        // feedback visual de carga
        document.getElementById("list-interviews").innerHTML = `
          <div class="col-span-full text-center py-8">
            <span class="text-xs text-slate-400">🔄 Sincronizando datos con Google Sheets...</span>
          </div>`;
        document.getElementById("list-evaluations").innerHTML = `
          <div class="col-span-full text-center py-8">
            <span class="text-xs text-slate-400">🔄 Sincronizando datos con Google Sheets...</span>
          </div>`;

        google.script.run
          .withSuccessHandler(function(response) {
            localData = response;
            renderTeachers();
            renderInterviews();
            renderEvaluations();
          })
          .withFailureHandler(function(err) {
            const errorMsg = `
              <div class="col-span-full bg-rose-50 border border-rose-200 text-rose-800 p-4 rounded-xl text-center">
                <p class="text-xs font-bold">⚠️ Error de sincronización con Google Sheets</p>
                <p class="text-[11px] text-slate-500 mt-1">${err.message}</p>
                <button onclick="loadSheetData()" class="mt-3 bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-bold px-4 py-1.5 rounded-lg">
                  Reintentar Sincronización
                </button>
              </div>`;
            document.getElementById("list-interviews").innerHTML = errorMsg;
            document.getElementById("list-evaluations").innerHTML = errorMsg;
          })
          .getSadosaData();
      }

      function renderTeachers() {
        const selInt = document.getElementById("int-teacher");
        const selEv = document.getElementById("ev-teacher");
        
        if (!selInt || !selEv) return;
        
        const optionsHTML = localData.teachers.map(t => `<option value="${t}">${t}</option>`).join("");
        selInt.innerHTML = optionsHTML;
        selEv.innerHTML = optionsHTML;
      }

      function getStatusBadgeClass(status) {
        if (status === "Asistió") return "bg-emerald-100 text-emerald-800 border-emerald-250";
        if (status === "No asistió") return "bg-rose-100 text-rose-800 border-rose-250";
        if (status === "Retraso") return "bg-amber-100 text-amber-800 border-amber-250";
        return "bg-slate-100 text-slate-600 border-slate-200";
      }

      function renderInterviews() {
        const container = document.getElementById("list-interviews");
        const filterC = document.getElementById("filter-course").value;
        const filterD = document.getElementById("filter-day").value;

        const filtered = (localData.interviews || []).filter(item => {
          return (filterC === "" || item.course === filterC) && 
                 (filterD === "" || item.day === filterD);
        });

        if (filtered.length === 0) {
          container.innerHTML = `
            <div class="col-span-full text-center bg-white border border-slate-200 rounded-xl p-8 text-xs text-slate-400 italic">
              No hay entrevistas programadas en esta sección.
            </div>`;
          return;
        }

        container.innerHTML = filtered.map(item => {
          const teacherControls = isTeacherMode ? `
            <div class="border-t border-slate-100 pt-2 flex justify-between items-center mt-2">
              <div class="flex gap-1">
                <button onclick="updateStatusFromUI('${item.id}', 'Asistió')" class="px-1.5 py-0.5 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold rounded text-[9px]">Asistió</button>
                <button onclick="updateStatusFromUI('${item.id}', 'No asistió')" class="px-1.5 py-0.5 bg-rose-600 hover:bg-rose-700 text-white font-semibold rounded text-[9px]">Faltó</button>
                <button onclick="updateStatusFromUI('${item.id}', 'Retraso')" class="px-1.5 py-0.5 bg-amber-500 hover:bg-amber-600 text-white font-semibold rounded text-[9px]">Retraso</button>
              </div>
              <button onclick="deleteRecordFromSheet('Entrevistas', '${item.id}')" class="text-[9px] text-red-600 font-bold hover:underline">
                Eliminar
              </button>
            </div>
          ` : `
            <div class="border-t border-slate-100 pt-1 mt-2 text-right">
              <span class="text-[9px] text-slate-400 font-mono">Sincronizado</span>
            </div>
          `;

          return `
            <div class="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col justify-between">
              <div>
                <div class="flex justify-between items-center mb-1">
                  <span class="bg-blue-50 text-blue-800 font-extrabold text-[9px] px-2 py-0.5 rounded uppercase">
                    ${item.course}
                  </span>
                  <span class="text-[9px] font-bold px-2 py-0.5 rounded border ${getStatusBadgeClass(item.parentStatus || 'Pendiente')}">
                    ${item.parentStatus || 'Pendiente'}
                  </span>
                </div>
                <h4 class="text-xs font-bold text-slate-800">Estudiante: ${item.studentName}</h4>
                <p class="text-[10px] text-slate-500 mt-0.5">Día: <b>${item.day}</b> • Hora: <b>${item.timeSlot}</b></p>
                <p class="text-[10px] text-slate-600 italic mt-1 bg-slate-50 p-2 rounded">Asunto: ${item.parentReason}</p>
                <p class="text-[9px] text-slate-400 mt-1">Docente: ${item.teacherName}</p>
              </div>
              ${teacherControls}
            </div>
          `;
        }).join("");
      }

      function updateStatusFromUI(id, status) {
        google.script.run
          .withSuccessHandler(function() {
            loadSheetData();
          })
          .updateInterviewStatus(id, status);
      }

      function renderEvaluations() {
        const container = document.getElementById("list-evaluations");
        const filterCourse = document.getElementById("filter-ev-course").value;

        const filtered = (localData.evaluations || []).filter(e => {
          return (filterCourse === "" || e.course === filterCourse);
        });

        if (filtered.length === 0) {
          container.innerHTML = `
            <div class="col-span-full text-center bg-white border border-slate-200 rounded-xl p-8 text-xs text-slate-400 italic">
              No hay exámenes agendados para este curso.
            </div>`;
          return;
        }

        container.innerHTML = filtered.map(item => {
          const colorClass = SUBJECT_COLORS[item.subject] || "bg-slate-50 border-slate-200 text-slate-800";
          const blocks = ["1° y 2° Periodo", "3° y 4° Periodo", "5° y 6° Periodo"];
          const blockLabel = blocks[item.blockIndex] || "Examen";

          const deleteBtn = isTeacherMode ? `
            <button onclick="deleteRecordFromSheet('Evaluaciones', '${item.id}')" class="text-[10px] text-red-500 hover:underline font-bold mt-2">
              Eliminar Evaluación
            </button>
          ` : "";

          return `
            <div class="border-2 rounded-xl p-4 shadow-xs flex flex-col justify-between ${colorClass}">
              <div>
                <div class="flex items-center justify-between mb-1">
                  <span class="text-[9px] font-extrabold uppercase bg-white/65 px-2 py-0.5 rounded-full">
                    ${item.subject}
                  </span>
                  <span class="text-[9px] font-mono opacity-80 font-bold">${item.date}</span>
                </div>
                <h4 class="text-xs font-extrabold text-slate-900">${item.course} - ${item.trimester}</h4>
                <p class="text-[11px] font-medium mt-1 leading-relaxed">Tema: <b>${item.topic}</b></p>
                <p class="text-[10px] opacity-75 mt-0.5">Bloque: ${blockLabel} • Docente: ${item.teacherName}</p>
              </div>
              ${deleteBtn}
            </div>
          `;
        }).join("");
      }

      function handleSaveInterview(e) {
        e.preventDefault();
        const data = {
          day: document.getElementById("int-day").value,
          course: document.getElementById("int-course").value,
          timeSlot: document.getElementById("int-time").value,
          teacherName: document.getElementById("int-teacher").value,
          studentName: document.getElementById("int-student").value,
          parentName: document.getElementById("int-parent").value,
          parentReason: document.getElementById("int-reason").value,
          parentStatus: "Pendiente"
        };

        google.script.run
          .withSuccessHandler(function() {
            document.getElementById("int-student").value = "";
            document.getElementById("int-parent").value = "";
            loadSheetData();
            alert("¡Entrevista programada y guardada exitosamente!");
          })
          .withFailureHandler(function(err) {
            alert("Error al guardar: " + err.message);
          })
          .saveInterviewFromUI(data);
      }

      function handleSaveEvaluation(e) {
        e.preventDefault();
        const data = {
          date: document.getElementById("ev-date").value,
          trimester: document.getElementById("ev-trimester").value,
          course: document.getElementById("ev-course").value,
          subject: document.getElementById("ev-subject").value,
          blockIndex: Number(document.getElementById("ev-block").value),
          topic: document.getElementById("ev-topic").value,
          teacherName: document.getElementById("ev-teacher").value
        };

        google.script.run
          .withSuccessHandler(function() {
            document.getElementById("ev-topic").value = "";
            loadSheetData();
            alert("¡Examen programado en el calendario de evaluaciones!");
          })
          .withFailureHandler(function(err) {
            alert("⚠️ Error: " + err.message);
          })
          .saveEvaluationFromUI(data);
      }

      function deleteRecordFromSheet(sheetName, id) {
        if (confirm("¿Estás seguro de que deseas eliminar permanentemente este registro de Google Sheets?")) {
          google.script.run
            .withSuccessHandler(function() {
              loadSheetData();
            })
            .withFailureHandler(function(err) {
              alert("Error al eliminar: " + err.message);
            })
            .deleteRecord(sheetName, id);
        }
      }
    </script>
  </body>
</html>
```
