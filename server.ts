import express from "express";
import path from "path";
import fs from "fs";
import { createServer as createViteServer } from "vite";

const app = express();
const PORT = 3000;
const DB_FILE = path.join(process.cwd(), "database.json");

// Middleware
app.use(express.json());

// Default teachers list
const DEFAULT_TEACHERS = [
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

// Helper functions for reading/writing DB_FILE
interface DBState {
  interviews: any[];
  evaluations: any[];
  teachers: string[];
}

function readDB(): DBState {
  try {
    if (!fs.existsSync(DB_FILE)) {
      const initialState: DBState = {
        interviews: [],
        evaluations: [],
        teachers: DEFAULT_TEACHERS
      };
      fs.writeFileSync(DB_FILE, JSON.stringify(initialState, null, 2));
      return initialState;
    }
    const raw = fs.readFileSync(DB_FILE, "utf-8");
    const data = JSON.parse(raw);
    let currentTeachers = data.teachers || DEFAULT_TEACHERS;
    
    // Auto-migrate if old teachers are still in database
    if (currentTeachers.some((t: string) => t.includes("Marina Valenzuela") || t.includes("Elizabeth Mamani"))) {
      currentTeachers = DEFAULT_TEACHERS;
      data.teachers = DEFAULT_TEACHERS;
      fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
    }

    return {
      interviews: data.interviews || [],
      evaluations: data.evaluations || [],
      teachers: currentTeachers
    };
  } catch (error) {
    console.error("Error reading database:", error);
    return { interviews: [], evaluations: [], teachers: DEFAULT_TEACHERS };
  }
}

function writeDB(data: DBState) {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2));
  } catch (error) {
    console.error("Error writing database:", error);
  }
}

// REST API Endpoints

// Get all DB stats
app.get("/api/data", (req, res) => {
  const db = readDB();
  res.json(db);
});

// Update or Create an Interview Slot
app.post("/api/interviews", (req, res) => {
  const { id, day, timeSlot, course, teacherName, parentReason, parentStatus, parentName, studentName } = req.body;
  if (!day || !timeSlot || !course || !teacherName || !parentReason) {
    return res.status(400).json({ error: "Faltan campos obligatorios para agendar la entrevista" });
  }

  const db = readDB();
  const existingIndex = db.interviews.findIndex((item) => item.id === id);

  const interviewData = {
    id: id || `interview-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    day,
    timeSlot,
    course,
    teacherName,
    parentReason,
    parentStatus: parentStatus || "Pendiente",
    parentName: parentName || "",
    studentName: studentName || ""
  };

  if (existingIndex > -1) {
    db.interviews[existingIndex] = interviewData;
  } else {
    db.interviews.push(interviewData);
  }

  writeDB(db);
  res.json({ success: true, interview: interviewData });
});

// Update only Parent Attendance Status
app.post("/api/interviews/status", (req, res) => {
  const { id, parentStatus } = req.body;
  if (!id || !parentStatus) {
    return res.status(400).json({ error: "Falta ID o estado" });
  }

  const db = readDB();
  const index = db.interviews.findIndex((item) => item.id === id);
  if (index === -1) {
    return res.status(404).json({ error: "Entrevista no encontrada" });
  }

  db.interviews[index].parentStatus = parentStatus;
  writeDB(db);
  res.json({ success: true, interview: db.interviews[index] });
});

// Delete an Interview Slot
app.delete("/api/interviews/:id", (req, res) => {
  const { id } = req.params;
  const db = readDB();
  db.interviews = db.interviews.filter((item) => item.id !== id);
  writeDB(db);
  res.json({ success: true });
});

// Update or occupy an Evaluation Slot
app.post("/api/evaluations", (req, res) => {
  const { date, course, trimester, blockIndex, subject, teacherName, topic } = req.body;
  if (!date || !course || !trimester || blockIndex === undefined || !subject) {
    return res.status(400).json({ error: "Faltan campos obligatorios para el rol de evaluación" });
  }

  const db = readDB();

  // Validate if the same slot (date, course, block) is already occupied
  const existingOverlap = db.evaluations.find(
    (ev) => ev.date === date && ev.course === course && ev.blockIndex === blockIndex
  );

  if (existingOverlap) {
    return res.status(409).json({
      error: `Este bloque de examen ya está ocupado por el área '${existingOverlap.subject}' asignado por ${existingOverlap.teacherName}.`
    });
  }

  const evaluationData = {
    id: `eval-${date}-${course}-${blockIndex}`,
    date,
    course,
    trimester,
    blockIndex,
    subject,
    teacherName: teacherName || "Profesor General",
    topic: topic || ""
  };

  db.evaluations.push(evaluationData);
  writeDB(db);
  res.json({ success: true, evaluation: evaluationData });
});

// Delete an Evaluation
app.delete("/api/evaluations/:id", (req, res) => {
  const { id } = req.params;
  const db = readDB();
  db.evaluations = db.evaluations.filter((ev) => ev.id !== id);
  writeDB(db);
  res.json({ success: true });
});

// Add a Teacher to the custom lists
app.post("/api/teachers", (req, res) => {
  const { name } = req.body;
  if (!name || name.trim() === "") {
    return res.status(400).json({ error: "Nombre inválido" });
  }

  const db = readDB();
  const trimmedName = name.trim();
  if (!db.teachers.includes(trimmedName)) {
    db.teachers.push(trimmedName);
    writeDB(db);
  }
  res.json({ success: true, teachers: db.teachers });
});

// Verify Master Password
app.post("/api/verify-password", (req, res) => {
  const { password } = req.body;
  if (password === "sadosaprimaria2026") {
    res.json({ success: true });
  } else {
    res.status(401).json({ error: "Contraseña maestra incorrecta" });
  }
});

// Integration of Vite middleware and production building
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running at http://0.0.0.0:${PORT}`);
  });
}

startServer();
