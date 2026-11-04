// ============================================================
// Aeries to SkillsUSA Roster Converter
// ============================================================

// ---------- Element references ----------
const steps = {
  1: document.getElementById("step-1"),
  2: document.getElementById("step-2"),
  3: document.getElementById("step-3"),
  4: document.getElementById("step-4"),
};

const templateInput   = document.getElementById("template-input");
const templateName    = document.getElementById("template-name");
const templateMessage = document.getElementById("template-message");
const templatePicker  = document.getElementById("template-picker");
const changeTemplateBtn = document.getElementById("change-template-btn");
const programSelect   = document.getElementById("program-select");
const programDetails  = document.getElementById("program-details");
const rosterInput     = document.getElementById("roster-input");
const rosterName      = document.getElementById("roster-name");
const rosterMessage   = document.getElementById("roster-message");
const rosterBtnText   = document.getElementById("roster-btn-text");
const studentPreview  = document.getElementById("student-preview");
const studentTableBody = document.getElementById("student-table-body");
const downloadBtn     = document.getElementById("download-btn");
const downloadMessage = document.getElementById("download-message");

// ---------- Settings ----------
const MAX_FILE_BYTES = 5 * 1024 * 1024; // 5 MB
const READ_TIMEOUT_MS = 15000;
const WORKER_URL = "js/sheet-reader.js";
const MAX_VALUE_LENGTH = 200;

// ---------- App state ----------
const state = {
  templateFileName: "",
  // Parsed template, kept so the output file can match it later:
  // { sheetName, headers: [...all column headers in order], programs: [...] }
  template: null,
  // Each program: { chapterId, chapterName, sectionId, trainingProgram }
  programs: [],
  selectedProgram: null,
  rosterFile: null,
  students: [], // [{ lastName, firstName, email, gender }], unique across all classes
};

// ============================================================
// TEMPLATE PARSING
// ============================================================

// Columns we must find in the template, keyed by the name used in code.
const TEMPLATE_COLUMNS = {
  chapterId: "ChapterID",
  chapterName: "ChapterName",
  sectionId: "SectionID",
  trainingProgram: "TrainingProgram",
};

// An error whose message is safe to show the user as-is.
class TemplateError extends Error {}

// "Training Program", " trainingprogram " and "TrainingProgram" all match.
function normalizeHeader(value) {
  return String(value ?? "").replace(/\s+/g, "").toLowerCase();
}

function cleanCell(value) {
  return String(value ?? "").trim();
}

// Values starting with characters a name, email or ID never starts with.
function looksLikeFormula(value) {
  return /^[=+\-@\t\r]/.test(String(value ?? ""));
}

function isUnexpectedValue(value) {
  return looksLikeFormula(value) || String(value ?? "").length > MAX_VALUE_LENGTH;
}

function isValidEmail(value) {
  return value.length <= 254 && /^[^\s@,;<>"]+@[^\s@,;<>"]+\.[^\s@,;<>"]+$/.test(value);
}

// ============================================================
// FILE READING
// ============================================================

let workerUrlPolicy = null;
function workerScriptUrl() {
  if (window.trustedTypes && window.trustedTypes.createPolicy) {
    if (!workerUrlPolicy) {
      workerUrlPolicy = window.trustedTypes.createPolicy("worker-url", {
        createScriptURL: (url) => {
          if (url !== WORKER_URL) throw new TypeError("Unexpected worker URL");
          return url;
        },
      });
    }
    return workerUrlPolicy.createScriptURL(WORKER_URL);
  }
  return WORKER_URL;
}

// Reads a file in a background worker so the page stays responsive.
// Falls back to reading on the page if workers can't start (e.g. file://).
function readSheets(buffer) {
  return new Promise((resolve, reject) => {
    let worker;
    try {
      worker = new Worker(workerScriptUrl());
    } catch (err) {
      try { resolve(readWorkbookRows(buffer)); } catch (e) { reject(e); }
      return;
    }

    let ready = false;
    const finish = () => {
      clearTimeout(timer);
      worker.terminate();
    };
    const timer = setTimeout(() => {
      finish();
      reject(new SheetReadError("timeout"));
    }, READ_TIMEOUT_MS);

    worker.onmessage = (event) => {
      const msg = event.data || {};
      if (msg.type === "ready") {
        ready = true;
        return;
      }
      finish();
      if (msg.type === "result" && Array.isArray(msg.sheets)) resolve(msg.sheets);
      else reject(new SheetReadError(typeof msg.code === "string" ? msg.code : "read"));
    };

    worker.onerror = (event) => {
      event.preventDefault();
      finish();
      if (ready) {
        reject(new SheetReadError("read"));
        return;
      }
      try { resolve(readWorkbookRows(buffer)); } catch (e) { reject(e); }
    };

    worker.postMessage(buffer);
  });
}

function readErrorMessage(err, fallback) {
  const code = err && err.code;
  if (code === "too-many-sheets") {
    return "This file has more sheets than expected. Check that you chose the right file.";
  }
  if (code === "too-large") {
    return "This file has more rows or columns than expected. Check that you chose the right file.";
  }
  if (code === "timeout") {
    return "This file took too long to read and was stopped. Check that you chose the right file.";
  }
  if (code === "missing-library") {
    return "The spreadsheet reader didn't load. Reload the page and try again.";
  }
  return fallback;
}

// Takes [{ name, rows }] and returns { sheetName, headers, programs, skippedCount }.
// Looks through every sheet for a header row containing "TrainingProgram".
function parseTemplate(sheets) {
  for (const { name: sheetName, rows } of sheets) {
    const headerRowIndex = rows.findIndex((row) =>
      row.some((cell) => normalizeHeader(cell) === "trainingprogram")
    );
    if (headerRowIndex === -1) continue;

    const headers = rows[headerRowIndex].map(cleanCell);
    if (headers.some(isUnexpectedValue)) {
      throw new TemplateError(
        "This template has unexpected content in its column headers. Download a fresh template from SkillsUSA and try again."
      );
    }

    // Find the column position of each required header.
    const col = {};
    const missing = [];
    for (const [key, label] of Object.entries(TEMPLATE_COLUMNS)) {
      const index = headers.findIndex(
        (h) => normalizeHeader(h) === normalizeHeader(label)
      );
      if (index === -1) missing.push(label);
      else col[key] = index;
    }
    if (missing.length) {
      throw new TemplateError(
        `This template is missing the ${missing.join(", ")} column${missing.length > 1 ? "s" : ""}. Download a fresh template from SkillsUSA and try again.`
      );
    }

    // One entry per unique training program + section.
    const programs = [];
    const seen = new Set();
    let skippedCount = 0;
    for (const row of rows.slice(headerRowIndex + 1)) {
      const program = {
        chapterId: cleanCell(row[col.chapterId]),
        chapterName: cleanCell(row[col.chapterName]),
        sectionId: cleanCell(row[col.sectionId]),
        trainingProgram: cleanCell(row[col.trainingProgram]),
      };
      if (!program.trainingProgram) continue;
      if (Object.values(program).some(isUnexpectedValue)) {
        skippedCount++;
        continue;
      }

      const key = `${program.sectionId}|${program.trainingProgram}`;
      if (seen.has(key)) continue;
      seen.add(key);
      programs.push(program);
    }

    if (!programs.length) {
      throw new TemplateError(
        skippedCount
          ? "None of the training programs in this template could be used because they contain unexpected characters."
          : "This template has a TrainingProgram column, but no programs are listed under it."
      );
    }

    return { sheetName, headers, programs, skippedCount };
  }

  throw new TemplateError(
    "No TrainingProgram column was found in this file. Check that you uploaded the SkillsUSA roster template."
  );
}

// ============================================================
// ROSTER PARSING (Aeries "Attendance Class Roster" export)
// ============================================================
//
// Layout notes from the sample export:
// - Each class starts with banner rows (school name, report title,
//   Period/Course/Term/Teacher, a blank row). These are skipped.
// - Each class then has a header row. The three headers we use sit in the
//   same column as their data (Student Name E, GN G, SEM K), so columns are
//   located from that header row rather than hardcoded.
// - Student Name is one cell in "LastName, FirstName" format.
// - Student ID, GR (grade) and the attendance columns are ignored.

const ROSTER_COLUMNS = {
  name: "Student Name",
  gender: "GN",
  email: "SEM",
};

class RosterError extends Error {}

// Returns { name: i, gender: i, email: i } if this row is a roster header
// row, otherwise null.
function findRosterHeader(row) {
  const normalized = row.map(normalizeHeader);
  const col = {};
  for (const [key, label] of Object.entries(ROSTER_COLUMNS)) {
    const index = normalized.indexOf(normalizeHeader(label));
    if (index === -1) return null;
    col[key] = index;
  }
  return col;
}

function stripMarkers(value) {
  return value.replace(/\*/g, "").replace(/\s+/g, " ").trim();
}

// "Doe, Jane" -> { lastName: "Doe", firstName: "Jane" }.
// Splits at the first comma. Returns null if the cell isn't in that format.
// Aeries marks some students with an asterisk (e.g. "*Doe" or "Doe*");
// asterisks are removed so only the name is kept.
function splitStudentName(value) {
  const text = cleanCell(value);
  const comma = text.indexOf(",");
  if (comma === -1) return null;

  const lastName = stripMarkers(text.slice(0, comma));
  const firstName = stripMarkers(text.slice(comma + 1));
  if (!lastName || !firstName) return null;
  return { lastName, firstName };
}

// Takes [{ name, rows }] and returns
// { students, classCount, duplicateCount, missingEmailCount, invalidEmailCount, unexpectedCount }.
function parseRoster(sheets) {
  const students = [];
  const seen = new Set();
  let classCount = 0;
  let duplicateCount = 0;
  let missingEmailCount = 0;
  let invalidEmailCount = 0;
  let unexpectedCount = 0;

  for (const { rows } of sheets) {
    let col = null; // column map from the most recent header row

    for (const row of rows) {
      const header = findRosterHeader(row);
      if (header) {
        col = header;
        classCount++;
        continue;
      }
      if (!col) continue; // banner rows before the first class

      // Banner rows have nothing in "Last, First" form in the name column,
      // so this check skips them.
      const name = splitStudentName(row[col.name]);
      if (!name) continue;

      const email = cleanCell(row[col.email]);
      const gender = cleanCell(row[col.gender]);

      if (!email) {
        missingEmailCount++;
        continue;
      }
      if ([name.lastName, name.firstName, email, gender].some(isUnexpectedValue)) {
        unexpectedCount++;
        continue;
      }
      if (!isValidEmail(email)) {
        invalidEmailCount++;
        continue;
      }

      // Same student in two classes = same name and same email.
      const key = [name.lastName, name.firstName, email]
        .map((v) => v.toLowerCase())
        .join("|");
      if (seen.has(key)) {
        duplicateCount++;
        continue;
      }
      seen.add(key);
      students.push({ lastName: name.lastName, firstName: name.firstName, email, gender });
    }
  }

  if (classCount === 0) {
    throw new RosterError(
      "No class roster was found in this file. Check that you uploaded the Aeries Class Roster export; it needs Student Name, GN and SEM columns."
    );
  }
  if (!students.length) {
    throw new RosterError(
      "No students with a valid email address in the SEM column were found in this file."
    );
  }

  return { students, classCount, duplicateCount, missingEmailCount, invalidEmailCount, unexpectedCount };
}

// ============================================================
// OUTPUT FILE
// ============================================================

// Uses the template's own column order. Every template column is included;
// columns we don't fill are left blank.
function buildOutputRows(template, program, students) {
  // Drop empty trailing header cells, keep everything else as-is.
  const headers = [...template.headers];
  while (headers.length && !headers[headers.length - 1]) headers.pop();

  const valueFor = new Map([
    ["chapterid",       () => toNumberIfSafe(program.chapterId)],
    ["chaptername",     () => program.chapterName],
    ["sectionid",       () => toNumberIfSafe(program.sectionId)],
    ["trainingprogram", () => program.trainingProgram],
    ["memberfirst",     (s) => s.firstName],
    ["memberlast",      (s) => s.lastName],
    ["gender",          (s) => s.gender],
    ["email",           (s) => s.email],
  ]);

  const rows = [headers];
  for (const student of students) {
    rows.push(
      headers.map((h) => {
        const getter = valueFor.get(normalizeHeader(h));
        return getter ? getter(student) ?? "" : "";
      })
    );
  }
  return rows;
}

// The template stores IDs as numbers; match that, but never drop a leading zero.
function toNumberIfSafe(value) {
  return /^[1-9]\d{0,14}$/.test(value) ? Number(value) : value;
}

// Excel sheet names: max 31 characters, none of [ ] : * ? / \
function safeSheetName(name) {
  const cleaned = String(name ?? "")
    .replace(/[\[\]:*?\/\\]/g, "")
    .replace(/^'+|'+$/g, "")
    .trim()
    .slice(0, 31);
  return cleaned || "Sheet1";
}

// "Program A (TeacherA)" -> "Program_A_TeacherA"
function safeFileName(text) {
  return text.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "Program";
}

function todayStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function downloadOutput() {
  const rows = buildOutputRows(state.template, state.selectedProgram, state.students);

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  // Readable column widths based on the longest value in each column.
  sheet["!cols"] = rows[0].map((_, c) => ({
    wch: Math.min(50, Math.max(...rows.map((r) => String(r[c] ?? "").length)) + 2),
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(state.template.sheetName));

  const fileName = `SkillsUSA_${safeFileName(state.selectedProgram.trainingProgram)}_${todayStamp()}.xlsx`;

  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  return { fileName, count: rows.length - 1 };
}

// ============================================================
// UI HELPERS
// ============================================================

// Mark one step as current: earlier steps complete, later ones locked.
function goToStep(n) {
  // Any change to an earlier step makes a previous download message stale.
  setMessage(downloadMessage, "");

  Object.entries(steps).forEach(([num, el]) => {
    const i = Number(num);
    el.classList.remove("is-locked", "is-active", "is-complete");

    if (i < n) {
      el.classList.add("is-complete");
    } else if (i === n) {
      el.classList.add("is-active");
    } else {
      el.classList.add("is-locked");
    }
  });

  // Locked steps are disabled through their own controls; the help
  // buttons stay usable on every step.
  programSelect.disabled = n < 2;
  rosterInput.disabled   = n < 3;
  downloadBtn.disabled   = n < 4;
}

function showFileName(labelEl, file) {
  if (file) {
    labelEl.textContent = file.name;
    labelEl.classList.add("has-file");
  } else {
    labelEl.textContent = "No file chosen";
    labelEl.classList.remove("has-file");
  }
}

// type: "success", "error", or "" for neutral text
function setMessage(el, text, type = "") {
  el.textContent = text;
  el.classList.remove("is-success", "is-error");
  if (type) el.classList.add(`is-${type}`);
}

// Fill the Step 2 dropdown. Option values are array indexes into state.programs.
function setProgramPlaceholder(text) {
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = text;
  programSelect.replaceChildren(placeholder);
}

function populatePrograms(programs) {
  setProgramPlaceholder("Select a training program");

  programs.forEach((p, index) => {
    const opt = document.createElement("option");
    opt.value = String(index);
    opt.textContent = p.trainingProgram;
    programSelect.appendChild(opt);
  });
}

function showStudentPreview(students) {
  studentTableBody.replaceChildren();
  students.forEach((s) => {
    const tr = document.createElement("tr");
    [s.lastName, s.firstName, s.email, s.gender].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    studentTableBody.appendChild(tr);
  });
  studentPreview.hidden = students.length === 0;
  studentPreview.open = false;
}

function showProgramDetails(program) {
  programDetails.textContent = program
    ? `${program.chapterName} (Chapter ID ${program.chapterId}), Section ID ${program.sectionId}`
    : "";
}

// ============================================================
// SESSION MEMORY
// ============================================================
// The parsed template and chosen program are kept in sessionStorage
// (this browser tab only, cleared when the tab closes). If the page is
// reloaded by anything outside the app, Steps 1 and 2 come back as they were.
// Rosters are never stored.

const SESSION_KEY = "skillsusa-roster-converter";

function saveSession() {
  try {
    if (!state.template) {
      sessionStorage.removeItem(SESSION_KEY);
      return;
    }
    const selectedIndex = state.selectedProgram
      ? state.programs.indexOf(state.selectedProgram)
      : -1;
    sessionStorage.setItem(
      SESSION_KEY,
      JSON.stringify({
        templateFileName: state.templateFileName,
        template: {
          sheetName: state.template.sheetName,
          headers: state.template.headers,
          programs: state.template.programs,
        },
        selectedIndex,
      })
    );
  } catch (err) {
    // Storage can be unavailable (private mode, file:// in some browsers).
  }
}

function isValidSavedTemplate(saved) {
  const isString = (v) => typeof v === "string";
  const isCleanString = (v) => isString(v) && !isUnexpectedValue(v);
  const t = saved && saved.template;
  return (
    !!t &&
    isString(saved.templateFileName) &&
    Number.isInteger(saved.selectedIndex) &&
    isString(t.sheetName) &&
    Array.isArray(t.headers) && t.headers.every(isCleanString) &&
    Array.isArray(t.programs) && t.programs.length > 0 &&
    t.programs.every((p) =>
      p && isCleanString(p.chapterId) && isCleanString(p.chapterName) &&
      isCleanString(p.sectionId) && isCleanString(p.trainingProgram) &&
      p.trainingProgram !== ""
    )
  );
}

function restoreSession() {
  try {
    const saved = JSON.parse(sessionStorage.getItem(SESSION_KEY) || "null");
    if (!isValidSavedTemplate(saved)) {
      sessionStorage.removeItem(SESSION_KEY);
      return false;
    }

    applyTemplate(
      {
        sheetName: saved.template.sheetName,
        headers: saved.template.headers,
        skippedCount: 0,
        programs: saved.template.programs.map((p) => ({
          chapterId: p.chapterId,
          chapterName: p.chapterName,
          sectionId: p.sectionId,
          trainingProgram: p.trainingProgram,
        })),
      },
      saved.templateFileName
    );
    if (saved.selectedIndex >= 0) {
      selectProgram(saved.selectedIndex);
    }
    return true;
  } catch (err) {
    return false;
  }
}

// ============================================================
// STATE CHANGES
// ============================================================

// Soft reset of Step 3 and later: clears the roster only.
function clearRoster() {
  state.rosterFile = null;
  state.students = [];
  rosterInput.value = "";
  showFileName(rosterName, null);
  rosterBtnText.textContent = "Choose roster file";
  setMessage(rosterMessage, "");
  showStudentPreview([]);
}

// Soft reset of Step 2 and later: clears the program choice and roster.
function clearProgram() {
  state.selectedProgram = null;
  programSelect.value = "";
  showProgramDetails(null);
  clearRoster();
}

// Soft reset back to Step 1: clears everything.
function clearTemplate() {
  state.templateFileName = "";
  state.template = null;
  state.programs = [];
  templateInput.value = "";
  showFileName(templateName, null);
  setMessage(templateMessage, "");
  setProgramPlaceholder("Upload a template first");
  templatePicker.hidden = false;
  changeTemplateBtn.hidden = true;
  clearProgram();
  saveSession();
}

// Put a parsed template into the page and unlock Step 2.
function applyTemplate(template, fileName) {
  state.template = template;
  state.templateFileName = fileName;
  state.programs = template.programs;
  populatePrograms(template.programs);

  templateName.textContent = fileName;
  templateName.classList.add("has-file");
  templatePicker.hidden = true;
  changeTemplateBtn.hidden = false;

  const count = template.programs.length;
  let text = `Found ${count} training program${count === 1 ? "" : "s"}.`;
  if (template.skippedCount) {
    text += ` ${template.skippedCount} skipped because ${template.skippedCount === 1 ? "it contains" : "they contain"} unexpected characters.`;
  }
  setMessage(templateMessage, text, "success");
  goToStep(2);
  saveSession();
}

function selectProgram(index) {
  const program = state.programs[index];
  if (!program) return;
  state.selectedProgram = program;
  programSelect.value = String(index);
  showProgramDetails(program);
  goToStep(3);
  saveSession();
}

// ============================================================
// STEP HANDLERS
// ============================================================

// ---------- Step 1: template upload ----------
templateInput.addEventListener("change", async () => {
  const file = templateInput.files[0] || null;
  if (!file) return; // picker was cancelled; keep things as they are

  showFileName(templateName, file);

  if (!/\.xlsx$/i.test(file.name)) {
    setMessage(templateMessage, "The template must be an .xlsx file. Choose the file you downloaded from SkillsUSA.", "error");
    templateInput.value = "";
    return;
  }

  if (file.size > MAX_FILE_BYTES) {
    setMessage(templateMessage, "This file is larger than 5 MB, which is too big for a SkillsUSA template. Check that you chose the right file.", "error");
    templateInput.value = "";
    return;
  }

  setMessage(templateMessage, "Reading template…");

  try {
    const sheets = await readSheets(await file.arrayBuffer());
    applyTemplate(parseTemplate(sheets), file.name);
  } catch (err) {
    console.error(err);
    const text = err instanceof TemplateError
      ? err.message
      : readErrorMessage(err, "This file couldn't be read as an Excel workbook. Try downloading the template again.");
    setMessage(templateMessage, text, "error");
    templateInput.value = "";
  }
});

// "Upload a different template": soft reset to Step 1.
changeTemplateBtn.addEventListener("click", () => {
  clearTemplate();
  goToStep(1);
  templateInput.focus();
});

// ---------- Step 2: program selection ----------
// Changing the program is a soft reset to Step 3: the template stays,
// the roster is cleared so a new one can be uploaded for this program.
programSelect.addEventListener("change", () => {
  const value = programSelect.value;
  clearProgram();

  if (value === "") {
    goToStep(2);
    saveSession();
    return;
  }
  selectProgram(Number(value));
});

// ---------- Step 3: roster upload ----------
// Choosing a new roster replaces the current one; template and program stay.
rosterInput.addEventListener("change", async () => {
  const file = rosterInput.files[0] || null;
  if (!file) return; // picker was cancelled; keep the current roster

  clearRoster();
  showFileName(rosterName, file);
  goToStep(3);

  if (!/\.xlsx$/i.test(file.name)) {
    setMessage(rosterMessage, "The roster must be an .xlsx file. Export it from Aeries as Excel.", "error");
    rosterInput.value = "";
    return;
  }

  if (file.size > MAX_FILE_BYTES) {
    setMessage(rosterMessage, "This file is larger than 5 MB, which is too big for a class roster. Check that you chose the right file.", "error");
    rosterInput.value = "";
    return;
  }

  setMessage(rosterMessage, "Reading roster…");

  try {
    const sheets = await readSheets(await file.arrayBuffer());
    const result = parseRoster(sheets);

    state.rosterFile = file;
    state.students = result.students;
    showStudentPreview(result.students);
    rosterBtnText.textContent = "Choose a different roster";

    const n = result.students.length;
    const parts = [
      `Found ${n} student${n === 1 ? "" : "s"} across ${result.classCount} class${result.classCount === 1 ? "" : "es"}.`,
    ];
    if (result.duplicateCount) {
      parts.push(`${result.duplicateCount} duplicate${result.duplicateCount === 1 ? "" : "s"} removed.`);
    }
    if (result.missingEmailCount) {
      parts.push(`${result.missingEmailCount} skipped for a blank SEM email.`);
    }
    if (result.invalidEmailCount) {
      parts.push(`${result.invalidEmailCount} skipped for an invalid email address.`);
    }
    if (result.unexpectedCount) {
      parts.push(`${result.unexpectedCount} skipped for unexpected characters.`);
    }
    setMessage(rosterMessage, parts.join(" "), "success");
    goToStep(4);
  } catch (err) {
    console.error(err);
    const text = err instanceof RosterError
      ? err.message
      : readErrorMessage(err, "This file couldn't be read as an Excel workbook. Try exporting the roster from Aeries again.");
    setMessage(rosterMessage, text, "error");
  } finally {
    // Clearing the input lets the same file be chosen again later.
    rosterInput.value = "";
  }
});

// ---------- Step 4: download ----------
// Downloading changes nothing on the page, so another roster or program
// can be used straight away.
downloadBtn.addEventListener("click", () => {
  if (!state.template || !state.selectedProgram || !state.students.length) {
    setMessage(downloadMessage, "Finish Steps 1 to 3 before downloading.", "error");
    return;
  }
  try {
    const { fileName, count } = downloadOutput();
    setMessage(downloadMessage, `Downloaded ${fileName} with ${count} student${count === 1 ? "" : "s"}.`, "success");
  } catch (err) {
    console.error(err);
    setMessage(downloadMessage, "The Excel file couldn't be created. Reload the page and try again.", "error");
  }
});

// ---------- Help panels ("i" buttons) ----------
// Each .info-btn has aria-controls pointing at its .info-panel.
// Only one panel is open at a time. The panel closes from its "i" button,
// its X button, or the Escape key.

const infoButtons = document.querySelectorAll(".info-btn");

function setHelpOpen(button, open) {
  const panel = document.getElementById(button.getAttribute("aria-controls"));
  if (!panel) return;
  button.setAttribute("aria-expanded", String(open));
  panel.hidden = !open;
}

infoButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const opening = button.getAttribute("aria-expanded") !== "true";
    infoButtons.forEach((b) => setHelpOpen(b, false));
    if (opening) setHelpOpen(button, true);
  });
});

document.querySelectorAll(".info-close").forEach((closeBtn) => {
  closeBtn.addEventListener("click", () => {
    const panel = closeBtn.closest(".info-panel");
    const button = document.querySelector(`.info-btn[aria-controls="${CSS.escape(panel.id)}"]`);
    if (button) {
      setHelpOpen(button, false);
      button.focus();
    }
  });
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  // Escape with an enlarged image open closes only the image.
  if (lightbox.open) return;
  infoButtons.forEach((button) => {
    if (button.getAttribute("aria-expanded") === "true") {
      setHelpOpen(button, false);
      button.focus();
    }
  });
});

// ---------- Help images: hints and enlarged view ----------
// Any link inside a help panel that wraps an image opens that image
// enlarged over the page instead of in a new tab. Paragraphs directly
// above an image get a short "click to enlarge" hint.

const lightbox = document.getElementById("lightbox");
const lightboxImg = document.getElementById("lightbox-img");

document.querySelectorAll(".info-panel-body a").forEach((link) => {
  const img = link.querySelector("img");
  if (!img) return;

  link.classList.add("image-link");
  link.removeAttribute("target"); // the overlay replaces the new tab

  const prev = link.previousElementSibling;
  if (prev && prev.tagName === "P" && !prev.querySelector(".enlarge-hint")) {
    const hint = document.createElement("span");
    hint.className = "enlarge-hint";
    hint.textContent = " Click the image to enlarge it.";
    prev.appendChild(hint);
  }
});

function localImageUrl(href) {
  try {
    const url = new URL(href, document.baseURI);
    const sameSite = url.protocol === "file:"
      ? location.protocol === "file:"
      : url.origin === location.origin;
    return sameSite && /\.(png|jpe?g|gif|webp)$/i.test(url.pathname) ? url.href : null;
  } catch (err) {
    return null;
  }
}

document.addEventListener("click", (event) => {
  const link = event.target.closest(".info-panel-body a.image-link");
  if (!link) return;
  event.preventDefault();

  const img = link.querySelector("img");
  const src = localImageUrl(link.getAttribute("href") || img.getAttribute("src"));
  if (!src) return;
  lightboxImg.src = src;
  lightboxImg.alt = img.alt;
  lightbox.showModal(); // darkens and disables the rest of the page
});

// Clicking anywhere except the image itself closes the overlay.
lightbox.addEventListener("click", (event) => {
  if (event.target !== lightboxImg) lightbox.close();
});

lightbox.addEventListener("close", () => {
  lightboxImg.removeAttribute("src");
});

// ---------- Init ----------
goToStep(1);
restoreSession();
