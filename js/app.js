// ============================================================
// Aeries to SkillsUSA Roster Converter
// ============================================================
//
// Security notes
// - Runs entirely offline. Every file it needs is in this project; the
//   Content-Security-Policy in index.html blocks all network connections.
// - Uploaded files are read in memory and never sent anywhere.
// - Text from uploaded files is only ever placed in the page with
//   textContent (never innerHTML), so it can't run as code.
// - harden.js freezes Object.prototype before SheetJS loads.
// - Only the parsed SkillsUSA template (program names and IDs) is kept in
//   sessionStorage. Student data is never stored.

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

// ---------- App state ----------
// Largest upload accepted. Real templates and rosters are far smaller;
// the cap stops an oversized or malicious file from freezing the tab.
const MAX_FILE_BYTES = 5 * 1024 * 1024; // 5 MB

// SheetJS read settings: values only, nothing the app doesn't use.
const READ_OPTIONS = {
  type: "array",
  cellFormula: false, // ignore formulas, read the stored values
  cellHTML: false,    // don't build HTML versions of cells
  cellStyles: false,
  bookVBA: false,     // never load macros
  sheetRows: 20000,   // stop reading past 20,000 rows per sheet
};

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

// Reads a SheetJS workbook and returns { sheetName, headers, programs }.
// Looks through every sheet for a header row containing "TrainingProgram".
function parseTemplate(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    // header: 1 gives an array of rows; raw: false keeps IDs as text
    // so values like leading zeros are not lost.
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });

    const headerRowIndex = rows.findIndex((row) =>
      row.some((cell) => normalizeHeader(cell) === "trainingprogram")
    );
    if (headerRowIndex === -1) continue;

    const headers = rows[headerRowIndex].map(cleanCell);

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
    for (const row of rows.slice(headerRowIndex + 1)) {
      const program = {
        chapterId: cleanCell(row[col.chapterId]),
        chapterName: cleanCell(row[col.chapterName]),
        sectionId: cleanCell(row[col.sectionId]),
        trainingProgram: cleanCell(row[col.trainingProgram]),
      };
      if (!program.trainingProgram) continue;

      const key = `${program.sectionId}|${program.trainingProgram}`;
      if (seen.has(key)) continue;
      seen.add(key);
      programs.push(program);
    }

    if (!programs.length) {
      throw new TemplateError(
        "This template has a TrainingProgram column, but no programs are listed under it."
      );
    }

    return { sheetName, headers, programs };
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

// "Doe, Jane" -> { lastName: "Doe", firstName: "Jane" }.
// Splits at the first comma. Returns null if the cell isn't in that format.
function splitStudentName(value) {
  const text = cleanCell(value);
  const comma = text.indexOf(",");
  if (comma === -1) return null;

  const lastName = text.slice(0, comma).trim();
  const firstName = text.slice(comma + 1).trim();
  if (!lastName || !firstName) return null;
  return { lastName, firstName };
}

// Reads a SheetJS workbook and returns
// { students, classCount, duplicateCount, missingEmailCount }.
function parseRoster(workbook) {
  const students = [];
  const seen = new Set();
  let classCount = 0;
  let duplicateCount = 0;
  let missingEmailCount = 0;

  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });

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
      "No students with an email address in the SEM column were found in this file."
    );
  }

  return { students, classCount, duplicateCount, missingEmailCount };
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

  // A Map (not a plain object) so a header such as "constructor" or
  // "__proto__" can't match a built-in object property.
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
  XLSX.utils.book_append_sheet(workbook, sheet, state.template.sheetName || "Sheet1");

  const fileName = `SkillsUSA_${safeFileName(state.selectedProgram.trainingProgram)}_${todayStamp()}.xlsx`;

  // Save through a temporary link so the page itself never navigates.
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
function populatePrograms(programs) {
  programSelect.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select a training program";
  programSelect.appendChild(placeholder);

  programs.forEach((p, index) => {
    const opt = document.createElement("option");
    opt.value = String(index);
    opt.textContent = p.trainingProgram;
    programSelect.appendChild(opt);
  });
}

function showStudentPreview(students) {
  studentTableBody.innerHTML = "";
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
        template: state.template,
        selectedIndex,
      })
    );
  } catch (err) {
    // Storage can be unavailable (private mode, file:// in some browsers).
  }
}

// Checks that restored data has exactly the expected shape before use.
function isValidSavedTemplate(saved) {
  const isString = (v) => typeof v === "string";
  const t = saved && saved.template;
  return (
    !!t &&
    isString(saved.templateFileName) &&
    Number.isInteger(saved.selectedIndex) &&
    isString(t.sheetName) &&
    Array.isArray(t.headers) && t.headers.every(isString) &&
    Array.isArray(t.programs) && t.programs.length > 0 &&
    t.programs.every((p) =>
      p && isString(p.chapterId) && isString(p.chapterName) &&
      isString(p.sectionId) && isString(p.trainingProgram)
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

    applyTemplate(saved.template, saved.templateFileName);
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
  programSelect.innerHTML = '<option value="">Upload a template first</option>';
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
  setMessage(templateMessage, `Found ${count} training program${count === 1 ? "" : "s"}.`, "success");
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

  if (typeof XLSX === "undefined") {
    setMessage(templateMessage, "The spreadsheet reader didn't load. Check that xlsx.full.min.js is in the js folder, then reload the page.", "error");
    templateInput.value = "";
    return;
  }

  setMessage(templateMessage, "Reading template…");

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, READ_OPTIONS);
    applyTemplate(parseTemplate(workbook), file.name);
  } catch (err) {
    console.error(err);
    const text = err instanceof TemplateError
      ? err.message
      : "This file couldn't be read as an Excel workbook. Try downloading the template again.";
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
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, READ_OPTIONS);
    const result = parseRoster(workbook);

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
    setMessage(rosterMessage, parts.join(" "), "success");
    goToStep(4);
  } catch (err) {
    console.error(err);
    const text = err instanceof RosterError
      ? err.message
      : "This file couldn't be read as an Excel workbook. Try exporting the roster from Aeries again.";
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
    const button = document.querySelector(`.info-btn[aria-controls="${panel.id}"]`);
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

document.addEventListener("click", (event) => {
  const link = event.target.closest(".info-panel-body a.image-link");
  if (!link) return;
  event.preventDefault();

  const img = link.querySelector("img");
  lightboxImg.src = link.getAttribute("href") || img.getAttribute("src");
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
