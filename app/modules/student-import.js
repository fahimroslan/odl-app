// app/modules/student-import.js
import { getAllRecords, upsertMany, STORES } from "../db/db.js";

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

function normalizeKey(key) {
  return String(key ?? "")
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

function getRowValue(row, keys) {
  for (const key of keys) {
    if (key in row) return row[key];
  }
  return "";
}

function normalizeStudentRow(rawRow, intake, intakeYear, intakeMonth) {
  const row = {};
  for (const key of Object.keys(rawRow)) {
    row[normalizeKey(key)] = rawRow[key];
  }

  const studentId = String(
    getRowValue(row, ["studentid", "id", "studentno", "matric", "matricno", "student_id"])
  ).trim();
  const name = String(getRowValue(row, ["name", "fullname", "studentname"])).trim();
  const ic = String(getRowValue(row, ["ic", "icno", "nric", "nricno"])).trim();

  if (!studentId || !name) return null;

  return {
    studentId,
    name,
    ic,
    intake,
    intakeYear,
    intakeMonth,
  };
}

function validateIntake(year, month) {
  if (!year || !month) throw new Error("Intake year and month are required.");
  if (!/^\d{4}$/.test(year)) throw new Error("Intake year must be YYYY.");
  if (!/^(02|09)$/.test(month)) {
    throw new Error("Intake month must be 02 or 09.");
  }
  return `${year}-${month}`;
}

export function initStudentImport({ onDataChanged } = {}) {
  const elYear = document.getElementById("studentIntakeYear");
  const elMonth = document.getElementById("studentIntakeMonth");
  const elFile = document.getElementById("studentCsvFile");
  const btnImport = document.getElementById("btnImportStudents");
  const elLog = document.getElementById("studentImportLog");

  if (!btnImport) return;

  btnImport.addEventListener("click", async () => {
    if (elLog) elLog.textContent = "";
    try {
      const year = String(elYear?.value ?? "").trim();
      const month = String(elMonth?.value ?? "").trim();
      const intake = validateIntake(year, month);
      const file = elFile?.files?.[0];
      if (!file) throw new Error("No CSV file selected.");

      appendLog(elLog, `Parsing: ${file.name} ...`);

      const parsed = await new Promise((resolve, reject) => {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          dynamicTyping: false,
          complete: (res) => resolve(res),
          error: (err) => reject(err),
        });
      });

      if (parsed.errors?.length) {
        appendLog(elLog, "CSV parse warnings/errors:");
        for (const e of parsed.errors.slice(0, 10)) appendLog(elLog, `- ${e.message}`);
      }

      const existing = await getAllRecords(STORES.students);
      const existingIds = new Set(
        existing.map((s) => String(s.studentId ?? "").trim()).filter(Boolean)
      );
      const seenIds = new Set();
      const newStudents = [];
      const duplicateIds = [];
      let missingRequired = 0;

      for (const row of parsed.data) {
        const normalized = normalizeStudentRow(row, intake, year, month);
        if (!normalized) {
          missingRequired += 1;
          continue;
        }
        const id = String(normalized.studentId).trim();
        if (seenIds.has(id) || existingIds.has(id)) {
          duplicateIds.push(id);
          continue;
        }
        seenIds.add(id);
        newStudents.push(normalized);
      }

      appendLog(elLog, `Intake set: ${intake}`);
      appendLog(elLog, `Rows read: ${parsed.data.length}`);
      appendLog(elLog, `New students: ${newStudents.length}`);
      appendLog(elLog, `Skipped missing required fields: ${missingRequired}`);
      appendLog(elLog, `Duplicates skipped: ${duplicateIds.length}`);
      if (duplicateIds.length) {
        appendLog(elLog, "Duplicate Student IDs (first 10):");
        duplicateIds.slice(0, 10).forEach((id) => appendLog(elLog, `- ${id}`));
      }

      if (newStudents.length) {
        await upsertMany(STORES.students, newStudents);
        appendLog(elLog, "Student import complete.");
        if (onDataChanged) await onDataChanged();
      } else {
        appendLog(elLog, "No new students to import.");
      }
    } catch (e) {
      appendLog(elLog, `ERROR: ${e.message ?? e}`);
    }
  });
}
