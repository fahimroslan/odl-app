// app/modules/backup.js
import { getAllRecords, STORES } from "../db/db.js";
import { importCSVFile } from "./results.js";

function escapeCsv(value) {
  const s = String(value ?? "");
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

export function buildBackupCsv({ students, courses, results }) {
  const studentMap = new Map(
    students.map((s) => [String(s.studentId ?? "").trim(), s])
  );
  const courseMap = new Map(
    courses.map((c) => [String(c.courseCode ?? "").trim(), c])
  );

  const header = [
    "ID",
    "Name",
    "IC",
    "Intake",
    "IntakeYear",
    "IntakeMonth",
    "CourseCode",
    "Title",
    "Credits",
    "IsMPUCourse",
    "Session",
    "Semester",
    "Mark",
    "Letter",
    "Point",
    "GradePoints",
    "PassedCourse",
    "CreditsEarned",
  ];

  const rows = results.map((result) => {
    const studentId = String(result.studentId ?? "").trim();
    const courseCode = String(result.courseCode ?? "").trim();
    const student = studentMap.get(studentId) ?? {};
    const course = courseMap.get(courseCode) ?? {};

    return [
      studentId,
      student.name ?? "",
      student.ic ?? "",
      student.intake ?? "",
      student.intakeYear ?? "",
      student.intakeMonth ?? "",
      courseCode,
      course.title ?? "",
      course.credits ?? "",
      course.isMPUCourse ?? "",
      result.session ?? "",
      result.semester ?? "",
      result.mark ?? "",
      result.letter ?? "",
      result.point ?? "",
      result.gradePoints ?? "",
      result.passedCourse ?? "",
      result.creditsEarned ?? "",
    ].map(escapeCsv);
  });

  return [header.map(escapeCsv).join(","), ...rows.map((r) => r.join(","))].join("\n");
}

export function initBackup({ onDataChanged } = {}) {
  const btnBackupDownload = document.getElementById("btnBackupDownload");
  const btnBackupImport = document.getElementById("btnBackupImport");
  const elBackupLog = document.getElementById("backupLog");
  const elBackupFile = document.getElementById("backupFile");

  if (btnBackupDownload) {
    btnBackupDownload.addEventListener("click", async () => {
      if (elBackupLog) elBackupLog.textContent = "";
      try {
        const [students, courses, results] = await Promise.all([
          getAllRecords(STORES.students),
          getAllRecords(STORES.courses),
          getAllRecords(STORES.results),
        ]);
        const csv = buildBackupCsv({ students, courses, results });
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "odl-backup.csv";
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        appendLog(elBackupLog, "Backup downloaded.");
      } catch (e) {
        appendLog(elBackupLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnBackupImport) {
    btnBackupImport.addEventListener("click", async () => {
      if (elBackupLog) elBackupLog.textContent = "";
      try {
        const file = elBackupFile?.files?.[0];
        await importCSVFile(file, (msg) => appendLog(elBackupLog, msg));
        if (onDataChanged) await onDataChanged();
      } catch (e) {
        appendLog(elBackupLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }
}
