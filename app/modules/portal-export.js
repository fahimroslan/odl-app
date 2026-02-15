// app/modules/portal-export.js
import { getAllRecords, STORES } from "../db/db.js";

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

const CONFIRM_RESULT_KEY = "odlConfirmResultSlipsAt";
const CONFIRM_ENROLL_KEY = "odlConfirmEnrollmentSlipsAt";

function countMissing(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === "string" && value.trim() === "") return true;
  return false;
}

function buildPortalExport({ students, courses, results }) {
  let offeredCourseCodes = [];
  try {
    const stored = JSON.parse(localStorage.getItem("odlEnrollmentOfferedCourses") || "[]");
    if (Array.isArray(stored)) {
      offeredCourseCodes = stored.map((code) => String(code ?? "").trim()).filter(Boolean);
    }
  } catch (_) {
    offeredCourseCodes = [];
  }
  const missing = {
    studentName: 0,
    studentIc: 0,
    studentIntake: 0,
    resultStudentId: 0,
    resultCourseCode: 0,
    resultSession: 0,
    resultSemester: 0,
    resultMark: 0,
  };

  for (const student of students) {
    if (countMissing(student?.name)) missing.studentName += 1;
    if (countMissing(student?.ic)) missing.studentIc += 1;
    if (countMissing(student?.intake)) missing.studentIntake += 1;
  }

  for (const result of results) {
    if (countMissing(result?.studentId)) missing.resultStudentId += 1;
    if (countMissing(result?.courseCode)) missing.resultCourseCode += 1;
    if (countMissing(result?.session)) missing.resultSession += 1;
    if (countMissing(result?.semester)) missing.resultSemester += 1;
    if (countMissing(result?.mark)) missing.resultMark += 1;
  }

  const payload = {
    version: "odl-portal-export-v1",
    exportedAt: new Date().toISOString(),
    offeredCourseCodes,
    stats: {
      totalStudents: students.length,
      totalCourses: courses.length,
      totalResults: results.length,
      missing,
    },
    students,
    courses,
    results,
  };

  return payload;
}

export function initPortalExport() {
  const btnPortalExport = document.getElementById("btnPortalExport");
  const btnPortalUpload = document.getElementById("btnPortalUpload");
  const elPortalUploadUrl = document.getElementById("portalUploadUrl");
  const elPortalExportLog = document.getElementById("portalExportLog");

  if (!btnPortalUpload && !btnPortalExport) return;

  const buildPayload = async () => {
    const [students, courses, results] = await Promise.all([
      getAllRecords(STORES.students),
      getAllRecords(STORES.courses),
      getAllRecords(STORES.results),
    ]);
    return buildPortalExport({ students, courses, results });
  };

  if (btnPortalExport) {
    btnPortalExport.addEventListener("click", async () => {
      if (elPortalExportLog) elPortalExportLog.textContent = "";
      try {
        appendLog(elPortalExportLog, "Preparing export...");
        const payload = await buildPayload();
        const json = JSON.stringify(payload, null, 2);
        const blob = new Blob([json], { type: "application/json;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `odl-portal-export-${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);

        appendLog(elPortalExportLog, "Export generated.");
        appendLog(elPortalExportLog, `Students: ${payload.stats.totalStudents}`);
        appendLog(elPortalExportLog, `Courses: ${payload.stats.totalCourses}`);
        appendLog(elPortalExportLog, `Results: ${payload.stats.totalResults}`);
        appendLog(elPortalExportLog, "Missing data summary:");
        appendLog(elPortalExportLog, `- Missing student name: ${payload.stats.missing.studentName}`);
        appendLog(elPortalExportLog, `- Missing student IC: ${payload.stats.missing.studentIc}`);
        appendLog(elPortalExportLog, `- Missing student intake: ${payload.stats.missing.studentIntake}`);
        appendLog(elPortalExportLog, `- Missing result student ID: ${payload.stats.missing.resultStudentId}`);
        appendLog(elPortalExportLog, `- Missing result course code: ${payload.stats.missing.resultCourseCode}`);
        appendLog(elPortalExportLog, `- Missing result session: ${payload.stats.missing.resultSession}`);
        appendLog(elPortalExportLog, `- Missing result semester: ${payload.stats.missing.resultSemester}`);
        appendLog(elPortalExportLog, `- Missing result mark: ${payload.stats.missing.resultMark}`);
      } catch (e) {
        appendLog(elPortalExportLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnPortalUpload) {
    btnPortalUpload.addEventListener("click", async () => {
      if (elPortalExportLog) elPortalExportLog.textContent = "";
      try {
        let url = String(elPortalUploadUrl?.value ?? "").trim();
        if (!url) throw new Error("Portal upload URL is required.");
        if (!url.includes("/api/admin/import-portal-export")) {
          url = url.replace(/\/+$/, "");
          url = `${url}/api/admin/import-portal-export`;
          if (elPortalUploadUrl) elPortalUploadUrl.value = url;
        }
        const resultConfirmed = localStorage.getItem(CONFIRM_RESULT_KEY);
        const enrollmentConfirmed = localStorage.getItem(CONFIRM_ENROLL_KEY);
        if (!resultConfirmed || !enrollmentConfirmed) {
          throw new Error("Please confirm Result Slips and Enrollment Slips on the Students tab first.");
        }
        appendLog(elPortalExportLog, "Preparing export for upload...");
        btnPortalUpload.disabled = true;
        const originalLabel = btnPortalUpload.textContent;
        btnPortalUpload.textContent = "Uploading...";
        const payload = await buildPayload();
        const payloadSize = new Blob([JSON.stringify(payload)]).size;
        appendLog(elPortalExportLog, `Payload size: ${(payloadSize / (1024 * 1024)).toFixed(2)} MB`);
        appendLog(elPortalExportLog, `Uploading to ${url} ...`);
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 120000);
        const res = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
          signal: controller.signal,
        });
        clearTimeout(timeoutId);
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || `Upload failed (${res.status})`);
        }
        appendLog(elPortalExportLog, "Upload successful.");
        appendLog(elPortalExportLog, `Finished at ${new Date().toLocaleTimeString()}`);
        btnPortalUpload.textContent = originalLabel;
        btnPortalUpload.disabled = false;
      } catch (e) {
        const msg = e?.name === "AbortError" ? "Upload timed out (120s)." : e?.message ?? e;
        appendLog(elPortalExportLog, `ERROR: ${msg}`);
        if (btnPortalUpload) {
          btnPortalUpload.textContent = "Upload to Portal";
          btnPortalUpload.disabled = false;
        }
      }
    });
  }
}
