// app/modules/imports-ui.js
import { importCSVFile, previewNewResults, uploadNewResults } from "./results.js";
import { getAllRecords, getCounts, STORES } from "../db/db.js";

function appendLog(el, msg) {
  if (!el) return;
  el.textContent += `${msg}\n`;
}

export function initImports({ onDataChanged }) {
  const elFileImports = document.getElementById("csvFileImports");
  const btnImportImports = document.getElementById("btnImportImports");
  const elLogImports = document.getElementById("logImports");

  const elNewResultYear = document.getElementById("newResultYear");
  const elNewResultMonth = document.getElementById("newResultMonth");
  const elNewResultCourse = document.getElementById("newResultCourse");
  const elNewResultFile = document.getElementById("newResultFile");
  const btnUploadNewResults = document.getElementById("btnUploadNewResults");
  const elNewResultLog = document.getElementById("newResultLog");
  const btnPreviewResults = document.getElementById("btnPreviewResults");
  const elPreviewModal = document.getElementById("resultPreviewModal");
  const elPreviewBody = document.getElementById("resultPreviewBody");
  const elPreviewTable = document.getElementById("resultPreviewTable");
  const elPreviewEmpty = document.getElementById("resultPreviewEmpty");
  const elPreviewCount = document.getElementById("resultPreviewCount");
  const btnSelectAllPreview = document.getElementById("btnSelectAllPreview");
  const btnSelectNonePreview = document.getElementById("btnSelectNonePreview");
  const btnConfirmPreview = document.getElementById("btnConfirmPreview");
  const btnClosePreview = document.getElementById("btnClosePreview");
  let previewCache = null;

  const loadCourses = async () => {
    if (!elNewResultCourse) return;
    const courses = await getAllRecords(STORES.courses);
    const sorted = courses
      .filter((course) => course?.courseCode)
      .sort((a, b) => String(a.courseCode).localeCompare(String(b.courseCode)));
    elNewResultCourse.textContent = "";
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "Course";
    elNewResultCourse.appendChild(defaultOption);
    for (const course of sorted) {
      const option = document.createElement("option");
      option.value = String(course.courseCode ?? "");
      option.textContent = course.title
        ? `${course.courseCode} - ${course.title}`
        : String(course.courseCode ?? "");
      elNewResultCourse.appendChild(option);
    }
  };

  if (btnImportImports) {
    btnImportImports.addEventListener("click", async () => {
      if (elLogImports) elLogImports.textContent = "";
      try {
        const file = elFileImports?.files?.[0];
        await importCSVFile(file, (msg) => appendLog(elLogImports, msg));
        const counts = await getCounts();
        appendLog(
          elLogImports,
          `DB counts -> students: ${counts.students}, courses: ${counts.courses}, results: ${counts.results}`
        );
        if (onDataChanged) await onDataChanged();
      } catch (e) {
        appendLog(elLogImports, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  const openPreview = async () => {
      if (elNewResultLog) elNewResultLog.textContent = "";
      try {
        const year = String(elNewResultYear?.value ?? "").trim();
        const month = String(elNewResultMonth?.value ?? "").trim();
        const courseCode = String(elNewResultCourse?.value ?? "").trim();
        const session = year && month ? `${year}/${month}` : "";
        const file = elNewResultFile?.files?.[0];

        const preview = await previewNewResults({ file, session, courseCode });
        previewCache = { preview, session, courseCode, file };

        if (elPreviewCount) {
          elPreviewCount.textContent = String(preview.affectedStudents.length);
        }
        if (preview.missingNameNewStudents) {
          appendLog(
            elNewResultLog,
            `WARNING: ${preview.missingNameNewStudents} new student(s) have no name in this file. Upload a student list CSV to fill names.`
          );
        }
        if (preview.skippedMismatch) {
          appendLog(
            elNewResultLog,
            `WARNING: ${preview.skippedMismatch} row(s) were blocked because ID+Name did not match existing student records.`
          );
          const mismatches = preview.mismatchedExisting ?? [];
          mismatches.slice(0, 10).forEach((row) => {
            appendLog(
              elNewResultLog,
              `- Row ${row.rowNo}: ${row.studentId} | existing="${row.existingName}" vs uploaded="${row.incomingName}"`
            );
          });
        }

        if (!preview.affectedStudents.length) {
          elPreviewBody.textContent = "";
          elPreviewEmpty.style.display = "block";
          if (elPreviewTable) elPreviewTable.style.display = "none";
        } else {
          elPreviewEmpty.style.display = "none";
          if (elPreviewTable) elPreviewTable.style.display = "table";
          elPreviewBody.textContent = "";
          for (const student of preview.affectedStudents) {
            const tr = document.createElement("tr");
            const tdCheck = document.createElement("td");
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.checked = true;
            checkbox.dataset.studentId = student.studentId;
            tdCheck.appendChild(checkbox);
            tr.appendChild(tdCheck);

            const tdId = document.createElement("td");
            tdId.textContent = student.studentId;
            tr.appendChild(tdId);

            const tdName = document.createElement("td");
            tdName.textContent = student.name || "-";
            tr.appendChild(tdName);

            const tdNew = document.createElement("td");
            tdNew.textContent = student.isNew ? "New" : "Existing";
            tr.appendChild(tdNew);

            const tdCount = document.createElement("td");
            tdCount.textContent = String(student.count);
            tr.appendChild(tdCount);

            elPreviewBody.appendChild(tr);
          }
        }

        if (elPreviewModal) {
          elPreviewModal.classList.add("active");
          elPreviewModal.setAttribute("aria-hidden", "false");
        }
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
      }
  };

  if (btnPreviewResults) {
    btnPreviewResults.addEventListener("click", async () => {
      await openPreview();
    });
  }

  if (btnUploadNewResults) {
    btnUploadNewResults.addEventListener("click", async () => {
      if (elNewResultLog) elNewResultLog.textContent = "";
      try {
        const year = String(elNewResultYear?.value ?? "").trim();
        const month = String(elNewResultMonth?.value ?? "").trim();
        const courseCode = String(elNewResultCourse?.value ?? "").trim();
        const session = year && month ? `${year}/${month}` : "";
        const file = elNewResultFile?.files?.[0];
        if (previewCache?.session !== session || previewCache?.courseCode !== courseCode || previewCache?.file !== file) {
          previewCache = null;
        }
        if (!previewCache) {
          await openPreview();
          return;
        }
        appendLog(elNewResultLog, "Use Confirm Update in the preview window.");
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnSelectAllPreview) {
    btnSelectAllPreview.addEventListener("click", () => {
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        cb.checked = true;
      });
    });
  }

  if (btnSelectNonePreview) {
    btnSelectNonePreview.addEventListener("click", () => {
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        cb.checked = false;
      });
    });
  }

  if (btnConfirmPreview) {
    btnConfirmPreview.addEventListener("click", async () => {
      if (!previewCache) return;
      const selected = [];
      elPreviewBody?.querySelectorAll("input[type='checkbox']").forEach((cb) => {
        if (cb.checked) selected.push(cb.dataset.studentId);
      });
      try {
        if (!selected.length) {
          appendLog(elNewResultLog, "No students selected.");
          return;
        }
        const result = await uploadNewResults({
          session: previewCache.session,
          file: previewCache.file,
          courseCode: previewCache.courseCode,
          selectedStudentIds: selected,
          logFn: (msg) => appendLog(elNewResultLog, msg),
        });
        if (onDataChanged) await onDataChanged();
        previewCache = null;
        if (elPreviewModal) {
          elPreviewModal.classList.remove("active");
          elPreviewModal.setAttribute("aria-hidden", "true");
        }
      } catch (e) {
        appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`);
      }
    });
  }

  if (btnClosePreview) {
    btnClosePreview.addEventListener("click", () => {
      if (elPreviewModal) {
        elPreviewModal.classList.remove("active");
        elPreviewModal.setAttribute("aria-hidden", "true");
      }
    });
  }

  if (elPreviewModal) {
    elPreviewModal.addEventListener("click", (event) => {
      if (event.target === elPreviewModal) {
        elPreviewModal.classList.remove("active");
        elPreviewModal.setAttribute("aria-hidden", "true");
      }
    });
  }

  loadCourses().catch((e) => appendLog(elNewResultLog, `ERROR: ${e.message ?? e}`));
}
