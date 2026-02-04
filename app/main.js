// app/main.js
import { deleteCourse, listCourses, normalizeCourseCode, upsertCourse } from "./modules/course.js";
import { initImports } from "./modules/imports-ui.js";
import { computeStudentCgpa } from "./modules/stats.js";
import { getGradeForMark } from "./modules/results.js";
import { initBackup } from "./modules/backup.js";
import { getAllRecords, STORES, upsertMany } from "./db/db.js";

const btnRefresh = document.getElementById("btnRefresh");

const elTotalStudents = document.getElementById("totalStudents");
const elTotalCourses = document.getElementById("totalCourses");
const elActiveCourses = document.getElementById("activeCourses");
const elTotalResults = document.getElementById("totalResults");
const elActiveStudents = document.getElementById("activeStudents");
const elStudentsWithFailures = document.getElementById("studentsWithFailures");
const elStudentsIncomplete = document.getElementById("studentsIncomplete");
const elStudentsLowCgpa = document.getElementById("studentsLowCgpa");
const elStudentsResultIssues = document.getElementById("studentsResultIssues");
const elAvgResultsPerStudent = document.getElementById("avgResultsPerStudent");
const btnViewFailures = document.getElementById("btnViewFailures");
const btnViewIncomplete = document.getElementById("btnViewIncomplete");
const btnViewLowCgpa = document.getElementById("btnViewLowCgpa");
const btnViewResultIssues = document.getElementById("btnViewResultIssues");
const elIssueTypeSelect = document.getElementById("issueTypeSelect");
const elIssueSearch = document.getElementById("issueSearch");
const elIssueListTitle = document.getElementById("issueListTitle");
const elIssueListCount = document.getElementById("issueListCount");
const elIssueListEmpty = document.getElementById("issueListEmpty");
const elIssueListTable = document.getElementById("issueListTable");
const elIssueListBody = document.getElementById("issueListBody");
const elIssuePager = document.getElementById("issuePager");
const btnIssuePrev = document.getElementById("btnIssuePrev");
const btnIssueNext = document.getElementById("btnIssueNext");
const elIssuePageMeta = document.getElementById("issuePageMeta");
const elLastUpdated = document.getElementById("lastUpdated");
const elCourseCoverageModal = document.getElementById("courseCoverageModal");
const btnCloseCoverage = document.getElementById("btnCloseCoverage");
const btnExportCoverage = document.getElementById("btnExportCoverage");
const elCoverageCourseMeta = document.getElementById("coverageCourseMeta");
const elCoverageTakenCount = document.getElementById("coverageTakenCount");
const elCoverageNotTakenCount = document.getElementById("coverageNotTakenCount");
const elCoverageTakenList = document.getElementById("coverageTakenList");
const elCoverageNotTakenList = document.getElementById("coverageNotTakenList");
const elCourseForm = document.getElementById("courseForm");
const elCourseCode = document.getElementById("courseCode");
const elCourseTitle = document.getElementById("courseTitle");
const elCourseCredits = document.getElementById("courseCredits");
const elCourseMPU = document.getElementById("courseMPU");
const btnCourseReset = document.getElementById("btnCourseReset");
const elCourseFormHint = document.getElementById("courseFormHint");
const elCourseListBody = document.getElementById("courseListBody");
const elCourseListTable = document.getElementById("courseListTable");
const elCourseListEmpty = document.getElementById("courseListEmpty");
const elCourseTotalCount = document.getElementById("courseTotalCount");
const elCourseTotalCredits = document.getElementById("courseTotalCredits");
const tabButtons = document.querySelectorAll("[data-tab]");
const tabPanels = document.querySelectorAll(".tab-panel");
const resultSubtabButtons = document.querySelectorAll(".result-subtab-btn");
const resultSubtabPanels = document.querySelectorAll(".result-subtab-panel");
const elResultTotal = document.getElementById("resultTotal");
const elResultPassRate = document.getElementById("resultPassRate");
const elResultAvgMark = document.getElementById("resultAvgMark");
const elResultStudents = document.getElementById("resultStudents");
const elResultSessionSelect = document.getElementById("resultSessionSelect");
const elResultSearch = document.getElementById("resultSearch");
const elResultViewMode = document.getElementById("resultViewMode");
const elResultTable = document.getElementById("resultTable");
const elResultTableBody = document.getElementById("resultTableBody");
const elResultTableEmpty = document.getElementById("resultTableEmpty");
const elCourseDetailStats = document.getElementById("courseDetailStats");
const elCourseDetailRecords = document.getElementById("courseDetailRecords");
const elCourseDetailPassRate = document.getElementById("courseDetailPassRate");
const elCourseDetailAvgMark = document.getElementById("courseDetailAvgMark");
const elCourseDetailStudents = document.getElementById("courseDetailStudents");
const elCourseDetailHeader = document.getElementById("courseDetailHeader");
const elCourseDetailTitle = document.getElementById("courseDetailTitle");
const btnBackToCourseList = document.getElementById("btnBackToCourseList");
const elStudentEditModal = document.getElementById("studentEditModal");
const btnCloseStudentEdit = document.getElementById("btnCloseStudentEdit");
const btnSaveStudentEdit = document.getElementById("btnSaveStudentEdit");
const elEditStudentId = document.getElementById("editStudentId");
const elEditStudentName = document.getElementById("editStudentName");
const elEditStudentIc = document.getElementById("editStudentIc");
const elEditStudentIntake = document.getElementById("editStudentIntake");
const btnPrintSlip = document.getElementById("btnPrintSlip");
const elResultSlipModal = document.getElementById("resultSlipModal");
const btnCloseSlip = document.getElementById("btnCloseSlip");
const elSlipStudentMeta = document.getElementById("slipStudentMeta");
const elSlipCGPA = document.getElementById("slipCGPA");
const elSlipTerms = document.getElementById("slipTerms");
const elResultPager = document.getElementById("resultPager");
const btnPrevPage = document.getElementById("btnPrevPage");
const btnNextPage = document.getElementById("btnNextPage");
const elResultPageMeta = document.getElementById("resultPageMeta");
const elMarkEditModal = document.getElementById("markEditModal");
const btnCloseMarkEdit = document.getElementById("btnCloseMarkEdit");
const btnConfirmMarkEdit = document.getElementById("btnConfirmMarkEdit");
const elMarkEditStudent = document.getElementById("markEditStudent");
const elMarkEditCourse = document.getElementById("markEditCourse");
const elMarkEditSession = document.getElementById("markEditSession");
const elMarkEditSemester = document.getElementById("markEditSemester");
const elMarkEditCurrent = document.getElementById("markEditCurrent");
const elMarkEditInput = document.getElementById("markEditInput");
const elMarkEditLetter = document.getElementById("markEditLetter");
const elMarkEditPoint = document.getElementById("markEditPoint");
const elMarkEditStatus = document.getElementById("markEditStatus");
const elMarkEditWarning = document.getElementById("markEditWarning");
const elMarkEditVerify = document.getElementById("markEditVerify");
const elEnrollmentCourseGrid = document.getElementById("enrollmentCourseGrid");
const btnEnrollmentSelectAll = document.getElementById("btnEnrollmentSelectAll");
const btnEnrollmentClearAll = document.getElementById("btnEnrollmentClearAll");
const btnEnrollmentBuild = document.getElementById("btnEnrollmentBuild");
const elEnrollmentSelectedCount = document.getElementById("enrollmentSelectedCount");
const elEnrollmentSearch = document.getElementById("enrollmentSearch");
const elEnrollmentStudentCount = document.getElementById("enrollmentStudentCount");
const elEnrollmentListEmpty = document.getElementById("enrollmentListEmpty");
const elEnrollmentListTable = document.getElementById("enrollmentListTable");
const elEnrollmentListBody = document.getElementById("enrollmentListBody");
const elEnrollmentSlipModal = document.getElementById("enrollmentSlipModal");
const btnCloseEnrollmentSlip = document.getElementById("btnCloseEnrollmentSlip");
const elEnrollmentSlipCount = document.getElementById("enrollmentSlipCount");
const elEnrollmentSlipStudentMeta = document.getElementById("enrollmentSlipStudentMeta");
const elEnrollmentSlipBody = document.getElementById("enrollmentSlipBody");
const numberFmt = new Intl.NumberFormat();
const percentFmt = new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: 1 });
const resultState = {
  students: [],
  courses: [],
  results: [],
};
const resultPaging = {
  page: 1,
  pageSize: 20,
};
let activeCourseDetailCode = "";
let activeStudentEditId = "";
let activeMarkEditId = "";
let activeMarkEditValue = "";
let activeMarkEditRow = null;
let activeEnrollmentStudentId = "";
const resultSortState = {
  key: "",
  dir: "asc",
};
const courseCoverageDetails = new Map();
let activeCoverageCourse = "";
let activeCourseEditCode = "";
const studentIssueLists = {
  failures: [],
  incomplete: [],
  lowCgpa: [],
  resultIssues: [],
};

const issueTitles = {
  failures: "Students with failures",
  incomplete: "Incomplete student info",
  lowCgpa: "CGPA below 2.00",
  resultIssues: "Result issues",
};
const issuePaging = {
  page: 1,
  pageSize: 20,
};
const enrollmentState = {
  slips: [],
  offeredCourseCodes: [],
};
const ENROLLMENT_MAX_OFFERED = 15;
const ENROLLMENT_TARGET_CREDITS = 90;

function setActiveResultSubtab(subtabId) {
  resultSubtabButtons.forEach((button) => {
    const isActive = button.dataset.subtab === subtabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  resultSubtabPanels.forEach((panel) => {
    const isActive = panel.id === `result-subtab-${subtabId}`;
    panel.classList.toggle("active", isActive);
  });
}


function setText(el, value) {
  if (!el) return;
  el.textContent = value;
}

function setLastUpdated() {
  if (!elLastUpdated) return;
  const now = new Date();
  elLastUpdated.textContent = `Last updated: ${now.toLocaleString()}`;
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function isFailedResult(result) {
  if (result.passedCourse === false) return true;
  const letter = String(result.letter ?? "").trim().toUpperCase();
  if (letter === "F" || letter === "FAIL") return true;
  const mark = toNumber(result.mark);
  if (mark !== null && mark < 40) return true;
  return false;
}

function isMissingMark(result) {
  if (result.mark === null || result.mark === undefined) return true;
  const s = String(result.mark).trim();
  return s === "";
}

function getGradePoints(result, courseCredits) {
  const direct = toNumber(result.gradePoints);
  const point = toNumber(result.point);
  const credits = toNumber(courseCredits);
  if (direct !== null) {
    if (direct === 0 && point !== null && point > 0 && credits !== null && credits > 0) {
      return point * credits;
    }
    return direct;
  }
  if (point !== null && credits !== null) return point * credits;
  return null;
}

function buildStudentMap() {
  const map = new Map();
  for (const student of resultState.students) {
    const studentId = String(student.studentId ?? "").trim();
    if (!studentId) continue;
    map.set(studentId, student);
  }
  return map;
}

function buildCourseMap() {
  const map = new Map();
  for (const course of resultState.courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    map.set(code, course);
  }
  return map;
}

function renderResultStats(results, studentIds) {
  const total = results.length;
  const passed = results.filter((row) => row.passedCourse === true).length;
  const failed = results.filter((row) => isFailedResult(row)).length;
  const missingMark = results.filter((row) => isMissingMark(row)).length;
  const passRate = total ? passed / total : 0;
  const marks = results.map((row) => toNumber(row.mark)).filter((mark) => mark !== null);
  const avgMark = marks.length ? marks.reduce((a, b) => a + b, 0) / marks.length : null;

  setText(elResultTotal, numberFmt.format(total));
  setText(elResultPassRate, total ? percentFmt.format(passRate) : "--");
  if (elResultPassRate) {
    elResultPassRate.title = `Failed: ${numberFmt.format(failed)} | No mark: ${numberFmt.format(missingMark)}`;
  }
  setText(elResultAvgMark, avgMark === null ? "--" : avgMark.toFixed(1));
  setText(elResultStudents, numberFmt.format(studentIds.size));
}

function renderResultTableHeader(mode) {
  const thead = elResultTable.querySelector("thead");
  if (!thead) return;
  if (mode === "students") {
    thead.innerHTML = `
      <tr>
        <th class="sortable" data-sort="studentId">Student ID</th>
        <th class="sortable" data-sort="studentName">Student Name</th>
        <th class="sortable" data-sort="currentSemester">Current Semester</th>
        <th class="sortable" data-sort="cgpa">CGPA</th>
        <th class="sortable" data-sort="records">Records</th>
        <th>Action</th>
      </tr>
    `;
    return;
  }

  if (mode === "courses") {
    thead.innerHTML = `
      <tr>
        <th class="sortable" data-sort="courseLabel">Course</th>
        <th class="sortable" data-sort="records">Records</th>
        <th class="sortable" data-sort="passRate">Pass Rate</th>
        <th class="sortable" data-sort="avgMark">Avg Mark</th>
        <th>Action</th>
      </tr>
    `;
    return;
  }

  const actionHeader = mode === "detail" ? "<th>Action</th>" : "";
  thead.innerHTML = `
    <tr>
      <th class="sortable" data-sort="studentId">Student ID</th>
      <th class="sortable" data-sort="studentName">Student Name</th>
      <th class="sortable" data-sort="session">Session</th>
      <th class="sortable" data-sort="semester">Semester</th>
      <th class="sortable" data-sort="mark">Mark</th>
      <th class="sortable" data-sort="letter">Letter</th>
      <th class="sortable" data-sort="point">Point</th>
      <th class="sortable" data-sort="status">Status</th>
      ${actionHeader}
    </tr>
  `;
}

function renderResultTable(rows, mode) {
  if (!elResultTable || !elResultTableBody || !elResultTableEmpty || !elResultPager) return;
  if (!rows.length) {
    elResultTable.style.display = "none";
    elResultTableEmpty.style.display = "block";
    elResultPager.style.display = "none";
    return;
  }

  elResultTableEmpty.style.display = "none";
  elResultTable.style.display = "table";
  elResultTableBody.textContent = "";
  renderResultTableHeader(mode);
  const thead = elResultTable.querySelector("thead");
  if (thead) {
    thead.querySelectorAll("th.sortable").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.dataset.sort || "";
        if (!key) return;
        if (resultSortState.key === key) {
          resultSortState.dir = resultSortState.dir === "asc" ? "desc" : "asc";
        } else {
          resultSortState.key = key;
          resultSortState.dir = "asc";
        }
        renderResultsView();
      });
    });
  }


  if (mode === "students") {
    elResultPager.style.display = "flex";
    for (const row of rows) {
      const tr = document.createElement("tr");

      const tdStudentId = document.createElement("td");
      tdStudentId.textContent = row.studentId ?? "-";
      tr.appendChild(tdStudentId);

      const tdStudentName = document.createElement("td");
      if (row.studentId) {
        const btn = document.createElement("button");
        btn.className = "link-button";
        btn.type = "button";
        btn.textContent = row.studentName || "-";
        btn.addEventListener("click", () => openStudentEdit(row.studentId));
        tdStudentName.appendChild(btn);
      } else {
        tdStudentName.textContent = row.studentName || "-";
      }
      tr.appendChild(tdStudentName);

      const tdSemester = document.createElement("td");
      tdSemester.textContent = row.currentSemester ?? "-";
      tr.appendChild(tdSemester);

      const tdCgpa = document.createElement("td");
      tdCgpa.textContent = row.cgpa ?? "--";
      tr.appendChild(tdCgpa);

      const tdCount = document.createElement("td");
      tdCount.textContent = numberFmt.format(row.records);
      tr.appendChild(tdCount);

      const tdAction = document.createElement("td");
      const btnView = document.createElement("button");
      btnView.className = "btn secondary small";
      btnView.type = "button";
      btnView.textContent = "View results";
      btnView.addEventListener("click", () => {
        openSlip(row.studentId);
      });
      tdAction.appendChild(btnView);
      tr.appendChild(tdAction);

      elResultTableBody.appendChild(tr);
    }
    return;
  }

  if (mode === "courses") {
    elResultPager.style.display = "flex";
    for (const row of rows) {
      const tr = document.createElement("tr");

      const tdCourse = document.createElement("td");
      tdCourse.textContent = row.courseLabel;
      tr.appendChild(tdCourse);

      const tdCount = document.createElement("td");
      tdCount.textContent = numberFmt.format(row.records);
      tr.appendChild(tdCount);

      const tdPass = document.createElement("td");
      tdPass.textContent = row.passRate === null ? "--" : percentFmt.format(row.passRate);
      tr.appendChild(tdPass);

      const tdAvg = document.createElement("td");
      tdAvg.textContent = row.avgMark === null ? "--" : row.avgMark.toFixed(1);
      tr.appendChild(tdAvg);

      const tdAction = document.createElement("td");
      const btnView = document.createElement("button");
      btnView.className = "btn secondary small";
      btnView.type = "button";
      btnView.textContent = "View results";
      btnView.addEventListener("click", () => {
        setActiveTab("results");
        renderCourseResults(row.courseCode);
      });
      tdAction.appendChild(btnView);
      tr.appendChild(tdAction);

      elResultTableBody.appendChild(tr);
    }
    return;
  }

  elResultPager.style.display = "none";
  for (const row of rows) {
    const tr = document.createElement("tr");

    const tdStudentId = document.createElement("td");
    tdStudentId.textContent = row.studentId ?? "-";
    tr.appendChild(tdStudentId);

    const tdStudentName = document.createElement("td");
    if (row.studentId) {
      const btn = document.createElement("button");
      btn.className = "link-button";
      btn.type = "button";
      btn.textContent = row.studentName || "-";
      btn.addEventListener("click", () => openStudentEdit(row.studentId));
      tdStudentName.appendChild(btn);
    } else {
      tdStudentName.textContent = row.studentName || "-";
    }
    tr.appendChild(tdStudentName);

    const tdSession = document.createElement("td");
    tdSession.textContent = row.session || "-";
    tr.appendChild(tdSession);

    const tdSemester = document.createElement("td");
    tdSemester.textContent = row.semester ?? "-";
    tr.appendChild(tdSemester);

    const tdMark = document.createElement("td");
    tdMark.textContent = row.mark ?? "-";
    if (row.failed) tdMark.classList.add("fail");
    tr.appendChild(tdMark);

    const tdLetter = document.createElement("td");
    tdLetter.textContent = row.letter || "-";
    if (row.failed) tdLetter.classList.add("fail");
    tr.appendChild(tdLetter);

    const tdPoint = document.createElement("td");
    tdPoint.textContent = row.point ?? "-";
    tr.appendChild(tdPoint);

    const tdStatus = document.createElement("td");
    if (row.failed) {
      tdStatus.textContent = "Failed";
      tdStatus.classList.add("remark");
    } else {
      tdStatus.textContent = row.passedCourse === true ? "Passed" : "-";
    }
    tr.appendChild(tdStatus);

    if (mode === "detail") {
      const tdAction = document.createElement("td");
      tdAction.className = "table-actions";
      const btnEdit = document.createElement("button");
      btnEdit.className = "btn secondary small";
      btnEdit.type = "button";
      btnEdit.textContent = "Edit";
      btnEdit.addEventListener("click", () => openMarkEdit(row));
      tdAction.appendChild(btnEdit);
      tr.appendChild(tdAction);
    }

    elResultTableBody.appendChild(tr);
  }
}

function buildSlip(studentId) {
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();
  const student = studentMap.get(studentId);

  if (!student) {
    elSlipStudentMeta.textContent = "Select a student to generate a slip.";
    elSlipCGPA.textContent = "--";
    elSlipTerms.textContent = "";
    return;
  }

  const studentResults = resultState.results.filter(
    (row) => String(row.studentId ?? "").trim() === studentId
  );

  if (!studentResults.length) {
    elSlipStudentMeta.textContent = "No results found for this student.";
    elSlipCGPA.textContent = "--";
    elSlipTerms.textContent = "";
    return;
  }

  elSlipStudentMeta.textContent = "";
  const table = document.createElement("table");
  const tbody = document.createElement("tbody");

  const makeCell = (label, value) => {
    const td = document.createElement("td");
    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    td.appendChild(labelSpan);
    td.appendChild(document.createElement("br"));
    td.appendChild(document.createTextNode(String(value)));
    return td;
  };

  const row1 = document.createElement("tr");
  row1.appendChild(makeCell("Student Name", student.name ?? "-"));
  row1.appendChild(makeCell("Student ID", student.studentId ?? studentId));
  tbody.appendChild(row1);

  const row2 = document.createElement("tr");
  row2.appendChild(makeCell("IC", student.ic ?? "-"));
  row2.appendChild(makeCell("Intake", student.intake ?? "-"));
  tbody.appendChild(row2);

  table.appendChild(tbody);
  elSlipStudentMeta.appendChild(table);

  const termGroups = new Map();
  for (const row of studentResults) {
    const session = String(row.session ?? "").trim();
    const semester = row.semester ?? "";
    const key = `${session}||${semester}`;
    if (!termGroups.has(key)) {
      termGroups.set(key, { session, semester, rows: [] });
    }
    termGroups.get(key).rows.push(row);
  }

  const termList = [...termGroups.values()].sort((a, b) => {
    if (a.session === b.session) {
      return String(a.semester ?? "").localeCompare(String(b.semester ?? ""));
    }
    return a.session.localeCompare(b.session);
  });

  let totalCredits = 0;
  let totalGradePoints = 0;

  elSlipTerms.textContent = "";

  const failedBuckets = [];

  for (const term of termList) {
    const section = document.createElement("div");
    section.className = "slip-section";

    const heading = document.createElement("h4");
    heading.textContent = `Session ${term.session || "-"} - Semester ${term.semester ?? "-"}`;
    section.appendChild(heading);

    const table = document.createElement("table");
    table.className = "slip-table";
    const colgroup = document.createElement("colgroup");
    const colWidths = ["14%", "32%", "8%", "8%", "8%", "8%", "22%"];
    for (const width of colWidths) {
      const col = document.createElement("col");
      col.style.width = width;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Course</th>
        <th>Title</th>
        <th>Credits</th>
        <th>Mark</th>
        <th>Letter</th>
        <th>Point</th>
        <th>Grade Pts</th>
      </tr>
    `;
    table.appendChild(thead);
    const tbody = document.createElement("tbody");

    let termCredits = 0;
    let termCreditsEarned = 0;
    let termGradePoints = 0;

    for (const row of term.rows) {
      const tr = document.createElement("tr");
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      const course = courseMap.get(courseCode);
      const credits = toNumber(course?.credits);
      const gradePoints = getGradePoints(row, credits);
      const failed = isFailedResult(row);

      const tdCourse = document.createElement("td");
      tdCourse.textContent = courseCode || "-";
      tr.appendChild(tdCourse);

      const tdTitle = document.createElement("td");
      tdTitle.textContent = course?.title || "-";
      tr.appendChild(tdTitle);

      const tdCredits = document.createElement("td");
      tdCredits.textContent = credits === null ? "-" : numberFmt.format(credits);
      tr.appendChild(tdCredits);

      const tdMark = document.createElement("td");
      tdMark.textContent = row.mark ?? "-";
      if (failed) tdMark.classList.add("fail");
      tr.appendChild(tdMark);

      const tdLetter = document.createElement("td");
      tdLetter.textContent = row.letter ?? "-";
      if (failed) tdLetter.classList.add("fail");
      tr.appendChild(tdLetter);

      const tdPoint = document.createElement("td");
      tdPoint.textContent = row.point ?? "-";
      tr.appendChild(tdPoint);

      const tdGradePoints = document.createElement("td");
      tdGradePoints.textContent = gradePoints === null ? "-" : gradePoints.toFixed(2);
      tr.appendChild(tdGradePoints);

      if (failed) {
        failedBuckets.push({
          session: term.session,
          semester: term.semester,
          courseCode,
          title: course?.title || "-",
          mark: row.mark ?? "-",
          letter: row.letter ?? "-",
          point: row.point ?? "-",
        });
      }

      if (credits !== null) {
        termCredits += credits;
        const earned = toNumber(row.creditsEarned);
        if (earned !== null) {
          termCreditsEarned += earned;
        } else if (!failed) {
          termCreditsEarned += credits;
        }
      }
      if (credits !== null && gradePoints !== null) {
        termGradePoints += gradePoints;
      }

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);

    const tfoot = document.createElement("tfoot");
    const cumulativeGradePoints = totalGradePoints + termGradePoints;
    const cumulativeCredits = totalCredits + termCredits;
    const termGpa = termCredits ? termGradePoints / termCredits : null;
    const cumulativeCgpa = cumulativeCredits ? cumulativeGradePoints / cumulativeCredits : null;

    const tr = document.createElement("tr");
    const summaryCells = [
      ["Credits Attempted", termCredits ? numberFmt.format(termCredits) : "-"],
      ["Credits Earned", termCreditsEarned ? numberFmt.format(termCreditsEarned) : "-"],
      ["Grade Pts (Sem)", termGradePoints ? termGradePoints.toFixed(2) : "-"],
      ["Grade Pts (Cum)", cumulativeGradePoints ? cumulativeGradePoints.toFixed(2) : "-"],
      ["GPA", termGpa === null ? "--" : termGpa.toFixed(2)],
      ["CGPA", cumulativeCgpa === null ? "--" : cumulativeCgpa.toFixed(2)],
    ];

    summaryCells.forEach(([label, value], index) => {
      const td = document.createElement("td");
      const strong = document.createElement("strong");
      strong.textContent = `${label}: `;
      td.appendChild(strong);
      td.appendChild(document.createTextNode(String(value)));
      if (index === summaryCells.length - 1) td.colSpan = 2;
      tr.appendChild(td);
    });

    tfoot.appendChild(tr);

    table.appendChild(tfoot);
    section.appendChild(table);

    const termSummary = document.createElement("div");
    termSummary.className = "slip-meta";
    const creditsLine = document.createElement("div");
    const creditsLabel = document.createElement("strong");
    creditsLabel.textContent = "Credits: ";
    creditsLine.appendChild(creditsLabel);
    creditsLine.appendChild(
      document.createTextNode(termCredits ? numberFmt.format(termCredits) : "-")
    );
    termSummary.appendChild(creditsLine);

    const gpaLine = document.createElement("div");
    const gpaLabel = document.createElement("strong");
    gpaLabel.textContent = "GPA: ";
    gpaLine.appendChild(gpaLabel);
    gpaLine.appendChild(
      document.createTextNode(termGpa === null ? "--" : termGpa.toFixed(2))
    );
    termSummary.appendChild(gpaLine);
    section.appendChild(termSummary);

    totalCredits += termCredits;
    totalGradePoints += termGradePoints;

    elSlipTerms.appendChild(section);
  }

  const cgpa = totalCredits ? totalGradePoints / totalCredits : null;
  elSlipCGPA.textContent = cgpa === null ? "--" : cgpa.toFixed(2);

  if (failedBuckets.length) {
    const failedSection = document.createElement("div");
    failedSection.className = "slip-section";
    const heading = document.createElement("h4");
    heading.textContent = "Failed Courses";
    failedSection.appendChild(heading);

    const table = document.createElement("table");
    table.className = "slip-table";
    const colgroup = document.createElement("colgroup");
    const colWidths = ["16%", "10%", "14%", "30%", "10%", "10%"];
    for (const width of colWidths) {
      const col = document.createElement("col");
      col.style.width = width;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Session</th>
        <th>Sem</th>
        <th>Course</th>
        <th>Title</th>
        <th>Mark</th>
        <th>Letter</th>
      </tr>
    `;
    table.appendChild(thead);
    const tbody = document.createElement("tbody");

    for (const row of failedBuckets) {
      const tr = document.createElement("tr");
      const cells = [
        row.session || "-",
        row.semester ?? "-",
        row.courseCode || "-",
        row.title || "-",
        row.mark ?? "-",
        row.letter ?? "-",
      ];
      for (const value of cells) {
        const td = document.createElement("td");
        td.textContent = value;
        td.classList.add("fail");
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    table.appendChild(tbody);
    failedSection.appendChild(table);
    elSlipTerms.appendChild(failedSection);
  }
}

function openSlip(studentId) {
  buildSlip(studentId);
  if (!elResultSlipModal) return;
  elResultSlipModal.classList.add("active");
  elResultSlipModal.setAttribute("aria-hidden", "false");
}

function closeSlip() {
  if (!elResultSlipModal) return;
  elResultSlipModal.classList.remove("active");
  elResultSlipModal.setAttribute("aria-hidden", "true");
}

function populateResultFilters() {
  if (!elResultSessionSelect) return;
  const studentMap = buildStudentMap();

  const currentSession = elResultSessionSelect.value;

  const sessions = [...new Set(
    [...studentMap.values()]
      .map((student) => String(student.intake ?? "").trim())
      .filter(Boolean)
  )].sort();
  elResultSessionSelect.textContent = "";
  const sessionDefault = document.createElement("option");
  sessionDefault.value = "";
  sessionDefault.textContent = "All intakes";
  elResultSessionSelect.appendChild(sessionDefault);
  for (const session of sessions) {
    const option = document.createElement("option");
    option.value = session;
    option.textContent = session;
    elResultSessionSelect.appendChild(option);
  }

  if (currentSession) elResultSessionSelect.value = currentSession;
}

function renderResultsView() {
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();

  const search = elResultSearch?.value?.trim().toLowerCase() ?? "";
  const intakeFilter = elResultSessionSelect?.value ?? "";
  const viewMode = elResultViewMode?.value ?? "students";

  const filtered = resultState.results.filter((row) => {
    const studentId = String(row.studentId ?? "").trim();
    if (intakeFilter) {
      const student = studentMap.get(studentId);
      const intake = String(student?.intake ?? "").trim();
      if (intake !== intakeFilter) return false;
    }
    if (activeCourseDetailCode) {
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      if (courseCode !== activeCourseDetailCode) return false;
    }
    if (search) {
      const student = studentMap.get(studentId);
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      const course = courseMap.get(courseCode);
      const label = `${student?.studentId ?? ""} ${student?.name ?? ""} ${courseCode} ${course?.title ?? ""}`
        .toLowerCase();
      if (!label.includes(search)) return false;
    }
    return true;
  });

  const studentIds = new Set(filtered.map((row) => String(row.studentId ?? "").trim()).filter(Boolean));

  renderResultStats(filtered, studentIds);
  if (activeCourseDetailCode) {
    const stats = filtered.reduce(
      (acc, row) => {
        acc.records += 1;
        if (row.passedCourse === true) acc.passed += 1;
        const mark = toNumber(row.mark);
        if (mark !== null) acc.marks.push(mark);
        if (row.studentId) acc.students.add(String(row.studentId).trim());
        return acc;
      },
      { records: 0, passed: 0, marks: [], students: new Set() }
    );
    const avg = stats.marks.length ? stats.marks.reduce((a, b) => a + b, 0) / stats.marks.length : null;
    const passRate = stats.records ? stats.passed / stats.records : null;
    if (elCourseDetailStats) elCourseDetailStats.style.display = "grid";
    if (elCourseDetailRecords) setText(elCourseDetailRecords, numberFmt.format(stats.records));
    if (elCourseDetailPassRate) setText(elCourseDetailPassRate, passRate === null ? "--" : percentFmt.format(passRate));
    if (elCourseDetailAvgMark) setText(elCourseDetailAvgMark, avg === null ? "--" : avg.toFixed(1));
    if (elCourseDetailStudents) setText(elCourseDetailStudents, numberFmt.format(stats.students.size));
    const courseMap = buildCourseMap();
    const course = courseMap.get(activeCourseDetailCode);
    const courseLabel = course
      ? `${activeCourseDetailCode} - ${course.title ?? ""}`.trim()
      : activeCourseDetailCode;
    if (elCourseDetailHeader) elCourseDetailHeader.style.display = "flex";
    if (elCourseDetailTitle) elCourseDetailTitle.textContent = courseLabel;
    const rows = filtered.map((row) => {
      const studentId = String(row.studentId ?? "").trim();
      const student = studentMap.get(studentId);
      return {
        resultId: row.resultId,
        courseCode: activeCourseDetailCode,
        courseLabel,
        studentId,
        studentName: student?.name ?? "",
        session: row.session ?? "",
        semester: row.semester ?? "",
        mark: row.mark ?? "",
        letter: row.letter ?? "",
        point: row.point ?? "",
        passedCourse: row.passedCourse,
        failed: isFailedResult(row),
        creditsEarned: row.creditsEarned ?? null,
      };
    });
    const sortedRows = sortResultRows(rows);
    const totalPages = Math.max(1, Math.ceil(rows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const pageRows = sortedRows.slice(startIndex, startIndex + resultPaging.pageSize);
    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;
    renderResultTable(pageRows, "detail");
    if (elResultPager) elResultPager.style.display = "flex";
    return;
  } else if (elCourseDetailStats) {
    elCourseDetailStats.style.display = "none";
    if (elCourseDetailHeader) elCourseDetailHeader.style.display = "none";
  }
  if (viewMode === "courses" && !activeCourseDetailCode) {
    const courseCounts = new Map();
    for (const row of filtered) {
      const courseCode = normalizeCourseCode(row.courseCode ?? "");
      if (!courseCode) continue;
      const entry = courseCounts.get(courseCode) ?? {
        records: 0,
        passed: 0,
        marks: [],
      };
      entry.records += 1;
      if (row.passedCourse === true) entry.passed += 1;
      const mark = toNumber(row.mark);
      if (mark !== null) entry.marks.push(mark);
      courseCounts.set(courseCode, entry);
    }

    const listRows = [...courseCounts.entries()]
      .map(([courseCode, entry]) => {
        const course = courseMap.get(courseCode);
        const courseLabel = course ? `${courseCode} - ${course.title ?? ""}`.trim() : courseCode;
        const avgMark = entry.marks.length
          ? entry.marks.reduce((a, b) => a + b, 0) / entry.marks.length
          : null;
        const passRate = entry.records ? entry.passed / entry.records : null;
        return {
          courseCode,
          courseLabel,
          records: entry.records,
          avgMark,
          passRate,
        };
      })
      .sort((a, b) => a.courseLabel.localeCompare(b.courseLabel));

    const totalPages = Math.max(1, Math.ceil(listRows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const tableRows = sortResultRows(listRows).slice(startIndex, startIndex + resultPaging.pageSize);

    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;

    renderResultTable(tableRows, "courses");
  } else {
    const studentCounts = new Map();
    const studentSemesterMax = new Map();
    for (const row of filtered) {
      const studentId = String(row.studentId ?? "").trim();
      if (!studentId) continue;
      studentCounts.set(studentId, (studentCounts.get(studentId) ?? 0) + 1);
      const sem = toNumber(row.semester);
      if (sem !== null) {
        const current = studentSemesterMax.get(studentId) ?? 0;
        if (sem > current) studentSemesterMax.set(studentId, sem);
      }
    }

    const cgpaByStudent = computeStudentCgpa(filtered, resultState.courses);

    const listRows = [...studentCounts.entries()]
      .map(([studentId, records]) => {
        const student = studentMap.get(studentId) ?? {};
        const cgpa = cgpaByStudent.get(studentId) ?? null;
        return {
          studentId,
          studentName: student.name ?? "",
          records,
          currentSemester: studentSemesterMax.get(studentId) ?? "-",
          cgpa: cgpa === null ? "--" : cgpa.toFixed(2),
        };
      })
      .sort((a, b) => {
        const aName = String(a.studentName ?? "").toLowerCase();
        const bName = String(b.studentName ?? "").toLowerCase();
        if (aName && bName) return aName.localeCompare(bName);
        return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
      });

    const totalPages = Math.max(1, Math.ceil(listRows.length / resultPaging.pageSize));
    if (resultPaging.page > totalPages) resultPaging.page = totalPages;
    const startIndex = (resultPaging.page - 1) * resultPaging.pageSize;
    const tableRows = sortResultRows(listRows).slice(startIndex, startIndex + resultPaging.pageSize);

    if (elResultPageMeta) elResultPageMeta.textContent = `Page ${resultPaging.page} of ${totalPages}`;
    if (btnPrevPage) btnPrevPage.disabled = resultPaging.page <= 1;
    if (btnNextPage) btnNextPage.disabled = resultPaging.page >= totalPages;

    renderResultTable(tableRows, "students");
  }
}

function sortResultRows(rows) {
  const sortKey = resultSortState.key;
  if (!sortKey) return rows;
  const getSortValue = (row, key) => {
    if (key === "status") {
      if (row.failed) return "failed";
      if (row.passedCourse === true) return "passed";
      return "unknown";
    }
    return row[key];
  };
  return rows.slice().sort((a, b) => {
    const av = getSortValue(a, sortKey);
    const bv = getSortValue(b, sortKey);
    if (av === null || av === undefined) return 1;
    if (bv === null || bv === undefined) return -1;
    if (typeof av === "number" && typeof bv === "number") {
      return resultSortState.dir === "asc" ? av - bv : bv - av;
    }
    const as = String(av).toLowerCase();
    const bs = String(bv).toLowerCase();
    return resultSortState.dir === "asc" ? as.localeCompare(bs) : bs.localeCompare(as);
  });
}

function openStudentEdit(studentId) {
  const student = resultState.students.find(
    (s) => String(s.studentId ?? "").trim() === studentId
  );
  if (!student || !elStudentEditModal) return;
  activeStudentEditId = studentId;
  elEditStudentId.value = studentId;
  elEditStudentName.value = student.name ?? "";
  elEditStudentIc.value = student.ic ?? "";
  elEditStudentIntake.value = student.intake ?? "";
  elStudentEditModal.classList.add("active");
  elStudentEditModal.setAttribute("aria-hidden", "false");
}

function closeStudentEdit() {
  activeStudentEditId = "";
  if (!elStudentEditModal) return;
  elStudentEditModal.classList.remove("active");
  elStudentEditModal.setAttribute("aria-hidden", "true");
}

function openMarkEdit(row) {
  if (!row?.resultId || !elMarkEditModal) return;
  activeMarkEditId = row.resultId;
  activeMarkEditRow = row;
  activeMarkEditValue = row.mark ?? "";
  if (elMarkEditVerify) elMarkEditVerify.checked = false;
  updateMarkEditPreview();
  elMarkEditModal.classList.add("active");
  elMarkEditModal.setAttribute("aria-hidden", "false");
  queueMicrotask(() => elMarkEditInput?.focus());
}

function closeMarkEdit() {
  activeMarkEditId = "";
  activeMarkEditValue = "";
  activeMarkEditRow = null;
  if (!elMarkEditModal) return;
  elMarkEditModal.classList.remove("active");
  elMarkEditModal.setAttribute("aria-hidden", "true");
}

function updateMarkEditPreview() {
  if (!activeMarkEditRow) return;
  const row = activeMarkEditRow;
  const studentMap = buildStudentMap();
  const courseMap = buildCourseMap();
  const student = studentMap.get(String(row.studentId ?? "").trim());
  const course = courseMap.get(normalizeCourseCode(row.courseCode ?? ""));
  const studentLabel = student ? `${student.studentId ?? ""} - ${student.name ?? ""}`.trim() : row.studentId ?? "-";
  const courseLabel = course
    ? `${normalizeCourseCode(row.courseCode ?? "")} - ${course.title ?? ""}`.trim()
    : normalizeCourseCode(row.courseCode ?? "");

  if (elMarkEditStudent) elMarkEditStudent.textContent = studentLabel || "-";
  if (elMarkEditCourse) elMarkEditCourse.textContent = courseLabel || "-";
  if (elMarkEditSession) elMarkEditSession.textContent = row.session ?? "-";
  if (elMarkEditSemester) elMarkEditSemester.textContent = row.semester ?? "-";
  if (elMarkEditCurrent) elMarkEditCurrent.textContent = row.mark ?? "-";
  if (elMarkEditInput) elMarkEditInput.value = activeMarkEditValue ?? "";

  const raw = String(activeMarkEditValue ?? "").trim();
  let mark = null;
  let letter = "";
  let point = null;
  let passedCourse = null;
  let warning = "";
  let isValid = true;

  if (raw) {
    const n = Number(raw);
    if (!Number.isFinite(n)) {
      isValid = false;
      warning = "Mark must be a number.";
    } else if (n < 0 || n > 100) {
      isValid = false;
      warning = "Mark must be between 0 and 100.";
    } else {
      mark = n;
      const grade = getGradeForMark(mark, course?.isMPUCourse === true);
      letter = grade.letter;
      point = grade.point;
      passedCourse = grade.passed;
    }
  } else {
    warning = "Blank mark will clear letter/point/status.";
  }

  if (elMarkEditLetter) elMarkEditLetter.textContent = letter || "-";
  if (elMarkEditPoint) elMarkEditPoint.textContent = point === null ? "-" : String(point);
  if (elMarkEditStatus) {
    elMarkEditStatus.textContent = passedCourse === null ? "-" : (passedCourse ? "Passed" : "Failed");
    elMarkEditStatus.classList.toggle("remark", passedCourse === false);
  }
  if (elMarkEditWarning) {
    elMarkEditWarning.textContent = warning;
    elMarkEditWarning.style.display = warning ? "block" : "none";
  }
  if (btnConfirmMarkEdit) {
    const verified = elMarkEditVerify?.checked;
    btnConfirmMarkEdit.disabled = !isValid || !verified;
  }
}

async function confirmMarkEdit() {
  if (!activeMarkEditRow) return;
  const current = resultState.results.find((r) => r.resultId === activeMarkEditRow.resultId);
  if (!current) return;

  const raw = String(activeMarkEditValue ?? "").trim();
  let mark = null;
  let letter = "";
  let point = null;
  let passedCourse = null;

  if (raw) {
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 0 || n > 100) {
      alert("Mark must be a number between 0 and 100.");
      return;
    }
    mark = n;
    const courseCode = normalizeCourseCode(current.courseCode ?? "");
    const course = buildCourseMap().get(courseCode);
    const grade = getGradeForMark(mark, course?.isMPUCourse === true);
    letter = grade.letter;
    point = grade.point;
    passedCourse = grade.passed;
  }

  const updated = {
    ...current,
    mark,
    letter,
    point,
    gradePoints: null,
    passedCourse,
  };

  try {
    await upsertMany(STORES.results, [updated]);
    const index = resultState.results.findIndex((r) => r.resultId === current.resultId);
    if (index >= 0) resultState.results[index] = updated;
    closeMarkEdit();
    await refreshStats();
    renderResultsView();
  } catch (e) {
    console.error(e);
    alert(`Failed to save mark: ${e.message ?? e}`);
  }
}


function renderCourseResults(courseCode) {
  activeCourseDetailCode = normalizeCourseCode(courseCode);
  if (elResultViewMode) elResultViewMode.value = "courses";
  resultPaging.page = 1;
  renderResultsView();
}

async function refreshResults() {
  const [students, courses, results] = await Promise.all([
    getAllRecords(STORES.students),
    getAllRecords(STORES.courses),
    getAllRecords(STORES.results),
  ]);

  resultState.students = students;
  resultState.courses = courses;
  resultState.results = results;
  resultPaging.page = 1;

  populateResultFilters();
  renderResultsView();
  refreshEnrollmentModule();
}

function updateEnrollmentSelectedCount() {
  if (!elEnrollmentSelectedCount) return;
  const selectedCount = enrollmentState.offeredCourseCodes.length;
  elEnrollmentSelectedCount.textContent = `${numberFmt.format(selectedCount)} selected (max ${ENROLLMENT_MAX_OFFERED})`;
}

function getSelectedEnrollmentCourseCodes() {
  if (!elEnrollmentCourseGrid) return [];
  return [...elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item.selected")]
    .map((item) => normalizeCourseCode(item.dataset.courseCode))
    .filter(Boolean);
}

function populateEnrollmentCourseOptions() {
  if (!elEnrollmentCourseGrid) return;
  const previous = new Set((enrollmentState.offeredCourseCodes ?? []).slice(0, ENROLLMENT_MAX_OFFERED));
  const courses = [...buildCourseMap().values()].sort((a, b) =>
    normalizeCourseCode(a.courseCode ?? "").localeCompare(normalizeCourseCode(b.courseCode ?? ""))
  );

  elEnrollmentCourseGrid.textContent = "";
  for (const course of courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    const item = document.createElement("button");
    item.type = "button";
    item.className = "enrollment-course-item";
    item.dataset.courseCode = code;
    if (previous.has(code)) item.classList.add("selected");
    item.addEventListener("click", () => {
      const isSelected = item.classList.contains("selected");
      if (!isSelected) {
        const selectedCodes = getSelectedEnrollmentCourseCodes();
        if (selectedCodes.length >= ENROLLMENT_MAX_OFFERED) {
          alert(`Maximum ${ENROLLMENT_MAX_OFFERED} offered courses allowed.`);
          return;
        }
        item.classList.add("selected");
      } else {
        item.classList.remove("selected");
      }
      enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
      updateEnrollmentSelectedCount();
    });
    const content = document.createElement("div");
    const strong = document.createElement("strong");
    strong.textContent = code;
    const span = document.createElement("span");
    span.textContent = String(course.title ?? "").trim() || "Untitled course";
    content.appendChild(strong);
    content.appendChild(span);
    item.appendChild(content);
    elEnrollmentCourseGrid.appendChild(item);
  }

  enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
  updateEnrollmentSelectedCount();
}

function getEnrollmentStudents() {
  const studentMap = buildStudentMap();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    if (!studentId || studentMap.has(studentId)) continue;
    studentMap.set(studentId, {
      studentId,
      name: "",
      intake: "",
      ic: "",
    });
  }
  return [...studentMap.values()].sort((a, b) => {
    const aName = String(a.name ?? "").trim().toLowerCase();
    const bName = String(b.name ?? "").trim().toLowerCase();
    if (aName && bName) return aName.localeCompare(bName);
    if (aName) return -1;
    if (bName) return 1;
    return String(a.studentId ?? "").localeCompare(String(b.studentId ?? ""));
  });
}

function buildEnrollmentSlips(options = {}) {
  const suppressMinAlert = options.suppressMinAlert === true;
  const selectedCodes = getSelectedEnrollmentCourseCodes();
  enrollmentState.offeredCourseCodes = selectedCodes;
  updateEnrollmentSelectedCount();

  if (!selectedCodes.length) {
    enrollmentState.slips = [];
    activeEnrollmentStudentId = "";
    renderEnrollmentList();
    closeEnrollmentSlip();
    if (!suppressMinAlert) {
      alert("Please select at least 1 offered course.");
    }
    return;
  }

  const courseMap = buildCourseMap();
  const progressByStudent = new Map();
  for (const row of resultState.results) {
    const studentId = String(row.studentId ?? "").trim();
    const courseCode = normalizeCourseCode(row.courseCode ?? "");
    if (!studentId || !courseCode) continue;
    let progress = progressByStudent.get(studentId);
    if (!progress) {
      progress = { attempted: new Set(), passed: new Set(), completedCredits: 0, currentSemester: null };
      progressByStudent.set(studentId, progress);
    }
    progress.attempted.add(courseCode);
    const semester = toNumber(row.semester);
    if (semester !== null) {
      progress.currentSemester =
        progress.currentSemester === null ? semester : Math.max(progress.currentSemester, semester);
    }
    if (row.passedCourse === true) {
      if (!progress.passed.has(courseCode)) {
        progress.passed.add(courseCode);
        const earned = toNumber(row.creditsEarned);
        const fallbackCredits = toNumber(courseMap.get(courseCode)?.credits);
        progress.completedCredits += earned ?? fallbackCredits ?? 0;
      }
    }
  }

  const slips = [];
  for (const student of getEnrollmentStudents()) {
    const studentId = String(student.studentId ?? "").trim();
    if (!studentId) continue;
    const progress = progressByStudent.get(studentId) ?? {
      attempted: new Set(),
      passed: new Set(),
      completedCredits: 0,
      currentSemester: null,
    };
    const courses = [];
    for (const courseCode of selectedCodes) {
      if (progress.passed.has(courseCode)) continue;
      const course = courseMap.get(courseCode);
      courses.push({
        courseCode,
        title: String(course?.title ?? "").trim(),
        remark: progress.attempted.has(courseCode) ? "Repeat" : "New",
      });
    }
    slips.push({
      studentId,
      name: String(student.name ?? "").trim(),
      intake: String(student.intake ?? "").trim(),
      currentSemester: progress.currentSemester ?? null,
      completedCredits: progress.completedCredits ?? 0,
      courses,
    });
  }

  enrollmentState.slips = slips;
  renderEnrollmentList();
  if (activeEnrollmentStudentId && !slips.some((slip) => slip.studentId === activeEnrollmentStudentId)) {
    closeEnrollmentSlip();
  }
}

function renderEnrollmentList() {
  if (!elEnrollmentListTable || !elEnrollmentListEmpty || !elEnrollmentListBody) return;
  const search = String(elEnrollmentSearch?.value ?? "").trim().toLowerCase();
  const filtered = search
    ? enrollmentState.slips.filter((slip) =>
        `${slip.studentId} ${slip.name}`.toLowerCase().includes(search)
      )
    : enrollmentState.slips;

  if (elEnrollmentStudentCount) {
    elEnrollmentStudentCount.textContent = numberFmt.format(filtered.length);
  }

  if (!filtered.length) {
    elEnrollmentListBody.textContent = "";
    elEnrollmentListTable.style.display = "none";
    elEnrollmentListEmpty.style.display = "block";
    return;
  }

  elEnrollmentListEmpty.style.display = "none";
  elEnrollmentListTable.style.display = "table";
  elEnrollmentListBody.textContent = "";

  for (const slip of filtered) {
    const tr = document.createElement("tr");

    const tdId = document.createElement("td");
    tdId.textContent = slip.studentId || "-";
    tr.appendChild(tdId);

    const tdName = document.createElement("td");
    tdName.textContent = slip.name || "-";
    tr.appendChild(tdName);

    const tdIntake = document.createElement("td");
    tdIntake.textContent = slip.intake || "-";
    tr.appendChild(tdIntake);

    const tdSemester = document.createElement("td");
    tdSemester.textContent = slip.currentSemester === null ? "-" : String(slip.currentSemester);
    tr.appendChild(tdSemester);

    const tdProgress = document.createElement("td");
    const completedCredits = toNumber(slip.completedCredits) ?? 0;
    const progressRatio = ENROLLMENT_TARGET_CREDITS
      ? Math.min(1, Math.max(0, completedCredits / ENROLLMENT_TARGET_CREDITS))
      : 0;
    tdProgress.textContent = `${numberFmt.format(completedCredits)} / ${ENROLLMENT_TARGET_CREDITS}`;
    tdProgress.title = percentFmt.format(progressRatio);
    tr.appendChild(tdProgress);

    const tdCount = document.createElement("td");
    tdCount.textContent = numberFmt.format(slip.courses.length);
    tr.appendChild(tdCount);

    const tdStatus = document.createElement("td");
    tdStatus.textContent = slip.courses.length ? "Enrollment Needed" : "No Enrollment";
    tr.appendChild(tdStatus);

    const tdAction = document.createElement("td");
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View slip";
    btnView.addEventListener("click", () => {
      openEnrollmentSlip(slip.studentId);
    });
    tdAction.appendChild(btnView);
    tr.appendChild(tdAction);

    elEnrollmentListBody.appendChild(tr);
  }
}

function openEnrollmentSlip(studentId) {
  activeEnrollmentStudentId = String(studentId ?? "").trim();
  if (!activeEnrollmentStudentId || !elEnrollmentSlipModal) return;
  renderEnrollmentSlip();
  elEnrollmentSlipModal.classList.add("active");
  elEnrollmentSlipModal.setAttribute("aria-hidden", "false");
}

function closeEnrollmentSlip() {
  activeEnrollmentStudentId = "";
  if (!elEnrollmentSlipModal) return;
  elEnrollmentSlipModal.classList.remove("active");
  elEnrollmentSlipModal.setAttribute("aria-hidden", "true");
}

function renderEnrollmentSlip() {
  if (!elEnrollmentSlipBody) return;
  const slip = enrollmentState.slips.find((row) => row.studentId === activeEnrollmentStudentId);
  if (!slip) {
    if (elEnrollmentSlipCount) elEnrollmentSlipCount.textContent = "0 course(s)";
    if (elEnrollmentSlipStudentMeta) elEnrollmentSlipStudentMeta.textContent = "--";
    elEnrollmentSlipBody.textContent = "";
    return;
  }

  if (elEnrollmentSlipStudentMeta) {
    elEnrollmentSlipStudentMeta.textContent = `${slip.studentId} - ${slip.name || "-"} | Intake: ${slip.intake || "-"}`;
  }
  if (elEnrollmentSlipCount) {
    elEnrollmentSlipCount.textContent = `${numberFmt.format(slip.courses.length)} course(s)`;
  }

  elEnrollmentSlipBody.textContent = "";
  if (!slip.courses.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 3;
    td.textContent = "No enrollment needed.";
    tr.appendChild(td);
    elEnrollmentSlipBody.appendChild(tr);
    return;
  }

  for (const course of slip.courses) {
    const tr = document.createElement("tr");
    const tdCode = document.createElement("td");
    tdCode.textContent = course.courseCode;
    tr.appendChild(tdCode);
    const tdTitle = document.createElement("td");
    tdTitle.textContent = course.title || "-";
    tr.appendChild(tdTitle);
    const tdRemark = document.createElement("td");
    tdRemark.textContent = course.remark;
    tr.appendChild(tdRemark);
    elEnrollmentSlipBody.appendChild(tr);
  }
}

function refreshEnrollmentModule() {
  populateEnrollmentCourseOptions();
  if (enrollmentState.offeredCourseCodes.length) {
    buildEnrollmentSlips();
    return;
  }
  enrollmentState.slips = [];
  renderEnrollmentList();
  closeEnrollmentSlip();
}

function setActiveTab(tabId) {
  for (const button of tabButtons) {
    const isActive = button.dataset.tab === tabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  for (const panel of tabPanels) {
    const isActive = panel.id === `panel-${tabId}`;
    panel.classList.toggle("active", isActive);
  }
}

function renderCourseList(courses) {
  if (!elCourseListTable || !elCourseListEmpty || !elCourseListBody) return;
  if (!courses.length) {
    elCourseListTable.style.display = "none";
    elCourseListEmpty.style.display = "block";
    return;
  }

  elCourseListEmpty.style.display = "none";
  elCourseListTable.style.display = "table";
  elCourseListBody.textContent = "";

  for (const course of courses) {
    const detail = courseCoverageDetails.get(course.courseCode);
    const taken = detail ? detail.taken.length : 0;
    const total = detail ? detail.totalStudents : 0;
    const coverageLabel = total ? `${taken} / ${total}` : "-";
    const tr = document.createElement("tr");

    const tdCode = document.createElement("td");
    tdCode.textContent = course.courseCode;
    tr.appendChild(tdCode);

    const tdTitle = document.createElement("td");
    tdTitle.textContent = course.title || "-";
    tr.appendChild(tdTitle);

    const tdCoverage = document.createElement("td");
    tdCoverage.textContent = coverageLabel;
    tr.appendChild(tdCoverage);

    const tdCredits = document.createElement("td");
    tdCredits.textContent = course.credits === null ? "-" : numberFmt.format(course.credits);
    tr.appendChild(tdCredits);

    const tdMPU = document.createElement("td");
    if (course.isMPUCourse === null) {
      tdMPU.textContent = "-";
    } else {
      tdMPU.textContent = course.isMPUCourse ? "Yes" : "No";
    }
    tr.appendChild(tdMPU);

    const tdAction = document.createElement("td");
    tdAction.className = "table-actions";
    const btnEdit = document.createElement("button");
    btnEdit.className = "btn secondary small";
    btnEdit.type = "button";
    btnEdit.textContent = "Edit";
    btnEdit.addEventListener("click", () => {
      fillCourseForm(course);
      elCourseCode.focus();
    });
    tdAction.appendChild(btnEdit);

    const btnDelete = document.createElement("button");
    btnDelete.className = "btn secondary small";
    btnDelete.type = "button";
    btnDelete.textContent = "Delete";
    btnDelete.addEventListener("click", async () => {
      const courseCode = normalizeCourseCode(course.courseCode ?? "");
      const detail = courseCoverageDetails.get(courseCode);
      if (detail && detail.taken.length > 0) {
        alert("Cannot delete this course because results already exist for it.");
        return;
      }
      const ok = window.confirm(`Delete course ${courseCode}?`);
      if (!ok) return;
      try {
        await deleteCourse(courseCode);
        if (activeCourseEditCode === courseCode) resetCourseForm();
        await refreshCourses();
        await refreshStats();
        await refreshResults();
      } catch (e) {
        console.error(e);
        alert(`Failed to delete course: ${e.message ?? e}`);
      }
    });
    tdAction.appendChild(btnDelete);
    tr.appendChild(tdAction);

    const tdView = document.createElement("td");
    tdView.className = "table-actions";
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View list";
    btnView.addEventListener("click", () => {
      openCourseCoverageDetail(course.courseCode);
    });
    tdView.appendChild(btnView);
    tr.appendChild(tdView);

    elCourseListBody.appendChild(tr);
  }
}

function renderCourseStats(courses) {
  const totalCourses = courses.length;
  const totalCredits = courses.reduce((sum, course) => {
    const credits = toNumber(course?.credits);
    return sum + (credits ?? 0);
  }, 0);
  setText(elCourseTotalCount, numberFmt.format(totalCourses));
  setText(elCourseTotalCredits, numberFmt.format(totalCredits));
}

function fillCourseForm(course) {
  activeCourseEditCode = normalizeCourseCode(course.courseCode ?? "");
  elCourseCode.value = course.courseCode ?? "";
  elCourseTitle.value = course.title ?? "";
  elCourseCredits.value = course.credits ?? "";
  if (course.isMPUCourse === null || course.isMPUCourse === undefined) {
    elCourseMPU.value = "";
  } else {
    elCourseMPU.value = course.isMPUCourse ? "true" : "false";
  }

  const detail = courseCoverageDetails.get(activeCourseEditCode);
  const hasCoverage = detail && detail.taken.length > 0;
  if (hasCoverage) {
    elCourseCode.setAttribute("disabled", "disabled");
    if (elCourseFormHint) {
      elCourseFormHint.textContent = "Course code is locked because results exist for this course.";
    }
  } else {
    elCourseCode.removeAttribute("disabled");
    if (elCourseFormHint) {
      elCourseFormHint.textContent = "Use the same course code to update an existing course.";
    }
  }
}

function resetCourseForm() {
  elCourseForm.reset();
  elCourseCode.value = "";
  elCourseTitle.value = "";
  elCourseCredits.value = "";
  elCourseMPU.value = "";
  activeCourseEditCode = "";
  elCourseCode.removeAttribute("disabled");
  if (elCourseFormHint) {
    elCourseFormHint.textContent = "Use the same course code to update an existing course.";
  }
}

async function refreshCourses() {
  const courses = await listCourses();
  renderCourseStats(courses);
  renderCourseList(courses);
}

async function refreshStats() {
  try {
    const [students, courses, results] = await Promise.all([
      getAllRecords(STORES.students),
      getAllRecords(STORES.courses),
      getAllRecords(STORES.results),
    ]);

    const totalStudents = students.length;
    const totalResults = results.length;
    const resultsByStudent = new Map();
    const failedStudents = new Set();
    const incompleteStudents = new Set();
    let incompleteCount = 0;
    const lowCgpaStudents = new Set();
    const resultIssueStudents = new Set();
    const seenCoursesByStudent = new Map();

    const statsByCourse = new Map();
    for (const result of results) {
      const courseCode = normalizeCourseCode(result.courseCode ?? "");
      if (!courseCode) continue;
      let entry = statsByCourse.get(courseCode);
      if (!entry) {
        entry = { takenIds: new Set(), results: 0 };
        statsByCourse.set(courseCode, entry);
      }
      entry.results += 1;
      if (result.studentId) entry.takenIds.add(result.studentId);

      const studentId = String(result.studentId ?? "").trim();
      if (studentId) {
        resultsByStudent.set(studentId, (resultsByStudent.get(studentId) ?? 0) + 1);
        if (isFailedResult(result)) failedStudents.add(studentId);
        if (isMissingMark(result)) resultIssueStudents.add(studentId);

        if (!seenCoursesByStudent.has(studentId)) {
          seenCoursesByStudent.set(studentId, new Set());
        }
        const seen = seenCoursesByStudent.get(studentId);
        if (seen.has(courseCode)) {
          resultIssueStudents.add(studentId);
        } else {
          seen.add(courseCode);
        }
      }
    }

    const courseMap = new Map();
    for (const course of courses) {
      const courseCode = normalizeCourseCode(course.courseCode ?? "");
      if (!courseCode) continue;
      courseMap.set(courseCode, String(course.title ?? "").trim());
    }
    for (const [courseCode] of statsByCourse) {
      if (!courseMap.has(courseCode)) courseMap.set(courseCode, "");
    }

    const studentMap = new Map();
    for (const student of students) {
      const id = String(student.studentId ?? "").trim();
      const ic = String(student.ic ?? "").trim();
      if (!id || !ic) {
        if (id) {
          incompleteStudents.add(id);
        } else {
          incompleteCount += 1;
        }
      }
      if (!id) continue;
      studentMap.set(id, student);
    }
    const allStudentIds = [...studentMap.keys()];

    courseCoverageDetails.clear();
    const totalCourses = courseMap.size;
    let activeCourses = 0;
    for (const [code, title] of [...courseMap.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
      const entry = statsByCourse.get(code);
      const taken = entry ? entry.takenIds.size : 0;
      if (taken > 0) activeCourses += 1;
      const coverage = totalStudents > 0 ? Math.round((taken / totalStudents) * 100) : 0;
      const takenIds = entry ? [...entry.takenIds] : [];
      const notTakenIds = allStudentIds.filter((id) => !(entry?.takenIds?.has(id)));
      const toRecord = (id) => {
        const student = studentMap.get(id);
        return { id, name: String(student?.name ?? "").trim() };
      };
      courseCoverageDetails.set(code, {
        code,
        title,
        coverage,
        totalStudents,
        taken: takenIds.map(toRecord).sort((a, b) => (a.id + a.name).localeCompare(b.id + b.name)),
        notTaken: notTakenIds
          .map(toRecord)
          .sort((a, b) => (a.id + a.name).localeCompare(b.id + b.name)),
      });
    }

    const cgpaByStudentAll = computeStudentCgpa(results, courses);
    for (const [studentId, cgpa] of cgpaByStudentAll.entries()) {
      if (cgpa < 2.0) lowCgpaStudents.add(studentId);
    }

    setText(elTotalStudents, numberFmt.format(totalStudents));
    setText(elTotalCourses, numberFmt.format(totalCourses));
    setText(elActiveCourses, numberFmt.format(activeCourses));
    setText(elTotalResults, numberFmt.format(totalResults));
    const getStudentRecord = (studentId) => {
      const student = studentMap.get(studentId) ?? {};
      return {
        studentId,
        name: student.name ?? "",
        ic: student.ic ?? "",
        intake: student.intake ?? "",
      };
    };

    studentIssueLists.failures = [...failedStudents].map(getStudentRecord);
    studentIssueLists.incomplete = [
      ...incompleteStudents,
      ...Array.from({ length: incompleteCount }, () => ""),
    ].map((id) => (id ? getStudentRecord(id) : { studentId: "", name: "", ic: "", intake: "" }));
    studentIssueLists.lowCgpa = [...lowCgpaStudents].map(getStudentRecord);
    studentIssueLists.resultIssues = [...resultIssueStudents].map(getStudentRecord);

    if (elIssueTypeSelect) {
      renderIssueList(elIssueTypeSelect.value);
    }

    setText(elActiveStudents, numberFmt.format(resultsByStudent.size));
    setText(elStudentsWithFailures, numberFmt.format(failedStudents.size));
    setText(elStudentsIncomplete, numberFmt.format(incompleteStudents.size + incompleteCount));
    setText(elStudentsLowCgpa, numberFmt.format(lowCgpaStudents.size));
    setText(elStudentsResultIssues, numberFmt.format(resultIssueStudents.size));
    const avgResults =
      totalStudents > 0 ? (totalResults / totalStudents) : 0;
    setText(elAvgResultsPerStudent, totalStudents ? avgResults.toFixed(1) : "--");
    setLastUpdated();
  } catch (e) {
    console.error(e);
  }
}

function openCourseCoverageDetail(courseCode) {
  if (!elCourseCoverageModal || !elCoverageCourseMeta || !elCoverageTakenCount || !elCoverageNotTakenCount || !elCoverageTakenList || !elCoverageNotTakenList) {
    return;
  }
  const detail = courseCoverageDetails.get(courseCode);
  if (!detail) return;
  activeCoverageCourse = courseCode;
  elCoverageCourseMeta.textContent = `${detail.code} ${detail.title ? "- " + detail.title : ""}`;
  elCoverageTakenCount.textContent = numberFmt.format(detail.taken.length);
  elCoverageNotTakenCount.textContent = numberFmt.format(detail.notTaken.length);
  elCoverageTakenList.textContent = detail.taken.length
    ? detail.taken.map((row) => `${row.id} - ${row.name}`.trim()).join("\n")
    : "—";
  elCoverageNotTakenList.textContent = detail.notTaken.length
    ? detail.notTaken.map((row) => `${row.id} - ${row.name}`.trim()).join("\n")
    : "—";
  elCourseCoverageModal.classList.add("active");
  elCourseCoverageModal.setAttribute("aria-hidden", "false");
}

function closeCourseCoverageDetail() {
  activeCoverageCourse = "";
  elCourseCoverageModal.classList.remove("active");
  elCourseCoverageModal.setAttribute("aria-hidden", "true");
}

function renderIssueList(type) {
  if (!elIssueListTable || !elIssueListEmpty || !elIssueListBody) return;
  const list = studentIssueLists[type] ?? [];
  const search = (elIssueSearch?.value ?? "").trim().toLowerCase();
  const filtered = search
    ? list.filter((student) => {
        const label = `${student.studentId ?? ""} ${student.name ?? ""}`.toLowerCase();
        return label.includes(search);
      })
    : list;

  if (elIssueListTitle) elIssueListTitle.textContent = issueTitles[type] ?? "Student List";
  if (elIssueListCount) elIssueListCount.textContent = numberFmt.format(filtered.length);

  const totalPages = Math.max(1, Math.ceil(filtered.length / issuePaging.pageSize));
  if (issuePaging.page > totalPages) issuePaging.page = totalPages;
  const startIndex = (issuePaging.page - 1) * issuePaging.pageSize;
  const pageRows = filtered.slice(startIndex, startIndex + issuePaging.pageSize);

  if (!pageRows.length) {
    elIssueListTable.style.display = "none";
    elIssueListEmpty.style.display = "block";
    if (elIssuePager) elIssuePager.style.display = "none";
    return;
  }

  if (elIssuePager) elIssuePager.style.display = "flex";
  if (elIssuePageMeta) elIssuePageMeta.textContent = `Page ${issuePaging.page} of ${totalPages}`;
  if (btnIssuePrev) btnIssuePrev.disabled = issuePaging.page <= 1;
  if (btnIssueNext) btnIssueNext.disabled = issuePaging.page >= totalPages;

  elIssueListEmpty.style.display = "none";
  elIssueListTable.style.display = "table";
  elIssueListBody.textContent = "";
  for (const student of pageRows) {
    const tr = document.createElement("tr");
    const tdId = document.createElement("td");
    tdId.textContent = student.studentId || "-";
    tr.appendChild(tdId);
    const tdName = document.createElement("td");
    tdName.textContent = student.name || "-";
    tr.appendChild(tdName);
    const tdIc = document.createElement("td");
    tdIc.textContent = student.ic || "-";
    tr.appendChild(tdIc);
    const tdIntake = document.createElement("td");
    tdIntake.textContent = student.intake || "-";
    tr.appendChild(tdIntake);
    const tdAction = document.createElement("td");
    const btnView = document.createElement("button");
    btnView.className = "btn secondary small";
    btnView.type = "button";
    btnView.textContent = "View results";
    btnView.addEventListener("click", () => {
      if (student.studentId) {
        openSlip(student.studentId);
      }
    });
    tdAction.appendChild(btnView);
    tr.appendChild(tdAction);
    elIssueListBody.appendChild(tr);
  }
}

function exportCoverageCSV() {
  if (!activeCoverageCourse) return;
  const detail = courseCoverageDetails.get(activeCoverageCourse);
  if (!detail) return;

  const rows = [
    ["Course Code", "Course Title", "Status", "Student ID", "Student Name"],
  ];

  for (const student of detail.taken) {
    rows.push([detail.code, detail.title, "Taken", student.id, student.name]);
  }
  for (const student of detail.notTaken) {
    rows.push([detail.code, detail.title, "Not Taken", student.id, student.name]);
  }

  const csv = rows
    .map((row) =>
      row
        .map((value) => {
          const s = String(value ?? "");
          if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
          return s;
        })
        .join(",")
    )
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const safeCode = detail.code || "course";
  link.href = url;
  link.download = `coverage-${safeCode}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

if (btnRefresh) {
  btnRefresh.addEventListener("click", async () => {
    await refreshStats();
    await refreshCourses();
    await refreshResults();
  });
}

for (const button of tabButtons) {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
  });
}

if (btnEnrollmentSelectAll) {
  btnEnrollmentSelectAll.addEventListener("click", () => {
    if (!elEnrollmentCourseGrid) return;
    const items = [...elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item")];
    items.forEach((item, index) => {
      item.classList.toggle("selected", index < ENROLLMENT_MAX_OFFERED);
    });
    if (items.length > ENROLLMENT_MAX_OFFERED) {
      alert(`Selected first ${ENROLLMENT_MAX_OFFERED} courses (maximum allowed).`);
    }
    enrollmentState.offeredCourseCodes = getSelectedEnrollmentCourseCodes();
    updateEnrollmentSelectedCount();
  });
}

if (btnEnrollmentClearAll) {
  btnEnrollmentClearAll.addEventListener("click", () => {
    if (!elEnrollmentCourseGrid) return;
    for (const item of elEnrollmentCourseGrid.querySelectorAll(".enrollment-course-item")) {
      item.classList.remove("selected");
    }
    enrollmentState.offeredCourseCodes = [];
    updateEnrollmentSelectedCount();
    buildEnrollmentSlips({ suppressMinAlert: true });
  });
}

if (btnEnrollmentBuild) {
  btnEnrollmentBuild.addEventListener("click", () => {
    buildEnrollmentSlips();
  });
}

if (elEnrollmentSearch) {
  elEnrollmentSearch.addEventListener("input", () => {
    renderEnrollmentList();
  });
}

resultSubtabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveResultSubtab(button.dataset.subtab);
  });
});

if (elResultSessionSelect) {
  elResultSessionSelect.addEventListener("change", () => {
    resultPaging.page = 1;
    activeCourseDetailCode = "";
    renderResultsView();
  });
}
if (elResultSearch) {
  elResultSearch.addEventListener("input", () => {
    resultPaging.page = 1;
    renderResultsView();
  });
}
if (elResultViewMode) {
  elResultViewMode.addEventListener("change", () => {
    resultPaging.page = 1;
    activeCourseDetailCode = "";
    renderResultsView();
  });
}


if (btnBackToCourseList) {
  btnBackToCourseList.addEventListener("click", () => {
    activeCourseDetailCode = "";
    resultPaging.page = 1;
    renderResultsView();
  });
}

if (btnCloseStudentEdit) {
  btnCloseStudentEdit.addEventListener("click", () => {
    closeStudentEdit();
  });
}

if (btnSaveStudentEdit) {
  btnSaveStudentEdit.addEventListener("click", async () => {
    if (!activeStudentEditId) return;
    const updated = {
      studentId: activeStudentEditId,
      name: elEditStudentName.value.trim(),
      ic: elEditStudentIc.value.trim(),
      intake: elEditStudentIntake.value.trim(),
    };
    try {
      await upsertMany(STORES.students, [updated]);
      await refreshResults();
      await refreshStats();
      closeStudentEdit();
    } catch (e) {
      console.error(e);
    }
  });
}

if (elStudentEditModal) {
  elStudentEditModal.addEventListener("click", (event) => {
    if (event.target === elStudentEditModal) closeStudentEdit();
  });
}


if (btnPrintSlip) {
  btnPrintSlip.addEventListener("click", () => {
    window.print();
  });
}

if (btnCloseSlip) {
  btnCloseSlip.addEventListener("click", () => {
    closeSlip();
  });
}

if (btnCloseEnrollmentSlip) {
  btnCloseEnrollmentSlip.addEventListener("click", () => {
    closeEnrollmentSlip();
  });
}

if (btnCloseMarkEdit) {
  btnCloseMarkEdit.addEventListener("click", () => {
    closeMarkEdit();
  });
}

if (btnConfirmMarkEdit) {
  btnConfirmMarkEdit.addEventListener("click", () => {
    confirmMarkEdit();
  });
}

if (elMarkEditInput) {
  elMarkEditInput.addEventListener("input", () => {
    activeMarkEditValue = elMarkEditInput.value;
    updateMarkEditPreview();
  });
}

if (elMarkEditVerify) {
  elMarkEditVerify.addEventListener("change", () => {
    updateMarkEditPreview();
  });
}

if (elMarkEditModal) {
  elMarkEditModal.addEventListener("click", (event) => {
    if (event.target === elMarkEditModal) closeMarkEdit();
  });
}

if (elResultSlipModal) {
  elResultSlipModal.addEventListener("click", (event) => {
    if (event.target === elResultSlipModal) {
      closeSlip();
    }
  });
}

if (elEnrollmentSlipModal) {
  elEnrollmentSlipModal.addEventListener("click", (event) => {
    if (event.target === elEnrollmentSlipModal) {
      closeEnrollmentSlip();
    }
  });
}

if (btnCloseCoverage) {
  btnCloseCoverage.addEventListener("click", () => {
    closeCourseCoverageDetail();
  });
}

if (btnExportCoverage) {
  btnExportCoverage.addEventListener("click", () => {
    exportCoverageCSV();
  });
}

if (elCourseCoverageModal) {
  elCourseCoverageModal.addEventListener("click", (event) => {
    if (event.target === elCourseCoverageModal) {
      closeCourseCoverageDetail();
    }
  });
}

const openIssuesTab = (type) => {
  if (elIssueTypeSelect) elIssueTypeSelect.value = type;
  issuePaging.page = 1;
  if (elIssueSearch) elIssueSearch.value = "";
  setActiveTab("results");
  setActiveResultSubtab("issues");
  renderIssueList(type);
};

if (btnViewFailures) {
  btnViewFailures.addEventListener("click", () => openIssuesTab("failures"));
}
if (btnViewIncomplete) {
  btnViewIncomplete.addEventListener("click", () => openIssuesTab("incomplete"));
}
if (btnViewLowCgpa) {
  btnViewLowCgpa.addEventListener("click", () => openIssuesTab("lowCgpa"));
}
if (btnViewResultIssues) {
  btnViewResultIssues.addEventListener("click", () => openIssuesTab("resultIssues"));
}

if (elIssueTypeSelect) {
  elIssueTypeSelect.addEventListener("change", () => {
    issuePaging.page = 1;
    renderIssueList(elIssueTypeSelect.value);
  });
}

if (elIssueSearch) {
  elIssueSearch.addEventListener("input", () => {
    issuePaging.page = 1;
    renderIssueList(elIssueTypeSelect?.value ?? "failures");
  });
}

if (btnIssuePrev) {
  btnIssuePrev.addEventListener("click", () => {
    if (issuePaging.page > 1) {
      issuePaging.page -= 1;
      renderIssueList(elIssueTypeSelect?.value ?? "failures");
    }
  });
}

if (btnIssueNext) {
  btnIssueNext.addEventListener("click", () => {
    issuePaging.page += 1;
    renderIssueList(elIssueTypeSelect?.value ?? "failures");
  });
}

if (btnPrevPage) {
  btnPrevPage.addEventListener("click", () => {
    if (resultPaging.page > 1) {
      resultPaging.page -= 1;
      renderResultsView();
    }
  });
}

if (btnNextPage) {
  btnNextPage.addEventListener("click", () => {
    resultPaging.page += 1;
    renderResultsView();
  });
}

if (elCourseForm) {
  elCourseForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const courseCode = normalizeCourseCode(elCourseCode.value);
    try {
      if (activeCourseEditCode && activeCourseEditCode !== courseCode) {
        const detail = courseCoverageDetails.get(activeCourseEditCode);
        if (detail && detail.taken.length > 0) {
          throw new Error("Cannot change course code when results exist. Create a new course instead.");
        }
      }
      await upsertCourse({
        courseCode,
        title: elCourseTitle.value,
        credits: elCourseCredits.value,
        isMPUCourse: elCourseMPU.value,
      });
      console.log(`Course saved: ${courseCode}`);
      resetCourseForm();
      await refreshCourses();
      await refreshStats();
      await refreshResults();
    } catch (e) {
      console.error(e);
    }
  });
}

if (btnCourseReset) {
  btnCourseReset.addEventListener("click", () => {
    resetCourseForm();
  });
}

Promise.resolve().then(async () => {
  try {
    initImports({
      onDataChanged: async () => {
        await refreshStats();
        await refreshCourses();
        await refreshResults();
      },
    });
  } catch (e) {
    console.error(e);
  }
  try {
    initBackup({
      onDataChanged: async () => {
        await refreshStats();
        await refreshCourses();
        await refreshResults();
      },
    });
  } catch (e) {
    console.error(e);
  }
  await refreshStats();
  await refreshCourses();
  await refreshResults();
  setActiveResultSubtab("overview");
});
