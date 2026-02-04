// app/modules/stats.js
import { normalizeCourseCode } from "./course.js";

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
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

export function computeStudentCgpa(results, courses) {
  const courseMap = new Map();
  for (const course of courses) {
    const code = normalizeCourseCode(course.courseCode ?? "");
    if (!code) continue;
    courseMap.set(code, course);
  }

  const totalsByStudent = new Map();
  for (const result of results) {
    const studentId = String(result.studentId ?? "").trim();
    if (!studentId) continue;
    const courseCode = normalizeCourseCode(result.courseCode ?? "");
    const course = courseMap.get(courseCode);
    const credits = toNumber(course?.credits);
    const gradePoints = getGradePoints(result, credits);
    if (credits === null || gradePoints === null) continue;
    const totals = totalsByStudent.get(studentId) ?? { credits: 0, gradePoints: 0 };
    totals.credits += credits;
    totals.gradePoints += gradePoints;
    totalsByStudent.set(studentId, totals);
  }

  const cgpaByStudent = new Map();
  for (const [studentId, totals] of totalsByStudent.entries()) {
    if (totals.credits > 0) {
      cgpaByStudent.set(studentId, totals.gradePoints / totals.credits);
    }
  }

  return cgpaByStudent;
}
