// app/modules/course.js
import { deleteRecord, getAllRecords, upsertMany, STORES } from "../db/db.js";

function toNumberOrNull(value) {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function toBoolOrNull(value) {
  if (value === null || value === undefined) return null;
  if (value === true || value === false) return value;
  const s = String(value).trim().toLowerCase();
  if (s === "") return null;
  if (["true", "1", "yes", "y"].includes(s)) return true;
  if (["false", "0", "no", "n"].includes(s)) return false;
  return null;
}

export function normalizeCourseCode(value) {
  return String(value ?? "").trim().toUpperCase();
}

export function normalizeCourseInput(course) {
  if (!course) return null;
  const courseCode = normalizeCourseCode(course.courseCode ?? course.code ?? "");
  if (!courseCode) return null;
  return {
    courseCode,
    title: String(course.title ?? "").trim(),
    credits: toNumberOrNull(course.credits),
    isMPUCourse: toBoolOrNull(course.isMPUCourse),
  };
}

export async function upsertCourse(course) {
  const normalized = normalizeCourseInput(course);
  if (!normalized) throw new Error("Course code is required.");
  await upsertMany(STORES.courses, [normalized]);
  return normalized;
}

export async function listCourses() {
  const courses = await getAllRecords(STORES.courses);
  return courses
    .map((course) => normalizeCourseInput(course))
    .filter(Boolean)
    .sort((a, b) => a.courseCode.localeCompare(b.courseCode));
}

export async function deleteCourse(courseCode) {
  const code = normalizeCourseCode(courseCode);
  if (!code) throw new Error("Course code is required.");
  await deleteRecord(STORES.courses, code);
  return code;
}
