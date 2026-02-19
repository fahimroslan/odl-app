// app/modules/enrollment-optimizer.js
function intakeToRank(intake) {
  const raw = String(intake ?? "").trim();
  const match = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (!match) return Number.MAX_SAFE_INTEGER;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month)) return Number.MAX_SAFE_INTEGER;
  return year * 12 + month;
}

class MinCostMaxFlow {
  constructor(n) {
    this.n = n;
    this.adj = Array.from({ length: n }, () => []);
  }

  addEdge(u, v, cap, cost) {
    const fwd = { v, cap, cost, rev: this.adj[v].length };
    const rev = { v: u, cap: 0, cost: -cost, rev: this.adj[u].length };
    this.adj[u].push(fwd);
    this.adj[v].push(rev);
  }

  minCostMaxFlow(source, sink, maxFlow = Infinity) {
    let flow = 0;
    let cost = 0;
    const n = this.n;
    const potential = new Array(n).fill(0);
    const dist = new Array(n);
    const prevV = new Array(n);
    const prevE = new Array(n);

    while (flow < maxFlow) {
      dist.fill(Infinity);
      dist[source] = 0;
      const visited = new Array(n).fill(false);

      for (;;) {
        let v = -1;
        for (let i = 0; i < n; i += 1) {
          if (!visited[i] && (v === -1 || dist[i] < dist[v])) v = i;
        }
        if (v === -1 || dist[v] === Infinity) break;
        visited[v] = true;
        for (let i = 0; i < this.adj[v].length; i += 1) {
          const e = this.adj[v][i];
          if (e.cap <= 0) continue;
          const nd = dist[v] + e.cost + potential[v] - potential[e.v];
          if (nd < dist[e.v]) {
            dist[e.v] = nd;
            prevV[e.v] = v;
            prevE[e.v] = i;
          }
        }
      }

      if (dist[sink] === Infinity) break;
      for (let i = 0; i < n; i += 1) {
        if (dist[i] < Infinity) potential[i] += dist[i];
      }

      let add = maxFlow - flow;
      for (let v = sink; v !== source; v = prevV[v]) {
        const e = this.adj[prevV[v]][prevE[v]];
        add = Math.min(add, e.cap);
      }

      for (let v = sink; v !== source; v = prevV[v]) {
        const e = this.adj[prevV[v]][prevE[v]];
        e.cap -= add;
        this.adj[v][e.rev].cap += add;
        cost += add * e.cost;
      }

      flow += add;
    }

    return { flow, cost };
  }
}

function assignGreedy({
  students,
  offeredCourseCodes,
  progressByStudent,
  maxPerStudent,
  maxPerStudentById,
  priorityCourseCodes,
  priorityForAll,
  studentRepeatCourses,
  repeatWeight,
  totalSlots,
}) {
  const assignments = new Map();
  for (const student of students) {
    assignments.set(String(student?.studentId ?? "").trim(), []);
  }

  if (!students.length || !offeredCourseCodes.length || totalSlots <= 0) return assignments;

  const courseCodes = offeredCourseCodes.slice();
  const prioritySet = new Set((priorityCourseCodes ?? []).map((code) => String(code ?? "").trim()));
  const courseCount = courseCodes.length;
  const base = Math.floor(totalSlots / courseCount);
  const extra = totalSlots % courseCount;
  const capacities = new Map();
  courseCodes.forEach((code, idx) => {
    let cap = base + (idx < extra ? 1 : 0);
    if (base === 0) cap = 1;
    capacities.set(code, cap);
  });

  const load = new Map(courseCodes.map((code) => [code, 0]));
  const usePriorityForAll = Boolean(priorityForAll) && prioritySet.size > 0;
  const sortedStudents = students.slice().sort((a, b) => {
    if (usePriorityForAll) {
      return String(a?.studentId ?? "").localeCompare(String(b?.studentId ?? ""));
    }
    return intakeToRank(a?.intake) - intakeToRank(b?.intake);
  });

  for (const student of sortedStudents) {
    const studentId = String(student?.studentId ?? "").trim();
    if (!studentId) continue;
    const progress = progressByStudent.get(studentId) ?? { passed: new Set() };
    const eligible = courseCodes.filter((code) => !progress.passed?.has(code));
    const overrideMax = maxPerStudentById?.get(studentId);
    const studentMax = Number.isFinite(overrideMax) ? Math.max(0, overrideMax) : maxPerStudent;
    const slots = Math.min(studentMax, eligible.length);
    if (!slots) continue;

    const repeats = studentRepeatCourses?.get(studentId) ?? new Set();
    const chosen = [];
    for (let i = 0; i < slots; i += 1) {
      const available = eligible
        .filter((code) => (capacities.get(code) ?? 0) > (load.get(code) ?? 0) && !chosen.includes(code))
        .sort((a, b) => {
          const ra = repeats.has(a) ? 0 : 1;
          const rb = repeats.has(b) ? 0 : 1;
          if (ra !== rb) return ra - rb;
          const pa = prioritySet.has(a) ? 0 : 1;
          const pb = prioritySet.has(b) ? 0 : 1;
          if (pa !== pb) return pa - pb;
          const la = load.get(a) ?? 0;
          const lb = load.get(b) ?? 0;
          if (la !== lb) return la - lb;
          return String(a).localeCompare(String(b));
        });
      if (!available.length) break;
      const pick = available[0];
      chosen.push(pick);
      load.set(pick, (load.get(pick) ?? 0) + 1);
    }

    assignments.set(studentId, chosen);
  }

  return assignments;
}

export function assignBalancedCourses({
  students,
  offeredCourseCodes,
  progressByStudent,
  maxPerStudent = 6,
  maxPerStudentById = null,
  mode = "auto",
  maxSlotsForMcmf = 1200,
  priorityCourseCodes = [],
  priorityWeight = 50,
  priorityForAll = true,
  studentRepeatCourses = new Map(),
  repeatWeight = 200,
}) {
  const assignments = new Map();
  for (const student of students) {
    assignments.set(String(student?.studentId ?? "").trim(), []);
  }

  if (!students.length || !offeredCourseCodes.length) return assignments;

  const courseCodes = offeredCourseCodes.slice();
  const courseIndex = new Map(courseCodes.map((code, idx) => [code, idx]));

  const studentSlots = [];
  const eligibleByStudent = new Map();
  let totalSlots = 0;

  for (const student of students) {
    const studentId = String(student?.studentId ?? "").trim();
    const progress = progressByStudent.get(studentId) ?? { passed: new Set() };
    const eligible = courseCodes.filter((code) => !progress.passed?.has(code));
    eligibleByStudent.set(studentId, eligible);
    const overrideMax = maxPerStudentById?.get(studentId);
    const studentMax = Number.isFinite(overrideMax) ? Math.max(0, overrideMax) : maxPerStudent;
    const slots = Math.min(studentMax, eligible.length);
    studentSlots.push(slots);
    totalSlots += slots;
  }

  if (!totalSlots) return assignments;

  const courseCount = courseCodes.length;
  const base = Math.floor(totalSlots / courseCount);
  const extra = totalSlots % courseCount;

  if (mode === "greedy" || (mode === "auto" && totalSlots > maxSlotsForMcmf)) {
    return assignGreedy({
      students,
      offeredCourseCodes,
      progressByStudent,
      maxPerStudent,
      maxPerStudentById,
      priorityCourseCodes,
      priorityForAll,
      studentRepeatCourses,
      repeatWeight,
      totalSlots,
    });
  }

  const nodeCount = 2 + students.length + courseCount + 1;
  const source = 0;
  const sink = 1;
  const studentOffset = 2;
  const courseOffset = studentOffset + students.length;
  const dummyNode = courseOffset + courseCount;

  const mcmf = new MinCostMaxFlow(nodeCount);

  students.forEach((student, i) => {
    const slots = studentSlots[i];
    if (slots <= 0) return;
    mcmf.addEdge(source, studentOffset + i, slots, 0);
  });

  courseCodes.forEach((_, j) => {
    let cap = base + (j < extra ? 1 : 0);
    if (base === 0) cap = 1;
    mcmf.addEdge(courseOffset + j, sink, cap, 0);
  });

  mcmf.addEdge(dummyNode, sink, totalSlots, 0);

  const prioritySet = new Set(
    (priorityCourseCodes ?? []).map((code) => String(code ?? "").trim()).filter(Boolean)
  );

  students.forEach((student, i) => {
    const studentId = String(student?.studentId ?? "").trim();
    const eligible = eligibleByStudent.get(studentId) ?? [];
    const rank = intakeToRank(student?.intake);
    const baseCost = Number.isFinite(rank) ? rank * 100 : 9_999_900;
    const repeatSet = studentRepeatCourses?.get(studentId) ?? new Set();

    for (const code of eligible) {
      const idx = courseIndex.get(code);
      if (idx === undefined) continue;
      const tieBreaker = idx;
      const isPriority = prioritySet.has(code);
      const priorityBias = isPriority ? -priorityWeight : 0;
      const edgeBaseCost = priorityForAll && isPriority ? 0 : baseCost;
      const repeatBias = repeatSet.has(code) ? -repeatWeight : 0;
      mcmf.addEdge(
        studentOffset + i,
        courseOffset + idx,
        1,
        edgeBaseCost + tieBreaker + priorityBias + repeatBias
      );
    }

    const slots = studentSlots[i];
    if (slots > 0) {
      mcmf.addEdge(studentOffset + i, dummyNode, slots, baseCost + 9_999);
    }
  });

  mcmf.minCostMaxFlow(source, sink, totalSlots);

  students.forEach((student, i) => {
    const studentId = String(student?.studentId ?? "").trim();
    if (!studentId) return;
    const edges = mcmf.adj[studentOffset + i];
    const assigned = [];
    for (const e of edges) {
      if (e.v >= courseOffset && e.v < courseOffset + courseCount && e.cap === 0) {
        const code = courseCodes[e.v - courseOffset];
        if (code) assigned.push(code);
      }
    }
    assignments.set(studentId, assigned);
  });

  return assignments;
}
