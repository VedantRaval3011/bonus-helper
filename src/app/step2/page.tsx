"use client";

import React from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import { useState } from "react";
import ExcelJS from "exceljs";

interface CellTriplet {
  A: number;
  B: number;
  C: number;
  D: number;
  E: number;
}

interface ReportData {
  months: string[];
  departments: { name: string; data: Record<string, CellTriplet> }[];
}

interface EmployeeMonthlySalary {
  name: string;
  employeeCode: string;
  department: string;
  monthlySalaries: Record<string, number>;
  monthlyDepartments?: Record<string, string>;
  source: string;
}

interface EmployeeComparison {
  name: string;
  employeeCode: string;
  department: string;
  actualSalaries: Record<string, number>;
  hrSalaries: Record<string, number>;
  hasMismatch: boolean;
  missingInHR: boolean;
  missingInActual: boolean;
  totalActual: number;
  totalHR: number;
  totalDifference: number;
  monthsWithMismatch: string[];
  monthlyDepartments?: Record<string, string>;
}

interface ComparisonResults {
  staffComparisons: EmployeeComparison[];
  workerComparisons: EmployeeComparison[];
}

export default function Step2Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [reportData, setReportData] = useState<ReportData | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isComparing, setIsComparing] = useState(false);
  const [comparisonResults, setComparisonResults] =
    useState<ComparisonResults | null>(null);
  const [activeTab, setActiveTab] = useState<"staff" | "worker">("staff");
  const [isLoadingTab, setIsLoadingTab] = useState(false);
  const extras: Record<string, CellTriplet> = {};
  const [ignoredEmployees, setIgnoredEmployees] = useState<Set<string>>(
    new Set()
  );
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [password, setPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [actualPercentageFile, setActualPercentageFile] = useState<File | null>(
    null
  );

  // In your Step2 component file
  // 1) helper to post messages
  async function postAuditMessages(
    items: any[],
    batchId?: string,
    step?: number
  ) {
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId, step, items }),
    });
  }
  // 2) build metric snapshots (A,B,C,D,E) for all months from reportData
  function buildMetricSnapshotMessages(reportData: ReportData) {
    // Expecting rows: Software, HR, A,N,C (extras), diff
    const software =
      reportData.departments.find((d) => d.name === "Software")?.data || {};
    const hr = reportData.departments.find((d) => d.name === "HR")?.data || {};
    const extras =
      reportData.departments.find((d) => d.name === "A,N,C")?.data || {};
    const diff =
      reportData.departments.find((d) => d.name === "diff")?.data || {};

    const metrics = ["A", "B", "C", "D", "E"] as const;

    const items = [];
    for (const m of reportData.months) {
      const snapshot: Record<string, any> = {};
      for (const k of metrics) {
        snapshot[k] = {
          software: software[m]?.[k] ?? null,
          hr: hr[m]?.[k] ?? null,
          diff: diff[m]?.[k] ?? null,
          extras: k === "B" ? extras[m]?.[k] ?? null : null, // only B has extras
        };
      }
      items.push({
        level: "info",
        tag: "metric-snapshot",
        text: `Metric snapshot for ${m}`,
        scope: "global",
        source: "step2",
        meta: { month: m, snapshot },
      });
    }
    return items;
  }

  // 3) build mismatch messages (only if comparisonResults is present)
  function buildMismatchMessages(comparisonResults: ComparisonResults | null) {
    if (!comparisonResults) return [];
    const months = generateMonthHeaders(); // you already have this helper

    const items: any[] = [];
    const buildFor = (
      list: EmployeeComparison[],
      scope: "staff" | "worker"
    ) => {
      for (const emp of list) {
        // missing in HR / missing in Actual
        if (emp.missingInHR) {
          items.push({
            level: "warning",
            tag: "missing-in-hr",
            text: `Missing in HR: ${emp.employeeCode} ${emp.name} (${emp.department})`,
            scope,
            source: "step2",
            meta: {
              employeeCode: emp.employeeCode,
              name: emp.name,
              department: emp.department,
            },
          });
        }
        if (emp.missingInActual) {
          items.push({
            level: "warning",
            tag: "missing-in-actual",
            text: `Missing in Actual: ${emp.employeeCode} ${emp.name} (${emp.department})`,
            scope,
            source: "step2",
            meta: {
              employeeCode: emp.employeeCode,
              name: emp.name,
              department: emp.department,
            },
          });
        }

        // per-month mismatches (abs diff >= 1) excluding ignorable months
        for (const m of months) {
          const actual = emp.actualSalaries[m] ?? 0;
          const hr = emp.hrSalaries[m] ?? 0;
          const diff = actual - hr;
          const hasDiff = Math.abs(diff) >= 1;
          const monthDept = emp.monthlyDepartments?.[m] || emp.department;
          const shouldIgnore =
            (monthDept || "").toString().toUpperCase() === "C" ||
            (monthDept || "").toString().toUpperCase() === "A" ||
            (emp.employeeCode || "").toString().toUpperCase() === "N";

          if (hasDiff && !shouldIgnore) {
            items.push({
              level: "error",
              tag: "mismatch",
              text: `[${scope}] ${emp.employeeCode} ${emp.name} ${m} diff=${diff} (actual=${actual}, hr=${hr})`,
              scope,
              source: "step2",
              meta: {
                employeeCode: emp.employeeCode,
                name: emp.name,
                department: emp.department,
                month: m,
                actual,
                hr,
                diff,
              },
            });
          }
        }
      }
    };

    buildFor(comparisonResults.staffComparisons, "staff");
    buildFor(comparisonResults.workerComparisons, "worker");
    return items;
  }

  async function storeAuditForGenerateReport(
    reportData: ReportData,
    comparisonResults: ComparisonResults | null
  ) {
    const batchId =
      crypto.randomUUID?.() || Math.random().toString(36).slice(2);
    const items = [
      ...buildMetricSnapshotMessages(reportData),
      ...buildMismatchMessages(comparisonResults),
    ];
    if (items.length > 0) {
      await postAuditMessages(items, batchId, 2); // step 2
    }
  }

 const calculateMonthlyDifferences = (): {
  canProceed: boolean;
  monthlyStats: Record<string, { totalDiff: number; employeeCount: number }>;
} => {
  if (!comparisonResults) return { canProceed: true, monthlyStats: {} };

  const months = generateMonthHeaders();
  const monthlyStats: Record<
  string,
    { totalDiff: number; employeeCount: number }
  > = {};
  let canProceed = true;

  months.forEach((month) => {
    let totalDiff = 0;
    let employeeCount = 0;

    // Calculate for staff
    comparisonResults.staffComparisons.forEach((emp) => {
      const actualSal = emp.actualSalaries[month] || 0;
      const hrSal = emp.hrSalaries[month] || 0;
      const diff = Math.abs(actualSal - hrSal);

      if (actualSal > 0 || hrSal > 0) {
        employeeCount++;
        totalDiff += diff;
      }
    });

    // Calculate for workers
    comparisonResults.workerComparisons.forEach((emp) => {
      const aRaw = emp.actualSalaries[month];
      const hRaw = emp.hrSalaries[month];
      const hasA = aRaw !== undefined && !isNaN(aRaw as number);
      const hasH = hRaw !== undefined && !isNaN(hRaw as number);
      const a = hasA ? (aRaw as number) : 0;
      const h = hasH ? (hRaw as number) : 0;
      if (hasA || hasH) {
        employeeCount++;
        totalDiff += Math.abs(a - h);
      }
    });

    monthlyStats[month] = { totalDiff, employeeCount };

    // ✅ NEW: Only block if absolute difference exceeds employee count
    // This allows for negative differences (like -41) as long as |diff| < employeeCount
    if (Math.abs(totalDiff) > employeeCount) {
      canProceed = false;
      console.log(`⚠️ Month ${month}: |diff| (${Math.abs(totalDiff)}) > employees (${employeeCount})`);
    } else {
      console.log(`✅ Month ${month}: |diff| (${Math.abs(totalDiff)}) <= employees (${employeeCount})`);
    }
  });

  return { canProceed, monthlyStats };
};

  const getCellNumericValue = (
    cell: ExcelJS.Cell
  ): { hasValue: boolean; value: number } => {
    const v: any = cell.value;

    // Null/undefined are blank
    if (v === null || v === undefined) {
      return { hasValue: false, value: 0 };
    }

    // ✅ Handle formula objects (both regular and shared formulas)
    if (typeof v === "object" && v !== null) {
      // Check if it has a result field
      if ("result" in v) {
        const res = v.result;

        if (res !== null && res !== undefined) {
          if (typeof res === "number") {
            return { hasValue: true, value: res };
          }

          if (typeof res === "string") {
            const s = String(res).trim();
            if (!s || s === "-") {
              return { hasValue: false, value: 0 };
            }
            const n = Number(s.replace(/,/g, ""));
            if (Number.isFinite(n)) {
              return { hasValue: true, value: n };
            }
          }
        }
      }

      // ✅ For shared formulas without results, check the cell's master formula
      // If sharedFormula exists but no result, try cell.text
      if ("sharedFormula" in v && !("result" in v)) {
        // Fall through to cell.text check below
      } else {
        // Other object types without valid result = no value
        return { hasValue: false, value: 0 };
      }
    }

    // Strings
    if (typeof v === "string") {
      const s = v.trim();
      if (!s || s === "-") {
        return { hasValue: false, value: 0 };
      }
      const n = Number(s.replace(/,/g, ""));
      return Number.isFinite(n)
        ? { hasValue: true, value: n }
        : { hasValue: false, value: 0 };
    }

    // Direct numbers
    if (typeof v === "number") {
      return { hasValue: true, value: v };
    }

    // ✅ Last resort: cell.text (for shared formulas without cached results)
    const t = cell.text?.trim();

    // ✅ SPECIAL CASE: If cell.text is empty but we know it's a formula cell,
    // treat it as 0 (value exists, just equals zero)
    if (typeof v === "object" && "sharedFormula" in v && (!t || t === "")) {
      return { hasValue: true, value: 0 }; // ✅ This fixes April!
    }

    if (!t || t === "-") {
      return { hasValue: false, value: 0 };
    }

    const n = Number(t.replace(/,/g, ""));
    return Number.isFinite(n)
      ? { hasValue: true, value: n }
      : { hasValue: false, value: 0 };
  };

const handleMoveToStep3 = () => {
  if (!reportData || !comparisonResults) {
    router.push("/step3");
    return;
  }

  let requirePassword = false;
  const months = reportData.months;
  const diffDept = reportData.departments.find(d => d.name === 'diff');
  
  if (diffDept) {
    for (const month of months) {
      const monthData = diffDept.data[month];
      if (!monthData) continue;

      // Count employees for this month
      let employeeCount = 0;
      comparisonResults.staffComparisons.forEach((emp) => {
        if ((emp.actualSalaries[month] || 0) > 0 || (emp.hrSalaries[month] || 0) > 0) {
          employeeCount++;
        }
      });
      comparisonResults.workerComparisons.forEach((emp) => {
        const aRaw = emp.actualSalaries[month];
        const hRaw = emp.hrSalaries[month];
        if ((aRaw !== undefined && !isNaN(aRaw as number)) || 
            (hRaw !== undefined && !isNaN(hRaw as number))) {
          employeeCount++;
        }
      });

      // Check each metric individually
      const metrics = ['A', 'B', 'C', 'D', 'E'] as const;
      for (const metric of metrics) {
        const diff = Math.abs(monthData[metric] || 0);
        if (diff > employeeCount) {
          console.log(`⚠️ ${month} ${metric}: |diff| (${diff}) > employees (${employeeCount})`);
          requirePassword = true;
          break;
        }
      }
      
      if (requirePassword) break;
    }
  }

  if (requirePassword) {
    setShowPasswordModal(true);
    setPassword("");
    setPasswordError("");
  } else {
    router.push("/step3");
  }
};

  const handlePasswordSubmit = () => {
    const expectedPassword = process.env.NEXT_PUBLIC_NEXT_PASSWORD;

    if (password === expectedPassword) {
      setShowPasswordModal(false);
      router.push("/step3");
    } else {
      setPasswordError("Incorrect password. Please try again.");
    }
  };

  const handleModalClose = () => {
    setShowPasswordModal(false);
    setPassword("");
    setPasswordError("");
  };

  const handleTabSwitch = async (tab: "staff" | "worker") => {
    setIsLoadingTab(true);
    // Simulate async loading (you can adjust timing)
    await new Promise((resolve) => setTimeout(resolve, 300));
    setActiveTab(tab);
    setIsLoadingTab(false);
  };

  type FileSlot = { type: string; file: File | null };

  const pickFile = (pred: (s: FileSlot) => boolean): File | null => {
    const slot = (fileSlots as FileSlot[]).find(pred);
    return slot?.file ?? null;
  };

  const staffFile =
    pickFile((s) => s.type === "Indiana-Staff") ??
    pickFile((s) => !!s.file && /staff/i.test(s.file.name));

  const workerFile =
    pickFile((s) => s.type === "Indiana-Worker") ??
    pickFile((s) => !!s.file && /worker/i.test(s.file.name));

  const monthWiseFile =
    pickFile((s) => s.type === "Month-Wise-Sheet") ??
    pickFile((s) => !!s.file && /month.*wise/i.test(s.file.name));

  const bonusFile =
    pickFile((s) => s.type === "Bonus-Calculation-Sheet") ??
    pickFile(
      (s) =>
        !!s.file &&
        /bonus.*final.*calculation|bonus.*2024-25/i.test(s.file.name)
    );

  const bonusSummaryFile =
    pickFile((s) => s.type === "Bonus-Summery") ??
    pickFile((s) => !!s.file && /bonus.*summ?ery/i.test(s.file.name));

  const generateMonthHeaders = () => {
    const months: string[] = [];
    const today = new Date();
    const startDate = new Date(today.getFullYear() - 1, 10, 1);
    for (let i = 0; i < 11; i++) {
      const d = new Date(startDate);
      d.setMonth(startDate.getMonth() + i);
      const m = d.toLocaleDateString("en-US", { month: "short" });
      const y = d.getFullYear().toString().slice(-2);
      months.push(`${m}-${y}`);
    }
    months.push("Oct-25");
    return months;
  };

  const sheetNameToMonthKey = (sheetName: string): string | null => {
    const m = sheetName
      .toUpperCase()
      .match(/(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JUL|JULY|AUG|SEP)-(\d{2})/);
    if (!m) return null;
    const mm = m[1];
    const yy = m[2];
    const map: Record<string, string> = {
      NOV: "Nov",
      DEC: "Dec",
      JAN: "Jan",
      FEB: "Feb",
      MAR: "Mar",
      APR: "Apr",
      MAY: "May",
      JUN: "Jun",
      JUL: "Jul",
      JULY: "Jul",
      AUG: "Aug",
      SEP: "Sep",
    };
    return `${map[mm]}-${yy}`;
  };

  // Helper function to check if employee should be ignored
  const shouldIgnoreEmployee = (
    emp: EmployeeComparison,
    months: string[]
  ): { shouldIgnore: boolean; validMonths: string[] } => {
    // Check if employee has departments C, A, or employee ID is 'N'
    const hasIgnorableDept =
      ["C", "A"].includes(emp.department.toUpperCase()) ||
      emp.employeeCode.toUpperCase() === "N";

    if (!hasIgnorableDept) {
      return { shouldIgnore: false, validMonths: [] };
    }

    // Track which months the employee was in ignorable departments
    // We need to extract department info from monthly data
    // For now, if they are currently in C/A/N, mark them as ignorable
    return { shouldIgnore: true, validMonths: months };
  };

  // Function to toggle ignore for C, A, N employees
  const toggleIgnoreSpecialDepts = () => {
    if (!comparisonResults) return;

    const specialEmployees = new Set<string>();

    // Check both staff and worker
    [
      ...comparisonResults.staffComparisons,
      ...comparisonResults.workerComparisons,
    ].forEach((emp) => {
      if (
        ["C", "A"].includes(emp.department.toUpperCase()) ||
        emp.employeeCode.toUpperCase() === "N"
      ) {
        specialEmployees.add(emp.employeeCode || emp.name);
      }
    });

    if (ignoredEmployees.size > 0) {
      // Clear ignored
      setIgnoredEmployees(new Set());
    } else {
      // Set ignored
      setIgnoredEmployees(specialEmployees);
    }
  };

  const num = (cell: ExcelJS.Cell): number => {
    const v: any = cell.value;
    if (typeof v === "number") return v;
    if (v && typeof v === "object" && "result" in v)
      return Number(v.result) || 0;
    const t = cell.text?.replace(/,/g, "") || "";
    const n = Number(t);
    return Number.isFinite(n) ? n : 0;
  };

  const round = (x: number) => Math.round(x || 0);

  const findRowByLabel = (
    ws: ExcelJS.Worksheet,
    predicate: (t: string) => boolean
  ): number => {
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 12; c++) {
        const t = row.getCell(c).value?.toString().toUpperCase() || "";
        if (predicate(t)) return r;
      }
    }
    return -1;
  };

  const findColByHeader = (
    ws: ExcelJS.Worksheet,
    must: string[],
    any: string[] = [],
    rowScan = 5
  ): { col: number; headerRow: number } => {
    for (let r = 1; r <= rowScan; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= row.cellCount; c++) {
        const t = row.getCell(c).value?.toString().toUpperCase() || "";
        if (
          must.every((k) => t.includes(k.toUpperCase())) &&
          (any.length === 0 || any.some((k) => t.includes(k.toUpperCase())))
        ) {
          return { col: c, headerRow: r };
        }
      }
    }
    return { col: -1, headerRow: -1 };
  };

  const getGrossSalaryGrandTotal = (ws: ExcelJS.Worksheet): number => {
    let gsCol = -1,
      headerRow = -1;
    for (let r = 1; r <= 5 && gsCol < 0; r++) {
      ws.getRow(r).eachCell((cell, c) => {
        const t = cell.value?.toString().toUpperCase() || "";
        if (t.includes("GROSS") && t.includes("SALARY")) {
          gsCol = c;
          headerRow = r;
        }
      });
    }
    if (gsCol < 0) {
      gsCol = 9;
      headerRow = 2;
    }

    let gtRow = -1;
    for (let r = 1; r <= ws.rowCount && gtRow < 0; r++) {
      for (let c = 1; c <= 12; c++) {
        const t = ws.getRow(r).getCell(c).value?.toString().toUpperCase() || "";
        if (t.includes("GRAND") && t.includes("TOTAL")) {
          gtRow = r;
          break;
        }
      }
    }

    if (gtRow > 0) {
      const v = num(ws.getRow(gtRow).getCell(gsCol));
      if (v) return v;
    }
    let sum = 0;
    const endRow = gtRow > 0 ? gtRow - 1 : ws.rowCount;
    for (let r = headerRow + 1; r <= endRow; r++)
      sum += num(ws.getRow(r).getCell(gsCol));
    return sum;
  };

  const readColumnGrandTotal = (
    ws: ExcelJS.Worksheet,
    targetCol: number,
    headerRow: number
  ): number => {
    if (targetCol <= 0) return 0;
    const gtRow = findRowByLabel(
      ws,
      (t) => t.includes("GRAND") && t.includes("TOTAL")
    );
    if (gtRow > 0) {
      const totalCell = ws.getRow(gtRow).getCell(targetCol);
      const direct = num(totalCell);
      if (direct) return direct;
      let sum = 0;
      for (let r = headerRow + 1; r < gtRow; r++)
        sum += num(ws.getRow(r).getCell(targetCol));
      return sum;
    }
    let sum = 0;
    for (let r = Math.max(2, headerRow + 1); r <= ws.rowCount; r++)
      sum += num(ws.getRow(r).getCell(targetCol));
    return sum;
  };

  const sumWorkerSalary1 = (ws: ExcelJS.Worksheet): number => {
    let { col, headerRow } = findColByHeader(ws, ["SALARY", "1"]);
    if (col < 0) {
      col = 9;
      headerRow = 2;
    }
    let sum = 0,
      blanks = 0;
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const v = num(ws.getRow(r).getCell(col));
      if (v > 0) {
        sum += v;
        blanks = 0;
      } else if (++blanks > 8) break;
    }
    return sum;
  };

  const readStaffSalary1Total = (ws: ExcelJS.Worksheet): number => {
    const header3 = ws.getRow(3);
    let salary1Col = -1;
    header3.eachCell((cell, c) => {
      const t = cell.value?.toString().toUpperCase() || "";
      if (t.includes("SALARY") && t.includes("1")) salary1Col = c;
    });
    if (salary1Col <= 0) return 0;
    const totalRow = findRowByLabel(ws, (t) => t.includes("TOTAL"));
    if (totalRow > 0) return num(ws.getRow(totalRow).getCell(salary1Col));
    return 0;
  };

  const readStaffGrossTotal = (ws: ExcelJS.Worksheet): number => {
    let { col } = findColByHeader(ws, ["GROSS", "SALARY"]);
    if (col < 0) col = 18;
    const totalRow = findRowByLabel(ws, (t) => t.includes("TOTAL"));
    if (totalRow > 0) return num(ws.getRow(totalRow).getCell(col));
    return 0;
  };

  // Calculate B Extras - Sum of SALARY1 from departments C, A, N (and M) from Worker file
  const calculateBExtras = (
    workerWb: ExcelJS.Workbook,
    workerMap: Record<string, string>,
    monthKey: string
  ): number => {
    const ws = workerWb.getWorksheet(workerMap[monthKey]);
    if (!ws) return 0;

    // Find header row
    let headerRow = -1;
    for (let r = 1; r <= 6; r++) {
      const t = ws.getRow(r).getCell(2).value?.toString().toUpperCase() || "";
      if (t.includes("EMP") && t.includes("ID")) {
        headerRow = r;
        break;
      }
    }

    if (headerRow < 0) return 0;

    const empIdCol = 2;
    const deptCol = 3;
    const empNameCol = 4;
    const salary1Col = 9;

    let total = 0;
    const extrasDepts = ["C", "A", "N"]; // Include M as well

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const empId = row.getCell(empIdCol).value?.toString().trim() || "";
      const dept =
        row.getCell(deptCol).value?.toString().trim().toUpperCase() || "";
      const name =
        row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";

      if (!name || name.includes("TOTAL")) continue;

      const salary1 = num(row.getCell(salary1Col));

      // Include if department is C, A, M, or if EMP ID is 'N'
      if (extrasDepts.includes(dept) || empId.toUpperCase() === "N") {
        if (salary1 > 0) {
          total += salary1;
        }
      }
    }

    return total;
  };

  // D(HR) calculation - Average of GROSS SALARY from Bonus Summary file (Nov-24 to Sep-25)
  const getDHR = (bonusSummaryWb: ExcelJS.Workbook | null): number => {
    if (!bonusSummaryWb) {
      console.log("⚠️ D(HR): Bonus Summary file not provided");
      return 0;
    }

    // Get the first worksheet
    const ws =
      bonusSummaryWb.getWorksheet("Sheet1") ?? bonusSummaryWb.worksheets[0];
    if (!ws) {
      console.log("❌ D(HR): No worksheet found in Bonus Summary");
      return 0;
    }

    // Find header row (contains "GROSS SALARY")
    let headerRow = -1;
    let grossCol = -1;
    let monthCol = -1;

    for (let r = 1; r <= 5; r++) {
      const row = ws.getRow(r);
      row.eachCell((cell, c) => {
        const text = cell.value?.toString().toUpperCase() || "";
        if (text.includes("GROSS") && text.includes("SALARY")) {
          grossCol = c;
          headerRow = r;
        }
        if (text.includes("MONTH")) {
          monthCol = c;
        }
      });
      if (grossCol > 0) break;
    }

    if (grossCol < 0) {
      console.log("❌ D(HR): Could not find GROSS SALARY column");
      return 0;
    }
    if (monthCol < 0) monthCol = 2; // Default to column 2

    // Read GROSS SALARY values from Nov-24 to Sep-25
    const grossValues: number[] = [];
    const targetMonths = [10, 11, 0, 1, 2, 3, 4, 5, 6, 7, 8]; // Nov(10) to Sep(8), wrapping year

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const monthCell = ws.getRow(r).getCell(monthCol).value;
      const grossVal = num(ws.getRow(r).getCell(grossCol));

      if (monthCell instanceof Date && grossVal > 0) {
        const month = monthCell.getMonth();
        const year = monthCell.getFullYear();

        // Include Nov-2024 through Sep-2025 (skip Oct-2024 and Oct-2025)
        if (
          (year === 2024 && month >= 10) || // Nov-24, Dec-24
          (year === 2025 && month <= 8) // Jan-25 through Sep-25
        ) {
          grossValues.push(grossVal);
          console.log(
            `✅ D(HR) ${monthCell.toLocaleDateString("en-US", {
              month: "short",
              year: "2-digit",
            })}: ₹${grossVal.toLocaleString()}`
          );
        }
      }
    }

    if (grossValues.length === 0) {
      console.log("❌ D(HR): No valid GROSS SALARY values found");
      return 0;
    }

    const average =
      grossValues.reduce((sum, val) => sum + val, 0) / grossValues.length;

    console.log(`\n📊 D(HR) Calculation:`);
    console.log(`   Total months: ${grossValues.length}`);
    console.log(
      `   Sum: ₹${grossValues
        .reduce((sum, val) => sum + val, 0)
        .toLocaleString()}`
    );
    console.log(`   Average (D(HR)): ₹${average.toLocaleString()}`);

    return average;
  };

  // NEW: C(HR) calculation - Get GROSS SALARY average from Month-Wise Sheet across 11 months
  // C(HR) - Calculate per-employee October average from GROSS SALARY, then sum
  const getCHR = (
    monthWiseWb: ExcelJS.Workbook | null,
    workerWb: ExcelJS.Workbook,
    workerMap: Record<string, string>
  ): number => {
    const monthsToAverage = [
      "Nov-24",
      "Dec-24",
      "Jan-25",
      "Feb-25",
      "Mar-25",
      "Apr-25",
      "May-25",
      "Jun-25",
      "Jul-25",
      "Aug-25",
      "Sep-25",
    ];

    // Store each employee's monthly GROSS SALARY values
    const employees: Record<string, number[]> = {};

    for (const monthKey of monthsToAverage) {
      const monthSheet = (monthWiseWb || workerWb).getWorksheet(
        workerMap[monthKey]
      );
      if (!monthSheet) continue;

      // Find header row and columns
      let headerRow = -1;
      let grossSalaryCol = -1;
      let empIdCol = -1;
      let empNameCol = -1;

      for (let r = 1; r <= 6; r++) {
        const row = monthSheet.getRow(r);
        const cell2Text = row.getCell(2).value?.toString().toUpperCase() || "";

        if (cell2Text.includes("EMP") && cell2Text.includes("ID")) {
          headerRow = r;
          empIdCol = 2;
          empNameCol = 4;
        }

        // Find GROSS SALARY column
        row.eachCell((cell, c) => {
          const text = cell.value?.toString().toUpperCase() || "";
          if (text.includes("GROSS") && text.includes("SALARY")) {
            grossSalaryCol = c;
          }
        });

        if (headerRow > 0 && grossSalaryCol > 0) break;
      }

      if (headerRow < 0 || grossSalaryCol < 0) continue;

      // Read each employee's GROSS SALARY for this month
      // In the getCHR function, find and replace:
      // Find and replace this section:
      for (let r = headerRow + 1; r <= monthSheet.rowCount; r++) {
        const row = monthSheet.getRow(r);
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";
        const empName =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";

        if (!empName || empName.includes("TOTAL")) continue;

        // ❌ OLD CODE:
        // const grossSalary = num(row.getCell(grossSalaryCol));

        // ✅ NEW CODE:
        const cellResult = getCellNumericValue(row.getCell(grossSalaryCol));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = [];
        }

        // Include the value if cell has any value (including 0)
        if (cellResult.hasValue) {
          employees[key].push(cellResult.value);
        }
      }
    }

    // Calculate October value for each employee, then sum
    let totalOctober = 0;
    let employeeCount = 0;

    for (const key in employees) {
      const values = employees[key];
      if (values.length > 0) {
        const average =
          values.reduce((sum, val) => sum + val, 0) / values.length;
        totalOctober += average;
        employeeCount++;
      }
    }

    console.log(`\n📊 C(HR) Employee-Level Calculation:`);
    console.log(`   Employees processed: ${employeeCount}`);
    console.log(`   Total October C(HR): ₹${totalOctober.toLocaleString()}`);

    return totalOctober;
  };

  // E(HR) calculation - Average of FD column from Bonus Summary file (Nov-24 to Sep-25)
  const getEHR = (bonusSummaryWb: ExcelJS.Workbook | null): number => {
    if (!bonusSummaryWb) {
      console.log("⚠️ E(HR): Bonus Summary file not provided");
      return 0;
    }

    // Get the first worksheet
    const ws =
      bonusSummaryWb.getWorksheet("Sheet1") ?? bonusSummaryWb.worksheets[0];
    if (!ws) {
      console.log("❌ E(HR): No worksheet found in Bonus Summary");
      return 0;
    }

    // Find header row (contains "FD")
    let headerRow = -1;
    let fdCol = -1;
    let monthCol = -1;

    for (let r = 1; r <= 5; r++) {
      const row = ws.getRow(r);
      row.eachCell((cell, c) => {
        const text = cell.value?.toString().toUpperCase() || "";
        if (text === "FD" || text.includes("FD")) {
          fdCol = c;
          headerRow = r;
        }
        if (text.includes("MONTH")) {
          monthCol = c;
        }
      });
      if (fdCol > 0) break;
    }

    if (fdCol < 0) {
      console.log("❌ E(HR): Could not find FD column");
      return 0;
    }
    if (monthCol < 0) monthCol = 2; // Default to column 2

    // Read FD values from Nov-24 to Sep-25
    const fdValues: number[] = [];

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const monthCell = ws.getRow(r).getCell(monthCol).value;
      const fdVal = num(ws.getRow(r).getCell(fdCol));

      if (monthCell instanceof Date && fdVal > 0) {
        const month = monthCell.getMonth();
        const year = monthCell.getFullYear();

        // Include Nov-2024 through Sep-2025 (skip Oct-2024 and Oct-2025)
        if (
          (year === 2024 && month >= 10) || // Nov-24, Dec-24
          (year === 2025 && month <= 8) // Jan-25 through Sep-25
        ) {
          fdValues.push(fdVal);
          console.log(
            `✅ E(HR) ${monthCell.toLocaleDateString("en-US", {
              month: "short",
              year: "2-digit",
            })}: ₹${fdVal.toLocaleString()}`
          );
        }
      }
    }

    if (fdValues.length === 0) {
      console.log("❌ E(HR): No valid FD values found");
      return 0;
    }

    const average =
      fdValues.reduce((sum, val) => sum + val, 0) / fdValues.length;

    console.log(`\n📊 E(HR) Calculation:`);
    console.log(`   Total months: ${fdValues.length}`);
    console.log(
      `   Sum of FD: ₹${fdValues
        .reduce((sum, val) => sum + val, 0)
        .toLocaleString()}`
    );
    console.log(`   Average (E(HR)): ₹${average.toLocaleString()}`);

    return average;
  };

  const getBonusMonthlyTotals = (bonusWb: ExcelJS.Workbook) => {
    const map: Record<string, number> = {
      "Nov-24": 6,
      "Dec-24": 7,
      "Jan-25": 8,
      "Feb-25": 9,
      "Mar-25": 10,
      "Apr-25": 11,
      "May-25": 12,
      "Jun-25": 13,
      "Jul-25": 14,
      "Aug-25": 15,
      "Sep-25": 16,
    };
    const out: Record<string, { worker: number; staff: number }> = {};
    const w = bonusWb.getWorksheet("Worker");
    const s = bonusWb.getWorksheet("Staff");
    const findGT = (ws?: ExcelJS.Worksheet) =>
      ws
        ? findRowByLabel(ws, (t) => t.includes("GRAND") && t.includes("TOTAL"))
        : -1;
    const wGT = findGT(w),
      sGT = findGT(s);
    for (const m of Object.keys(map)) {
      const col = map[m];
      out[m] = {
        worker: wGT > 0 && w ? num(w.getRow(wGT).getCell(col)) : 0,
        staff: sGT > 0 && s ? num(s.getRow(sGT).getCell(col)) : 0,
      };
    }
    return out;
  };

  const getBonusGrossTotals = (
    bonusWb: ExcelJS.Workbook
  ): { worker: number; staff: number } => {
    const out: { worker: number; staff: number } = { worker: 0, staff: 0 };
    const readGross = (ws?: ExcelJS.Worksheet) => {
      if (!ws) return 0;
      const gt = findRowByLabel(
        ws,
        (t) => t.includes("GRAND") && t.includes("TOTAL")
      );
      if (gt < 0) return 0;
      return num(ws.getRow(gt).getCell(18));
    };
    out.worker = readGross(bonusWb.getWorksheet("Worker"));
    out.staff = readGross(bonusWb.getWorksheet("Staff"));
    return out;
  };

  const getBonusSummeryMonthlyGross = (
    wb: ExcelJS.Workbook
  ): Record<string, number> => {
    const ws = wb.getWorksheet("Sheet1") ?? wb.worksheets[0];
    if (!ws) return {};

    let grossCol = -1,
      monthCol = -1,
      headerRow = -1;
    for (let r = 1; r <= 5 && grossCol < 0; r++) {
      ws.getRow(r).eachCell((cell, c) => {
        const t = cell.value?.toString().toUpperCase() || "";
        if (t.includes("GROSS") && t.includes("SALARY")) {
          grossCol = c;
          headerRow = r;
        }
        if (t.includes("MONTH")) {
          monthCol = c;
        }
      });
    }
    if (grossCol < 0) {
      grossCol = 3;
      headerRow = 1;
    }
    if (monthCol < 0) {
      monthCol = 2;
    }

    const out: Record<string, number> = {};
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const monthCell = ws.getRow(r).getCell(monthCol).value;
      const grossVal = num(ws.getRow(r).getCell(grossCol));

      if (monthCell && grossVal > 0) {
        let monthKey = "";
        if (monthCell instanceof Date) {
          const m = monthCell.toLocaleDateString("en-US", { month: "short" });
          const y = monthCell.getFullYear().toString().slice(-2);
          monthKey = `${m}-${y}`;
        } else if (typeof monthCell === "string" && monthCell.includes("-")) {
          const d = new Date(monthCell);
          const m = d.toLocaleDateString("en-US", { month: "short" });
          const y = d.getFullYear().toString().slice(-2);
          monthKey = `${m}-${y}`;
        }

        if (monthKey && monthKey !== "Oct-25") out[monthKey] = grossVal;
      }
    }
    return out;
  };

  const getBonusSummeryMonthlyFD = (
    wb: ExcelJS.Workbook
  ): Record<string, number> => {
    const ws = wb.getWorksheet("Sheet1") ?? wb.worksheets[0];
    if (!ws) return {};

    let fdCol = -1,
      monthCol = -1,
      headerRow = -1;
    for (let r = 1; r <= 5 && fdCol < 0; r++) {
      ws.getRow(r).eachCell((cell, c) => {
        const t = cell.value?.toString().toUpperCase() || "";
        if (t.includes("FD")) {
          fdCol = c;
          headerRow = r;
        }
        if (t.includes("MONTH")) {
          monthCol = c;
        }
      });
    }
    if (fdCol < 0) {
      fdCol = 5;
      headerRow = 1;
    }
    if (monthCol < 0) {
      monthCol = 2;
    }

    const out: Record<string, number> = {};
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const monthCell = ws.getRow(r).getCell(monthCol).value;
      const fdVal = num(ws.getRow(r).getCell(fdCol));

      if (monthCell && fdVal > 0) {
        let monthKey = "";
        if (monthCell instanceof Date) {
          const m = monthCell.toLocaleDateString("en-US", { month: "short" });
          const y = monthCell.getFullYear().toString().slice(-2);
          monthKey = `${m}-${y}`;
        } else if (typeof monthCell === "string" && monthCell.includes("-")) {
          const d = new Date(monthCell);
          const m = d.toLocaleDateString("en-US", { month: "short" });
          const y = d.getFullYear().toString().slice(-2);
          monthKey = `${m}-${y}`;
        }

        if (monthKey && monthKey !== "Oct-25") out[monthKey] = fdVal;
      }
    }
    return out;
  };

  const calculateOctoberAverageForWorker = (
    workerWb: ExcelJS.Workbook,
    workerMap: Record<string, string>
  ): number => {
    const monthsToAverage = [
      "Nov-24",
      "Dec-24",
      "Jan-25",
      "Feb-25",
      "Mar-25",
      "Apr-25",
      "May-25",
      "Jun-25",
      "Jul-25",
      "Aug-25",
      "Sep-25",
    ];

    const employees: Record<
      string,
      { salaries: number[]; hasSeptSalary: boolean }
    > = {};

    for (const monthKey of monthsToAverage) {
      const ws = workerWb.getWorksheet(workerMap[monthKey]);
      if (!ws) continue;

      let headerRow = -1;
      for (let r = 1; r <= 6; r++) {
        const t = ws.getRow(r).getCell(2).value?.toString().toUpperCase() || "";
        if (t.includes("EMP") && t.includes("ID")) {
          headerRow = r;
          break;
        }
      }

      if (headerRow < 0) continue;

      const empIdCol = 2;
      const empNameCol = 4;
      const salary1Col = 9;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const empName = row
          .getCell(empNameCol)
          .value?.toString()
          .trim()
          .toUpperCase();
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";

        if (!empName || empName === "" || empName.includes("TOTAL")) continue;

        const cellResult = getCellNumericValue(row.getCell(salary1Col));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = { salaries: [], hasSeptSalary: false };
        }

        // Include the value if cell has any value (including 0)
        if (cellResult.hasValue) {
          employees[key].salaries.push(cellResult.value);
        }

        // For September check, still use > 0 since we want employees with actual salary
        if (
          monthKey === "Sep-25" &&
          cellResult.hasValue &&
          cellResult.value > 0
        ) {
          employees[key].hasSeptSalary = true;
        }
      }
    }

    let totalOctober = 0;
    for (const key of Object.keys(employees)) {
      const emp = employees[key];

      if (emp.hasSeptSalary && emp.salaries.length > 0) {
        const average =
          emp.salaries.reduce((sum, val) => sum + val, 0) / emp.salaries.length;
        totalOctober += average;
      }
    }

    return totalOctober;
  };

  // NEW: Calculate Staff GROSS SALARY October average for C(Software)
  const calculateStaffGrossOctober = (
    staffWb: ExcelJS.Workbook,
    staffMap: Record<string, string>
  ): number => {
    const monthsToAverage = [
      "Nov-24",
      "Dec-24",
      "Jan-25",
      "Feb-25",
      "Mar-25",
      "Apr-25",
      "May-25",
      "Jun-25",
      "Jul-25",
      "Aug-25",
      "Sep-25",
    ];

    const employees: Record<
      string,
      {
        grossSalaries: number[];
        salary1Values: number[];
        hasSeptSalary1: boolean;
      }
    > = {};

    for (const monthKey of monthsToAverage) {
      const ws = staffWb.getWorksheet(staffMap[monthKey]);
      if (!ws) continue;

      let headerRow = -1;
      for (let r = 1; r <= 5; r++) {
        const row = ws.getRow(r);
        const cellText = row.getCell(2).value?.toString().toUpperCase() || "";
        if (cellText.includes("EMP") && cellText.includes("ID")) {
          headerRow = r;
          break;
        }
      }

      if (headerRow < 0) continue;

      const empIdCol = 2;
      const empNameCol = 5;
      const salary1Col = 15; // SALARY1 column
      const grossSalCol = 18; // GROSS SALARY column

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const empName = row
          .getCell(empNameCol)
          .value?.toString()
          .trim()
          .toUpperCase();
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";

        if (!empName || empName === "" || empName.includes("TOTAL")) continue;

        const salary1Result = getCellNumericValue(row.getCell(salary1Col));
        const grossSalResult = getCellNumericValue(row.getCell(grossSalCol));

        const key = `${empId}_${empName}`;
        if (!employees[key]) {
          employees[key] = {
            grossSalaries: [],
            salary1Values: [],
            hasSeptSalary1: false,
          };
        }

        // Track SALARY1 for Sept condition - only if cell has value (including 0)
        if (salary1Result.hasValue) {
          employees[key].salary1Values.push(salary1Result.value);
        }

        // Track GROSS SALARY for averaging - include zeros!
        if (grossSalResult.hasValue) {
          employees[key].grossSalaries.push(grossSalResult.value);
        }

        // Check September SALARY1 condition - check for actual value > 0
        if (
          monthKey === "Sep-25" &&
          salary1Result.hasValue &&
          salary1Result.value > 0
        ) {
          employees[key].hasSeptSalary1 = true;
        }
      }
    }

    let totalOctober = 0;
    for (const key of Object.keys(employees)) {
      const emp = employees[key];

      // Only include if employee has September SALARY1 > 0
      if (emp.hasSeptSalary1 && emp.grossSalaries.length > 0) {
        const average =
          emp.grossSalaries.reduce((sum, val) => sum + val, 0) /
          emp.grossSalaries.length;
        totalOctober += average;
      }
    }

    console.log(
      `✅ Staff GROSS SALARY October Total: ₹${totalOctober.toLocaleString()}`
    );
    return totalOctober;
  };

  const extractStaffEmployees = (
    wb: ExcelJS.Workbook,
    sheetNames: string[]
  ) => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();

    for (const sheetName of sheetNames) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;

      let headerRow = -1;
      for (let r = 1; r <= 5; r++) {
        const t = ws.getRow(r).getCell(2).value?.toString().toUpperCase() || "";
        if (t.includes("EMP") && t.includes("ID")) {
          headerRow = r;
          break;
        }
      }
      if (headerRow < 0) continue;

      const monthKey = sheetNameToMonthKey(sheetName);
      if (!monthKey) continue;

      const empIdCol = 2,
        deptCol = 3,
        empNameCol = 5,
        salary1Col = 15;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);

        const empName =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";
        if (!empName || empName.includes("TOTAL")) continue;

        const empId = row.getCell(empIdCol).value?.toString().trim() || "";
        const dept = (
          row.getCell(deptCol).value?.toString().trim() || ""
        ).toUpperCase();

        const key = empId || empName;

        // 🐛 DETAILED DEBUG: Check what's actually in the cell
        const cell = row.getCell(salary1Col);
        const cellValue = cell.value;
        const cellText = cell.text;

        if (empName.includes("SANJAY") && empName.includes("RATHOD")) {
          console.log(`\n🔍 ${monthKey} SANJAY RATHOD RAW DATA:`);
          console.log(`   cell.value =`, cellValue);
          console.log(`   cell.text =`, cellText);
          console.log(`   typeof cell.value =`, typeof cellValue);

          if (cellValue && typeof cellValue === "object") {
            console.log(`   cell.value.result =`, (cellValue as any).result);
            console.log(`   typeof result =`, typeof (cellValue as any).result);
          }
        }

        const s1 = getCellNumericValue(cell);

        if (!employees[key]) {
          employees[key] = {
            name: empName,
            employeeCode: empId || empName,
            department: dept,
            monthlySalaries: {},
            monthlyDepartments: {},
            source: "Staff",
          };
        }

        if (empName.includes("SANJAY") && empName.includes("RATHOD")) {
          console.log(
            `   getCellNumericValue returned: hasValue=${s1.hasValue}, value=${s1.value}`
          );
        }

        if (s1.hasValue) {
          employees[key].monthlySalaries[monthKey] = s1.value;

          if (empName.includes("SANJAY") && empName.includes("RATHOD")) {
            console.log(`   ✅ ${monthKey} INCLUDED: ${s1.value}`);
          }
        } else {
          if (empName.includes("SANJAY") && empName.includes("RATHOD")) {
            console.log(`   ❌ ${monthKey} SKIPPED\n`);
          }
        }

        if (dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }
      }
    }

    // Rest of the function remains the same...
    const monthsReversed = [...months].reverse();
    for (const key in employees) {
      const emp = employees[key];
      let finalDept = "";

      for (const m of monthsReversed) {
        const d = emp.monthlyDepartments?.[m]?.toUpperCase();
        if (d && !["C", "A"].includes(d)) {
          finalDept = d;
          break;
        }
      }
      if (!finalDept && emp.monthlyDepartments?.["Sep-25"]) {
        finalDept = emp.monthlyDepartments["Sep-25"].toUpperCase();
      }
      if (!finalDept) {
        for (const m of monthsReversed) {
          const d = emp.monthlyDepartments?.[m]?.toUpperCase();
          if (d) {
            finalDept = d;
            break;
          }
        }
      }

      emp.department = finalDept || emp.department?.toUpperCase();
      if (!emp.employeeCode) emp.employeeCode = key;
    }

    return employees;
  };
  const extractWorkerEmployees = (
    wb: ExcelJS.Workbook,
    sheetNames: string[]
  ): Record<string, EmployeeMonthlySalary> => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();

    // First pass: collect all data
    for (const sheetName of sheetNames) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;

      let headerRow = -1;
      for (let r = 1; r <= 6; r++) {
        const t = ws.getRow(r).getCell(2).value?.toString().toUpperCase() || "";
        if (t.includes("EMP") && t.includes("ID")) {
          headerRow = r;
          break;
        }
      }

      if (headerRow < 0) continue;

      const monthKey = sheetNameToMonthKey(sheetName);
      if (!monthKey) continue;

      const empIdCol = 2;
      const deptCol = 3;
      const empNameCol = 4;
      const salary1Col = 9;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";
        const id = row.getCell(empIdCol).value?.toString().trim() || "";
        const dept = row.getCell(deptCol).value?.toString().trim() || "";

        if (!name || name.includes("TOTAL")) continue;

        const s1 = num(row.getCell(salary1Col));
        const key = id || name;
        const res = getCellNumericValue(row.getCell(salary1Col));
        if (!employees[key]) {
          employees[key] = {
            name,
            employeeCode: id,
            department: "",
            monthlySalaries: {},
            monthlyDepartments: {},
            source: "Worker",
          };
        }

        if (res.hasValue) {
          employees[key].monthlySalaries[monthKey] = res.value; // includes explicit 0
        }
        if (monthKey && dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }

        // Track department for each month
        if (monthKey && dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }
      }
    }

    // Second pass: determine the best department for each employee
    // Priority: 1) Latest non-C/A department, 2) September department, 3) Any department
    for (const key in employees) {
      const emp = employees[key];
      let finalDept = "";

      // Get departments in reverse chronological order (Sep-25 to Nov-24)
      const monthsReversed = [...months].reverse();

      // First, try to find the latest non-C/A department
      for (const month of monthsReversed) {
        const dept = emp.monthlyDepartments?.[month];
        if (dept && !["C", "A"].includes(dept.toUpperCase())) {
          finalDept = dept;
          break;
        }
      }

      // If no non-C/A department found, use September department if available
      if (!finalDept && emp.monthlyDepartments?.["Sep-25"]) {
        finalDept = emp.monthlyDepartments["Sep-25"];
      }

      // If still no department, use any available department
      if (!finalDept) {
        for (const month of monthsReversed) {
          const dept = emp.monthlyDepartments?.[month];
          if (dept) {
            finalDept = dept;
            break;
          }
        }
      }

      emp.department = finalDept;
    }

    return employees;
  };

  const extractHREmployees = (
    bonusWb: ExcelJS.Workbook,
    sheetName: "Staff" | "Worker"
  ): Record<string, EmployeeMonthlySalary> => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();
    const monthColMap: Record<string, number> = {
      "Nov-24": 6,
      "Dec-24": 7,
      "Jan-25": 8,
      "Feb-25": 9,
      "Mar-25": 10,
      "Apr-25": 11,
      "May-25": 12,
      "Jun-25": 13,
      "Jul-25": 14,
      "Aug-25": 15,
      "Sep-25": 16,
      "Oct-25": 17,
    };

    const ws = bonusWb.getWorksheet(sheetName);
    if (!ws) return employees;

    let headerRow = -1;
    for (let r = 1; r <= 5; r++) {
      const row = ws.getRow(r);
      const cellText = row.getCell(2).value?.toString().toUpperCase();
      if (cellText?.includes("EMP") && cellText?.includes("CODE")) {
        headerRow = r;
        break;
      }
    }

    if (headerRow <= 0) return employees;

    const empCodeCol = 2;
    const deptCol = 3;
    const empNameCol = 4;

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const code = row.getCell(empCodeCol).value?.toString().trim() || ""; // ✅ Default to ""
      const name =
        row.getCell(empNameCol).value?.toString().trim().toUpperCase() || ""; // ✅ Default to ""
      const dept = row.getCell(deptCol).value?.toString().trim() || ""; // ✅ Default to ""

      // ✅ Skip if no name or code
      if (!name || !code || name === "GRAND TOTAL") continue;

      const key = `${code}|${name}`;

      if (!employees[key]) {
        employees[key] = {
          name,
          employeeCode: code, // ✅ Now guaranteed to be string
          department: dept, // ✅ Now guaranteed to be string
          monthlySalaries: {},
          source: "HR",
        };
        months.forEach((m) => (employees[key].monthlySalaries[m] = 0));
      }

      // ✅ SUM all values for the same employee instead of replacing
      for (const [mKey, colNum] of Object.entries(monthColMap)) {
        const v = num(row.getCell(colNum));
        if (v > 0) {
          // ✅ ADD to existing value instead of replacing
          employees[key].monthlySalaries[mKey] += v;
        }
      }
    }

    return employees;
  };

  const compareEmployees = (
    actualByKey: Record<string, EmployeeMonthlySalary>,
    hrByKey: Record<string, EmployeeMonthlySalary>,
    months: string[]
  ): EmployeeComparison[] => {
    const allKeys = new Set([
      ...Object.keys(actualByKey),
      ...Object.keys(hrByKey),
    ]);

    // ✅ NEW: Group employees by employee code to handle duplicates
    const groupedActual: Record<string, EmployeeMonthlySalary[]> = {};
    const groupedHR: Record<string, EmployeeMonthlySalary[]> = {};

    // Group actual employees by employee code
    for (const key of Object.keys(actualByKey)) {
      const emp = actualByKey[key];
      const code = emp.employeeCode;
      if (!groupedActual[code]) {
        groupedActual[code] = [];
      }
      groupedActual[code].push(emp);
    }

    // Group HR employees by employee code
    for (const key of Object.keys(hrByKey)) {
      const emp = hrByKey[key];
      const code = emp.employeeCode;
      if (!groupedHR[code]) {
        groupedHR[code] = [];
      }
      groupedHR[code].push(emp);
    }

    // Get all unique employee codes
    const allCodes = new Set([
      ...Object.keys(groupedActual),
      ...Object.keys(groupedHR),
    ]);

    const out: EmployeeComparison[] = [];

    for (const code of allCodes) {
      const actualEmps = groupedActual[code] || [];
      const hrEmps = groupedHR[code] || [];

      // Get employee info (use first record)
      const firstActual = actualEmps[0];
      const firstHR = hrEmps[0];

      const name = (firstActual?.name || firstHR?.name).toUpperCase();
      const dept = firstActual?.department || firstHR?.department;

      const actualSalaries: Record<string, number> = {};
      const hrSalaries: Record<string, number> = {};
      const monthsWithMismatch: string[] = [];
      let hasMismatch = false;

      for (const m of months) {
        // ✅ SUM all values for this employee code across all records
        let av = 0;
        for (const emp of actualEmps) {
          const salary = emp?.monthlySalaries?.[m];
          const monthDept = emp?.monthlyDepartments?.[m];
          if (
            monthDept &&
            monthDept.toUpperCase() === (dept || "").toUpperCase() &&
            salary !== undefined &&
            !isNaN(salary)
          ) {
            av += salary;
          }
        }

        // ✅ SUM all HR values for this employee code
        let hv = 0;
        for (const emp of hrEmps) {
          hv += emp?.monthlySalaries?.[m] || 0;
        }

        actualSalaries[m] = av;
        hrSalaries[m] = hv;

        // Check for mismatch
        const md = firstActual?.monthlyDepartments?.[m];
        const shouldIgnoreMonth =
          (md ? ["C", "A"].includes(md.toUpperCase()) : true) ||
          code.toUpperCase() === "N";

        if (Math.abs(av - hv) > 1 && !shouldIgnoreMonth) {
          hasMismatch = true;
          monthsWithMismatch.push(m);
        }
      }

      const totalActual = Object.values(actualSalaries).reduce(
        (s, v) => s + v,
        0
      );
      const totalHR = Object.values(hrSalaries).reduce((s, v) => s + v, 0);

      out.push({
        name,
        employeeCode: code,
        department: dept,
        actualSalaries,
        hrSalaries,
        hasMismatch,
        missingInHR: hrEmps.length === 0,
        missingInActual: actualEmps.length === 0,
        totalActual,
        totalHR,
        totalDifference: totalActual - totalHR,
        monthsWithMismatch,
        monthlyDepartments: firstActual?.monthlyDepartments,
      });
    }

    // Sort logic
    out.sort((x, y) => {
      if (x.missingInHR !== y.missingInHR) return x.missingInHR ? -1 : 1;
      if (x.missingInActual !== y.missingInActual)
        return x.missingInActual ? -1 : 1;
      if (x.hasMismatch !== y.hasMismatch) return x.hasMismatch ? -1 : 1;
      return 0;
    });

    return out;
  };

  const runSalaryComparison = async () => {
    // Get actualPercentageFile from fileSlots context
    const actualPercentageFileFromContext = pickFile(
      (s) =>
        s.type === "Actual-Percentage-Bonus" ||
        (!!s.file && /actual.*percentage.*bonus/i.test(s.file.name))
    );

    if (!staffFile || !workerFile || !bonusFile) {
      alert(
        "Please upload Indiana Staff, Indiana Worker, and Bonus Calculation files for comparison"
      );
      return;
    }

    setIsComparing(true);

    try {
      console.log("Starting detailed employee comparison...");

      const staffWb = new ExcelJS.Workbook();
      const workerWb = new ExcelJS.Workbook();
      const bonusWb = new ExcelJS.Workbook();

      await staffWb.xlsx.load(await staffFile.arrayBuffer());
      await workerWb.xlsx.load(await workerFile.arrayBuffer());
      await bonusWb.xlsx.load(await bonusFile.arrayBuffer());

      // Load Actual Percentage file if available and extract employee IDs from Average sheet
      const averageSheetEmployeeIds = new Set<string>();
      if (actualPercentageFileFromContext) {
        const actualPercentageWb = new ExcelJS.Workbook();
        await actualPercentageWb.xlsx.load(
          await actualPercentageFileFromContext.arrayBuffer()
        );

        // Extract employee IDs from Average sheet
        const ws = actualPercentageWb.getWorksheet("Average");
        if (ws) {
          const empCodeCol = 2; // Column B
          for (let r = 2; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            const empCode = row.getCell(empCodeCol).value?.toString().trim();
            if (empCode && empCode !== "") {
              averageSheetEmployeeIds.add(empCode);
            }
          }
          console.log(
            `Found ${averageSheetEmployeeIds.size} employee IDs in Average sheet:`,
            Array.from(averageSheetEmployeeIds)
          );
        } else {
          console.log("Average sheet not found in Actual Percentage file");
        }
      } else {
        console.log(
          "Actual Percentage Bonus file not provided - skipping Average sheet check"
        );
      }

      const staffSheetNames = staffWb.worksheets
        .map((ws) => ws.name)
        .filter(
          (name) =>
            /(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JULY|AUG|SEP)-\d{2}/i.test(
              name
            ) &&
            name.includes("O") &&
            !/(OCT)-\d{2}/i.test(name)
        );

      const workerSheetNames = workerWb.worksheets
        .map((ws) => ws.name)
        .filter(
          (name) =>
            /(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JULY|AUG|SEP)-\d{2}/i.test(
              name
            ) &&
            name.includes("W") &&
            !/(OCT)-\d{2}/i.test(name)
        );

      console.log("Staff sheets:", staffSheetNames);
      console.log("Worker sheets:", workerSheetNames);

      const staffEmployees = extractStaffEmployees(staffWb, staffSheetNames);
      const workerEmployees = extractWorkerEmployees(
        workerWb,
        workerSheetNames
      );

      // Function to calculate October average for actual employees
      const calculateOctoberAverageForActualEmployees = (
        employees: Record<string, EmployeeMonthlySalary>,
        averageSheetEmployeeIds: Set<string>
      ) => {
        const monthsToAverage = [
          "Nov-24",
          "Dec-24",
          "Jan-25",
          "Feb-25",
          "Mar-25",
          "Apr-25",
          "May-25",
          "Jun-25",
          "Jul-25",
          "Aug-25",
          "Sep-25",
        ];

        for (const key in employees) {
          const emp = employees[key];

          // Skip assigning Oct-25 for Average-sheet employees (leave unset)
          if (averageSheetEmployeeIds.has(emp.employeeCode)) {
            continue;
          }

          // If September salary missing/zero, skip assigning October (leave unset)
          const septSalary = emp.monthlySalaries["Sep-25"];
          if (
            septSalary === undefined ||
            septSalary === null ||
            septSalary === 0
          ) {
            continue;
          }

          const finalDept = (emp.department || "").toUpperCase();
          const values: number[] = [];

          for (const month of monthsToAverage) {
            const salary = emp.monthlySalaries[month];

            // ✅ NEW: Check if salary is actually present AND has a meaningful value
            // undefined/null = month was blank in Excel (exclude from average)
            // 0 or positive number = month had value in Excel (include in average, even if 0)
            if (salary === undefined || salary === null) {
              continue; // Skip months with no data
            }

            // Require a recorded month department; no fallback to current/Final
            const monthDept = emp.monthlyDepartments?.[month];
            if (!monthDept) continue;

            const md = monthDept.toUpperCase();
            if (md === "C" || md !== finalDept) continue;

            // ✅ Include the value (even if it's 0) because cell had a value
            values.push(salary);
          }

          if (values.length > 0) {
            const average =
              values.reduce((sum, val) => sum + val, 0) / values.length;
            emp.monthlySalaries["Oct-25"] = Math.round(average * 2) / 2;
            if (emp.name.includes("SANJAY") && emp.name.includes("RATHOD")) {
              console.log(`\n📊 SANJAY RATHOD October Calculation:`);
              console.log(`   Values: [${values.join(", ")}]`);
              console.log(`   Average: ${average}`);
            }
            if (emp.monthlyDepartments) {
              emp.monthlyDepartments["Oct-25"] =
                emp.monthlyDepartments["Sep-25"] || emp.department;
            }
          } else {
            // Leave Oct-25 salary unset when no eligible months
            if (emp.monthlyDepartments) {
              emp.monthlyDepartments["Oct-25"] =
                emp.monthlyDepartments["Sep-25"] || emp.department;
            }
          }
        }
      };

      // Apply October calculation with Average sheet employee IDs
      calculateOctoberAverageForActualEmployees(
        staffEmployees,
        averageSheetEmployeeIds
      );
      calculateOctoberAverageForActualEmployees(
        workerEmployees,
        averageSheetEmployeeIds
      );

      const hrStaffEmployees = extractHREmployees(bonusWb, "Staff");
      const hrWorkerEmployees = extractHREmployees(bonusWb, "Worker");

      console.log(`Staff employees: ${Object.keys(staffEmployees).length}`);
      console.log(`Worker employees: ${Object.keys(workerEmployees).length}`);
      console.log(
        `HR Staff employees: ${Object.keys(hrStaffEmployees).length}`
      );
      console.log(
        `HR Worker employees: ${Object.keys(hrWorkerEmployees).length}`
      );

      const months = generateMonthHeaders();

      const staffComparisons = compareEmployees(
        staffEmployees,
        hrStaffEmployees,
        months
      );
      const workerComparisons = compareEmployees(
        workerEmployees,
        hrWorkerEmployees,
        months
      );

      console.log(`Staff comparisons: ${staffComparisons.length}`);
      console.log(`Worker comparisons: ${workerComparisons.length}`);

      setComparisonResults({
        staffComparisons,
        workerComparisons,
      });

      await downloadDetailedComparison(
        staffComparisons,
        workerComparisons,
        months
      );
    } catch (error) {
      console.error("Error during comparison:", error);
      alert(
        "Error performing salary comparison. Please check console for details."
      );
    } finally {
      setIsComparing(false);
    }
  };

  const downloadDetailedComparison = async (
    staffComparisons: EmployeeComparison[],
    workerComparisons: EmployeeComparison[],
    months: string[]
  ) => {
    const wb = new ExcelJS.Workbook();

    const staffWs = wb.addWorksheet("Staff Comparison");
    const staffHeader = [
      "Employee ID",
      "Employee Name",
      "Department",
      "Status",
      ...months.flatMap((m) => [`${m} (Actual)`, `${m} (HR)`, `${m} (Diff)`]),
      "Total Actual",
      "Total HR",
      "Total Difference",
    ];
    staffWs.addRow(staffHeader);
    staffWs.getRow(1).font = { bold: true };
    staffWs.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };

    staffComparisons.forEach((emp) => {
      let status = "Correct";
      if (emp.missingInHR) status = "⚠️ Missing in HR";
      else if (emp.hasMismatch) status = "❌ Mismatch";

      const row = [
        emp.employeeCode,
        emp.name,
        emp.department,
        status,
        ...months.flatMap((m) => [
          emp.actualSalaries[m] || 0,
          emp.hrSalaries[m] || 0,
          (emp.actualSalaries[m] || 0) - (emp.hrSalaries[m] || 0),
        ]),
        emp.totalActual,
        emp.totalHR,
        emp.totalDifference,
      ];
      const addedRow = staffWs.addRow(row);

      if (emp.missingInHR) {
        addedRow.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFCCCC" },
        };
      } else if (emp.hasMismatch) {
        addedRow.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF4CC" },
        };
      }
    });

    const workerWs = wb.addWorksheet("Worker Comparison");
    const workerHeader = [
      "Employee ID",
      "Employee Name",
      "Department",
      "Status",
      ...months.flatMap((m) => [`${m} (Actual)`, `${m} (HR)`, `${m} (Diff)`]),
      "Total Actual",
      "Total HR",
      "Total Difference",
    ];
    workerWs.addRow(workerHeader);
    workerWs.getRow(1).font = { bold: true };
    workerWs.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };

    workerComparisons.forEach((emp) => {
      let status = "Correct";
      if (emp.missingInHR) status = "⚠️ Missing in HR";
      else if (emp.hasMismatch) status = "❌ Mismatch";

      const row = [
        emp.employeeCode,
        emp.name,
        emp.department,
        status,
        ...months.flatMap((m) => [
          emp.actualSalaries[m] || 0,
          emp.hrSalaries[m] || 0,
          (emp.actualSalaries[m] || 0) - (emp.hrSalaries[m] || 0),
        ]),
        emp.totalActual,
        emp.totalHR,
        emp.totalDifference,
      ];
      const addedRow = workerWs.addRow(row);

      if (emp.missingInHR) {
        addedRow.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFCCCC" },
        };
      } else if (emp.hasMismatch) {
        addedRow.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF4CC" },
        };
      }
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `employee_comparison_${
      new Date().toISOString().split("T")[0]
    }.xlsx`;
    a.click();
  };

  const processExcelFiles = async () => {
    if (!staffFile || !workerFile) {
      alert("Please upload both Indiana Staff and Indiana Worker files");
      return;
    }
    setIsGenerating(true);

    try {
      const staffWb = new ExcelJS.Workbook();
      const workerWb = new ExcelJS.Workbook();
      await staffWb.xlsx.load(await staffFile.arrayBuffer());
      await workerWb.xlsx.load(await workerFile.arrayBuffer());

      let monthWiseWb: ExcelJS.Workbook | null = null;
      if (monthWiseFile) {
        monthWiseWb = new ExcelJS.Workbook();
        await monthWiseWb.xlsx.load(await monthWiseFile.arrayBuffer());
      }

      let bonusWb: ExcelJS.Workbook | null = null;
      let bonusMonthly: Record<string, { worker: number; staff: number }> = {};
      let bonusGross: { worker: number; staff: number } = {
        worker: 0,
        staff: 0,
      };
      if (bonusFile) {
        bonusWb = new ExcelJS.Workbook();
        await bonusWb.xlsx.load(await bonusFile.arrayBuffer());
        bonusMonthly = getBonusMonthlyTotals(bonusWb);
        bonusGross = getBonusGrossTotals(bonusWb);
      }

      let bonusSummaryWb: ExcelJS.Workbook | null = null;
      let bonusSummaryMonthlyGross: Record<string, number> = {};
      let bonusSummaryMonthlyFD: Record<string, number> = {};
      if (bonusSummaryFile) {
        bonusSummaryWb = new ExcelJS.Workbook();
        await bonusSummaryWb.xlsx.load(await bonusSummaryFile.arrayBuffer());
        bonusSummaryMonthlyGross = getBonusSummeryMonthlyGross(bonusSummaryWb);
        bonusSummaryMonthlyFD = getBonusSummeryMonthlyFD(bonusSummaryWb);
      }

      const months = generateMonthHeaders();

      const staffMap: Record<string, string> = {
        "Nov-24": "NOV-24 O",
        "Dec-24": "DEC-24 O",
        "Jan-25": "JAN-25 O",
        "Feb-25": "FEB-25 O",
        "Mar-25": "MAR-25 O",
        "Apr-25": "APR-25 O",
        "May-25": "MAY-25 O",
        "Jun-25": "JUN-25 O",
        "Jul-25": "JULY-25 O",
        "Aug-25": "AUG-25 O",
        "Sep-25": "SEP-25 O",
        "Oct-25": "OCT-25 O",
      };
      const workerMap: Record<string, string> = {
        "Nov-24": "NOV-24 W",
        "Dec-24": "DEC-24 W",
        "Jan-25": "JAN-25 W",
        "Feb-25": "FEB-25 W",
        "Mar-25": "MAR-25 W",
        "Apr-25": "APR-25 W",
        "May-25": "MAY-25 W",
        "Jun-25": "JUN-25 W",
        "Jul-25": "JULY-25 W",
        "Aug-25": "AUG-25 W",
        "Sep-25": "SEP-25 W",
        "Oct-25": "OCT-25 W",
      };

      const software: Record<string, CellTriplet> = {};
      const hr: Record<string, CellTriplet> = {};
      const diff: Record<string, CellTriplet> = {};

      for (const m of months) {
        if (m === "Oct-25") {
          // ✅ A value calculation - Sum from comparison results
          // A_v1: Sum of all Staff employees' October ACTUAL values from comparison table
          let A_v1 = 0;
          if (comparisonResults) {
            comparisonResults.staffComparisons.forEach((emp) => {
              const octActual = emp.actualSalaries["Oct-25"] || 0;
              A_v1 += octActual;
            });
          }

          // A_v2: Sum of all Staff employees' October HR values from comparison table
          let A_v2 = 0;
          if (comparisonResults) {
            comparisonResults.staffComparisons.forEach((emp) => {
              const octHR = emp.hrSalaries["Oct-25"] || 0;
              A_v2 += octHR;
            });
          }

          console.log("=== A Value Calculation (October) ===");
          console.log(`A_v1 (Staff Actual Sum): ${A_v1.toLocaleString()}`);
          console.log(`A_v2 (Staff HR Sum): ${A_v2.toLocaleString()}`);

          const A_diff = A_v2 - A_v1;

          // ✅ B value calculation - NEW: Use comparison results for both staff and worker
          let B_v1 = 0;
          let B_v2 = 0;

          if (comparisonResults) {
            // Sum all Staff October ACTUAL values
            comparisonResults.staffComparisons.forEach((emp) => {
              const octActual = emp.actualSalaries["Oct-25"] || 0;
              B_v1 += octActual;
            });

            // Sum all Worker October ACTUAL values
            comparisonResults.workerComparisons.forEach((emp) => {
              const octActual = emp.actualSalaries["Oct-25"] || 0;
              B_v1 += octActual;
            });

            // Sum all Staff October HR values
            comparisonResults.staffComparisons.forEach((emp) => {
              const octHR = emp.hrSalaries["Oct-25"] || 0;
              B_v2 += octHR;
            });

            // Sum all Worker October HR values
            comparisonResults.workerComparisons.forEach((emp) => {
              const octHR = emp.hrSalaries["Oct-25"] || 0;
              B_v2 += octHR;
            });
          }

          console.log("=== B Value Calculation (October) ===");
          console.log(
            `B_v1 (Staff + Worker Actual Sum): ${B_v1.toLocaleString()}`
          );
          console.log(`B_v2 (Staff + Worker HR Sum): ${B_v2.toLocaleString()}`);

          // ✅ Calculate October B Extras from comparison results
          let B_extras = 0;

          if (comparisonResults) {
            // Sum all Worker October ACTUAL values where department is C, A, or employee ID is 'N'
            comparisonResults.workerComparisons.forEach((emp) => {
              const dept = emp.department.toUpperCase();
              const empId = emp.employeeCode.toUpperCase();

              // Check if employee is in C, A department or has ID 'N'
              if (["C", "A"].includes(dept) || empId === "N") {
                const octActual = emp.actualSalaries["Oct-25"] || 0;
                B_extras += octActual;
              }
            });

            console.log(
              `✅ B_extras for October (from comparison): ₹${B_extras.toLocaleString()}`
            );
          }

          const B_diff = B_v2 - B_v1 + B_extras;

          // C value calculation
          // C, D, E values are zero for October
          const C_v1 = 0;
          const C_v2 = 0;
          const C_diff = 0;

          const D_v1 = 0;
          const D_v2 = 0;
          const D_diff = 0;

          const E_v1 = 0;
          const E_v2 = 0;
          const E_diff = 0;

          software[m] = {
            A: round(A_v1),
            B: round(B_v1),
            C: round(C_v1),
            D: round(D_v1),
            E: round(E_v1),
          };

          hr[m] = {
            A: round(A_v2),
            B: round(B_v2),
            C: round(C_v2),
            D: round(D_v2),
            E: round(E_v2),
          };

          extras[m] = {
            A: 0,
            B: round(B_extras),
            C: 0,
            D: 0,
            E: 0,
          };

          diff[m] = {
            A: round(A_diff),
            B: round(B_diff),
            C: round(C_diff),
            D: round(D_diff),
            E: round(E_diff),
          };

          continue;
        }

        // Original logic for other months (Nov-24 to Sep-25)
        const staffSheet = staffWb.getWorksheet(staffMap[m]);
        const workerMonthSheet = (monthWiseWb || workerWb).getWorksheet(
          workerMap[m]
        );

        // A value
        let A_v1 = 0;
        if (staffSheet) {
          A_v1 = readStaffSalary1Total(staffSheet);
        }

        let A_v2 = 0;
        if (workerMonthSheet) {
          const { col, headerRow } = findColByHeader(
            workerMonthSheet,
            ["WD"],
            ["SALARY"]
          );
          A_v2 = readColumnGrandTotal(workerMonthSheet, col, headerRow);
        }

        const A_diff = A_v2 - A_v1;

        // B value
        let B_v1staff = 0,
          B_v1worker = 0;
        if (staffSheet) {
          B_v1staff = readStaffSalary1Total(staffSheet);
        }

        const workerSheetForMonth = workerWb.getWorksheet(workerMap[m]);
        if (workerSheetForMonth) {
          B_v1worker = sumWorkerSalary1(workerSheetForMonth);
        }

        const B_v1 = B_v1staff + B_v1worker;
        const B_v2 = bonusMonthly[m]
          ? bonusMonthly[m].worker + bonusMonthly[m].staff
          : 0;

        // Calculate B Extras
        const B_extras = calculateBExtras(workerWb, workerMap, m);

        // Update diff calculation to include extras
        const B_diff = B_v2 - B_v1 + B_extras;

        // C value
        let C_g1worker = 0,
          C_g1staffGross = 0;
        if (workerSheetForMonth) {
          C_g1worker = sumWorkerSalary1(workerSheetForMonth);
        }
        if (staffSheet) {
          C_g1staffGross = readStaffGrossTotal(staffSheet);
        }

        const C_v1 = C_g1worker + C_g1staffGross;

        let C_v2 = 0;
        const monthWiseMonth = (monthWiseWb || workerWb).getWorksheet(
          workerMap[m]
        );
        if (monthWiseMonth) {
          C_v2 = getGrossSalaryGrandTotal(monthWiseMonth);
        }

        const C_diff = C_v2 - C_v1;

        // D value
        let D_1worker = 0,
          D_1staffGross = 0;
        if (workerSheetForMonth) {
          D_1worker = sumWorkerSalary1(workerSheetForMonth);
        }
        if (staffSheet) {
          D_1staffGross = readStaffGrossTotal(staffSheet);
        }

        const D_v1 = D_1worker + D_1staffGross;
        const D_v2 = bonusSummaryMonthlyGross[m] ?? 0;

        // E value
        let E_1worker = 0,
          E_1staffGross = 0;
        if (workerSheetForMonth) {
          E_1worker = sumWorkerSalary1(workerSheetForMonth);
        }
        if (staffSheet) {
          E_1staffGross = readStaffGrossTotal(staffSheet);
        }

        const E_v1 = (E_1worker + E_1staffGross) * 0.0833;
        const E_v2 = bonusSummaryMonthlyFD[m] ?? 0;

        software[m] = {
          A: round(A_v1),
          B: round(B_v1),
          C: round(C_v1),
          D: round(D_v1),
          E: round(E_v1),
        };

        hr[m] = {
          A: round(A_v2),
          B: round(B_v2),
          C: round(C_v2),
          D: round(D_v2),
          E: round(E_v2),
        };

        diff[m] = {
          A: round(A_diff),
          B: round(B_diff),
          C: round(C_diff),
          D: round(D_v2 - D_v1),
          E: round(E_v2 - E_v1),
        };

        extras[m] = {
          A: 0,
          B: round(B_extras),
          C: 0,
          D: 0,
          E: 0,
        };
      }
      setReportData({
        months,
        departments: [
          { name: "Software", data: software },
          { name: "HR", data: hr },
          { name: "A,N,C", data: extras },
          { name: "diff", data: diff },
        ],
      });

      // Persist audit logs only for Generate Report
      try {
        await storeAuditForGenerateReport(
          {
            months,
            departments: [
              { name: "Software", data: software },
              { name: "HR", data: hr },
              { name: "A,N,C", data: extras },
              { name: "diff", data: diff },
            ],
          },
          comparisonResults // may be null; handled in builder
        );
      } catch (e) {
        console.error("Audit store failed", e);
      }
    } catch (e) {
      console.error(e);
      alert("Error processing files. Please verify formats and sheet names.");
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadReport = async () => {
    if (!reportData) return;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Report");

    const hdr: (string | number)[] = [""];
    reportData.months.forEach((m) => hdr.push(m, "", "", "", ""));
    ws.addRow(hdr);
    reportData.months.forEach((_, idx) =>
      ws.mergeCells(1, 2 + idx * 5, 1, 6 + idx * 5)
    );

    const sub: (string | number)[] = [""];
    reportData.months.forEach(() => sub.push("A", "B", "C", "D", "E"));
    ws.addRow(sub);

    reportData.departments.forEach((d) => {
      const row: (string | number)[] = [d.name];
      reportData.months.forEach((m) => {
        row.push(
          d.data[m]?.A ?? 0,
          d.data[m]?.B ?? 0,
          d.data[m]?.C ?? 0,
          d.data[m]?.D ?? 0,
          d.data[m]?.E ?? 0
        );
      });
      ws.addRow(row);
    });
    ws.getRow(1).font = { bold: true };
    ws.getRow(2).font = { bold: true };
    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" },
    };

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `report_${new Date().toISOString().split("T")[0]}.xlsx`;
    a.click();
  };

  const BonusCard = () => (
    <div
      className={`border-2 rounded-lg p-6 ${
        bonusFile
          ? "border-green-300 bg-green-50"
          : "border-gray-300 bg-gray-50"
      }`}
    >
      <div className="flex items-center gap-3 mb-4">
        {bonusFile ? (
          <svg
            className="w-8 h-8 text-green-600"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
            />
          </svg>
        ) : (
          <svg
            className="w-8 h-8 text-gray-400"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M12 9v2m0 4h.01m-6.938 4h13.856"
            />
          </svg>
        )}
        <div>
          <h3 className="text-lg font-bold text-gray-800">Bonus Sheet</h3>
          <span className="text-xs bg-gray-200 text-gray-700 px-2 py-0.5 rounded font-medium">
            Optional
          </span>
        </div>
      </div>
      {bonusFile ? (
        <div className="space-y-2">
          <div className="bg-white rounded p-3 border border-green-200">
            <p className="text-sm font-medium text-gray-800 truncate">
              {bonusFile.name}
            </p>
            <p className="text-xs text-gray-500 mt-1">
              Size: {(bonusFile.size / 1024).toFixed(2)} KB
            </p>
          </div>
          <div className="flex items-center gap-2 text-xs text-green-700 bg-green-100 px-3 py-2 rounded">
            <svg
              className="w-4 h-4"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M5 13l4 4L19 7"
              />
            </svg>
            File is cached and ready
          </div>
        </div>
      ) : (
        <div className="bg-white rounded p-3 border border-gray-200">
          <p className="text-sm text-gray-600 font-medium">No file uploaded</p>
          <p className="text-xs text-gray-500 mt-1">This file is optional</p>
        </div>
      )}
    </div>
  );

  const BonusSummaryCard = () => (
    <div
      className={`border-2 rounded-lg p-6 ${
        bonusSummaryFile
          ? "border-green-300 bg-green-50"
          : "border-gray-300 bg-gray-50"
      }`}
    >
      <div className="flex items-center gap-3 mb-4">
        {bonusSummaryFile ? (
          <svg
            className="w-8 h-8 text-green-600"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
            />
          </svg>
        ) : (
          <svg
            className="w-8 h-8 text-gray-400"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M12 9v2m0 4h.01m-6.938 4h13.856"
            />
          </svg>
        )}
        <div>
          <h3 className="text-lg font-bold text-gray-800">Bonus Summery</h3>
          <span className="text-xs bg-gray-200 text-gray-700 px-2 py-0.5 rounded font-medium">
            Optional
          </span>
        </div>
      </div>
      {bonusSummaryFile ? (
        <div className="space-y-2">
          <div className="bg-white rounded p-3 border border-green-200">
            <p className="text-sm font-medium text-gray-800 truncate">
              {bonusSummaryFile.name}
            </p>
            <p className="text-xs text-gray-500 mt-1">
              Size: {(bonusSummaryFile.size / 1024).toFixed(2)} KB
            </p>
          </div>
          <div className="flex items-center gap-2 text-xs text-green-700 bg-green-100 px-3 py-2 rounded">
            <svg
              className="w-4 h-4"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M5 13l4 4L19 7"
              />
            </svg>
            File is cached and ready
          </div>
        </div>
      ) : (
        <div className="bg-white rounded p-3 border border-gray-200">
          <p className="text-sm text-gray-600 font-medium">No file uploaded</p>
          <p className="text-xs text-gray-500 mt-1">This file is optional</p>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-5 px-4">
      <div className="mx-auto max-w-7xl">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 2 - Generate Report
              </h1>
              <p className="text-gray-600 mt-2">
                Review cached files and generate your report
              </p>
            </div>
            <button
              onClick={() => router.push("/")}
              className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
            >
              ← Back to Step 1
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-5 gap-6 mb-8">
            <div
              className={`border-2 rounded-lg p-6 ${
                staffFile
                  ? "border-green-300 bg-green-50"
                  : "border-red-300 bg-red-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    staffFile ? "text-green-600" : "text-red-600"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d={
                      staffFile
                        ? "M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                        : "M6 18L18 6M6 6l12 12"
                    }
                  />
                </svg>
                <h3 className="text-lg font-bold text-gray-800">
                  Indiana Staff
                </h3>
              </div>
              {staffFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {staffFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(staffFile.size / 1024).toFixed(2)} KB
                    </p>
                  </div>
                  <div className="flex items-center gap-2 text-xs text-green-700 bg-green-100 px-3 py-2 rounded">
                    <svg
                      className="w-4 h-4"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M5 13l4 4L19 7"
                      />
                    </svg>
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-red-200">
                  <p className="text-sm text-red-700 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    Please go back to Step 1 and upload this file
                  </p>
                </div>
              )}
            </div>

            <div
              className={`border-2 rounded-lg p-6 ${
                workerFile
                  ? "border-green-300 bg-green-50"
                  : "border-red-300 bg-red-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    workerFile ? "text-green-600" : "text-red-600"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d={
                      workerFile
                        ? "M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                        : "M6 18L18 6M6 6l12 12"
                    }
                  />
                </svg>
                <h3 className="text-lg font-bold text-gray-800">
                  Indiana Worker
                </h3>
              </div>
              {workerFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {workerFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(workerFile.size / 1024).toFixed(2)} KB
                    </p>
                  </div>
                  <div className="flex items-center gap-2 text-xs text-green-700 bg-green-100 px-3 py-2 rounded">
                    <svg
                      className="w-4 h-4"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M5 13l4 4L19 7"
                      />
                    </svg>
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-red-200">
                  <p className="text-sm text-red-700 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    Please go back to Step 1 and upload this file
                  </p>
                </div>
              )}
            </div>

            <div
              className={`border-2 rounded-lg p-6 ${
                monthWiseFile
                  ? "border-green-300 bg-green-50"
                  : "border-gray-300 bg-gray-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    monthWiseFile ? "text-green-600" : "text-gray-400"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
                <div>
                  <h3 className="text-lg font-bold text-gray-800">
                    Month Wise Sheet
                  </h3>
                  <span className="text-xs bg-gray-200 text-gray-700 px-2 py-0.5 rounded font-medium">
                    Optional
                  </span>
                </div>
              </div>
              {monthWiseFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {monthWiseFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(monthWiseFile.size / 1024).toFixed(2)} KB
                    </p>
                  </div>
                  <div className="flex items-center gap-2 text-xs text-green-700 bg-green-100 px-3 py-2 rounded">
                    <svg
                      className="w-4 h-4"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M5 13l4 4L19 7"
                      />
                    </svg>
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-gray-200">
                  <p className="text-sm text-gray-600 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    This file is optional
                  </p>
                </div>
              )}
            </div>

            {BonusCard()}
            {BonusSummaryCard()}
          </div>

          <div className="flex justify-center gap-4 mb-8">
            {/* Step 1: Run Comparison button - Always clickable */}
            <button
              onClick={runSalaryComparison}
              disabled={!staffFile || !workerFile || !bonusFile || isComparing}
              className="px-8 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center gap-2"
            >
              {isComparing ? (
                <>
                  <svg
                    className="animate-spin h-5 w-5"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    />
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    />
                  </svg>
                  Comparing...
                </>
              ) : (
                <>
                  <svg
                    className="w-5 h-5"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"
                    />
                  </svg>
                  {comparisonResults ? "Re-run Comparison" : "Run Comparison"}
                </>
              )}
            </button>

            {/* Step 2: Generate Report button - Only visible after comparison */}
            {comparisonResults && (
              <button
                onClick={processExcelFiles}
                disabled={!staffFile || !workerFile || isGenerating}
                className="px-8 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center gap-2"
              >
                {isGenerating ? (
                  <>
                    <svg
                      className="animate-spin h-5 w-5"
                      fill="none"
                      viewBox="0 0 24 24"
                    >
                      <circle
                        className="opacity-25"
                        cx="12"
                        cy="12"
                        r="10"
                        stroke="currentColor"
                        strokeWidth="4"
                      />
                      <path
                        className="opacity-75"
                        fill="currentColor"
                        d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                      />
                    </svg>
                    Generating Report...
                  </>
                ) : (
                  <>
                    <svg
                      className="w-5 h-5"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                      />
                    </svg>
                    {reportData ? "Re-generate Report" : "Generate Report"}
                  </>
                )}
              </button>
            )}

            {/* Step 3: Move to Step 3 button - Only visible after report is generated */}
            {reportData && (
              <button
                onClick={handleMoveToStep3}
                className="px-8 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition flex items-center gap-2"
              >
                <svg
                  className="w-5 h-5"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M13 7l5 5m0 0l-5 5m5-5H6"
                  />
                </svg>
                Move to Step 3
              </button>
            )}
          </div>

          {comparisonResults && (
            <div className="mt-8 mb-8">
              <div className="bg-gradient-to-r from-purple-50 to-indigo-50 rounded-lg p-6 border-2 border-purple-200">
                <h2 className="text-2xl font-bold text-gray-800 mb-4 flex items-center gap-2">
                  <svg
                    className="w-6 h-6 text-purple-600"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"
                    />
                  </svg>
                  Employee-wise Salary Comparison
                </h2>

                <div className="flex gap-2 mb-6">
                  <button
                    onClick={() => handleTabSwitch("staff")}
                    className={`px-6 py-2 rounded-lg font-medium transition ${
                      activeTab === "staff"
                        ? "bg-purple-600 text-white"
                        : "bg-white text-gray-700 hover:bg-gray-100"
                    }`}
                  >
                    Staff Comparison (
                    {comparisonResults.staffComparisons.length})
                  </button>

                  <button
                    onClick={() => handleTabSwitch("worker")}
                    className={`px-6 py-2 rounded-lg font-medium transition ${
                      activeTab === "worker"
                        ? "bg-purple-600 text-white"
                        : "bg-white text-gray-700 hover:bg-gray-100"
                    }`}
                  >
                    Worker Comparison (
                    {comparisonResults.workerComparisons.length})
                  </button>
                  <button
                    onClick={toggleIgnoreSpecialDepts}
                    className={`px-6 py-2 rounded-lg font-medium transition ml-auto flex items-center gap-2 ${
                      ignoredEmployees.size > 0
                        ? "bg-orange-600 text-white hover:bg-orange-700"
                        : "bg-gray-600 text-white hover:bg-gray-700"
                    }`}
                  >
                    {ignoredEmployees.size > 0 ? (
                      <>
                        <svg
                          className="w-5 h-5"
                          fill="none"
                          stroke="currentColor"
                          viewBox="0 0 24 24"
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth={2}
                            d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"
                          />
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth={2}
                            d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"
                          />
                        </svg>
                        Show C, A, N ({ignoredEmployees.size})
                      </>
                    ) : (
                      <>
                        <svg
                          className="w-5 h-5"
                          fill="none"
                          stroke="currentColor"
                          viewBox="0 0 24 24"
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth={2}
                            d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21"
                          />
                        </svg>
                        Ignore C, A, N
                      </>
                    )}
                  </button>
                </div>

                {activeTab === "staff" && (
                  <div className="bg-white rounded-lg overflow-x-auto">
                    {isLoadingTab ? (
                      // ✅ ADD THIS LOADING STATE
                      <div className="flex items-center justify-center py-20">
                        <div className="text-center">
                          <svg
                            className="animate-spin h-12 w-12 text-purple-600 mx-auto mb-4"
                            fill="none"
                            viewBox="0 0 24 24"
                          >
                            <circle
                              className="opacity-25"
                              cx="12"
                              cy="12"
                              r="10"
                              stroke="currentColor"
                              strokeWidth="4"
                            ></circle>
                            <path
                              className="opacity-75"
                              fill="currentColor"
                              d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                            ></path>
                          </svg>
                          <p className="text-gray-600 font-medium">
                            Loading Staff Data...
                          </p>
                        </div>
                      </div>
                    ) : (
                      <div className="max-h-[600px] overflow-y-auto">
                        <table className="w-full border-collapse text-sm">
                          <thead className="sticky top-0 bg-gray-100 z-10">
                            <tr>
                              <th className="border px-3 py-2 text-left">
                                Employee ID
                              </th>
                              <th className="border px-3 py-2 text-left">
                                Employee Name
                              </th>
                              <th className="border px-3 py-2 text-left">
                                Department
                              </th>
                              <th className="border px-3 py-2 text-left">
                                Status
                              </th>
                              {generateMonthHeaders().map((m) => (
                                <th
                                  key={m}
                                  className="border px-2 py-2 text-center"
                                  colSpan={3}
                                >
                                  {m}
                                </th>
                              ))}
                              <th className="border px-3 py-2 text-right">
                                Total Diff
                              </th>
                            </tr>
                            <tr className="bg-gray-50">
                              <th className="border px-3 py-1"></th>
                              <th className="border px-3 py-1"></th>
                              <th className="border px-3 py-1"></th>
                              <th className="border px-3 py-1"></th>
                              {generateMonthHeaders().map((m) => (
                                <React.Fragment key={m}>
                                  <th className="border px-1 py-1 text-xs">
                                    Actual
                                  </th>
                                  <th className="border px-1 py-1 text-xs">
                                    HR
                                  </th>
                                  <th className="border px-1 py-1 text-xs">
                                    Diff
                                  </th>
                                </React.Fragment>
                              ))}
                              <th className="border px-3 py-1"></th>
                            </tr>
                          </thead>
                          <tbody>
                            {comparisonResults.staffComparisons.map(
                              (emp, idx) => {
                                const bgColor = emp.missingInHR
                                  ? "bg-red-50"
                                  : emp.hasMismatch
                                  ? "bg-yellow-50"
                                  : "";

                                return (
                                  <tr
                                    key={idx}
                                    className={`${bgColor} hover:bg-gray-50`}
                                  >
                                    <td className="border px-3 py-2 text-xs">
                                      {emp.employeeCode}
                                    </td>
                                    <td className="border px-3 py-2 font-medium">
                                      {emp.name}
                                    </td>
                                    <td className="border px-3 py-2 text-xs">
                                      {emp.department}
                                    </td>
                                    <td className="border px-3 py-2 text-xs">
                                      {emp.missingInHR ? (
                                        <span className="text-red-600 font-bold">
                                          ⚠️ Missing
                                        </span>
                                      ) : emp.hasMismatch ? (
                                        <span className="text-orange-600 font-bold">
                                          ❌ Mismatch
                                        </span>
                                      ) : (
                                        <span className="text-green-600">
                                          ✓ Match
                                        </span>
                                      )}
                                    </td>
                                    {generateMonthHeaders().map((m) => {
                                      const actualSal =
                                        emp.actualSalaries[m] || 0;
                                      const hrSal = emp.hrSalaries[m] || 0;
                                      const diff = actualSal - hrSal;
                                      const hasDiff = Math.abs(diff) > 1;

                                      // Check the department for THIS specific month
                                      const monthDept =
                                        emp.monthlyDepartments?.[m] ||
                                        emp.department;
                                      const shouldIgnoreThisMonth =
                                        ["C", "A"].includes(
                                          monthDept.toUpperCase()
                                        ) ||
                                        emp.employeeCode.toUpperCase() === "N";

                                      // If ignore button is active AND this month is in C/A dept, grey it out
                                      const isIgnoredByUser =
                                        ignoredEmployees.has(
                                          emp.employeeCode || emp.name
                                        );
                                      const shouldGreyOut =
                                        isIgnoredByUser &&
                                        shouldIgnoreThisMonth;

                                      return (
                                        <React.Fragment key={m}>
                                          <td
                                            className={`border px-2 py-1 text-right text-xs ${
                                              shouldGreyOut
                                                ? "bg-gray-200 text-gray-400"
                                                : ""
                                            }`}
                                          >
                                            {actualSal > 0
                                              ? actualSal.toLocaleString()
                                              : "-"}
                                          </td>
                                          <td
                                            className={`border px-2 py-1 text-right text-xs ${
                                              shouldGreyOut
                                                ? "bg-gray-200 text-gray-400"
                                                : ""
                                            }`}
                                          >
                                            {hrSal > 0
                                              ? hrSal.toLocaleString()
                                              : "-"}
                                          </td>
                                          <td
                                            className={`border px-2 py-1 text-right text-xs font-medium ${
                                              shouldGreyOut
                                                ? "bg-gray-200 text-gray-400" // Grey out ignored months
                                                : hasDiff
                                                ? "bg-yellow-200 text-red-600" // Only show error for non-ignored months
                                                : "text-gray-400"
                                            }`}
                                          >
                                            {shouldGreyOut
                                              ? "Ignored"
                                              : diff !== 0
                                              ? diff.toLocaleString()
                                              : "-"}
                                          </td>
                                        </React.Fragment>
                                      );
                                    })}

                                    <td
                                      className={`border px-3 py-2 text-right font-bold ${
                                        Math.abs(emp.totalDifference) > 1
                                          ? "text-red-600"
                                          : "text-gray-400"
                                      }`}
                                    >
                                      {emp.totalDifference.toLocaleString()}
                                    </td>
                                  </tr>
                                );
                              }
                            )}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {activeTab === "worker" && (
                  <div className="bg-white rounded-lg overflow-x-auto">
                    <div className="max-h-[600px] overflow-y-auto">
                      <table className="w-full border-collapse text-sm">
                        <thead className="sticky top-0 bg-gray-100 z-10">
                          <tr>
                            <th className="border px-3 py-2 text-left">
                              Employee ID
                            </th>
                            <th className="border px-3 py-2 text-left">
                              Employee Name
                            </th>
                            <th className="border px-3 py-2 text-left">
                              Department
                            </th>
                            <th className="border px-3 py-2 text-left">
                              Status
                            </th>
                            {generateMonthHeaders().map((m) => (
                              <th
                                key={m}
                                className="border px-2 py-2 text-center"
                                colSpan={3}
                              >
                                {m}
                              </th>
                            ))}
                            <th className="border px-3 py-2 text-right">
                              Total Diff
                            </th>
                          </tr>
                          <tr className="bg-gray-50">
                            <th className="border px-3 py-1"></th>
                            <th className="border px-3 py-1"></th>
                            <th className="border px-3 py-1"></th>
                            <th className="border px-3 py-1"></th>
                            {generateMonthHeaders().map((m) => (
                              <React.Fragment key={m}>
                                <th className="border px-1 py-1 text-xs">
                                  Actual
                                </th>
                                <th className="border px-1 py-1 text-xs">HR</th>
                                <th className="border px-1 py-1 text-xs">
                                  Diff
                                </th>
                              </React.Fragment>
                            ))}
                            <th className="border px-3 py-1"></th>
                          </tr>
                        </thead>
                        <tbody>
                          {comparisonResults.workerComparisons.map(
                            (emp, idx) => {
                              const isIgnoredEmployee = ignoredEmployees.has(
                                emp.employeeCode || emp.name
                              );

                              // ✅ NEW: Check if employee should be completely hidden
                              // Only hide if ALL their months are in C/A departments or they're ID 'N'
                              let hasAnyNonIgnoredMonth = false;
                              const months = generateMonthHeaders();

                              for (const m of months) {
                                const monthDept =
                                  emp.monthlyDepartments?.[m] || emp.department;
                                const isIgnoredDept =
                                  ["C", "A"].includes(
                                    monthDept.toUpperCase()
                                  ) || emp.employeeCode.toUpperCase() === "N";

                                if (!isIgnoredDept) {
                                  hasAnyNonIgnoredMonth = true;
                                  break;
                                }
                              }

                              // ✅ If ignore button is active AND employee has NO non-ignored months, hide entire row
                              if (isIgnoredEmployee && !hasAnyNonIgnoredMonth) {
                                return null;
                              }

                              // Calculate if there are REAL errors (excluding ignored months)
                              let hasRealErrors = false;

                              for (const m of months) {
                                const actualSal = emp.actualSalaries[m] || 0;
                                const hrSal = emp.hrSalaries[m] || 0;
                                const diff = Math.abs(actualSal - hrSal);

                                // Check if this month should be ignored
                                const monthDept =
                                  emp.monthlyDepartments?.[m] || emp.department;
                                const isIgnoredMonth =
                                  isIgnoredEmployee &&
                                  (["C", "A"].includes(
                                    monthDept.toUpperCase()
                                  ) ||
                                    emp.employeeCode.toUpperCase() === "N");

                                // Only count as error if diff exists AND month is not ignored
                                if (diff > 1 && !isIgnoredMonth) {
                                  hasRealErrors = true;
                                  break;
                                }
                              }

                              // Determine background color based on REAL errors only
                              const bgColor = emp.missingInHR
                                ? "bg-red-50"
                                : hasRealErrors
                                ? "bg-yellow-50"
                                : "";

                              // Determine status text
                              let statusText = "✓ Match";
                              let statusColor = "text-green-600";

                              if (emp.missingInHR) {
                                statusText = "⚠️ Missing in HR";
                                statusColor = "text-orange-600";
                              } else if (hasRealErrors) {
                                statusText = "✗ Mismatch";
                                statusColor = "text-red-600";
                              }

                              return (
                                <tr
                                  key={idx}
                                  className={`${bgColor} hover:bg-gray-50`}
                                >
                                  <td className="border px-3 py-2 text-center text-xs">
                                    {emp.employeeCode}
                                  </td>

                                  {/* ✅ Column 2: Employee Name */}
                                  <td className="border px-3 py-2 font-medium text-sm sticky left-0 bg-white">
                                    {emp.name}
                                  </td>

                                  {/* ✅ Column 3: Department */}
                                  <td className="border px-3 py-2 text-center text-xs">
                                    {emp.department}
                                  </td>

                                  {/* ✅ Column 4: Status */}
                                  <td
                                    className={`border px-3 py-2 text-center text-xs font-medium ${statusColor}`}
                                  >
                                    {statusText}
                                  </td>

                                  {/* Monthly cells */}
                                  {months.map((m) => {
                                    const actualSal =
                                      emp.actualSalaries[m] || 0;
                                    const hrSal = emp.hrSalaries[m] || 0;
                                    const diff = actualSal - hrSal;
                                    const hasDiff = Math.abs(diff) > 1;

                                    // Check if THIS SPECIFIC MONTH should be ignored
                                    const monthDept =
                                      emp.monthlyDepartments?.[m] ||
                                      emp.department;
                                    const shouldIgnoreThisMonth =
                                      ["C", "A"].includes(
                                        monthDept.toUpperCase()
                                      ) ||
                                      emp.employeeCode.toUpperCase() === "N";

                                    // If ignore button is active AND this month should be ignored
                                    const shouldHide =
                                      isIgnoredEmployee &&
                                      shouldIgnoreThisMonth;

                                    // ✅ FIX: Render empty cells instead of null to maintain alignment
                                    if (shouldHide) {
                                      return (
                                        <React.Fragment key={m}>
                                          <td className="border px-2 py-1 bg-gray-100"></td>
                                          <td className="border px-2 py-1 bg-gray-100"></td>
                                          <td className="border px-2 py-1 bg-gray-100"></td>
                                        </React.Fragment>
                                      );
                                    }

                                    return (
                                      <React.Fragment key={m}>
                                        <td className="border px-2 py-1 text-right text-xs">
                                          {actualSal > 0
                                            ? actualSal.toLocaleString()
                                            : "-"}
                                        </td>
                                        <td className="border px-2 py-1 text-right text-xs">
                                          {hrSal > 0
                                            ? hrSal.toLocaleString()
                                            : "-"}
                                        </td>
                                        <td
                                          className={`border px-2 py-1 text-right text-xs font-medium ${
                                            hasDiff
                                              ? "bg-yellow-200 text-red-600"
                                              : "text-gray-400"
                                          }`}
                                        >
                                          {diff !== 0
                                            ? diff.toLocaleString()
                                            : "-"}
                                        </td>
                                      </React.Fragment>
                                    );
                                  })}
                                </tr>
                              );
                            }
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                <div className="text-sm text-gray-600 mt-4">
                  📄 A detailed Excel report with all employee comparisons has
                  been downloaded automatically
                </div>
              </div>
            </div>
          )}

          {reportData && (
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-2xl font-bold text-gray-800">
                  Report Preview
                </h2>
                <button
                  onClick={downloadReport}
                  className="px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition flex items-center gap-2"
                >
                  Download Excel
                </button>
              </div>
              <div className="bg-gray-900 rounded-lg p-6 overflow-x-auto">
                <table className="w-full border-collapse border border-gray-600">
                  <thead>
                    <tr className="bg-gray-800">
                      <th
                        className="border border-gray-600 px-4 py-3 text-white"
                        rowSpan={2}
                      ></th>
                      {reportData.months.map((m, i) => (
                        <th
                          key={i}
                          className="border border-gray-600 px-2 py-3 text-white"
                          colSpan={5}
                        >
                          {m}
                        </th>
                      ))}
                    </tr>
                    <tr className="bg-gray-800">
                      {reportData.months.map((_, i) => (
                        <React.Fragment key={i}>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            A
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            B
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            C
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            D
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            E
                          </th>
                        </React.Fragment>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {reportData.departments.map((d, i) => (
                      <tr key={i} className="bg-gray-800 hover:bg-gray-700">
                        <td className="border border-gray-600 px-4 py-3 text-white font-semibold">
                          {d.name}
                        </td>
                        {reportData.months.map((m, j) => (
                          <React.Fragment key={j}>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.A ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.B ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.C ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.D ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.E ?? 0}
                            </td>
                          </React.Fragment>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>
      {showPasswordModal && (
        <div className="fixed inset-0 bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-8 max-w-md w-full mx-4">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-xl font-bold text-gray-800">
                Password Required
              </h3>
              <button
                onClick={handleModalClose}
                className="text-gray-400 hover:text-gray-600"
              >
                <svg
                  className="w-6 h-6"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M6 18L18 6M6 6l12 12"
                  />
                </svg>
              </button>
            </div>

            <div className="mb-4">
              <p className="text-gray-600 mb-4">
                The total differences exceed the number of employees for some
                months. Please enter the password to proceed to Step 3.
              </p>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                onKeyPress={(e) => e.key === "Enter" && handlePasswordSubmit()}
                placeholder="Enter password"
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500 focus:border-transparent"
                autoFocus
              />
              {passwordError && (
                <p className="text-red-600 text-sm mt-2">{passwordError}</p>
              )}
            </div>

            <div className="flex gap-3">
              <button
                onClick={handleModalClose}
                className="flex-1 px-4 py-2 bg-gray-300 text-gray-700 rounded-lg hover:bg-gray-400 transition"
              >
                Cancel
              </button>
              <button
                onClick={handlePasswordSubmit}
                className="flex-1 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition"
              >
                Submit
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
