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
        const actualSal = emp.actualSalaries[month] || 0;
        const hrSal = emp.hrSalaries[month] || 0;
        const diff = Math.abs(actualSal - hrSal);

        if (actualSal > 0 || hrSal > 0) {
          employeeCount++;
          totalDiff += diff;
        }
      });

      monthlyStats[month] = { totalDiff, employeeCount };

      // Check if difference exceeds employee count
      if (totalDiff > employeeCount) {
        canProceed = false;
      }
    });

    return { canProceed, monthlyStats };
  };

  const handleMoveToStep3 = () => {
    const { canProceed } = calculateMonthlyDifferences();

    if (canProceed) {
      // Direct navigation to Step 3
      router.push("/step3");
    } else {
      // Show password modal
      setShowPasswordModal(true);
      setPassword("");
      setPasswordError("");
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

  const getBonusOctoberTotals = (
    bonusWb: ExcelJS.Workbook
  ): { worker: number; staff: number } => {
    const out = { worker: 0, staff: 0 };

    const readOctober = (ws?: ExcelJS.Worksheet, sheetName: string = "") => {
      if (!ws) {
        console.log(`❌ Sheet ${sheetName} not found`);
        return 0;
      }

      let headerRow = -1;
      let octoberCol = -1;

      for (let r = 1; r <= 3; r++) {
        const row = ws.getRow(r);
        for (let c = 16; c <= 18; c++) {
          const cellVal = row.getCell(c).value;

          if (cellVal instanceof Date) {
            if (cellVal.getMonth() === 9 && cellVal.getFullYear() === 2025) {
              headerRow = r;
              octoberCol = c;
              console.log(
                `✅ ${sheetName}: Found October at row ${r}, col ${c}`
              );
              break;
            }
          }

          const cellText = cellVal?.toString().toUpperCase() || "";
          if (cellText.includes("SALARY") && cellText.includes("12")) {
            headerRow = r;
            octoberCol = c;
            console.log(
              `✅ ${sheetName}: Found Salary12 at row ${r}, col ${c}`
            );
            break;
          }
        }
        if (octoberCol > 0) break;
      }

      if (headerRow < 0 || octoberCol < 0) {
        console.log(`❌ ${sheetName}: Could not find October column`);
        return 0;
      }

      let grandTotalRow = -1;
      const lastRow = ws.rowCount;

      for (let r = lastRow; r >= lastRow - 30; r--) {
        const nameCell = ws.getRow(r).getCell(4);
        const nameText = nameCell.value?.toString().toUpperCase() || "";

        if (nameText.includes("GRAND") && nameText.includes("TOTAL")) {
          grandTotalRow = r;
          console.log(`✅ ${sheetName}: Found Grand Total at row ${r}`);
          break;
        }
      }

      if (grandTotalRow < 0) {
        console.log(`❌ ${sheetName}: Could not find Grand Total row`);
        return 0;
      }

      const totalValue = num(ws.getRow(grandTotalRow).getCell(octoberCol));
      console.log(`✅ ${sheetName}: Grand Total value = ${totalValue}`);
      return totalValue;
    };

    out.worker = readOctober(bonusWb.getWorksheet("Worker"), "Worker");
    out.staff = readOctober(bonusWb.getWorksheet("Staff"), "Staff");

    console.log(`\n📊 B(HR) Calculation:`);
    console.log(`   Worker:  ₹${out.worker.toLocaleString()}`);
    console.log(`   Staff:   ₹${out.staff.toLocaleString()}`);
    console.log(`   Total:   ₹${(out.worker + out.staff).toLocaleString()}`);

    return out;
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
      for (let r = headerRow + 1; r <= monthSheet.rowCount; r++) {
        const row = monthSheet.getRow(r);
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";
        const empName =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";

        if (!empName || empName.includes("TOTAL")) continue;

        const grossSalary = num(row.getCell(grossSalaryCol));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = [];
        }

        if (grossSalary > 0) {
          employees[key].push(grossSalary);
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

        const salary1 = num(row.getCell(salary1Col));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = { salaries: [], hasSeptSalary: false };
        }

        if (salary1 > 0) {
          employees[key].salaries.push(salary1);
        }

        if (monthKey === "Sep-25" && salary1 > 0) {
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

  const calculateOctoberAverageForStaff = (
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
      { salaries: number[]; hasSeptSalary: boolean }
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
      const salary1Col = 15;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const empName = row
          .getCell(empNameCol)
          .value?.toString()
          .trim()
          .toUpperCase();
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";

        if (!empName || empName === "" || empName.includes("TOTAL")) continue;

        const salary1 = num(row.getCell(salary1Col));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = { salaries: [], hasSeptSalary: false };
        }

        if (salary1 > 0) {
          employees[key].salaries.push(salary1);
        }

        if (monthKey === "Sep-25" && salary1 > 0) {
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

        const salary1 = num(row.getCell(salary1Col));
        const grossSal = num(row.getCell(grossSalCol));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = {
            grossSalaries: [],
            salary1Values: [],
            hasSeptSalary1: false,
          };
        }

        // Track SALARY1 for Sept condition
        if (salary1 > 0) {
          employees[key].salary1Values.push(salary1);
        }

        // Track GROSS SALARY for averaging
        if (grossSal > 0) {
          employees[key].grossSalaries.push(grossSal);
        }

        // Check September SALARY1 condition
        if (monthKey === "Sep-25" && salary1 > 0) {
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

  // A(HR) - Calculate per-employee October average from WD SALARY, then sum
  const calculateOctoberAverageForHR = (
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

    // Store each employee's monthly WD SALARY values
    const employees: Record<string, number[]> = {};

    for (const monthKey of monthsToAverage) {
      const monthSheet = (monthWiseWb || workerWb).getWorksheet(
        workerMap[monthKey]
      );
      if (!monthSheet) continue;

      // Find header row and columns
      let headerRow = -1;
      let wdSalaryCol = -1;
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

        // Find WD SALARY column
        row.eachCell((cell, c) => {
          const text = cell.value?.toString().toUpperCase() || "";
          if (text.includes("WD") && text.includes("SALARY")) {
            wdSalaryCol = c;
          }
        });

        if (headerRow > 0 && wdSalaryCol > 0) break;
      }

      if (headerRow < 0 || wdSalaryCol < 0) continue;

      // Read each employee's WD SALARY for this month
      for (let r = headerRow + 1; r <= monthSheet.rowCount; r++) {
        const row = monthSheet.getRow(r);
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";
        const empName =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";

        if (!empName || empName.includes("TOTAL")) continue;

        const wdSalary = num(row.getCell(wdSalaryCol));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = [];
        }

        if (wdSalary > 0) {
          employees[key].push(wdSalary);
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

    console.log(`\n📊 A(HR) Employee-Level Calculation:`);
    console.log(`   Employees processed: ${employeeCount}`);
    console.log(`   Total October A(HR): ₹${totalOctober.toLocaleString()}`);

    return totalOctober;
  };

  const extractStaffEmployees = (
    wb: ExcelJS.Workbook,
    sheetNames: string[]
  ): Record<string, EmployeeMonthlySalary> => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();

    for (const sheetName of sheetNames) {
      const ws = wb.getWorksheet(sheetName);
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

      const monthKey = sheetNameToMonthKey(sheetName);
      if (!monthKey) continue;

      const empIdCol = 2;
      const deptCol = 3;
      const empNameCol = 5;
      const salary1Col = 15;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const empName = row
          .getCell(empNameCol)
          .value?.toString()
          .trim()
          .toUpperCase();
        const empId = row.getCell(empIdCol).value?.toString().trim() || "";
        const dept = row.getCell(deptCol).value?.toString().trim() || "";

        if (!empName || empName === "" || empName.includes("TOTAL")) continue;

        const salary1 = num(row.getCell(salary1Col));
        const key = empId || empName;

        if (!employees[key]) {
          employees[key] = {
            name: empName,
            employeeCode: empId,
            department: dept,
            monthlySalaries: {},
            source: "Staff",
          };
          months.forEach((m) => (employees[key].monthlySalaries[m] = 0));
        }

        if (salary1 > 0) {
          employees[key].monthlySalaries[monthKey] = salary1;
        }

        if (!employees[key].employeeCode && empId) {
          employees[key].employeeCode = empId;
        }
        if (!employees[key].department && dept) {
          employees[key].department = dept;
        }
      }
    }

    return employees;
  };

  const extractWorkerEmployees = (
    wb: ExcelJS.Workbook,
    sheetNames: string[]
  ): Record<string, EmployeeMonthlySalary> => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();

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

        if (!employees[key]) {
          employees[key] = {
            name,
            employeeCode: id,
            department: dept,
            monthlySalaries: {},
            monthlyDepartments: {}, // Add this
            source: "Worker",
          };
          months.forEach((m) => (employees[key].monthlySalaries[m] = 0));
        }

        if (s1 > 0) {
          employees[key].monthlySalaries[monthKey] = s1;
        }

        // Track department for each month
        if (monthKey && dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }

        // Update latest department
        if (!employees[key].department && dept) {
          employees[key].department = dept;
        }
      }
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
    };

    const ws = bonusWb.getWorksheet(sheetName);
    if (!ws) return employees;

    let headerRow = -1;
    for (let r = 1; r <= 5; r++) {
      const row = ws.getRow(r);
      const cellText = row.getCell(2).value?.toString().toUpperCase() || "";
      if (cellText.includes("EMP") && cellText.includes("CODE")) {
        headerRow = r;
        break;
      }
    }

    if (headerRow < 0) return employees;

    const empCodeCol = 2;
    const deptCol = 3;
    const empNameCol = 4;

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const code = row.getCell(empCodeCol).value?.toString().trim() || "";
      const name =
        row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";
      const dept = row.getCell(deptCol).value?.toString().trim() || "";

      if (!name || name === "GRAND TOTAL") continue;

      const key = code || name;

      if (!employees[key]) {
        employees[key] = {
          name,
          employeeCode: code,
          department: dept,
          monthlySalaries: {},
          source: "HR",
        };
        months.forEach((m) => (employees[key].monthlySalaries[m] = 0));
      }

      for (const [mKey, colNum] of Object.entries(monthColMap)) {
        const v = num(row.getCell(colNum));
        if (v > 0) {
          const currentVal = employees[key].monthlySalaries[mKey] || 0;
          if (currentVal === 0 || v > currentVal) {
            employees[key].monthlySalaries[mKey] = v;
          }
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
  const allKeys = new Set([...Object.keys(actualByKey), ...Object.keys(hrByKey)]);
  const out: EmployeeComparison[] = [];

  for (const key of allKeys) {
    const a = actualByKey[key];
    const h = hrByKey[key];

    const name = (a?.name || h?.name || "").toUpperCase();
    const code = a?.employeeCode || h?.employeeCode || "";
    const dept = a?.department || h?.department || "";

    const actualSalaries: Record<string, number> = {};
    const hrSalaries: Record<string, number> = {};
    const monthsWithMismatch: string[] = [];
    let hasMismatch = false;

    for (const m of months) {
      let av = a?.monthlySalaries[m] || 0;
      const hv = h?.monthlySalaries[m] || 0;

      // Get department for this specific month
      const monthDept = (a?.monthlyDepartments?.[m] || dept).toUpperCase();
      
      // If department is C for this month, treat salary1 as null (0)
      if (monthDept === "C") {
        av = 0; // Ignore the actual salary1 value
      }

      actualSalaries[m] = av;
      hrSalaries[m] = hv;

      // Check if this month should be ignored for mismatch detection
      const shouldIgnoreMonth =
        ["C", "A"].includes(monthDept) || code.toUpperCase() === "N";

      // Only flag as mismatch if not in C/A department and not employee N
      if (Math.abs(av - hv) > 1 && !shouldIgnoreMonth) {
        hasMismatch = true;
        monthsWithMismatch.push(m);
      }
    }

    const totalActual = Object.values(actualSalaries).reduce((s, v) => s + v, 0);
    const totalHR = Object.values(hrSalaries).reduce((s, v) => s + v, 0);

    out.push({
      name,
      employeeCode: code,
      department: dept,
      actualSalaries,
      hrSalaries,
      hasMismatch,
      missingInHR: !h,
      missingInActual: !a,
      totalActual,
      totalHR,
      totalDifference: totalActual - totalHR,
      monthsWithMismatch,
      monthlyDepartments: a?.monthlyDepartments,
    });
  }

  out.sort((x, y) => {
    if (x.missingInHR !== y.missingInHR) return x.missingInHR ? -1 : 1;
    if (x.missingInActual !== y.missingInActual) return x.missingInActual ? -1 : 1;
    if (x.hasMismatch !== y.hasMismatch) return x.hasMismatch ? -1 : 1;
    return 0;
  });

  return out;
};


  const runSalaryComparison = async () => {
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
      const calculateOctoberAverageForActualEmployees = (
        employees: Record<string, EmployeeMonthlySalary>
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
          const values: number[] = [];

          // Collect all salary values from Nov-24 to Sep-25
          for (const month of monthsToAverage) {
            const salary = emp.monthlySalaries[month] || 0;

            // Get department for this specific month
            const monthDept = emp.monthlyDepartments?.[month] || emp.department;

            // Ignore salary1 if department is C for this month
            if (monthDept.toUpperCase() === "C") {
              continue; // Skip this month's salary
            }

            if (salary > 0) {
              values.push(salary);
            }
          }

          // Calculate average and assign to October
          if (values.length > 0) {
            const average =
              values.reduce((sum, val) => sum + val, 0) / values.length;
            emp.monthlySalaries["Oct-25"] = Math.round(average);
          } else {
            // If all months were ignored (all dept C), set October to 0
            emp.monthlySalaries["Oct-25"] = 0;
          }
        }
      };

      // Apply October calculation to both staff and worker actual data
      calculateOctoberAverageForActualEmployees(staffEmployees);
      calculateOctoberAverageForActualEmployees(workerEmployees);

      const hrStaffEmployees = extractHREmployees(bonusWb, "Staff");
      const hrWorkerEmployees = extractHREmployees(bonusWb, "Worker");

      // ✅ NEW: Calculate October average for HR employees (Nov-24 to Sep-25)
      const calculateOctoberAverageForHREmployees = (
        employees: Record<string, EmployeeMonthlySalary>
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
          const values: number[] = [];

          // Collect all salary values from Nov-24 to Sep-25
          for (const month of monthsToAverage) {
            const salary = emp.monthlySalaries[month] || 0;
            if (salary > 0) {
              values.push(salary);
            }
          }

          // Calculate average and assign to October
          if (values.length > 0) {
            const average =
              values.reduce((sum, val) => sum + val, 0) / values.length;
            emp.monthlySalaries["Oct-25"] = Math.round(average);
          }
        }
      };

      // Apply October calculation to both staff and worker HR data
      calculateOctoberAverageForHREmployees(hrStaffEmployees);
      calculateOctoberAverageForHREmployees(hrWorkerEmployees);

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
          // A value calculation
          const A_v1 = calculateOctoberAverageForStaff(staffWb, staffMap);
          const A_v2 = calculateOctoberAverageForHR(
            monthWiseWb,
            workerWb,
            workerMap
          );
          const A_diff = A_v2 - A_v1;

          // B value calculation
          // B value calculation (EXISTING)
          const B_staffOct = calculateOctoberAverageForStaff(staffWb, staffMap);
          const B_workerOct = calculateOctoberAverageForWorker(
            workerWb,
            workerMap
          );
          const B_v1 = B_staffOct + B_workerOct;
          let B_v2 = 0;
          if (bonusWb) {
            const bonusOctober = getBonusOctoberTotals(bonusWb);
            B_v2 = bonusOctober.staff + bonusOctober.worker;
          }

          // NEW: Calculate October B Extras as average of 11 months
          let totalExtras = 0;
          const extrasMonths = [
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
          for (const em of extrasMonths) {
            totalExtras += calculateBExtras(workerWb, workerMap, em);
          }
          const B_extras = totalExtras / extrasMonths.length;

          // NEW: Update diff calculation
          const B_diff = B_v2 - B_v1 + B_extras;

          // NEW: C value calculation
          const C_staffGross = calculateStaffGrossOctober(staffWb, staffMap);
          const C_workerSalary1 = calculateOctoberAverageForWorker(
            workerWb,
            workerMap
          );
          const C_v1 = C_staffGross + C_workerSalary1;

          // C(HR) = Average of GROSS SALARY from Month-Wise Sheet across 11 months
          const C_v2 = getCHR(monthWiseWb, workerWb, workerMap); // ✅ CORRECT - passing monthWiseWb
          const C_diff = C_v2 - C_v1;

          const D_v1 = C_v1; // Reuse C(Software) value

          // D(HR) = Average of GROSS SALARY from Bonus Summary (Nov-24 to Sep-25)
          const D_v2 = getDHR(bonusSummaryWb);
          const D_diff = D_v2 - D_v1;

          const E_v1 = D_v1 * 0.0833; // 8.33% of D(Software)

          // E(HR) = Average of FD from Bonus Summary (Nov-24 to Sep-25)
          const E_v2 = getEHR(bonusSummaryWb);
          const E_diff = E_v2 - E_v1;

          // Update the software, hr, and diff objects for October
          software[m] = {
            A: round(A_v1),
            B: round(B_v1),
            C: round(C_v1),
            D: round(D_v1),
            E: round(E_v1), // 8.33% of D(Software)
          };
          hr[m] = {
            A: round(A_v2),
            B: round(B_v2),
            C: round(C_v2),
            D: round(D_v2),
            E: round(E_v2), // Average FD from Bonus Summary
          };
          extras[m] = {
            // ✅ This line should already be here
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

        let A_v1 = 0;
        if (staffSheet) A_v1 = readStaffSalary1Total(staffSheet);
        let A_v2 = 0;
        if (workerMonthSheet) {
          const { col, headerRow } = findColByHeader(workerMonthSheet, [
            "WD",
            "SALARY",
          ]);
          A_v2 = readColumnGrandTotal(workerMonthSheet, col, headerRow);
        }
        const A_diff = A_v2 - A_v1;

        let B_v1_staff = 0,
          B_v1_worker = 0;
        if (staffSheet) B_v1_staff = readStaffSalary1Total(staffSheet);
        const workerSheetForMonth = workerWb.getWorksheet(workerMap[m]);
        if (workerSheetForMonth)
          B_v1_worker = sumWorkerSalary1(workerSheetForMonth);
        const B_v1 = B_v1_staff + B_v1_worker;
        const B_v2 = bonusMonthly[m]
          ? bonusMonthly[m].worker + bonusMonthly[m].staff
          : 0;

        // NEW: Calculate B Extras
        const B_extras = calculateBExtras(workerWb, workerMap, m);

        // NEW: Update diff calculation to include extras
        const B_diff = B_v2 - B_v1 + B_extras; // Changed formula

        let C_g1_worker = 0,
          C_g1_staffGross = 0;
        if (workerSheetForMonth)
          C_g1_worker = sumWorkerSalary1(workerSheetForMonth);
        if (staffSheet) C_g1_staffGross = readStaffGrossTotal(staffSheet);
        const C_v1 = C_g1_worker + C_g1_staffGross;

        let C_v2 = 0;
        const monthWiseMonth = (monthWiseWb || workerWb).getWorksheet(
          workerMap[m]
        );
        if (monthWiseMonth) C_v2 = getGrossSalaryGrandTotal(monthWiseMonth);
        const C_diff = C_v2 - C_v1;

        let D1_worker = 0,
          D1_staffGross = 0;
        if (workerSheetForMonth)
          D1_worker = sumWorkerSalary1(workerSheetForMonth);
        if (staffSheet) D1_staffGross = readStaffGrossTotal(staffSheet);
        const D_v1 = D1_worker + D1_staffGross;
        const D_v2 = bonusSummaryMonthlyGross[m] ?? 0;

        let E1_worker = 0,
          E1_staffGross = 0;
        if (workerSheetForMonth)
          E1_worker = sumWorkerSalary1(workerSheetForMonth);
        if (staffSheet) E1_staffGross = readStaffGrossTotal(staffSheet);
        const E_v1 = (E1_worker + E1_staffGross) * 0.0833;
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
          { name: "A,N,C", data: extras }, // NEW ROW
          { name: "diff", data: diff },
        ],
      });
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
            <button
              onClick={processExcelFiles}
              disabled={!staffFile || !workerFile || isGenerating}
              className="px-8 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center gap-2"
            >
              {isGenerating ? "Generating Report..." : "Generate Report"}
            </button>

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
                  Run Comparison
                </>
              )}
            </button>
            {(reportData || comparisonResults) && (
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
