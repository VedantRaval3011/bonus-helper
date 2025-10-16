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

interface MonthlyTotals {
  software: number;
  hr: number;
  difference: number;
}

export default function Step2Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [reportData, setReportData] = useState<ReportData | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isComparing, setIsComparing] = useState(false);
  const [comparisonResults, setComparisonResults] =
    useState<ComparisonResults | null>(null);
  const [monthlyTotals, setMonthlyTotals] = useState<Record<string, MonthlyTotals> | null>(null);
  const [activeTab, setActiveTab] = useState<"staff" | "worker">("staff");
  const [isLoadingTab, setIsLoadingTab] = useState(false);
  const [ignoredEmployees, setIgnoredEmployees] = useState<Set<string>>(
    new Set()
  );
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [password, setPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");

  const handleTabSwitch = async (tab: "staff" | "worker") => {
    setIsLoadingTab(true);
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

const bonusFile =
  pickFile((s) => s.type === "Bonus-Calculation-Sheet") ??
  pickFile(
    (s) =>
      !!s.file &&
      /bonus.*final.*calculation|bonus.*2024-25|sci.*prec.*final.*calculation|final.*calculation.*sheet|nrtm.*final.*bonus.*calculation|nutra.*bonus.*calculation|sci.*prec.*life.*science.*bonus.*calculation/i.test(
        s.file.name
      )
  );


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
      .match(/(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JUL|JULY|AUG|SEP)[-\s]?(\d{2})/);
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

  const getCellNumericValue = (
    cell: ExcelJS.Cell
  ): { hasValue: boolean; value: number } => {
    const v: any = cell.value;

    if (v === null || v === undefined) {
      return { hasValue: false, value: 0 };
    }

    if (typeof v === "object" && v !== null) {
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
      if ("sharedFormula" in v && !("result" in v)) {
        // Fall through to cell.text check
      } else {
        return { hasValue: false, value: 0 };
      }
    }

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

    if (typeof v === "number") {
      return { hasValue: true, value: v };
    }

    const t = cell.text?.trim();
    if (typeof v === "object" && "sharedFormula" in v && (!t || t === "")) {
      return { hasValue: true, value: 0 };
    }

    if (!t || t === "-") {
      return { hasValue: false, value: 0 };
    }

    const n = Number(t.replace(/,/g, ""));
    return Number.isFinite(n)
      ? { hasValue: true, value: n }
      : { hasValue: false, value: 0 };
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

  // Extract SALARY1 values from Staff Tulsi files (Software data)
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

      const empIdCol = 2;    // Column B
      const deptCol = 3;     // Column C
      const empNameCol = 4;  // Column D
      const salary1Col = 11; // Column K (SALARY1)

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

        const cell = row.getCell(salary1Col);
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

        if (s1.hasValue) {
          employees[key].monthlySalaries[monthKey] = s1.value;
        }

        if (dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }
      }
    }

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

  // Extract SALARY1 values from Worker Tulsi files (Software data)
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

      const empIdCol = 2;    // Column B
      const deptCol = 3;     // Column C
      const empNameCol = 4;  // Column D
      const salary1Col = 11; // Column K (SALARY1)

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name =
          row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";
        const id = row.getCell(empIdCol).value?.toString().trim() || "";
        const dept = row.getCell(deptCol).value?.toString().trim() || "";

        if (!name || name.includes("TOTAL")) continue;

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
          employees[key].monthlySalaries[monthKey] = res.value;
        }
        if (monthKey && dept) {
          employees[key].monthlyDepartments![monthKey] = dept;
        }
      }
    }

    for (const key in employees) {
      const emp = employees[key];
      let finalDept = "";

      const monthsReversed = [...months].reverse();

      for (const month of monthsReversed) {
        const dept = emp.monthlyDepartments?.[month];
        if (dept && !["C", "A"].includes(dept.toUpperCase())) {
          finalDept = dept;
          break;
        }
      }

      if (!finalDept && emp.monthlyDepartments?.["Sep-25"]) {
        finalDept = emp.monthlyDepartments["Sep-25"];
      }

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

  // Extract HR employees from sci-prc-final calculation sheet
  const extractHREmployeesFromFinalCalc = (
    finalCalcWb: ExcelJS.Workbook
  ): Record<string, EmployeeMonthlySalary> => {
    const employees: Record<string, EmployeeMonthlySalary> = {};
    const months = generateMonthHeaders();

    const ws = finalCalcWb.getWorksheet("Sheet1") ?? finalCalcWb.worksheets[0];
    if (!ws) return employees;

    let headerRow = -1;
    for (let r = 1; r <= 5; r++) {
      const row = ws.getRow(r);
      const cell = row.getCell(2).value?.toString().toUpperCase();
      if (cell?.includes("EMP") && cell?.includes("CODE")) {
        headerRow = r;
        break;
      }
    }

    if (headerRow <= 0) return employees;

    const empCodeCol = 2;  // Column B
    const deptCol = 3;     // Column C
    const empNameCol = 4;  // Column D

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

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const code = row.getCell(empCodeCol).value?.toString().trim() || "";
      const name =
        row.getCell(empNameCol).value?.toString().trim().toUpperCase() || "";
      const dept = row.getCell(deptCol).value?.toString().trim() || "";

      if (!name || !code || name === "GRAND TOTAL") continue;

      const key = `${code}|${name}`;

      if (!employees[key]) {
        employees[key] = {
          name,
          employeeCode: code,
          department: dept,
          monthlySalaries: {},
          source: "HR",
        };
      }

      let sum = 0;
      let count = 0;

      for (const [mKey, colNum] of Object.entries(monthColMap)) {
        const v = num(row.getCell(colNum));
        if (v > 0) {
          employees[key].monthlySalaries[mKey] = v;
          sum += v;
          count++;
        }
      }

      // Calculate Oct-25 as average of all 11 months
      if (count > 0) {
        employees[key].monthlySalaries["Oct-25"] =
          Math.round((sum / count) * 100) / 100;
      }
    }

    return employees;
  };

  const compareEmployees = (
    actualByKey: Record<string, EmployeeMonthlySalary>,
    hrByKey: Record<string, EmployeeMonthlySalary>,
    months: string[]
  ): EmployeeComparison[] => {
    const groupedActual: Record<string, EmployeeMonthlySalary[]> = {};
    const groupedHR: Record<string, EmployeeMonthlySalary[]> = {};

    for (const key of Object.keys(actualByKey)) {
      const emp = actualByKey[key];
      const code = emp.employeeCode;
      if (!groupedActual[code]) {
        groupedActual[code] = [];
      }
      groupedActual[code].push(emp);
    }

    for (const key of Object.keys(hrByKey)) {
      const emp = hrByKey[key];
      const code = emp.employeeCode;
      if (!groupedHR[code]) {
        groupedHR[code] = [];
      }
      groupedHR[code].push(emp);
    }

    const allCodes = new Set([
      ...Object.keys(groupedActual),
      ...Object.keys(groupedHR),
    ]);

    const out: EmployeeComparison[] = [];

    for (const code of allCodes) {
      const actualEmps = groupedActual[code] || [];
      const hrEmps = groupedHR[code] || [];

      const firstActual = actualEmps[0];
      const firstHR = hrEmps[0];

      const name = (firstActual?.name || firstHR?.name).toUpperCase();
      const dept = firstActual?.department || firstHR?.department;

      const actualSalaries: Record<string, number> = {};
      const hrSalaries: Record<string, number> = {};
      const monthsWithMismatch: string[] = [];
      let hasMismatch = false;

      for (const m of months) {
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

        let hv = 0;
        for (const emp of hrEmps) {
          hv += emp?.monthlySalaries?.[m] || 0;
        }

        actualSalaries[m] = av;
        hrSalaries[m] = hv;

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

    out.sort((x, y) => {
      if (x.missingInHR !== y.missingInHR) return x.missingInHR ? -1 : 1;
      if (x.missingInActual !== y.missingInActual)
        return x.missingInActual ? -1 : 1;
      if (x.hasMismatch !== y.hasMismatch) return x.hasMismatch ? -1 : 1;
      return 0;
    });

    return out;
  };

  const calculateMonthlyTotals = (
    staffComparisons: EmployeeComparison[],
    workerComparisons: EmployeeComparison[],
    months: string[]
  ): Record<string, MonthlyTotals> => {
    const totals: Record<string, MonthlyTotals> = {};

    for (const month of months) {
      let softwareTotal = 0;
      let hrTotal = 0;

      // Calculate for staff
      staffComparisons.forEach((emp) => {
        softwareTotal += emp.actualSalaries[month] || 0;
        hrTotal += emp.hrSalaries[month] || 0;
      });

      // Calculate for workers
      workerComparisons.forEach((emp) => {
        softwareTotal += emp.actualSalaries[month] || 0;
        hrTotal += emp.hrSalaries[month] || 0;
      });

      totals[month] = {
        software: Math.round(softwareTotal),
        hr: Math.round(hrTotal),
        difference: Math.round(softwareTotal - hrTotal),
      };
    }

    return totals;
  };

  const runSalaryComparison = async () => {
    if (!staffFile || !workerFile || !bonusFile) {
      alert(
        "Please upload Staff Tulsi, Worker Tulsi, and Sci-Prc-Final Calculation files for comparison"
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
            /(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JULY|AUG|SEP)[-\s]?\d{2}/i.test(
              name
            ) &&
            /SP/i.test(name) &&
            !/OCT/i.test(name)
        );

      const workerSheetNames = workerWb.worksheets
        .map((ws) => ws.name)
        .filter(
          (name) =>
            /(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JULY|AUG|SEP)[-\s]?\d{2}/i.test(
              name
            ) &&
            /SP/i.test(name) &&
            !/OCT/i.test(name)
        );

      console.log("Staff sheets:", staffSheetNames);
      console.log("Worker sheets:", workerSheetNames);

      const staffEmployees = extractStaffEmployees(staffWb, staffSheetNames);
      const workerEmployees = extractWorkerEmployees(
        workerWb,
        workerSheetNames
      );

      // Calculate October average for software employees
      const calculateOctoberAverage = (
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

            if (salary === undefined || salary === null) {
              continue;
            }

            const monthDept = emp.monthlyDepartments?.[month];
            if (!monthDept) continue;

            const md = monthDept.toUpperCase();
            if (md === "C" || md !== finalDept) continue;

            values.push(salary);
          }

          if (values.length > 0) {
            const average =
              values.reduce((sum, val) => sum + val, 0) / values.length;
            emp.monthlySalaries["Oct-25"] = Math.round(average * 2) / 2;

            if (emp.monthlyDepartments) {
              emp.monthlyDepartments["Oct-25"] =
                emp.monthlyDepartments["Sep-25"] || emp.department;
            }
          } else {
            if (emp.monthlyDepartments) {
              emp.monthlyDepartments["Oct-25"] =
                emp.monthlyDepartments["Sep-25"] || emp.department;
            }
          }
        }
      };

      calculateOctoberAverage(staffEmployees);
      calculateOctoberAverage(workerEmployees);

      // Extract HR data from sci-prc-final calculation sheet
      const hrAllEmployees = extractHREmployeesFromFinalCalc(bonusWb);

      const hrStaffEmployees: Record<string, EmployeeMonthlySalary> = {};
      const hrWorkerEmployees: Record<string, EmployeeMonthlySalary> = {};

      for (const key in hrAllEmployees) {
        const emp = hrAllEmployees[key];
        if (emp.department.toUpperCase() === "S") {
          hrStaffEmployees[key] = emp;
        } else if (emp.department.toUpperCase() === "W") {
          hrWorkerEmployees[key] = emp;
        }
      }

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

      // Calculate monthly totals
      const totals = calculateMonthlyTotals(
        staffComparisons,
        workerComparisons,
        months
      );

      setComparisonResults({
        staffComparisons,
        workerComparisons,
      });
      setMonthlyTotals(totals);

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
      ...months.flatMap((m) => [`${m} (Software)`, `${m} (HR)`, `${m} (Diff)`]),
      "Total Software",
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
      ...months.flatMap((m) => [`${m} (Software)`, `${m} (HR)`, `${m} (Diff)`]),
      "Total Software",
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

  const handleGenerateReport = async () => {
    if (!monthlyTotals || !comparisonResults) {
      alert("Please run comparison first!");
      return;
    }

    setIsGenerating(true);

    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Monthly Comparison Report");

      const months = generateMonthHeaders();

      // Header row
      const headerRow = ["Source", ...months];
      ws.addRow(headerRow);
      ws.getRow(1).font = { bold: true, size: 12 };
      ws.getRow(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4472C4" },
      };
      ws.getRow(1).font = { ...ws.getRow(1).font, color: { argb: "FFFFFFFF" } };

      // Software row
      const softwareRow = [
        "Software (SALARY1)",
        ...months.map((m) => monthlyTotals[m]?.software || 0),
      ];
      ws.addRow(softwareRow);
      ws.getRow(2).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE7E6E6" },
      };

      // HR row
      const hrRow = [
        "HR (Sci-Prc-Final)",
        ...months.map((m) => monthlyTotals[m]?.hr || 0),
      ];
      ws.addRow(hrRow);
      ws.getRow(3).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFFFF" },
      };

      // Difference row
      const diffRow = [
        "Difference",
        ...months.map((m) => monthlyTotals[m]?.difference || 0),
      ];
      ws.addRow(diffRow);
      ws.getRow(4).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFC7CE" },
      };
      ws.getRow(4).font = { bold: true };

      // Format number columns
      for (let col = 2; col <= months.length + 1; col++) {
        ws.getColumn(col).numFmt = "#,##0";
        ws.getColumn(col).width = 15;
      }

      // Format first column
      ws.getColumn(1).width = 25;

      // Add borders
      for (let row = 1; row <= 4; row++) {
        for (let col = 1; col <= months.length + 1; col++) {
          const cell = ws.getCell(row, col);
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }

      // Download the file
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `monthly_comparison_report_${
        new Date().toISOString().split("T")[0]
      }.xlsx`;
      a.click();

      alert("Report generated successfully!");
    } catch (error) {
      console.error("Error generating report:", error);
      alert("Error generating report. Please check console for details.");
    } finally {
      setIsGenerating(false);
    }
  };

  const toggleIgnoreSpecialDepts = () => {
    if (!comparisonResults) return;

    const specialEmployees = new Set<string>();

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
      setIgnoredEmployees(new Set());
    } else {
      setIgnoredEmployees(specialEmployees);
    }
  };

  const handleMoveToStep3 = () => {
    router.push("step3");
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-5 px-4">
      <div className="mx-auto max-w-7xl">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 2 - Salary Comparison
              </h1>
              <p className="text-gray-600 mt-2">
                Compare Software data (SALARY1 from Tulsi) vs HR data
                (Sci-Prc-Final Calculation Sheet)
              </p>
            </div>
            <button
              onClick={() => router.push("/")}
              className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
            >
              ← Back to Step 1
            </button>
          </div>

          {/* File status cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
            {/* Staff File Card */}
            <div
              className={`border-2 rounded-lg p-6 ${
                staffFile
                  ? "border-green-300 bg-green-50"
                  : "border-gray-300 bg-gray-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                {staffFile ? (
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
                  <h3 className="text-lg font-bold text-gray-800">
                    Staff Tulsi
                  </h3>
                </div>
              </div>
              {staffFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {staffFile.name}
                    </p>
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-gray-200">
                  <p className="text-sm text-gray-600 font-medium">
                    No file uploaded
                  </p>
                </div>
              )}
            </div>

            {/* Worker File Card */}
            <div
              className={`border-2 rounded-lg p-6 ${
                workerFile
                  ? "border-green-300 bg-green-50"
                  : "border-gray-300 bg-gray-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                {workerFile ? (
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
                  <h3 className="text-lg font-bold text-gray-800">
                    Worker Tulsi
                  </h3>
                </div>
              </div>
              {workerFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {workerFile.name}
                    </p>
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-gray-200">
                  <p className="text-sm text-gray-600 font-medium">
                    No file uploaded
                  </p>
                </div>
              )}
            </div>

            {/* Final Calculation Sheet Card */}
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
                  <h3 className="text-lg font-bold text-gray-800">
                    Sci-Prc-Final Calculation
                  </h3>
                </div>
              </div>
              {bonusFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {bonusFile.name}
                    </p>
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-gray-200">
                  <p className="text-sm text-gray-600 font-medium">
                    No file uploaded
                  </p>
                </div>
              )}
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex gap-4 mb-8">
            <button
              onClick={runSalaryComparison}
              disabled={isComparing || !staffFile || !workerFile || !bonusFile}
              className={`flex-1 py-3 rounded-lg font-semibold transition ${
                isComparing || !staffFile || !workerFile || !bonusFile
                  ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                  : "bg-blue-600 text-white hover:bg-blue-700"
              }`}
            >
              {isComparing ? (
                <span className="flex items-center justify-center gap-2">
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
                </span>
              ) : (
                "Compare Salaries"
              )}
            </button>

            {monthlyTotals && (
              <button
                onClick={handleGenerateReport}
                disabled={isGenerating}
                className={`px-6 py-3 rounded-lg font-semibold transition ${
                  isGenerating
                    ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                    : "bg-purple-600 text-white hover:bg-purple-700"
                }`}
              >
                {isGenerating ? (
                  <span className="flex items-center justify-center gap-2">
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
                    Generating...
                  </span>
                ) : (
                  "📊 Generate Report"
                )}
              </button>
            )}

            {comparisonResults && (
              <button
                onClick={handleMoveToStep3}
                className="px-6 py-3 bg-green-600 text-white rounded-lg font-semibold hover:bg-green-700 transition"
              >
                Proceed to Step 3 →
              </button>
            )}
          </div>

          {/* Monthly Totals Summary Table */}
          {monthlyTotals && (
            <div className="mb-8 bg-blue-50 rounded-lg p-6 border-2 border-blue-200">
              <h2 className="text-xl font-bold text-gray-800 mb-4">
                Monthly Comparison Summary
              </h2>
              <div className="overflow-x-auto">
                <table className="w-full text-sm border-collapse">
                  <thead>
                    <tr className="bg-blue-600 text-white">
                      <th className="border border-blue-700 px-3 py-2 text-left">
                        Source
                      </th>
                      {generateMonthHeaders().map((m) => (
                        <th
                          key={m}
                          className="border border-blue-700 px-3 py-2 text-center"
                        >
                          {m}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="bg-gray-100">
                      <td className="border border-gray-300 px-3 py-2 font-semibold">
                        Software (SALARY1)
                      </td>
                      {generateMonthHeaders().map((m) => (
                        <td
                          key={m}
                          className="border border-gray-300 px-3 py-2 text-right"
                        >
                          {monthlyTotals[m]?.software.toLocaleString()}
                        </td>
                      ))}
                    </tr>
                    <tr className="bg-white">
                      <td className="border border-gray-300 px-3 py-2 font-semibold">
                        HR (Sci-Prc-Final)
                      </td>
                      {generateMonthHeaders().map((m) => (
                        <td
                          key={m}
                          className="border border-gray-300 px-3 py-2 text-right"
                        >
                          {monthlyTotals[m]?.hr.toLocaleString()}
                        </td>
                      ))}
                    </tr>
                    <tr className="bg-red-100">
                      <td className="border border-gray-300 px-3 py-2 font-bold">
                        Difference
                      </td>
                      {generateMonthHeaders().map((m) => (
                        <td
                          key={m}
                          className={`border border-gray-300 px-3 py-2 text-right font-bold ${
                            Math.abs(monthlyTotals[m]?.difference) > 10
                              ? "text-red-600"
                              : "text-green-600"
                          }`}
                        >
                          {monthlyTotals[m]?.difference.toLocaleString()}
                        </td>
                      ))}
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Comparison Results Table */}
          {comparisonResults && (
            <div className="mt-8">
              <div className="flex gap-4 mb-6">
                <button
                  onClick={() => handleTabSwitch("staff")}
                  className={`px-6 py-3 rounded-lg font-semibold transition ${
                    activeTab === "staff"
                      ? "bg-blue-600 text-white"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  Staff ({comparisonResults.staffComparisons.length})
                </button>
                <button
                  onClick={() => handleTabSwitch("worker")}
                  className={`px-6 py-3 rounded-lg font-semibold transition ${
                    activeTab === "worker"
                      ? "bg-blue-600 text-white"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  Worker ({comparisonResults.workerComparisons.length})
                </button>
                <button
                  onClick={toggleIgnoreSpecialDepts}
                  className={`px-6 py-3 rounded-lg font-semibold transition ml-auto ${
                    ignoredEmployees.size > 0
                      ? "bg-orange-600 text-white"
                      : "bg-gray-200 text-gray-700 hover:bg-gray-300"
                  }`}
                >
                  {ignoredEmployees.size > 0 ? "Show" : "Hide"} C, A, N
                </button>
              </div>

              {isLoadingTab ? (
                <div className="flex justify-center items-center py-20">
                  <svg
                    className="animate-spin h-12 w-12 text-blue-600"
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
                </div>
              ) : (
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
                          <th className="border px-3 py-2 text-left">Status</th>
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
                                Software
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
                        {(activeTab === "staff"
                          ? comparisonResults.staffComparisons
                          : comparisonResults.workerComparisons
                        ).map((emp, idx) => {
                          const isIgnoredEmployee = ignoredEmployees.has(
                            emp.employeeCode || emp.name
                          );

                          let hasAnyNonIgnoredMonth = false;
                          const months = generateMonthHeaders();

                          for (const m of months) {
                            const monthDept =
                              emp.monthlyDepartments?.[m] || emp.department;
                            const isIgnoredDept =
                              ["C", "A"].includes(monthDept.toUpperCase()) ||
                              emp.employeeCode.toUpperCase() === "N";

                            if (!isIgnoredDept) {
                              hasAnyNonIgnoredMonth = true;
                              break;
                            }
                          }

                          if (isIgnoredEmployee && !hasAnyNonIgnoredMonth) {
                            return null;
                          }

                          let hasRealErrors = false;
                          for (const m of months) {
                            const actualSal = emp.actualSalaries[m] || 0;
                            const hrSal = emp.hrSalaries[m] || 0;
                            const diff = Math.abs(actualSal - hrSal);

                            const monthDept =
                              emp.monthlyDepartments?.[m] || emp.department;
                            const isIgnoredMonth =
                              isIgnoredEmployee &&
                              (["C", "A"].includes(monthDept.toUpperCase()) ||
                                emp.employeeCode.toUpperCase() === "N");

                            if (diff > 1 && !isIgnoredMonth) {
                              hasRealErrors = true;
                              break;
                            }
                          }

                          const bgColor = emp.missingInHR
                            ? "bg-red-50"
                            : hasRealErrors
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
                              {months.map((m) => {
                                const actualSal = emp.actualSalaries[m] || 0;
                                const hrSal = emp.hrSalaries[m] || 0;
                                const diff = actualSal - hrSal;
                                const hasDiff = Math.abs(diff) > 1;

                                const monthDept =
                                  emp.monthlyDepartments?.[m] ||
                                  emp.department;
                                const shouldIgnoreThisMonth =
                                  ["C", "A"].includes(
                                    monthDept.toUpperCase()
                                  ) ||
                                  emp.employeeCode.toUpperCase() === "N";

                                const shouldGreyOut =
                                  isIgnoredEmployee && shouldIgnoreThisMonth;

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
                                          ? "bg-gray-200 text-gray-400"
                                          : hasDiff
                                          ? "bg-yellow-200 text-red-600"
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
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
