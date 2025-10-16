"use client";

import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

export default function Step8Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isLogicMinimized, setIsLogicMinimized] = useState(true);

  // 🎯 Sorting state
  const [sortConfig, setSortConfig] = useState<{
    key: string | null;
    direction: "asc" | "desc" | null;
  }>({ key: null, direction: null });

  type FileSlot = { type: string; file: File | null };

  const pickFile = (pred: (s: FileSlot) => boolean): File | null => {
    const slot = (fileSlots as FileSlot[]).find(pred);
    return slot?.file ?? null;
  };

  const staffFile =
    pickFile((s) => s.type === "Indiana-Staff") ??
    pickFile((s) => !!s.file && /staff/i.test(s.file.name));

const bonusFile =
  pickFile((s) => s.type === "Bonus-Calculation-Sheet") ??
  pickFile(
    (s) =>
      !!s.file &&
      /bonus.*final.*calculation|bonus.*2024-25|sci.*prec.*final.*calculation|final.*calculation.*sheet|nrtm.*final.*bonus.*calculation|nutra.*bonus.*calculation|sci.*prec.*life.*science.*bonus.*calculation/i.test(
        s.file.name
      )
  );

  const dueVoucherFile =
    pickFile((s) => s.type === "Due-Voucher-List") ??
    pickFile((s) => !!s.file && /due.*voucher/i.test(s.file.name));

  const actualPercentageFile =
    pickFile((s) => s.type === "Actual-Percentage-Bonus-Data") ??
    pickFile((s) => !!s.file && /actual.*percentage.*bonus/i.test(s.file.name));

  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  const MONTH_NAME_MAP: Record<string, number> = {
    JAN: 1,
    JANUARY: 1,
    FEB: 2,
    FEBRUARY: 2,
    MAR: 3,
    MARCH: 3,
    APR: 4,
    APRIL: 4,
    MAY: 5,
    JUN: 6,
    JUNE: 6,
    JUL: 7,
    JULY: 7,
    AUG: 8,
    AUGUST: 8,
    SEP: 9,
    SEPT: 9,
    SEPTEMBER: 9,
    OCT: 10,
    OCTOBER: 10,
    NOV: 11,
    NOVEMBER: 11,
    DEC: 12,
    DECEMBER: 12,
  };

  const pad2 = (n: number) => String(n).padStart(2, "0");

  const parseMonthFromSheetName = (sheetName: string): string | null => {
    const s = String(sheetName || "")
      .trim()
      .toUpperCase();

    const yyyymm = s.match(/(20\d{2})\D{0,2}(\d{1,2})/);
    if (yyyymm) {
      const y = Number(yyyymm[1]);
      const m = Number(yyyymm[2]);
      if (y >= 2000 && m >= 1 && m <= 12) return `${y}-${pad2(m)}`;
    }

    const mon = s.match(
      /\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b/
    );
    const monthFull = s.match(
      /\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/
    );
    const y2or4 = s.match(/\b(20\d{2}|\d{2})\b/);
    const monthToken = (monthFull?.[1] || mon?.[1]) as string | undefined;

    if (monthToken && y2or4) {
      let y = Number(y2or4[1]);
      if (y < 100) y += 2000;
      const m = MONTH_NAME_MAP[monthToken];
      if (m) return `${y}-${pad2(m)}`;
    }

    return null;
  };

  const AVG_WINDOW: string[] = [
    "2024-11",
    "2024-12",
    "2025-01",
    "2025-02",
    "2025-03",
    "2025-04",
    "2025-05",
    "2025-06",
    "2025-07",
    "2025-08",
    "2025-09",
  ];

  const EXCLUDED_MONTHS: string[] = ["2025-10", "2024-10"];
  const EXCLUDE_OCTOBER_EMPLOYEES = new Set<number>([
    937, 1039, 1065, 1105, 59, 161,
  ]);

  const TOLERANCE = 12;
  const REGISTER_PERCENTAGE = 8.33;
  const SPECIAL_PERCENTAGE = 12.0;
  const SPECIAL_GROSS_MULTIPLIER = 0.6; // 60% for 12% employees

  const calculatePercentageByMonths = (monthsWorked: number): number => {
    if (monthsWorked < 12) {
      return 10.0;
    } else if (monthsWorked >= 12 && monthsWorked < 24) {
      return 12.0;
    } else {
      return 8.33;
    }
  };

  const calculateMonthsFromDOJ = (dateValue: any): number => {
    if (!dateValue) return 0;

    let doj: Date;

    if (typeof dateValue === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      doj = new Date(excelEpoch.getTime() + dateValue * 86400000);
    } else if (typeof dateValue === "string") {
      const trimmed = dateValue.trim();
      const ddmmyyMatch = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);

      if (ddmmyyMatch) {
        let day = parseInt(ddmmyyMatch[1], 10);
        let month = parseInt(ddmmyyMatch[2], 10);
        let year = parseInt(ddmmyyMatch[3], 10);

        if (year < 100) {
          year += year < 50 ? 2000 : 1900;
        }

        doj = new Date(year, month - 1, day);
      } else {
        doj = new Date(dateValue);
      }
    } else {
      doj = new Date(dateValue);
    }

    if (isNaN(doj.getTime())) {
      return 0;
    }

    const referenceDate = new Date(2025, 9, 30);
    const yearsDiff = referenceDate.getFullYear() - doj.getFullYear();
    const monthsDiff = referenceDate.getMonth() - doj.getMonth();
    const daysDiff = referenceDate.getDate() - doj.getDate();

    let totalMonths = yearsDiff * 12 + monthsDiff;

    if (daysDiff < 0) {
      totalMonths--;
    }

    return Math.max(0, totalMonths);
  };

  const calculateGross2 = (grossSal: number, percentage: number): number => {
    if (percentage === 8.33) {
      return grossSal;
    } else if (percentage > 8.33) {
      return grossSal * 0.6;
    } else {
      return 0;
    }
  };

  const calculateActual = (
    grossSal: number,
    gross2: number,
    percentage: number
  ): number => {
    const pct = Number(percentage);

    if (pct === 8.33) {
      return (grossSal * pct) / 100;
    } else if (pct > 8.33) {
      return (gross2 * pct) / 100;
    } else {
      return 0;
    }
  };

  // === Step 8 Audit Helpers ===
  async function postAuditMessagesStep8(items: any[], batchId?: string) {
    const bid =
      batchId ||
      (typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2));
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId: bid, step: 8, items }),
    });
    return bid;
  }

  function buildStep8MismatchMessages(rows: any[]) {
    const items: any[] = [];
    for (const r of rows) {
      if (r?.status === "Mismatch") {
        items.push({
          level: "error",
          tag: "mismatch",
          text: `[step8] ${r.employeeId} ${r.employeeName} diff=${Number(
            r.difference ?? 0
          ).toFixed(2)}`,
          scope:
            r.department === "Staff"
              ? "staff"
              : r.department === "Worker"
              ? "worker"
              : "global",
          source: "step8",
          meta: {
            employeeId: r.employeeId,
            name: r.employeeName,
            department: r.department,
            grossSal: r.grossSal,
            adjustedGross: r.adjustedGross,
            excludedOctober: r.excludedOctober,
            registerPercentage: r.registerPercentage,
            actualPercentage: r.actualPercentage,
            percentageSource: r.percentageSource,
            registerCalculated: r.registerCalculated,
            actualCalculated: r.actualCalculated,
            reimSoftware: r.reimSoftware,
            reimHR: r.reimHR,
            paymentStatus: r.paymentStatus,
            diff: r.difference,
            tolerance: TOLERANCE,
          },
        });
      }
    }
    return items;
  }

  function buildStep8SummaryMessage(rows: any[]) {
    const total = rows.length || 0;
    const matches = rows.filter((r) => r.status === "Match").length;
    const mismatches = rows.filter((r) => r.status === "Mismatch").length;

    const staffRows = rows.filter((r) => r.department === "Staff");
    const workerRows = rows.filter((r) => r.department === "Worker");

    const staffMismatch = staffRows.filter(
      (r) => r.status === "Mismatch"
    ).length;
    const workerMismatch = workerRows.filter(
      (r) => r.status === "Mismatch"
    ).length;

    const customPercentageCount = rows.filter(
      (r) => r.percentageSource === "Custom"
    ).length;
    const zeroOctoberCount = rows.filter((r) => r.excludedOctober).length;
    const alreadyPaidCount = rows.filter(
      (r) => r.paymentStatus === "Already Paid"
    ).length;
    const unpaidCount = rows.filter((r) => r.paymentStatus === "Unpaid").length;

    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const staffGrossSalSum = sum(staffRows.map((r) => Number(r.grossSal || 0)));
    const staffRegisterSum = sum(
      staffRows.map((r) => Number(r.registerCalculated || 0))
    );
    const staffActualSum = sum(
      staffRows.map((r) => Number(r.actualCalculated || 0))
    );
    const staffReimSoftwareSum = sum(
      staffRows.map((r) => Number(r.reimSoftware || 0))
    );
    const staffReimHRSum = sum(staffRows.map((r) => Number(r.reimHR || 0)));

    const workerGrossSalSum = sum(
      workerRows.map((r) => Number(r.grossSal || 0))
    );
    const workerRegisterSum = sum(
      workerRows.map((r) => Number(r.registerCalculated || 0))
    );
    const workerActualSum = sum(
      workerRows.map((r) => Number(r.actualCalculated || 0))
    );
    const workerReimSoftwareSum = sum(
      workerRows.map((r) => Number(r.reimSoftware || 0))
    );
    const workerReimHRSum = sum(workerRows.map((r) => Number(r.reimHR || 0)));

    return {
      level: "info",
      tag: "summary",
      text: `Step8 run: total=${total} match=${matches} mismatch=${mismatches}`,
      scope: "global",
      source: "step8",
      meta: {
        totals: {
          total,
          matches,
          mismatches,
          tolerance: TOLERANCE,
          customPercentageCount,
          zeroOctoberCount,
          alreadyPaidCount,
          unpaidCount,
        },
        staff: {
          count: staffRows.length,
          mismatches: staffMismatch,
          grossSalSum: staffGrossSalSum,
          registerSum: staffRegisterSum,
          actualSum: staffActualSum,
          reimSoftwareSum: staffReimSoftwareSum,
          reimHRSum: staffReimHRSum,
        },
        worker: {
          count: workerRows.length,
          mismatches: workerMismatch,
          grossSalSum: workerGrossSalSum,
          registerSum: workerRegisterSum,
          actualSum: workerActualSum,
          reimSoftwareSum: workerReimSoftwareSum,
          reimHRSum: workerReimHRSum,
        },
      },
    };
  }

  useEffect(() => {
    if (typeof window === "undefined") return;
    if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

    const batchId = `step8-${Date.now()}-${Math.random()
      .toString(36)
      .slice(2, 8)}`;
    const items = [
      buildStep8SummaryMessage(comparisonData),
      ...buildStep8MismatchMessages(comparisonData),
    ];

    postAuditMessagesStep8(items, batchId).catch((err) =>
      console.error("Auto-audit step8 failed", err)
    );
  }, [comparisonData]);

  const processFiles = async () => {
    if (!staffFile || !bonusFile) {
      setError("Both Staff file and Bonus Calculation file are required");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log("📊 STEP 8: Reimbursement with Adj. Gross & Payment Status");
      // ========== LOAD DUE VOUCHER LIST ==========
      const dueVoucherMap = new Map<
        number,
        { dueVC: number; alreadyPaid: number; name: string; dept: string }
      >();

      if (dueVoucherFile) {
        console.log("\n📋 Loading Due Voucher List...");
        const dueVoucherBuffer = await dueVoucherFile.arrayBuffer();
        const dueVoucherWorkbook = XLSX.read(dueVoucherBuffer);

        const dueVoucherSheet =
          dueVoucherWorkbook.Sheets[dueVoucherWorkbook.SheetNames[0]];
        const dueVoucherData: any[][] = XLSX.utils.sheet_to_json(
          dueVoucherSheet,
          { header: 1 }
        );

        // Find header row
        let headerIdx = -1;
        for (let i = 0; i < dueVoucherData.length; i++) {
          const row = dueVoucherData[i];
          if (
            row &&
            row.some((cell: any) =>
              /EMP.*CODE|EMPCODE/i.test(String(cell ?? ""))
            )
          ) {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx !== -1) {
          const headers = dueVoucherData[headerIdx];
          const empCodeIdx = headers.findIndex((h: any) =>
            /EMP.*CODE|EMPCODE/i.test(String(h ?? ""))
          );
          const empNameIdx = headers.findIndex((h: any) =>
            /EMP.*NAME/i.test(String(h ?? ""))
          );
          const deptIdx = headers.findIndex((h: any) =>
            /DEPT/i.test(String(h ?? ""))
          );
          const dueVCIdx = headers.findIndex((h: any) =>
            /DUE.*VC/i.test(String(h ?? ""))
          );
          const alreadyPaidIdx = headers.findIndex((h: any) =>
            /ALREADY.*PAID/i.test(String(h ?? ""))
          );

          console.log(`\n✅ Processing Due Voucher List:`);
          for (let i = headerIdx + 1; i < dueVoucherData.length; i++) {
            const row = dueVoucherData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            if (!empCode || isNaN(empCode)) continue;

            const empName = String(row[empNameIdx] || "").trim();
            const dept = String(row[deptIdx] || "").trim();
            const dueVC = Number(row[dueVCIdx]) || 0;
            const alreadyPaid = Number(row[alreadyPaidIdx]) || 0;

            // Only add employees with Due VC > 0 (these are unpaid)
            if (dueVC > 0) {
              dueVoucherMap.set(empCode, {
                dueVC: dueVC,
                alreadyPaid: alreadyPaid,
                name: empName,
                dept: dept,
              });
              console.log(`   Emp ${empCode}: Due VC = ₹${dueVC} → UNPAID`);
            }
          }
        }

        console.log(
          `\n📊 Loaded ${dueVoucherMap.size} employees with Due VC (unpaid)`
        );
      } else {
        console.log("\n⚠️ Due Voucher List file not found");
      }

      console.log("=".repeat(60));

      // ========== LOAD ACTUAL PERCENTAGE FILE ==========
      const customPercentageMap = new Map<number, number>();
      const zeroOctoberEmployees = new Set<number>();
      const employeesInPerSheet = new Set<number>(); // NEW: Track employees in Per sheet

      if (actualPercentageFile) {
        console.log("\n📋 Loading Actual Percentage Bonus Data...");
        const percentageBuffer = await actualPercentageFile.arrayBuffer();
        const percentageWorkbook = XLSX.read(percentageBuffer);

        const perSheet = percentageWorkbook.Sheets["Per"];
        if (perSheet) {
          const perData: any[][] = XLSX.utils.sheet_to_json(perSheet, {
            header: 1,
          });

          console.log(
            "\n✅ Processing 'Per' sheet (Custom Percentages for ACTUAL):"
          );
          for (let i = 1; i < perData.length; i++) {
            const row = perData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[1]);
            const bonusPercentage = Number(row[4]);

            if (
              empCode &&
              !isNaN(empCode) &&
              bonusPercentage &&
              !isNaN(bonusPercentage)
            ) {
              customPercentageMap.set(empCode, bonusPercentage);
              employeesInPerSheet.add(empCode); // Track this employee
              console.log(
                `   Emp ${empCode}: Custom percentage for ACTUAL = ${bonusPercentage}% (will use Adj. Gross for Register)`
              );
            }
          }
        }

        const avgSheet = percentageWorkbook.Sheets["Average"];
        if (avgSheet) {
          const avgData: any[][] = XLSX.utils.sheet_to_json(avgSheet, {
            header: 1,
          });

          console.log("\n✅ Processing 'Average' sheet (Zero October):");
          for (let i = 1; i < avgData.length; i++) {
            const row = avgData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[1]);

            if (empCode && !isNaN(empCode)) {
              zeroOctoberEmployees.add(empCode);
              console.log(
                `   Emp ${empCode}: October excluded from GROSS & Register`
              );
            }
          }
        }

        console.log(
          `\n📊 Loaded ${customPercentageMap.size} custom percentages for ACTUAL`
        );
        console.log(
          `📊 Loaded ${zeroOctoberEmployees.size} zero-October employees`
        );
        console.log(
          `📊 Loaded ${employeesInPerSheet.size} employees in Per sheet (will use Adj. Gross for Register)`
        );
      } else {
        console.log(
          "\n⚠️ Actual Percentage file not found - using calculated percentages"
        );
      }

      // ========== LOAD BONUS FILE ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const hrReimDataMap: Map<
        number,
        {
          reimHR: number;
          unpaidHR: number;
          alreadyPaidHR: number;
          name: string;
          dept: string;
        }
      > = new Map();

      // Process Staff sheet (2nd sheet)
      if (bonusWorkbook.SheetNames.length > 1) {
        const staffSheetName = bonusWorkbook.SheetNames[1];
        const staffSheet = bonusWorkbook.Sheets[staffSheetName];
        const staffData: any[][] = XLSX.utils.sheet_to_json(staffSheet, {
          header: 1,
        });

        console.log(`\n📄 Processing bonus sheet: ${staffSheetName}`);

        for (let rowIdx = 0; rowIdx < staffData.length; rowIdx++) {
          const row = staffData[rowIdx];
          if (!row || row.length === 0) continue;

          const hasEmpCode = row.some((cell: any) =>
            /EMP.*CODE|EMPCODE/i.test(String(cell ?? ""))
          );

          if (hasEmpCode) {
            const headers = row;

            const empCodeIdx = headers.findIndex((h: any) =>
              /EMP.*CODE|EMPCODE/i.test(String(h ?? ""))
            );
            const deptIdx = headers.findIndex((h: any) =>
              /DEPTT|DEPT|DEPARTMENT/i.test(String(h ?? ""))
            );
            const empNameIdx = headers.findIndex((h: any) =>
              /EMP.*NAME|EMPNAME|EMPLOYEE.*NAME/i.test(String(h ?? ""))
            );
            const reimIdx = headers.findIndex((h: any) =>
              /^REIM\.?$/i.test(String(h ?? "").trim())
            );
            const unpaidIdx = headers.findIndex((h: any) =>
              /^UNPAID$/i.test(String(h ?? "").trim())
            );
            const alreadyPaidIdx = headers.findIndex((h: any) =>
              /^ALREADY.*PAID$/i.test(String(h ?? "").trim())
            );

            if (empCodeIdx === -1 || reimIdx === -1) continue;

            for (let i = rowIdx + 1; i < staffData.length; i++) {
              const dataRow = staffData[i];
              if (!dataRow || dataRow.length === 0) continue;

              if (
                dataRow.some((cell: any) =>
                  /EMP.*CODE|EMPCODE|^SR.*NO/i.test(String(cell ?? ""))
                )
              ) {
                break;
              }

              const empCode = Number(dataRow[empCodeIdx]);
              if (!empCode || isNaN(empCode)) continue;

              const empName = String(dataRow[empNameIdx] || "").trim();
              const dept = String(dataRow[deptIdx] || "").trim();
              const reimHR = Number(dataRow[reimIdx]) || 0;
              const unpaidHR =
                unpaidIdx !== -1 ? Number(dataRow[unpaidIdx]) || 0 : 0;
              const alreadyPaidHR =
                alreadyPaidIdx !== -1
                  ? Number(dataRow[alreadyPaidIdx]) || 0
                  : 0;

              if (!hrReimDataMap.has(empCode)) {
                hrReimDataMap.set(empCode, {
                  reimHR: reimHR,
                  unpaidHR: unpaidHR,
                  alreadyPaidHR: alreadyPaidHR,
                  name: empName,
                  dept: dept || "Staff",
                });
              } else {
                const existing = hrReimDataMap.get(empCode)!;
                existing.reimHR += reimHR;
                existing.unpaidHR += unpaidHR;
                existing.alreadyPaidHR += alreadyPaidHR;
              }
            }
          }
        }
      }

      console.log(`\n✅ Bonus data loaded: ${hrReimDataMap.size} employees`);

      // ========== LOAD STAFF FILE ==========
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);

      const staffEmployees: Map<
        number,
        {
          name: string;
          dept: string;
          dateOfJoining: any;
          months: Map<string, number>;
        }
      > = new Map();

      for (let sheetName of staffWorkbook.SheetNames) {
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";

        if (EXCLUDED_MONTHS.includes(monthKey)) {
          continue;
        }

        const sheet = staffWorkbook.Sheets[sheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let headerIdx = -1;
        for (let i = 0; i < Math.min(15, data.length); i++) {
          if (
            data[i] &&
            data[i].some((v: any) => {
              const t = norm(v);
              return t === "EMPID" || t === "EMPCODE";
            })
          ) {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx === -1) continue;

        const headers = data[headerIdx];
        const empIdIdx = headers.findIndex((h: any) =>
          ["EMPID", "EMPCODE"].includes(norm(h))
        );
        const empNameIdx = headers.findIndex((h: any) =>
          /EMPLOYEE\s*NAME/i.test(String(h ?? ""))
        );
        const salary1Idx = headers.findIndex(
          (h: any) =>
            /SALARY-?1/i.test(String(h ?? "")) || norm(h) === "SALARY1"
        );
        const dojIdx = headers.findIndex((h: any) =>
          /DATE\s*OF\s*JOINING|DOJ|JOINING\s*DATE/i.test(String(h ?? ""))
        );

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) {
          continue;
        }

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim();
          const salary1 = Number(row[salary1Idx]) || 0;
          const doj = dojIdx !== -1 ? row[dojIdx] : null;

          if (!empId || isNaN(empId) || !empName) continue;

          if (!staffEmployees.has(empId)) {
            staffEmployees.set(empId, {
              name: empName,
              dept: "Staff",
              dateOfJoining: doj,
              months: new Map(),
            });
          } else {
            const existingEmp = staffEmployees.get(empId)!;
            if (!existingEmp.dateOfJoining && doj) {
              existingEmp.dateOfJoining = doj;
            }
          }

          const emp = staffEmployees.get(empId)!;
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
        }
      }

      console.log(
        `\n✅ Staff data extracted: ${staffEmployees.size} employees`
      );

      // ========== CALCULATE GROSS SAL WITH CONDITIONAL OCTOBER ==========
      const softwareEmployeesData: Map<
        number,
        {
          name: string;
          dept: string;
          dateOfJoining: any;
          grossSal: number;
          excludedOctober: boolean;
        }
      > = new Map();

      for (const [empId, rec] of staffEmployees) {
        let baseSum = 0;
        for (const v of rec.months.values()) {
          baseSum += Number(v) || 0;
        }

        let estOct = 0;
        const hasSep2025 = rec.months.has("2025-09");

        if (hasSep2025 && !zeroOctoberEmployees.has(empId)) {
          const values: number[] = [];
          for (const mk of AVG_WINDOW) {
            const v = rec.months.get(mk);
            if (v != null && !isNaN(Number(v))) {
              values.push(Number(v));
            }
          }
          estOct = values.length
            ? values.reduce((a, b) => a + b, 0) / values.length
            : 0;
        } else if (zeroOctoberEmployees.has(empId)) {
          estOct = 0;
          console.log(
            `🔴 Emp ${empId}: October excluded from GROSS & Register`
          );
        }

        const total = baseSum + estOct;

        softwareEmployeesData.set(empId, {
          name: rec.name,
          dept: rec.dept,
          dateOfJoining: rec.dateOfJoining,
          grossSal: total,
          excludedOctober: zeroOctoberEmployees.has(empId),
        });
      }

      console.log(
        `\n✅ GROSS SAL calculated: ${softwareEmployeesData.size} employees`
      );

      // ========== CALCULATE REIMBURSEMENT ==========
      const comparison: any[] = [];

      for (const [empId, softwareData] of softwareEmployeesData) {
        const hrData = hrReimDataMap.get(empId);

        const monthsFromDOJ = calculateMonthsFromDOJ(
          softwareData.dateOfJoining
        );

        let percentageForActual: number;
        let percentageSource: string;

        if (customPercentageMap.has(empId)) {
          percentageForActual = customPercentageMap.get(empId)!;
          percentageSource = "Custom";
        } else {
          percentageForActual = calculatePercentageByMonths(monthsFromDOJ);
          percentageSource = "Calculated";
        }

        // NEW: Use Adj. Gross for Register if employee is in Per sheet
        let adjustedGrossForRegister: number;
        let registerCalculated: number;

        if (employeesInPerSheet.has(empId)) {
          // Employee in Per sheet: use 60% of gross for register
          adjustedGrossForRegister =
            softwareData.grossSal * SPECIAL_GROSS_MULTIPLIER;
          registerCalculated = adjustedGrossForRegister * percentageForActual / 100;

          console.log(
            `🎯 Emp ${empId}: In Per sheet - Using Adj. Gross for Register | Gross=₹${softwareData.grossSal.toFixed(
              2
            )} → 60%=₹${adjustedGrossForRegister.toFixed(
              2
            )} → Register(8.33%)=₹${registerCalculated.toFixed(2)}`
          );
        } else {
          // Not in Per sheet: use full gross for register
          adjustedGrossForRegister = softwareData.grossSal;
          registerCalculated =
            (softwareData.grossSal * REGISTER_PERCENTAGE) / 100;
        }

        const gross2 = calculateGross2(
          softwareData.grossSal,
          percentageForActual
        );
        const actualCalculated = calculateActual(
          softwareData.grossSal,
          gross2,
          percentageForActual
        );

        // Determine payment status from Due VC
        let paymentStatus = "None";
        const dueVCData = dueVoucherMap.get(empId);

        if (dueVCData) {
          // Employee has Due VC > 0, they are UNPAID
          paymentStatus = "Unpaid";
          console.log(`🔴 Emp ${empId}: UNPAID - Due VC = ₹${dueVCData.dueVC}`);
        } else if (hrData && hrData.alreadyPaidHR > 0) {
          // Check HR data for Already Paid
          paymentStatus = "Already Paid";
        }

        // Set Reim(Software) to 0 if Unpaid or Already Paid
        let reimSoftware: number;
        if (paymentStatus === "Unpaid") {
          reimSoftware = 0;
          console.log(`🔒 Emp ${empId}: UNPAID - Reim(Software) set to 0`);
        } else if (paymentStatus === "Already Paid") {
          reimSoftware = 0;
          console.log(
            `🔒 Emp ${empId}: Already Paid - Reim(Software) set to 0`
          );
        } else {
          reimSoftware = registerCalculated - actualCalculated;
        }

        const reimHR = hrData?.reimHR || 0;

        const difference = reimSoftware - reimHR;
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        comparison.push({
          employeeId: empId,
          employeeName: hrData?.name || softwareData.name,
          department: hrData?.dept || softwareData.dept,
          grossSal: softwareData.grossSal,
          adjustedGross: adjustedGrossForRegister,
          excludedOctober: softwareData.excludedOctober,
          registerPercentage: REGISTER_PERCENTAGE,
          actualPercentage: percentageForActual,
          percentageSource: percentageSource,
          registerCalculated: registerCalculated,
          actualCalculated: actualCalculated,
          reimSoftware: reimSoftware,
          reimHR: reimHR,
          paymentStatus: paymentStatus,
          difference: difference,
          status: status,
          dueVC: dueVCData?.dueVC || 0,
        });
      }

      // === Merge HR-only Unpaid employees not present in Staff sheets (softwareEmployeesData) ===
      for (const [hrEmpId, hrRec] of hrReimDataMap) {
        // Skip if already present from software side
        const alreadyInComparison = comparison.some(
          (r) => r.employeeId === hrEmpId
        );
        if (alreadyInComparison) continue;

        // Only bring in when HR shows Unpaid or Already Paid info; prioritize showing Unpaid
        const isUnpaid = (hrRec.unpaidHR || 0) > 0;
        const isAlreadyPaid = (hrRec.alreadyPaidHR || 0) > 0;

        if (!isUnpaid && !isAlreadyPaid) {
          // No actionable payment status; skip
          continue;
        }

        const paymentStatus = isUnpaid ? "Unpaid" : "Already Paid";

        // When Staff sheets lack this employee, software side has no GROSS context.
        // Populate minimal fields; set Reim(SW)=0 as per rule for Unpaid/Already Paid.
        const registerPercentage = REGISTER_PERCENTAGE;
        const actualPercentage = isUnpaid ? 0 : 0; // unknown, not used since Reim(SW)=0
        const percentageSource = "HR-only";
        const grossSal = 0;
        const adjustedGross = 0;
        const registerCalculated = 0;
        const actualCalculated = 0;
        const reimSoftware = 0; // CRITICAL: zero for Unpaid/Already Paid in software calculation
        const reimHR = hrRec.reimHR || 0;

        // For visibility: department/name from HR where available
        const department = hrRec.dept || "Staff";
        const employeeName = hrRec.name || "";

        const difference = reimSoftware - reimHR;
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        // Add employees from Due VC list who aren't in software data
        for (const [dueEmpId, dueRec] of dueVoucherMap) {
          const alreadyInComparison = comparison.some(
            (r) => r.employeeId === dueEmpId
          );
          if (alreadyInComparison) continue;

          const hrData = hrReimDataMap.get(dueEmpId);

          comparison.push({
            employeeId: dueEmpId,
            employeeName: dueRec.name,
            department: dueRec.dept,
            grossSal: 0,
            adjustedGross: 0,
            excludedOctober: false,
            registerPercentage: 0,
            actualPercentage: 0,
            percentageSource: "Due VC Only",
            registerCalculated: 0,
            actualCalculated: 0,
            reimSoftware: 0,
            reimHR: hrData?.reimHR || 0,
            dueVC: dueRec.dueVC,
            paymentStatus: "Unpaid",
            difference: 0 - (hrData?.reimHR || 0),
            status: "Mismatch",
          });
        }

        comparison.push({
          employeeId: hrEmpId,
          employeeName,
          department,
          grossSal,
          adjustedGross,
          excludedOctober: false,
          registerPercentage,
          actualPercentage,
          percentageSource,
          registerCalculated,
          actualCalculated,
          reimSoftware,
          reimHR,
          paymentStatus,
          difference,
          status,
        });
      }

      comparison.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(comparison);
      setFilteredData(comparison);

      console.log(
        "\n✅ Reimbursement calculation completed with Adj. Gross & Payment Status"
      );
    } catch (err: any) {
      setError(`Error processing files: ${err.message}`);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  useEffect(() => {
    if (staffFile && bonusFile) {
      processFiles();
    }
  }, [staffFile, bonusFile, actualPercentageFile]);

  // 🎯 Sorting logic
  useEffect(() => {
    let sorted = [...comparisonData];

    if (sortConfig.key && sortConfig.direction) {
      sorted.sort((a, b) => {
        const aValue = a[sortConfig.key!];
        const bValue = b[sortConfig.key!];

        let comparison = 0;
        if (typeof aValue === "number" && typeof bValue === "number") {
          comparison = aValue - bValue;
        } else if (typeof aValue === "string" && typeof bValue === "string") {
          comparison = aValue.localeCompare(bValue);
        } else if (typeof aValue === "boolean" && typeof bValue === "boolean") {
          comparison = aValue === bValue ? 0 : aValue ? 1 : -1;
        }

        return sortConfig.direction === "asc" ? comparison : -comparison;
      });
    }

    setFilteredData(sorted);
  }, [comparisonData, sortConfig]);

  // 🎯 Handle column sorting
  const handleSort = (key: string) => {
    let direction: "asc" | "desc" = "asc";

    if (sortConfig.key === key) {
      if (sortConfig.direction === "asc") {
        direction = "desc";
      } else if (sortConfig.direction === "desc") {
        setSortConfig({ key: null, direction: null });
        return;
      }
    }

    setSortConfig({ key, direction });
  };

  // 🎯 Sort icon component
  const SortIcon = ({ columnKey }: { columnKey: string }) => {
    const isActive = sortConfig.key === columnKey;

    return (
      <div className="inline-flex flex-col ml-1">
        <svg
          className={`w-3 h-3 ${
            isActive && sortConfig.direction === "asc"
              ? "text-blue-600"
              : "text-gray-400"
          }`}
          fill="currentColor"
          viewBox="0 0 20 20"
        >
          <path d="M5 10l5-5 5 5H5z" />
        </svg>
        <svg
          className={`w-3 h-3 -mt-1 ${
            isActive && sortConfig.direction === "desc"
              ? "text-blue-600"
              : "text-gray-400"
          }`}
          fill="currentColor"
          viewBox="0 0 20 20"
        >
          <path d="M15 10l-5 5-5-5h10z" />
        </svg>
      </div>
    );
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      maximumFractionDigits: 2,
    }).format(value);
  };

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(
      comparisonData.map((row) => ({
        "Employee ID": row.employeeId,
        "Employee Name": row.employeeName,
        Department: row.department,
        "GROSS SAL": row.grossSal,
        "Adj. Gross": row.adjustedGross,
        "October Excluded": row.excludedOctober ? "Yes" : "No",
        "Register %": row.registerPercentage + "%",
        "Actual %": row.actualPercentage + "%",
        "% Source": row.percentageSource,
        "Register (Calculated)": row.registerCalculated,
        "Actual (Calculated)": row.actualCalculated,
        "Reim(Software)": row.reimSoftware,
        "Reim(HR)": row.reimHR,
        "Payment Status": row.paymentStatus,
        Difference: row.difference,
        Status: row.status,
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Step8 Reimbursement");
    XLSX.writeFile(wb, `Step8-Reimbursement-Calculation.xlsx`);
  };

  // ========== Calculate Grand Totals ==========
const calculateGrandTotals = () => {
  const totalReimSoftware = filteredData.reduce(
    (sum, row) => sum + Number(row.reimSoftware || 0),
    0
  );
  const totalReimHR = filteredData.reduce(
    (sum, row) => sum + Number(row.reimHR || 0),
    0
  );
  const totalDifference = totalReimSoftware - totalReimHR;
  
  // Grand Total Status: Mismatch if |difference| > number of employees
  const employeeCount = filteredData.length;
  const grandTotalStatus = Math.abs(totalDifference) > employeeCount ? "Mismatch" : "Match";

  return {
    totalReimSoftware,
    totalReimHR,
    totalDifference,
    grandTotalStatus,
  };
};


  const grandTotals = calculateGrandTotals();

  const FileCard = ({
    title,
    file,
    description,
    optional = false,
  }: {
    title: string;
    file: File | null;
    description: string;
    optional?: boolean;
  }) => (
    <div
      className={`border-2 rounded-lg p-6 ${
        file
          ? "border-green-300 bg-green-50"
          : optional
          ? "border-gray-300 bg-gray-50"
          : "border-red-300 bg-red-50"
      }`}
    >
      {file ? (
        <div className="space-y-3">
          <div className="bg-white rounded-lg p-4 border border-green-200">
            <div className="flex items-center justify-between mb-2">
              <p className="text-sm font-medium text-gray-800 truncate flex-1 mr-2">
                {file.name}
              </p>
              <span className="text-xs bg-green-100 text-green-700 px-2 py-1 rounded font-medium">
                Cached
              </span>
            </div>
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
            File is ready
          </div>
        </div>
      ) : (
        <div className="bg-white rounded-lg p-4 border border-gray-200">
          <div
            className={`flex items-center gap-2 ${
              optional ? "text-gray-500" : "text-red-600"
            } mb-2`}
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
                d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16.5c-.77.833.192 2.5 1.732 2.5z"
              />
            </svg>
            <span className="font-medium">
              {optional ? "Optional - Not uploaded" : "File not found"}
            </span>
          </div>
          <p className="text-xs text-gray-500">
            {optional ? "Will use calculated percentages" : "Upload in Step 1"}
          </p>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-purple-50 to-pink-100 py-5 px-4">
      <div className="mx-auto max-w-full">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 8 - Reimbursement Calculation
              </h1>
              <p className="text-gray-600 mt-2">
                Formula: Reimbursement = Register - Actual (with Adj. Gross &
                Payment Status)
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("step1")}
                className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Back to Step 1
              </button>
              <button
                onClick={() => router.push("step7")}
                className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Back to Step 7
              </button>
              <button
                onClick={() => router.push("step9")}
                className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Move to Step 9
              </button>
            </div>
          </div>

          {/* Logic Explanation */}
          <div className="mb-8 bg-blue-50 border border-blue-200 rounded-lg overflow-hidden">
            <div
              className="flex justify-between items-center p-4 cursor-pointer hover:bg-blue-100 transition-colors"
              onClick={() => setIsLogicMinimized(!isLogicMinimized)}
            >
              <h3 className="font-bold text-blue-900 flex items-center gap-2">
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
                    d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
                Reimbursement Calculation Logic
              </h3>
              <button className="text-blue-600 hover:text-blue-800 transition-colors">
                <svg
                  className={`w-5 h-5 transition-transform duration-200 ${
                    isLogicMinimized ? "" : "rotate-180"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M5 15l7-7 7 7"
                  />
                </svg>
              </button>
            </div>

            {!isLogicMinimized && (
              <div className="px-6 pb-6">
                <div className="text-sm text-blue-800 space-y-3">
                  <div>
                    <p className="font-semibold mb-1">Formula:</p>
                    <p className="ml-4 text-lg">
                      <strong>Reimbursement = Register - Actual</strong>
                    </p>
                  </div>

                  <div>
                    <p className="font-semibold mb-1">Calculation Steps:</p>
                    <ul className="list-disc ml-8 space-y-1">
                      <li>
                        <strong>GROSS SAL:</strong> Sum of monthly SALARY1 +
                        October estimate
                      </li>
                      <li>
                        <strong>Adj. Gross:</strong> For employees in Per sheet
                        → GROSS × 60%, others → GROSS
                      </li>
                      <li>
                        <strong>Register:</strong> Adj. Gross × 8.33%
                      </li>
                      <li>
                        <strong>Actual:</strong> Calculated using custom or
                        DOJ-based percentage
                      </li>
                      <li>
                        <strong>Reimbursement:</strong> Register - Actual (set
                        to 0 if Already Paid or Unpaid)
                      </li>
                    </ul>
                  </div>

                  <div className="bg-purple-50 border border-purple-200 rounded p-3 mt-3">
                    <p className="font-semibold text-purple-900 mb-1">
                      🎯 Adj. Gross Rule:
                    </p>
                    <p className="ml-4 text-purple-800">
                      Employees in "Per" sheet use <strong>Gross × 60%</strong>{" "}
                      for Register calculation
                    </p>
                  </div>

                  <div className="bg-red-50 border border-red-200 rounded p-3">
                    <p className="font-semibold text-red-900 mb-1">
                      🔒 Payment Status Rule:
                    </p>
                    <p className="ml-4 text-red-800">
                      If employee is <strong>Already Paid</strong> or{" "}
                      <strong>Unpaid</strong>, Reim(Software) is set to{" "}
                      <strong>0</strong>
                    </p>
                  </div>

                  <div className="bg-violet-50 border border-violet-200 rounded p-3">
                    <p className="font-semibold text-purple-900 mb-1">
                      🔴 October Exclusion Rule:
                    </p>
                    <p className="ml-4 text-green-800">
                      Employees in "Average" sheet have{" "}
                      <strong>October estimate = 0</strong>
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data for GROSS SAL calculation"
            />
            <FileCard
              title="Bonus Calculation Sheet"
              file={bonusFile}
              description="Reim (HR), Unpaid, Already Paid values"
            />
            <FileCard
              title="Actual Percentage Bonus Data"
              file={actualPercentageFile}
              description="Custom percentages (Per) & October exclusion (Average)"
              optional={true}
            />
          </div>

          {[staffFile, bonusFile].filter(Boolean).length < 2 && (
            <div className="mt-8 bg-yellow-50 border border-yellow-200 rounded-lg p-4">
              <div className="flex items-center gap-3">
                <svg
                  className="w-6 h-6 text-yellow-600"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16.5c-.77.833.192 2.5 1.732 2.5z"
                  />
                </svg>
                <div>
                  <h3 className="font-medium text-yellow-800">
                    Some files are missing
                  </h3>
                  <p className="text-sm text-yellow-600 mt-1">
                    Please upload all required files in Step 1
                  </p>
                </div>
              </div>
            </div>
          )}

          {isProcessing && (
            <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-4">
              <div className="flex items-center gap-3">
                <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
                <p className="text-blue-800">
                  Calculating Reimbursement with Adj. Gross & Payment Status...
                </p>
              </div>
            </div>
          )}

          {error && (
            <div className="mt-8 bg-red-50 border border-red-200 rounded-lg p-4">
              <div className="flex items-center gap-3">
                <svg
                  className="w-6 h-6 text-red-600"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
                <p className="text-red-800">{error}</p>
              </div>
            </div>
          )}

          {comparisonData.length > 0 && (
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-xl font-bold text-gray-800">
                  Reimbursement Comparison Results
                </h2>
                <button
                  onClick={exportToExcel}
                  className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition flex items-center gap-2"
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
                      d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                    />
                  </svg>
                  Export to Excel
                </button>
              </div>

              <div className="border border-gray-300 rounded-lg overflow-hidden">
                <div className="overflow-x-auto">
                  <div className="max-h-[600px] overflow-y-auto">
                    <table className="w-full border-collapse">
                      <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                          <th
                            className="border border-gray-300 px-4 py-2 text-left cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("employeeId")}
                          >
                            <div className="flex items-center">
                              Employee ID
                              <SortIcon columnKey="employeeId" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-left cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("employeeName")}
                          >
                            <div className="flex items-center">
                              Employee Name
                              <SortIcon columnKey="employeeName" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-left cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("department")}
                          >
                            <div className="flex items-center">
                              Department
                              <SortIcon columnKey="department" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-center cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("actualPercentage")}
                          >
                            <div className="flex items-center justify-center">
                              Actual %
                              <SortIcon columnKey="actualPercentage" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-right cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("registerCalculated")}
                          >
                            <div className="flex items-center justify-end">
                              Register (8.33%)
                              <SortIcon columnKey="registerCalculated" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-right cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("actualCalculated")}
                          >
                            <div className="flex items-center justify-end">
                              Actual
                              <SortIcon columnKey="actualCalculated" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-right cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("reimSoftware")}
                          >
                            <div className="flex items-center justify-end">
                              Reim(Software)
                              <SortIcon columnKey="reimSoftware" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-right cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("reimHR")}
                          >
                            <div className="flex items-center justify-end">
                              Reim(HR)
                              <SortIcon columnKey="reimHR" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-right cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("difference")}
                          >
                            <div className="flex items-center justify-end">
                              Difference
                              <SortIcon columnKey="difference" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-4 py-2 text-center cursor-pointer hover:bg-gray-200 select-none bg-gray-100"
                            onClick={() => handleSort("status")}
                          >
                            <div className="flex items-center justify-center">
                              Status
                              <SortIcon columnKey="status" />
                            </div>
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredData.map((row, idx) => (
                          <tr
                            key={idx}
                            className={
                              idx % 2 === 0 ? "bg-white" : "bg-gray-50"
                            }
                          >
                            <td className="border border-gray-300 px-4 py-2">
                              {row.employeeId}
                            </td>
                            <td className="border border-gray-300 px-4 py-2">
                              <div className="flex items-center gap-2">
                                <span>{row.employeeName}</span>
                                {row.excludedOctober && (
                                  <span className="text-xs bg-orange-100 text-orange-700 px-2 py-0.5 rounded">
                                    Oct=0
                                  </span>
                                )}
                                {row.paymentStatus === "Already Paid" && (
                                  <span className="text-xs bg-teal-100 text-teal-700 px-2 py-0.5 rounded font-medium">
                                    Already Paid
                                  </span>
                                )}
                                {row.paymentStatus === "Unpaid" && (
                                  <span className="text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded font-medium">
                                    Unpaid
                                  </span>
                                )}
                              </div>
                            </td>
                            <td className="border border-gray-300 px-4 py-2">
                              <span className="px-2 py-1 bg-indigo-100 text-indigo-800 rounded text-xs font-medium">
                                {row.department}
                              </span>
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-center">
                              <div className="flex flex-col items-center gap-1">
                                <span
                                  className={`px-2 py-1 rounded text-xs font-medium ${
                                    row.actualPercentage === 10.0
                                      ? "bg-red-100 text-red-800"
                                      : row.actualPercentage === 12.0
                                      ? "bg-yellow-100 text-yellow-800"
                                      : "bg-green-100 text-green-800"
                                  }`}
                                >
                                  {row.actualPercentage}%
                                </span>
                                {row.percentageSource === "Custom" && (
                                  <span className="text-xs bg-purple-100 text-purple-700 px-1.5 py-0.5 rounded">
                                    Custom
                                  </span>
                                )}
                              </div>
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right text-blue-600 font-medium">
                              {formatCurrency(row.registerCalculated)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right text-purple-600 font-medium">
                              {formatCurrency(row.actualCalculated)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right font-bold text-green-600">
                              {formatCurrency(row.reimSoftware)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right font-bold text-orange-600">
                              {formatCurrency(row.reimHR)}
                            </td>
                            <td
                              className={`border border-gray-300 px-4 py-2 text-right font-medium ${
                                Math.abs(row.difference) <= TOLERANCE
                                  ? "text-green-600"
                                  : "text-red-600"
                              }`}
                            >
                              {formatCurrency(row.difference)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-center">
                              <span
                                className={`px-3 py-1 rounded-full text-sm font-medium ${
                                  row.status === "Match"
                                    ? "bg-green-100 text-green-800"
                                    : "bg-red-100 text-red-800"
                                }`}
                              >
                                {row.status}
                              </span>
                            </td>
                          </tr>
                        ))}

                        {/* ========== GRAND TOTAL ROW ========== */}
                        <tr className="bg-purple-100 font-bold sticky bottom-0">
                          <td
                            colSpan={6}
                            className="border border-gray-300 px-4 py-3 text-right"
                          >
                            <span className="text-lg">GRAND TOTAL</span>
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-purple-900">
                            {formatCurrency(grandTotals.totalReimSoftware)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-purple-900">
                            {formatCurrency(grandTotals.totalReimHR)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-green-700">
                            {formatCurrency(grandTotals.totalDifference)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-center">
  <span
    className={`px-3 py-1 rounded-full text-sm font-medium ${
      grandTotals.grandTotalStatus === "Match"
        ? "bg-green-200 text-green-900"
        : "bg-red-200 text-red-900"
    }`}
  >
    {grandTotals.grandTotalStatus}
  </span>
</td>

                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>

              <div className="mt-4 flex justify-between items-center text-sm text-gray-600">
                <div>Total Employees: {filteredData.length}</div>
                <div>
                  Matches:{" "}
                  {filteredData.filter((r) => r.status === "Match").length} |
                  Mismatches:{" "}
                  {filteredData.filter((r) => r.status === "Mismatch").length} |
                  Already Paid:{" "}
                  {
                    filteredData.filter(
                      (r) => r.paymentStatus === "Already Paid"
                    ).length
                  }{" "}
                  | Unpaid:{" "}
                  {
                    filteredData.filter((r) => r.paymentStatus === "Unpaid")
                      .length
                  }
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
