"use client";

import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

type SortDirection = "asc" | "desc" | null;
type SortableColumn = 
  | "employeeId"
  | "employeeName"
  | "department"
  | "monthsOfService"
  | "isEligible"
  | "percentage"
  | "grossSalarySoftware"
  | "adjustedGross"
  | "registerSoftware"
  | "unpaidSoftware"
  | "alreadyPaid"
  | "loanDeduction"
  | "finalRTGSSoftware"
  | "finalRTGSHR"
  | "difference"
  | "status";

export default function Step9Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [departmentFilter, setDepartmentFilter] = useState<string>("All");
  const [eligibilityFilter, setEligibilityFilter] = useState<string>("All");

  // 🎯 Sorting state
  const [sortColumn, setSortColumn] = useState<SortableColumn | null>(null);
  const [sortDirection, setSortDirection] = useState<SortDirection>(null);

  // === Step 9 Audit Helpers ===
  const TOLERANCE_STEP9 = 12;

  async function postAuditMessagesStep9(items: any[], batchId?: string) {
    const bid =
      batchId ||
      (typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2));
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId: bid, step: 9, items }),
    });
    return bid;
  }

  function buildStep9MismatchMessages(rows: any[]) {
    const items: any[] = [];
    for (const r of rows) {
      if (r?.status === "Mismatch") {
        items.push({
          level: "error",
          tag: "mismatch",
          text: `[step9] ${r.employeeId} ${r.employeeName} diff=${Number(
            r.difference ?? 0
          ).toFixed(2)}`,
          scope:
            r.department === "Staff"
              ? "staff"
              : r.department === "Worker"
              ? "worker"
              : "global",
          source: "step9",
          meta: {
            employeeId: r.employeeId,
            name: r.employeeName,
            department: r.department,
            monthsOfService: r.monthsOfService,
            isEligible: r.isEligible,
            percentage: r.percentage,
            grossSalarySoftware: r.grossSalarySoftware,
            adjustedGross: r.adjustedGross,
            registerSoftware: r.registerSoftware,
            unpaidSoftware: r.unpaidSoftware,
            alreadyPaid: r.alreadyPaid,
            loanDeduction: r.loanDeduction,
            finalRTGSSoftware: r.finalRTGSSoftware,
            finalRTGSHR: r.finalRTGSHR,
            hrSheets: r.hrSheets,
            hrSheetCount: r.hrSheets?.length || 0,
            diff: r.difference,
            tolerance: TOLERANCE_STEP9,
          },
        });
      }
    }
    return items;
  }

  function buildStep9SummaryMessage(rows: any[]) {
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

    const eligible = rows.filter((r) => r.isEligible).length;
    const notEligible = rows.filter((r) => !r.isEligible).length;
    const multiSheetCount = rows.filter((r) => r.hrSheets?.length > 1).length;
    const specialPercentageCount = rows.filter(
      (r) => r.percentage === 12.0
    ).length;

    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const staffGrossSalSum = sum(
      staffRows.map((r) => Number(r.grossSalarySoftware || 0))
    );
    const staffRegisterSum = sum(
      staffRows.map((r) => Number(r.registerSoftware || 0))
    );
    const staffUnpaidSum = sum(
      staffRows.map((r) => Number(r.unpaidSoftware || 0))
    );
    const staffAlreadyPaidSum = sum(
      staffRows.map((r) => Number(r.alreadyPaid || 0))
    );
    const staffLoanSum = sum(
      staffRows.map((r) => Number(r.loanDeduction || 0))
    );
    const staffFinalRTGSSWSum = sum(
      staffRows.map((r) => Number(r.finalRTGSSoftware || 0))
    );
    const staffFinalRTGSHRSum = sum(
      staffRows.map((r) => Number(r.finalRTGSHR || 0))
    );

    const workerGrossSalSum = sum(
      workerRows.map((r) => Number(r.grossSalarySoftware || 0))
    );
    const workerRegisterSum = sum(
      workerRows.map((r) => Number(r.registerSoftware || 0))
    );
    const workerUnpaidSum = sum(
      workerRows.map((r) => Number(r.unpaidSoftware || 0))
    );
    const workerAlreadyPaidSum = sum(
      workerRows.map((r) => Number(r.alreadyPaid || 0))
    );
    const workerLoanSum = sum(
      workerRows.map((r) => Number(r.loanDeduction || 0))
    );
    const workerFinalRTGSSWSum = sum(
      workerRows.map((r) => Number(r.finalRTGSSoftware || 0))
    );
    const workerFinalRTGSHRSum = sum(
      workerRows.map((r) => Number(r.finalRTGSHR || 0))
    );

    return {
      level: "info",
      tag: "summary",
      text: `Step9 run: total=${total} match=${matches} mismatch=${mismatches}`,
      scope: "global",
      source: "step9",
      meta: {
        totals: {
          total,
          matches,
          mismatches,
          tolerance: TOLERANCE_STEP9,
          eligible,
          notEligible,
          multiSheetCount,
          specialPercentageCount,
        },
        staff: {
          count: staffRows.length,
          mismatches: staffMismatch,
          grossSalSum: staffGrossSalSum,
          registerSum: staffRegisterSum,
          unpaidSum: staffUnpaidSum,
          alreadyPaidSum: staffAlreadyPaidSum,
          loanSum: staffLoanSum,
          finalRTGSSWSum: staffFinalRTGSSWSum,
          finalRTGSHRSum: staffFinalRTGSHRSum,
        },
        worker: {
          count: workerRows.length,
          mismatches: workerMismatch,
          grossSalSum: workerGrossSalSum,
          registerSum: workerRegisterSum,
          unpaidSum: workerUnpaidSum,
          alreadyPaidSum: workerAlreadyPaidSum,
          loanSum: workerLoanSum,
          finalRTGSSWSum: workerFinalRTGSSWSum,
          finalRTGSHRSum: workerFinalRTGSHRSum,
        },
      },
    };
  }

  async function handleSaveAuditStep9(rows: any[]) {
    if (!rows || rows.length === 0) return;
    const items = [
      buildStep9SummaryMessage(rows),
      ...buildStep9MismatchMessages(rows),
    ];
    if (items.length === 0) return;
    await postAuditMessagesStep9(items);
  }

  function djb2Hash(str: string) {
    let h = 5381;
    for (let i = 0; i < str.length; i++) h = (h << 5) + h + str.charCodeAt(i);
    return (h >>> 0).toString(36);
  }

  function buildRunKeyStep9(rows: any[]) {
    const sig = rows
      .map(
        (r) =>
          `${r.employeeId}|${r.department}|${
            Number(r.finalRTGSSoftware) || 0
          }|${Number(r.finalRTGSHR) || 0}|${Number(r.difference) || 0}|${
            r.status
          }|${r.hrSheets?.length || 0}`
      )
      .join(";");
    return djb2Hash(sig);
  }

  useEffect(() => {
    if (typeof window === "undefined") return;
    if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

    const runKey = buildRunKeyStep9(comparisonData);
    const markerKey = `audit_step9_${runKey}`;

    if (sessionStorage.getItem(markerKey)) return;

    sessionStorage.setItem(markerKey, "1");
    const deterministicBatchId = `step9-${runKey}`;

    const items = [
      buildStep9SummaryMessage(comparisonData),
      ...buildStep9MismatchMessages(comparisonData),
    ];

    postAuditMessagesStep9(items, deterministicBatchId).catch((err) => {
      console.error("Auto-audit step9 failed", err);
      sessionStorage.removeItem(markerKey);
    });
  }, [comparisonData]);

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
        /bonus.*final.*calculation|bonus.*2024-25/i.test(s.file.name)
    );

  const actualPercentageFile =
    pickFile((s) => s.type === "Actual-Percentage-Bonus-Data") ??
    pickFile((s) => !!s.file && /actual.*percentage/i.test(s.file.name));

  const dueVoucherFile =
    pickFile((s) => s.type === "Due-Voucher-List") ??
    pickFile((s) => !!s.file && /due.*voucher/i.test(s.file.name));

  const loanDeductionFile =
    pickFile((s) => s.type === "Loan-Deduction") ??
    pickFile((s) => !!s.file && /loan.*deduction/i.test(s.file.name));

  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  const getCellValue = (cell: any): { hasValue: boolean; value: number } => {
    if (cell == null || cell === "") {
      return { hasValue: false, value: 0 };
    }

    if (typeof cell === "number") {
      return { hasValue: true, value: cell };
    }

    if (typeof cell === "string") {
      const trimmed = cell.trim();
      if (!trimmed || trimmed === "-") {
        return { hasValue: false, value: 0 };
      }
      const parsed = Number(trimmed.replace(/,/g, ""));
      if (!isNaN(parsed)) {
        return { hasValue: true, value: parsed };
      }
      return { hasValue: false, value: 0 };
    }

    return { hasValue: false, value: 0 };
  };

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
  const EXCLUDED_DEPARTMENTS = ["C", "CASH", "A"];

  const EXCLUDE_OCTOBER_EMPLOYEES = new Set<number>([
    937, 1039, 1065, 1105, 59, 161,
  ]);

  const DEFAULT_PERCENTAGE = 8.33;
  const SPECIAL_PERCENTAGE = 12.0;
  const SPECIAL_GROSS_MULTIPLIER = 0.6; // 60% of gross for 12% employees
  const TOLERANCE = 12;

  const referenceDate = new Date(Date.UTC(2025, 9, 30));

  function parseDOJ(raw: any): Date | null {
    if (raw == null || raw === "") return null;

    if (typeof raw === "number") {
      const excelEpoch = Date.UTC(1899, 11, 30);
      return new Date(excelEpoch + raw * 86400000);
    }

    if (typeof raw === "string") {
      let s = raw.trim();

      if (/\d{4}-\d{2}-\d{2}\s+\d/.test(s)) {
        s = s.split(/\s+/)[0];
      }

      s = s.replace(/[.\/]/g, "-");

      const m = /^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/.exec(s);
      if (m) {
        let [_, d, mo, y] = m;
        let year = Number(
          y.length === 2 ? (Number(y) <= 29 ? "20" + y : "19" + y) : y
        );
        let month = Number(mo) - 1;
        let day = Number(d);
        const dt = new Date(Date.UTC(year, month, day));
        return isNaN(dt.getTime()) ? null : dt;
      }

      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        const dt = new Date(s + "T00:00:00Z");
        return isNaN(dt.getTime()) ? null : dt;
      }
    }

    return null;
  }

  function monthsBetween(start: Date, end: Date): number {
    const sy = start.getUTCFullYear();
    const sm = start.getUTCMonth();
    const sd = start.getUTCDate();
    const ey = end.getUTCFullYear();
    const em = end.getUTCMonth();
    const ed = end.getUTCDate();

    let months = (ey - sy) * 12 + (em - sm);

    if (ed < sd) {
      months -= 1;
    }

    return Math.max(0, months);
  }

  const calculateMonthsOfService = (dateOfJoining: any): number => {
    const doj = parseDOJ(dateOfJoining);
    if (!doj) return 0;
    return monthsBetween(doj, referenceDate);
  };

  const processFiles = async () => {
    if (
      !staffFile ||
      !workerFile ||
      !bonusFile ||
      !actualPercentageFile ||
      !dueVoucherFile ||
      !loanDeductionFile
    ) {
      setError("All six files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log(
        "📊 STEP 9: Final RTGS Comparison (with adj.gross + already paid logic)"
      );
      console.log("=".repeat(60));
      console.log("⚡ NEW FORMULA: Register - Unpaid - Loan - Already Paid = Final RTGS");
      console.log("⚡ 12% employees: (Step3-Gross × 60%) × 12% = Register");
      console.log("📌 8.33% employees: Step3-Gross × 8.33% = Register");

      // ========== LOAD LOAN DEDUCTION DATA ==========
      const loanBuffer = await loanDeductionFile.arrayBuffer();
      const loanWorkbook = XLSX.read(loanBuffer);
      const loanSheet = loanWorkbook.Sheets[loanWorkbook.SheetNames[0]];
      const loanData: any[][] = XLSX.utils.sheet_to_json(loanSheet, {
        header: 1,
      });

      const loanMap: Map<number, number> = new Map();
      const loanHeaderRow = 1;
      const empIdIdx = 1;
      const loanIdx = 5;

      for (let i = loanHeaderRow + 1; i < loanData.length; i++) {
        const row = loanData[i];
        if (!row || row.length === 0) continue;

        const empId = Number(row[empIdIdx]);
        const loanAmount = Number(row[loanIdx]) || 0;

        if (empId && !isNaN(empId) && loanAmount > 0) {
          loanMap.set(empId, loanAmount);
          console.log(`💰 Loan: Emp ${empId} = ₹${loanAmount}`);
        }
      }

      console.log(`✅ Loan deduction data loaded: ${loanMap.size} employees`);

      // ========== LOAD ACTUAL PERCENTAGE DATA ==========
      const actualPercentageBuffer = await actualPercentageFile.arrayBuffer();
      const actualPercentageWorkbook = XLSX.read(actualPercentageBuffer);
      const actualPercentageSheet =
        actualPercentageWorkbook.Sheets[actualPercentageWorkbook.SheetNames[0]];
      const actualPercentageData: any[][] = XLSX.utils.sheet_to_json(
        actualPercentageSheet,
        { header: 1 }
      );

      const specialPercentageEmployees = new Set<number>();
      let headerRow = -1;

      for (let i = 0; i < Math.min(10, actualPercentageData.length); i++) {
        if (
          actualPercentageData[i] &&
          actualPercentageData[i].some((v: any) => {
            const t = norm(v);
            return t === "EMPCODE" || t === "EMPLOYEECODE";
          })
        ) {
          headerRow = i;
          break;
        }
      }

      if (headerRow !== -1) {
        const headers = actualPercentageData[headerRow];
        const empCodeIdx = headers.findIndex((h: any) =>
          ["EMPCODE", "EMPLOYEECODE"].includes(norm(h))
        );
        const percentageIdx = headers.findIndex((h: any) =>
          /BONUS.*PERCENTAGE|PERCENTAGE/i.test(String(h ?? ""))
        );

        if (empCodeIdx !== -1 && percentageIdx !== -1) {
          for (let i = headerRow + 1; i < actualPercentageData.length; i++) {
            const row = actualPercentageData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const percentage = Number(row[percentageIdx]);

            if (
              empCode &&
              !isNaN(empCode) &&
              percentage === SPECIAL_PERCENTAGE
            ) {
              specialPercentageEmployees.add(empCode);
            }
          }
        }
      }

      console.log(
        `✅ Special percentage employees: ${specialPercentageEmployees.size}`
      );

      // ========== LOAD DUE VOUCHER DATA ==========
      const dueVoucherBuffer = await dueVoucherFile.arrayBuffer();
      const dueVoucherWorkbook = XLSX.read(dueVoucherBuffer);
      const dueVoucherSheet =
        dueVoucherWorkbook.Sheets[dueVoucherWorkbook.SheetNames[0]];
      const dueVoucherData: any[][] = XLSX.utils.sheet_to_json(
        dueVoucherSheet,
        { header: 1 }
      );

      const dueVCMap: Map<number, number> = new Map();
      let dueVCHeaderRow = -1;

      for (let i = 0; i < Math.min(10, dueVoucherData.length); i++) {
        if (
          dueVoucherData[i] &&
          dueVoucherData[i].some((v: any) => {
            const t = norm(v);
            return t === "EMPCODE" || t === "EMPLOYEECODE";
          })
        ) {
          dueVCHeaderRow = i;
          break;
        }
      }

      if (dueVCHeaderRow !== -1) {
        const headers = dueVoucherData[dueVCHeaderRow];
        const empCodeIdx = headers.findIndex((h: any) =>
          ["EMPCODE", "EMPLOYEECODE"].includes(norm(h))
        );
        const dueVCIdx = headers.findIndex((h: any) =>
          /DUE.*VC|DUEVC/i.test(norm(h))
        );

        if (empCodeIdx !== -1 && dueVCIdx !== -1) {
          for (let i = dueVCHeaderRow + 1; i < dueVoucherData.length; i++) {
            const row = dueVoucherData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const dueVC = Number(row[dueVCIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              dueVCMap.set(empCode, dueVC);
            }
          }
        }
      }

      console.log(`✅ Due VC data loaded: ${dueVCMap.size} employees`);

      // ========== 🆕 LOAD "ALREADY PAID" (PAID) DATA FROM BONUS FILE ==========
      const alreadyPaidMap: Map<number, { paid: number; sheets: string[] }> = new Map();

      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      console.log("\n💳 LOADING 'ALREADY PAID' (PAID) DATA FROM BONUS SHEETS");
      console.log("=".repeat(60));

      for (const sheetName of bonusWorkbook.SheetNames) {
        if (sheetName === "Loan Ded.") {
          console.log(`⏭️ Skipping sheet: ${sheetName}`);
          continue;
        }

        console.log(`📄 Processing Bonus sheet for Paid: ${sheetName}`);
        const sheet = bonusWorkbook.Sheets[sheetName];
        const sheetData: any[][] = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
        });

        let sheetHeaderRow = -1;
        for (let i = 0; i < Math.min(10, sheetData.length); i++) {
          if (
            sheetData[i] &&
            sheetData[i].some((v: any) => {
              const t = norm(v);
              return t === "EMPCODE" || t === "EMPLOYEECODE";
            })
          ) {
            sheetHeaderRow = i;
            break;
          }
        }

        if (sheetHeaderRow === -1) {
          console.log(`⚠️ No header found in ${sheetName}`);
          continue;
        }

        const headers = sheetData[sheetHeaderRow];
        const empCodeIdx = headers.findIndex((h: any) =>
          ["EMPCODE", "EMPLOYEECODE", "EMP CODE"].includes(
            String(h ?? "")
              .trim()
              .toUpperCase()
              .replace(/\s+/g, "")
          )
        );

        // 🎯 Find "PAID" column (similar to how Step 6 does it)
        const paidIdx = headers.findIndex((h: any) => {
          const headerStr = String(h ?? "")
            .trim()
            .toUpperCase();
          return (
            headerStr === "PAID" ||
            headerStr === "ALREADY PAID" ||
            headerStr === "ALREADYPAID" ||
            /^PAID$/i.test(headerStr)
          );
        });

        if (empCodeIdx === -1 || paidIdx === -1) {
          console.log(
            `⚠️ Required columns not found in ${sheetName} (Emp: ${empCodeIdx}, Paid: ${paidIdx})`
          );
          continue;
        }

        console.log(
          `  ✓ Found columns - EmpCode at ${empCodeIdx}, Paid at ${paidIdx}`
        );

        let recordsInSheet = 0;
        for (let i = sheetHeaderRow + 1; i < sheetData.length; i++) {
          const row = sheetData[i];
          if (!row || row.length === 0) continue;

          const empCodeRaw = row[empCodeIdx];
          const paidRaw = row[paidIdx];

          if (
            empCodeRaw == null ||
            empCodeRaw === "" ||
            paidRaw == null ||
            paidRaw === ""
          )
            continue;

          const empCode = Number(empCodeRaw);
          const paid = Number(paidRaw);

          if (isNaN(empCode) || isNaN(paid)) continue;

          recordsInSheet++;

          if (!alreadyPaidMap.has(empCode)) {
            alreadyPaidMap.set(empCode, {
              paid: paid,
              sheets: [sheetName],
            });
          } else {
            const existing = alreadyPaidMap.get(empCode)!;
            existing.paid += paid;
            existing.sheets.push(sheetName);
            console.log(
              `  🔄 Emp ${empCode}: Adding Paid ₹${paid.toFixed(
                2
              )} from ${sheetName} (Total: ₹${existing.paid.toFixed(2)})`
            );
          }
        }

        console.log(
          `  ✅ Processed ${recordsInSheet} 'Paid' records from ${sheetName}`
        );
      }

      console.log(
        `✅ Already Paid data loaded: ${alreadyPaidMap.size} employees`
      );

      // ========== LOAD BONUS FILE FOR FINAL RTGS (HR) ==========
      const hrFinalRTGSData: Map<
        number,
        { finalRTGS: number; sheets: string[] }
      > = new Map();

      console.log("\n📊 LOADING 'FINAL RTGS' DATA FROM BONUS SHEETS");
      console.log("=".repeat(60));

      for (const sheetName of bonusWorkbook.SheetNames) {
        if (sheetName === "Loan Ded.") {
          console.log(`⏭️ Skipping sheet: ${sheetName}`);
          continue;
        }

        console.log(`📄 Processing Bonus sheet for Final RTGS: ${sheetName}`);
        const sheet = bonusWorkbook.Sheets[sheetName];
        const sheetData: any[][] = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
        });

        let sheetHeaderRow = -1;
        for (let i = 0; i < Math.min(10, sheetData.length); i++) {
          if (
            sheetData[i] &&
            sheetData[i].some((v: any) => {
              const t = norm(v);
              return t === "EMPCODE" || t === "EMPLOYEECODE";
            })
          ) {
            sheetHeaderRow = i;
            break;
          }
        }

        if (sheetHeaderRow === -1) {
          console.log(`⚠️ No header found in ${sheetName}`);
          continue;
        }

        const headers = sheetData[sheetHeaderRow];
        const empCodeIdx = headers.findIndex((h: any) =>
          ["EMPCODE", "EMPLOYEECODE", "EMP CODE"].includes(
            String(h ?? "")
              .trim()
              .toUpperCase()
              .replace(/\s+/g, "")
          )
        );

        const finalRTGSIdx = headers.findIndex((h: any) => {
          const headerStr = String(h ?? "")
            .trim()
            .toUpperCase();
          return (
            headerStr === "FINAL RTGS" ||
            headerStr === "FINALRTGS" ||
            headerStr === "FINAL RTGS.1" ||
            /FINAL.*RTGS/i.test(headerStr)
          );
        });

        if (empCodeIdx === -1 || finalRTGSIdx === -1) {
          console.log(
            `⚠️ Required columns not found in ${sheetName} (Emp: ${empCodeIdx}, RTGS: ${finalRTGSIdx})`
          );
          continue;
        }

        console.log(
          `  ✓ Found columns - EmpCode at ${empCodeIdx}, Final RTGS at ${finalRTGSIdx}`
        );

        let recordsInSheet = 0;
        for (let i = sheetHeaderRow + 1; i < sheetData.length; i++) {
          const row = sheetData[i];
          if (!row || row.length === 0) continue;

          const empCodeRaw = row[empCodeIdx];
          const finalRTGSRaw = row[finalRTGSIdx];

          if (
            empCodeRaw == null ||
            empCodeRaw === "" ||
            finalRTGSRaw == null ||
            finalRTGSRaw === ""
          )
            continue;

          const empCode = Number(empCodeRaw);
          const finalRTGS = Number(finalRTGSRaw);

          if (isNaN(empCode) || isNaN(finalRTGS)) continue;

          recordsInSheet++;

          if (!hrFinalRTGSData.has(empCode)) {
            hrFinalRTGSData.set(empCode, {
              finalRTGS: finalRTGS,
              sheets: [sheetName],
            });
          } else {
            const existing = hrFinalRTGSData.get(empCode)!;
            existing.finalRTGS += finalRTGS;
            existing.sheets.push(sheetName);
            console.log(
              `  🔄 Emp ${empCode}: Adding ₹${finalRTGS.toFixed(
                2
              )} from ${sheetName} (Total: ₹${existing.finalRTGS.toFixed(2)})`
            );
          }
        }

        console.log(
          `  ✅ Processed ${recordsInSheet} records from ${sheetName}`
        );
      }

      const multiSheetEmployees = Array.from(hrFinalRTGSData.entries()).filter(
        ([_, data]) => data.sheets.length > 1
      );
      console.log(
        `\n📊 Employees appearing in multiple sheets: ${multiSheetEmployees.length}`
      );
      multiSheetEmployees.forEach(([empId, data]) => {
        console.log(
          `  Emp ${empId}: ₹${data.finalRTGS.toFixed(
            2
          )} across [${data.sheets.join(", ")}]`
        );
      });

      console.log(
        `✅ HR Final RTGS data loaded: ${hrFinalRTGSData.size} employees`
      );

      // ========== COMPUTE GROSS SALARY ==========
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);

      const staffEmployees: Map<
        number,
        {
          name: string;
          dept: string;
          months: Map<string, { hasValue: boolean; value: number }>;
          dateOfJoining: any;
        }
      > = new Map();

      for (let sheetName of staffWorkbook.SheetNames) {
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";

        if (EXCLUDED_MONTHS.includes(monthKey)) {
          console.log(`🚫 SKIP Staff: ${sheetName} (${monthKey}) - EXCLUDED`);
          continue;
        }

        if (!AVG_WINDOW.includes(monthKey)) {
          console.log(
            `⏭️ SKIP Staff: ${sheetName} (${monthKey}) - NOT IN WINDOW`
          );
          continue;
        }

        console.log(`✅ Processing Staff: ${sheetName} -> ${monthKey}`);

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
            /^\s*SALARY\s*-?\s*1\s*$/i.test(String(h ?? "")) ||
            norm(h) === "SALARY1"
        );

        let dojIdx = headers.findIndex((h: any) => {
          const headerStr = String(h ?? "").trim();
          return /DATE.*OF.*JOINING|DOJ|JOINING.*DATE|DATE.*JOINING|D\.O\.J/i.test(
            headerStr
          );
        });

        if (dojIdx === -1) {
          for (
            let i = Math.max(0, headers.length - 3);
            i < headers.length;
            i++
          ) {
            const h = String(headers[i] ?? "")
              .trim()
              .toLowerCase();
            if (
              h.includes("date") ||
              h.includes("joining") ||
              h.includes("doj")
            ) {
              dojIdx = i;
              break;
            }
          }
        }

        if (dojIdx === -1 && headers.length > 15) {
          dojIdx = headers.length - 1;
        }

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) continue;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "")
            .trim()
            .toUpperCase();

          const salary1Result = getCellValue(row[salary1Idx]);

          const doj = dojIdx !== -1 && row.length > dojIdx ? row[dojIdx] : null;

          if (!empId || isNaN(empId) || !empName) continue;

          if (!staffEmployees.has(empId)) {
            staffEmployees.set(empId, {
              name: empName,
              dept: "Staff",
              months: new Map(),
              dateOfJoining: doj,
            });
          }

          const emp = staffEmployees.get(empId)!;

          const existing = emp.months.get(monthKey);
          if (existing) {
            if (salary1Result.hasValue) {
              emp.months.set(monthKey, {
                hasValue: true,
                value: existing.value + salary1Result.value,
              });
            }
          } else {
            emp.months.set(monthKey, salary1Result);
          }
        }
      }

      console.log(`✅ Staff employees: ${staffEmployees.size}`);

      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);

      const workerEmployees: Map<
        number,
        {
          name: string;
          dept: string;
          months: Map<string, { hasValue: boolean; value: number }>;
          dateOfJoining: any;
        }
      > = new Map();

      for (let sheetName of workerWorkbook.SheetNames) {
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";

        if (EXCLUDED_MONTHS.includes(monthKey)) {
          console.log(`🚫 SKIP Worker: ${sheetName} (${monthKey}) - EXCLUDED`);
          continue;
        }

        if (!AVG_WINDOW.includes(monthKey)) {
          console.log(
            `⏭️ SKIP Worker: ${sheetName} (${monthKey}) - NOT IN WINDOW`
          );
          continue;
        }

        console.log(`✅ Processing Worker: ${sheetName} -> ${monthKey}`);

        const sheet = workerWorkbook.Sheets[sheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let headerIdx = -1;
        for (let i = 0; i < Math.min(5, data.length); i++) {
          if (data[i] && data[i].some((v: any) => norm(v) === "EMPID")) {
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

        const deptIdx = headers.findIndex((h: any) => {
          const normalized = norm(h);
          return (
            normalized === "DEPT" ||
            normalized === "DEPARTMENT" ||
            normalized === "DEPTT"
          );
        });

        const salary1Idx = 8;

        let dojIdx = headers.findIndex((h: any) => {
          const headerStr = String(h ?? "").trim();
          return /DATE.*OF.*JOINING|DOJ|JOINING.*DATE|DATE.*JOINING|D\.O\.J/i.test(
            headerStr
          );
        });

        if (dojIdx === -1 && headers.length > 15) {
          dojIdx = headers.length - 1;
        }

        if (empIdIdx === -1 || empNameIdx === -1) continue;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "")
            .trim()
            .toUpperCase();

          const salary1Result = getCellValue(row[salary1Idx]);

          const doj = dojIdx !== -1 && row.length > dojIdx ? row[dojIdx] : null;

          if (deptIdx !== -1) {
            const dept = String(row[deptIdx] || "")
              .trim()
              .toUpperCase();
            if (EXCLUDED_DEPARTMENTS.includes(dept)) {
              continue;
            }
          }

          if (!empId || isNaN(empId) || !empName) continue;

          if (!workerEmployees.has(empId)) {
            workerEmployees.set(empId, {
              name: empName,
              dept: "Worker",
              months: new Map(),
              dateOfJoining: doj,
            });
          }

          const emp = workerEmployees.get(empId)!;

          const existing = emp.months.get(monthKey);
          if (existing) {
            if (salary1Result.hasValue) {
              emp.months.set(monthKey, {
                hasValue: true,
                value: existing.value + salary1Result.value,
              });
            }
          } else {
            emp.months.set(monthKey, salary1Result);
          }
        }
      }

      console.log(`✅ Worker employees: ${workerEmployees.size}`);

      // ========== COMPUTE SOFTWARE TOTALS WITH OCTOBER ESTIMATION ==========
      const employeeData: Map<
        number,
        { name: string; dept: string; grossSalary: number; dateOfJoining: any }
      > = new Map();

      const foldMonthly = (
        src: Map<
          number,
          {
            name: string;
            dept: string;
            months: Map<string, { hasValue: boolean; value: number }>;
            dateOfJoining: any;
          }
        >
      ) => {
        for (const [empId, rec] of src) {
          let baseSum = 0;
          const monthsIncluded: { month: string; value: number }[] = [];

          for (const mk of AVG_WINDOW) {
            const cellData = rec.months.get(mk);
            if (cellData && cellData.hasValue) {
              baseSum += cellData.value;
              monthsIncluded.push({ month: mk, value: cellData.value });
            }
          }

          let estOct = 0;
          let total = baseSum;

          const sepData = rec.months.get("2025-09");
          const hasSep2025 = sepData && sepData.hasValue && sepData.value > 0;
          const isExcluded = EXCLUDE_OCTOBER_EMPLOYEES.has(empId);

          if (isExcluded) {
            console.log(
              `🚫 EMP ${empId} (${
                rec.name
              }): IN EXCLUDE LIST - Base only = ₹${baseSum.toFixed(2)}`
            );
          } else if (hasSep2025 && monthsIncluded.length > 0) {
            const values = monthsIncluded.map((m) => m.value);
            estOct = values.reduce((a, b) => a + b, 0) / values.length;
            total = baseSum + estOct;

            console.log(
              `📊 EMP ${empId} (${rec.name}): Avg from ${
                monthsIncluded.length
              } months with values = ₹${estOct.toFixed(
                2
              )}, Total = ₹${total.toFixed(2)}`
            );
          }

          if (!employeeData.has(empId)) {
            employeeData.set(empId, {
              name: rec.name,
              dept: rec.dept,
              grossSalary: total,
              dateOfJoining: rec.dateOfJoining,
            });
          } else {
            employeeData.get(empId)!.grossSalary += total;
          }
        }
      };

      foldMonthly(staffEmployees);
      foldMonthly(workerEmployees);

      console.log(`✅ Employee data loaded: ${employeeData.size} employees`);

      // ========== 🆕 CALCULATE FINAL RTGS WITH NEW FORMULA ==========
      console.log("\n💰 CALCULATING FINAL RTGS WITH NEW FORMULA");
      console.log("=".repeat(60));
      console.log("Formula: Register - Unpaid - Loan - Already Paid = Final RTGS");
      console.log("=".repeat(60));

      const comparison: any[] = [];

      for (const [empId, empData] of employeeData) {
        const percentage = specialPercentageEmployees.has(empId)
          ? SPECIAL_PERCENTAGE
          : DEFAULT_PERCENTAGE;

        // Apply adj.gross logic just like Step-5
        const shouldApplyGrossAdjustment = percentage === SPECIAL_PERCENTAGE;
        
        let adjustedGross: number;
        let registerSoftware: number;
        
        if (shouldApplyGrossAdjustment) {
          // For 12% employees: adj.gross = gross × 60%
          adjustedGross = empData.grossSalary * SPECIAL_GROSS_MULTIPLIER;
          registerSoftware = (adjustedGross * percentage) / 100;
        } else {
          // For 8.33% employees: use full gross
          adjustedGross = empData.grossSalary;
          registerSoftware = (empData.grossSalary * percentage) / 100;
        }

        const monthsOfService = calculateMonthsOfService(empData.dateOfJoining);

        let isEligible = true;
        if (empData.dept === "Worker") {
          isEligible = monthsOfService >= 6;
        }

        let unpaidSoftware = dueVCMap.get(empId) || 0;

        if (!isEligible) {
          unpaidSoftware = registerSoftware;
        }

        const loanDeduction = loanMap.get(empId) || 0;
        
        // 🆕 Get Already Paid amount
        const alreadyPaidData = alreadyPaidMap.get(empId);
        const alreadyPaid = alreadyPaidData?.paid || 0;

        // 🆕 NEW FORMULA: Register - Unpaid - Loan - Already Paid
        const finalRTGSSoftware =
          registerSoftware - unpaidSoftware - loanDeduction - alreadyPaid;

        const hrData = hrFinalRTGSData.get(empId);
        const finalRTGSHR = hrData?.finalRTGS || 0;
        const hrSheets = hrData?.sheets || [];

        const difference = finalRTGSSoftware - finalRTGSHR;
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        comparison.push({
          employeeId: empId,
          employeeName: empData.name,
          department: empData.dept,
          monthsOfService: monthsOfService,
          isEligible: isEligible,
          percentage: percentage,
          grossSalarySoftware: empData.grossSalary,
          adjustedGross: adjustedGross,
          registerSoftware: registerSoftware,
          unpaidSoftware: unpaidSoftware,
          alreadyPaid: alreadyPaid,
          loanDeduction: loanDeduction,
          finalRTGSSoftware: finalRTGSSoftware,
          finalRTGSHR: finalRTGSHR,
          hrSheets: hrSheets,
          difference: difference,
          status: status,
        });

        console.log(
          `Emp ${empId}: Register=₹${registerSoftware.toFixed(
            2
          )} - Unpaid=₹${unpaidSoftware.toFixed(
            2
          )} - Loan=₹${loanDeduction.toFixed(
            2
          )} - AlreadyPaid=₹${alreadyPaid.toFixed(
            2
          )} = Final RTGS (SW)=₹${finalRTGSSoftware.toFixed(
            2
          )}, Final RTGS (HR)=₹${finalRTGSHR.toFixed(2)}`
        );
      }

      comparison.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(comparison);
      setFilteredData(comparison);

      console.log("✅ Final RTGS comparison completed with new formula");
    } catch (err: any) {
      setError(`Error processing files: ${err.message}`);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  useEffect(() => {
    if (
      staffFile &&
      workerFile &&
      bonusFile &&
      actualPercentageFile &&
      dueVoucherFile &&
      loanDeductionFile
    ) {
      processFiles();
    }
    // eslint-disable-next-line
  }, [
    staffFile,
    workerFile,
    bonusFile,
    actualPercentageFile,
    dueVoucherFile,
    loanDeductionFile,
  ]);

  // Apply filters and sorting
  useEffect(() => {
    let filtered = comparisonData;

    if (departmentFilter !== "All") {
      filtered = filtered.filter((row) => row.department === departmentFilter);
    }

    if (eligibilityFilter !== "All") {
      if (eligibilityFilter === "Eligible") {
        filtered = filtered.filter((row) => row.isEligible);
      } else if (eligibilityFilter === "Not Eligible") {
        filtered = filtered.filter((row) => !row.isEligible);
      }
    }

    // Apply sorting
    if (sortColumn && sortDirection) {
      filtered = [...filtered].sort((a, b) => {
        let aVal = a[sortColumn];
        let bVal = b[sortColumn];

        // Handle boolean values
        if (typeof aVal === "boolean") {
          aVal = aVal ? 1 : 0;
          bVal = bVal ? 1 : 0;
        }

        // Handle string values
        if (typeof aVal === "string") {
          aVal = aVal.toLowerCase();
          bVal = bVal.toLowerCase();
        }

        if (sortDirection === "asc") {
          return aVal > bVal ? 1 : aVal < bVal ? -1 : 0;
        } else {
          return aVal < bVal ? 1 : aVal > bVal ? -1 : 0;
        }
      });
    }

    setFilteredData(filtered);
  }, [departmentFilter, eligibilityFilter, comparisonData, sortColumn, sortDirection]);

  // Sort handler
  const handleSort = (column: SortableColumn, direction: "asc" | "desc") => {
    setSortColumn(column);
    setSortDirection(direction);
  };

  // Sort button component
  const SortButtons = ({ column }: { column: SortableColumn }) => {
    return (
      <div className="inline-flex flex-col ml-1">
        <button
          onClick={() => handleSort(column, "asc")}
          className={`leading-none ${
            sortColumn === column && sortDirection === "asc"
              ? "text-blue-600"
              : "text-gray-400 hover:text-gray-600"
          }`}
          title="Sort Ascending"
        >
          <svg className="w-2 h-2" fill="currentColor" viewBox="0 0 10 10">
            <path d="M5 2l3 3H2z" />
          </svg>
        </button>
        <button
          onClick={() => handleSort(column, "desc")}
          className={`leading-none ${
            sortColumn === column && sortDirection === "desc"
              ? "text-blue-600"
              : "text-gray-400 hover:text-gray-600"
          }`}
          title="Sort Descending"
        >
          <svg className="w-2 h-2" fill="currentColor" viewBox="0 0 10 10">
            <path d="M5 8L2 5h6z" />
          </svg>
        </button>
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
    const dataToExport =
      departmentFilter === "All" && eligibilityFilter === "All"
        ? comparisonData
        : filteredData;

    const ws = XLSX.utils.json_to_sheet(
      dataToExport.map((row) => ({
        "Emp ID": row.employeeId,
        Name: row.employeeName,
        Dept: row.department,
        MOS: row.monthsOfService,
        Eligible: row.isEligible ? "YES" : "NO",
        "%": row.percentage,
        Gross: row.grossSalarySoftware,
        "Adj. Gross": row.adjustedGross,
        Register: row.registerSoftware,
        Unpaid: row.unpaidSoftware,
        "Already Paid": row.alreadyPaid,
        Loan: row.loanDeduction,
        "Final RTGS (Software)": row.finalRTGSSoftware,
        "Final RTGS (HR)": row.finalRTGSHR,
        "HR Sheets": row.hrSheets.join(", "),
        Difference: row.difference,
        Status: row.status,
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Final RTGS Comparison");
    XLSX.writeFile(
      wb,
      `Step9-Final-RTGS-Comparison-${departmentFilter}-${eligibilityFilter}.xlsx`
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-orange-50 to-orange-100">
      <div className="container mx-auto px-4 py-8 max-w-[92rem]">
        <div className="bg-white rounded-lg shadow-lg p-6">
          <div className="flex items-center justify-between mb-6">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 9: Final RTGS Comparison
              </h1>
              <p className="text-sm text-gray-600 mt-2">
                Formula: Register - Unpaid - Loan - Already Paid = Final RTGS
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("/step8")}
                className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Back to Step 8
              </button>
              <button
                onClick={() => router.push("/")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                Back to Step 1
              </button>
            </div>
          </div>

          {error && (
            <div className="mb-6 p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg">
              {error}
            </div>
          )}

          {isProcessing && (
            <div className="mb-6 p-4 bg-blue-100 border border-blue-400 text-blue-700 rounded-lg">
              Processing files...
            </div>
          )}

          {comparisonData.length > 0 && (
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <div className="flex items-center gap-4">
                  <h2 className="text-xl font-bold text-gray-800">
                    Final RTGS Comparison Results
                  </h2>
                  <select
                    value={departmentFilter}
                    onChange={(e) => setDepartmentFilter(e.target.value)}
                    className="px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500"
                  >
                    <option value="All">All Departments</option>
                    <option value="Staff">Staff Only</option>
                    <option value="Worker">Worker Only</option>
                  </select>
                  <select
                    value={eligibilityFilter}
                    onChange={(e) => setEligibilityFilter(e.target.value)}
                    className="px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-orange-500"
                  >
                    <option value="All">All Eligibility</option>
                    <option value="Eligible">Eligible Only</option>
                    <option value="Not Eligible">Not Eligible Only</option>
                  </select>
                </div>
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

              <div className="overflow-x-auto">
                <table className="min-w-full bg-white border border-gray-300">
                  <thead className="bg-gray-100">
                    <tr>
                      <th className="px-4 py-2 border">
                        Emp ID <SortButtons column="employeeId" />
                      </th>
                      <th className="px-4 py-2 border">
                        Name <SortButtons column="employeeName" />
                      </th>
                      <th className="px-4 py-2 border">
                        Dept <SortButtons column="department" />
                      </th>
                      <th className="px-4 py-2 border">
                        MOS <SortButtons column="monthsOfService" />
                      </th>
                      <th className="px-4 py-2 border">
                        Eligible <SortButtons column="isEligible" />
                      </th>
                      <th className="px-4 py-2 border">
                        % <SortButtons column="percentage" />
                      </th>
                      <th className="px-4 py-2 border">
                        Gross <SortButtons column="grossSalarySoftware" />
                      </th>
                      <th className="px-4 py-2 border">
                        Adj. Gross <SortButtons column="adjustedGross" />
                      </th>
                      <th className="px-4 py-2 border">
                        Register <SortButtons column="registerSoftware" />
                      </th>
                      <th className="px-4 py-2 border">
                        Unpaid <SortButtons column="unpaidSoftware" />
                      </th>
                      <th className="px-4 py-2 border">
                        Already Paid <SortButtons column="alreadyPaid" />
                      </th>
                      <th className="px-4 py-2 border">
                        Loan <SortButtons column="loanDeduction" />
                      </th>
                      <th className="px-4 py-2 border">
                        Final RTGS (SW) <SortButtons column="finalRTGSSoftware" />
                      </th>
                      <th className="px-4 py-2 border">
                        Final RTGS (HR) <SortButtons column="finalRTGSHR" />
                      </th>
                      <th className="px-4 py-2 border">
                        Difference <SortButtons column="difference" />
                      </th>
                      <th className="px-4 py-2 border">
                        Status <SortButtons column="status" />
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, idx) => (
                      <tr
                        key={idx}
                        className={
                          row.status === "Match"
                            ? "bg-green-50"
                            : "bg-red-50"
                        }
                      >
                        <td className="px-4 py-2 border">{row.employeeId}</td>
                        <td className="px-4 py-2 border">{row.employeeName}</td>
                        <td className="px-4 py-2 border">{row.department}</td>
                        <td className="px-4 py-2 border">{row.monthsOfService}</td>
                        <td className="px-4 py-2 border">
                          {row.isEligible ? "YES" : "NO"}
                        </td>
                        <td className="px-4 py-2 border">{row.percentage}%</td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.grossSalarySoftware)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.adjustedGross)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.registerSoftware)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.unpaidSoftware)}
                        </td>
                        <td className="px-4 py-2 border text-right bg-yellow-50">
                          {formatCurrency(row.alreadyPaid)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.loanDeduction)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.finalRTGSSoftware)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.finalRTGSHR)}
                        </td>
                        <td className="px-4 py-2 border text-right">
                          {formatCurrency(row.difference)}
                        </td>
                        <td className="px-4 py-2 border text-center">
                          <span
                            className={`px-2 py-1 rounded ${
                              row.status === "Match"
                                ? "bg-green-200 text-green-800"
                                : "bg-red-200 text-red-800"
                            }`}
                          >
                            {row.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="mt-4 flex justify-between items-center text-sm text-gray-600">
                <div>
                  Total: {filteredData.length} | Staff:{" "}
                  {filteredData.filter((r) => r.department === "Staff").length}{" "}
                  | Worker:{" "}
                  {filteredData.filter((r) => r.department === "Worker").length}
                </div>
                <div>
                  Matches:{" "}
                  {filteredData.filter((r) => r.status === "Match").length} |
                  Mismatches:{" "}
                  {filteredData.filter((r) => r.status === "Mismatch").length}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
