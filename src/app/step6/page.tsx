"use client";

import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

export default function Step6Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [departmentFilter, setDepartmentFilter] = useState<string>("All");
  const [eligibilityFilter, setEligibilityFilter] = useState<string>("All");

  // === Step 6 Audit Helpers ===
const TOLERANCE_STEP6 = 12; // Step 6 uses ±12 to mark Match vs Mismatch

async function postAuditMessagesStep6(items: any[], batchId?: string) {
  const bid =
    batchId ||
    (typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2));
  await fetch('/api/audit/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ batchId: bid, step: 6, items }),
  });
  return bid;
}

function buildStep6MismatchMessages(rows: any[]) {
  // Expecting rows with: { employeeId, employeeName, department, monthsOfService, isEligible, percentage, grossSalarySoftware, registerSoftware, registerHR, unpaidSoftware, unpaidHR, hrOccurrences, difference, status, validationError }
  const items: any[] = [];
  for (const r of rows) {
    if (r?.status === 'Mismatch' || r?.status === 'Error') {
      items.push({
        level: r.status === 'Error' ? 'error' : 'warning',
        tag: r.status === 'Error' ? 'validation-error' : 'mismatch',
        text: `[step6] ${r.employeeId} ${r.employeeName} ${r.status === 'Error' ? 'Validation Error' : `diff=${Number(r.difference ?? 0).toFixed(2)}`}`,
        scope: r.department === 'Staff' ? 'staff' : r.department === 'Worker' ? 'worker' : 'global',
        source: 'step6',
        meta: {
          employeeId: r.employeeId,
          name: r.employeeName,
          department: r.department,
          monthsOfService: r.monthsOfService,
          isEligible: r.isEligible,
          percentage: r.percentage,
          grossSalarySoftware: r.grossSalarySoftware,
          registerSoftware: r.registerSoftware,
          registerHR: r.registerHR,
          unpaidSoftware: r.unpaidSoftware,
          unpaidHR: r.unpaidHR,
          hrOccurrences: r.hrOccurrences,
          diff: r.difference,
          status: r.status,
          validationError: r.validationError || null,
          tolerance: TOLERANCE_STEP6,
        },
      });
    }
  }
  return items;
}

function buildStep6SummaryMessage(rows: any[]) {
  const total = rows.length || 0;
  const matches = rows.filter((r) => r.status === 'Match').length;
  const mismatches = rows.filter((r) => r.status === 'Mismatch').length;
  const errors = rows.filter((r) => r.status === 'Error').length;

  const staffRows = rows.filter((r) => r.department === 'Staff');
  const workerRows = rows.filter((r) => r.department === 'Worker');

  const staffMismatch = staffRows.filter((r) => r.status === 'Mismatch' || r.status === 'Error').length;
  const workerMismatch = workerRows.filter((r) => r.status === 'Mismatch' || r.status === 'Error').length;

  const eligible = rows.filter((r) => r.isEligible).length;
  const notEligible = rows.filter((r) => !r.isEligible).length;
  const duplicates = rows.filter((r) => r.hrOccurrences > 1).length;

  const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
  const staffRegSWSum = sum(staffRows.map((r) => Number(r.registerSoftware || 0)));
  const staffRegHRSum = sum(staffRows.map((r) => Number(r.registerHR || 0)));
  const staffUnpaidSWSum = sum(staffRows.map((r) => Number(r.unpaidSoftware || 0)));
  const staffUnpaidHRSum = sum(staffRows.map((r) => Number(r.unpaidHR || 0)));

  const workerRegSWSum = sum(workerRows.map((r) => Number(r.registerSoftware || 0)));
  const workerRegHRSum = sum(workerRows.map((r) => Number(r.registerHR || 0)));
  const workerUnpaidSWSum = sum(workerRows.map((r) => Number(r.unpaidSoftware || 0)));
  const workerUnpaidHRSum = sum(workerRows.map((r) => Number(r.unpaidHR || 0)));

  return {
    level: 'info',
    tag: 'summary',
    text: `Step6 run: total=${total} match=${matches} mismatch=${mismatches} error=${errors}`,
    scope: 'global',
    source: 'step6',
    meta: {
      totals: { total, matches, mismatches, errors, eligible, notEligible, duplicates, tolerance: TOLERANCE_STEP6 },
      staff: {
        count: staffRows.length,
        issues: staffMismatch,
        registerSWSum: staffRegSWSum,
        registerHRSum: staffRegHRSum,
        unpaidSWSum: staffUnpaidSWSum,
        unpaidHRSum: staffUnpaidHRSum,
      },
      worker: {
        count: workerRows.length,
        issues: workerMismatch,
        registerSWSum: workerRegSWSum,
        registerHRSum: workerRegHRSum,
        unpaidSWSum: workerUnpaidSWSum,
        unpaidHRSum: workerUnpaidHRSum,
      },
    },
  };
}

async function handleSaveAuditStep6(rows: any[]) {
  if (!rows || rows.length === 0) return;
  const items = [buildStep6SummaryMessage(rows), ...buildStep6MismatchMessages(rows)];
  if (items.length === 0) return;
  await postAuditMessagesStep6(items);
}

// Stable hash for run signature
function djb2Hash(str: string) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h) + str.charCodeAt(i);
  return (h >>> 0).toString(36);
}

function buildRunKeyStep6(rows: any[]) {
  const sig = rows
    .map(r =>
      `${r.employeeId}|${r.department}|${Number(r.unpaidSoftware)||0}|${Number(r.unpaidHR)||0}|${Number(r.difference)||0}|${r.status}|${r.isEligible}`
    )
    .join(';');
  return djb2Hash(sig);
}

useEffect(() => {
  if (typeof window === 'undefined') return; // SSR guard
  if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

  const runKey = buildRunKeyStep6(comparisonData);
  const markerKey = `audit_step6_${runKey}`;

  if (sessionStorage.getItem(markerKey)) return; // prevent duplicate on refresh/StrictMode

  sessionStorage.setItem(markerKey, '1');
  const deterministicBatchId = `step6-${runKey}`;

  const items = [buildStep6SummaryMessage(comparisonData), ...buildStep6MismatchMessages(comparisonData)];

  postAuditMessagesStep6(items, deterministicBatchId).catch(err => {
    console.error('Auto-audit step6 failed', err);
    sessionStorage.removeItem(markerKey); // allow retry on next refresh if failed
  });
}, [comparisonData]);

useEffect(() => {
  if (typeof window === 'undefined') return;
  if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

  const batchId = `step6-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const items = [buildStep6SummaryMessage(comparisonData), ...buildStep6MismatchMessages(comparisonData)];

  postAuditMessagesStep6(items, batchId).catch(err => console.error('Auto-audit step6 failed', err));
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

  // Helper to normalize header text
  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  // Constants from Step 5
  const MONTH_NAME_MAP: Record<string, number> = {
    JAN: 1, JANUARY: 1,
    FEB: 2, FEBRUARY: 2,
    MAR: 3, MARCH: 3,
    APR: 4, APRIL: 4,
    MAY: 5,
    JUN: 6, JUNE: 6,
    JUL: 7, JULY: 7,
    AUG: 8, AUGUST: 8,
    SEP: 9, SEPT: 9, SEPTEMBER: 9,
    OCT: 10, OCTOBER: 10,
    NOV: 11, NOVEMBER: 11,
    DEC: 12, DECEMBER: 12,
  };

  const pad2 = (n: number) => String(n).padStart(2, "0");

  const parseMonthFromSheetName = (sheetName: string): string | null => {
    const s = String(sheetName || "").trim().toUpperCase();
    
    const yyyymm = s.match(/(20\d{2})\D{0,2}(\d{1,2})/);
    if (yyyymm) {
      const y = Number(yyyymm[1]);
      const m = Number(yyyymm[2]);
      if (y >= 2000 && m >= 1 && m <= 12) return `${y}-${pad2(m)}`;
    }
    
    const mon = s.match(/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b/);
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
    "2024-11", "2024-12", "2025-01", "2025-02", "2025-03", "2025-04",
    "2025-05", "2025-06", "2025-07", "2025-08", "2025-09",
  ];

  const EXCLUDED_MONTHS: string[] = ["2025-10", "2024-10"];
  const EXCLUDED_DEPARTMENTS = ["C", "CASH", "A"];

  const EXCLUDE_OCTOBER_EMPLOYEES = new Set<number>([
    937, 1039, 1065, 1105, 59, 161
  ]);

  const DEFAULT_PERCENTAGE = 8.33;
  const SPECIAL_PERCENTAGE = 12.0;
  const TOLERANCE = 12;

  // ✅ CORRECTED: Reference date should be october 31, 2025 (end of bonus period)
  const referenceDate = new Date(Date.UTC(2025, 9, 30)); // 2025-10-31 (UTC)

  // Parse DOJ from various formats
  function parseDOJ(raw: any): Date | null {
    if (raw == null || raw === '') return null;
    
    if (typeof raw === 'number') {
      const excelEpoch = Date.UTC(1899, 11, 30);
      return new Date(excelEpoch + raw * 86400000);
    }
    
    if (typeof raw === 'string') {
      let s = raw.trim();
      
      if (/\d{4}-\d{2}-\d{2}\s+\d/.test(s)) {
        s = s.split(/\s+/)[0];
      }
      
      s = s.replace(/[.\/]/g, '-');

      const m = /^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/.exec(s);
      if (m) {
        let [_, d, mo, y] = m;
        let year = Number(y.length === 2 ? (Number(y) <= 29 ? '20' + y : '19' + y) : y);
        let month = Number(mo) - 1;
        let day = Number(d);
        const dt = new Date(Date.UTC(year, month, day));
        return isNaN(dt.getTime()) ? null : dt;
      }

      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        const dt = new Date(s + 'T00:00:00Z');
        return isNaN(dt.getTime()) ? null : dt;
      }
    }
    
    return null;
  }

  // ✅ CORRECTED: Proper month calculation that handles all edge cases
  function monthsBetween(start: Date, end: Date): number {
    const sy = start.getUTCFullYear();
    const sm = start.getUTCMonth();
    const sd = start.getUTCDate();
    const ey = end.getUTCFullYear();
    const em = end.getUTCMonth();
    const ed = end.getUTCDate();
    
    // Calculate raw month difference
    let months = (ey - sy) * 12 + (em - sm);
    
    // Adjust for incomplete months
    // If the day of end date is before the day of start date, subtract 1 month
    if (ed < sd) {
      months -= 1;
    }
    
    return Math.max(0, months);
  }

  const calculateMonthsOfService = (dateOfJoining: any): number => {
    const doj = parseDOJ(dateOfJoining);
    if (!doj) return 0;
    
    const months = monthsBetween(doj, referenceDate);
    
    // Debug logging
    console.log(`DOJ: ${doj.toISOString().split('T')[0]} → MOS: ${months} months (as of ${referenceDate.toISOString().split('T')[0]})`);
    
    return months;
  };

  const processFiles = async () => {
    if (!staffFile || !workerFile || !bonusFile || !actualPercentageFile || !dueVoucherFile) {
      setError("All five files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log("📊 STEP 6: Unpaid Verification)");
      console.log("=".repeat(60));
      console.log(`✅ CORRECTED Reference Date: ${referenceDate.toISOString().split('T')[0]} (Sep 30, 2025)`);
      console.log(`Bonus Period: November 2024 - September 2025`);

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

            if (empCode && !isNaN(empCode) && percentage === SPECIAL_PERCENTAGE) {
              specialPercentageEmployees.add(empCode);
            }
          }
        }
      }

      console.log(`✅ Special percentage employees: ${specialPercentageEmployees.size}`);

      // ========== LOAD DUE VOUCHER DATA ==========
      const dueVoucherBuffer = await dueVoucherFile.arrayBuffer();
      const dueVoucherWorkbook = XLSX.read(dueVoucherBuffer);
      const dueVoucherSheet = dueVoucherWorkbook.Sheets[dueVoucherWorkbook.SheetNames[0]];
      const dueVoucherData: any[][] = XLSX.utils.sheet_to_json(dueVoucherSheet, { header: 1 });

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

      // ========== LOAD BONUS FILE WITH ACCUMULATION FOR DUPLICATES ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const hrUnpaidData: Map<
        number, 
        { unpaidHR: number; registerHR: number; dept: string; occurrences: number }
      > = new Map();

      // Process Worker sheet (1st sheet)
      if (bonusWorkbook.SheetNames.length > 0) {
        const workerSheetName = bonusWorkbook.SheetNames[0];
        console.log(`📄 Processing Bonus Worker sheet: ${workerSheetName}`);
        const workerSheet = bonusWorkbook.Sheets[workerSheetName];
        const workerData: any[][] = XLSX.utils.sheet_to_json(workerSheet, { header: 1 });

        let workerHeaderRow = -1;
        for (let i = 0; i < Math.min(10, workerData.length); i++) {
          if (
            workerData[i] &&
            workerData[i].some((v: any) => {
              const t = norm(v);
              return t === "EMPCODE" || t === "EMPLOYEECODE";
            })
          ) {
            workerHeaderRow = i;
            break;
          }
        }

        if (workerHeaderRow !== -1) {
          const headers = workerData[workerHeaderRow];
          const empCodeIdx = headers.findIndex((h: any) =>
            ["EMPCODE", "EMPLOYEECODE"].includes(norm(h))
          );
          const registerIdx = 18; // Column S
          
          let dueVCIdx = headers.findIndex((h: any) => {
            const headerStr = String(h ?? "").trim();
            return /DUE\s*VC|DUEVC/i.test(headerStr);
          });
          
          if (dueVCIdx === -1) {
            dueVCIdx = 19; // Column T
          }

          for (let i = workerHeaderRow + 1; i < workerData.length; i++) {
            const row = workerData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;
            const unpaidHR = Number(row[dueVCIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrUnpaidData.has(empCode)) {
                const existing = hrUnpaidData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.unpaidHR += unpaidHR;
                existing.occurrences += 1;
                console.log(
                  `🔄 Worker Emp ${empCode}: Duplicate found - Adding Register: ₹${registerHR.toFixed(2)}, Unpaid: ₹${unpaidHR.toFixed(2)}, Total Register: ₹${existing.registerHR.toFixed(2)}, Total Unpaid: ₹${existing.unpaidHR.toFixed(2)} (${existing.occurrences} occurrences)`
                );
              } else {
                hrUnpaidData.set(empCode, {
                  registerHR: registerHR,
                  unpaidHR: unpaidHR,
                  dept: "Worker",
                  occurrences: 1,
                });
              }
            }
          }
        }
      }

      // Process Staff sheet (2nd sheet)
      if (bonusWorkbook.SheetNames.length > 1) {
        const staffSheetName = bonusWorkbook.SheetNames[1];
        console.log(`📄 Processing Bonus Staff sheet: ${staffSheetName}`);
        const staffSheet = bonusWorkbook.Sheets[staffSheetName];
        const staffData: any[][] = XLSX.utils.sheet_to_json(staffSheet, { header: 1 });

        let staffHeaderRow = -1;
        for (let i = 0; i < Math.min(10, staffData.length); i++) {
          if (
            staffData[i] &&
            staffData[i].some((v: any) => {
              const t = norm(v);
              return t === "EMPCODE" || t === "EMPLOYEECODE";
            })
          ) {
            staffHeaderRow = i;
            break;
          }
        }

        if (staffHeaderRow !== -1) {
          const headers = staffData[staffHeaderRow];
          const empCodeIdx = headers.findIndex((h: any) =>
            ["EMPCODE", "EMPLOYEECODE"].includes(norm(h))
          );
          const registerIdx = 19; // Column T
          const unpaidIdx = 21; // Column V

          for (let i = staffHeaderRow + 1; i < staffData.length; i++) {
            const row = staffData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;
            const unpaidHR = Number(row[unpaidIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrUnpaidData.has(empCode)) {
                const existing = hrUnpaidData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.unpaidHR += unpaidHR;
                existing.occurrences += 1;
                console.log(
                  `🔄 Staff Emp ${empCode}: Duplicate found - Adding Register: ₹${registerHR.toFixed(2)}, Unpaid: ₹${unpaidHR.toFixed(2)}, Total Register: ₹${existing.registerHR.toFixed(2)}, Total Unpaid: ₹${existing.unpaidHR.toFixed(2)} (${existing.occurrences} occurrences)`
                );
              } else {
                hrUnpaidData.set(empCode, {
                  registerHR: registerHR,
                  unpaidHR: unpaidHR,
                  dept: "Staff",
                  occurrences: 1,
                });
              }
            }
          }
        }
      }

      console.log(`✅ HR Unpaid data loaded: ${hrUnpaidData.size} employees`);

      // ========== COMPUTE GROSS SALARY (EXACT STEP-5 LOGIC WITH OCTOBER ESTIMATION) ==========
      
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);
      
      const staffEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number>; dateOfJoining: any }
      > = new Map();

      for (let sheetName of staffWorkbook.SheetNames) {
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";
        
        if (EXCLUDED_MONTHS.includes(monthKey)) {
          console.log(`🚫 SKIP Staff: ${sheetName} (${monthKey}) - EXCLUDED`);
          continue;
        }
        
        if (!AVG_WINDOW.includes(monthKey)) {
          console.log(`⏭️ SKIP Staff: ${sheetName} (${monthKey}) - NOT IN WINDOW`);
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
          return /DATE.*OF.*JOINING|DOJ|JOINING.*DATE|DATE.*JOINING|D\.O\.J/i.test(headerStr);
        });

        if (dojIdx === -1) {
          for (let i = Math.max(0, headers.length - 3); i < headers.length; i++) {
            const h = String(headers[i] ?? "").trim().toLowerCase();
            if (h.includes("date") || h.includes("joining") || h.includes("doj")) {
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
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;
          const doj = (dojIdx !== -1 && row.length > dojIdx) ? row[dojIdx] : null;

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
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
        }
      }

      console.log(`✅ Staff employees: ${staffEmployees.size}`);

      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);
      
      const workerEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number>; dateOfJoining: any }
      > = new Map();

      for (let sheetName of workerWorkbook.SheetNames) {
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";
        
        if (EXCLUDED_MONTHS.includes(monthKey)) {
          console.log(`🚫 SKIP Worker: ${sheetName} (${monthKey}) - EXCLUDED`);
          continue;
        }
        
        if (!AVG_WINDOW.includes(monthKey)) {
          console.log(`⏭️ SKIP Worker: ${sheetName} (${monthKey}) - NOT IN WINDOW`);
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
          return normalized === "DEPT" || normalized === "DEPARTMENT" || normalized === "DEPTT";
        });
        
        const salary1Idx = 8; // Column I

        let dojIdx = headers.findIndex((h: any) => {
          const headerStr = String(h ?? "").trim();
          return /DATE.*OF.*JOINING|DOJ|JOINING.*DATE|DATE.*JOINING|D\.O\.J/i.test(headerStr);
        });

        if (dojIdx === -1 && headers.length > 15) {
          dojIdx = headers.length - 1;
        }

        if (empIdIdx === -1 || empNameIdx === -1) continue;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;
          const doj = (dojIdx !== -1 && row.length > dojIdx) ? row[dojIdx] : null;

          if (deptIdx !== -1) {
            const dept = String(row[deptIdx] || "").trim().toUpperCase();
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
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
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
          { name: string; dept: string; months: Map<string, number>; dateOfJoining: any }
        >
      ) => {
        for (const [empId, rec] of src) {
          let baseSum = 0;
          const monthsIncluded: { month: string; value: number }[] = [];
          
          for (const mk of AVG_WINDOW) {
            const v = rec.months.get(mk);
            if (v != null && !isNaN(Number(v)) && Number(v) > 0) {
              baseSum += Number(v);
              monthsIncluded.push({ month: mk, value: Number(v) });
            }
          }

          let estOct = 0;
          let total = baseSum;
          const hasSep2025 = rec.months.has("2025-09") && (rec.months.get("2025-09") || 0) > 0;
          const isExcluded = EXCLUDE_OCTOBER_EMPLOYEES.has(empId);

          if (isExcluded) {
            console.log(
              `🚫 EMP ${empId} (${rec.name}): IN EXCLUDE LIST - Base only = ₹${baseSum.toFixed(2)}`
            );
          } else if (hasSep2025 && monthsIncluded.length > 0) {
            const values = monthsIncluded.map(m => m.value);
            estOct = values.reduce((a, b) => a + b, 0) / values.length;
            total = baseSum + estOct;
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

      // ========== CALCULATE UNPAID WITH CORRECTED MOS ==========
      const comparison: any[] = [];

      for (const [empId, empData] of employeeData) {
        const percentage = specialPercentageEmployees.has(empId)
          ? SPECIAL_PERCENTAGE
          : DEFAULT_PERCENTAGE;

        const registerSoftware = (empData.grossSalary * percentage) / 100;

        const monthsOfService = calculateMonthsOfService(empData.dateOfJoining);

        let isEligible = true;
        if (empData.dept === "Worker") {
          isEligible = monthsOfService >= 6;
        }

        let unpaidSoftware = dueVCMap.get(empId) || 0;

        if (!isEligible) {
          unpaidSoftware = registerSoftware;
        }

        const hrData = hrUnpaidData.get(empId);
        const registerHR = hrData?.registerHR || 0;
        const unpaidHR = hrData?.unpaidHR || 0;
        const occurrences = hrData?.occurrences || 0;

        const difference = unpaidSoftware - unpaidHR;
        let status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        let validationError = "";
        if (!isEligible && Math.abs(unpaidSoftware - registerSoftware) > TOLERANCE) {
          validationError = "Employee is not eligible, so their Unpaid value must be equal to the Register.";
          status = "Error";
        }

        comparison.push({
          employeeId: empId,
          employeeName: empData.name,
          department: empData.dept,
          monthsOfService: monthsOfService,
          isEligible: isEligible,
          percentage: percentage,
          grossSalarySoftware: empData.grossSalary,
          registerSoftware: registerSoftware,
          registerHR: registerHR,
          unpaidSoftware: unpaidSoftware,
          unpaidHR: unpaidHR,
          hrOccurrences: occurrences,
          difference: difference,
          status: status,
          validationError: validationError,
        });
      }

      comparison.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(comparison);
      setFilteredData(comparison);


     

    } catch (err: any) {
      setError(`Error processing files: ${err.message}`);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  useEffect(() => {
    if (staffFile && workerFile && bonusFile && actualPercentageFile && dueVoucherFile) {
      processFiles();
    }
    // eslint-disable-next-line
  }, [staffFile, workerFile, bonusFile, actualPercentageFile, dueVoucherFile]);

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

    setFilteredData(filtered);
  }, [departmentFilter, eligibilityFilter, comparisonData]);

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      maximumFractionDigits: 2,
    }).format(value);
  };

  const exportToExcel = () => {
    const dataToExport = departmentFilter === "All" && eligibilityFilter === "All"
      ? comparisonData
      : filteredData;

    const ws = XLSX.utils.json_to_sheet(
      dataToExport.map((row) => ({
        "Employee ID": row.employeeId,
        "Employee Name": row.employeeName,
        Department: row.department,
        "Months of Service": row.monthsOfService,
        "Eligible": row.isEligible ? "YES" : "NO",
        "Percentage (%)": row.percentage,
        "Gross Salary (Software)": row.grossSalarySoftware,
        "Register (Software)": row.registerSoftware,
        "Register (HR)": row.registerHR,
        "HR Occurrences": row.hrOccurrences,
        "Unpaid (Software)": row.unpaidSoftware,
        "Unpaid (HR)": row.unpaidHR,
        Difference: row.difference,
        Status: row.status,
        "Validation Error": row.validationError || "",
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Unpaid Verification");
    XLSX.writeFile(wb, `Step6-Unpaid-Verification-CORRECTED-${departmentFilter}-${eligibilityFilter}.xlsx`);
  };

  const FileCard = ({
    title,
    file,
    description,
  }: {
    title: string;
    file: File | null;
    description: string;
  }) => (
    <div
      className={`border-2 rounded-lg p-6 ${
        file ? "border-green-300 bg-green-50" : "border-red-300 bg-red-50"
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
        <div className="bg-white rounded-lg p-4 border border-red-200">
          <div className="flex items-center gap-2 text-red-600 mb-2">
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
            <span className="font-medium">File not found</span>
          </div>
          <p className="text-xs text-gray-500">Upload in Step 1</p>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-orange-50 to-red-100 py-5 px-4">
      <div className="mx-auto max-w-7xl">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            
            <div className="flex gap-3">
              <button
                onClick={() => router.push("/step5")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                ← Back to Step 5
              </button>
              <button
                onClick={() => router.push("/")}
                className="px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition"
              >
                Back to Step 1
              </button>
              <button
                onClick={() => router.push("/step7")}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-green-700 transition"
              >
                Move to Step 7
              </button>
            </div>
          </div>

          

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data (Nov-24 to Sep-25)"
            />
            <FileCard
              title="Indiana Worker"
              file={workerFile}
              description="Worker salary data (excludes Dept C, CASH, A)"
            />
            <FileCard
              title="Bonus Calculation Sheet"
              file={bonusFile}
              description="Register & Unpaid (HR): Worker Col T, Staff Col V"
            />
            <FileCard
              title="Actual Percentage Data"
              file={actualPercentageFile}
              description="Employees with 12% bonus"
            />
            <FileCard
              title="Due Voucher List"
              file={dueVoucherFile}
              description="DUE VC values for Unpaid (Software)"
            />
          </div>

          {[staffFile, workerFile, bonusFile, actualPercentageFile, dueVoucherFile].filter(
            Boolean
          ).length < 5 && (
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
                  Processing with corrected MOS calculation (Sep 30, 2025)...
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
                <div className="flex items-center gap-4">
                  <h2 className="text-xl font-bold text-gray-800">
                    Unpaid Verification Results (CORRECTED MOS)
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
                <table className="w-full border-collapse text-sm">
                  <thead>
                    <tr className="bg-gray-100">
                      <th className="border border-gray-300 px-3 py-2 text-left">
                        Emp ID
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-left">
                        Name
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-left">
                        Dept
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-center">
                        MOS
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-center">
                        Eligible
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-center">
                        %
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Gross (SW)
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Register (SW)
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Register (HR)
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-center">
                        HR Entries
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Unpaid (SW)
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Unpaid (HR)
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-right">
                        Diff
                      </th>
                      <th className="border border-gray-300 px-3 py-2 text-center">
                        Status
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, idx) => (
                      <tr
                        key={idx}
                        className={`${
                          idx % 2 === 0 ? "bg-white" : "bg-gray-50"
                        } ${row.validationError ? "bg-red-50" : ""} ${
                          row.hrOccurrences > 1 ? "bg-yellow-50" : ""
                        } ${row.employeeId === 922 ? "bg-green-100 font-bold" : ""}`}
                      >
                        <td className="border border-gray-300 px-3 py-2">
                          {row.employeeId}
                          {row.employeeId === 922 && (
                            <span className="ml-1 text-xs text-green-700">✅</span>
                          )}
                        </td>
                        <td className="border border-gray-300 px-3 py-2">
                          {row.employeeName}
                        </td>
                        <td className="border border-gray-300 px-3 py-2">
                          <span
                            className={`px-2 py-1 rounded text-xs font-medium ${
                              row.department === "Staff"
                                ? "bg-blue-100 text-blue-800"
                                : "bg-purple-100 text-purple-800"
                            }`}
                          >
                            {row.department}
                          </span>
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-center">
                          {row.monthsOfService || 0}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-center">
                          <span
                            className={`px-2 py-1 rounded text-xs font-medium ${
                              row.isEligible
                                ? "bg-green-100 text-green-800"
                                : "bg-red-100 text-red-800"
                            }`}
                          >
                            {row.isEligible ? "YES" : "NO"}
                          </span>
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-center">
                          <span
                            className={`px-2 py-1 rounded text-xs font-medium ${
                              row.percentage === 12.0
                                ? "bg-yellow-100 text-yellow-800"
                                : "bg-gray-100 text-gray-800"
                            }`}
                          >
                            {row.percentage}%
                          </span>
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-right">
                          {formatCurrency(row.grossSalarySoftware)}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-right">
                          {formatCurrency(row.registerSoftware)}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-right">
                          {formatCurrency(row.registerHR)}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-center">
                          <span
                            className={`px-2 py-1 rounded text-xs font-medium ${
                              row.hrOccurrences > 1
                                ? "bg-orange-100 text-orange-800"
                                : "bg-gray-100 text-gray-800"
                            }`}
                          >
                            {row.hrOccurrences}
                          </span>
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-right font-medium text-blue-600">
                          {formatCurrency(row.unpaidSoftware)}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-right font-medium text-purple-600">
                          {formatCurrency(row.unpaidHR)}
                        </td>
                        <td
                          className={`border border-gray-300 px-3 py-2 text-right font-medium ${
                            Math.abs(row.difference) <= TOLERANCE
                              ? "text-green-600"
                              : "text-red-600"
                          }`}
                        >
                          {formatCurrency(row.difference)}
                        </td>
                        <td className="border border-gray-300 px-3 py-2 text-center">
                          <span
                            className={`px-3 py-1 rounded-full text-xs font-medium ${
                              row.status === "Match"
                                ? "bg-green-100 text-green-800"
                                : row.status === "Error"
                                ? "bg-red-100 text-red-800"
                                : "bg-orange-100 text-orange-800"
                            }`}
                            title={row.validationError || ""}
                          >
                            {row.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {filteredData.some((r) => r.validationError) && (
                <div className="mt-4 bg-red-50 border border-red-200 rounded-lg p-4">
                  <h3 className="font-medium text-red-800 mb-2">
                    ⚠️ Validation Errors Found
                  </h3>
                  <div className="text-sm text-red-700 space-y-1">
                    {filteredData
                      .filter((r) => r.validationError)
                      .map((row) => (
                        <p key={row.employeeId}>
                          <strong>Emp {row.employeeId} ({row.employeeName}):</strong>{" "}
                          {row.validationError}
                        </p>
                      ))}
                  </div>
                </div>
              )}

              <div className="mt-4 flex justify-between items-center text-sm text-gray-600">
                <div>
                  Total: {filteredData.length} | Staff:{" "}
                  {filteredData.filter((r) => r.department === "Staff").length}{" "}
                  | Worker:{" "}
                  {filteredData.filter((r) => r.department === "Worker").length}
                </div>
                <div>
                  Eligible:{" "}
                  {filteredData.filter((r) => r.isEligible).length} |
                  Not Eligible:{" "}
                  {filteredData.filter((r) => !r.isEligible).length}
                </div>
                <div>
                  Matches:{" "}
                  {filteredData.filter((r) => r.status === "Match").length} |
                  Mismatches:{" "}
                  {filteredData.filter((r) => r.status === "Mismatch").length} |
                  Errors:{" "}
                  {filteredData.filter((r) => r.status === "Error").length} |
                  Duplicates:{" "}
                  {filteredData.filter((r) => r.hrOccurrences > 1).length}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}