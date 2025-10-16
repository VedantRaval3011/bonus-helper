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

  // Sorting state
  const [sortConfig, setSortConfig] = useState<{
    key: string | null;
    direction: "asc" | "desc" | null;
  }>({ key: null, direction: null });

  // === Step 6 Audit Helpers ===
  const TOLERANCE_STEP6 = 12;

  async function postAuditMessagesStep6(items: any[], batchId?: string) {
    const bid =
      batchId ||
      (typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2));
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId: bid, step: 6, items }),
    });
    return bid;
  }

  function buildStep6MismatchMessages(rows: any[]) {
    const items: any[] = [];
    for (const r of rows) {
      if (r?.status === "Mismatch" || r?.status === "Error") {
        items.push({
          level: r.status === "Error" ? "error" : "warning",
          tag: r.status === "Error" ? "validation-error" : "mismatch",
          text: `[step6] ${r.employeeId} ${r.employeeName} ${
            r.status === "Error"
              ? "Validation Error"
              : `diff=${Number(r.difference ?? 0).toFixed(2)}`
          }`,
          scope:
            r.department === "Staff"
              ? "staff"
              : r.department === "Worker"
              ? "worker"
              : "global",
          source: "step6",
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
            alreadyPaidSoftware: r.alreadyPaidSoftware,
            alreadyPaidHR: r.alreadyPaidHR,
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
    const matches = rows.filter((r) => r.status === "Match").length;
    const mismatches = rows.filter((r) => r.status === "Mismatch").length;
    const errors = rows.filter((r) => r.status === "Error").length;

    const staffRows = rows.filter((r) => r.department === "Staff");
    const workerRows = rows.filter((r) => r.department === "Worker");

    const staffMismatch = staffRows.filter(
      (r) => r.status === "Mismatch" || r.status === "Error"
    ).length;
    const workerMismatch = workerRows.filter(
      (r) => r.status === "Mismatch" || r.status === "Error"
    ).length;

    const eligible = rows.filter((r) => r.isEligible).length;
    const notEligible = rows.filter((r) => !r.isEligible).length;
    const duplicates = rows.filter((r) => r.hrOccurrences > 1).length;

    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const staffRegSWSum = sum(
      staffRows.map((r) => Number(r.registerSoftware || 0))
    );
    const staffRegHRSum = sum(staffRows.map((r) => Number(r.registerHR || 0)));
    const staffUnpaidSWSum = sum(
      staffRows.map((r) => Number(r.unpaidSoftware || 0))
    );
    const staffUnpaidHRSum = sum(staffRows.map((r) => Number(r.unpaidHR || 0)));
    const staffAlreadyPaidSWSum = sum(
      staffRows.map((r) => Number(r.alreadyPaidSoftware || 0))
    );
    const staffAlreadyPaidHRSum = sum(
      staffRows.map((r) => Number(r.alreadyPaidHR || 0))
    );

    const workerRegSWSum = sum(
      workerRows.map((r) => Number(r.registerSoftware || 0))
    );
    const workerRegHRSum = sum(
      workerRows.map((r) => Number(r.registerHR || 0))
    );
    const workerUnpaidSWSum = sum(
      workerRows.map((r) => Number(r.unpaidSoftware || 0))
    );
    const workerUnpaidHRSum = sum(
      workerRows.map((r) => Number(r.unpaidHR || 0))
    );
    const workerAlreadyPaidSWSum = sum(
      workerRows.map((r) => Number(r.alreadyPaidSoftware || 0))
    );
    const workerAlreadyPaidHRSum = sum(
      workerRows.map((r) => Number(r.alreadyPaidHR || 0))
    );

    return {
      level: "info",
      tag: "summary",
      text: `Step6 run: total=${total} match=${matches} mismatch=${mismatches} error=${errors}`,
      scope: "global",
      source: "step6",
      meta: {
        totals: {
          total,
          matches,
          mismatches,
          errors,
          eligible,
          notEligible,
          duplicates,
          tolerance: TOLERANCE_STEP6,
        },
        staff: {
          count: staffRows.length,
          issues: staffMismatch,
          registerSWSum: staffRegSWSum,
          registerHRSum: staffRegHRSum,
          unpaidSWSum: staffUnpaidSWSum,
          unpaidHRSum: staffUnpaidHRSum,
          alreadyPaidSWSum: staffAlreadyPaidSWSum,
          alreadyPaidHRSum: staffAlreadyPaidHRSum,
        },
        worker: {
          count: workerRows.length,
          issues: workerMismatch,
          registerSWSum: workerRegSWSum,
          registerHRSum: workerRegHRSum,
          unpaidSWSum: workerUnpaidSWSum,
          unpaidHRSum: workerUnpaidHRSum,
          alreadyPaidSWSum: workerAlreadyPaidSWSum,
          alreadyPaidHRSum: workerAlreadyPaidHRSum,
        },
      },
    };
  }

  async function handleSaveAuditStep6(rows: any[]) {
    if (!rows || rows.length === 0) return;
    const items = [
      buildStep6SummaryMessage(rows),
      ...buildStep6MismatchMessages(rows),
    ];
    if (items.length === 0) return;
    await postAuditMessagesStep6(items);
  }

  function djb2Hash(str: string) {
    let h = 5381;
    for (let i = 0; i < str.length; i++) h = (h << 5) + h + str.charCodeAt(i);
    return (h >>> 0).toString(36);
  }

  function buildRunKeyStep6(rows: any[]) {
    const sig = rows
      .map(
        (r) =>
          `${r.employeeId}|${r.department}|${Number(r.unpaidSoftware) || 0}|${
            Number(r.unpaidHR) || 0
          }|${Number(r.alreadyPaidSoftware) || 0}|${Number(r.alreadyPaidHR) || 0}|${Number(r.difference) || 0}|${r.status}|${r.isEligible}`
      )
      .join(";");
    return djb2Hash(sig);
  }

  useEffect(() => {
    if (typeof window === "undefined") return;
    if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

    const runKey = buildRunKeyStep6(comparisonData);
    const markerKey = `audit_step6_${runKey}`;

    if (sessionStorage.getItem(markerKey)) return;

    sessionStorage.setItem(markerKey, "1");
    const deterministicBatchId = `step6-${runKey}`;

    const items = [
      buildStep6SummaryMessage(comparisonData),
      ...buildStep6MismatchMessages(comparisonData),
    ];
    postAuditMessagesStep6(items, deterministicBatchId).catch((err) => {
      console.error("Auto-audit step6 failed", err);
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
      /bonus.*final.*calculation|bonus.*2024-25|sci.*prec.*final.*calculation|final.*calculation.*sheet|nrtm.*final.*bonus.*calculation|nutra.*bonus.*calculation|sci.*prec.*life.*science.*bonus.*calculation/i.test(
        s.file.name
      )
  );


  const dueVoucherFile =
    pickFile((s) => s.type === "Due-Voucher-List") ??
    pickFile((s) => !!s.file && /due.*voucher/i.test(s.file.name));

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

    const months = monthsBetween(doj, referenceDate);

    console.log(
      `DOJ: ${doj.toISOString().split("T")[0]} → MOS: ${months} months (as of ${
        referenceDate.toISOString().split("T")[0]
      })`
    );

    return months;
  };

  const processFiles = async () => {
    // ✅ Updated: Only require 4 files now (removed actualPercentageFile)
    if (!staffFile || !workerFile || !bonusFile || !dueVoucherFile) {
      setError("Required files: Indiana Staff, Indiana Worker, Bonus Calculation Sheet, and Due Voucher List");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log("📊 STEP 6: Unpaid Verification (Simplified - No Percentage File)");
      console.log("=".repeat(60));
      console.log(
        `✅ Reference Date: ${
          referenceDate.toISOString().split("T")[0]
        } (Oct 30, 2025)`
      );
      console.log(`Bonus Period: November 2024 - September 2025`);
      console.log(`✅ Default Percentage: ${DEFAULT_PERCENTAGE}%`);

      // ✅ Load Already Paid (Software) from Due Voucher file column F
      const dueVoucherBuffer = await dueVoucherFile.arrayBuffer();
      const dueVoucherWorkbook = XLSX.read(dueVoucherBuffer);
      const dueVoucherSheet =
        dueVoucherWorkbook.Sheets[dueVoucherWorkbook.SheetNames[0]];
      const dueVoucherData: any[][] = XLSX.utils.sheet_to_json(
        dueVoucherSheet,
        { header: 1 }
      );

      const dueVCMap: Map<number, number> = new Map();
      const alreadyPaidSoftwareMap: Map<number, number> = new Map();
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
        const alreadyPaidIdx = 5; // Column F

        if (empCodeIdx !== -1 && dueVCIdx !== -1) {
          for (let i = dueVCHeaderRow + 1; i < dueVoucherData.length; i++) {
            const row = dueVoucherData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const dueVC = Number(row[dueVCIdx]) || 0;
            const alreadyPaid = Number(row[alreadyPaidIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              dueVCMap.set(empCode, dueVC);
              alreadyPaidSoftwareMap.set(empCode, alreadyPaid);
            }
          }
        }
      }

      console.log(`✅ Due VC data loaded: ${dueVCMap.size} employees`);
      console.log(`✅ Already Paid (Software) data loaded: ${alreadyPaidSoftwareMap.size} employees`);

      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const hrUnpaidData: Map<
        number,
        {
          unpaidHR: number;
          registerHR: number;
          alreadyPaidHR: number;
          dept: string;
          occurrences: number;
        }
      > = new Map();

      // Process Worker sheet (first sheet)
      if (bonusWorkbook.SheetNames.length > 0) {
        const workerSheetName = bonusWorkbook.SheetNames[0];
        console.log(`📄 Processing Bonus Worker sheet: ${workerSheetName}`);
        const workerSheet = bonusWorkbook.Sheets[workerSheetName];
        const workerData: any[][] = XLSX.utils.sheet_to_json(workerSheet, {
          header: 1,
        });

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
          const registerIdx = 19; // Column T

          let dueVCIdx = headers.findIndex((h: any) => {
            const headerStr = String(h ?? "").trim();
            return /DUE\s*VC|DUEVC/i.test(headerStr);
          });

          if (dueVCIdx === -1) {
            dueVCIdx = 21; // Column V
          }

          const alreadyPaidIdx = 22; // Column W

          for (let i = workerHeaderRow + 1; i < workerData.length; i++) {
            const row = workerData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;
            const unpaidHR = Number(row[dueVCIdx]) || 0;
            const alreadyPaidHR = Number(row[alreadyPaidIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrUnpaidData.has(empCode)) {
                const existing = hrUnpaidData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.unpaidHR += unpaidHR;
                existing.alreadyPaidHR += alreadyPaidHR;
                existing.occurrences += 1;
              } else {
                hrUnpaidData.set(empCode, {
                  registerHR: registerHR,
                  unpaidHR: unpaidHR,
                  alreadyPaidHR: alreadyPaidHR,
                  dept: "Worker",
                  occurrences: 1,
                });
              }
            }
          }
        }
      }

      // Process Staff sheet (second sheet)
      if (bonusWorkbook.SheetNames.length > 1) {
        const staffSheetName = bonusWorkbook.SheetNames[1];
        console.log(`📄 Processing Bonus Staff sheet: ${staffSheetName}`);
        const staffSheet = bonusWorkbook.Sheets[staffSheetName];
        const staffData: any[][] = XLSX.utils.sheet_to_json(staffSheet, {
          header: 1,
        });

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
          const alreadyPaidIdx = 22; // Column W

          for (let i = staffHeaderRow + 1; i < staffData.length; i++) {
            const row = staffData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;
            const unpaidHR = Number(row[unpaidIdx]) || 0;
            const alreadyPaidHR = Number(row[alreadyPaidIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrUnpaidData.has(empCode)) {
                const existing = hrUnpaidData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.unpaidHR += unpaidHR;
                existing.alreadyPaidHR += alreadyPaidHR;
                existing.occurrences += 1;
              } else {
                hrUnpaidData.set(empCode, {
                  registerHR: registerHR,
                  unpaidHR: unpaidHR,
                  alreadyPaidHR: alreadyPaidHR,
                  dept: "Staff",
                  occurrences: 1,
                });
              }
            }
          }
        }
      }

      console.log(`✅ HR Unpaid data loaded: ${hrUnpaidData.size} employees`);

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

      const comparison: any[] = [];

      for (const [empId, empData] of employeeData) {
        // ✅ Use default percentage for all employees
        const percentage = DEFAULT_PERCENTAGE;

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

        const alreadyPaidSoftware = alreadyPaidSoftwareMap.get(empId) || 0;

        const hrData = hrUnpaidData.get(empId);
        const registerHR = hrData?.registerHR || 0;
        const unpaidHR = hrData?.unpaidHR || 0;
        const alreadyPaidHR = hrData?.alreadyPaidHR || 0;
        const occurrences = hrData?.occurrences || 0;

        const difference = unpaidSoftware - unpaidHR;
        let status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        let validationError = "";
        if (
          !isEligible &&
          Math.abs(unpaidSoftware - registerSoftware) > TOLERANCE
        ) {
          validationError =
            "Employee is not eligible, so their Unpaid value must be equal to the Register.";
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
          alreadyPaidSoftware: alreadyPaidSoftware,
          alreadyPaidHR: alreadyPaidHR,
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
    // ✅ Updated: Only check for 4 required files
    if (staffFile && workerFile && bonusFile && dueVoucherFile) {
      processFiles();
    }
    // eslint-disable-next-line
  }, [staffFile, workerFile, bonusFile, dueVoucherFile]);

  useEffect(() => {
    let filtered = [...comparisonData];

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

    if (sortConfig.key && sortConfig.direction) {
      filtered.sort((a, b) => {
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

    setFilteredData(filtered);
  }, [departmentFilter, eligibilityFilter, comparisonData, sortConfig]);

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
    const dataToExport =
      departmentFilter === "All" && eligibilityFilter === "All"
        ? comparisonData
        : filteredData;

    const ws = XLSX.utils.json_to_sheet(
      dataToExport.map((row) => ({
        "Employee ID": row.employeeId,
        "Employee Name": row.employeeName,
        Department: row.department,
        "Months of Service": row.monthsOfService,
        Eligible: row.isEligible ? "YES" : "NO",
        "Percentage (%)": row.percentage,
        "Gross Salary (Software)": row.grossSalarySoftware,
        "Register (Software)": row.registerSoftware,
        "Register (HR)": row.registerHR,
        "HR Occurrences": row.hrOccurrences,
        "Unpaid (Software)": row.unpaidSoftware,
        "Unpaid (HR)": row.unpaidHR,
        "Already Paid (Software)": row.alreadyPaidSoftware,
        "Already Paid (HR)": row.alreadyPaidHR,
        Difference: row.difference,
        Status: row.status,
        "Validation Error": row.validationError || "",
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Unpaid Verification");
    XLSX.writeFile(
      wb,
      `Step6-Unpaid-Verification-${departmentFilter}-${eligibilityFilter}.xlsx`
    );
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
      <div className="max-w-full">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 6 - Unpaid Verification
              </h1>
              <p className="text-sm text-gray-600 mt-2">
                Default Percentage: {DEFAULT_PERCENTAGE}% for all employees
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("step5")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                ← Back to Step 5
              </button>
              <button
                onClick={() => router.push("step1")}
                className="px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition"
              >
                Back to Step 1
              </button>
              <button
                onClick={() => router.push("step7")}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-green-700 transition"
              >
                Move to Step 7
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
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
              description="Register, Unpaid & Already Paid (HR): Worker Col S,T,W | Staff Col T,V,W"
            />
            <FileCard
              title="Due Voucher List"
              file={dueVoucherFile}
              description="DUE VC & Already Paid (Software) - Col F"
            />
          </div>

          {[staffFile, workerFile, bonusFile, dueVoucherFile].filter(Boolean)
            .length < 4 && (
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
                <p className="text-blue-800">Processing unpaid verification...</p>
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
                    Unpaid Verification Results
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

              <div className="relative">
                <div className="overflow-x-auto">
                  <div className="max-h-[600px] overflow-y-auto">
                    <table className="w-full border-collapse text-sm">
                      <thead className="sticky top-0 bg-gray-100 z-10">
                        <tr className="bg-gray-100">
                          <th
                            className="border border-gray-300 px-3 py-2 text-left cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("employeeId")}
                          >
                            <div className="flex items-center">
                              Emp ID
                              <SortIcon columnKey="employeeId" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-left cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("employeeName")}
                          >
                            <div className="flex items-center">
                              Name
                              <SortIcon columnKey="employeeName" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-left cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("department")}
                          >
                            <div className="flex items-center">
                              Dept
                              <SortIcon columnKey="department" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-center cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("monthsOfService")}
                          >
                            <div className="flex items-center justify-center">
                              MOS
                              <SortIcon columnKey="monthsOfService" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-center cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("isEligible")}
                          >
                            <div className="flex items-center justify-center">
                              Eligible
                              <SortIcon columnKey="isEligible" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-center cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("percentage")}
                          >
                            <div className="flex items-center justify-center">
                              %
                              <SortIcon columnKey="percentage" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("grossSalarySoftware")}
                          >
                            <div className="flex items-center justify-end">
                              Gross (SW)
                              <SortIcon columnKey="grossSalarySoftware" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("registerSoftware")}
                          >
                            <div className="flex items-center justify-end">
                              Register (SW)
                              <SortIcon columnKey="registerSoftware" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("registerHR")}
                          >
                            <div className="flex items-center justify-end">
                              Register (HR)
                              <SortIcon columnKey="registerHR" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("unpaidSoftware")}
                          >
                            <div className="flex items-center justify-end">
                              Unpaid (SW)
                              <SortIcon columnKey="unpaidSoftware" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("unpaidHR")}
                          >
                            <div className="flex items-center justify-end">
                              Unpaid (HR)
                              <SortIcon columnKey="unpaidHR" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("alreadyPaidSoftware")}
                          >
                            <div className="flex items-center justify-end">
                              Already Paid (SW)
                              <SortIcon columnKey="alreadyPaidSoftware" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("alreadyPaidHR")}
                          >
                            <div className="flex items-center justify-end">
                              Already Paid (HR)
                              <SortIcon columnKey="alreadyPaidHR" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-right cursor-pointer hover:bg-gray-200 select-none"
                            onClick={() => handleSort("difference")}
                          >
                            <div className="flex items-center justify-end">
                              Diff
                              <SortIcon columnKey="difference" />
                            </div>
                          </th>
                          <th
                            className="border border-gray-300 px-3 py-2 text-center cursor-pointer hover:bg-gray-200 select-none"
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
                            className={`${
                              idx % 2 === 0 ? "bg-white" : "bg-gray-50"
                            } ${row.validationError ? "bg-red-50" : ""} ${
                              row.hrOccurrences > 1 ? "bg-yellow-50" : ""
                            }`}
                          >
                            <td className="border border-gray-300 px-3 py-2">
                              {row.employeeId}
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
                              <span className="px-2 py-1 rounded text-xs font-medium bg-gray-100 text-gray-800">
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
                            <td className="border border-gray-300 px-3 py-2 text-right font-medium text-blue-600">
                              {formatCurrency(row.unpaidSoftware)}
                            </td>
                            <td className="border border-gray-300 px-3 py-2 text-right font-medium text-purple-600">
                              {formatCurrency(row.unpaidHR)}
                            </td>
                            <td className="border border-gray-300 px-3 py-2 text-right font-medium text-green-600">
                              {formatCurrency(row.alreadyPaidSoftware)}
                            </td>
                            <td className="border border-gray-300 px-3 py-2 text-right font-medium text-teal-600">
                              {formatCurrency(row.alreadyPaidHR)}
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
                      <tfoot className="sticky bottom-0 bg-yellow-100 z-10">
                        <tr className="font-bold border-t-4 border-yellow-600">
                          <td
                            colSpan={9}
                            className="border border-gray-300 px-3 py-3 text-right bg-yellow-100"
                          >
                            GRAND TOTAL:
                          </td>
                          <td className="border border-gray-300 px-3 py-3 text-right text-blue-700 bg-yellow-100">
                            {formatCurrency(
                              filteredData.reduce(
                                (sum, row) => sum + (row.unpaidSoftware || 0),
                                0
                              )
                            )}
                          </td>
                          <td className="border border-gray-300 px-3 py-3 text-right text-purple-700 bg-yellow-100">
                            {formatCurrency(
                              filteredData.reduce(
                                (sum, row) => sum + (row.unpaidHR || 0),
                                0
                              )
                            )}
                          </td>
                          <td className="border border-gray-300 px-3 py-3 text-right text-green-700 bg-yellow-100">
                            {formatCurrency(
                              filteredData.reduce(
                                (sum, row) =>
                                  sum + (row.alreadyPaidSoftware || 0),
                                0
                              )
                            )}
                          </td>
                          <td className="border border-gray-300 px-3 py-3 text-right text-teal-700 bg-yellow-100">
                            {formatCurrency(
                              filteredData.reduce(
                                (sum, row) => sum + (row.alreadyPaidHR || 0),
                                0
                              )
                            )}
                          </td>
                          <td
                            colSpan={2}
                            className="border border-gray-300 bg-yellow-100"
                          ></td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
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
                          <strong>
                            Emp {row.employeeId} ({row.employeeName}):
                          </strong>{" "}
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
                  Eligible: {filteredData.filter((r) => r.isEligible).length} |
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
