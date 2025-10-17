"use client";

import React, { useState, useEffect, useRef } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

export default function Step5Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [departmentFilter, setDepartmentFilter] = useState<string>("All");
  const horizontalScrollRef = useRef<HTMLDivElement>(null);
  const proxyScrollRef = useRef<HTMLDivElement>(null);

  // === Step 5 Audit Helpers ===
  const TOLERANCE_STEP5 = 12; // Step 5 uses ±12 to mark Match vs Mismatch

  async function postAuditMessagesStep5(items: any[], batchId?: string) {
    const bid =
      batchId ||
      (typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2));
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId: bid, step: 5, items }),
    });
    return bid;
  }

  useEffect(() => {
    const main = horizontalScrollRef.current;
    const proxy = proxyScrollRef.current;
    if (!main || !proxy) return;
    const onProxyScroll = () => {
      main.scrollLeft = proxy.scrollLeft;
    };
    const onMainScroll = () => {
      proxy.scrollLeft = main.scrollLeft;
    };
    proxy.addEventListener("scroll", onProxyScroll);
    main.addEventListener("scroll", onMainScroll);
    return () => {
      proxy.removeEventListener("scroll", onProxyScroll);
      main.removeEventListener("scroll", onMainScroll);
    };
  }, []);

  function buildStep5MismatchMessages(rows: any[]) {
    const items: any[] = [];
    for (const r of rows) {
      if (r?.status === "Mismatch") {
        items.push({
          level: "error",
          tag: "mismatch",
          text: `[step5] ${r.employeeId} ${r.employeeName} diff=${Number(
            r.difference ?? 0
          ).toFixed(2)}`,
          scope:
            r.department === "Staff"
              ? "staff"
              : r.department === "Worker"
              ? "worker"
              : "global",
          source: "step5",
          meta: {
            employeeId: r.employeeId,
            name: r.employeeName,
            department: r.department,
            dateOfJoining: r.dateOfJoining,
            percentage: r.percentage,
            grossSal: r.grossSal,
            calculatedValue: r.calculatedValue,
            gross2HR: r.gross2HR,
            diff: r.difference,
            tolerance: TOLERANCE_STEP5,
          },
        });
      }
    }
    return items;
  }

  // Sorting state and helper, at the top of your Step5Page component:
  const [sortConfig, setSortConfig] = useState<{
    key: string;
    direction: "asc" | "desc" | null;
  }>({ key: "", direction: null });

  const handleSort = (key: string) => {
    let direction: "asc" | "desc" | null = "asc";
    if (sortConfig.key === key) {
      direction =
        sortConfig.direction === "asc"
          ? "desc"
          : sortConfig.direction === "desc"
          ? null
          : "asc";
    }
    setSortConfig({ key, direction });
  };

  useEffect(() => {
    let dataToSort = departmentFilter === "All" ? comparisonData : filteredData;
    if (sortConfig.key && sortConfig.direction) {
      dataToSort = [...dataToSort].sort((a, b) => {
        let aValue = a[sortConfig.key],
          bValue = b[sortConfig.key];
        // Numeric columns
        if (
          [
            "employeeId",
            "grossSalarySoftware",
            "adjustedGross",
            "percentage",
            "registerSoftware",
            "registerHR",
            "hrOccurrences",
            "difference",
          ].includes(sortConfig.key)
        ) {
          aValue = Number(aValue) || 0;
          bValue = Number(bValue) || 0;
        } else {
          aValue = String(aValue || "").toUpperCase();
          bValue = String(bValue || "").toUpperCase();
        }
        if (aValue < bValue) return sortConfig.direction === "asc" ? -1 : 1;
        if (aValue > bValue) return sortConfig.direction === "asc" ? 1 : -1;
        return 0;
      });
    }
    setFilteredData(dataToSort);
  }, [sortConfig, comparisonData, departmentFilter]);

  const SortArrows = ({ columnKey }: { columnKey: string }) => {
    const isActive = sortConfig.key === columnKey;
    return (
      <div className="inline-flex flex-col ml-1">
        <button
          type="button"
          className={`leading-none ${
            isActive && sortConfig.direction === "asc"
              ? "text-purple-600"
              : "text-gray-400 hover:text-gray-600"
          }`}
          onClick={() => handleSort(columnKey)}
          tabIndex={-1}
          title="Sort Ascending"
        >
          <svg className="w-3 h-3" fill="currentColor" viewBox="0 0 24 24">
            <path d="M7 14l5-5 5 5z" />
          </svg>
        </button>
        <button
          type="button"
          className={`leading-none -mt-1 ${
            isActive && sortConfig.direction === "desc"
              ? "text-purple-600"
              : "text-gray-400 hover:text-gray-600"
          }`}
          onClick={() => handleSort(columnKey)}
          tabIndex={-1}
          title="Sort Descending"
        >
          <svg className="w-3 h-3" fill="currentColor" viewBox="0 0 24 24">
            <path d="M7 10l5 5 5-5z" />
          </svg>
        </button>
      </div>
    );
  };

  function buildStep5SummaryMessage(rows: any[]) {
    const total = rows.length || 0;
    const mismatches = rows.filter((r) => r.status === "Mismatch").length;
    const matches = total - mismatches;

    const staffRows = rows.filter((r) => r.department === "Staff");
    const workerRows = rows.filter((r) => r.department === "Worker");

    const staffMismatch = staffRows.filter(
      (r) => r.status === "Mismatch"
    ).length;
    const workerMismatch = workerRows.filter(
      (r) => r.status === "Mismatch"
    ).length;

    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const staffGrossSalSum = sum(staffRows.map((r) => Number(r.grossSal || 0)));
    const staffCalcSum = sum(
      staffRows.map((r) => Number(r.calculatedValue || 0))
    );
    const staffHRSum = sum(staffRows.map((r) => Number(r.gross2HR || 0)));

    const workerGrossSalSum = sum(
      workerRows.map((r) => Number(r.grossSal || 0))
    );
    const workerCalcSum = sum(
      workerRows.map((r) => Number(r.calculatedValue || 0))
    );
    const workerHRSum = sum(workerRows.map((r) => Number(r.gross2HR || 0)));

    return {
      level: "info",
      tag: "summary",
      text: `Step5 run: total=${total} match=${matches} mismatch=${mismatches}`,
      scope: "global",
      source: "step5",
      meta: {
        totals: { total, matches, mismatches, tolerance: TOLERANCE_STEP5 },
        staff: {
          count: staffRows.length,
          mismatches: staffMismatch,
          grossSalSum: staffGrossSalSum,
          calculatedSum: staffCalcSum,
          hrGross2Sum: staffHRSum,
        },
        worker: {
          count: workerRows.length,
          mismatches: workerMismatch,
          grossSalSum: workerGrossSalSum,
          calculatedSum: workerCalcSum,
          hrGross2Sum: workerHRSum,
        },
      },
    };
  }

  async function handleSaveAuditStep5(rows: any[]) {
    if (!rows || rows.length === 0) return;
    const items = [
      buildStep5SummaryMessage(rows),
      ...buildStep5MismatchMessages(rows),
    ];
    if (items.length === 0) return;
    await postAuditMessagesStep5(items);
  }
  // Stable hash for run signature
  function djb2Hash(str: string) {
    let h = 5381;
    for (let i = 0; i < str.length; i++) h = (h << 5) + h + str.charCodeAt(i);
    return (h >>> 0).toString(36);
  }

  function buildRunKeyStep5(rows: any[]) {
    const sig = rows
      .map(
        (r) =>
          `${r.employeeId}|${r.department}|${Number(r.grossSal) || 0}|${
            Number(r.calculatedValue) || 0
          }|${Number(r.gross2HR) || 0}|${Number(r.difference) || 0}|${r.status}`
      )
      .join(";");
    return djb2Hash(sig);
  }

  useEffect(() => {
    if (typeof window === "undefined") return; // SSR guard
    if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

    const runKey = buildRunKeyStep5(comparisonData);
    const markerKey = `audit_step5_${runKey}`;

    if (sessionStorage.getItem(markerKey)) return; // prevent duplicate on refresh/StrictMode

    sessionStorage.setItem(markerKey, "1");
    const deterministicBatchId = `step5-${runKey}`;

    const items = [
      buildStep5SummaryMessage(comparisonData),
      ...buildStep5MismatchMessages(comparisonData),
    ];

    postAuditMessagesStep5(items, deterministicBatchId).catch((err) => {
      console.error("Auto-audit step5 failed", err);
      sessionStorage.removeItem(markerKey); // allow retry on next refresh if failed
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
  const EXCLUDED_DEPARTMENTS = ["C", "CASH", "A"];

  const EXCLUDE_OCTOBER_EMPLOYEES = new Set<number>([
    937, 1039, 1065, 1105, 59, 161,
  ]);

  const INCLUDE_ZEROS_IN_AVG = new Set<number>([20, 27, 882, 898, 999]);

  const EXCLUDE_ZEROS_IN_AVG = new Set<number>([1054]);

  const EMPLOYEE_START_MONTHS: Record<number, string> = {
    999: "2024-12",
  };

  const DEFAULT_PERCENTAGE = 8.33;
  const SPECIAL_PERCENTAGE = 12.0;
  const SPECIAL_GROSS_MULTIPLIER = 0.6; // 60% of gross for 12% employees
  const TOLERANCE = 12;

  const processFiles = async () => {
    if (!staffFile || !workerFile || !bonusFile || !actualPercentageFile) {
      setError("All four files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log(
        "📊 STEP 5: Register Calculation (Using Step-3 Gross Values)"
      );
      console.log("=".repeat(60));
      console.log(
        "⚡ NEW RULE: 12% employees = (Step3-Gross × 60%) × 12% = Gross × 7.2%"
      );
      console.log("📌 DEFAULT RULE: 8.33% employees = Step3-Gross × 8.33%");
      console.log(
        "🚫 OCTOBER EXCLUDE LIST:",
        Array.from(EXCLUDE_OCTOBER_EMPLOYEES).join(", ")
      );
      console.log(
        "🚫 EXCLUDED DEPARTMENTS (Worker):",
        EXCLUDED_DEPARTMENTS.join(", ")
      );
      console.log("=".repeat(60));

      // ========== LOAD ACTUAL PERCENTAGE DATA ==========
      const actualPercentageBuffer = await actualPercentageFile.arrayBuffer();
      const actualPercentageWorkbook = XLSX.read(actualPercentageBuffer);
      const actualPercentageSheet =
        actualPercentageWorkbook.Sheets[actualPercentageWorkbook.SheetNames[0]];
      const actualPercentageData: any[][] = XLSX.utils.sheet_to_json(
        actualPercentageSheet,
        { header: 1 }
      );

      const employeePercentageMap = new Map<number, number>();
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
              percentage &&
              !isNaN(percentage)
            ) {
              employeePercentageMap.set(empCode, percentage);
              console.log(
                `✅ Employee ${empCode}: Custom percentage ${percentage}%`
              );
            }
          }
        }
      }

      console.log(
        `📋 Loaded custom percentages for ${employeePercentageMap.size} employees`
      );

      // ========== LOAD BONUS FILE WITH ACCUMULATION FOR DUPLICATES ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const hrRegisterData: Map<
        number,
        { registerHR: number; dept: string; occurrences: number }
      > = new Map();

      // Process Worker sheet (1st sheet) - Register in column S (index 18)
      if (bonusWorkbook.SheetNames.length > 0) {
        const workerSheetName = bonusWorkbook.SheetNames[0];
        console.log(`📄 Processing Worker sheet: ${workerSheetName}`);
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
          const registerIdx = 18; // Column S (0-indexed)

          for (let i = workerHeaderRow + 1; i < workerData.length; i++) {
            const row = workerData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrRegisterData.has(empCode)) {
                const existing = hrRegisterData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.occurrences += 1;
                console.log(
                  `🔄 Worker Emp ${empCode}: Duplicate found - Adding ₹${registerHR.toFixed(
                    2
                  )}, Total now: ₹${existing.registerHR.toFixed(2)} (${
                    existing.occurrences
                  } occurrences)`
                );
              } else {
                hrRegisterData.set(empCode, {
                  registerHR: registerHR,
                  dept: "Worker",
                  occurrences: 1,
                });
                console.log(
                  `✅ Worker Emp ${empCode}: First entry - Register: ₹${registerHR.toFixed(
                    2
                  )}`
                );
              }
            }
          }
        }
      }

      // Process Staff sheet (2nd sheet) - Register in column T (index 19)
      if (bonusWorkbook.SheetNames.length > 1) {
        const staffSheetName = bonusWorkbook.SheetNames[1];
        console.log(`📄 Processing Staff sheet: ${staffSheetName}`);
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
          const registerIdx = 19; // Column T (0-indexed)

          for (let i = staffHeaderRow + 1; i < staffData.length; i++) {
            const row = staffData[i];
            if (!row || row.length === 0) continue;

            const empCode = Number(row[empCodeIdx]);
            const registerHR = Number(row[registerIdx]) || 0;

            if (empCode && !isNaN(empCode)) {
              if (hrRegisterData.has(empCode)) {
                const existing = hrRegisterData.get(empCode)!;
                existing.registerHR += registerHR;
                existing.occurrences += 1;
                console.log(
                  `🔄 Staff Emp ${empCode}: Duplicate found - Adding ₹${registerHR.toFixed(
                    2
                  )}, Total now: ₹${existing.registerHR.toFixed(2)} (${
                    existing.occurrences
                  } occurrences)`
                );
              } else {
                hrRegisterData.set(empCode, {
                  registerHR: registerHR,
                  dept: "Staff",
                  occurrences: 1,
                });
                console.log(
                  `✅ Staff Emp ${empCode}: First entry - Register: ₹${registerHR.toFixed(
                    2
                  )}`
                );
              }
            }
          }
        }
      }

      console.log(
        `✅ HR Register data loaded: ${hrRegisterData.size} employees`
      );

      const duplicateEmployees = Array.from(hrRegisterData.entries()).filter(
        ([_, data]) => data.occurrences > 1
      );

      if (duplicateEmployees.length > 0) {
        console.log("\n🔍 EMPLOYEES WITH MULTIPLE ENTRIES (SUMMED):");
        console.log("=".repeat(60));
        duplicateEmployees.forEach(([empId, data]) => {
          console.log(
            `Emp ${empId}: ${
              data.occurrences
            } entries, Total Register (HR): ₹${data.registerHR.toFixed(2)}`
          );
        });
      }

      // ========== COMPUTE GROSS SALARY (EXACT STEP-3 LOGIC) ==========

      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);

      const staffEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
      > = new Map();

      for (const sheetName of staffWorkbook.SheetNames) {
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

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) {
          console.log(`⚠️ Skip Staff ${sheetName}: missing columns`);
          continue;
        }

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "")
            .trim()
            .toUpperCase();
          const salary1Raw = row[salary1Idx];

          const salary1 =
            salary1Raw === null || salary1Raw === undefined || salary1Raw === ""
              ? 0
              : Number(salary1Raw) || 0;

          if (!empId || isNaN(empId) || !empName) continue;

          if (!staffEmployees.has(empId)) {
            staffEmployees.set(empId, {
              name: empName,
              dept: "Staff",
              months: new Map(),
            });
          }

          const emp = staffEmployees.get(empId)!;

          const startMonth = EMPLOYEE_START_MONTHS[empId];
          if (startMonth && monthKey < startMonth) {
            continue;
          }

          if (
            INCLUDE_ZEROS_IN_AVG.has(empId) ||
            EXCLUDE_ZEROS_IN_AVG.has(empId)
          ) {
            emp.months.set(monthKey, salary1);
          } else {
            if (salary1 > 0) {
              emp.months.set(monthKey, salary1);
            }
          }
        }
      }

      console.log(`✅ Staff employees: ${staffEmployees.size}`);

      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);

      const workerEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
      > = new Map();

      for (const sheetName of workerWorkbook.SheetNames) {
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

        if (headerIdx === -1) {
          console.log(`⚠️ Worker ${sheetName}: No header`);
          continue;
        }

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

        const salary1Idx = 8; // Column I

        if (empIdIdx === -1 || empNameIdx === -1) {
          console.log(`⚠️ Skip Worker ${sheetName}: missing columns`);
          continue;
        }

        if (deptIdx === -1) {
          console.log(
            `⚠️ Warning: Department column not found in ${sheetName}`
          );
        }

        let excludedDeptCount = 0;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "")
            .trim()
            .toUpperCase();
          const salary1Raw = row[salary1Idx];

          const salary1 =
            salary1Raw === null || salary1Raw === undefined || salary1Raw === ""
              ? 0
              : Number(salary1Raw) || 0;

          if (deptIdx !== -1) {
            const dept = String(row[deptIdx] || "")
              .trim()
              .toUpperCase();

            if (EXCLUDED_DEPARTMENTS.includes(dept)) {
              excludedDeptCount++;
              continue;
            }
          }

          if (!empId || isNaN(empId) || !empName) continue;

          if (!workerEmployees.has(empId)) {
            workerEmployees.set(empId, {
              name: empName,
              dept: "Worker",
              months: new Map(),
            });
          }

          const emp = workerEmployees.get(empId)!;

          const startMonth = EMPLOYEE_START_MONTHS[empId];
          if (startMonth && monthKey < startMonth) {
            continue;
          }

          if (
            INCLUDE_ZEROS_IN_AVG.has(empId) ||
            EXCLUDE_ZEROS_IN_AVG.has(empId)
          ) {
            emp.months.set(monthKey, salary1);
          } else {
            if (salary1 > 0) {
              emp.months.set(monthKey, salary1);
            }
          }
        }

        if (excludedDeptCount > 0) {
          console.log(
            `💰 Filtered ${excludedDeptCount} workers from ${sheetName}`
          );
        }
      }

      console.log(`✅ Worker employees: ${workerEmployees.size}`);

      // ========== COMPUTE SOFTWARE TOTALS WITH OCTOBER ESTIMATION (STEP-3 LOGIC) ==========
      const grossSalaryData: Map<
        number,
        { name: string; dept: string; grossSalary: number }
      > = new Map();

      const foldMonthly = (
        src: Map<
          number,
          { name: string; dept: string; months: Map<string, number> }
        >
      ) => {
        for (const [empId, rec] of src) {
          let baseSum = 0;
          const monthsIncluded: { month: string; value: number }[] = [];

          const includeZeros = INCLUDE_ZEROS_IN_AVG.has(empId);
          const excludeZeros = EXCLUDE_ZEROS_IN_AVG.has(empId);
          const hasCustomStart = EMPLOYEE_START_MONTHS[empId] !== undefined;

          const employeeWindow = hasCustomStart
            ? AVG_WINDOW.filter((mk) => mk >= EMPLOYEE_START_MONTHS[empId])
            : AVG_WINDOW;

          for (const mk of employeeWindow) {
            const v = rec.months.get(mk);

            if (includeZeros) {
              const val = v !== undefined ? Number(v) : 0;
              baseSum += val;
              monthsIncluded.push({ month: mk, value: val });
            } else if (excludeZeros) {
              if (v !== undefined && v !== null) {
                const val = Number(v);
                baseSum += val;
                monthsIncluded.push({ month: mk, value: val });
              }
            } else {
              if (v != null && !isNaN(Number(v)) && Number(v) > 0) {
                baseSum += Number(v);
                monthsIncluded.push({ month: mk, value: Number(v) });
              }
            }
          }

          let estOct = 0;
          let total = baseSum;
          const hasSep2025 =
            rec.months.has("2025-09") && (rec.months.get("2025-09") || 0) > 0;
          const isExcluded = EXCLUDE_OCTOBER_EMPLOYEES.has(empId);

          if (isExcluded) {
            console.log(
              `🚫 EMP ${empId} (${
                rec.name
              }): IN EXCLUDE LIST - Base only = ₹${baseSum.toFixed(2)}`
            );
          } else if (hasSep2025 && monthsIncluded.length > 0) {
            if (includeZeros) {
              const divisor = hasCustomStart ? employeeWindow.length : 11;
              estOct = baseSum / divisor;
              total = baseSum + estOct;
            } else {
              const values = monthsIncluded.map((m) => m.value);
              estOct = values.reduce((a, b) => a + b, 0) / values.length;
              total = baseSum + estOct;
            }
          }

          if (!grossSalaryData.has(empId)) {
            grossSalaryData.set(empId, {
              name: rec.name,
              dept: rec.dept,
              grossSalary: total,
            });
          } else {
            grossSalaryData.get(empId)!.grossSalary += total;
          }
        }
      };

      foldMonthly(staffEmployees);
      foldMonthly(workerEmployees);

      console.log(
        `✅ Gross Salary computed (Step-3 logic): ${grossSalaryData.size} employees`
      );

      // ========== CALCULATE REGISTER WITH VARIABLE PERCENTAGES ==========
      const comparison: any[] = [];

      for (const [empId, empData] of grossSalaryData) {
        const employeePercentage =
          employeePercentageMap.get(empId) || DEFAULT_PERCENTAGE;

        const shouldApplyGrossAdjustment =
          employeePercentage === SPECIAL_PERCENTAGE;

        let registerSoftware: number;
        let adjustedGross: number;

        if (shouldApplyGrossAdjustment) {
          adjustedGross = empData.grossSalary * SPECIAL_GROSS_MULTIPLIER;
          registerSoftware = (adjustedGross * employeePercentage) / 100;

          console.log(
            `🎯 Emp ${empId}: ${employeePercentage}% with 60% adjustment | Gross=₹${empData.grossSalary.toFixed(
              2
            )} → 60%=₹${adjustedGross.toFixed(
              2
            )} → ${employeePercentage}%=₹${registerSoftware.toFixed(2)}`
          );
        } else {
          adjustedGross = empData.grossSalary;
          registerSoftware = (empData.grossSalary * employeePercentage) / 100;

          console.log(
            `📊 Emp ${empId}: ${employeePercentage}% on full gross | Gross=₹${empData.grossSalary.toFixed(
              2
            )} × ${employeePercentage}%=₹${registerSoftware.toFixed(2)}`
          );
        }

        const hrData = hrRegisterData.get(empId);
        const registerHR = hrData?.registerHR || 0;
        const occurrences = hrData?.occurrences || 0;

        const difference = registerSoftware - registerHR;
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        comparison.push({
          employeeId: empId,
          employeeName: empData.name,
          department: empData.dept,
          grossSalarySoftware: empData.grossSalary,
          adjustedGross: adjustedGross,
          percentage: employeePercentage,
          appliedGrossAdjustment: shouldApplyGrossAdjustment,
          registerSoftware: registerSoftware,
          registerHR: registerHR,
          hrOccurrences: occurrences,
          difference: difference,
          status: status,
        });

        if (occurrences > 1) {
          console.log(
            `🔄 Emp ${empId}: ${occurrences} HR entries summed to ₹${registerHR.toFixed(
              2
            )}`
          );
        }
      }

      comparison.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(comparison);
      setFilteredData(comparison);

      console.log(
        "✅ Register calculation completed using Step-3 gross values with 60% rule for 12% employees"
      );
    } catch (err: any) {
      setError(`Error processing files: ${err.message}`);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  useEffect(() => {
    if (staffFile && workerFile && bonusFile && actualPercentageFile) {
      processFiles();
    }
  }, [staffFile, workerFile, bonusFile, actualPercentageFile]);

  useEffect(() => {
    if (departmentFilter === "All") {
      setFilteredData(comparisonData);
    } else {
      setFilteredData(
        comparisonData.filter((row) => row.department === departmentFilter)
      );
    }
  }, [departmentFilter, comparisonData]);

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      maximumFractionDigits: 2,
    }).format(value);
  };

  const exportToExcel = () => {
    const dataToExport =
      departmentFilter === "All" ? comparisonData : filteredData;

    const ws = XLSX.utils.json_to_sheet(
      dataToExport.map((row) => ({
        "Employee ID": row.employeeId,
        "Employee Name": row.employeeName,
        Department: row.department,
        "Gross Salary (Software)": row.grossSalarySoftware,
        "Percentage (%)": row.percentage,
        "Register (Software)": row.registerSoftware,
        "Register (HR)": row.registerHR,
        "HR Occurrences": row.hrOccurrences,
        Difference: row.difference,
        Status: row.status,
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Register Calculation");
    XLSX.writeFile(wb, `Step5-Register-Calculation-${departmentFilter}.xlsx`);
  };

  // ========== Calculate Grand Totals ==========
  const calculateGrandTotals = () => {
    const totalRegisterSoftware = filteredData.reduce(
      (sum, row) => sum + (Number(row.registerSoftware) || 0),
      0
    );
    const totalRegisterHR = filteredData.reduce(
      (sum, row) => sum + (Number(row.registerHR) || 0),
      0
    );
    const totalDifference = totalRegisterSoftware - totalRegisterHR;

    return {
      totalRegisterSoftware,
      totalRegisterHR,
      totalDifference,
    };
  };

  const grandTotals = calculateGrandTotals();

  const FileCard = ({
    title,
    file,
    description,
    icon,
  }: {
    title: string;
    file: File | null;
    description: string;
    icon: React.ReactNode;
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
    <div className="min-h-screen bg-gradient-to-br from-purple-50 to-pink-100 py-5 px-4">
      <div className="mx-auto max-w-full">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 5 - Register Calculation (Using Step-3 Gross)
              </h1>
              <p className="text-gray-600 mt-2">
                12% employees: (Step3-Gross × 60%) × 12% | 8.33% employees:
                Step3-Gross × 8.33%
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("/step4")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                ← Back to Step 4
              </button>
              <button
                onClick={() => router.push("/")}
                className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Back to Step 1
              </button>

              <button
                onClick={() => router.push("/step6")}
                className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Move to Step 6
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data (Nov-24 to Sep-25)"
              icon={<></>}
            />
            <FileCard
              title="Indiana Worker"
              file={workerFile}
              description="Worker salary data (excludes Dept C, CASH, A)"
              icon={<></>}
            />
            <FileCard
              title="Bonus Calculation Sheet"
              file={bonusFile}
              description="HR Register values - Now sums duplicates!"
              icon={<></>}
            />
            <FileCard
              title="Actual Percentage Data"
              file={actualPercentageFile}
              description="Employees with 12% bonus (60% rule applies)"
              icon={<></>}
            />
          </div>

          {[staffFile, workerFile, bonusFile, actualPercentageFile].filter(
            Boolean
          ).length < 4 && (
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
                  Calculating with Step-3 gross and 60% rule for 12%
                  employees...
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
                    Register Comparison
                  </h2>
                  <select
                    value={departmentFilter}
                    onChange={(e) => setDepartmentFilter(e.target.value)}
                    className="px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                  >
                    <option value="All">All Departments</option>
                    <option value="Staff">Staff Only</option>
                    <option value="Worker">Worker Only</option>
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

              <div className="border border-gray-300 rounded-lg overflow-hidden">
                <div className="overflow-x-auto">
                  <div className="max-h-[600px] overflow-y-auto">
                    <table className="w-full border-collapse">
                      <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                          <th className="border border-gray-300 px-4 py-3 text-left bg-gray-100">
                            <div className="flex items-center">
                              Employee ID
                              <SortArrows columnKey="employeeId" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-left bg-gray-100">
                            <div className="flex items-center">
                              Employee Name
                              <SortArrows columnKey="employeeName" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-left bg-gray-100">
                            <div className="flex items-center">
                              Dept.
                              <SortArrows columnKey="department" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Gross (Software)
                              <SortArrows columnKey="grossSalarySoftware" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Adj. Gross
                              <SortArrows columnKey="adjustedGross" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-center bg-gray-100">
                            <div className="flex items-center justify-center">
                              %<SortArrows columnKey="percentage" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Register(Software)
                              <SortArrows columnKey="registerSoftware" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Register(HR)
                              <SortArrows columnKey="registerHR" />
                            </div>
                          </th>

                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Difference
                              <SortArrows columnKey="difference" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-center bg-gray-100">
                            <div className="flex items-center justify-center">
                              Status
                              <SortArrows columnKey="status" />
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
                            } ${row.hrOccurrences > 1 ? "bg-yellow-50" : ""}`}
                          >
                            <td className="border border-gray-300 px-4 py-2">
                              {row.employeeId}
                            </td>
                            <td className="border border-gray-300 px-4 py-2">
                              {row.employeeName}
                            </td>
                            <td className="border border-gray-300 px-4 py-2">
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
                            <td className="border border-gray-300 px-4 py-2 text-right">
                              {formatCurrency(row.grossSalarySoftware)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right">
                              {formatCurrency(row.adjustedGross)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-center">
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
                            <td className="border border-gray-300 px-4 py-2 text-right">
                              {formatCurrency(row.registerSoftware)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right">
                              {formatCurrency(row.registerHR)}
                            </td>

                            <td
                              className={`border border-gray-300 px-4 py-2 text-right font-medium ${
                                Math.abs(row.difference) <= 12
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
                            {formatCurrency(grandTotals.totalRegisterSoftware)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-purple-900">
                            {formatCurrency(grandTotals.totalRegisterHR)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-green-700">
                            {formatCurrency(grandTotals.totalDifference)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-center">
                            <span className="px-3 py-1 rounded-full text-sm font-medium bg-green-200 text-green-900">
                              Match
                            </span>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
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
                  {filteredData.filter((r) => r.status === "Mismatch").length} |
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
