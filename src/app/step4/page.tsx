"use client";

import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

export default function Step4Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

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
        /bonus.*final.*calculation|bonus.*2024-25/i.test(s.file.name)
    );

  const actualPercentageFile =
    pickFile((s) => s.type === "Actual-Percentage-Bonus-Data") ??
    pickFile((s) => !!s.file && /actual.*percentage/i.test(s.file.name));

  // Helper to normalize header text
  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  // Month parsing constants
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
    const s = String(sheetName ?? "")
      .trim()
      .toUpperCase();

    // Case 1: YYYY-MM
    const yyyymm = s.match(/(\d{4})-(\d{1,2})/);
    if (yyyymm) {
      const y = Number(yyyymm[1]);
      const m = Number(yyyymm[2]);
      if (y >= 2000 && m >= 1 && m <= 12) {
        return `${y}-${pad2(m)}`;
      }
    }

    // Case 2: MON or MONTH with 2/4 digit year nearby
    const mon = s.match(
      /(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)/
    );
    const monthFull = s.match(
      /(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)/
    );
    const y2or4 = s.match(/(\d{2,4})/);
    const monthToken = (monthFull?.[1] ?? mon?.[1]) as string | undefined;

    if (monthToken && y2or4) {
      let y = Number(y2or4[1]);
      if (y < 100) y += 2000;
      const m = MONTH_NAME_MAP[monthToken];
      if (m) return `${y}-${pad2(m)}`;
    }

    return null;
  };

  // Months to average for October 2025 estimate
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

  // === Step 4 Audit Helpers ===
  const TOLERANCE = 12; // Step 4 uses ±12 to mark Match vs Mismatch

  async function postAuditMessagesStep4(items: any[], batchId?: string) {
    const bid =
      batchId ||
      (typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2));
    await fetch("/api/audit/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId: bid, step: 4, items }),
    });
    return bid;
  }

  function buildStep4MismatchMessages(rows: any[]) {
    const items: any[] = [];
    for (const r of rows) {
      if (r?.status === "Mismatch") {
        items.push({
          level: "error",
          tag: "mismatch",
          text: `[step4] ${r.employeeId} ${r.employeeName} diff=${Number(
            r.difference ?? 0
          ).toFixed(2)}`,
          scope: "staff",
          source: "step4",
          meta: {
            employeeId: r.employeeId,
            name: r.employeeName,
            department: r.department,
            percentage: r.percentage,
            grossSal: r.grossSal,
            calculatedValue: r.calculatedValue,
            gross2HR: r.gross2HR,
            diff: r.difference,
            tolerance: TOLERANCE,
          },
        });
      }
    }
    return items;
  }

  function buildStep4SummaryMessage(rows: any[]) {
    const total = rows.length || 0;
    const mismatches = rows.filter((r) => r.status === "Mismatch").length;
    const matches = total - mismatches;

    const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
    const grossSalSum = sum(rows.map((r) => Number(r.grossSal || 0)));
    const calcSum = sum(rows.map((r) => Number(r.calculatedValue || 0)));
    const hrSum = sum(rows.map((r) => Number(r.gross2HR || 0)));

    return {
      level: "info",
      tag: "summary",
      text: `Step4 run: total=${total} match=${matches} mismatch=${mismatches}`,
      scope: "staff",
      source: "step4",
      meta: {
        totals: {
          total,
          matches,
          mismatches,
          tolerance: TOLERANCE,
          grossSalSum,
          calculatedSum: calcSum,
          hrGross2Sum: hrSum,
        },
      },
    };
  }

  async function handleSaveAuditStep4(rows: any[]) {
    if (!rows || rows.length === 0) return;
    const items = [
      buildStep4SummaryMessage(rows),
      ...buildStep4MismatchMessages(rows),
    ];
    if (items.length === 0) return;
    await postAuditMessagesStep4(items);
  }

  // Stable hash for a run signature
  function djb2Hash(str: string) {
    let h = 5381;
    for (let i = 0; i < str.length; i++) h = (h << 5) + h + str.charCodeAt(i);
    return (h >>> 0).toString(36);
  }

  function buildRunKeyStep4(rows: any[]) {
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
    if (typeof window === "undefined") return;
    if (!Array.isArray(comparisonData) || comparisonData.length === 0) return;

    const runKey = buildRunKeyStep4(comparisonData);
    const markerKey = `audit_step4_${runKey}`;

    if (sessionStorage.getItem(markerKey)) return;

    sessionStorage.setItem(markerKey, "1");
    const deterministicBatchId = `step4-${runKey}`;

    const items = [
      buildStep4SummaryMessage(comparisonData),
      ...buildStep4MismatchMessages(comparisonData),
    ];

    postAuditMessagesStep4(items, deterministicBatchId).catch((err) => {
      console.error("Auto-audit step4 failed", err);
      sessionStorage.removeItem(markerKey);
    });
  }, [comparisonData]);

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
    let dataToSort = comparisonData;
    if (sortConfig.key && sortConfig.direction) {
      dataToSort = [...dataToSort].sort((a, b) => {
        let aValue = a[sortConfig.key],
          bValue = b[sortConfig.key];
        if (
          [
            "employeeId",
            "percentage",
            "grossSal",
            "calculatedValue",
            "gross2HR",
            "difference",
          ].includes(sortConfig.key)
        ) {
          aValue = Number(aValue) || 0;
          bValue = Number(bValue) || 0;
        } else if (
          sortConfig.key === "employeeName" ||
          sortConfig.key === "department" ||
          sortConfig.key === "status"
        ) {
          aValue = String(aValue || "").toUpperCase();
          bValue = String(bValue || "").toUpperCase();
        }
        if (aValue < bValue) return sortConfig.direction === "asc" ? -1 : 1;
        if (aValue > bValue) return sortConfig.direction === "asc" ? 1 : -1;
        return 0;
      });
    }
    setFilteredData(dataToSort);
  }, [sortConfig, comparisonData]);

  const SortArrows = ({ columnKey }: { columnKey: string }) => {
    const isActive = sortConfig.key === columnKey;
    return (
      <div className="inline-flex flex-col ml-1">
        <button
          type="button"
          className={`leading-none ${
            isActive && sortConfig.direction === "asc"
              ? "text-indigo-600"
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
              ? "text-indigo-600"
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

  const calculatePercentage = (dateOfJoining: any): number => {
    if (!dateOfJoining) return 0;

    let doj: Date;

    if (typeof dateOfJoining === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      doj = new Date(excelEpoch.getTime() + dateOfJoining * 86400000);
    } else {
      doj = new Date(dateOfJoining);
    }

    if (isNaN(doj.getTime())) return 0;

    const referenceDate = new Date(2025, 8, 30);

    const yearsDiff = referenceDate.getFullYear() - doj.getFullYear();
    const monthsDiff = referenceDate.getMonth() - doj.getMonth();
    const daysDiff = referenceDate.getDate() - doj.getDate();

    let totalMonths = yearsDiff * 12 + monthsDiff;

    if (daysDiff < 0) {
      totalMonths--;
    }

    console.log(
      `DOJ: ${doj.toLocaleDateString()}, Reference: ${referenceDate.toLocaleDateString()}, Total Months: ${totalMonths}`
    );

    if (totalMonths < 12) {
      return 10;
    } else if (totalMonths >= 12 && totalMonths < 24) {
      return 12;
    } else {
      return 8.33;
    }
  };

  const applyBonusFormula = (grossSal: number, percentage: number): number => {
    if (percentage === 8.33) {
      return grossSal;
    } else if (percentage > 8.33) {
      return grossSal * 0.6;
    } else {
      return 0;
    }
  };

  const processFiles = async () => {
    if (!staffFile || !bonusFile) {
      setError("Both Indiana Staff and Bonus Sheet files are required");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      let employeesToIgnoreOctober = new Set<number>();

      if (actualPercentageFile) {
        try {
          console.log("Processing Actual Percentage file...");
          const actualBuffer = await actualPercentageFile.arrayBuffer();
          const actualWorkbook = XLSX.read(actualBuffer);

          const averageSheetName = actualWorkbook.SheetNames.find(
            (name) => name.toLowerCase() === "average"
          );

          if (averageSheetName) {
            console.log(`Found Average sheet: ${averageSheetName}`);
            const averageSheet = actualWorkbook.Sheets[averageSheetName];
            const averageData: any[][] = XLSX.utils.sheet_to_json(
              averageSheet,
              {
                header: 1,
              }
            );

            let avgHeaderRow = -1;
            for (let i = 0; i < Math.min(10, averageData.length); i++) {
              const row = averageData[i];
              if (
                row &&
                row.some((cell: any) => {
                  const cellStr = String(cell || "")
                    .toUpperCase()
                    .replace(/\s+/g, "");
                  return cellStr.includes("EMP") && cellStr.includes("CODE");
                })
              ) {
                avgHeaderRow = i;
                break;
              }
            }

            if (avgHeaderRow !== -1) {
              const avgHeaders = averageData[avgHeaderRow];
              const avgEmpCodeIdx = avgHeaders.findIndex((h: any) => {
                const hStr = String(h || "")
                  .toUpperCase()
                  .replace(/\s+/g, "");
                return (
                  hStr.includes("EMP") &&
                  (hStr.includes("CODE") || hStr === "EMPCODE")
                );
              });

              if (avgEmpCodeIdx !== -1) {
                for (let i = avgHeaderRow + 1; i < averageData.length; i++) {
                  const row = averageData[i];
                  if (!row || row.length === 0) continue;

                  const empId = Number(row[avgEmpCodeIdx]);
                  if (empId && !isNaN(empId)) {
                    employeesToIgnoreOctober.add(empId);
                  }
                }

                console.log(
                  `✅ Found ${employeesToIgnoreOctober.size} employees in Average sheet - will set October salary to 0 for these employees`
                );
                console.log(
                  "Employee IDs to ignore October estimate:",
                  Array.from(employeesToIgnoreOctober)
                );
              }
            }
          } else {
            console.log("Average sheet not found in Actual Percentage file");
          }
        } catch (err: any) {
          console.error(
            "Error processing Actual Percentage file:",
            err.message
          );
        }
      } else {
        console.log(
          "Actual Percentage file not found - proceeding without October filtering"
        );
      }

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
        const sheet = staffWorkbook.Sheets[sheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";

        console.log(`Staff sheet: ${sheetName} → monthKey: ${monthKey}`);

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

        console.log(`📋 Sheet: ${sheetName}`);
        console.log(`  Headers found at row ${headerIdx}`);
        console.log(`  EMP ID column index: ${empIdIdx}`);
        console.log(`  Name column index: ${empNameIdx}`);
        console.log(`  SALARY1 column index: ${salary1Idx}`);
        console.log(`  DOJ column index: ${dojIdx}`);
        if (dojIdx !== -1) {
          console.log(`  DOJ column header: "${headers[dojIdx]}"`);
        } else {
          console.warn(`  ⚠️ DOJ column NOT FOUND in sheet ${sheetName}`);
        }

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) {
          console.log(
            `Skipping staff sheet ${sheetName}: missing required columns`
          );
          continue;
        }

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "")
            .trim()
            .toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;
          const doj = dojIdx !== -1 ? row[dojIdx] : null;

          if (!empId || isNaN(empId) || !empName) continue;

          if (empId === 554) {
            console.log(`🔍 FOUND EMPLOYEE 554 in sheet ${sheetName}:`);
            console.log(`  Name: ${empName}`);
            console.log(`  DOJ raw value: ${doj}`);
            console.log(`  DOJ type: ${typeof doj}`);
            console.log(`  DOJ index: ${dojIdx}`);
          }

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
              console.log(
                `Updated DOJ for employee ${empId} from sheet ${sheetName}`
              );
            }
          }

          const emp = staffEmployees.get(empId)!;
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
        }
      }

      console.log(`Total staff employees: ${staffEmployees.size}`);

      const softwareEmployeesTotals: Map<
        number,
        {
          name: string;
          dept: string;
          dateOfJoining: any;
          grossSal: number;
        }
      > = new Map();

      for (const [empId, rec] of staffEmployees) {
        let baseSum = 0;
        for (const v of rec.months.values()) {
          baseSum += Number(v) || 0;
        }

        let estOct = 0;
        const hasSep2025 = rec.months.has("2025-09");
        const isInAverageSheet = employeesToIgnoreOctober.has(empId);

        if (hasSep2025 && !isInAverageSheet) {
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

          console.log(
            `Employee ${empId} (${
              rec.name
            }): Has Sep 2025 data, NOT in Average sheet, Base sum = ${baseSum}, Oct estimate = ${estOct}, Total = ${
              baseSum + estOct
            }`
          );
        } else if (isInAverageSheet) {
          console.log(
            `🚫 Employee ${empId} (${rec.name}): Found in Average sheet - IGNORING October estimate (set to 0), Total = ${baseSum}`
          );
        } else {
          console.log(
            `Employee ${empId} (${rec.name}): Missing Sep 2025 data, skipping Oct estimate, Total = ${baseSum}`
          );
        }

        const total = baseSum + estOct;

        softwareEmployeesTotals.set(empId, {
          name: rec.name,
          dept: rec.dept,
          dateOfJoining: rec.dateOfJoining,
          grossSal: total,
        });
      }

      console.log(
        `Total employees with GROSS SAL. (Software): ${softwareEmployeesTotals.size}`
      );

      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const bonusGross2Map: Map<number, number> = new Map();
      const bonusEmployeeNames: Map<number, string> = new Map();
      const bonusEmployeeDepts: Map<number, string> = new Map();

      const staffSheetName = bonusWorkbook.SheetNames.find(
        (name) => name.toLowerCase() === "staff"
      );

      if (!staffSheetName) {
        console.error("Staff sheet not found in bonus file!");
        setError("Staff sheet not found in Bonus Calculation Sheet");
        setIsProcessing(false);
        return;
      }

      console.log(`Processing bonus sheet: ${staffSheetName}`);

      const bonusSheet = bonusWorkbook.Sheets[staffSheetName];
      const bonusSheetData: any[][] = XLSX.utils.sheet_to_json(bonusSheet, {
        header: 1,
      });

      let bonusHeaderRow = -1;
      for (let i = 0; i < Math.min(5, bonusSheetData.length); i++) {
        const row = bonusSheetData[i];
        if (
          row &&
          row.some((cell: any) => {
            const cellStr = String(cell || "").toUpperCase();
            return cellStr.includes("EMP") && cellStr.includes("CODE");
          })
        ) {
          bonusHeaderRow = i;
          break;
        }
      }

      if (bonusHeaderRow === -1) {
        console.error(`Cannot locate header row in Staff sheet`);
        setError("Cannot locate header row in Staff sheet of Bonus file");
        setIsProcessing(false);
        return;
      }

      const headers = bonusSheetData[bonusHeaderRow];

      const empCodeIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("EMP") && hStr.includes("CODE");
      });

      const empNameIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("EMP") && hStr.includes("NAME");
      });

      const deptIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("DEPT") || hStr === "DEPTT." || hStr === "DEPTT";
      });

      const gross2Idx = headers.findIndex((h: any) => {
        const hStr = String(h || "")
          .trim()
          .toUpperCase();
        return (
          hStr === "GROSS 02" ||
          hStr === "GROSS02" ||
          (hStr.includes("GROSS") &&
            (hStr.includes("02") || hStr.includes(" 2")))
        );
      });

      if (empCodeIdx === -1) {
        console.error(`Cannot locate EMP Code column in Staff sheet`);
        setError("Cannot locate EMP Code column in Staff sheet of Bonus file");
        setIsProcessing(false);
        return;
      }

      if (gross2Idx === -1) {
        console.error(`Cannot locate GROSS 02 column in Staff sheet`);
        setError("Cannot locate GROSS 02 column in Staff sheet of Bonus file");
        setIsProcessing(false);
        return;
      }

      console.log(
        `Staff sheet: Found EmpCode at index ${empCodeIdx}, Department at index ${deptIdx}, GROSS 02 at index ${gross2Idx}`
      );

      let processedCount = 0;
      const duplicateTracker: Map<number, number> = new Map();

      for (let i = bonusHeaderRow + 1; i < bonusSheetData.length; i++) {
        const row = bonusSheetData[i];
        if (!row || row.length === 0) continue;

        const empId = Number(row[empCodeIdx]);
        const gross2 = Number(row[gross2Idx]) || 0;
        const empName = empNameIdx !== -1 ? String(row[empNameIdx] || "") : "";
        const dept = deptIdx !== -1 ? String(row[deptIdx] || "") : "Staff";

        if (!empId || isNaN(empId)) continue;

        if (bonusGross2Map.has(empId)) {
          const existingSum = bonusGross2Map.get(empId)!;
          const newSum = existingSum + gross2;
          bonusGross2Map.set(empId, newSum);

          const occurrenceCount = (duplicateTracker.get(empId) || 1) + 1;
          duplicateTracker.set(empId, occurrenceCount);

          console.log(
            `Duplicate found - Employee ${empId} (${empName}): Adding GROSS 02 = ${gross2} to existing ${existingSum} = ${newSum} (occurrence #${occurrenceCount})`
          );
        } else {
          bonusGross2Map.set(empId, gross2);
          bonusEmployeeNames.set(empId, empName);
          bonusEmployeeDepts.set(empId, dept);
          duplicateTracker.set(empId, 1);
          processedCount++;
          console.log(
            `New employee ${empId} (${empName}, Dept: ${dept}): GROSS 02 = ${gross2}`
          );
        }
      }

      console.log(`Staff sheet: Processed ${processedCount} unique employees`);
      console.log(
        `Total unique employees in bonus Staff sheet: ${bonusGross2Map.size}`
      );

      const employeesWithDuplicates = Array.from(
        duplicateTracker.entries()
      ).filter(([_, count]) => count > 1);

      if (employeesWithDuplicates.length > 0) {
        console.log(
          `\n${employeesWithDuplicates.length} employees with multiple entries:`
        );
        employeesWithDuplicates.forEach(([empId, count]) => {
          const totalGross2 = bonusGross2Map.get(empId) || 0;
          const name = bonusEmployeeNames.get(empId) || "Unknown";
          const dept = bonusEmployeeDepts.get(empId) || "Unknown";
          console.log(
            `  Employee ${empId} (${name}, ${dept}): ${count} entries, Total GROSS 02 = ${totalGross2}`
          );
        });
      }

      console.log(
        `\nSample GROSS 02 values from HR (Staff - ALL occurrences summed):`
      );
      let sampleCount = 0;
      for (const [empId, gross2] of bonusGross2Map) {
        if (sampleCount < 5) {
          const occurrences = duplicateTracker.get(empId) || 1;
          const name = bonusEmployeeNames.get(empId) || "Unknown";
          const dept = bonusEmployeeDepts.get(empId) || "Unknown";
          console.log(
            `  EMP ${empId} (${name}, ${dept}): ${gross2} (from ${occurrences} occurrence${
              occurrences > 1 ? "s" : ""
            })`
          );
          sampleCount++;
        }
      }

      const calculationResults: any[] = [];

      for (const [empId, empData] of softwareEmployeesTotals) {
        const percentage = calculatePercentage(empData.dateOfJoining);
        const gross2HR = bonusGross2Map.get(empId) || 0;
        const department = bonusEmployeeDepts.get(empId) || empData.dept;

        const calculatedValue = applyBonusFormula(empData.grossSal, percentage);

        const difference = calculatedValue - gross2HR;
        const status = Math.abs(difference) <= 12 ? "Match" : "Mismatch";

        calculationResults.push({
          employeeId: empId,
          employeeName: empData.name,
          department: department,
          dateOfJoining: empData.dateOfJoining,
          percentage: percentage,
          grossSal: empData.grossSal,
          calculatedValue: calculatedValue,
          gross2HR: gross2HR,
          difference: difference,
          status: status,
        });
      }

      for (const [empId, gross2] of bonusGross2Map) {
        if (!softwareEmployeesTotals.has(empId)) {
          const name =
            bonusEmployeeNames.get(empId) || "NOT FOUND IN STAFF FILE";
          const dept = bonusEmployeeDepts.get(empId) || "Unknown";
          calculationResults.push({
            employeeId: empId,
            employeeName: name,
            department: dept,
            dateOfJoining: null,
            percentage: 0,
            grossSal: 0,
            calculatedValue: 0,
            gross2HR: gross2,
            difference: -gross2,
            status: "Mismatch",
          });
        }
      }

      calculationResults.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(calculationResults);
      setFilteredData(calculationResults);
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
  }, [staffFile, bonusFile]);

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: "INR",
      maximumFractionDigits: 2,
    }).format(value);
  };

  const formatDate = (dateValue: any) => {
    if (!dateValue) return "N/A";

    try {
      let date: Date;
      if (typeof dateValue === "number") {
        const excelEpoch = new Date(1899, 11, 30);
        date = new Date(excelEpoch.getTime() + dateValue * 86400000);
      } else if (typeof dateValue === "string") {
        const ddmmyyMatch = dateValue.match(
          /^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$/
        );
        if (ddmmyyMatch) {
          let [, day, month, year] = ddmmyyMatch;
          let y = parseInt(year);
          if (y < 100) y += 2000;
          date = new Date(y, parseInt(month) - 1, parseInt(day));
        } else {
          date = new Date(dateValue);
        }
      } else {
        date = new Date(dateValue);
      }

      if (isNaN(date.getTime())) return "Invalid Date";

      return date.toLocaleDateString("en-IN", {
        year: "numeric",
        month: "short",
        day: "numeric",
      });
    } catch {
      return "Invalid Date";
    }
  };

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(
      comparisonData.map((row) => ({
        "Employee ID": row.employeeId,
        "Employee Name": row.employeeName,
        Department: row.department,
        "Date of Joining": formatDate(row.dateOfJoining),
        "Percentage (%)": row.percentage,
        "Gross(Software)": row.grossSal,
        "Gross2(Software)": row.calculatedValue,
        "Gross2(HR)": row.gross2HR,
        Difference: row.difference,
        Status: row.status,
      }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Step 4 Comparison");
    XLSX.writeFile(wb, `Step4-Gross2-Comparison-Staff.xlsx`);
  };

  // ========== Calculate Grand Totals ==========
  const calculateGrandTotals = () => {
    const totalGross2Software = filteredData.reduce(
      (sum, row) => sum + (Number(row.calculatedValue) || 0),
      0
    );
    const totalGross2HR = filteredData.reduce(
      (sum, row) => sum + (Number(row.gross2HR) || 0),
      0
    );
    const totalDifference = totalGross2Software - totalGross2HR;

    return {
      totalGross2Software,
      totalGross2HR,
      totalDifference,
    };
  };

  const grandTotals = calculateGrandTotals();

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
            File is ready for processing
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
          <p className="text-xs text-gray-500">
            This file was not uploaded in the previous steps
          </p>
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
                Step 4 - Staff Bonus Calculation
              </h1>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("/step3")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                ← Back to Step 3
              </button>
              <button
                onClick={() => router.push("/")}
                className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
              >
                Back to Step 1
              </button>
            </div>
          </div>

          <div className="mb-8 bg-blue-50 border border-blue-200 rounded-lg p-6">
            <h3 className="font-bold text-blue-900 mb-3 flex items-center gap-2">
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
              Calculation Formula (Staff Only)
            </h3>
            <div className="text-sm text-blue-800 space-y-2">
              <p>
                <strong>Excel Formula:</strong> =IF(X=8.33, Q, IF(X&gt;8.33,
                Q*0.6, ""))
              </p>
              <p>
                <strong>Where:</strong>
              </p>
              <ul className="list-disc ml-6 space-y-1">
                <li>X = Percentage (calculated from Date of Joining)</li>
                <li>
                  Q = Gross(Software) (sum of monthly salaries + Oct 2025
                  estimate)
                </li>
              </ul>
              <p>
                <strong>Logic:</strong>
              </p>
              <ul className="list-disc ml-6 space-y-1">
                <li>
                  If percentage = 8.33% → Gross2(Software) = Gross(Software)
                </li>
                <li>
                  If percentage &gt; 8.33% (10% or 12%) → Gross2(Software) =
                  Gross(Software) × 0.6
                </li>
              </ul>
              <p>
                <strong>October Filtering:</strong> If employee ID exists in the
                "Average" sheet of Actual Percentage file, October estimate is
                set to 0
              </p>
              <p>
                <strong>Gross2(HR):</strong> SUM of all "GROSS 02" values from
                bonus file (for duplicate employee IDs)
              </p>
              <p className="text-xs text-blue-600 mt-2">
                Percentage is calculated based on service period as of Oct 12,
                2025: &lt;12 months = 10% | 12-23 months = 12% | ≥24 months =
                8.33%
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data"
            />

            <FileCard
              title="Bonus Sheet"
              file={bonusFile}
              description="Bonus calculation data with GROSS 02 (Staff sheet only)"
            />

            <FileCard
              title="Actual Percentage"
              file={actualPercentageFile}
              description="Optional: Controls October estimate filtering"
            />
          </div>

          {(!staffFile || !bonusFile) && (
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
                    Required files are missing
                  </h3>
                  <p className="text-sm text-yellow-600 mt-1">
                    Please upload Indiana Staff and Bonus Sheet files in Step 1
                    to proceed.
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
                  Processing files and calculating values...
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
                  Staff Bonus Comparison Results
                </h2>
                <div className="flex gap-3">
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
                  <button
                    onClick={() => router.push("/step5")}
                    className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition flex items-center gap-2"
                  >
                    Move to Step 5
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
                  </button>
                </div>
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
                              Department
                              <SortArrows columnKey="department" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-left bg-gray-100">
                            <div className="flex items-center">
                              Date of Joining
                              <SortArrows columnKey="dateOfJoining" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              %
                              <SortArrows columnKey="percentage" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Gross(Software)
                              <SortArrows columnKey="grossSal" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Gross2(Software)
                              <SortArrows columnKey="calculatedValue" />
                            </div>
                          </th>
                          <th className="border border-gray-300 px-4 py-3 text-right bg-gray-100">
                            <div className="flex items-center justify-end">
                              Gross2(HR)
                              <SortArrows columnKey="gross2HR" />
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
                            className={
                              idx % 2 === 0 ? "bg-white" : "bg-gray-50"
                            }
                          >
                            <td className="border border-gray-300 px-4 py-2">
                              {row.employeeId}
                            </td>
                            <td className="border border-gray-300 px-4 py-2">
                              {row.employeeName}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-center">
                              <span className="px-2 py-1 bg-indigo-100 text-indigo-800 rounded text-xs font-medium">
                                {row.department}
                              </span>
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-sm">
                              {formatDate(row.dateOfJoining)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right font-medium">
                              {row.percentage}%
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right">
                              {formatCurrency(row.grossSal)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right font-medium text-blue-600">
                              {formatCurrency(row.calculatedValue)}
                            </td>
                            <td className="border border-gray-300 px-4 py-2 text-right font-medium text-purple-600">
                              {formatCurrency(row.gross2HR)}
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
                        <tr className="bg-indigo-100 font-bold sticky bottom-0">
                          <td
                            colSpan={6}
                            className="border border-gray-300 px-4 py-3 text-right"
                          >
                            <span className="text-lg">GRAND TOTAL</span>
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-indigo-900">
                            {formatCurrency(grandTotals.totalGross2Software)}
                          </td>
                          <td className="border border-gray-300 px-4 py-3 text-right text-indigo-900">
                            {formatCurrency(grandTotals.totalGross2HR)}
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
                <div>Total Staff Employees: {filteredData.length}</div>
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
