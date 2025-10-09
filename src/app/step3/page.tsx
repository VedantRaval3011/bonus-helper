"use client";

import React, { useState, useEffect } from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import * as XLSX from "xlsx";

export default function Step3Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [comparisonData, setComparisonData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [departmentFilter, setDepartmentFilter] = useState<string>("All");

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

  // Helper to normalize header text
  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  // NEW: month parsing + constants
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

  // Try to parse "YYYY-MM" from a sheet name like "Nov-24", "November 2024", "2025-09", etc.
  const parseMonthFromSheetName = (sheetName: string): string | null => {
    const s = String(sheetName || "").trim().toUpperCase();

    // Case 1: YYYY[-/_ ]MM
    const yyyymm = s.match(/(20\d{2})\D{0,2}(\d{1,2})/);
    if (yyyymm) {
      const y = Number(yyyymm[1]);
      const m = Number(yyyymm[2]);
      if (y >= 2000 && m >= 1 && m <= 12) return `${y}-${pad2(m)}`;
    }

    // Case 2: MON or MONTH with 2/4 digit year nearby
    const mon = s.match(/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b/);
    const monthFull = s.match(
      /\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/
    );
    const y2or4 = s.match(/\b(20\d{2}|\d{2})\b/);

    const monthToken = (monthFull?.[1] || mon?.[1]) as string | undefined;
    if (monthToken && y2or4) {
      let y = Number(y2or4[1]);
      if (y < 100) y += 2000; // assume 20xx for 2-digit years
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

  // Tolerance threshold for matching
  const TOLERANCE = 12;

  const processFiles = async () => {
    if (!staffFile || !workerFile || !bonusFile) {
      setError("All three files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      // Read Staff file
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);

      // REPLACE the staffEmployees declaration with months map
      const staffEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
      > = new Map();

      // Process Staff sheets - sum SALARY1 column per month
      for (let sheetName of staffWorkbook.SheetNames) {
        const sheet = staffWorkbook.Sheets[sheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Inside the Staff sheet loop, right after you compute headers and indices:
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";
        console.log(`Staff sheet: ${sheetName} -> monthKey: ${monthKey}`);

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
          console.log(
            `Skipping staff sheet ${sheetName}: missing required columns`
          );
          continue;
        }

        // REPLACE the row accumulation block inside Staff loop:
        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;
          if (!empId || isNaN(empId) || !empName) continue;

          if (!staffEmployees.has(empId)) {
            staffEmployees.set(empId, {
              name: empName,
              dept: "Staff",
              months: new Map(),
            });
          }
          const emp = staffEmployees.get(empId)!;
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
        }
      }

      // Read Worker file
      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);

      // REPLACE the workerEmployees declaration with months map
      const workerEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
      > = new Map();

      // Process Worker sheets - sum Salary1 column (column I, index 8) per month
      for (let sheetName of workerWorkbook.SheetNames) {
        console.log(`Processing worker sheet: ${sheetName}`);
        const sheet = workerWorkbook.Sheets[sheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Inside the Worker sheet loop, after headers are prepared:
        const monthKey = parseMonthFromSheetName(sheetName) ?? "unknown";
        console.log(`Worker sheet: ${sheetName} -> monthKey: ${monthKey}`);

        // Find header row
        let headerIdx = -1;
        for (let i = 0; i < Math.min(5, data.length); i++) {
          if (data[i] && data[i].some((v: any) => norm(v) === "EMPID")) {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx === -1) {
          console.log(`Sheet ${sheetName}: Cannot find header row`);
          continue;
        }

        const headers = data[headerIdx];

        // Find Employee ID column
        const empIdIdx = headers.findIndex((h: any) =>
          ["EMPID", "EMPCODE"].includes(norm(h))
        );

        // Find Employee Name column
        const empNameIdx = headers.findIndex((h: any) =>
          /EMPLOYEE\s*NAME/i.test(String(h ?? ""))
        );

        // SALARY1 is at column I (index 8 - 0-based indexing) for workers
        const salary1Idx = 8; // Column I

        // Log the column header at this position for verification
        if (headers[salary1Idx]) {
          console.log(
            `Worker sheet ${sheetName}: Column I header is "${headers[salary1Idx]}"`
          );
        }

        if (empIdIdx === -1 || empNameIdx === -1) {
          console.log(
            `Skipping worker sheet ${sheetName}: missing EmpId or Name columns`
          );
          continue;
        }

        // REPLACE the row accumulation block inside Worker loop:
        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;
          if (!empId || isNaN(empId) || !empName) continue;

          if (!workerEmployees.has(empId)) {
            workerEmployees.set(empId, {
              name: empName,
              dept: "Worker",
              months: new Map(),
            });
          }
          const emp = workerEmployees.get(empId)!;
          emp.months.set(monthKey, (emp.months.get(monthKey) || 0) + salary1);
        }
      }

      console.log(`Total worker employees processed: ${workerEmployees.size}`);

      // Read Bonus file
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);
      const bonusEmployees: Map<
        number,
        { name: string; grossSalary: number; dept: string }
      > = new Map();

      // Process ALL sheets in the bonus workbook
      for (let sheetName of bonusWorkbook.SheetNames) {
        console.log(`Processing bonus sheet: ${sheetName}`);

        const bonusSheet = bonusWorkbook.Sheets[sheetName];
        const bonusData: any[][] = XLSX.utils.sheet_to_json(bonusSheet, {
          header: 1,
        });

        // Detect header row for HR bonus sheet
        let bonusHeaderRow = -1;
        for (let i = 0; i < Math.min(8, bonusData.length); i++) {
          if (
            bonusData[i] &&
            bonusData[i].includes("EMP Code") &&
            bonusData[i].includes("EMP. NAME")
          ) {
            bonusHeaderRow = i;
            break;
          }
        }

        if (bonusHeaderRow === -1) {
          console.log(`Skipping sheet ${sheetName}: Cannot locate header row`);
          continue;
        }

        const headers = bonusData[bonusHeaderRow];
        const empCodeIdx = headers.indexOf("EMP Code");
        const empNameIdx = headers.indexOf("EMP. NAME");
        const deptIdx = headers.indexOf("Deptt.");

        // Find the Gross column - it could be "GROSS", "Gross", or "GROSS SAL."
        const grossIdx = headers.findIndex(
          (h: any) =>
            typeof h === "string" &&
            /^(GROSS|Gross|GROSS SAL\.)$/i.test(h.trim())
        );

        if (grossIdx === -1) {
          console.log(
            `Sheet ${sheetName}: Cannot locate Gross/GROSS SAL. column, skipping`
          );
          continue;
        }

        console.log(
          `Sheet ${sheetName}: Found columns - EmpCode:${empCodeIdx}, Name:${empNameIdx}, Dept:${deptIdx}, Gross:${grossIdx}`
        );

        // Process rows in this sheet
        for (let i = bonusHeaderRow + 1; i < bonusData.length; i++) {
          const row = bonusData[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empCodeIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const dept = String(row[deptIdx] || "").trim().toUpperCase();
          const gross = Number(row[grossIdx]) || 0;

          // Skip invalid rows
          if (!empId || isNaN(empId) || !empName || isNaN(gross)) continue;

          // Map department: W = Worker, M/S = Staff
          const deptType =
            dept === "W"
              ? "Worker"
              : dept === "M" || dept === "S"
              ? "Staff"
              : "Unknown";

          if (bonusEmployees.has(empId)) {
            const existing = bonusEmployees.get(empId)!;
            existing.grossSalary += gross;
          } else {
            bonusEmployees.set(empId, {
              name: empName,
              grossSalary: gross,
              dept: deptType,
            });
          }
        }

        console.log(
          `Sheet ${sheetName}: Processed ${bonusEmployees.size} total employees so far`
        );
      }

      // NEW: fold Staff and Worker monthly maps into software totals with Oct-2025 estimate
      const softwareEmployeesTotals: Map<
        number,
        { name: string; dept: string; grossSalary: number }
      > = new Map();

      // Helper to fold one map into totals
      const foldMonthly = (
        src: Map<
          number,
          { name: string; dept: string; months: Map<string, number> }
        >
      ) => {
        for (const [empId, rec] of src) {
          // sum all known months
          let baseSum = 0;
          for (const v of rec.months.values()) baseSum += Number(v) || 0;

          // NEW CONDITION: Only calculate October 2025 estimate if September 2025 data exists
          let estOct = 0;
          const hasSep2025 = rec.months.has("2025-09");
          
          if (hasSep2025) {
            // compute mean across window months that exist for this employee
            const values: number[] = [];
            for (const mk of AVG_WINDOW) {
              const v = rec.months.get(mk);
              if (v != null && !isNaN(Number(v))) values.push(Number(v));
            }
            estOct = values.length
              ? values.reduce((a, b) => a + b, 0) / values.length
              : 0;

            console.log(
              `Employee ${empId} (${rec.name}): Has Sep 2025 data, Base sum = ${baseSum}, Oct estimate = ${estOct}, Total = ${
                baseSum + estOct
              }`
            );
          } else {
            console.log(
              `Employee ${empId} (${rec.name}): Missing Sep 2025 data, skipping Oct estimate, Total = ${baseSum}`
            );
          }

          const total = baseSum + estOct;

          // write / merge into totals map
          const prev = softwareEmployeesTotals.get(empId);
          if (!prev) {
            softwareEmployeesTotals.set(empId, {
              name: rec.name,
              dept: rec.dept,
              grossSalary: total,
            });
          } else {
            // In case the same EmpID appears in both maps, add up (rare)
            prev.grossSalary += total;
            // Prefer Staff over Worker name/department if needed
            if (prev.dept !== "Staff" && rec.dept === "Staff") {
              prev.name = rec.name;
              prev.dept = rec.dept;
            }
          }
        }
      };

      // Fold Staff and Worker into totals
      foldMonthly(staffEmployees);
      foldMonthly(workerEmployees);

      // Use software totals (with Oct-2025 estimate where applicable)
      const allEmployees = softwareEmployeesTotals;

      // Build comparison using software totals (with Oct-2025 estimate)
      const comparison: any[] = [];
      for (const [empId, empData] of allEmployees) {
        const b = bonusEmployees.get(empId);

        const difference = empData.grossSalary - (b?.grossSalary || 0);
        
        // NEW: Use tolerance of ±12 for matching
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        comparison.push({
          employeeId: empId,
          employeeName: empData.name,
          department: empData.dept,
          grossSalarySoftware: empData.grossSalary,
          grossSalaryHR: b?.grossSalary || 0,
          difference: difference,
          status: status,
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
    if (staffFile && workerFile && bonusFile) {
      processFiles();
    }
    // eslint-disable-next-line
  }, [staffFile, workerFile, bonusFile]);

  // Apply department filter
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
        "Gross Salary (HR)": row.grossSalaryHR,
        Difference: row.difference,
        Status: row.status,
      }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Gross Salary Comparison");
    XLSX.writeFile(
      wb,
      `Step3-Gross-Salary-Comparison-${departmentFilter}.xlsx`
    );
  };

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
          {/* Header */}
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 3 - Gross Salary Comparison
              </h1>
              <p className="text-gray-600 mt-2">
                Compare gross salaries between Software and HR data (Oct 2025 estimated only if Sep 2025 exists, ±12 tolerance)
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => router.push("/step2")}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
              >
                ← Back to Step 2
              </button>
              <button
                onClick={() => router.push("/")}
                className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
              >
                Back to Step 1
              </button>
            </div>
          </div>

          {/* File Cards */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data file"
              icon={
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
                    d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z"
                  />
                </svg>
              }
            />

            <FileCard
              title="Indiana Worker"
              file={workerFile}
              description="Worker salary data file"
              icon={
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
                    d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"
                  />
                </svg>
              }
            />

            <FileCard
              title="Bonus Sheet"
              file={bonusFile}
              description="Bonus calculation data"
              icon={
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
                    d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1"
                  />
                </svg>
              }
            />
          </div>

          {/* Missing Files Alert */}
          {[staffFile, workerFile, bonusFile].filter(Boolean).length < 3 && (
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
                    Please upload all required files in Step 1 to proceed with
                    full processing capabilities.
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* Processing Status */}
          {isProcessing && (
            <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-4">
              <div className="flex items-center gap-3">
                <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
                <p className="text-blue-800">
                  Processing files and calculating gross salaries (Oct 2025 estimated only if Sep 2025 exists)...
                </p>
              </div>
            </div>
          )}

          {/* Error Message */}
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

          {/* Comparison Table */}
          {comparisonData.length > 0 && (
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <div className="flex items-center gap-4">
                  <h2 className="text-xl font-bold text-gray-800">
                    Gross Salary Comparison Results
                  </h2>

                  {/* Department Filter */}
                  <select
                    value={departmentFilter}
                    onChange={(e) => setDepartmentFilter(e.target.value)}
                    className="px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
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

              <div className="overflow-x-auto">
                <table className="w-full border-collapse">
                  <thead>
                    <tr className="bg-gray-100">
                      <th className="border border-gray-300 px-4 py-2 text-left">
                        Employee ID
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-left">
                        Employee Name
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-left">
                        Department
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Gross Salary (Software)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Gross Salary (HR)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Difference
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-center">
                        Status
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, idx) => (
                      <tr
                        key={idx}
                        className={idx % 2 === 0 ? "bg-white" : "bg-gray-50"}
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
                          {formatCurrency(row.grossSalaryHR)}
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
                  </tbody>
                </table>
              </div>

              <div className="mt-4 flex justify-between items-center text-sm text-gray-600">
                <div>
                  Total Employees: {filteredData.length} | Staff:{" "}
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
