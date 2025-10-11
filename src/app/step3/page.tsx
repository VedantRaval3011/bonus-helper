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
  
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [password, setPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");

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

  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  const MONTH_NAME_MAP: Record<string, number> = {
    JAN: 1, JANUARY: 1, FEB: 2, FEBRUARY: 2, MAR: 3, MARCH: 3,
    APR: 4, APRIL: 4, MAY: 5, JUN: 6, JUNE: 6, JUL: 7, JULY: 7,
    AUG: 8, AUGUST: 8, SEP: 9, SEPT: 9, SEPTEMBER: 9,
    OCT: 10, OCTOBER: 10, NOV: 11, NOVEMBER: 11, DEC: 12, DECEMBER: 12,
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
  const TOLERANCE = 12;

  // Employees who will NOT get October estimate
  const EXCLUDE_OCTOBER_EMPLOYEES = new Set<number>([
    937, 1039, 1065, 1105, 59, 161
  ]);

  // *** UPDATED: Added employee 20 (Sanjay Rathod) ***
  // Employees who should calculate Oct estimate INCLUDING zeros
  const INCLUDE_ZEROS_IN_AVG = new Set<number>([
    20,   // SANJAY RATHOD (Staff - zero in April) - ADDED!
    27,   // KIRAN SASANIYA (zero in July, Aug)
    882,  // SHRADDHA DHODHAKIYA (zero in June)
    898,  // RAMNIK SOLANKI (zero in March, April)
    999,  // HANSHABEN PARMAR (zero in April, May) - starts from Dec-24
  ]);

  // Employees who should calculate Oct estimate EXCLUDING zeros
  const EXCLUDE_ZEROS_IN_AVG = new Set<number>([
    1054, // Different employee (joined April 2025)
  ]);

  // Employee-specific work start months
  const EMPLOYEE_START_MONTHS: Record<number, string> = {
    999: "2024-12",  // Hanshaben Parmar joined December 2024
  };

  const processFiles = async () => {
    if (!staffFile || !workerFile || !bonusFile) {
      setError("All three files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("=".repeat(60));
      console.log("🚫 OCTOBER EXCLUDE LIST:", Array.from(EXCLUDE_OCTOBER_EMPLOYEES).join(", "));
      console.log("✅ INCLUDE ZEROS IN AVG:", Array.from(INCLUDE_ZEROS_IN_AVG).join(", "));
      console.log("⭕ EXCLUDE ZEROS IN AVG:", Array.from(EXCLUDE_ZEROS_IN_AVG).join(", "));
      console.log("📅 CUSTOM START MONTHS:", JSON.stringify(EMPLOYEE_START_MONTHS));
      console.log("🚫 EXCLUDED DEPARTMENTS: C (Cash), A");
      console.log("=".repeat(60));

      // ========== PROCESS STAFF FILE ==========
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);
      
      const staffEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
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

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) {
          console.log(`⚠️ Skip Staff ${sheetName}: missing columns`);
          continue;
        }

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1Raw = row[salary1Idx];
          
          const salary1 = (salary1Raw === null || salary1Raw === undefined || salary1Raw === "") 
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
          
          // Check if employee should be processed for this month
          const startMonth = EMPLOYEE_START_MONTHS[empId];
          if (startMonth && monthKey < startMonth) {
            continue;
          }
          
          // Store zero values for special employees
          if (INCLUDE_ZEROS_IN_AVG.has(empId) || EXCLUDE_ZEROS_IN_AVG.has(empId)) {
            emp.months.set(monthKey, salary1);
          } else {
            if (salary1 > 0) {
              emp.months.set(monthKey, salary1);
            }
          }
        }
      }

      console.log(`✅ Staff employees: ${staffEmployees.size}`);

      // ========== PROCESS WORKER FILE ==========
      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);
      
      const workerEmployees: Map<
        number,
        { name: string; dept: string; months: Map<string, number> }
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
          return normalized === "DEPT" || normalized === "DEPARTMENT" || normalized === "DEPTT";
        });
        
        const salary1Idx = 8; // Column I

        if (empIdIdx === -1 || empNameIdx === -1) {
          console.log(`⚠️ Skip Worker ${sheetName}: missing columns`);
          continue;
        }

        if (deptIdx === -1) {
          console.log(`⚠️ Warning: Department column not found in ${sheetName}`);
        }

        let excludedDeptCount = 0;
        const excludedDeptBreakdown: Record<string, number> = {};

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1Raw = row[salary1Idx];
          
          const salary1 = (salary1Raw === null || salary1Raw === undefined || salary1Raw === "") 
            ? 0 
            : Number(salary1Raw) || 0;

          if (deptIdx !== -1) {
            const dept = String(row[deptIdx] || "").trim().toUpperCase();
            
            if (EXCLUDED_DEPARTMENTS.includes(dept)) {
              excludedDeptCount++;
              excludedDeptBreakdown[dept] = (excludedDeptBreakdown[dept] || 0) + 1;
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
          
          // Check if employee should be processed for this month
          const startMonth = EMPLOYEE_START_MONTHS[empId];
          if (startMonth && monthKey < startMonth) {
            continue;
          }
          
          // Store zero values for special employees
          if (INCLUDE_ZEROS_IN_AVG.has(empId) || EXCLUDE_ZEROS_IN_AVG.has(empId)) {
            emp.months.set(monthKey, salary1);
          } else {
            if (salary1 > 0) {
              emp.months.set(monthKey, salary1);
            }
          }
        }

        if (excludedDeptCount > 0) {
          const breakdown = Object.entries(excludedDeptBreakdown)
            .map(([dept, count]) => `${dept}=${count}`)
            .join(", ");
          console.log(`💰 Filtered ${excludedDeptCount} workers from ${sheetName} (${breakdown})`);
        }
      }

      console.log(`✅ Worker employees: ${workerEmployees.size}`);

      // ========== PROCESS BONUS FILE ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);
      
      const bonusEmployees: Map<
        number,
        { name: string; grossSalary: number; dept: string }
      > = new Map();

      for (let sheetName of bonusWorkbook.SheetNames) {
        console.log(`📊 Processing bonus: ${sheetName}`);
        const bonusSheet = bonusWorkbook.Sheets[sheetName];
        const bonusData: any[][] = XLSX.utils.sheet_to_json(bonusSheet, {
          header: 1,
        });

        const headerRows: number[] = [];
        for (let i = 0; i < bonusData.length; i++) {
          if (
            bonusData[i] &&
            bonusData[i].includes("EMP Code") &&
            bonusData[i].includes("EMP. NAME")
          ) {
            headerRows.push(i);
          }
        }

        console.log(`   Found ${headerRows.length} data section(s) in ${sheetName}`);

        for (let sectionIdx = 0; sectionIdx < headerRows.length; sectionIdx++) {
          const bonusHeaderRow = headerRows[sectionIdx];
          const nextHeaderRow = headerRows[sectionIdx + 1] || bonusData.length;
          
          console.log(`   📋 Processing section ${sectionIdx + 1} (rows ${bonusHeaderRow} to ${nextHeaderRow - 1})`);

          const headers = bonusData[bonusHeaderRow];
          const empCodeIdx = headers.indexOf("EMP Code");
          const empNameIdx = headers.indexOf("EMP. NAME");
          const deptIdx = headers.indexOf("Deptt.");
          const grossIdx = headers.findIndex(
            (h: any) =>
              typeof h === "string" &&
              /^(GROSS|Gross|GROSS SAL\.)$/i.test(h.trim())
          );

          if (grossIdx === -1) {
            console.log(`   ⚠️ Skip section ${sectionIdx + 1}: No Gross column`);
            continue;
          }

          for (let i = bonusHeaderRow + 1; i < nextHeaderRow; i++) {
            const row = bonusData[i];
            if (!row || row.length === 0) continue;

            const empId = Number(row[empCodeIdx]);
            const empName = String(row[empNameIdx] || "").trim().toUpperCase();
            const dept = String(row[deptIdx] || "").trim().toUpperCase();
            const grossRaw = row[grossIdx];
            
            const gross = (grossRaw === null || grossRaw === undefined || grossRaw === "") 
              ? 0 
              : Number(grossRaw) || 0;

            if (!empId || isNaN(empId) || !empName) continue;
            if (gross === 0) continue;

            const deptType =
              dept === "W" ? "Worker" : dept === "M" || dept === "S" ? "Staff" : "Unknown";

            if (bonusEmployees.has(empId)) {
              const existing = bonusEmployees.get(empId)!;
              const prevGross = existing.grossSalary;
              existing.grossSalary += gross;
              
              console.log(`   🔄 EMP ${empId}: Adding ₹${gross.toFixed(2)} to existing ₹${prevGross.toFixed(2)} = ₹${existing.grossSalary.toFixed(2)}`);
            } else {
              bonusEmployees.set(empId, {
                name: empName,
                grossSalary: gross,
                dept: deptType,
              });
              console.log(`   ➕ EMP ${empId}: New entry with ₹${gross.toFixed(2)}`);
            }
          }
        }
      }

      console.log(`✅ Bonus employees: ${bonusEmployees.size}`);

      // ========== COMPUTE SOFTWARE TOTALS ==========
      const softwareEmployeesTotals: Map<
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
          
          // Build custom window for employees with start months
          const employeeWindow = hasCustomStart
            ? AVG_WINDOW.filter(mk => mk >= EMPLOYEE_START_MONTHS[empId])
            : AVG_WINDOW;
          
          // Collect all months in window
          for (const mk of employeeWindow) {
            const v = rec.months.get(mk);
            
            if (includeZeros) {
              // For employees with genuine zero months, include them
              const val = v !== undefined ? Number(v) : 0;
              baseSum += val;
              monthsIncluded.push({ month: mk, value: val });
            } else if (excludeZeros) {
              // For mid-year joiners, only count months they were employed
              if (v !== undefined && v !== null) {
                const val = Number(v);
                baseSum += val;
                monthsIncluded.push({ month: mk, value: val });
              }
            } else {
              // Normal employees: only non-zero months
              if (v != null && !isNaN(Number(v)) && Number(v) > 0) {
                baseSum += Number(v);
                monthsIncluded.push({ month: mk, value: Number(v) });
              }
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
            // Calculate October estimate
            if (includeZeros) {
              // Use employeeWindow length for custom start dates
              const divisor = hasCustomStart ? employeeWindow.length : 11;
              estOct = baseSum / divisor;
              
              // Special logging for specific employees
              if (empId === 20) {
                console.log(
                  `✅ EMP ${empId} (${rec.name}): STAFF + INCLUDE ZERO (April)\n` +
                  `   All 11 months counted\n` +
                  `   Base Sum: ₹${baseSum.toFixed(2)}\n` +
                  `   Oct Estimate (÷${divisor}): ₹${estOct.toFixed(2)}\n` +
                  `   TOTAL: ₹${(baseSum + estOct).toFixed(2)}`
                );
              } else if (empId === 999) {
                console.log(
                  `✅ EMP ${empId} (${rec.name}): CUSTOM START (Dec-24) + INCLUDE ZEROS (Apr & May)\n` +
                  `   Work window: Dec-24 to Sep-25 (${divisor} months)\n` +
                  `   Base Sum: ₹${baseSum.toFixed(2)}\n` +
                  `   Oct Estimate (÷${divisor}): ₹${estOct.toFixed(2)}\n` +
                  `   TOTAL: ₹${(baseSum + estOct).toFixed(2)}`
                );
              } else {
                console.log(
                  `✅ EMP ${empId} (${rec.name}): INCLUDING ZEROS\n` +
                  `   Months: ${divisor} (${hasCustomStart ? 'custom start' : 'all months'})\n` +
                  `   Base Sum: ₹${baseSum.toFixed(2)}\n` +
                  `   Oct Estimate: ₹${estOct.toFixed(2)}\n` +
                  `   TOTAL: ₹${(baseSum + estOct).toFixed(2)}`
                );
              }
            } else {
              // Average of only counted months
              const values = monthsIncluded.map(m => m.value);
              estOct = values.reduce((a, b) => a + b, 0) / values.length;
              
              if (excludeZeros) {
                console.log(
                  `⭕ EMP ${empId} (${rec.name}): EXCLUDING ZEROS (mid-year joiner)\n` +
                  `   Months counted: ${monthsIncluded.length}\n` +
                  `   Base Sum: ₹${baseSum.toFixed(2)}\n` +
                  `   Oct Estimate: ₹${estOct.toFixed(2)}\n` +
                  `   TOTAL: ₹${(baseSum + estOct).toFixed(2)}`
                );
              }
            }
            total = baseSum + estOct;
          } else {
            console.log(
              `⚠️ EMP ${empId} (${rec.name}): No Sep-25 data - Base only = ₹${baseSum.toFixed(2)}`
            );
          }

          const prev = softwareEmployeesTotals.get(empId);
          if (!prev) {
            softwareEmployeesTotals.set(empId, {
              name: rec.name,
              dept: rec.dept,
              grossSalary: total,
            });
          } else {
            prev.grossSalary += total;
            if (prev.dept !== "Staff" && rec.dept === "Staff") {
              prev.name = rec.name;
              prev.dept = rec.dept;
            }
          }
        }
      };

      foldMonthly(staffEmployees);
      foldMonthly(workerEmployees);

      // ========== BUILD COMPARISON ==========
      const comparison: any[] = [];
      
      for (const [empId, empData] of softwareEmployeesTotals) {
        const b = bonusEmployees.get(empId);
        const difference = empData.grossSalary - (b?.grossSalary || 0);
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

  const handleMoveToStep4 = () => {
    const mismatchCount = comparisonData.filter((r) => r.status === "Mismatch").length;
    
    if (mismatchCount > 1) {
      setShowPasswordModal(true);
      setPassword("");
      setPasswordError("");
    } else {
      router.push("/step4");
    }
  };

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    const correctPassword = process.env.NEXT_PUBLIC_NEXT_PASSWORD;
    
    if (password === correctPassword) {
      setShowPasswordModal(false);
      router.push("/step4");
    } else {
      setPasswordError("Incorrect password. Please try again.");
      setPassword("");
    }
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
          <p className="text-xs text-gray-500">
            Upload in Step 1
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
                Step 3 - Gross Salary Comparison
              </h1>
              <p className="text-gray-600 mt-2">
                Nov-24 to Sep-25 + Oct-25 avg (Special: Staff ID 20 Apr=0, Worker ID 999 Dec start Apr&May=0, IDs 27,882,898 zeros | Excludes Dept A&C, IDs: 937,1039,1065,1105,59,161)
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

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data"
              icon={<></>}
            />
            <FileCard
              title="Indiana Worker"
              file={workerFile}
              description="Worker salary data (Dept A & C excluded)"
              icon={<></>}
            />
            <FileCard
              title="Bonus Sheet"
              file={bonusFile}
              description="Bonus calculation (all sections)"
              icon={<></>}
            />
          </div>

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
                  Processing with special handling (Staff ID 20 Apr=0, Worker ID 999 Dec-start)...
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
                    Comparison Results
                  </h2>
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
                    onClick={handleMoveToStep4}
                    className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition flex items-center gap-2"
                  >
                    Move to Step 4
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
                        Gross (Software)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Gross (HR)
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

      {/* Password Modal */}
      {showPasswordModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full mx-4">
            <div className="flex items-center justify-between mb-6">
              <h3 className="text-2xl font-bold text-gray-800">
                Password Required
              </h3>
              <button
                onClick={() => setShowPasswordModal(false)}
                className="text-gray-400 hover:text-gray-600 transition"
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
            
            <p className="text-gray-600 mb-6">
              There are more than 1 mismatches. Please enter the password to proceed to Step 4.
            </p>

            <form onSubmit={handlePasswordSubmit}>
              <div className="mb-4">
                <label
                  htmlFor="password"
                  className="block text-sm font-medium text-gray-700 mb-2"
                >
                  Password
                </label>
                <input
                  type="password"
                  id="password"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    setPasswordError("");
                  }}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  placeholder="Enter password"
                  autoFocus
                />
                {passwordError && (
                  <p className="mt-2 text-sm text-red-600 flex items-center gap-1">
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
                        d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                      />
                    </svg>
                    {passwordError}
                  </p>
                )}
              </div>

              <div className="flex gap-3">
                <button
                  type="button"
                  onClick={() => setShowPasswordModal(false)}
                  className="flex-1 px-4 py-2 bg-gray-200 text-gray-800 rounded-lg hover:bg-gray-300 transition"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
                >
                  Submit
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}