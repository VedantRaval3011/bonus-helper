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

  // **CORRECTED PERCENTAGE CALCULATION** - Calculate based on current date (Oct 12, 2025)
  const calculatePercentage = (dateOfJoining: any): number => {
    if (!dateOfJoining) return 0;

    let doj: Date;

    // Handle Excel serial date number
    if (typeof dateOfJoining === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      doj = new Date(excelEpoch.getTime() + dateOfJoining * 86400000);
    } else {
      doj = new Date(dateOfJoining);
    }

    // Check if date is valid
    if (isNaN(doj.getTime())) return 0;

    // Calculate to 30th Oct 2025
    const referenceDate = new Date(2025, 8, 30); // FIXED: Month is 0-indexed, so 9 = October

    // Calculate the difference in years and months
    const yearsDiff = referenceDate.getFullYear() - doj.getFullYear();
    const monthsDiff = referenceDate.getMonth() - doj.getMonth();
    const daysDiff = referenceDate.getDate() - doj.getDate();

    // Calculate total months
    let totalMonths = yearsDiff * 12 + monthsDiff;

    // If the day of the reference date is before the day of joining,
    // subtract one month because a full month hasn't been completed yet
    if (daysDiff < 0) {
      totalMonths--;
    }

    console.log(
      `DOJ: ${doj.toLocaleDateString()}, Reference: ${referenceDate.toLocaleDateString()}, Total Months: ${totalMonths}`
    );

    // Apply percentage rules based on months of service
    if (totalMonths < 12) {
      return 10; // Less than 1 year → 10%
    } else if (totalMonths >= 12 && totalMonths < 24) {
      return 12; // 1 year to less than 2 years → 12%
    } else {
      return 8.33; // 2 years or more → 8.33%
    }
  };

  // *CORRECTED FORMULA*: =IF(X=8.33, Q, IF(X>8.33, Q*0.6, ""))
  // Where X = percentage, Q = Gross SAL. (our gross2Software)
  const applyBonusFormula = (grossSal: number, percentage: number): number => {
    if (percentage === 8.33) {
      return grossSal;
    } else if (percentage > 8.33) {
      // This applies to 10% and 12% cases
      return grossSal * 0.6;
    } else {
      // For percentage < 8.33, Excel returns "" (empty), we'll return 0
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
      // ========== STEP 0: Read Actual Percentage file and get Average sheet employee IDs ==========
      let employeesToIgnoreOctober = new Set<number>();

      if (actualPercentageFile) {
        try {
          console.log("Processing Actual Percentage file...");
          const actualBuffer = await actualPercentageFile.arrayBuffer();
          const actualWorkbook = XLSX.read(actualBuffer);

          // Read the Average sheet to get employee IDs to ignore
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

            // Find header row in Average sheet
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
                // Collect all employee IDs from Average sheet
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
          // Don't fail the entire process, just log and continue
        }
      } else {
        console.log(
          "Actual Percentage file not found - proceeding without October filtering"
        );
      }

      // ========== STEP 1: Process Staff file to get Gross SAL. (Software) ==========
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

      // In processFiles function, around line 670-730 where Staff file is processed:

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

        // **ADD THIS DEBUG LOG**
        console.log(`📋 Sheet: ${sheetName}`);
        console.log(`  Headers found at row ${headerIdx}`);
        console.log(`  EMP ID column index: ${empIdIdx}`);
        console.log(`  Name column index: ${empNameIdx}`);
        console.log(`  SALARY1 column index: ${salary1Idx}`);
        console.log(`  DOJ column index: ${dojIdx}`); // ← Check if this is -1
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

          // **ADD THIS DEBUG LOG FOR EMPLOYEE 554**
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
            // **CRITICAL FIX**: If DOJ already exists, don't overwrite it with null
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

      // ========== STEP 2: Calculate GROSS SAL. (Software) with Oct-2025 estimate ==========
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
        // Sum all known months
        let baseSum = 0;
        for (const v of rec.months.values()) {
          baseSum += Number(v) || 0;
        }

        // Only calculate October 2025 estimate if September 2025 data exists
        // AND employee is NOT in the Average sheet
        let estOct = 0;
        const hasSep2025 = rec.months.has("2025-09");
        const isInAverageSheet = employeesToIgnoreOctober.has(empId);

        if (hasSep2025 && !isInAverageSheet) {
          // Compute mean across window months that exist for this employee
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

      // ========== STEP 3: Read Bonus file to get GROSS 02 (HR) and Department - SUM ALL OCCURRENCES ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      // Map to store SUM of all Gross2 values per employee (Staff only)
      const bonusGross2Map: Map<number, number> = new Map();
      // Map to track employee names and departments
      const bonusEmployeeNames: Map<number, string> = new Map();
      const bonusEmployeeDepts: Map<number, string> = new Map();

      // Only process "Staff" sheet
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

      // Find header row - it's at row 1 for the Staff sheet
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

      // Find EMP Code column
      const empCodeIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("EMP") && hStr.includes("CODE");
      });

      // Find Employee Name column
      const empNameIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("EMP") && hStr.includes("NAME");
      });

      // Find Department column
      const deptIdx = headers.findIndex((h: any) => {
        const hStr = String(h || "").toUpperCase();
        return hStr.includes("DEPT") || hStr === "DEPTT." || hStr === "DEPTT";
      });

      // Find GROSS 02 column
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

      // **KEY FIX**: SUM all GROSS 02 values for each employee ID
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

        // **CRITICAL CHANGE**: Add to existing sum instead of replacing
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

      // Log employees with multiple entries
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

      // ========== STEP 4: Build final comparison data ==========
      const calculationResults: any[] = [];

      for (const [empId, empData] of softwareEmployeesTotals) {
        const percentage = calculatePercentage(empData.dateOfJoining);
        const gross2HR = bonusGross2Map.get(empId) || 0;
        const department = bonusEmployeeDepts.get(empId) || empData.dept;

        // Apply formula: IF(percentage=8.33, grossSal, IF(percentage>8.33, grossSal*0.6, grossSal*0.6))
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

      // Check for employees in bonus sheet but not in staff
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
    // eslint-disable-next-line
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
        // Excel serial date number
        const excelEpoch = new Date(1899, 11, 30);
        date = new Date(excelEpoch.getTime() + dateValue * 86400000);
      } else if (typeof dateValue === "string") {
        // Try parsing DD.MM.YY or DD-MM-YY format (Indian format)
        const ddmmyyMatch = dateValue.match(
          /^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})$/
        );
        if (ddmmyyMatch) {
          let [, day, month, year] = ddmmyyMatch;
          let y = parseInt(year);
          if (y < 100) y += 2000; // Convert 23 to 2023
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
        "GROSS SAL. (Software)": row.grossSal,
        "Calculated Value (GROSS 02)": row.calculatedValue,
        "GROSS 02 (HR)": row.gross2HR,
        Difference: row.difference,
        Status: row.status,
      }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Step 4 Comparison");
    XLSX.writeFile(wb, `Step4-Gross2-Comparison-Staff.xlsx`);
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
                Step 4 - Staff Bonus Calculation
              </h1>
              <p className="text-gray-600 mt-2">
                Calculate staff bonuses using the formula: IF(percentage=8.33,
                GROSS SAL., IF(percentage&gt;8.33, GROSS SAL.*0.6, GROSS
                SAL.*0.6))
              </p>
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

          {/* Formula Explanation */}
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
                  Q = GROSS SAL. (sum of monthly salaries + Oct 2025 estimate)
                </li>
              </ul>
              <p>
                <strong>Logic:</strong>
              </p>
              <ul className="list-disc ml-6 space-y-1">
                <li>If percentage = 8.33% → Calculated Value = GROSS SAL.</li>
                <li>
                  If percentage &gt; 8.33% (10% or 12%) → Calculated Value =
                  GROSS SAL. × 0.6
                </li>
              </ul>
              <p>
                <strong>October Filtering:</strong> If employee ID exists in the
                "Average" sheet of Actual Percentage file, October estimate is
                set to 0
              </p>
              <p>
                <strong>GROSS 02 (HR):</strong> SUM of all "GROSS 02" values
                from bonus file (for duplicate employee IDs)
              </p>
              <p className="text-xs text-blue-600 mt-2">
                <strong>✅ CORRECTED:</strong> Percentage is calculated based on
                service period as of Oct 12, 2025: &lt;12 months = 10% | 12-23
                months = 12% | ≥24 months = 8.33%
              </p>
            </div>
          </div>

          {/* File Cards */}
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

          {/* Missing Files Alert */}
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

          {/* Processing Status */}
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
                      <th className="border border-gray-300 px-4 py-2 text-left">
                        Date of Joining
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        %
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        GROSS SAL. (Software)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Calculated (GROSS 02)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        GROSS 02 (HR)
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
                  </tbody>
                </table>
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
