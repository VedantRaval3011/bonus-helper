"use client";

import React, { useState, useEffect } from "react";
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

  // Helper to normalize header text
  const norm = (x: any) =>
    String(x ?? "")
      .replace(/\s+/g, "")
      .replace(/[-_.]/g, "")
      .toUpperCase();

  // Default and special percentages
  const DEFAULT_PERCENTAGE = 8.33;
  const SPECIAL_PERCENTAGE = 12.0;
  const TOLERANCE = 12;

  const processFiles = async () => {
    if (!staffFile || !workerFile || !bonusFile || !actualPercentageFile) {
      setError("All four files are required for processing");
      return;
    }

    setIsProcessing(true);
    setError(null);

    try {
      console.log("="  .repeat(60));
      console.log("📊 STEP 5: Register Calculation");
      console.log("="  .repeat(60));

      // ========== LOAD ACTUAL PERCENTAGE DATA ==========
      const actualPercentageBuffer = await actualPercentageFile.arrayBuffer();
      const actualPercentageWorkbook = XLSX.read(actualPercentageBuffer);
      const actualPercentageSheet =
        actualPercentageWorkbook.Sheets[actualPercentageWorkbook.SheetNames[0]];
      const actualPercentageData: any[][] = XLSX.utils.sheet_to_json(
        actualPercentageSheet,
        { header: 1 }
      );

      // Find employees with special percentage (12%)
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
              console.log(
                `✨ Employee ${empCode}: Special percentage ${SPECIAL_PERCENTAGE}%`
              );
            }
          }
        }
      }

      console.log(
        `📋 Employees with ${SPECIAL_PERCENTAGE}% bonus: ${Array.from(
          specialPercentageEmployees
        ).join(", ")}`
      );

      // ========== LOAD BONUS FILE FOR HR REGISTER VALUES ==========
      const bonusBuffer = await bonusFile.arrayBuffer();
      const bonusWorkbook = XLSX.read(bonusBuffer);

      const hrRegisterData: Map<
        number,
        { registerHR: number; dept: string }
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
              hrRegisterData.set(empCode, {
                registerHR: registerHR,
                dept: "Worker",
              });
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
              hrRegisterData.set(empCode, {
                registerHR: registerHR,
                dept: "Staff",
              });
            }
          }
        }
      }

      console.log(`✅ HR Register data loaded: ${hrRegisterData.size} employees`);

      // ========== LOAD GROSS SALARY DATA FROM STEP 2 ==========
      // We'll compute gross salary from Staff and Worker files
      const grossSalaryData: Map<
        number,
        { name: string; dept: string; grossSalary: number }
      > = new Map();

      // Process Staff file for Gross Salary
      const staffBuffer = await staffFile.arrayBuffer();
      const staffWorkbook = XLSX.read(staffBuffer);

      for (let sheetName of staffWorkbook.SheetNames) {
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

        if (empIdIdx === -1 || empNameIdx === -1 || salary1Idx === -1) continue;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;

          if (!empId || isNaN(empId) || !empName) continue;

          if (!grossSalaryData.has(empId)) {
            grossSalaryData.set(empId, {
              name: empName,
              dept: "Staff",
              grossSalary: 0,
            });
          }

          const emp = grossSalaryData.get(empId)!;
          emp.grossSalary += salary1;
        }
      }

      // Process Worker file for Gross Salary
      const workerBuffer = await workerFile.arrayBuffer();
      const workerWorkbook = XLSX.read(workerBuffer);

      for (let sheetName of workerWorkbook.SheetNames) {
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
        const salary1Idx = 8; // Column I

        if (empIdIdx === -1 || empNameIdx === -1) continue;

        for (let i = headerIdx + 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;

          const empId = Number(row[empIdIdx]);
          const empName = String(row[empNameIdx] || "").trim().toUpperCase();
          const salary1 = Number(row[salary1Idx]) || 0;

          if (!empId || isNaN(empId) || !empName) continue;

          if (!grossSalaryData.has(empId)) {
            grossSalaryData.set(empId, {
              name: empName,
              dept: "Worker",
              grossSalary: 0,
            });
          }

          const emp = grossSalaryData.get(empId)!;
          emp.grossSalary += salary1;
        }
      }

      console.log(`✅ Gross Salary data loaded: ${grossSalaryData.size} employees`);

      // ========== CALCULATE REGISTER ==========
      const comparison: any[] = [];

      for (const [empId, empData] of grossSalaryData) {
        // Determine percentage
        const percentage = specialPercentageEmployees.has(empId)
          ? SPECIAL_PERCENTAGE
          : DEFAULT_PERCENTAGE;

        // Calculate Register (Software) = Gross Salary × (Percentage / 100)
        const registerSoftware = (empData.grossSalary * percentage) / 100;

        // Get Register (HR) from bonus file
        const hrData = hrRegisterData.get(empId);
        const registerHR = hrData?.registerHR || 0;

        // Calculate difference
        const difference = registerSoftware - registerHR;
        const status = Math.abs(difference) <= TOLERANCE ? "Match" : "Mismatch";

        comparison.push({
          employeeId: empId,
          employeeName: empData.name,
          department: empData.dept,
          grossSalarySoftware: empData.grossSalary,
          percentage: percentage,
          registerSoftware: registerSoftware,
          registerHR: registerHR,
          difference: difference,
          status: status,
        });

        console.log(
          `💰 Emp ${empId}: Gross=₹${empData.grossSalary.toFixed(
            2
          )}, ${percentage}% → Register(SW)=₹${registerSoftware.toFixed(
            2
          )}, Register(HR)=₹${registerHR.toFixed(2)}, Diff=₹${difference.toFixed(2)}`
        );
      }

      comparison.sort((a, b) => a.employeeId - b.employeeId);
      setComparisonData(comparison);
      setFilteredData(comparison);

      console.log("✅ Register calculation completed");
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
    // eslint-disable-next-line
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
        Difference: row.difference,
        Status: row.status,
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Register Calculation");
    XLSX.writeFile(wb, `Step5-Register-Calculation-${departmentFilter}.xlsx`);
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
          <p className="text-xs text-gray-500">Upload in Step 1</p>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-purple-50 to-pink-100 py-5 px-4">
      <div className="mx-auto max-w-7xl">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Step 5 - Register Calculation
              </h1>
              <p className="text-gray-600 mt-2">
                Register (Software) = Gross Salary × Percentage (8.33% default,
                12% special)
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
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
            <FileCard
              title="Indiana Staff"
              file={staffFile}
              description="Staff salary data"
              icon={<></>}
            />
            <FileCard
              title="Indiana Worker"
              file={workerFile}
              description="Worker salary data"
              icon={<></>}
            />
            <FileCard
              title="Bonus Calculation Sheet"
              file={bonusFile}
              description="HR Register values (Worker: Col S, Staff: Col T)"
              icon={<></>}
            />
            <FileCard
              title="Actual Percentage Data"
              file={actualPercentageFile}
              description="Employees with 12% bonus percentage"
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
                  Calculating Register (Software vs HR)...
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
                      <th className="border border-gray-300 px-4 py-2 text-center">
                        %
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Register (Software)
                      </th>
                      <th className="border border-gray-300 px-4 py-2 text-right">
                        Register (HR)
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
                  {filteredData.filter((r) => r.status === "Mismatch").length} |
                  Special %:{" "}
                  {filteredData.filter((r) => r.percentage === 12.0).length}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
