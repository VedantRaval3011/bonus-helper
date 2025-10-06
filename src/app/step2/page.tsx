"use client";

import React from "react";
import { useRouter } from "next/navigation";
import { useFileContext } from "@/contexts/FileContext";
import { useState } from "react";
import ExcelJS from "exceljs";

interface CellTriplet {
  A: number;
  B: number;
  C: number;
}
interface ReportData {
  months: string[];
  departments: { name: string; data: Record<string, CellTriplet> }[];
}

export default function Step2Page() {
  const router = useRouter();
  const { fileSlots } = useFileContext();
  const [reportData, setReportData] = useState<ReportData | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);

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

  const monthWiseFile =
    pickFile((s) => s.type === "Month-Wise-Sheet") ??
    pickFile((s) => !!s.file && /month.*wise/i.test(s.file.name)); // optional [file:1]

  const bonusFile =
    pickFile((s) => s.type === "Bonus-Calculation-Sheet") ??
    pickFile(
      (s) =>
        !!s.file &&
        /bonus.*final.*calculation|bonus.*2024-25/i.test(s.file.name)
    ); // robust match [file:5]

  // Months: Nov of previous FY through Oct of current FY dynamically
  const generateMonthHeaders = () => {
    const months: string[] = [];
    const today = new Date();
    const startDate = new Date(today.getFullYear() - 1, 10, 1); // Nov prev year
    for (let i = 0; i < 12; i++) {
      const d = new Date(startDate);
      d.setMonth(startDate.getMonth() + i);
      const m = d.toLocaleDateString("en-US", { month: "short" });
      const y = d.getFullYear().toString().slice(-2);
      months.push(`${m}-${y}`);
    }
    return months;
  };

  // Utilities
  const num = (cell: ExcelJS.Cell): number => {
    const v: any = cell.value;
    if (typeof v === "number") return v;
    if (v && typeof v === "object" && "result" in v)
      return Number(v.result) || 0;
    const t = cell.text?.replace(/,/g, "") || "";
    const n = Number(t);
    return Number.isFinite(n) ? n : 0;
  };

  const round = (x: number) => Math.round(x || 0);

  // Find label row anywhere in first 12 columns
  const findRowByLabel = (
    ws: ExcelJS.Worksheet,
    predicate: (t: string) => boolean
  ): number => {
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 12; c++) {
        const t = row.getCell(c).value?.toString().toUpperCase() || "";
        if (predicate(t)) return r;
      }
    }
    return -1;
  };

  // Find header column by keywords in first 5 rows
  const findColByHeader = (
    ws: ExcelJS.Worksheet,
    must: string[],
    any: string[] = [],
    rowScan = 5
  ): { col: number; headerRow: number } => {
    for (let r = 1; r <= rowScan; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= row.cellCount; c++) {
        const t = row.getCell(c).value?.toString().toUpperCase() || "";
        if (
          must.every((k) => t.includes(k.toUpperCase())) &&
          (any.length === 0 || any.some((k) => t.includes(k.toUpperCase())))
        ) {
          return { col: c, headerRow: r };
        }
      }
    }
    return { col: -1, headerRow: -1 };
  };

  // place near your other helpers
  const getGrossSalaryGrandTotal = (ws: ExcelJS.Worksheet): number => {
    // find "GROSS SALARY" column (fallback col 9 = I)
    let gsCol = -1,
      headerRow = -1;
    for (let r = 1; r <= 5 && gsCol < 0; r++) {
      ws.getRow(r).eachCell((cell, c) => {
        const t = cell.value?.toString().toUpperCase() || "";
        if (t.includes("GROSS") && t.includes("SALARY")) {
          gsCol = c;
          headerRow = r;
        }
      });
    }
    if (gsCol < 0) {
      gsCol = 9;
      headerRow = 2;
    } // Column I is GROSS SALARY in Month‑Wise sheets [file:1]

    // find GRAND TOTAL row (label often in col A/B; scan several columns)
    let gtRow = -1;
    for (let r = 1; r <= ws.rowCount && gtRow < 0; r++) {
      for (let c = 1; c <= 12; c++) {
        const t = ws.getRow(r).getCell(c).value?.toString().toUpperCase() || "";
        if (t.includes("GRAND") && t.includes("TOTAL")) {
          gtRow = r;
          break;
        }
      }
    }

    if (gtRow > 0) {
      const v = num(ws.getRow(gtRow).getCell(gsCol));
      if (v) return v; // direct or cached formula result [file:1]
    }
    // fallback: sum column between header and total
    let sum = 0;
    const endRow = gtRow > 0 ? gtRow - 1 : ws.rowCount;
    for (let r = headerRow + 1; r <= endRow; r++)
      sum += num(ws.getRow(r).getCell(gsCol));
    return sum; // e.g., NOV-24 W shows GRAND TOTAL 4,753,310 in GROSS SALARY (col I) [file:1]
  };

  // Read a column total by GRAND TOTAL row; fallback sum between header and total
  const readColumnGrandTotal = (
    ws: ExcelJS.Worksheet,
    targetCol: number,
    headerRow: number
  ): number => {
    if (targetCol <= 0) return 0;
    const gtRow = findRowByLabel(
      ws,
      (t) => t.includes("GRAND") && t.includes("TOTAL")
    );
    if (gtRow > 0) {
      const totalCell = ws.getRow(gtRow).getCell(targetCol);
      const direct = num(totalCell);
      if (direct) return direct;
      // Fallback: sum cells from headerRow+1 to gtRow-1
      let sum = 0;
      for (let r = headerRow + 1; r < gtRow; r++)
        sum += num(ws.getRow(r).getCell(targetCol));
      return sum;
    }
    // If no GT row, fallback to summing all numeric cells below header
    let sum = 0;
    for (let r = Math.max(2, headerRow + 1); r <= ws.rowCount; r++)
      sum += num(ws.getRow(r).getCell(targetCol));
    return sum;
  };

  // Worker Salary1 sum (no TOTAL row present in Worker sheets)
  const sumWorkerSalary1 = (ws: ExcelJS.Worksheet): number => {
    // Find Salary1 col; fallback to column 9 (I) per worker format
    let { col, headerRow } = findColByHeader(ws, ["SALARY", "1"]);
    if (col < 0) {
      col = 9;
      headerRow = 2;
    }
    let sum = 0,
      blanks = 0;
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const v = num(ws.getRow(r).getCell(col));
      if (v > 0) {
        sum += v;
        blanks = 0;
      } else if (++blanks > 8) break;
    }
    return sum;
  };

  // Staff Salary1 TOTAL from TOTAL row
  const readStaffSalary1Total = (ws: ExcelJS.Worksheet): number => {
    const header3 = ws.getRow(3);
    let salary1Col = -1;
    header3.eachCell((cell, c) => {
      const t = cell.value?.toString().toUpperCase() || "";
      if (t.includes("SALARY") && t.includes("1")) salary1Col = c;
    });
    if (salary1Col <= 0) return 0;
    const totalRow = findRowByLabel(ws, (t) => t.includes("TOTAL"));
    if (totalRow > 0) return num(ws.getRow(totalRow).getCell(salary1Col));
    return 0;
  };

  // Staff GROSS SALARY TOTAL (Column R typical)
  const readStaffGrossTotal = (ws: ExcelJS.Worksheet): number => {
    // Try to find GROSS SALARY column from header rows; fallback to R(18)
    let { col } = findColByHeader(ws, ["GROSS", "SALARY"]);
    if (col < 0) col = 18; // R [file:2]
    const totalRow = findRowByLabel(ws, (t) => t.includes("TOTAL"));
    if (totalRow > 0) return num(ws.getRow(totalRow).getCell(col));
    return 0;
  };

  // Bonus file helpers (B and C lanes)
  const getBonusMonthlyTotals = (bonusWb: ExcelJS.Workbook) => {
    // F..Q are months (Nov..Oct) in Bonus Worker/Staff sheets [file:5]
    const map: Record<string, number> = {
      "Nov-24": 6,
      "Dec-24": 7,
      "Jan-25": 8,
      "Feb-25": 9,
      "Mar-25": 10,
      "Apr-25": 11,
      "May-25": 12,
      "Jun-25": 13,
      "Jul-25": 14,
      "Aug-25": 15,
      "Sep-25": 16,
      "Oct-25": 17,
    };
    const out: Record<string, { worker: number; staff: number }> = {};
    const w = bonusWb.getWorksheet("Worker");
    const s = bonusWb.getWorksheet("Staff");
    const findGT = (ws?: ExcelJS.Worksheet) =>
      ws
        ? findRowByLabel(ws, (t) => t.includes("GRAND") && t.includes("TOTAL"))
        : -1;
    const wGT = findGT(w),
      sGT = findGT(s);
    for (const m of Object.keys(map)) {
      const col = map[m];
      out[m] = {
        worker: wGT > 0 && w ? num(w.getRow(wGT).getCell(col)) : 0,
        staff: sGT > 0 && s ? num(s.getRow(sGT).getCell(col)) : 0,
      };
    }
    return out;
  };

  const getBonusGrossTotals = (
    bonusWb: ExcelJS.Workbook
  ): { worker: number; staff: number } => {
    const out: { worker: number; staff: number } = { worker: 0, staff: 0 };
    const readGross = (ws?: ExcelJS.Worksheet) => {
      if (!ws) return 0;
      const gt = findRowByLabel(
        ws,
        (t) => t.includes("GRAND") && t.includes("TOTAL")
      );
      if (gt < 0) return 0;
      // Column R (18) = Gross in Bonus sheets [file:5]
      return num(ws.getRow(gt).getCell(18));
    };
    out.worker = readGross(bonusWb.getWorksheet("Worker"));
    out.staff = readGross(bonusWb.getWorksheet("Staff"));
    return out;
  };

  const processExcelFiles = async () => {
    if (!staffFile || !workerFile) {
      alert("Please upload both Indiana Staff and Indiana Worker files");
      return;
    }
    setIsGenerating(true);

    try {
      // Load workbooks
      const staffWb = new ExcelJS.Workbook();
      const workerWb = new ExcelJS.Workbook();
      await staffWb.xlsx.load(await staffFile.arrayBuffer());
      await workerWb.xlsx.load(await workerFile.arrayBuffer());

      // Optional workbooks
      let monthWiseWb: ExcelJS.Workbook | null = null;
      if (monthWiseFile) {
        monthWiseWb = new ExcelJS.Workbook();
        await monthWiseWb.xlsx.load(await monthWiseFile.arrayBuffer());
      }

      let bonusWb: ExcelJS.Workbook | null = null;
      let bonusMonthly: Record<string, { worker: number; staff: number }> = {};
      let bonusGross: { worker: number; staff: number } = {
        worker: 0,
        staff: 0,
      };
      if (bonusFile) {
        bonusWb = new ExcelJS.Workbook();
        await bonusWb.xlsx.load(await bonusFile.arrayBuffer());
        bonusMonthly = getBonusMonthlyTotals(bonusWb);
        bonusGross = getBonusGrossTotals(bonusWb);
        console.log(
          "Bonus sheets loaded:",
          bonusWb.worksheets.map((ws) => ws.name)
        );
      } else {
        console.warn(
          "Bonus sheet not provided – B and C HR values will be 0 unless month-wise logic fills them."
        );
      }

      const months = generateMonthHeaders();

      // Sheet maps
      const staffMap: Record<string, string> = {
        "Nov-24": "NOV-24 O",
        "Dec-24": "DEC-24 O",
        "Jan-25": "JAN-25 O",
        "Feb-25": "FEB-25 O",
        "Mar-25": "MAR-25 O",
        "Apr-25": "APR-25 O",
        "May-25": "MAY-25 O",
        "Jun-25": "JUN-25 O",
        "Jul-25": "JULY-25 O",
        "Aug-25": "AUG-25 O",
        "Sep-25": "SEP-25 O",
        "Oct-25": "OCT-25 O",
      };
      const workerMap: Record<string, string> = {
        "Nov-24": "NOV-24 W",
        "Dec-24": "DEC-24 W",
        "Jan-25": "JAN-25 W",
        "Feb-25": "FEB-25 W",
        "Mar-25": "MAR-25 W",
        "Apr-25": "APR-25 W",
        "May-25": "MAY-25 W",
        "Jun-25": "JUN-25 W",
        "Jul-25": "JULY-25 W",
        "Aug-25": "AUG-25 W",
        "Sep-25": "SEP-25 W",
        "Oct-25": "OCT-25 W",
      };

      const software: Record<string, CellTriplet> = {};
      const hr: Record<string, CellTriplet> = {};
      const diff: Record<string, CellTriplet> = {};

      for (const m of months) {
        // A lane (existing)
        const staffSheet = staffWb.getWorksheet(staffMap[m]);
        const workerMonthSheet = (monthWiseWb || workerWb).getWorksheet(
          workerMap[m]
        );

        let A_v1 = 0;
        if (staffSheet) A_v1 = readStaffSalary1Total(staffSheet); // Software A from Staff Salary1 TOTAL [file:2]
        let A_v2 = 0;
        if (workerMonthSheet) {
          // WD SALARY column detection and total [file:1]
          const { col, headerRow } = findColByHeader(workerMonthSheet, [
            "WD",
            "SALARY",
          ]);
          A_v2 = readColumnGrandTotal(workerMonthSheet, col, headerRow);
        }
        const A_diff = A_v2 - A_v1;

        // B lane (existing)
        let B_v1_staff = 0,
          B_v1_worker = 0;
        if (staffSheet) B_v1_staff = readStaffSalary1Total(staffSheet);
        const workerSheetForMonth = workerWb.getWorksheet(workerMap[m]);
        if (workerSheetForMonth)
          B_v1_worker = sumWorkerSalary1(workerSheetForMonth); // sum Salary1 [file:3]
        const B_v1 = B_v1_staff + B_v1_worker;
        const B_v2 = bonusMonthly[m]
          ? bonusMonthly[m].worker + bonusMonthly[m].staff
          : 0; // Bonus Worker+Staff month [file:5]
        const B_diff = B_v2 - B_v1;

        // C lane (new): g1 vs g2
        // g1: Worker Salary1 sum + Staff GROSS SALARY TOTAL (R) from Staff file [file:2][file:3]
        let C_g1_worker = 0,
          C_g1_staffGross = 0;
        if (workerSheetForMonth)
          C_g1_worker = sumWorkerSalary1(workerSheetForMonth);
        if (staffSheet) C_g1_staffGross = readStaffGrossTotal(staffSheet);
        const C_v1 = C_g1_worker + C_g1_staffGross;

        // g2: Bonus file GROSS SALARY grand totals (Column R) Worker+Staff [file:5]
        let C_v2 = 0;
        const monthWiseMonth = (monthWiseWb || workerWb).getWorksheet(
          workerMap[m]
        ); // prefer Month‑Wise if loaded
        if (monthWiseMonth) C_v2 = getGrossSalaryGrandTotal(monthWiseMonth); // reads Column I GRAND TOTAL [file:1]
        const C_diff = C_v2 - C_v1;

        software[m] = { A: round(A_v1), B: round(B_v1), C: round(C_v1) };
        hr[m] = { A: round(A_v2), B: round(B_v2), C: round(C_v2) };
        diff[m] = { A: round(A_diff), B: round(B_diff), C: round(C_diff) };
      }

      setReportData({
        months,
        departments: [
          { name: "Software", data: software },
          { name: "HR", data: hr },
          { name: "diff", data: diff },
        ],
      });
    } catch (e) {
      console.error(e);
      alert("Error processing files. Please verify formats and sheet names.");
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadReport = async () => {
    if (!reportData) return;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Report");

    const hdr: (string | number)[] = [""];
    reportData.months.forEach((m) => hdr.push(m, "", ""));
    ws.addRow(hdr);
    reportData.months.forEach((_, idx) =>
      ws.mergeCells(1, 2 + idx * 3, 1, 4 + idx * 3)
    );

    const sub: (string | number)[] = [""];
    reportData.months.forEach(() => sub.push("A", "B", "C"));
    ws.addRow(sub);

    reportData.departments.forEach((d) => {
      const row: (string | number)[] = [d.name];
      reportData.months.forEach((m) => {
        row.push(d.data[m]?.A ?? 0, d.data[m]?.B ?? 0, d.data[m]?.C ?? 0);
      });
      ws.addRow(row);
    });

    ws.getRow(1).font = { bold: true };
    ws.getRow(2).font = { bold: true };
    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" },
    };

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `report_${new Date().toISOString().split("T")[0]}.xlsx`;
    a.click();
  };

  // UI
  const BonusCard = () => (
    <div
      className={`border-2 rounded-lg p-6 ${
        bonusFile
          ? "border-green-300 bg-green-50"
          : "border-gray-300 bg-gray-50"
      }`}
    >
      <div className="flex items-center gap-3 mb-4">
        {bonusFile ? (
          <svg
            className="w-8 h-8 text-green-600"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
            />
          </svg>
        ) : (
          <svg
            className="w-8 h-8 text-gray-400"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M12 9v2m0 4h.01m-6.938 4h13.856"
            />
          </svg>
        )}
        <div>
          <h3 className="text-lg font-bold text-gray-800">Bonus Sheet</h3>
          <span className="text-xs bg-gray-200 text-gray-700 px-2 py-0.5 rounded font-medium">
            Optional
          </span>
        </div>
      </div>
      {bonusFile ? (
        <div className="space-y-2">
          <div className="bg-white rounded p-3 border border-green-200">
            <p className="text-sm font-medium text-gray-800 truncate">
              {bonusFile.name}
            </p>
            <p className="text-xs text-gray-500 mt-1">
              Size: {(bonusFile.size / 1024).toFixed(2)} KB
            </p>
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
            File is cached and ready
          </div>
        </div>
      ) : (
        <div className="bg-white rounded p-3 border border-gray-200">
          <p className="text-sm text-gray-600 font-medium">No file uploaded</p>
          <p className="text-xs text-gray-500 mt-1">This file is optional</p>
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
                Step 2 - Generate Report
              </h1>
              <p className="text-gray-600 mt-2">
                Review cached files and generate your report
              </p>
            </div>
            <button
              onClick={() => router.push("/")}
              className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
            >
              ← Back to Step 1
            </button>
          </div>

          {/* File Status Cards */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
            {/* Indiana Staff */}
            <div
              className={`border-2 rounded-lg p-6 ${
                staffFile
                  ? "border-green-300 bg-green-50"
                  : "border-red-300 bg-red-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    staffFile ? "text-green-600" : "text-red-600"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d={
                      staffFile
                        ? "M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                        : "M6 18L18 6M6 6l12 12"
                    }
                  />
                </svg>
                <h3 className="text-lg font-bold text-gray-800">
                  Indiana Staff
                </h3>
              </div>
              {staffFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {staffFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(staffFile.size / 1024).toFixed(2)} KB
                    </p>
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
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-red-200">
                  <p className="text-sm text-red-700 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    Please go back to Step 1 and upload this file
                  </p>
                </div>
              )}
            </div>

            {/* Indiana Worker */}
            <div
              className={`border-2 rounded-lg p-6 ${
                workerFile
                  ? "border-green-300 bg-green-50"
                  : "border-red-300 bg-red-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    workerFile ? "text-green-600" : "text-red-600"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d={
                      workerFile
                        ? "M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                        : "M6 18L18 6M6 6l12 12"
                    }
                  />
                </svg>
                <h3 className="text-lg font-bold text-gray-800">
                  Indiana Worker
                </h3>
              </div>
              {workerFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {workerFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(workerFile.size / 1024).toFixed(2)} KB
                    </p>
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
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-red-200">
                  <p className="text-sm text-red-700 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    Please go back to Step 1 and upload this file
                  </p>
                </div>
              )}
            </div>

            {/* Month Wise */}
            <div
              className={`border-2 rounded-lg p-6 ${
                monthWiseFile
                  ? "border-green-300 bg-green-50"
                  : "border-gray-300 bg-gray-50"
              }`}
            >
              <div className="flex items-center gap-3 mb-4">
                <svg
                  className={`w-8 h-8 ${
                    monthWiseFile ? "text-green-600" : "text-gray-400"
                  }`}
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
                <div>
                  <h3 className="text-lg font-bold text-gray-800">
                    Month Wise Sheet
                  </h3>
                  <span className="text-xs bg-gray-200 text-gray-700 px-2 py-0.5 rounded font-medium">
                    Optional
                  </span>
                </div>
              </div>
              {monthWiseFile ? (
                <div className="space-y-2">
                  <div className="bg-white rounded p-3 border border-green-200">
                    <p className="text-sm font-medium text-gray-800 truncate">
                      {monthWiseFile.name}
                    </p>
                    <p className="text-xs text-gray-500 mt-1">
                      Size: {(monthWiseFile.size / 1024).toFixed(2)} KB
                    </p>
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
                    File is cached and ready
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded p-3 border border-gray-200">
                  <p className="text-sm text-gray-600 font-medium">
                    No file uploaded
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    This file is optional
                  </p>
                </div>
              )}
            </div>

            {/* Bonus */}
            {BonusCard()}
          </div>

          <div className="flex justify-center mb-8">
            <button
              onClick={processExcelFiles}
              disabled={!staffFile || !workerFile || isGenerating}
              className="px-8 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition disabled:bg-gray-400 disabled:cursor-not-allowed flex items-center gap-2"
            >
              {isGenerating ? "Generating Report..." : "Generate Report"}
            </button>
          </div>

          {reportData && (
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-2xl font-bold text-gray-800">
                  Report Preview
                </h2>
                <button
                  onClick={downloadReport}
                  className="px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition flex items-center gap-2"
                >
                  Download Excel
                </button>
              </div>
              <div className="bg-gray-900 rounded-lg p-6 overflow-x-auto">
                <table className="w-full border-collapse border border-gray-600">
                  <thead>
                    <tr className="bg-gray-800">
                      <th
                        className="border border-gray-600 px-4 py-3 text-white"
                        rowSpan={2}
                      ></th>
                      {reportData.months.map((m, i) => (
                        <th
                          key={i}
                          className="border border-gray-600 px-2 py-3 text-white"
                          colSpan={3}
                        >
                          {m}
                        </th>
                      ))}
                    </tr>
                    <tr className="bg-gray-800">
                      {reportData.months.map((_, i) => (
                        <React.Fragment key={i}>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            A
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            B
                          </th>
                          <th className="border border-gray-600 px-3 py-2 text-white text-sm">
                            C
                          </th>
                        </React.Fragment>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {reportData.departments.map((d, i) => (
                      <tr key={i} className="bg-gray-800 hover:bg-gray-700">
                        <td className="border border-gray-600 px-4 py-3 text-white font-semibold">
                          {d.name}
                        </td>
                        {reportData.months.map((m, j) => (
                          <React.Fragment key={j}>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.A ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.B ?? 0}
                            </td>
                            <td className="border border-gray-600 px-3 py-3 text-white text-center text-sm">
                              {d.data[m]?.C ?? 0}
                            </td>
                          </React.Fragment>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
