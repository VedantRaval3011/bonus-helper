import * as XLSX from "xlsx";

export interface ExcelValidationError {
  checkNumber: number;
  severity: "critical" | "high" | "medium";
  message: string;
  details: string;
  row?: number;
  column?: string;
  sheet?: string;
  fileOwner?: string; // NEW: the file to display this error under
  fileTypeOwner?: string;
}

export interface ValidationResult {
  fileName: string;
  fileType: string;
  status: "success" | "error";
  validationErrors: ExcelValidationError[];
  summary: {
    totalChecks: number;
    criticalIssues: number;
    highIssues: number;
    mediumIssues: number;
  };
}

interface EmployeeRecord {
  empId: string | number;
  empName: string;
  doj: Date | null;
  sheet: string;
  rowNum: number;
  sourceFileName: string; // NEW
  sourceFileType: string;
}

// Global storage for cross-file validation
class ValidationContext {
  private static allStaffRecords: Map<string, EmployeeRecord[]> = new Map();
  private static allWorkerRecords: Map<string, EmployeeRecord[]> = new Map();
  private static bonusRecords: Map<string, EmployeeRecord[]> = new Map();
  private static loanRecords: Map<string, EmployeeRecord[]> = new Map();
  private static dueVoucherRecords: Map<string, EmployeeRecord[]> = new Map();
  private static actualPercentageRecords: Map<string, EmployeeRecord[]> =
    new Map();

  static addStaffRecord(empId: string, record: EmployeeRecord) {
    if (!this.allStaffRecords.has(empId)) {
      this.allStaffRecords.set(empId, []);
    }
    this.allStaffRecords.get(empId)!.push(record);
  }

  static addWorkerRecord(empId: string, record: EmployeeRecord) {
    if (!this.allWorkerRecords.has(empId)) {
      this.allWorkerRecords.set(empId, []);
    }
    this.allWorkerRecords.get(empId)!.push(record);
  }

  static addBonusRecord(empId: string, record: EmployeeRecord) {
    if (!this.bonusRecords.has(empId)) {
      this.bonusRecords.set(empId, []);
    }
    this.bonusRecords.get(empId)!.push(record);
  }

  static addLoanRecord(empId: string, record: EmployeeRecord) {
    if (!this.loanRecords.has(empId)) {
      this.loanRecords.set(empId, []);
    }
    this.loanRecords.get(empId)!.push(record);
  }

  static addDueVoucherRecord(empId: string, record: EmployeeRecord) {
    if (!this.dueVoucherRecords.has(empId)) {
      this.dueVoucherRecords.set(empId, []);
    }
    this.dueVoucherRecords.get(empId)!.push(record);
  }
  static addActualPercentageRecord(empId: string, record: EmployeeRecord) {
    if (!this.actualPercentageRecords.has(empId)) {
      this.actualPercentageRecords.set(empId, []);
    }
    this.actualPercentageRecords.get(empId)!.push(record);
  }

  static getActualPercentageRecords() {
    return this.actualPercentageRecords;
  }

  static getAllStaffRecords() {
    return this.allStaffRecords;
  }

  static getAllWorkerRecords() {
    return this.allWorkerRecords;
  }

  static getBonusRecords() {
    return this.bonusRecords;
  }

  static getLoanRecords() {
    return this.loanRecords;
  }

  static getDueVoucherRecords() {
    return this.dueVoucherRecords;
  }

  static clear() {
    this.allStaffRecords.clear();
    this.allWorkerRecords.clear();
    this.bonusRecords.clear();
    this.loanRecords.clear();
    this.dueVoucherRecords.clear();
    this.actualPercentageRecords.clear();
  }
}

export class ExcelValidator {
  private workbook: XLSX.WorkBook;
  private validationErrors: ExcelValidationError[] = [];
  private fileName: string;
  private fileType: string;
  private allEmployeeRecords: Map<string, EmployeeRecord[]> = new Map();

  private isRowCompletelyEmpty(row: any, isStaff: boolean): boolean {
    const empId = this.getFieldValue(
      row,
      "EMP. ID",
      "EMP ID",
      "EMP.ID",
      "EMPID"
    );
    const empName = this.getFieldValue(
      row,
      "EMPLOYEE NAME",
      "EMP. NAME",
      "EMP NAME",
      "EMPNAME"
    );
    const dept = this.getFieldValue(row, "DEPT.", "DEPT");
    const salary = isStaff
      ? this.getFieldValue(row, "Salary", "SALARY1")
      : this.getFieldValue(row, "Salary1", "SALARY1");
    const workingDays = this.getFieldValue(row, "WD", "DAY");
    const basic = this.getFieldValue(row, "BASIC");
    const grossSalary = this.getFieldValue(row, "GROSS SALARY");

    // Check if all critical fields are empty or zero
    return (
      this.isEmptyValue(empId) &&
      this.isEmptyValue(empName) &&
      this.isEmptyValue(dept) &&
      (this.isEmptyValue(salary) || Number(salary) === 0) &&
      (this.isEmptyValue(workingDays) || Number(workingDays) === 0) &&
      (this.isEmptyValue(basic) || Number(basic) === 0) &&
      (this.isEmptyValue(grossSalary) || Number(grossSalary) === 0)
    );
  }

  constructor(buffer: Buffer, fileName: string) {
    this.workbook = XLSX.read(buffer);
    this.fileName = fileName;
    this.fileType = this.detectFileType(fileName);
  }

private detectFileType(fileName: string): string {
  const lowerName = fileName.toLowerCase();

  // First try filename-based detection (as fallback)
  if (lowerName.includes("actual-percentage") || lowerName.includes("actual percentage"))
    return "Actual-Percentage-Bonus-Data";
  if (lowerName.includes("bonus-summery") || lowerName.includes("bonus summery"))
    return "Bonus-Summery";
  if (lowerName.includes("bonus-final") || lowerName.includes("bonus final"))
    return "Bonus-Final-Calculation";
  if (lowerName.includes("indiana-staff") || lowerName.includes("indiana staff"))
    return "Indiana-Staff";
  if (lowerName.includes("indiana-worker") || lowerName.includes("indiana worker"))
    return "Indiana-Worker";
  if (lowerName.includes("month-wise") || lowerName.includes("month wise"))
    return "Month-Wise-Sheet";
  if (lowerName.includes("due-voucher") || lowerName.includes("due voucher"))
    return "Due-Voucher-List-Worker";
  if (lowerName.includes("loan-deduction") || lowerName.includes("loan deduction"))
    return "Loan-Deduction";

  // Content-based detection - analyze file structure
  return this.detectFileTypeByContent();
}

private detectFileTypeByContent(): string {
  // Get first sheet
  const firstSheetName = this.workbook.SheetNames[0];
  const sheet = this.workbook.Sheets[firstSheetName];
  
  if (!sheet) return "Unknown";

  // Extract headers from multiple rows to handle different header row positions
  const data = XLSX.utils.sheet_to_json(sheet, { defval: null, header: 1 }) as any[][];
  
  // Check first 5 rows for headers
  const headerRows = data.slice(0, 5);
  const allHeaders = new Set<string>();
  
  headerRows.forEach(row => {
    if (Array.isArray(row)) {
      row.forEach(cell => {
        if (cell && typeof cell === 'string') {
          allHeaders.add(cell.toString().trim().toUpperCase());
        }
      });
    }
  });

  // Convert to array for easier checking
  const headers = Array.from(allHeaders);

  // Helper function to check if headers contain specific patterns
  const hasHeaders = (...patterns: string[]) => {
    return patterns.every(pattern => 
      headers.some(h => h.includes(pattern.toUpperCase()))
    );
  };

  const hasAnyHeaders = (...patterns: string[]) => {
    return patterns.some(pattern => 
      headers.some(h => h.includes(pattern.toUpperCase()))
    );
  };

  // Detection logic based on unique column combinations

  // Indiana-Staff: Has "Salary" (not "Salary1"), "EMP. ID", "EMPLOYEE NAME"
  if (hasHeaders("EMP", "SALARY") && 
      !hasHeaders("SALARY1") && 
      hasAnyHeaders("EMPLOYEE NAME", "EMP. NAME")) {
    // Additional check for staff-specific fields
    if (hasAnyHeaders("BASIC", "D.A", "GROSS SALARY")) {
      return "Indiana-Staff";
    }
  }

  // Indiana-Worker: Has "Salary1", "EMP. ID", "DAY"/"WD"
  if (hasHeaders("SALARY1") || 
      (hasHeaders("EMP", "DAY") && hasAnyHeaders("EMPLOYEE NAME", "EMP NAME"))) {
    return "Indiana-Worker";
  }

  // Bonus-Final-Calculation: Has "EMP Code"/"EMP. CODE", "Register", "Due VC", "Final RTGS"
  if (hasHeaders("EMP", "REGISTER") && 
      (hasHeaders("DUE VC") || hasHeaders("FINAL RTGS"))) {
    return "Bonus-Final-Calculation";
  }

  // Bonus-Summery: Has "Percentage", "Category", but not detailed employee data
  if (hasHeaders("PERCENTAGE", "CATEGORY") && 
      !hasHeaders("EMP CODE") && 
      !hasHeaders("REGISTER")) {
    return "Bonus-Summery";
  }

  // Actual-Percentage-Bonus-Data: Has "Category", "Percentage" with potential EMP data
  if (hasHeaders("CATEGORY", "PERCENTAGE") && 
      (hasHeaders("EMP") || headers.length < 10)) {
    // Check if it's more detailed than Bonus-Summery
    if (hasAnyHeaders("EMP. ID", "EMP ID", "EMPLOYEE")) {
      return "Actual-Percentage-Bonus-Data";
    }
    return "Bonus-Summery";
  }

  // Month-Wise-Sheet: Has "EMP NAME", "ADJUSTMENT", monthly columns
  if (hasHeaders("EMP NAME", "ADJUSTMENT") || 
      (hasHeaders("EMP NAME") && hasAnyHeaders("JAN", "FEB", "MAR", "APR"))) {
    return "Month-Wise-Sheet";
  }

  // Due-Voucher-List-Worker: Has "Worker ID", "Worker Name", "Due Amount"
  if (hasHeaders("WORKER") && hasAnyHeaders("DUE AMOUNT", "AMOUNT")) {
    return "Due-Voucher-List-Worker";
  }

  // Loan-Deduction: Has "Loan Amount"/"Loan", "Employee ID"
  if (hasAnyHeaders("LOAN AMOUNT", "LOAN") && 
      hasAnyHeaders("EMPLOYEE ID", "EMP ID")) {
    return "Loan-Deduction";
  }

  return "Unknown";
}


  private addError(
    checkNumber: number,
    severity: "critical" | "high" | "medium",
    message: string,
    details: string,
    row?: number,
    column?: string,
    sheet?: string
  ) {
    this.validationErrors.push({
      checkNumber,
      severity,
      message,
      details,
      row,
      column,
      sheet,
    });
  }

  private getSheetData(sheetName?: string): any[] {
    const sheet = sheetName
      ? this.workbook.Sheets[sheetName]
      : this.workbook.Sheets[this.workbook.SheetNames[0]];

    if (!sheet) return [];

    // Determine header row based on file type
    let headerRow = 1; // Default: row 2 (0-indexed)

    if (this.fileType === "Indiana-Staff") {
      headerRow = 2; // Row 3 (0-indexed)
    }

    const data = XLSX.utils.sheet_to_json(sheet, {
      defval: null,
      header: 1,
    }) as any[][];

    // Get headers
    const headers = data[headerRow] || [];

    // Convert to object array
    const result: any[] = [];
    for (let i = headerRow + 1; i < data.length; i++) {
      const row: any = {};
      const dataRow = data[i];

      // Skip completely empty rows
      if (
        !dataRow ||
        dataRow.every(
          (cell) => cell === null || cell === undefined || cell === ""
        )
      ) {
        continue;
      }

      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        if (header) {
          row[header] = dataRow[j];
        }
      }
      result.push(row);
    }

    return result;
  }

  private getFieldValue(row: any, ...possibleNames: string[]): any {
    for (const name of possibleNames) {
      // Direct match
      if (row[name] !== undefined && row[name] !== null) return row[name];

      // Trimmed match
      const trimmedName = name.trim();
      if (row[trimmedName] !== undefined && row[trimmedName] !== null)
        return row[trimmedName];

      // Normalized match (remove extra spaces)
      for (const key of Object.keys(row)) {
        const normalizedKey = key.replace(/\s+/g, " ").trim();
        const normalizedName = name.replace(/\s+/g, " ").trim();

        if (normalizedKey === normalizedName) {
          return row[key];
        }

        // Also try case-insensitive match
        if (normalizedKey.toLowerCase() === normalizedName.toLowerCase()) {
          return row[key];
        }
      }
    }
    return undefined;
  }

  private parseDate(dateValue: any): Date | null {
    if (!dateValue) return null;

    // Excel numeric date
    if (typeof dateValue === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      return new Date(excelEpoch.getTime() + dateValue * 86400000);
    }

    // Already a Date object
    if (dateValue instanceof Date) {
      return dateValue;
    }

    // String date
    if (typeof dateValue === "string") {
      // Try DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY format
      const parts = dateValue.split(/[./-]/);
      if (parts.length === 3) {
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1;
        let year = parseInt(parts[2]);

        if (year < 100) {
          year += year < 50 ? 2000 : 1900;
        }

        if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
          return new Date(year, month, day);
        }
      }
    }

    const parsed = new Date(dateValue);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  private isValidDate(date: Date | null): boolean {
    return date !== null && !isNaN(date.getTime());
  }

  private isEmptyValue(value: any): boolean {
    return (
      value === null ||
      value === undefined ||
      value === "" ||
      (typeof value === "string" && value.trim() === "")
    );
  }

  private normalizeEmployeeName(name: string): string {
    if (!name) return "";
    return name.toString().trim().toUpperCase().replace(/\s+/g, " ");
  }

  private normalizeEmployeeId(id: any): string {
    if (this.isEmptyValue(id)) return "";
    return String(id).trim();
  }

  validateStaffWorkerTulsi(): void {
    const sheets = this.workbook.SheetNames;
    const currentDate = new Date();
    const isStaff = this.fileType === "Indiana-Staff";

    sheets.forEach((sheetName) => {
      const data = this.getSheetData(sheetName);
      const empIdsInSheet = new Set<string>();
      const empNamesInSheet: Map<
        string,
        { empId: string; doj: Date | null; rowNum: number }[]
      > = new Map();

      const actualRowNum = isStaff ? 4 : 3;

      data.forEach((row: any, index: number) => {
        const rowNum = actualRowNum + index;

        // ⭐ FIRST: Check if row is completely empty - skip immediately
        if (this.isRowCompletelyEmpty(row, isStaff)) {
          return; // Skip empty rows completely
        }

        // Now extract all field values
        const empId = this.getFieldValue(
          row,
          "EMP. ID",
          "EMP ID",
          "EMP.ID",
          "EMPID"
        );
        const empName = this.getFieldValue(
          row,
          "EMPLOYEE NAME",
          "EMP. NAME",
          "EMP NAME",
          "EMPNAME"
        );
        const doj = this.getFieldValue(row, "DOJ");
        const dept = this.getFieldValue(row, "DEPT.", "DEPT");

        const salary = isStaff
          ? this.getFieldValue(row, "Salary", "SALARY1")
          : this.getFieldValue(row, "Salary1", "SALARY1");
        const grossSalary = this.getFieldValue(row, "GROSS SALARY");
        const finalCheque = this.getFieldValue(row, "FINAL CHEQUE");

        const pf = this.getFieldValue(
          row,
          "PF-12%",
          "PF 12%",
          "PF-12 %",
          "PF 12 %",
          "BASIC+   DA",
          "PF"
        );
        const esic = this.getFieldValue(
          row,
          "ESIC   0.75%",
          "ESIC 0.75%",
          "ESIC 0.75 %",
          "ESIC   0.75 %",
          "ESIC"
        );
        const pt = this.getFieldValue(row, "PT", "PROF TAX");

        const workingDays = this.getFieldValue(row, "WD", "DAY");
        const leaveDays = this.getFieldValue(row, "LD", "LATE");

        const basic = this.getFieldValue(row, "BASIC");
        const da = this.getFieldValue(row, "D.A.", "DA");

        // Store employee record for cross-validation (only if has ID and name)
        const normalizedEmpId = this.normalizeEmployeeId(empId);
        const normalizedName = (empName ?? "").toString().trim(); // at least trim for display and matching
        if (!this.isEmptyValue(empId) && !this.isEmptyValue(empName)) {
          const record: EmployeeRecord = {
            empId: normalizedEmpId,
            empName: normalizedName,
            doj: this.parseDate(doj),
            sheet: sheetName,
            rowNum,
            sourceFileName: this.fileName,
            sourceFileType: this.fileType,
          };

          if (!this.allEmployeeRecords.has(normalizedName)) {
            this.allEmployeeRecords.set(normalizedName, []);
          }
          this.allEmployeeRecords.get(normalizedName)!.push(record);

          if (!empNamesInSheet.has(normalizedName)) {
            empNamesInSheet.set(normalizedName, []);
          }
          empNamesInSheet
            .get(normalizedName)!
            .push({ empId: normalizedEmpId, doj: this.parseDate(doj), rowNum });

          if (isStaff) {
            ValidationContext.addStaffRecord(normalizedEmpId, record);
          } else {
            ValidationContext.addWorkerRecord(normalizedEmpId, record);
          }
        }

        // Apply empty row check for all validations
        const hasBasicData =
          !this.isEmptyValue(empId) && !this.isEmptyValue(empName);

        // Check 4: Empty Employee ID
        if (hasBasicData && this.isEmptyValue(empId)) {
          this.addError(
            5,
            "critical",
            "Missing Employee ID",
            `Row ${rowNum} has no Employee ID`,
            rowNum,
            "EMP. ID",
            sheetName
          );
        }

        // Check 5: Empty Employee Name
        if (hasBasicData && this.isEmptyValue(empName)) {
          this.addError(
            6,
            "critical",
            "Missing Employee Name",
            `Employee ID ${empId} has no name specified`,
            rowNum,
            "EMPLOYEE NAME",
            sheetName
          );
        }

        // Check 6: Missing DOJ

        if (
          hasBasicData &&
          this.isEmptyValue(doj) &&
          normalizedEmpId.toUpperCase() !== "N"
        ) {
          this.addError(
            7,
            "critical",
            "Missing DOJ",
            `Employee ${empName} (${empId}) has no DOJ`,
            rowNum,
            "DOJ",
            sheetName
          );
        } else if (hasBasicData && normalizedEmpId.toUpperCase() !== "N") {
          const dojDate = this.parseDate(doj);

          if (!this.isValidDate(dojDate)) {
            this.addError(
              20,
              "medium",
              "Invalid DOJ format",
              `Employee ${empName} (${empId}) has invalid DOJ: ${doj}`,
              rowNum,
              "DOJ",
              sheetName
            );
          } else {
            if (dojDate && dojDate > currentDate) {
              this.addError(
                1,
                "critical",
                "Future joining date",
                `Employee ${empName} (${empId}) has DOJ ${dojDate.toLocaleDateString()} which is in the future`,
                rowNum,
                "DOJ",
                sheetName
              );
            }
          }
        }

        // Check 3: Empty Department
        if (hasBasicData && this.isEmptyValue(dept)) {
          this.addError(
            4,
            "critical",
            "Missing Department",
            `Employee ${empName} (${empId}) has no department specified`,
            rowNum,
            "DEPT",
            sheetName
          );
        }

        // Check 7-8: Negative salary values
        if (hasBasicData && !this.isEmptyValue(salary) && Number(salary) < 0) {
          this.addError(
            8,
            "critical",
            "Negative Salary",
            `Employee ${empName} (${empId}) has negative SALARY: ${salary}`,
            rowNum,
            "Salary",
            sheetName
          );
        }

        if (
          hasBasicData &&
          !this.isEmptyValue(grossSalary) &&
          Number(grossSalary) < 0
        ) {
          this.addError(
            8,
            "critical",
            "Negative Gross Salary",
            `Employee ${empName} (${empId}) has negative GROSS SALARY: ${grossSalary}`,
            rowNum,
            "GROSS SALARY",
            sheetName
          );
        }

        // Check 9: Zero Salary but Working Days > 0
        if (
          hasBasicData &&
          !this.isEmptyValue(workingDays) &&
          Number(workingDays) > 0 &&
          (this.isEmptyValue(salary) || Number(salary) === 0)
        ) {
          this.addError(
            9,
            "medium",
            "Zero salary with working days",
            `Employee ${empName} (${empId}) has ${workingDays} working days but zero salary`,
            rowNum,
            "Salary",
            sheetName
          );
        }

        // Check 10: Working Days > Total days
        const monthDays = 31;
        if (
          hasBasicData &&
          !this.isEmptyValue(workingDays) &&
          Number(workingDays) > monthDays &&
          normalizedEmpId.toUpperCase() !== "N"
        ) {
          this.addError(
            10,
            "critical",
            "Working Days exceed month days",
            `Employee ${empName} (${empId}): Working Days (${workingDays}) > ${monthDays}`,
            rowNum,
            "WD/DAY",
            sheetName
          );
        }

        // Check 11: Negative working days
        if (
          hasBasicData &&
          !this.isEmptyValue(workingDays) &&
          Number(workingDays) < 0
        ) {
          this.addError(
            11,
            "critical",
            "Negative Working Days",
            `Employee ${empName} (${empId}) has negative working days: ${workingDays}`,
            rowNum,
            "WD/DAY",
            sheetName
          );
        }

        // Check 12: LD + WD > Total days
        // if (
        //   hasBasicData &&
        //   !this.isEmptyValue(workingDays) &&
        //   !this.isEmptyValue(leaveDays) &&
        //   Number(workingDays) + Number(leaveDays) > monthDays
        // ) {
        //   this.addError(
        //     12,
        //     "medium",
        //     "Leave + Working Days exceed month",
        //     `Employee ${empName} (${empId}): WD (${workingDays}) + LD (${leaveDays}) > ${monthDays}`,
        //     rowNum,
        //     "WD/LD",
        //     sheetName
        //   );
        // }

        // Check 15: PF Deduction > Gross Salary
        if (
          hasBasicData &&
          !this.isEmptyValue(pf) &&
          !this.isEmptyValue(grossSalary) &&
          Number(pf) > Number(grossSalary)
        ) {
          this.addError(
            15,
            "critical",
            "PF exceeds Gross Salary",
            `Employee ${empName} (${empId}): PF (${pf}) > Gross Salary (${grossSalary})`,
            rowNum,
            "PF",
            sheetName
          );
        }

        // Check 16: Negative deductions
        if (hasBasicData && !this.isEmptyValue(pf) && Number(pf) < 0) {
          this.addError(
            16,
            "critical",
            "Negative PF Deduction",
            `Employee ${empName} (${empId}) has negative PF: ${pf}`,
            rowNum,
            "PF",
            sheetName
          );
        }

        if (hasBasicData && !this.isEmptyValue(esic) && Number(esic) < 0) {
          this.addError(
            16,
            "critical",
            "Negative ESIC Deduction",
            `Employee ${empName} (${empId}) has negative ESIC: ${esic}`,
            rowNum,
            "ESIC",
            sheetName
          );
        }

        if (hasBasicData && !this.isEmptyValue(pt) && Number(pt) < 0) {
          this.addError(
            16,
            "critical",
            "Negative PT Deduction",
            `Employee ${empName} (${empId}) has negative PT: ${pt}`,
            rowNum,
            "PT",
            sheetName
          );
        }

        // Check 17: PF ≠ 12% of (BASIC + DA) - Already has the correct check
        if (
          !this.isEmptyValue(empId) &&
          !this.isEmptyValue(empName) &&
          !this.isEmptyValue(basic) &&
          Number(basic) > 0 &&
          !this.isEmptyValue(da) &&
          Number(da) > 0 &&
          !this.isEmptyValue(pf) &&
          Number(pf) > 0
        ) {
          const expectedPF = (Number(basic) + Number(da)) * 0.12;
          const actualPF = Number(pf);
          const tolerance = 2;

          if (
            Math.abs(expectedPF - actualPF) > tolerance &&
            actualPF !== 1800
          ) {
            this.addError(
              17,
              "medium",
              "PF Calculation Error",
              `Employee ${empName} (${empId}): Expected PF ${expectedPF.toFixed(
                2
              )}, Got ${actualPF}`,
              rowNum,
              "PF",
              sheetName
            );
          }
        }

        // Check 19: FINAL CHEQUE > GROSS SALARY
        if (
          hasBasicData &&
          !this.isEmptyValue(finalCheque) &&
          !this.isEmptyValue(grossSalary) &&
          Number(finalCheque) > Number(grossSalary) * 1.1
        ) {
          this.addError(
            19,
            "critical",
            "Final Cheque significantly exceeds Gross Salary",
            `Employee ${empName} (${empId}): Final Cheque (${finalCheque}) > Gross Salary (${grossSalary})`,
            rowNum,
            "FINAL CHEQUE",
            sheetName
          );
        }

        // Check 21: Duplicate EMP ID in same sheet
        // Check 21: Duplicate EMP ID in same sheet
        if (hasBasicData && !this.isEmptyValue(empId)) {
          const empIdStr = this.normalizeEmployeeId(empId);
          // Skip duplicate check for Employee ID "N"
          if (empIdStr.toUpperCase() !== "N" && empIdsInSheet.has(empIdStr)) {
            this.addError(
              21,
              "critical",
              "Duplicate Employee ID in same sheet",
              `Employee ID ${empId} appears multiple times in ${sheetName}`,
              rowNum,
              "EMP. ID",
              sheetName
            );
          }
          // Only add to set if not "N" to allow multiple "N" entries
          if (empIdStr.toUpperCase() !== "N") {
            empIdsInSheet.add(empIdStr);
          }
        }

        // Check 22: Non-numeric values in numeric fields
        const numericFields = [
          { field: salary, name: "Salary" },
          { field: grossSalary, name: "Gross Salary" },
          { field: pf, name: "PF" },
          { field: esic, name: "ESIC" },
          { field: workingDays, name: "Working Days" },
        ];

        if (hasBasicData) {
          numericFields.forEach(({ field, name }) => {
            if (!this.isEmptyValue(field) && isNaN(Number(field))) {
              this.addError(
                22,
                "critical",
                "Invalid numeric value",
                `Employee ${empName} (${empId}): ${name} has non-numeric value: ${field}`,
                rowNum,
                name,
                sheetName
              );
            }
          });
        }
      });

      // Within-sheet check: Same name with different EMP IDs
      empNamesInSheet.forEach((records, empName) => {
        if (records.length > 1) {
          const uniqueIds = new Set(records.map((r) => r.empId));
          if (uniqueIds.size > 1) {
            const idList = Array.from(uniqueIds).join(", ");
            this.addError(
              3,
              "critical",
              "Same Name + Different EMP ID in same sheet",
              `Employee "${empName}" appears with different Employee IDs in ${sheetName}: ${idList}`,
              records[0].rowNum,
              "EMP. ID",
              sheetName
            );
          }
        }
      });
    });

    // PHASE 2: Cross-sheet validation remains unchanged
    this.allEmployeeRecords.forEach((records, empName) => {
      if (records.length > 1) {
        const dojGroups: Map<string, EmployeeRecord[]> = new Map();

        records.forEach((record) => {
          const dojKey = record.doj ? record.doj.toDateString() : "NO_DOJ";
          if (!dojGroups.has(dojKey)) {
            dojGroups.set(dojKey, []);
          }
          dojGroups.get(dojKey)!.push(record);
        });

        dojGroups.forEach((groupRecords, dojKey) => {
          if (groupRecords.length > 1) {
            const uniqueEmpIds = new Set(
              groupRecords.map((r) => this.normalizeEmployeeId(r.empId))
            );

            if (uniqueEmpIds.size > 1) {
              const idList = Array.from(uniqueEmpIds).join(", ");
              const sheetList = groupRecords
                .map((r) => `${r.sheet} (ID: ${r.empId})`)
                .join(", ");

              this.addError(
                2,
                "critical",
                "Same DOJ + Same Name + Different EMP ID",
                `Employee "${empName}" with same DOJ (${
                  dojKey !== "NO_DOJ" ? dojKey : "Missing"
                }) has DIFFERENT Employee IDs: ${idList}. Found in: ${sheetList}`,
                groupRecords[0].rowNum,
                "EMP. ID",
                groupRecords[0].sheet
              );
            }
          }
        });
      }
    });
  }

  validateBonusCalculation(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const empCode = this.getFieldValue(
        row,
        "EMP Code",
        "EMP ID",
        "EMP. CODE",
        "EMP. ID",
        "EMPID" // Add more variations
      );
      const empName = this.getFieldValue(
        row,
        "EMP. NAME",
        "EMP NAME",
        "EMPLOYEE NAME",
        "Employee Name" // Add more variations
      );
      const dept = this.getFieldValue(row, "Deptt.", "DEPT", "Department");
      const bonusPercent = this.getFieldValue(
        row,
        "%",
        "Percentage",
        "Bonus %"
      );
      const register = this.getFieldValue(row, "Register");
      const dueVC = this.getFieldValue(row, "Due VC");
      const finalRTGS = this.getFieldValue(row, "Final RTGS");

      // Skip if completely empty
      if (this.isEmptyValue(empCode) && this.isEmptyValue(empName)) return;

      // Store in global context
      if (!this.isEmptyValue(empCode)) {
        const normalizedEmpId = this.normalizeEmployeeId(empCode);
        ValidationContext.addBonusRecord(normalizedEmpId, {
          empId: normalizedEmpId, // use normalized, not raw
          empName: (empName ?? "").toString().trim(), // trim name; may still be empty
          doj: null,
          sheet: "Bonus",
          rowNum,
          sourceFileName: this.fileName,
          sourceFileType: this.fileType,
        });
      }

      const hasBasicData =
        !this.isEmptyValue(empCode) && !this.isEmptyValue(empName);

      // Check 23: Missing Employee Code
      if (hasBasicData && this.isEmptyValue(empCode)) {
        this.addError(
          23,
          "critical",
          "Missing Employee Code",
          `Row ${rowNum} in Bonus sheet has no Employee Code`,
          rowNum,
          "EMP Code"
        );
      }

      // Check 24: Missing Employee Name
      if (hasBasicData && this.isEmptyValue(empName)) {
        this.addError(
          24,
          "critical",
          "Missing Employee Name",
          `Employee Code ${empCode} has no name in Bonus sheet`,
          rowNum,
          "EMP. NAME"
        );
      }

      // Check 25: Missing Department
      if (hasBasicData && this.isEmptyValue(dept)) {
        this.addError(
          25,
          "medium",
          "Missing Department",
          `Employee ${empName} (${empCode}) has no department in Bonus sheet`,
          rowNum,
          "Deptt."
        );
      }

      // Check 26: Invalid Bonus Percentage
      if (hasBasicData && !this.isEmptyValue(bonusPercent)) {
        const percent = Number(bonusPercent);
        if (percent < 0 || percent > 100) {
          this.addError(
            26,
            "critical",
            "Invalid Bonus Percentage",
            `Employee ${empName} (${empCode}): Bonus % is ${percent}% (should be 0-100%)`,
            rowNum,
            "%"
          );
        }
      }

      // Check 27: Negative Monthly Salary
      const monthlyFields = Object.keys(row).filter(
        (key) =>
          key.toLowerCase().includes("2024") ||
          key.toLowerCase().includes("2025") ||
          key.toLowerCase().includes("jan") ||
          key.toLowerCase().includes("feb") ||
          key.toLowerCase().includes("mar") ||
          key.toLowerCase().includes("apr") ||
          key.toLowerCase().includes("may") ||
          key.toLowerCase().includes("jun") ||
          key.toLowerCase().includes("jul") ||
          key.toLowerCase().includes("aug") ||
          key.toLowerCase().includes("sep") ||
          key.toLowerCase().includes("oct") ||
          key.toLowerCase().includes("nov") ||
          key.toLowerCase().includes("dec")
      );

      if (hasBasicData) {
        monthlyFields.forEach((field) => {
          const value = row[field];
          if (!this.isEmptyValue(value) && Number(value) < 0) {
            this.addError(
              27,
              "critical",
              "Negative Monthly Salary",
              `Employee ${empName} (${empCode}): Negative salary in ${field}: ${value}`,
              rowNum,
              field
            );
          }
        });
      }

      // Check 33: Payment Split Error
      if (
        hasBasicData &&
        !this.isEmptyValue(register) &&
        !this.isEmptyValue(dueVC) &&
        !this.isEmptyValue(finalRTGS)
      ) {
        const sum = Number(dueVC || 0) + Number(finalRTGS);
        const expectedRegister = Number(register);
        const tolerance = 1;

        if (Math.abs(sum - expectedRegister) > tolerance) {
          this.addError(
            33,
            "medium",
            "Payment Split Error",
            `Employee ${empName} (${empCode}): Due VC (${dueVC}) + Final RTGS (${finalRTGS}) ≠ Register (${register})`,
            rowNum,
            "Register"
          );
        }
      }
    });
  }

  validateMonthWiseSheet(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const empName = this.getFieldValue(
        row,
        "EMP NAME",
        "EMPLOYEE NAME",
        "EMP. NAME"
      );
      const adjustment = this.getFieldValue(
        row,
        "ADJUSTMENT",
        "ADJ.",
        "EXT. ADJ"
      );

      // Skip if completely empty
      if (this.isEmptyValue(empName)) return;

      const hasBasicData = !this.isEmptyValue(empName);

      // Check 35: Missing Employee Name
      if (hasBasicData && this.isEmptyValue(empName)) {
        this.addError(
          35,
          "medium",
          "Missing Employee Name",
          `Row ${rowNum} in Month-Wise sheet has no Employee Name`,
          rowNum,
          "EMP NAME"
        );
      }

      // Check 37: Large Negative Adjustment
      if (
        hasBasicData &&
        !this.isEmptyValue(adjustment) &&
        Number(adjustment) < -10000
      ) {
        this.addError(
          37,
          "medium",
          "Large Negative Adjustment",
          `Employee ${empName}: Large negative adjustment detected: ${adjustment}`,
          rowNum,
          "ADJUSTMENT"
        );
      }
    });
  }

  validateDueVoucher(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const workerId = this.getFieldValue(
        row,
        "Worker ID",
        "WORKER ID",
        "EMP. ID",
        "EMP ID"
      );
      const workerName = this.getFieldValue(
        row,
        "Worker Name",
        "WORKER NAME",
        "EMPLOYEE NAME"
      );
      const dueAmount = this.getFieldValue(
        row,
        "Due Amount",
        "DUE AMOUNT",
        "AMOUNT"
      );

      // Skip if completely empty
      if (
        this.isEmptyValue(workerId) &&
        this.isEmptyValue(workerName) &&
        this.isEmptyValue(dueAmount)
      )
        return;

      // Store in global context
      if (!this.isEmptyValue(workerId)) {
        const normalizedWorkerId = this.normalizeEmployeeId(workerId);
        ValidationContext.addDueVoucherRecord(normalizedWorkerId, {
          empId: normalizedWorkerId,
          empName: workerName,
          doj: null,
          sheet: "Due Voucher",
          rowNum: rowNum,
          sourceFileName: this.fileName,
          sourceFileType: this.fileType,
        });
      }

      const hasBasicData =
        !this.isEmptyValue(workerId) && !this.isEmptyValue(workerName);

      // Check 38: Missing Worker ID
      if (hasBasicData && this.isEmptyValue(workerId)) {
        this.addError(
          38,
          "critical",
          "Missing Worker ID",
          `Row ${rowNum} in Due Voucher has no Worker ID`,
          rowNum,
          "Worker ID"
        );
      }

      // Check 39: Missing Worker Name
      if (hasBasicData && this.isEmptyValue(workerName)) {
        this.addError(
          39,
          "critical",
          "Missing Worker Name",
          `Worker ID ${workerId} has no name in Due Voucher`,
          rowNum,
          "Worker Name"
        );
      }

      // Check 40: Negative Due Amount
      if (
        hasBasicData &&
        !this.isEmptyValue(dueAmount) &&
        Number(dueAmount) < 0
      ) {
        this.addError(
          40,
          "critical",
          "Negative Due Amount",
          `Worker ${workerName} (${workerId}) has negative due amount: ${dueAmount}`,
          rowNum,
          "Due Amount"
        );
      }
    });
  }

  validateLoanDeduction(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const empId = this.getFieldValue(
        row,
        "Employee ID",
        "EMP ID",
        "EMP. ID",
        "EMPID"
      );
      const empName = this.getFieldValue(
        row,
        "Employee Name",
        "EMP NAME",
        "EMPLOYEE NAME",
        "EMP. NAME"
      );
      const loanAmount = this.getFieldValue(
        row,
        "Loan Amount",
        "LOAN",
        "Loan",
        "AMOUNT"
      );

      // Skip if completely empty
      if (
        this.isEmptyValue(empId) &&
        this.isEmptyValue(empName) &&
        this.isEmptyValue(loanAmount)
      )
        return;

      // Store in global context
      if (
        !this.isEmptyValue(empId) &&
        !this.isEmptyValue(loanAmount) &&
        Number(loanAmount) > 0
      ) {
        const normalizedEmpId = this.normalizeEmployeeId(empId);
        ValidationContext.addLoanRecord(normalizedEmpId, {
          empId: normalizedEmpId,
          empName: empName,
          doj: null,
          sheet: "Loan",
          rowNum: rowNum,
          sourceFileName: this.fileName,
          sourceFileType: this.fileType,
        });
      }

      const hasBasicData =
        !this.isEmptyValue(empId) && !this.isEmptyValue(empName);

      // Check 42: Missing Employee ID
      if (
        hasBasicData &&
        !this.isEmptyValue(loanAmount) &&
        Number(loanAmount) > 0 &&
        this.isEmptyValue(empId)
      ) {
        this.addError(
          42,
          "critical",
          "Missing Employee ID",
          `Row ${rowNum} in Loan Deduction has loan but no Employee ID`,
          rowNum,
          "Employee ID"
        );
      }

      // Check 43: Missing Employee Name
      if (
        hasBasicData &&
        !this.isEmptyValue(loanAmount) &&
        Number(loanAmount) > 0 &&
        this.isEmptyValue(empName)
      ) {
        this.addError(
          43,
          "critical",
          "Missing Employee Name",
          `Row ${rowNum} in Loan Deduction has loan but no Employee Name`,
          rowNum,
          "Employee Name"
        );
      }

      // Check 44: Negative Loan Amount
      if (
        hasBasicData &&
        !this.isEmptyValue(loanAmount) &&
        Number(loanAmount) < 0
      ) {
        this.addError(
          44,
          "critical",
          "Negative Loan Amount",
          `Employee ${empName} (${empId}) has negative loan amount: ${loanAmount}`,
          rowNum,
          "Loan Amount"
        );
      }
    });
  }

  validateBonusSummery(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const percentage = this.getFieldValue(
        row,
        "Percentage",
        "%",
        "PERCENTAGE"
      );
      const category = this.getFieldValue(row, "Category", "CATEGORY", "TYPE");

      // Skip if completely empty
      if (this.isEmptyValue(category) && this.isEmptyValue(percentage)) return;

      const hasBasicData = !this.isEmptyValue(category);

      // Check 50: Missing Percentage
      if (
        hasBasicData &&
        !this.isEmptyValue(category) &&
        this.isEmptyValue(percentage)
      ) {
        this.addError(
          50,
          "high",
          "Missing Percentage",
          `Row ${rowNum} in Bonus Summery has no percentage`,
          rowNum,
          "Percentage"
        );
      }
    });
  }

  validateActualPercentageBonus(): void {
    const data = this.getSheetData();
    const actualRowNum = 3;

    data.forEach((row: any, index: number) => {
      const rowNum = actualRowNum + index;
      const category = this.getFieldValue(row, "Category", "CATEGORY", "TYPE");
      const percentage = this.getFieldValue(
        row,
        "Percentage",
        "%",
        "PERCENTAGE"
      );

      // ADD THIS: Extract employee ID if present
      const empId = this.getFieldValue(
        row,
        "EMP. ID",
        "EMP ID",
        "Employee ID",
        "EMPID"
      );
      const empName = this.getFieldValue(
        row,
        "EMP. NAME",
        "EMP NAME",
        "Employee Name",
        "EMPLOYEE NAME"
      );

      // Skip if completely empty
      if (this.isEmptyValue(category) && this.isEmptyValue(percentage)) return;

      // ADD THIS: Store employee record if ID exists
      if (!this.isEmptyValue(empId)) {
        const normalizedEmpId = this.normalizeEmployeeId(empId);
        ValidationContext.addActualPercentageRecord(normalizedEmpId, {
          empId: normalizedEmpId,
          empName: empName || "",
          doj: null,
          sheet: "Actual Percentage",
          rowNum: rowNum,
          sourceFileName: this.fileName,
          sourceFileType: this.fileType,
        });
      }

      const hasBasicData = !this.isEmptyValue(category);

      // Check 51: Missing Percentage for Category
      if (
        hasBasicData &&
        !this.isEmptyValue(category) &&
        this.isEmptyValue(percentage)
      ) {
        this.addError(
          51,
          "high",
          "Missing Percentage for Category",
          `Category ${category} has no percentage defined`,
          rowNum,
          "Percentage"
        );
      }
    });
  }

  // Cross-file validations - called after all files are validated
  // 3) Restrict cross-file rule #27 to "Bonus-Final-Calculation" and route to that file
  static performCrossFileValidation(
    validators: ExcelValidator[]
  ): ExcelValidationError[] {
    const crossFileErrors: ExcelValidationError[] = [];

    const staffRecords = ValidationContext.getAllStaffRecords();
    const workerRecords = ValidationContext.getAllWorkerRecords();
    const bonusRecords = ValidationContext.getBonusRecords();
    const loanRecords = ValidationContext.getLoanRecords();
    const dueVoucherRecords = ValidationContext.getDueVoucherRecords();

    // Combine all salary records
    const allSalaryRecords = new Map([...staffRecords, ...workerRecords]);

    const takeNonEmpty = (v?: unknown) =>
      typeof v === "string" ? v.trim() : v == null ? "" : String(v).trim();

    // Check 41
    dueVoucherRecords.forEach((records, workerId) => {
      if (!workerRecords.has(workerId)) {
        const src = records[0];
        const displayWorkerId = src.empId || workerId || "Unknown ID";
        const displayWorkerName = src.empName || "Unknown Name";

        crossFileErrors.push({
          checkNumber: 41,
          severity: "high",
          message: "Worker in Due Voucher but not in Indiana-Worker",
          details: `Worker ID ${displayWorkerId} (${displayWorkerName}) appears in Due Voucher but not found in Indiana-Worker`,
          row: src.rowNum,
          sheet: "Due Voucher",
          fileOwner: src.sourceFileName,
          fileTypeOwner: src.sourceFileType,
        });
      }
    });

    // Check 45
    loanRecords.forEach((records, empId) => {
      if (!allSalaryRecords.has(empId)) {
        const src = records[0];
        const displayEmpId = src.empId || empId || "Unknown ID";
        const displayEmpName = src.empName || "Unknown Name";

        crossFileErrors.push({
          checkNumber: 45,
          severity: "high",
          message: "Employee with Loan not in Salary sheets",
          details: `Employee ID ${displayEmpId} (${displayEmpName}) has loan but not found in Indiana-Staff and Indiana-Worker`,
          row: src.rowNum,
          sheet: "Loan",
          fileOwner: src.sourceFileName,
          fileTypeOwner: src.sourceFileType,
        });
      }
    });

    return crossFileErrors;
  }

  validate(): ValidationResult {
    switch (this.fileType) {
      case "Indiana-Staff":
      case "Indiana-Worker":
        this.validateStaffWorkerTulsi();
        break;
      case "Bonus-Final-Calculation":
        this.validateBonusCalculation();
        break;
      case "Month-Wise-Sheet":
        this.validateMonthWiseSheet();
        break;
      case "Due-Voucher-List-Worker":
        this.validateDueVoucher();
        break;
      case "Loan-Deduction":
        this.validateLoanDeduction();
        break;
      case "Bonus-Summery":
        this.validateBonusSummery();
        break;
      case "Actual-Percentage-Bonus-Data":
        this.validateActualPercentageBonus();
        break;
    }

    const summary = {
      totalChecks: this.validationErrors.length,
      criticalIssues: this.validationErrors.filter(
        (e) => e.severity === "critical"
      ).length,
      highIssues: this.validationErrors.filter((e) => e.severity === "high")
        .length,
      mediumIssues: this.validationErrors.filter((e) => e.severity === "medium")
        .length,
    };

    return {
      fileName: this.fileName,
      fileType: this.fileType,
      status: this.validationErrors.length > 0 ? "error" : "success",
      validationErrors: this.validationErrors,
      summary,
    };
  }
}
