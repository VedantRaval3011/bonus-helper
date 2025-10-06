import { NextRequest, NextResponse } from "next/server";
import connectDB from "@/lib/mongodb";
import ValidationMessage from "@/lib/models/ValidationMessage";
import {
  ExcelValidator,
  ValidationResult,
} from "@/lib/validators/excelValidators";

export async function POST(request: NextRequest) {
  try {
    await connectDB();

    const formData = await request.formData();
    const files = formData.getAll("files") as File[];

    if (!files || files.length === 0) {
      return NextResponse.json({ error: "No files uploaded" }, { status: 400 });
    }

    // Phase 1: Validate each file individually
    const validators: ExcelValidator[] = [];
    const results: ValidationResult[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());

      const validator = new ExcelValidator(buffer, file.name);
      const validationResult = validator.validate();

      validators.push(validator);
      results.push(validationResult);
    }

    // Phase 2: Perform cross-file validation
    const crossFileErrors =
      ExcelValidator.performCrossFileValidation(validators);

    // Add cross-file errors to their appropriate source files
    if (crossFileErrors.length > 0) {
      crossFileErrors.forEach((error) => {
        // Find the correct file to assign this error to
        let targetResult: ValidationResult | undefined;

        // Use fileOwner or fileTypeOwner to find the right result
        if (error.fileOwner) {
          targetResult = results.find((r) => r.fileName === error.fileOwner);
        }

        // Fallback to fileTypeOwner if fileOwner doesn't match
        if (!targetResult && error.fileTypeOwner) {
          targetResult = results.find(
            (r) => r.fileType === error.fileTypeOwner
          );
        }

        // If we found the target file, add the error to it
        if (targetResult) {
          targetResult.validationErrors.push(error);
          targetResult.summary.totalChecks++;

          // Update severity counts
          if (error.severity === "critical") {
            targetResult.summary.criticalIssues++;
          } else if (error.severity === "high") {
            targetResult.summary.highIssues++;
          } else if (error.severity === "medium") {
            targetResult.summary.mediumIssues++;
          }

          targetResult.status = "error";
        } else {
          // Fallback: If no matching file found, add to first result as before
          // This shouldn't happen if fileOwner/fileTypeOwner are set correctly
          console.warn(
            `Could not find target file for cross-file error:`,
            error
          );
          if (results.length > 0) {
            results[0].validationErrors.push(error);
            results[0].summary.totalChecks++;
            if (error.severity === "critical")
              results[0].summary.criticalIssues++;
            else if (error.severity === "high") results[0].summary.highIssues++;
            else if (error.severity === "medium")
              results[0].summary.mediumIssues++;
            results[0].status = "error";
          }
        }
      });
    }
    // Save all results to MongoDB
    const savedResults = [];
    for (const result of results) {
      const message = await ValidationMessage.create(result);
      savedResults.push({
        id: message._id,
        fileName: result.fileName,
        fileType: result.fileType,
        status: result.status,
        summary: result.summary,
        validationErrors: result.validationErrors, // Include full error details
      });
    }

    return NextResponse.json({
      success: true,
      message: "Files validated successfully",
      results: savedResults,
      crossFileErrorCount: crossFileErrors.length,
    });
  } catch (error: any) {
    console.error("Upload error:", error);
    return NextResponse.json(
      { error: "Failed to process files", details: error.message },
      { status: 500 }
    );
  }
}

export const config = {
  api: {
    bodyParser: false,
  },
};
