import mongoose, { Schema, Document } from 'mongoose';

// Define custom error interface with a different name to avoid conflict
export interface IValidationError {
  checkNumber: number;
  severity: 'critical' | 'high' | 'medium';
  message: string;
  details: string;
  row?: number;
  column?: string;
}

export interface IValidationSummary {
  totalChecks: number;
  criticalIssues: number;
  highIssues: number;
  mediumIssues: number;
}

export interface IValidationMessage extends Document {
  fileName: string;
  fileType: string;
  uploadDate: Date;
  status: 'success' | 'error';
  validationErrors: IValidationError[]; // Renamed from 'errors' to avoid conflict
  summary: IValidationSummary;
}

const ValidationErrorSchema = new Schema({
  checkNumber: { type: Number, required: true },
  severity: { 
    type: String, 
    enum: ['critical', 'high', 'medium'], 
    required: true 
  },
  message: { type: String, required: true },
  details: { type: String, required: true },
  row: { type: Number },
  column: { type: String }
}, { _id: false });

const ValidationSummarySchema = new Schema({
  totalChecks: { type: Number, required: true, default: 0 },
  criticalIssues: { type: Number, required: true, default: 0 },
  highIssues: { type: Number, required: true, default: 0 },
  mediumIssues: { type: Number, required: true, default: 0 }
}, { _id: false });

const ValidationMessageSchema: Schema = new Schema({
  fileName: { type: String, required: true },
  fileType: { type: String, required: true },
  uploadDate: { type: Date, default: Date.now },
  status: { 
    type: String, 
    enum: ['success', 'error'], 
    required: true 
  },
  validationErrors: [ValidationErrorSchema], // Renamed from 'errors'
  summary: { type: ValidationSummarySchema, required: true }
}, {
  timestamps: true
});

// Add indexes for better query performance
ValidationMessageSchema.index({ fileName: 1, uploadDate: -1 });
ValidationMessageSchema.index({ status: 1 });

export default mongoose.models.ValidationMessage || 
  mongoose.model<IValidationMessage>('ValidationMessage', ValidationMessageSchema);
