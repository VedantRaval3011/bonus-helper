// models/AuditMessage.ts
import mongoose, { Schema, Model, InferSchemaType } from 'mongoose';

const AuditMessageSchema = new Schema(
  {
    batchId: { type: String, required: true, index: true },
    step: { type: Number, required: true, min: 2, max: 9, index: true },
    level: { type: String, enum: ['error', 'warning', 'info'], default: 'error', index: true },
    tag: { type: String, default: 'mismatch', index: true }, // 'mismatch' | 'missing-in-hr' | 'metric-snapshot' ...
    text: { type: String, required: true },
    scope: { type: String, enum: ['staff', 'worker', 'global'], default: 'global', index: true },
    source: { type: String, default: 'step2', index: true },
    meta: { type: Schema.Types.Mixed },
  },
  { timestamps: { createdAt: true, updatedAt: false } }
);

AuditMessageSchema.index({ createdAt: -1 });
AuditMessageSchema.index({ batchId: 1, createdAt: -1 });

export type AuditMessageDoc = InferSchemaType<typeof AuditMessageSchema>;

export const AuditMessage: Model<AuditMessageDoc> =
  mongoose.models.AuditMessage ||
  mongoose.model<AuditMessageDoc>('AuditMessage', AuditMessageSchema);
