// app/layout.tsx or wherever your root layout is
// Add this to wrap your children:
// <FileProvider>{children}</FileProvider>

'use client';

import React, { createContext, useContext, useState, ReactNode } from 'react';

interface FileSlot {
  type: string;
  label: string;
  file: File | null;
  required: boolean;
}

interface ValidationResults {
  results: any[];
  crossFileErrorCount: number;
}

interface FileContextType {
  fileSlots: FileSlot[];
  setFileSlots: (slots: FileSlot[]) => void;
  validationResults: ValidationResults | null;
  setValidationResults: (results: ValidationResults | null) => void;
  updateFileSlot: (index: number, file: File | null) => void;
  clearAllFiles: () => void;
}

const FileContext = createContext<FileContextType | undefined>(undefined);

const initialFileSlots: FileSlot[] = [
  { type: 'Indiana-Staff', label: 'Indiana Staff', file: null, required: true },
  { type: 'Indiana-Worker', label: 'Indiana Worker', file: null, required: true },
  { type: 'Bonus-Final-Calculation', label: 'Bonus Final Calculation', file: null, required: true },
  { type: 'Bonus-Summery', label: 'Bonus Summery', file: null, required: false },
  { type: 'Actual-Percentage-Bonus-Data', label: 'Actual Percentage Bonus Data', file: null, required: false },
  { type: 'Month-Wise-Sheet', label: 'Month Wise Sheet', file: null, required: false },
  { type: 'Due-Voucher-List-Worker', label: 'Due Voucher List (Worker)', file: null, required: false },
  { type: 'Loan-Deduction', label: 'Loan Deduction', file: null, required: false },
];

export function FileProvider({ children }: { children: ReactNode }) {
  const [fileSlots, setFileSlots] = useState<FileSlot[]>(initialFileSlots);
  const [validationResults, setValidationResults] = useState<ValidationResults | null>(null);

  const updateFileSlot = (index: number, file: File | null) => {
    setFileSlots(prev => {
      const newSlots = [...prev];
      newSlots[index].file = file;
      return newSlots;
    });
  };

  const clearAllFiles = () => {
    setFileSlots(initialFileSlots);
    setValidationResults(null);
  };

  return (
    <FileContext.Provider
      value={{
        fileSlots,
        setFileSlots,
        validationResults,
        setValidationResults,
        updateFileSlot,
        clearAllFiles,
      }}
    >
      {children}
    </FileContext.Provider>
  );
}

export function useFileContext() {
  const context = useContext(FileContext);
  if (context === undefined) {
    throw new Error('useFileContext must be used within a FileProvider');
  }
  return context;
}