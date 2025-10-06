'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { useFileContext } from '@/contexts/FileContext';

export default function HomePage() {
  const router = useRouter();
  const { 
    fileSlots, 
    setFileSlots, 
    validationResults, 
    setValidationResults 
  } = useFileContext();
  
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string>('');
  const [expandedErrors, setExpandedErrors] = useState<Set<string>>(new Set());
  const [selectedResult, setSelectedResult] = useState<number | null>(null);
  const [password, setPassword] = useState<string>('');
  const [passwordError, setPasswordError] = useState<string>('');
  const [showPasswordModal, setShowPasswordModal] = useState<boolean>(false);

  const NEXT_PASSWORD = process.env.NEXT_PUBLIC_NEXT_PASSWORD || 'defaultPassword';

  const handleFileChange = (index: number, e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const newFileSlots = [...fileSlots];
      newFileSlots[index].file = file;
      setFileSlots(newFileSlots);
      setValidationResults(null);
      setError('');
      setExpandedErrors(new Set());
      setSelectedResult(null);
      setPassword('');
      setPasswordError('');
      setShowPasswordModal(false);
    }
  };

  const removeFile = (index: number) => {
    const newFileSlots = [...fileSlots];
    newFileSlots[index].file = null;
    setFileSlots(newFileSlots);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    const missingRequired = fileSlots.filter(slot => slot.required && !slot.file);
    if (missingRequired.length > 0) {
      setError(`Please upload required files: ${missingRequired.map(s => s.label).join(', ')}`);
      return;
    }

    const uploadedFiles = fileSlots.filter(slot => slot.file !== null);
    if (uploadedFiles.length === 0) {
      setError('Please upload at least one file');
      return;
    }

    setUploading(true);
    setError('');

    try {
      const formData = new FormData();
      
      uploadedFiles.forEach(slot => {
        if (slot.file) {
          const originalName = slot.file.name;
          const newFileName = `${slot.type}_${originalName}`;
          const renamedFile = new File([slot.file], newFileName, { type: slot.file.type });
          formData.append('files', renamedFile);
        }
      });

      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Upload failed');
      }

      setValidationResults(data);
      if (data.results && data.results.length > 0) {
        setSelectedResult(0);
      }
    } catch (err: any) {
      setError(err.message || 'Failed to upload files');
    } finally {
      setUploading(false);
    }
  };

  const toggleErrorExpansion = (errorKey: string) => {
    const newExpanded = new Set(expandedErrors);
    if (newExpanded.has(errorKey)) {
      newExpanded.delete(errorKey);
    } else {
      newExpanded.add(errorKey);
    }
    setExpandedErrors(newExpanded);
  };

  const getSeverityColor = (severity: string) => {
    switch (severity) {
      case 'critical':
        return {
          bg: 'bg-red-50',
          border: 'border-red-500',
          text: 'text-red-800',
          badge: 'bg-red-200 text-red-800'
        };
      case 'high':
        return {
          bg: 'bg-orange-50',
          border: 'border-orange-500',
          text: 'text-orange-800',
          badge: 'bg-orange-200 text-orange-800'
        };
      case 'medium':
        return {
          bg: 'bg-yellow-50',
          border: 'border-yellow-500',
          text: 'text-yellow-800',
          badge: 'bg-yellow-200 text-yellow-800'
        };
      default:
        return {
          bg: 'bg-gray-50',
          border: 'border-gray-500',
          text: 'text-gray-800',
          badge: 'bg-gray-200 text-gray-800'
        };
    }
  };

  const hasHighOrCriticalErrors = () => {
    if (!validationResults || !validationResults.results) return false;
    return validationResults.results.some((result: any) => 
      result.summary.criticalIssues > 0 || result.summary.highIssues > 0
    );
  };

  const handleNext = () => {
    if (!hasHighOrCriticalErrors()) {
      router.push('/step2');
    } else {
      setShowPasswordModal(true);
    }
  };

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === NEXT_PASSWORD) {
      setPasswordError('');
      setShowPasswordModal(false);
      router.push('/step2');
    } else {
      setPasswordError('Incorrect password');
    }
  };

  const getUploadedCount = () => fileSlots.filter(slot => slot.file !== null).length;
  const getRequiredCount = () => fileSlots.filter(slot => slot.required).length;
  const getUploadedRequiredCount = () => fileSlots.filter(slot => slot.required && slot.file !== null).length;

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-12 px-4">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <div>
              <h1 className="text-3xl font-bold text-gray-800">
                Excel Validation - Step 1
              </h1>
              <p className="text-sm text-gray-600 mt-2">
                Upload files individually for validation
              </p>
            </div>
            <button
              onClick={() => router.push('/messages')}
              className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
            >
              View All Messages
            </button>
          </div>

          <form onSubmit={handleSubmit} className="space-y-6">
            <div className="bg-gradient-to-r from-blue-50 to-indigo-50 border border-indigo-200 rounded-lg p-4">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-gray-700">
                    Files Uploaded: <span className="text-indigo-600 font-bold">{getUploadedCount()}</span> / {fileSlots.length}
                  </p>
                  <p className="text-xs text-gray-600 mt-1">
                    Required: <span className={getUploadedRequiredCount() === getRequiredCount() ? 'text-green-600 font-semibold' : 'text-orange-600 font-semibold'}>
                      {getUploadedRequiredCount()}
                    </span> / {getRequiredCount()}
                  </p>
                </div>
                <div className="flex gap-3">
                  <div className="flex items-center gap-1">
                    <span className="w-2 h-2 bg-red-500 rounded-full"></span>
                    <span className="text-xs text-gray-600">Required</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <span className="w-2 h-2 bg-gray-400 rounded-full"></span>
                    <span className="text-xs text-gray-600">Optional</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {fileSlots.map((slot, index) => (
                <div
                  key={index}
                  className={`border-2 rounded-lg p-4 transition ${
                    slot.required
                      ? 'border-red-200 bg-red-50/30'
                      : 'border-gray-200 bg-gray-50/30'
                  } ${slot.file ? 'ring-2 ring-green-400' : ''}`}
                >
                  <div className="flex items-start justify-between mb-3">
                    <div className="flex items-center gap-2">
                      <span className={`w-2 h-2 rounded-full ${slot.required ? 'bg-red-500' : 'bg-gray-400'}`}></span>
                      <label className="text-sm font-semibold text-gray-800">
                        {slot.label}
                      </label>
                    </div>
                    {slot.required && (
                      <span className="text-xs bg-red-100 text-red-700 px-2 py-0.5 rounded font-medium">
                        Required
                      </span>
                    )}
                  </div>

                  {!slot.file ? (
                    <div className="relative">
                      <input
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => handleFileChange(index, e)}
                        className="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-white focus:outline-none p-2 file:mr-4 file:py-1 file:px-3 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100"
                        id={`file-${index}`}
                      />
                    </div>
                  ) : (
                    <div className="bg-white border border-green-300 rounded-lg p-3 flex items-center justify-between">
                      <div className="flex items-center gap-2 flex-1 min-w-0">
                        <svg className="w-5 h-5 text-green-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                        <div className="flex-1 min-w-0">
                          <p className="text-sm font-medium text-gray-900 truncate">
                            {slot.file.name}
                          </p>
                          <p className="text-xs text-gray-500">
                            {(slot.file.size / 1024).toFixed(2)} KB
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => removeFile(index)}
                        className="ml-2 text-red-500 hover:text-red-700 flex-shrink-0"
                        aria-label="Remove file"
                      >
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                      </button>
                    </div>
                  )}
                </div>
              ))}
            </div>

            <button
              type="submit"
              disabled={uploading || getUploadedCount() === 0}
              className="w-full py-3 px-4 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition"
            >
              {uploading ? (
                <span className="flex items-center justify-center">
                  <svg className="animate-spin -ml-1 mr-3 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                  </svg>
                  Validating {getUploadedCount()} file{getUploadedCount() !== 1 ? 's' : ''}...
                </span>
              ) : (
                `Validate ${getUploadedCount()} File${getUploadedCount() !== 1 ? 's' : ''}`
              )}
            </button>
          </form>

          {error && (
            <div className="mt-6 bg-red-50 border-l-4 border-red-500 p-4 rounded">
              <div className="flex">
                <div className="flex-shrink-0">
                  <svg className="h-5 w-5 text-red-400" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                  </svg>
                </div>
                <div className="ml-3">
                  <p className="text-sm text-red-700">{error}</p>
                </div>
              </div>
            </div>
          )}

          {validationResults && validationResults.results && validationResults.results.length > 0 && (
            <div className="mt-8 space-y-6">
              <div className="flex items-center justify-between">
                <h2 className="text-2xl font-semibold text-gray-800">Validation Results</h2>
                <div className="flex items-center gap-4">
                  {validationResults.crossFileErrorCount > 0 && (
                    <div className="bg-orange-100 border border-orange-300 text-orange-800 px-4 py-2 rounded-lg text-sm font-medium">
                      {validationResults.crossFileErrorCount} Cross-File Error(s) Detected
                    </div>
                  )}
                  <button
                    onClick={handleNext}
                    className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition"
                  >
                    Next
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                <div className="lg:col-span-1 space-y-2">
                  <h3 className="text-sm font-semibold text-gray-600 mb-3 uppercase tracking-wide">
                    Files ({validationResults.results.length})
                  </h3>
                  {validationResults.results.map((result: any, index: number) => {
                    const hasErrors = result.summary.totalChecks > 0;
                    const hasCritical = result.summary.criticalIssues > 0;

                    return (
                      <button
                        key={index}
                        onClick={() => setSelectedResult(index)}
                        className={`w-full text-left px-4 py-3 rounded-lg transition ${
                          selectedResult === index
                            ? 'bg-indigo-100 border-2 border-indigo-500'
                            : 'bg-gray-50 hover:bg-gray-100 border-2 border-transparent'
                        }`}
                      >
                        <div className="flex items-start justify-between">
                          <div className="flex-1 min-w-0">
                            <p className="font-medium text-sm truncate">{result.fileName}</p>
                            <p className="text-xs text-gray-500 mt-0.5">{result.fileType}</p>
                            <div className="flex items-center gap-2 mt-1">
                              {hasErrors ? (
                                <>
                                  <span className="text-xs text-red-600 font-medium">
                                    {result.summary.totalChecks} error(s)
                                  </span>
                                  {hasCritical && (
                                    <span className="flex-shrink-0 w-2 h-2 bg-red-500 rounded-full"></span>
                                  )}
                                </>
                              ) : (
                                <span className="text-xs text-green-600 font-medium">✓ Clean</span>
                              )}
                            </div>
                          </div>
                        </div>
                      </button>
                    );
                  })}
                </div>

                <div className="lg:col-span-3">
                  {selectedResult !== null && validationResults.results[selectedResult] ? (
                    <div className="space-y-4">
                      <div className="bg-gradient-to-r from-indigo-50 to-purple-50 rounded-lg p-6 border border-indigo-200">
                        <h3 className="text-xl font-bold text-gray-800 mb-2">
                          {validationResults.results[selectedResult].fileName}
                        </h3>
                        <p className="text-sm text-indigo-600 font-medium mb-4">
                          Type: {validationResults.results[selectedResult].fileType}
                        </p>
                        <div className="grid grid-cols-4 gap-4">
                          <div className="text-center bg-white rounded-lg p-3 shadow-sm">
                            <p className="text-3xl font-bold text-gray-800">
                              {validationResults.results[selectedResult].summary.totalChecks}
                            </p>
                            <p className="text-xs text-gray-500 mt-1">Total Issues</p>
                          </div>
                          <div className="text-center bg-white rounded-lg p-3 shadow-sm">
                            <p className="text-3xl font-bold text-red-600">
                              {validationResults.results[selectedResult].summary.criticalIssues}
                            </p>
                            <p className="text-xs text-gray-500 mt-1">Critical</p>
                          </div>
                          <div className="text-center bg-white rounded-lg p-3 shadow-sm">
                            <p className="text-3xl font-bold text-orange-600">
                              {validationResults.results[selectedResult].summary.highIssues}
                            </p>
                            <p className="text-xs text-gray-500 mt-1">High</p>
                          </div>
                          <div className="text-center bg-white rounded-lg p-3 shadow-sm">
                            <p className="text-3xl font-bold text-yellow-600">
                              {validationResults.results[selectedResult].summary.mediumIssues}
                            </p>
                            <p className="text-xs text-gray-500 mt-1">Medium</p>
                          </div>
                        </div>
                      </div>

                      {validationResults.results[selectedResult].validationErrors && 
                       validationResults.results[selectedResult].validationErrors.length > 0 ? (
                        <div className="space-y-3">
                          <div className="flex items-center justify-between">
                            <h4 className="font-semibold text-gray-800">
                              Error Details ({validationResults.results[selectedResult].validationErrors.length})
                            </h4>
                            <button
                              onClick={() => {
                                const allErrorKeys = validationResults.results[selectedResult].validationErrors.map(
                                  (_: any, idx: number) => `${selectedResult}-${idx}`
                                );
                                if (expandedErrors.size === allErrorKeys.length) {
                                  setExpandedErrors(new Set());
                                } else {
                                  setExpandedErrors(new Set(allErrorKeys));
                                }
                              }}
                              className="text-sm text-indigo-600 hover:text-indigo-800 font-medium"
                            >
                              {expandedErrors.size === validationResults.results[selectedResult].validationErrors.length
                                ? 'Collapse All'
                                : 'Expand All'}
                            </button>
                          </div>

                          <div className="max-h-[600px] overflow-y-auto space-y-3 pr-2">
                            {validationResults.results[selectedResult].validationErrors.map((error: any, errorIndex: number) => {
                              const errorKey = `${selectedResult}-${errorIndex}`;
                              const colors = getSeverityColor(error.severity);
                              const isExpanded = expandedErrors.has(errorKey);

                              return (
                                <div
                                  key={errorIndex}
                                  className={`${colors.bg} rounded-lg border-l-4 ${colors.border} overflow-hidden transition-all`}
                                >
                                  <div className="p-4">
                                    <div className="flex items-start justify-between mb-2">
                                      <div className="flex items-center gap-2 flex-1">
                                        <span className={`px-2 py-0.5 rounded text-xs font-bold ${colors.badge}`}>
                                          #{error.checkNumber}
                                        </span>
                                        <span className={`px-2 py-0.5 rounded text-xs font-medium ${colors.badge} uppercase`}>
                                          {error.severity}
                                        </span>
                                      </div>
                                      <button
                                        onClick={() => toggleErrorExpansion(errorKey)}
                                        className="text-gray-500 hover:text-gray-700 ml-2"
                                        aria-label={isExpanded ? 'Collapse' : 'Expand'}
                                      >
                                        <svg
                                          className={`w-5 h-5 transition-transform ${isExpanded ? 'rotate-180' : ''}`}
                                          fill="none"
                                          stroke="currentColor"
                                          viewBox="0 0 24 24"
                                        >
                                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                                        </svg>
                                      </button>
                                    </div>

                                    <h5 className={`font-semibold ${colors.text} mb-2`}>
                                      {error.message}
                                    </h5>

                                    <div className="flex flex-wrap gap-3 text-xs text-gray-600 mb-2">
                                      {error.sheet && (
                                        <div className="flex items-center gap-1 bg-white px-2 py-1 rounded">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                          </svg>
                                          <span className="font-medium">Sheet:</span>
                                          <span className="font-mono">{error.sheet}</span>
                                        </div>
                                      )}
                                      {error.row && (
                                        <div className="flex items-center gap-1 bg-white px-2 py-1 rounded">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14" />
                                          </svg>
                                          <span className="font-medium">Row:</span>
                                          <span className="font-mono">{error.row}</span>
                                        </div>
                                      )}
                                      {error.column && (
                                        <div className="flex items-center gap-1 bg-white px-2 py-1 rounded">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
                                          </svg>
                                          <span className="font-medium">Column:</span>
                                          <span className="font-mono">{error.column}</span>
                                        </div>
                                      )}
                                    </div>

                                    {isExpanded && (
                                      <div className="mt-3 pt-3 border-t border-gray-300">
                                        <p className="text-sm text-gray-700 bg-white p-3 rounded leading-relaxed">
                                          {error.details}
                                        </p>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      ) : (
                        <div className="text-center py-12 bg-green-50 rounded-lg border-2 border-green-200">
                          <svg className="w-20 h-20 mx-auto mb-4 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                          </svg>
                          <p className="text-lg font-semibold text-green-700">
                            No validation errors found!
                          </p>
                          <p className="text-sm text-green-600 mt-1">
                            This file passed all validation checks
                          </p>
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center h-64 text-gray-400">
                      <svg className="w-20 h-20 mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                      </svg>
                      <p className="text-lg">Select a file to view validation details</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {showPasswordModal && (
            <div className="fixed inset-0 bg-opacity-50 flex items-center justify-center z-50 back">
              <div className="bg-white shadow-2xl rounded-lg p-6 w-full max-w-md">
                <h3 className="text-lg font-semibold text-gray-800 mb-4">
                  Password Required
                </h3>
                <p className="text-sm text-gray-600 mb-4">
                  High or critical errors detected. Please enter the password to proceed to Step 2.
                </p>
                <form onSubmit={handlePasswordSubmit} className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Password
                    </label>
                    <input
                      type="password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                    {passwordError && (
                      <p className="mt-1 text-sm text-red-600">{passwordError}</p>
                    )}
                  </div>
                  <div className="flex justify-end gap-2">
                    <button
                      type="button"
                      onClick={() => setShowPasswordModal(false)}
                      className="px-4 py-2 bg-gray-200 text-gray-800 rounded-lg hover:bg-gray-300"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700"
                    >
                      Submit
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
