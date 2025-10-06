'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';

export default function MessagesPage() {
  const [password, setPassword] = useState('');
  const [authenticated, setAuthenticated] = useState(false);
  const [messages, setMessages] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [selectedFile, setSelectedFile] = useState<string | null>(null);
  const [expandedErrors, setExpandedErrors] = useState<Set<number>>(new Set());
  const router = useRouter();

  const handleAuth = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError('');

    try {
      const response = await fetch('/api/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password })
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Authentication failed');
      }

      setAuthenticated(true);
      setMessages(data.messages);
    } catch (err: any) {
      setError(err.message || 'Failed to authenticate');
    } finally {
      setLoading(false);
    }
  };

  const toggleErrorExpansion = (index: number) => {
    const newExpanded = new Set(expandedErrors);
    if (newExpanded.has(index)) {
      newExpanded.delete(index);
    } else {
      newExpanded.add(index);
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

  if (!authenticated) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center px-4">
        <div className="bg-white rounded-2xl shadow-xl p-8 w-full max-w-md">
          <h1 className="text-2xl font-bold text-gray-800 mb-6">Access Messages</h1>
          <form onSubmit={handleAuth} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Password
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                placeholder="Enter password"
              />
            </div>
            {error && (
              <p className="text-sm text-red-600">{error}</p>
            )}
            <button
              type="submit"
              disabled={loading}
              className="w-full py-2 px-4 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 disabled:bg-gray-400 transition"
            >
              {loading ? 'Authenticating...' : 'Access'}
            </button>
            <button
              type="button"
              onClick={() => router.push('/')}
              className="w-full py-2 px-4 bg-gray-200 text-gray-700 font-medium rounded-lg hover:bg-gray-300 transition"
            >
              Back to Upload
            </button>
          </form>
        </div>
      </div>
    );
  }

  const groupedMessages = messages.reduce((acc: any, msg: any) => {
    if (!acc[msg.fileName]) {
      acc[msg.fileName] = [];
    }
    acc[msg.fileName].push(msg);
    return acc;
  }, {});

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-12 px-4">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-8">
            <h1 className="text-3xl font-bold text-gray-800">Validation Messages</h1>
            <button
              onClick={() => router.push('/')}
              className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
            >
              Back to Upload
            </button>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
            {/* File List Sidebar */}
            <div className="lg:col-span-1 space-y-2">
              <h2 className="text-lg font-semibold text-gray-700 mb-4">Files</h2>
              {Object.keys(groupedMessages).map((fileName) => {
                const fileMessages = groupedMessages[fileName];
                const totalErrors = fileMessages.reduce((sum: number, msg: any) => 
                  sum + (msg.validationErrors?.length || 0), 0
                );
                const hasCritical = fileMessages.some((msg: any) => 
                  msg.validationErrors?.some((err: any) => err.severity === 'critical')
                );

                return (
                  <button
                    key={fileName}
                    onClick={() => setSelectedFile(fileName)}
                    className={`w-full text-left px-4 py-3 rounded-lg transition ${
                      selectedFile === fileName
                        ? 'bg-indigo-100 border-2 border-indigo-500'
                        : 'bg-gray-50 hover:bg-gray-100 border-2 border-transparent'
                    }`}
                  >
                    <div className="flex items-start justify-between">
                      <div className="flex-1 min-w-0">
                        <p className="font-medium text-sm truncate">{fileName}</p>
                        <p className="text-xs text-gray-500 mt-1">
                          {totalErrors} error(s)
                        </p>
                      </div>
                      {hasCritical && (
                        <span className="ml-2 flex-shrink-0 w-2 h-2 bg-red-500 rounded-full mt-1"></span>
                      )}
                    </div>
                  </button>
                );
              })}
            </div>

            {/* Messages Display */}
            <div className="lg:col-span-3">
              {selectedFile ? (
                <div className="space-y-4">
                  <h2 className="text-xl font-semibold text-gray-800 mb-4">
                    {selectedFile}
                  </h2>
                  {groupedMessages[selectedFile].map((msg: any) => (
                    <div key={msg._id} className="bg-gray-50 rounded-lg p-6 border border-gray-200">
                      <div className="flex items-center justify-between mb-4">
                        <span className={`px-3 py-1 rounded-full text-sm font-medium ${
                          msg.status === 'success'
                            ? 'bg-green-100 text-green-800'
                            : 'bg-red-100 text-red-800'
                        }`}>
                          {msg.status.toUpperCase()}
                        </span>
                        <span className="text-sm text-gray-500">
                          {new Date(msg.uploadDate).toLocaleString()}
                        </span>
                      </div>

                      <div className="grid grid-cols-4 gap-4 mb-4 p-4 bg-white rounded-lg">
                        <div className="text-center">
                          <p className="text-2xl font-bold text-gray-800">{msg.summary.totalChecks}</p>
                          <p className="text-xs text-gray-500">Total Issues</p>
                        </div>
                        <div className="text-center">
                          <p className="text-2xl font-bold text-red-600">{msg.summary.criticalIssues}</p>
                          <p className="text-xs text-gray-500">Critical</p>
                        </div>
                        <div className="text-center">
                          <p className="text-2xl font-bold text-orange-600">{msg.summary.highIssues}</p>
                          <p className="text-xs text-gray-500">High</p>
                        </div>
                        <div className="text-center">
                          <p className="text-2xl font-bold text-yellow-600">{msg.summary.mediumIssues}</p>
                          <p className="text-xs text-gray-500">Medium</p>
                        </div>
                      </div>

                      {msg.validationErrors && msg.validationErrors.length > 0 && (
                        <div className="space-y-3">
                          <h3 className="font-semibold text-gray-800 mb-3">Error Details</h3>
                          <div className="max-h-[600px] overflow-y-auto space-y-3 pr-2">
                            {msg.validationErrors.map((error: any, index: number) => {
                              const colors = getSeverityColor(error.severity);
                              const isExpanded = expandedErrors.has(index);

                              return (
                                <div
                                  key={index}
                                  className={`${colors.bg} rounded-lg border-l-4 ${colors.border} overflow-hidden transition-all`}
                                >
                                  <div className="p-4">
                                    {/* Error Header */}
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
                                        onClick={() => toggleErrorExpansion(index)}
                                        className="text-gray-500 hover:text-gray-700 ml-2"
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

                                    {/* Error Message */}
                                    <h4 className={`font-semibold ${colors.text} mb-1`}>
                                      {error.message}
                                    </h4>

                                    {/* Location Info */}
                                    <div className="flex flex-wrap gap-3 text-xs text-gray-600 mb-2">
                                      {error.sheet && (
                                        <div className="flex items-center gap-1">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                          </svg>
                                          <span className="font-medium">Sheet:</span>
                                          <span className="bg-white px-2 py-0.5 rounded">{error.sheet}</span>
                                        </div>
                                      )}
                                      {error.row && (
                                        <div className="flex items-center gap-1">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14" />
                                          </svg>
                                          <span className="font-medium">Row:</span>
                                          <span className="bg-white px-2 py-0.5 rounded font-mono">{error.row}</span>
                                        </div>
                                      )}
                                      {error.column && (
                                        <div className="flex items-center gap-1">
                                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
                                          </svg>
                                          <span className="font-medium">Column:</span>
                                          <span className="bg-white px-2 py-0.5 rounded">{error.column}</span>
                                        </div>
                                      )}
                                    </div>

                                    {/* Expandable Details */}
                                    {isExpanded && (
                                      <div className="mt-3 pt-3 border-t border-gray-300">
                                        <p className="text-sm text-gray-700 bg-white p-3 rounded">
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
                      )}

                      {(!msg.validationErrors || msg.validationErrors.length === 0) && (
                        <div className="text-center py-8 text-green-600">
                          <svg className="w-16 h-16 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                          </svg>
                          <p className="font-semibold">No validation errors found!</p>
                        </div>
                      )}
                    </div>
                  ))}
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
      </div>
    </div>
  );
}
