// app/audit/messages/page.tsx
'use client';

import React, { useEffect, useMemo, useState, useCallback, useRef } from 'react';
import { ChevronDown, ChevronUp, X, Filter } from 'lucide-react';

type Msg = {
  _id: string;
  batchId: string;
  step: number;
  level: 'error' | 'warning' | 'info';
  tag: string;
  text: string;
  scope: 'staff' | 'worker' | 'global';
  source: string;
  meta?: Record<string, any>;
  createdAt: string;
};

type Group = {
  batchId: string;
  step: number;
  startedAt: string;
  endedAt: string;
  count: number;
  levelCounts: { error: number; warning: number; info: number };
  tagCounts: Record<string, number>;
  items: Msg[];
};

type FilterState = {
  levels: Set<'error' | 'warning' | 'info'>;
  dateFrom: string;
  dateTo: string;
};

const steps = [2, 3, 4, 5, 6, 7, 8, 9];

export default function AuditMessagesPage() {
  const [activeStep, setActiveStep] = useState<number>(2);
  const [allData, setAllData] = useState<Record<number, Group[]>>({});
  const [loading, setLoading] = useState<Record<number, boolean>>({});
  const [includeSnapshots, setIncludeSnapshots] = useState(false);
  const [showFilters, setShowFilters] = useState(false);
  const [filters, setFilters] = useState<FilterState>({
    levels: new Set(['error', 'warning', 'info']),
    dateFrom: '',
    dateTo: '',
  });

  // Track which steps have been loaded
  const loadedStepsRef = useRef<Set<number>>(new Set());

  // Load data for a specific step only if not already loaded
  const loadStep = useCallback(async (step: number, forceReload = false) => {
    // Skip if already loaded and not forcing reload
    if (loadedStepsRef.current.has(step) && !forceReload) {
      return;
    }
    
    setLoading(prev => ({ ...prev, [step]: true }));
    try {
      const res = await fetch(`/api/audit/messages?grouped=true&step=${step}`, { cache: 'no-store' });
      const json = await res.json();
      setAllData(prev => ({ ...prev, [step]: (json.groups || []) as Group[] }));
      loadedStepsRef.current.add(step);
    } catch (error) {
      console.error(`Error loading step ${step}:`, error);
    } finally {
      setLoading(prev => ({ ...prev, [step]: false }));
    }
  }, []); // Empty dependency array - function doesn't depend on state

  // Refresh current step data
  const refreshCurrentStep = useCallback(() => {
    loadStep(activeStep, true);
  }, [activeStep, loadStep]);

  // Load active step on mount and step change (only if not already loaded)
  useEffect(() => {
    loadStep(activeStep);
  }, [activeStep, loadStep]);

  const handleStepChange = (step: number) => {
    setActiveStep(step);
  };

  const toggleFilter = (level: 'error' | 'warning' | 'info') => {
    setFilters(prev => {
      const newLevels = new Set(prev.levels);
      if (newLevels.has(level)) {
        newLevels.delete(level);
      } else {
        newLevels.add(level);
      }
      return { ...prev, levels: newLevels };
    });
  };

  const clearFilters = () => {
    setFilters({
      levels: new Set(['error', 'warning', 'info']),
      dateFrom: '',
      dateTo: '',
    });
  };

  const hasActiveFilters = filters.levels.size !== 3 || filters.dateFrom || filters.dateTo;

  const currentGroups = allData[activeStep] || [];
  const isLoading = loading[activeStep];

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="mx-auto max-w-7xl space-y-4">
        <header className="flex items-center justify-between flex-wrap gap-4">
          <h1 className="text-2xl font-semibold text-gray-800">Audit Dashboard</h1>
          <div className="flex items-center gap-3 flex-wrap">
            <label className="flex items-center gap-2 text-sm text-gray-700">
              <input
                type="checkbox"
                className="h-4 w-4 rounded border-gray-300"
                checked={includeSnapshots}
                onChange={(e) => setIncludeSnapshots(e.target.checked)}
              />
              Show metric snapshots
            </label>
            <button
              className="flex items-center gap-2 px-3 py-1.5 rounded border text-sm bg-white hover:bg-gray-100 transition-colors"
              onClick={() => setShowFilters(!showFilters)}
            >
              <Filter className="w-4 h-4" />
              Filters
              {hasActiveFilters && (
                <span className="ml-1 px-1.5 py-0.5 bg-indigo-600 text-white rounded-full text-xs">
                  {(filters.levels.size !== 3 ? 1 : 0) + (filters.dateFrom ? 1 : 0) + (filters.dateTo ? 1 : 0)}
                </span>
              )}
            </button>
            <button
              className="px-3 py-1.5 rounded border text-sm bg-white hover:bg-gray-100 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              onClick={refreshCurrentStep}
              disabled={isLoading}
            >
              {isLoading ? 'Loading...' : 'Refresh'}
            </button>
          </div>
        </header>

        {/* Filter Panel */}
        {showFilters && (
          <div className="bg-white rounded-lg shadow p-4 space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="font-medium text-gray-900">Filter Messages</h3>
              {hasActiveFilters && (
                <button
                  onClick={clearFilters}
                  className="text-sm text-indigo-600 hover:text-indigo-800 flex items-center gap-1"
                >
                  <X className="w-4 h-4" />
                  Clear all
                </button>
              )}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Level</label>
                <div className="flex gap-2 flex-wrap">
                  {(['error', 'warning', 'info'] as const).map(level => (
                    <button
                      key={level}
                      onClick={() => toggleFilter(level)}
                      className={`px-3 py-1.5 rounded text-sm border transition-colors ${
                        filters.levels.has(level)
                          ? level === 'error'
                            ? 'bg-red-100 text-red-700 border-red-300'
                            : level === 'warning'
                            ? 'bg-yellow-100 text-yellow-800 border-yellow-300'
                            : 'bg-blue-100 text-blue-700 border-blue-300'
                          : 'bg-gray-100 text-gray-500 border-gray-300'
                      }`}
                    >
                      {level.charAt(0).toUpperCase() + level.slice(1)}
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-2">
                <label className="block text-sm font-medium text-gray-700">Date Range</label>
                <div className="flex gap-2">
                  <input
                    type="date"
                    className="flex-1 px-3 py-1.5 border rounded text-sm"
                    value={filters.dateFrom}
                    onChange={(e) => setFilters(prev => ({ ...prev, dateFrom: e.target.value }))}
                    placeholder="From"
                  />
                  <input
                    type="date"
                    className="flex-1 px-3 py-1.5 border rounded text-sm"
                    value={filters.dateTo}
                    onChange={(e) => setFilters(prev => ({ ...prev, dateTo: e.target.value }))}
                    placeholder="To"
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Step Navigation */}
        <nav className="flex flex-wrap gap-2">
          {steps.map((s) => (
            <button
              key={s}
              onClick={() => handleStepChange(s)}
              className={
                'px-3 py-1.5 rounded border text-sm transition-colors ' +
                (activeStep === s
                  ? 'bg-indigo-600 text-white border-indigo-600'
                  : 'bg-white text-gray-700 hover:bg-gray-100')
              }
            >
              Step {s}
              {allData[s] && (
                <span className="ml-1 text-xs opacity-75">
                  ({allData[s].length})
                </span>
              )}
            </button>
          ))}
        </nav>

        {isLoading && (
          <div className="p-6 bg-white rounded-lg shadow text-gray-600">Loading…</div>
        )}

        {!isLoading && currentGroups.length === 0 && (
          <div className="p-6 bg-white rounded-lg shadow text-gray-600">
            No messages for Step {activeStep}.
          </div>
        )}

        {!isLoading &&
          currentGroups.map((g) => (
            <BatchSection
              key={g.batchId}
              g={g}
              includeSnapshots={includeSnapshots}
              filters={filters}
            />
          ))}
      </div>
    </div>
  );
}

function BatchSection({
  g,
  includeSnapshots,
  filters,
}: {
  g: Group;
  includeSnapshots: boolean;
  filters: FilterState;
}) {
  const [isExpanded, setIsExpanded] = useState(true);

  const filteredItems = useMemo(() => {
    return g.items
      .slice()
      .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime())
      .filter((m) => {
        if (!includeSnapshots && m.tag === 'metric-snapshot') return false;
        if (!filters.levels.has(m.level)) return false;

        const msgDate = new Date(m.createdAt);
        if (filters.dateFrom && msgDate < new Date(filters.dateFrom)) return false;
        if (filters.dateTo && msgDate > new Date(filters.dateTo + 'T23:59:59')) return false;

        return true;
      });
  }, [g.items, includeSnapshots, filters]);

  const chips = [
    { label: 'Errors', value: g.levelCounts.error, intent: 'red' },
    { label: 'Warnings', value: g.levelCounts.warning, intent: 'yellow' },
    { label: 'Info', value: g.levelCounts.info, intent: 'blue' },
  ];

  return (
    <section className="bg-white rounded-lg shadow">
      <div className="border-b p-4 flex items-center justify-between">
        <div className="flex flex-col">
          <span className="text-sm text-gray-500">Step {g.step} · Run {g.batchId.slice(0, 8)}</span>
          <span className="text-gray-800">
            {new Date(g.startedAt).toLocaleString()} → {new Date(g.endedAt).toLocaleString()}
          </span>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            {chips.map((c) => (
              <span
                key={c.label}
                className={
                  'inline-flex items-center gap-1 px-2 py-1 rounded text-xs ' +
                  (c.intent === 'red'
                    ? 'bg-red-100 text-red-700'
                    : c.intent === 'yellow'
                    ? 'bg-yellow-100 text-yellow-800'
                    : 'bg-blue-100 text-blue-700')
                }
              >
                {c.label}: {c.value}
              </span>
            ))}
            <span className="ml-2 text-xs text-gray-500">
              {filteredItems.length} / {g.count} messages
            </span>
          </div>
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1 hover:bg-gray-100 rounded transition-colors"
            aria-label={isExpanded ? 'Collapse' : 'Expand'}
          >
            {isExpanded ? (
              <ChevronUp className="w-5 h-5 text-gray-600" />
            ) : (
              <ChevronDown className="w-5 h-5 text-gray-600" />
            )}
          </button>
        </div>
      </div>

      {isExpanded && (
        <div className="divide-y max-h-96 overflow-y-auto">
          {filteredItems.length === 0 ? (
            <div className="p-4 text-center text-gray-500 text-sm">
              No messages match the current filters
            </div>
          ) : (
            filteredItems.map((m) => <MessageRow key={m._id} m={m} />)
          )}
        </div>
      )}
    </section>
  );
}

function MessageRow({ m }: { m: Msg }) {
  const [isExpanded, setIsExpanded] = useState(false);

  const levelClass =
    m.level === 'error'
      ? 'bg-red-100 text-red-700'
      : m.level === 'warning'
      ? 'bg-yellow-100 text-yellow-800'
      : 'bg-blue-100 text-blue-700';

  const isSnapshot = m.tag === 'metric-snapshot';
  const hasMetadata = m.meta && Object.keys(m.meta).length > 0;

  return (
    <div className="p-3 hover:bg-gray-50 transition-colors">
      <div className="flex items-start justify-between gap-2">
        <div className="flex-1 flex items-center gap-2 text-sm flex-wrap">
          <span className={'px-2 py-0.5 rounded text-xs font-medium ' + levelClass}>{m.level}</span>
          <span className="text-gray-800">{m.text}</span>
          <span className="text-gray-400">·</span>
          <span className="text-gray-500">{m.tag}</span>
          <span className="text-gray-400">·</span>
          <span className="text-gray-500">{m.scope}</span>
          <span className="text-gray-400">·</span>
          <span className="text-gray-400">{new Date(m.createdAt).toLocaleString()}</span>
        </div>
        {hasMetadata && (
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1 hover:bg-gray-200 rounded transition-colors flex-shrink-0"
            aria-label={isExpanded ? 'Hide details' : 'Show details'}
          >
            {isExpanded ? (
              <ChevronUp className="w-4 h-4 text-gray-600" />
            ) : (
              <ChevronDown className="w-4 h-4 text-gray-600" />
            )}
          </button>
        )}
      </div>

      {isExpanded && hasMetadata && (
        <div className="mt-2 pl-4">
          {!isSnapshot && <MismatchChips meta={m.meta} />}
          {isSnapshot && <SnapshotChips meta={m.meta} />}
        </div>
      )}
    </div>
  );
}

function Pill({
  children,
  intent = 'gray',
}: {
  children: React.ReactNode;
  intent?: 'gray' | 'green' | 'red' | 'yellow' | 'blue';
}) {
  const cls =
    intent === 'red'
      ? 'bg-red-50 text-red-700 border-red-200'
      : intent === 'green'
      ? 'bg-green-50 text-green-700 border-green-200'
      : intent === 'yellow'
      ? 'bg-yellow-50 text-yellow-800 border-yellow-200'
      : intent === 'blue'
      ? 'bg-blue-50 text-blue-700 border-blue-200'
      : 'bg-gray-50 text-gray-700 border-gray-200';
  return <span className={`inline-flex items-center px-2 py-0.5 rounded border text-xs ${cls}`}>{children}</span>;
}

function MismatchChips({ meta }: { meta?: Record<string, any> }) {
  if (!meta) return null;
  const emp = meta.employeeCode || meta.empCode;
  const name = meta.name;
  const dept = meta.department;
  const month = meta.month;
  const diff = typeof meta.diff === 'number' ? meta.diff : undefined;
  const actual = meta.actual;
  const hr = meta.hr;

  const diffIntent = typeof diff === 'number' ? (Math.abs(diff) >= 1 ? 'red' : 'green') : 'gray';

  return (
    <div className="flex flex-wrap gap-2">
      {emp && <Pill>{emp}</Pill>}
      {name && <Pill>{name}</Pill>}
      {dept && <Pill intent="blue">{dept}</Pill>}
      {month && <Pill intent="yellow">{month}</Pill>}
      {typeof actual === 'number' && <Pill>actual: {actual}</Pill>}
      {typeof hr === 'number' && <Pill>hr: {hr}</Pill>}
      {typeof diff === 'number' && <Pill intent={diffIntent as any}>diff: {diff}</Pill>}
    </div>
  );
}

function SnapshotChips({ meta }: { meta?: Record<string, any> }) {
  if (!meta || !meta.snapshot) return null;
  const month = meta.month;
  const keys: Array<'A' | 'B' | 'C' | 'D' | 'E'> = ['A', 'B', 'C', 'D', 'E'];

  return (
    <div className="flex flex-wrap items-center gap-2">
      {month && <Pill intent="yellow">{month}</Pill>}
      {keys.map((k) => {
        const s = meta.snapshot[k] || {};
        const d = typeof s.diff === 'number' ? s.diff : null;
        const intent = d == null ? 'gray' : Math.abs(d) >= 1 ? 'red' : 'green';
        return (
          <Pill key={k} intent={intent as any}>
            {k}: {d == null ? '—' : d}
          </Pill>
        );
      })}
    </div>
  );
}
