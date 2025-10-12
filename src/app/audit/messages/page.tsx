// app/audit/messages/page.tsx
'use client';

import React, { useEffect, useMemo, useState } from 'react';

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

const steps = [2, 3, 4, 5, 6, 7, 8, 9];

export default function AuditMessagesPage() {
  const [activeStep, setActiveStep] = useState<number>(2);
  const [groups, setGroups] = useState<Group[]>([]);
  const [loading, setLoading] = useState(true);
  const [includeSnapshots, setIncludeSnapshots] = useState(false); // off by default

  async function load(step: number) {
    setLoading(true);
    const res = await fetch(`/api/audit/messages?grouped=true&step=${step}`, { cache: 'no-store' });
    const json = await res.json();
    setGroups((json.groups || []) as Group[]);
    setLoading(false);
  }

  useEffect(() => {
    load(activeStep);
    const id = setInterval(() => load(activeStep), 10000);
    return () => clearInterval(id);
  }, [activeStep]);

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="mx-auto max-w-7xl space-y-4">
        <header className="flex items-center justify-between">
          <h1 className="text-2xl font-semibold text-gray-800">Audit Dashboard</h1>
          <div className="flex items-center gap-4">
            <label className="flex items-center gap-2 text-sm text-gray-700">
              <input
                type="checkbox"
                className="h-4 w-4"
                checked={includeSnapshots}
                onChange={(e) => setIncludeSnapshots(e.target.checked)}
              />
              Show metric snapshots
            </label>
            <button
              className="px-3 py-1.5 rounded border text-sm bg-white hover:bg-gray-100"
              onClick={() => load(activeStep)}
            >
              Refresh
            </button>
          </div>
        </header>

        <nav className="flex flex-wrap gap-2">
          {steps.map((s) => (
            <button
              key={s}
              onClick={() => setActiveStep(s)}
              className={
                'px-3 py-1.5 rounded border text-sm ' +
                (activeStep === s
                  ? 'bg-indigo-600 text-white border-indigo-600'
                  : 'bg-white text-gray-700 hover:bg-gray-100')
              }
            >
              Step {s}
            </button>
          ))}
        </nav>

        {loading && (
          <div className="p-6 bg-white rounded-lg shadow text-gray-600">Loading…</div>
        )}

        {!loading && groups.length === 0 && (
          <div className="p-6 bg-white rounded-lg shadow text-gray-600">
            No messages for Step {activeStep}.
          </div>
        )}

        {!loading &&
          groups.map((g) => <BatchSection key={g.batchId} g={g} includeSnapshots={includeSnapshots} />)}
      </div>
    </div>
  );
}

function BatchSection({ g, includeSnapshots }: { g: Group; includeSnapshots: boolean }) {
  const filteredItems = useMemo(() => {
    return g.items
      .slice()
      .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime())
      .filter((m) => includeSnapshots || m.tag !== 'metric-snapshot');
  }, [g.items, includeSnapshots]);

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
          <span className="ml-2 text-xs text-gray-500">{g.count} messages</span>
        </div>
      </div>

      <div className="divide-y">
        {filteredItems.map((m) => (
          <MessageRow key={m._id} m={m} />
        ))}
      </div>
    </section>
  );
}

function MessageRow({ m }: { m: Msg }) {
  const levelClass =
    m.level === 'error'
      ? 'bg-red-100 text-red-700'
      : m.level === 'warning'
      ? 'bg-yellow-100 text-yellow-800'
      : 'bg-blue-100 text-blue-700';

  const isSnapshot = m.tag === 'metric-snapshot';

  return (
    <div className="p-3 flex flex-col gap-2">
      <div className="flex items-center gap-2 text-sm">
        <span className={'px-2 py-0.5 rounded text-xs font-medium ' + levelClass}>{m.level}</span>
        <span className="text-gray-800">{m.text}</span>
        <span className="text-gray-400">·</span>
        <span className="text-gray-500">{m.tag}</span>
        <span className="text-gray-400">·</span>
        <span className="text-gray-500">{m.scope}</span>
        <span className="text-gray-400">·</span>
        <span className="text-gray-400">{new Date(m.createdAt).toLocaleString()}</span>
      </div>

      {!isSnapshot && <MismatchChips meta={m.meta} />}

      {isSnapshot && <SnapshotChips meta={m.meta} />}
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

  const diffIntent =
    typeof diff === 'number' ? (Math.abs(diff) >= 1 ? 'red' : 'green') : 'gray';

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
  // Expect meta: { month, snapshot: { A:{...}, B:{...}, C:{...}, D:{...}, E:{...} } }
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
        // Show just the diff chip, minimal dashboard; hide raw JSON
        return (
          <Pill key={k} intent={intent as any}>
            {k}: {d == null ? '—' : d}
          </Pill>
        );
      })}
    </div>
  );
}
