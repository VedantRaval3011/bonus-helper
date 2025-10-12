// app/api/audit/messages/route.ts
import { NextRequest, NextResponse } from 'next/server';
import dbConnect  from '@/lib/mongodb';
import crypto from 'crypto';
import { AuditMessage } from '@/lib/models/AuditMessage';

type IncomingItem = {
  step?: number;
  level?: 'error' | 'warning' | 'info';
  tag?: string;
  text: string;
  scope?: 'staff' | 'worker' | 'global';
  source?: string;
  meta?: Record<string, any>;
  createdAt?: string | Date;
};

function coerceStep(s?: any, src?: string): number | null {
  if (typeof s === 'number' && s >= 2 && s <= 9) return s;
  const m = src?.match(/step(\d+)/i);
  if (m) {
    const n = parseInt(m[1], 10);
    if (n >= 2 && n <= 9) return n;
  }
  return null;
}

export async function POST(req: NextRequest) {
  await dbConnect();

  const body = (await req.json()) as { batchId?: string; step?: number; items: IncomingItem[] };
  if (!body || !Array.isArray(body.items) || body.items.length === 0) {
    return NextResponse.json({ error: 'items required' }, { status: 400 });
  }

  const bodyStep = coerceStep(body.step);
  const batchId = body.batchId || crypto.randomUUID();
  const now = new Date();

  const docs = [];
  for (const it of body.items) {
    const step = coerceStep(it.step, it.source) ?? bodyStep;
    if (step == null) {
      return NextResponse.json({ error: 'step (2-9) required on body or item/source' }, { status: 400 });
    }
    docs.push({
      batchId,
      step,
      level: it.level || 'error',
      tag: it.tag || 'mismatch',
      text: it.text,
      scope: it.scope || 'global',
      source: it.source || `step${step}`,
      meta: it.meta || {},
      createdAt: it.createdAt ? new Date(it.createdAt) : now,
    });
  }

  const result = await AuditMessage.insertMany(docs);
  return NextResponse.json({ batchId, step: docs[0].step, inserted: result.length }, { status: 201 });
}

export async function GET(req: NextRequest) {
  await dbConnect();

  const { searchParams } = new URL(req.url);
  const limit = Math.min(parseInt(searchParams.get('limit') || '500', 10), 2000);
  const grouped = (searchParams.get('grouped') || 'false') === 'true';
  const stepParam = searchParams.get('step');
  const step = stepParam ? parseInt(stepParam, 10) : undefined;

  const query: any = {};
  if (step && step >= 2 && step <= 9) query.step = step;

  const messages = await AuditMessage.find(query)
    .sort({ createdAt: -1 })
    .limit(limit)
    .lean()
    .exec();

  if (!grouped) {
    return NextResponse.json({ messages });
  }

  // group by batchId and compute summaries
  const groupMap = new Map<
    string,
    {
      batchId: string;
      step: number;
      startedAt: string;
      endedAt: string;
      count: number;
      levelCounts: Record<'error' | 'warning' | 'info', number>;
      tagCounts: Record<string, number>;
      items: any[];
    }
  >();

  for (const m of messages) {
    const key = m.batchId as string;
    if (!groupMap.has(key)) {
      groupMap.set(key, {
        batchId: key,
        step: m.step as number,
        startedAt: m.createdAt as any,
        endedAt: m.createdAt as any,
        count: 0,
        levelCounts: { error: 0, warning: 0, info: 0 },
        tagCounts: {},
        items: [],
      });
    }
    const g = groupMap.get(key)!;
    g.items.push(m);
    g.count += 1;
    g.levelCounts[m.level as 'error' | 'warning' | 'info'] =
      (g.levelCounts[m.level as 'error' | 'warning' | 'info'] || 0) + 1;
    const tg = (m.tag as string) || 'misc';
    g.tagCounts[tg] = (g.tagCounts[tg] || 0) + 1;
    if (new Date(m.createdAt as any) < new Date(g.startedAt)) g.startedAt = m.createdAt as any;
    if (new Date(m.createdAt as any) > new Date(g.endedAt)) g.endedAt = m.createdAt as any;
  }

  const groups = Array.from(groupMap.values()).sort(
    (a, b) => new Date(b.endedAt).getTime() - new Date(a.endedAt).getTime()
  );

  return NextResponse.json({ groups });
}
