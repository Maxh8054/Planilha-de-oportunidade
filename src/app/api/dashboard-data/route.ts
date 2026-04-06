import { db } from '@/lib/db';
import { NextResponse } from 'next/server';

async function ensureTable() {
  await db.$executeRaw`
    CREATE TABLE IF NOT EXISTS dashboard_sync (
      id TEXT PRIMARY KEY,
      data TEXT NOT NULL DEFAULT '[]',
      total_records INTEGER NOT NULL DEFAULT 0,
      updated_at TEXT DEFAULT (datetime('now'))
    )
  `;
}

export async function GET() {
  try {
    await ensureTable();
    const rows = await db.$queryRaw<Array<{ id: string; data: string; total_records: number; updated_at: string }>>`
      SELECT id, data, total_records, updated_at FROM dashboard_sync WHERE id = 'singleton'
    `;

    if (rows.length === 0) {
      return NextResponse.json({ data: null, totalRecords: 0, updatedAt: null });
    }

    const row = rows[0];
    return NextResponse.json({
      data: JSON.parse(row.data),
      totalRecords: row.total_records,
      updatedAt: row.updated_at,
    });
  } catch (error) {
    console.error('GET error:', error);
    return NextResponse.json({ data: null, totalRecords: 0, updatedAt: null, error: String(error) });
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { data, totalRecords } = body;

    if (!data || !Array.isArray(data)) {
      return NextResponse.json({ error: 'Invalid data format' }, { status: 400 });
    }

    const jsonString = JSON.stringify(data);
    const count = totalRecords || data.length;

    await ensureTable();

    // Upsert using DELETE + INSERT
    await db.$executeRaw`DELETE FROM dashboard_sync WHERE id = 'singleton'`;
    await db.$executeRaw`
      INSERT INTO dashboard_sync (id, data, total_records, updated_at)
      VALUES ('singleton', ${jsonString}, ${count}, datetime('now'))
    `;

    return NextResponse.json({ success: true, totalRecords: count });
  } catch (error) {
    console.error('POST error:', error);
    return NextResponse.json({ error: String(error) }, { status: 500 });
  }
}
