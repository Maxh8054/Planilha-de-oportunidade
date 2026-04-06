import { sql } from '@vercel/postgres';
import { NextResponse } from 'next/server';

async function ensureTable() {
  try {
    await sql`
      CREATE TABLE IF NOT EXISTS dashboard_sync (
        id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
        data JSONB NOT NULL DEFAULT '[]'::jsonb,
        total_records INTEGER NOT NULL DEFAULT 0,
        updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      )
    `;
  } catch (e) {
    console.error('Table creation error:', e);
  }
}

export async function GET() {
  try {
    await ensureTable();
    const result = await sql`
      SELECT data, total_records, updated_at FROM dashboard_sync WHERE id = 1
    `;

    if (result.rows.length === 0) {
      return NextResponse.json({ data: null, totalRecords: 0, updatedAt: null });
    }

    const row = result.rows[0];
    return NextResponse.json({
      data: row.data,
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
    await ensureTable();
    const body = await request.json();
    const { data, totalRecords } = body;

    if (!data || !Array.isArray(data)) {
      return NextResponse.json({ error: 'Invalid data format' }, { status: 400 });
    }

    const jsonString = JSON.stringify(data);
    const count = totalRecords || data.length;

    // Delete existing then insert (simplest approach)
    await sql`DELETE FROM dashboard_sync WHERE id = 1`;
    await sql`
      INSERT INTO dashboard_sync (id, data, total_records, updated_at)
      VALUES (1, ${jsonString}, ${count}, NOW())
    `;

    return NextResponse.json({ success: true, totalRecords: count });
  } catch (error) {
    console.error('POST error:', error);
    return NextResponse.json({ error: String(error) }, { status: 500 });
  }
}
