import { NextRequest, NextResponse } from "next/server";

export async function POST(req: NextRequest) {
  const { url } = await req.json();
  const res = await fetch(url);
  if (!res.ok) return NextResponse.json({ error: "Failed to fetch" }, { status: 400 });
  const buffer = await res.arrayBuffer();
  return NextResponse.json({ data: Array.from(new Uint8Array(buffer)) });
}
