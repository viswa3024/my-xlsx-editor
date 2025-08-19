import { NextResponse } from "next/server";

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url);
  const fileUrl = searchParams.get("url");

  if (!fileUrl) {
    return NextResponse.json({ error: "Missing file URL" }, { status: 400 });
  }

  try {
    const res = await fetch(fileUrl);

    if (!res.ok) {
      return NextResponse.json({ error: "Failed to fetch CSV" }, { status: 500 });
    }

    // Pass CSV as plain text back to client
    const text = await res.text();
    return new NextResponse(text, {
      status: 200,
      headers: { "Content-Type": "text/csv" },
    });
  } catch (err) {
    return NextResponse.json({ error: "Fetch error" }, { status: 500 });
  }
}
