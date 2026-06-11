export default async function handler(req, res) {
  const FILE_ID = "1vNrcI1yQE06_1QV5e6KUiILSGapq7Dzw";

  // ── Password check (server-side) ──────────────────────────────────────────
  // The password is stored in a Vercel Environment Variable named
  // ACCESS_PASSWORD — it is NEVER sent to the browser, so it cannot be found
  // via F12 / DevTools. The client sends the entered password in the
  // "x-access-password" header; the server compares it here.
  const expected = process.env.ACCESS_PASSWORD;
  const provided = req.headers["x-access-password"];

  // If no env var is configured on the server, fail closed (deny access)
  // rather than accidentally serving the file to everyone.
  if (!expected) {
    return res.status(500).json({ error: "Server password not configured." });
  }
  if (provided !== expected) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  // ── Fetch the spreadsheet from Google Drive ────────────────────────────────
  const url = `https://docs.google.com/spreadsheets/d/${FILE_ID}/export?format=xlsx`;

  try {
    const response = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0",
      },
    });

    if (!response.ok) {
      return res.status(response.status).json({ error: `Google Drive returned ${response.status}` });
    }

    const buffer = await response.arrayBuffer();

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    // Do NOT cache an authenticated response on shared CDNs
    res.setHeader("Cache-Control", "private, no-store");
    res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
