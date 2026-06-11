export default async function handler(req, res) {
  const FILE_ID = "1ynIpgfPGAr5F6uQ-t0HkXZNShrzPngzd";
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

    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Cache-Control", "s-maxage=300"); // cache 5 phút
    res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
