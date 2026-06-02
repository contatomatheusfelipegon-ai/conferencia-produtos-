const { list, put } = require('@vercel/blob');

const BLOB_TOKEN = process.env.BLOB_READ_WRITE_TOKEN || 'vercel_blob_rw_XeCDHbJl1hpJh7WK_XUQ0FcVLIMA74HgSpO1FrJH1IfhqnW';

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { pedido, status } = req.body || {};
  if (!pedido || !status) return res.status(400).json({ error: 'Campos obrigatorios' });

  try {
    const { blobs } = await list({ prefix: `meta_${pedido}.json`, token: BLOB_TOKEN });
    if (!blobs.length) return res.status(404).json({ error: 'Pedido nao encontrado' });

    const resp = await fetch(blobs[0].url);
    const meta = await resp.json();
    meta.status = status;

    await put(`meta_${pedido}.json`, JSON.stringify(meta), {
      access: 'public',
      contentType: 'application/json',
      addRandomSuffix: false,
      token: BLOB_TOKEN
    });

    return res.status(200).json({ ok: true });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
