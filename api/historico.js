const { list } = require('@vercel/blob');

const BLOB_TOKEN = process.env.BLOB_READ_WRITE_TOKEN || 'vercel_blob_rw_XeCDHbJl1hpJh7WK_XUQ0FcVLIMA74HgSpO1FrJH1IfhqnW';

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).end();
  try {
    const { blobs } = await list({
      prefix: 'Conferencia_',
      token: BLOB_TOKEN
    });

    const arquivos = blobs.map(b => ({
      nome: b.pathname,
      pedido: b.pathname.replace('Conferencia_', '').replace('.xlsx', ''),
      url: b.url,
      data: new Date(b.uploadedAt).toLocaleDateString('pt-BR', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      })
    }));

    return res.status(200).json(arquivos);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
