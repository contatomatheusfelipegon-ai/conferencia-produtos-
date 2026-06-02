const { lerPdf } = require('../lib/parsePdf');
const { gerarExcel } = require('../lib/gerarExcel');
const { put } = require('@vercel/blob');
const multer = require('multer');

const BLOB_TOKEN = process.env.BLOB_READ_WRITE_TOKEN || 'vercel_blob_rw_XeCDHbJl1hpJh7WK_XUQ0FcVLIMA74HgSpO1FrJH1IfhqnW';

const upload = multer({ storage: multer.memoryStorage() });
module.exports.config = { api: { bodyParser: false } };

function runMiddleware(req, res, fn) {
  return new Promise((resolve, reject) => {
    fn(req, res, (result) => {
      if (result instanceof Error) return reject(result);
      return resolve(result);
    });
  });
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  await runMiddleware(req, res, upload.single('pdf'));
  if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo enviado' });

  try {
    const { cab, produtos } = await lerPdf(req.file.buffer);
    const nomeArquivo = `Conferencia_${cab.pedido}.xlsx`;
    const buffer = await gerarExcel(cab, produtos);

    // Salva o Excel
    const blob = await put(nomeArquivo, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      addRandomSuffix: false,
      token: BLOB_TOKEN
    });

    // Salva metadados como JSON
    const meta = {
      pedido: cab.pedido,
      cliente: cab.cliente,
      nome: nomeArquivo,
      url: blob.url,
      status: 'Aguardando',
      totalProdutos: produtos.length,
      totalUnidades: produtos.reduce((s, p) => s + p.qtde, 0),
      criadoEm: new Date().toISOString()
    };
    await put(`meta_${cab.pedido}.json`, JSON.stringify(meta), {
      access: 'public',
      contentType: 'application/json',
      addRandomSuffix: false,
      token: BLOB_TOKEN
    });

    return res.status(200).json({ ok: true, ...meta });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
