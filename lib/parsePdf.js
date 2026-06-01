const pdfParse = require('pdf-parse');

function extrairCabecalho(txt) {
  const cab = {};
  let m;

  m = txt.match(/Pedido de Venda\s*[-–]\s*(\d+)/);
  cab.pedido = m ? m[1] : '';

  m = txt.match(/Data Emiss[aã]o\s*:\s*(\d{2}\/\d{2}\/\d{4})/);
  cab.data_emissao = m ? m[1] : '';

  m = txt.match(/Cliente\s*:\s*(.+?)\s*\(M\d+/);
  cab.cliente = m ? m[1].trim() : '';

  m = txt.match(/Endere[çc]o\s*:\s*(.+?)(?=\n)/);
  cab.endereco = m ? m[1].trim() : '';

  m = txt.match(/Tipo Frete\s*:\s*(\w+)/);
  cab.tipo_frete = m ? m[1].trim() : '';

  m = txt.match(/Cond\.Pgt\.\s*:\s*\d+\s*-\s*(.+?)(?=\n)/);
  cab.cond_pgt = m ? m[1].trim() : '';

  m = txt.match(/CNPJ\/CPF\s*:\s*([\d.\/\-]+)/);
  cab.cnpj = m ? m[1].trim() : '';

  m = txt.match(/Repres\.\s*:\s*\w+\s*-\s*(.+?)(?=\n)/);
  cab.representante = m ? m[1].trim() : '';

  m = txt.match(/Volume\s*:\s*(\d+\s*VOLUMES?)/i);
  cab.volume = m ? m[1].trim() : '';

  m = txt.match(/Transp\.\s*:\s*(.+?)(?=\n)/);
  cab.transportadora = m ? m[1].trim() : '';

  return cab;
}

function limparDescricao(desc) {
  desc = desc.replace(/\s+/g, ' ').trim();
  desc = desc.replace(/(\d)\s+,(\d)/g, '$1,$2');
  desc = desc.replace(/\b([A-Z])\s+([A-Z])\b/g, '$1$2');
  desc = desc.replace('LEMO NADE', 'LEMONADE');
  return desc;
}

function extrairProdutos(txt) {
  const linhas = txt.split('\n');
  const PAD = /^(\d+)\s+(PN\w+)\s+(.+)\s+([A-Z0-9]{6,8})\s+(\d{2}\/\d{2})\s+[\d,]+\s+[\d*,]+/;
  const PAD_IGNORAR = /^[-_]{5,}|^Qtde\s|^Folha\.|^SIGA\s|^Hora\.|^Data\s|^Cliente\s|^Endere|^Tipo\s|^CNPJ|^Repres\.|^Volume\s|^Nat\.|^Frete|^\d+o\s+Vencto|^Transp\.|^Ender\.|^NUM\.CXA|^\(cid|^\d{3,}\s+[\d.,]|^SANTO\s|^RUA\s/;
  const PAD_CONT = /^[A-Za-z0-9\s\-,./'()]+$/;

  const produtosRaw = [];
  let i = 0;

  while (i < linhas.length) {
    const linha = linhas[i].trim();
    i++;
    if (!linha || PAD_IGNORAR.test(linha)) continue;
    const m = PAD.exec(linha);
    if (!m) continue;

    let qtde = parseInt(m[1]);
    const cod = m[2];
    let desc = m[3].trim();
    const lote = m[4];
    const valid = m[5];

    if (i < linhas.length) {
      const prox = linhas[i].trim();
      if (prox && PAD_CONT.test(prox) && !PAD.test(prox) && !PAD_IGNORAR.test(prox)) {
        desc = desc + ' ' + prox;
        i++;
      }
    }

    produtosRaw.push({
      qtde, cod,
      descricao: limparDescricao(desc),
      lote, validade: valid
    });
  }

  // Agrupa por código
  const grupos = {};
  for (const p of produtosRaw) {
    if (!grupos[p.cod]) {
      grupos[p.cod] = { qtde: 0, lotes: [], validades: [], descricao: '' };
    }
    grupos[p.cod].qtde += p.qtde;
    grupos[p.cod].descricao = p.descricao;
    if (!grupos[p.cod].lotes.includes(p.lote)) grupos[p.cod].lotes.push(p.lote);
    if (!grupos[p.cod].validades.includes(p.validade)) grupos[p.cod].validades.push(p.validade);
  }

  return Object.entries(grupos).map(([cod, d]) => ({
    cod,
    descricao: d.descricao,
    lote: d.lotes.join(' / '),
    validade: d.validades.join(' / '),
    qtde: d.qtde
  }));
}

async function lerPdf(buffer) {
  const data = await pdfParse(buffer);
  const txt = data.text;
  return {
    cab: extrairCabecalho(txt),
    produtos: extrairProdutos(txt)
  };
}

module.exports = { lerPdf };
