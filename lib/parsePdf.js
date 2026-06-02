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

/**
 * Quando o pdf-parse junta dois produtos na mesma linha (colunas lado a lado
 * no PDF do TOTVS), a descrição do primeiro produto acaba contendo o início
 * do segundo: "PRODUTO A ... QTD PNxxxx PRODUTO B LOTE DATA PRECO TOTAL"
 *
 * Esta função quebra a string capturada em múltiplos produtos sempre que
 * detecta um novo código PN embutido precedido de quantidade inteira.
 *
 * Retorna array de objetos { qtde, cod, descricao, lote, validade } ou null
 * se a linha não contiver produtos embutidos extras.
 */
function quebrarLinhaComposta(qtde, cod, descRaw, lote, validade) {
  // Padrão: "... <QTD> PNxxxx <RESTO>"
  // Ex: "PURO WHEY COOKIES'N CREAM - 909 GR 4 PN8093 PURO WHEY MORANGO - 909 GR2628063 03/28 129,97 519,88"
  const PAD_EMBUTIDO = /^(.*?)\s+(\d+)\s+(PN\w+)\s+(.+?)\s+([A-Z0-9]{6,8}[A-Z]?)\s+(\d{2}\/\d{2})\s+[\d,]+\s+[\d,]+\s*$/;

  const m = PAD_EMBUTIDO.exec(descRaw);
  if (!m) return null;

  const desc1    = m[1].trim();
  const qtde2    = parseInt(m[2]);
  const cod2     = m[3];
  const desc2Raw = m[4].trim();
  const lote2    = m[5];
  const valid2   = m[6];

  // Desc2 pode conter o lote colado — limpar sufixo de lote/preço se ainda sobrou
  const desc2 = desc2Raw.replace(/\s*[A-Z0-9]{6,8}[A-Z]?\s+\d{2}\/\d{2}.*$/, '').trim();

  return [
    { qtde, cod, descricao: limparDescricao(desc1), lote, validade },
    { qtde: qtde2, cod: cod2, descricao: limparDescricao(desc2 || desc2Raw), lote: lote2, validade: valid2 },
  ];
}

function extrairProdutos(txt) {
  const linhas = txt.split('\n');

  // Padrão principal de linha de produto
  const PAD = /^(\d+)\s+(PN\w+)\s+(.+)\s+([A-Z0-9]{6,8}[A-Z]?)\s+(\d{2}\/\d{2})\s+[\d,]+\s+[\d*,]+/;

  const PAD_IGNORAR = /^[-_]{5,}|^Qtde\s|^Folha\.|^SIGA\s|^Hora\.|^Data\s|^Cliente\s|^Endere|^Tipo\s|^CNPJ|^Repres\.|^Volume\s|^Nat\.|^Frete|^\d+o\s+Vencto|^Transp\.|^Ender\.|^NUM\.CXA|^\(cid|^\d{3,}\s+[\d.,]|^SANTO\s|^RUA\s/;
  const PAD_CONT   = /^[A-Za-z0-9\s\-,./'()]+$/;

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
    let desc  = m[3].trim();
    const lote    = m[4];
    const validade = m[5];

    // Absorver linha de continuação (descrição quebrada)
    if (i < linhas.length) {
      const prox = linhas[i].trim();
      if (prox && PAD_CONT.test(prox) && !PAD.test(prox) && !PAD_IGNORAR.test(prox)) {
        desc = desc + ' ' + prox;
        i++;
      }
    }

    // ── CORREÇÃO PRINCIPAL ──────────────────────────────────────────────
    // Verifica se a string de descrição contém um segundo produto embutido
    // (ocorre quando o TOTVS coloca dois produtos na mesma linha visual)
    const compostos = quebrarLinhaComposta(qtde, cod, desc, lote, validade);
    if (compostos) {
      produtosRaw.push(...compostos);
    } else {
      produtosRaw.push({
        qtde, cod,
        descricao: limparDescricao(desc),
        lote, validade
      });
    }
  }

  // Agrupa por código (mesmo produto pode aparecer em linhas separadas com lotes distintos)
  const grupos = {};
  for (const p of produtosRaw) {
    if (!grupos[p.cod]) {
      grupos[p.cod] = { qtde: 0, lotes: [], validades: [], descricao: '' };
    }
    grupos[p.cod].qtde += p.qtde;
    grupos[p.cod].descricao = p.descricao;
    if (!grupos[p.cod].lotes.includes(p.lote))
      grupos[p.cod].lotes.push(p.lote);
    if (!grupos[p.cod].validades.includes(p.validade))
      grupos[p.cod].validades.push(p.validade);
  }

  return Object.entries(grupos).map(([cod, d]) => ({
    cod,
    descricao: d.descricao,
    lote:      d.lotes.join(' / '),
    validade:  d.validades.join(' / '),
    qtde:      d.qtde
  }));
}

async function lerPdf(buffer) {
  const data = await pdfParse(buffer);
  const txt  = data.text;
  return {
    cab:     extrairCabecalho(txt),
    produtos: extrairProdutos(txt)
  };
}

module.exports = { lerPdf };
