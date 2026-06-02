/**
 * parsePdf.js — Parser para pedidos TOTVS/SIGA
 * Usa APENAS o pdf-parse (já no package.json). Zero dependências novas.
 */

'use strict';

const pdfParse = require('pdf-parse');

// ── Deduplicação de linha ─────────────────────────────────────────────────────
// O TOTVS renderiza cada linha duas vezes lado a lado no PDF.
function dedupLinha(linha) {
  const len = linha.length;
  if (len < 4) return linha;

  // Caso especial TOTVS 256297: "DESC ... QTD PNxxxx DESC ..."
  // A segunda metade começa com um número seguido de PNxxxx (a quantidade duplicada)
  // Ex: "PURO WHEY BAUN - 909 GR 4 PN8092 PURO WHEY BAUN - 909 GR 2628138..."
  const mDupPN = linha.match(/^(.+?)\s+\d+\s+PN\w+\s+\1/);
  if (mDupPN) {
    return mDupPN[1].trim();
  }

  // Caso especial: "DESC QTD PNxxxx ..." — corta antes do QTD PNxxxx repetido
  const mPN = linha.match(/^(.+?)\s+(\d+\s+PN\w+\s+.+)$/);
  if (mPN) {
    // Verifica se a segunda parte começa a repetir a primeira
    const primeira = mPN[1].trim();
    const segunda  = mPN[2];
    const inicioP  = primeira.slice(0, Math.min(10, primeira.length));
    if (segunda.includes(inicioP)) {
      return primeira;
    }
  }

  if (len <= 24) {
    for (let c = 2; c <= Math.floor(len / 2); c++) {
      if (linha.slice(0, c) === linha.slice(c, c * 2)) return linha.slice(0, c);
    }
    return linha;
  }

  const metade = Math.floor(len / 2);
  for (let corte = metade - 8; corte <= metade + 8; corte++) {
    if (corte < 6) continue;
    const p1 = linha.slice(0, corte);
    const p2 = linha.slice(corte, corte + p1.length);
    if (p1 === p2) return p1.trimEnd();
  }

  const inicio = linha.slice(0, Math.min(12, metade));
  const posRep = linha.indexOf(inicio, Math.max(4, metade - 10));
  if (posRep > 0 && posRep <= metade + 10) {
    return linha.slice(0, posRep).trimEnd();
  }

  return linha;
}

// ── Renderer: agrupa tokens por Y mantendo ordem por X ───────────────────────
function pageRenderComDedup(pageData) {
  return pageData.getTextContent().then(function (textContent) {
    // Agrupa por Y, mantendo tokens ordenados por X
    const mapaY = new Map();
    for (const item of textContent.items) {
      if (!item.str) continue;
      const y = Math.round(item.transform[5]);
      if (!mapaY.has(y)) mapaY.set(y, []);
      mapaY.get(y).push({ x: item.transform[4], str: item.str });
    }

    const ys = [...mapaY.keys()].sort((a, b) => b - a);

    return ys.map(y => {
      const tokens = mapaY.get(y).sort((a, b) => a.x - b.x);
      const linha = tokens.map(t => t.str).join('').trim();
      return dedupLinha(linha);
    }).filter(Boolean).join('\n');
  });
}

// ── Extração de texto com controle de folhas ──────────────────────────────────
async function extrairTexto(buffer) {
  const folhasVistas = new Set();
  const linhasFinais = [];

  await pdfParse(buffer, {
    pagerender: function (pageData) {
      return pageRenderComDedup(pageData).then(function (textoFolha) {
        const match   = textoFolha.match(/Folha\.\.: (\d+)/);
        const idFolha = match ? match[1] : `_${Date.now()}_${Math.random()}`;
        if (!folhasVistas.has(idFolha)) {
          folhasVistas.add(idFolha);
          linhasFinais.push(...textoFolha.split('\n').filter(l => l.trim()));
        }
        return textoFolha;
      });
    },
  });

  return linhasFinais;
}

// ── Cabeçalho ─────────────────────────────────────────────────────────────────
function extrairCabecalho(linhas) {
  const txt = linhas.join('\n');
  const m = (pat, flags) => {
    const r = new RegExp(pat, flags || '').exec(txt);
    return r ? r[1].trim() : '';
  };
  return {
    pedido:         m('Pedido de Venda\\s*[-\u2013]\\s*(\\d+)'),
    data_emissao:   m('Data Emiss[a\u00e3]o\\s*:\\s*(\\d{2}/\\d{2}/\\d{4})'),
    cliente:        m('Cliente\\s*:\\s*(.+?)\\s*\\(M\\d+'),
    endereco:       m('Endere[\u00e7c]o\\s*:\\s*(.+?)(?=\\n)'),
    tipo_frete:     m('Tipo Frete\\s*:\\s*(\\w+)'),
    cond_pgt:       m('Cond\\.Pgt\\.\\s*:\\s*\\d+\\s*-\\s*(.+?)(?=\\n)'),
    cnpj:           m('CNPJ/CPF\\s*:\\s*([\\d./\\-]+)'),
    representante:  m('Repres\\.\\s*:\\s*\\w+\\s*-\\s*(.+?)(?=\\n)'),
    volume:         m('Volume\\s*:\\s*(\\d+\\s*VOLUMES?)', 'i'),
    transportadora: m('Transp\\.\\s*:\\s*(.+?)(?=\\n)'),
  };
}

// ── Limpeza ───────────────────────────────────────────────────────────────────
function limparDescricao(desc) {
  return desc.replace(/\s+/g, ' ').replace(/(\d)\s+,(\d)/g, '$1,$2').trim();
}

// ── Padrões ───────────────────────────────────────────────────────────────────

// Lote TOTVS: começa com dígito, 6-8 chars alfanum, pode terminar com letra
// Ex: 2527038, 262803A, 2607073B
const RE_LOTE_CAP  = /(\d[A-Z0-9]{5,7}[A-Z]?)/;
const RE_VALID_CAP = /(\d{2}\/\d{2})/;

// Linha de produto completa: QTD PNxxxx TUDO-MAIS
// O "TUDO-MAIS" contém descrição + possivelmente lote + validade + valores
const PAD_LINHA_PROD = /^(\d+)\s+(PN\w+)\s+(.+)$/;

// Qualquer linha de produto
const PAD_QUALQUER_PROD = /^\d+\s+PN\w+/;

const PAD_IGNORAR = new RegExp([
  '^[-_]{5,}', '^Qtde\\s', '^Folha\\.', '^SIGA\\s', '^Hora\\.', '^Data\\s',
  '^Cliente\\s', '^Endere', '^Tipo\\s', '^CNPJ', '^Repres\\.', '^Volume\\s',
  '^Nat\\.', '^Frete', '^\\d+o\\s+Vencto', '^Transp\\.', '^Ender\\.',
  '^NUM\\.CXA', '\\(cid', '^\\d{3,}\\s+[\\d.,]',
  '^SANTO\\s', '^RUA\\s', '^GUARULHOS', '^RODOVIA',
  '^SEPARADO', '^EMBALADO', '^CONFERIDO',
].join('|'));

const PAD_CONT = /^[A-Za-z0-9\s\-,./'"()]+$/;

function ehContinuacao(linha) {
  if (!linha) return false;
  if (!PAD_CONT.test(linha)) return false;
  if (PAD_QUALQUER_PROD.test(linha)) return false;
  if (PAD_IGNORAR.test(linha)) return false;
  if (/^\d{6,}/.test(linha)) return false;
  if (/^\d+$/.test(linha)) return false;
  return true;
}

// ── Separa descrição / lote / validade de uma string "resto" ─────────────────
// A string "resto" é tudo depois de "QTD PNxxxx ": pode conter
// descrição, lote, validade, valores monetários misturados.
//
// Estratégia:
//   1. Encontrar o lote (6-8 dígitos+letra) e validade (MM/AA)
//   2. Tudo antes do lote é descrição
//   3. Lote e validade extraídos separadamente
function parseResto(resto) {
  // Remove duplicação do tipo "DESC QTD PNxxxx DESC" dentro do resto
  // Ex: "PURO WHEY BAUN - 909 GR 4 PN8092 PURO WHEY BAUN - 909 GR 2628138 05/28"
  const mDupResto = resto.match(/^(.+?)\s+\d+\s+PN\w+\s+\1/);
  if (mDupResto) {
    // A parte duplicada foi encontrada; o que sobra começa no lote
    resto = resto.slice(mDupResto[0].length).trim();
  } else {
    // Tenta cortar antes de "QTD PNxxxx" se a parte depois tiver lote/validade
    const mPN = resto.match(/^(.*?)\s+\d+\s+PN\w+\s+(.+)$/);
    if (mPN) {
      const parteDepois = mPN[2];
      if (/\d[A-Z0-9]{5,7}/.test(parteDepois) || /\d{2}\/\d{2}/.test(parteDepois)) {
        resto = mPN[1].trim() + ' ' + parteDepois;
      }
    }
  }

  // Busca lote: sequência que começa com dígito, 6-8 chars [A-Z0-9], opcionalmente termina com letra
  // e é seguida de espaços e uma validade MM/AA
  const RE_LOTE_COM_VAL = /(\d[A-Z0-9]{5,7}[A-Z]?)\s+(\d{2}\/\d{2})/;
  const m = RE_LOTE_COM_VAL.exec(resto);

  if (m) {
    const descRaw = resto.slice(0, m.index).trim();
    const lote    = m[1];
    const validade = m[2];
    return { desc: descRaw, lote, validade };
  }

  // Sem lote: tenta extrair só validade
  const mVal = RE_VALID_CAP.exec(resto);
  if (mVal) {
    const descRaw = resto.slice(0, mVal.index).trim();
    return { desc: descRaw, lote: '', validade: mVal[1] };
  }

  // Sem lote nem validade: remove valores monetários do fim
  const descRaw = resto.replace(/[\d,*]+\s*$/, '').replace(/[\d,*]+\s*$/, '').trim();
  return { desc: descRaw, lote: '', validade: '' };
}

// ── Extração de produtos ──────────────────────────────────────────────────────
function extrairProdutos(linhas) {
  const raw = [];
  let i = 0;

  while (i < linhas.length) {
    const linha = linhas[i].trim();
    i++;

    if (!linha || PAD_IGNORAR.test(linha)) continue;

    const mProd = PAD_LINHA_PROD.exec(linha);
    if (!mProd) continue;

    const [, qtdeStr, cod, restoRaw] = mProd;
    let { desc, lote, validade } = parseResto(restoRaw);

    // Absorve linha de continuação (ex: "909 GR", "- 1,8KG", "NADE - 909 GR")
    if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
      desc += ' ' + linhas[i].trim();
      i++;
    }

    // Remove lixo de valores monetários que possa ter sobrado na desc
    desc = desc.replace(/\s+[\d,*]+\s*$/, '').trim();

    raw.push({
      qtde:      parseInt(qtdeStr, 10),
      cod,
      descricao: limparDescricao(desc),
      lote,
      validade,
    });
  }

  // Agrupar por código (múltiplos lotes do mesmo SKU)
  const grupos = new Map();
  for (const p of raw) {
    if (!grupos.has(p.cod)) grupos.set(p.cod, { qtde: 0, lotes: [], validades: [], descricao: '' });
    const g = grupos.get(p.cod);
    g.qtde += p.qtde;
    g.descricao = p.descricao;
    if (p.lote     && !g.lotes.includes(p.lote))        g.lotes.push(p.lote);
    if (p.validade && !g.validades.includes(p.validade)) g.validades.push(p.validade);
  }

  return [...grupos.entries()].map(([cod, d]) => ({
    cod,
    descricao: d.descricao,
    lote:      d.lotes.join(' / '),
    validade:  d.validades.join(' / '),
    qtde:      d.qtde,
  }));
}

// ── Função principal exportada ────────────────────────────────────────────────
async function lerPdf(buffer) {
  const linhas = await extrairTexto(buffer);
  return {
    cab:      extrairCabecalho(linhas),
    produtos: extrairProdutos(linhas),
  };
}

module.exports = { lerPdf };
