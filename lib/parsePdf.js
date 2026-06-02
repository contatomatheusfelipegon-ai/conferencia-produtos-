/**
 * parsePdf.js — Parser para pedidos TOTVS/SIGA
 * Usa APENAS o pdf-parse (já no package.json). Zero dependências novas.
 */

'use strict';

const pdfParse = require('pdf-parse');

// ── Deduplicação de linha ─────────────────────────────────────────────────────
function dedupLinha(linha) {
  const len = linha.length;
  if (len < 4) return linha;

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

// ── Renderer customizado ──────────────────────────────────────────────────────
function pageRenderComDedup(pageData) {
  return pageData.getTextContent().then(function (textContent) {
    const mapaY = new Map();
    for (const item of textContent.items) {
      if (!item.str) continue;
      const y = Math.round(item.transform[5]);
      mapaY.set(y, (mapaY.get(y) || '') + item.str);
    }
    const ys = [...mapaY.keys()].sort((a, b) => b - a);
    return ys
      .map(y => dedupLinha(mapaY.get(y).trim()))
      .filter(s => s.length > 0)
      .join('\n');
  });
}

// ── Extração de texto ─────────────────────────────────────────────────────────
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
          linhasFinais.push(...textoFolha.split('\n'));
        }
        return textoFolha;
      });
    },
  });

  return linhasFinais.filter(l => l.trim().length > 0);
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

// Lote TOTVS: 6-8 dígitos + letra opcional (ex: 2628033, 262803A, 2607073B)
const RE_LOTE_STR = '[A-Z0-9]{6,8}[A-Z]?';
const RE_VALID_STR = '\\d{2}/\\d{2}';

// Linha COM lote separado por espaço (formato normal):
//   QTD  PNxxxx  DESC  LOTE  MM/AA  VL.UN  V.TOTAL
const PAD_COM_LOTE_ESP = new RegExp(
  `^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s+(${RE_LOTE_STR})\\s+(${RE_VALID_STR})\\s+[\\d,]+\\s+[\\d*,]+`
);

// Linha COM lote GRUDADO na descrição (sem espaço antes do lote):
//   QTD  PNxxxx  DESC_SEM_ESPACO_LOTE  MM/AA  VL.UN  V.TOTAL
// Detecta o lote como sequência de 6-8 dígitos+letra grudada no fim da desc
const PAD_COM_LOTE_GRUDADO = new RegExp(
  `^(\\d+)\\s+(PN\\w+)\\s+(.+?)(${RE_LOTE_STR})\\s+(${RE_VALID_STR})\\s+[\\d,]+\\s+[\\d*,]+`
);

// SEM lote: QTD  PNxxxx  DESC  VL.UN  V.TOTAL
const PAD_SEM_LOTE = /^(\d+)\s+(PN\w+)\s+(.+?)\s+([\d,]+)\s+([\d*,]+)\s*$/;

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

// ── Extração de produtos ──────────────────────────────────────────────────────
function extrairProdutos(linhas) {
  const raw = [];
  let i = 0;

  while (i < linhas.length) {
    const linha = linhas[i].trim();
    i++;
    if (!linha || PAD_IGNORAR.test(linha)) continue;

    // 1) COM lote separado por espaço (formato normal)
    let m = PAD_COM_LOTE_ESP.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw, lote, validade] = m;
      let desc = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
        desc += ' ' + linhas[i].trim(); i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(desc), lote, validade });
      continue;
    }

    // 2) COM lote grudado na descrição (sem espaço)
    m = PAD_COM_LOTE_GRUDADO.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw, lote, validade] = m;
      let desc = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
        desc += ' ' + linhas[i].trim(); i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(desc), lote, validade });
      continue;
    }

    // 3) SEM lote
    m = PAD_SEM_LOTE.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw] = m;
      if (/^[\d.,*]+$/.test(descRaw.trim())) continue;
      let desc = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
        desc += ' ' + linhas[i].trim(); i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(desc), lote: '', validade: '' });
    }
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

