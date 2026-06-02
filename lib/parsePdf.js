/**
 * parsePdf.js — Parser para pedidos TOTVS/SIGA
 *
 * Usa APENAS o pdf-parse que já está no package.json.
 * Zero dependências novas — só copiar e colar.
 *
 * Uso:
 *   const { lerPdf } = require('./parsePdf');
 *   const resultado = await lerPdf(buffer); // { cab, produtos }
 */

'use strict';

const pdfParse = require('pdf-parse');

// ── Remoção de conteúdo duplicado dentro de uma linha ────────────────────────
// O TOTVS renderiza cada linha duas vezes lado a lado no PDF.
// Exemplos: "KGKG"→"KG", "8 PN8092...8 PN8092..."→"8 PN8092..."
function dedupLinha(linha) {
  const len = linha.length;
  if (len < 4) return linha;

  // Strings curtas (≤24): busca prefixo repetido exato
  if (len <= 24) {
    for (let c = 2; c <= Math.floor(len / 2); c++) {
      if (linha.slice(0, c) === linha.slice(c, c * 2)) return linha.slice(0, c);
    }
    return linha;
  }

  // Strings longas: testa cortes ±8 ao redor da metade
  const metade = Math.floor(len / 2);
  for (let corte = metade - 8; corte <= metade + 8; corte++) {
    if (corte < 6) continue;
    const p1 = linha.slice(0, corte);
    const p2 = linha.slice(corte, corte + p1.length);
    if (p1 === p2) return p1.trimEnd();
  }

  // Fallback: localiza repetição pelo início da string
  const inicio = linha.slice(0, Math.min(12, metade));
  const posRep = linha.indexOf(inicio, Math.max(4, metade - 10));
  if (posRep > 0 && posRep <= metade + 10) {
    return linha.slice(0, posRep).trimEnd();
  }

  return linha;
}

// ── Renderer customizado para o pdf-parse ────────────────────────────────────
// Intercepta cada página, agrupa tokens por coordenada Y e deduplica.
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

// ── Extração de texto com controle de folhas duplicadas ──────────────────────
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

// ── Cabeçalho ────────────────────────────────────────────────────────────────
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

// ── Limpeza de descrição ──────────────────────────────────────────────────────
function limparDescricao(desc) {
  return desc.replace(/\s+/g, ' ').replace(/(\d)\s+,(\d)/g, '$1,$2').trim();
}

// ── Padrões ───────────────────────────────────────────────────────────────────
const RE_LOTE = '[A-Z0-9]{6,8}[A-Z]?';

// COM lote: QTD  PNxxxx  DESC  LOTE  MM/AA  VL.UN  V.TOTAL
const PAD_COM_LOTE = new RegExp(
  `^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s+(${RE_LOTE})\\s+(\\d{2}/\\d{2})\\s+[\\d,]+\\s+[\\d*,]+`
);

// SEM lote: QTD  PNxxxx  DESC  VL.UN  V.TOTAL
const PAD_SEM_LOTE = /^(\d+)\s+(PN\w+)\s+(.+?)\s+([\d,]+)\s+([\d*,]+)\s*$/;

// É uma linha de produto (com ou sem lote)?
const PAD_QUALQUER_PROD = /^\d+\s+PN\w+/;

// Linhas que devem ser ignoradas
const PAD_IGNORAR = new RegExp([
  '^[-_]{5,}', '^Qtde\\s', '^Folha\\.', '^SIGA\\s', '^Hora\\.', '^Data\\s',
  '^Cliente\\s', '^Endere', '^Tipo\\s', '^CNPJ', '^Repres\\.', '^Volume\\s',
  '^Nat\\.', '^Frete', '^\\d+o\\s+Vencto', '^Transp\\.', '^Ender\\.',
  '^NUM\\.CXA', '\\(cid', '^\\d{3,}\\s+[\\d.,]',
  '^SANTO\\s', '^RUA\\s', '^GUARULHOS', '^RODOVIA',
  '^SEPARADO', '^EMBALADO', '^CONFERIDO',
].join('|'));

// Linha de continuação de descrição:
// - Só texto/números/pontuação simples
// - NÃO começa com dígito seguido de PN (seria próximo produto)
// - NÃO começa com dígito seguido de lote/validade
const PAD_CONT = /^[A-Za-z0-9\s\-,./'"()]+$/;

function ehContinuacao(linha) {
  if (!linha) return false;
  if (!PAD_CONT.test(linha)) return false;
  if (PAD_QUALQUER_PROD.test(linha)) return false; // é um produto
  if (PAD_IGNORAR.test(linha)) return false;
  if (/^\d{6,}/.test(linha)) return false;          // lote sozinho
  if (/^\d+$/.test(linha)) return false;             // número puro
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

    // COM lote
    let m = PAD_COM_LOTE.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw, lote, validade] = m;
      let desc = descRaw.trim();
      // Absorve UMA linha de continuação (ex: "909 GR", "- 1,8KG")
      if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
        desc += ' ' + linhas[i].trim();
        i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(desc), lote, validade });
      continue;
    }

    // SEM lote
    m = PAD_SEM_LOTE.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw] = m;
      if (/^[\d.,*]+$/.test(descRaw.trim())) continue; // linha de total
      let desc = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].trim())) {
        desc += ' ' + linhas[i].trim();
        i++;
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
