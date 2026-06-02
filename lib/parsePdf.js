/**
 * parsePdf.js — Parser para pedidos TOTVS/SIGA
 * Usa APENAS o pdf-parse (já no package.json). Zero dependências novas.
 *
 * Estratégia: agrupa tokens por coordenada Y, mas separa por X para
 * distinguir a coluna de descrição da coluna de lote/validade.
 * Isso evita que letras do fim da descrição (ex: "R" de "GR") sejam
 * capturadas como parte do lote.
 */

'use strict';

const pdfParse = require('pdf-parse');

// ── Deduplicação de linha ─────────────────────────────────────────────────────
// O TOTVS renderiza cada linha duas vezes lado a lado no PDF.
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
// Agrupa tokens por Y e retorna objeto com colunas separadas por X.
// Formato de cada linha: { full, desc, resto }
//   full  = linha completa (para cabeçalho)
//   desc  = tokens até X < COLUNA_LOTE (coluna descrição)
//   resto = tokens a partir de X >= COLUNA_LOTE (coluna lote/validade/valores)
//
// A largura da página TOTVS é ~595pt. A coluna de lote começa ~350pt.
// Usamos 340 como limiar conservador.
const COLUNA_LOTE_X = 340;

function pageRenderComDedup(pageData) {
  return pageData.getTextContent().then(function (textContent) {
    // Mapa Y → { descTokens: [], restoTokens: [] }
    const mapaY = new Map();

    for (const item of textContent.items) {
      if (!item.str) continue;
      const y = Math.round(item.transform[5]);
      const x = item.transform[4];

      if (!mapaY.has(y)) mapaY.set(y, { desc: '', resto: '' });
      const linha = mapaY.get(y);

      if (x < COLUNA_LOTE_X) {
        linha.desc += item.str;
      } else {
        linha.resto += item.str;
      }
    }

    const ys = [...mapaY.keys()].sort((a, b) => b - a);

    // Retorna linhas como string especial separada por \x00
    // formato: "DESC\x00RESTO"
    return ys
      .map(y => {
        const { desc, resto } = mapaY.get(y);
        const d = dedupLinha(desc.trim());
        const r = dedupLinha(resto.trim());
        if (!d && !r) return null;
        return d + '\x00' + r;
      })
      .filter(Boolean)
      .join('\n');
  });
}

// ── Extração de texto ─────────────────────────────────────────────────────────
async function extrairTexto(buffer) {
  const folhasVistas = new Set();
  const linhasFinais = []; // cada item: { full, desc, resto }

  await pdfParse(buffer, {
    pagerender: function (pageData) {
      return pageRenderComDedup(pageData).then(function (textoFolha) {
        // Detectar folha para evitar duplicatas de página
        const fullText = textoFolha.replace(/\x00/g, ' ');
        const match    = fullText.match(/Folha\.\.: (\d+)/);
        const idFolha  = match ? match[1] : `_${Date.now()}_${Math.random()}`;

        if (!folhasVistas.has(idFolha)) {
          folhasVistas.add(idFolha);
          for (const rawLinha of textoFolha.split('\n')) {
            const sep  = rawLinha.indexOf('\x00');
            const desc = rawLinha.slice(0, sep).trim();
            const resto = rawLinha.slice(sep + 1).trim();
            const full = (desc + ' ' + resto).trim();
            if (full) linhasFinais.push({ full, desc, resto });
          }
        }
        return textoFolha;
      });
    },
  });

  return linhasFinais;
}

// ── Cabeçalho ─────────────────────────────────────────────────────────────────
function extrairCabecalho(linhas) {
  const txt = linhas.map(l => l.full).join('\n');
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

// Lote: começa com dígito, 6-8 chars alfanum + letra opcional no fim
const RE_LOTE     = /\d[A-Z0-9]{5,7}[A-Z]?/;
const RE_VALIDADE = /\d{2}\/\d{2}/;

// Linha de produto no campo desc: QTD  PNxxxx  DESCRIÇÃO
const PAD_PROD_DESC = /^(\d+)\s+(PN\w+)\s+(.+)$/;

// Linha de produto full (quando tudo vem junto sem separação de colunas)
// QTD  PNxxxx  DESC  LOTE  MM/AA  VL.UN  V.TOTAL
const RE_LOTE_STR = '\\d[A-Z0-9]{5,7}[A-Z]?';
const RE_VALID_STR = '\\d{2}/\\d{2}';
const PAD_FULL_COM_LOTE = new RegExp(
  `^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s*(${RE_LOTE_STR})\\s+(${RE_VALID_STR})(?:\\s+[\\d,*]+){1,2}\\s*$`
);
const PAD_FULL_SEM_LOTE = /^(\d+)\s+(PN\w+)\s+(.+?)\s+([\d,]+)\s+([\d*,]+)\s*$/;

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

// Extrai lote e validade de uma string "resto" (coluna direita)
function extrairLoteValidade(resto) {
  let lote = '', validade = '';
  const mLote = RE_LOTE.exec(resto);
  if (mLote) lote = mLote[0];
  const mVal = RE_VALIDADE.exec(resto);
  if (mVal) validade = mVal[0];
  return { lote, validade };
}

// ── Extração de produtos ──────────────────────────────────────────────────────
function extrairProdutos(linhas) {
  const raw = [];
  let i = 0;

  while (i < linhas.length) {
    const { full, desc, resto } = linhas[i];
    i++;

    if (!full || PAD_IGNORAR.test(full)) continue;

    // ── Caso 1: colunas separadas corretamente pelo X ────────────────────────
    // desc tem "QTD PNxxxx DESCRIÇÃO", resto tem lote e validade
    const mDesc = PAD_PROD_DESC.exec(desc);
    if (mDesc && resto) {
      const [, qtdeStr, cod, descRaw] = mDesc;
      let descFinal = descRaw.trim();

      // Absorve linha de continuação (ex: "909 GR", "- 1,8KG")
      if (i < linhas.length && ehContinuacao(linhas[i].desc.trim()) && !linhas[i].resto.trim()) {
        descFinal += ' ' + linhas[i].desc.trim(); i++;
      }

      const { lote, validade } = extrairLoteValidade(resto);
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(descFinal), lote, validade });
      continue;
    }

    // ── Caso 2: tudo numa linha só COM lote (duplicação ou layout compacto) ──
    let m = PAD_FULL_COM_LOTE.exec(full);
    if (m) {
      const [, qtdeStr, cod, descRaw, lote, validade] = m;
      let descFinal = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].full.trim())) {
        descFinal += ' ' + linhas[i].full.trim(); i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(descFinal), lote, validade });
      continue;
    }

    // ── Caso 3: tudo numa linha só SEM lote ──────────────────────────────────
    m = PAD_FULL_SEM_LOTE.exec(full);
    if (m) {
      const [, qtdeStr, cod, descRaw] = m;
      if (/^[\d.,*]+$/.test(descRaw.trim())) continue;
      let descFinal = descRaw.trim();
      if (i < linhas.length && ehContinuacao(linhas[i].full.trim())) {
        descFinal += ' ' + linhas[i].full.trim(); i++;
      }
      raw.push({ qtde: parseInt(qtdeStr, 10), cod, descricao: limparDescricao(descFinal), lote: '', validade: '' });
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

