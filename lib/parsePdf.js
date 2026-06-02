/**
 * lerPdf.js — Parser 100% Node.js para pedidos TOTVS/SIGA
 *
 * Dependência única: pdfjs-dist
 *   npm install pdfjs-dist
 *
 * Uso:
 *   const { lerPdf } = require('./lerPdf');
 *   const resultado = await lerPdf(buffer); // { cab, produtos }
 *
 * Resolve duplicação de conteúdo nos PDFs TOTVS:
 *   O SIGA renderiza cada linha duas vezes (lado a lado) no PDF.
 *   pdfjs-dist agrupa tokens por coordenada Y → detectamos e removemos a metade duplicada.
 */

'use strict';

// ── Importação compatível com CJS e ESM ──────────────────────────────────────
let _pdfjsPromise = null;
function getPdfjs() {
  if (!_pdfjsPromise) {
    _pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs').catch(() =>
      import('pdfjs-dist/build/pdf.mjs')
    );
  }
  return _pdfjsPromise;
}

// ── Remoção de conteúdo duplicado dentro de uma linha ────────────────────────
//
// O TOTVS renderiza cada linha lado a lado no PDF.
// O pdfjs concatena os dois blocos numa string só, resultando em duplicação.
//
// Exemplos reais observados:
//   "8 PN8092 PURO WHEY BAUN - 909 GR...8 PN8092 PURO WHEY BAUN - 909 GR..."
//   "KGKG" → "KG"
//   "TABSTABS" → "TABS"
//   "TABS" → "TABS"  (sem duplicação, não altera)
//
// Estratégia:
//   1. Strings curtas (≤ 24 chars): busca prefixo repetido exato.
//   2. Strings longas: testa cortes ao redor da metade.
//   3. Fallback: localiza o início da string mais adiante no texto.
function dedupLinha(linha) {
  const len = linha.length;
  if (len < 4) return linha;

  // Strings curtas: busca metade exata
  if (len <= 24) {
    for (let c = 2; c <= Math.floor(len / 2); c++) {
      const p1 = linha.slice(0, c);
      const p2 = linha.slice(c, c * 2);
      if (p1 === p2) return p1;
    }
    return linha;
  }

  // Strings longas: testa cortes ±8 ao redor da metade
  const metade = Math.floor(len / 2);
  for (let corte = metade - 8; corte <= metade + 8; corte++) {
    if (corte < 6) continue;
    if (corte * 2 > len + 4) break;
    const p1 = linha.slice(0, corte);
    const p2 = linha.slice(corte, corte + p1.length);
    if (p1 === p2) return p1.trimEnd();
  }

  // Fallback: detecta repetição pelo início da string
  const inicio = linha.slice(0, Math.min(12, metade));
  const posRep = linha.indexOf(inicio, Math.max(4, metade - 10));
  if (posRep > 0 && posRep <= metade + 10) {
    return linha.slice(0, posRep).trimEnd();
  }

  return linha;
}

// ── Extração de texto agrupando tokens por linha (coordenada Y) ──────────────
async function extrairTexto(buffer) {
  const pdfjsLib    = await getPdfjs();
  const getDocument = pdfjsLib.getDocument ?? pdfjsLib.default?.getDocument;

  const doc      = await getDocument({ data: new Uint8Array(buffer) }).promise;
  const numPages = doc.numPages;

  const folhasVistas = new Set();
  const linhasFinais = [];

  for (let p = 1; p <= numPages; p++) {
    const page    = await doc.getPage(p);
    const content = await page.getTextContent();

    // Agrupar tokens pela coordenada Y (arredondada ao inteiro mais próximo)
    // para reconstruir as linhas do documento
    const mapaY = new Map();
    for (const item of content.items) {
      if (!item.str) continue;
      const y = Math.round(item.transform[5]);
      mapaY.set(y, (mapaY.get(y) ?? '') + item.str);
    }

    // Ordenar Y de cima para baixo (Y maior = topo da página em coordenadas PDF)
    const ys = [...mapaY.keys()].sort((a, b) => b - a);

    // Montar linhas removendo duplicações
    const linhasDaPagina = ys
      .map(y => dedupLinha(mapaY.get(y).trim()))
      .filter(s => s.length > 0);

    // Detectar número da folha para ignorar páginas repetidas
    // (o TOTVS às vezes imprime a mesma folha duas vezes)
    const textoFolha = linhasDaPagina.join('\n');
    const matchFolha = textoFolha.match(/Folha\.\.: (\d+)/);
    const idFolha    = matchFolha ? matchFolha[1] : `pag${p}`;

    if (folhasVistas.has(idFolha)) continue;
    folhasVistas.add(idFolha);
    linhasFinais.push(...linhasDaPagina);
  }

  return linhasFinais;
}

// ── Cabeçalho ────────────────────────────────────────────────────────────────
function extrairCabecalho(linhas) {
  const txt = linhas.join('\n');

  /** Atalho: executa regex e retorna grupo 1 ou '' */
  const m = (pat, flags = '') => {
    const r = new RegExp(pat, flags).exec(txt);
    return r ? r[1].trim() : '';
  };

  return {
    pedido:         m('Pedido de Venda\\s*[-\u2013]\\s*(\\d+)'),
    data_emissao:   m('Data Emiss[aã]o\\s*:\\s*(\\d{2}/\\d{2}/\\d{4})'),
    cliente:        m('Cliente\\s*:\\s*(.+?)\\s*\\(M\\d+'),
    endereco:       m('Endere[çc]o\\s*:\\s*(.+?)(?=\\n)'),
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
  return desc
    .replace(/\s+/g, ' ')          // múltiplos espaços → um
    .replace(/(\d)\s+,(\d)/g, '$1,$2') // "1 ,8" → "1,8"
    .trim();
}

// ── Padrões de linha ──────────────────────────────────────────────────────────

// Lote TOTVS: 6-8 dígitos + letra opcional no fim (ex: 2628033, 262803A, 2607073B)
const RE_LOTE = '[A-Z0-9]{6,8}[A-Z]?';

// Linha COM lote e validade:
//   QTD  PNxxxx  DESCRIÇÃO  LOTE  MM/AA  VL.UN  V.TOTAL
const PAD_COM_LOTE = new RegExp(
  `^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s+(${RE_LOTE})\\s+(\\d{2}/\\d{2})\\s+[\\d,]+\\s+[\\d*,]+`
);

// Linha SEM lote (produto sem rastreabilidade):
//   QTD  PNxxxx  DESCRIÇÃO  VL.UN  V.TOTAL
// V.Total pode ser numérico ou ****,**
const PAD_SEM_LOTE = /^(\d+)\s+(PN\w+)\s+(.+?)\s+([\d,]+)\s+([\d*,]+)\s*$/;

// Linhas que devem ser ignoradas completamente
const PAD_IGNORAR = new RegExp([
  '^[-_]{5,}',            // linhas separadoras
  '^Qtde\\s',             // cabeçalho da tabela
  '^Folha\\.',            // número da folha
  '^SIGA\\s',             // cabeçalho SIGA
  '^Hora\\.',             // hora
  '^Data\\s',             // data
  '^Cliente\\s',          // dados do cliente
  '^Endere',              // endereço
  '^Tipo\\s',             // tipo frete
  '^CNPJ',                // CNPJ
  '^Repres\\.',           // representante
  '^Volume\\s',           // volume
  '^Nat\\.',              // natureza da operação
  '^Frete',               // linha de frete/total
  '^\\d+o\\s+Vencto',     // vencimentos
  '^Transp\\.',           // transportadora
  '^Ender\\.',            // endereço transportadora
  '^NUM\\.CXA',           // rodapé
  '\\(cid',               // artefato de encoding
  '^\\d{3,}\\s+[\\d.,]',  // linha de total (ex: "756  111.913,32")
  '^SANTO\\s',            // cidade
  '^RUA\\s',              // rua
  '^GUARULHOS',           // cidade
  '^RODOVIA',             // endereço transportadora
  '^SEPARADO',            // rodapé
  '^EMBALADO',
  '^CONFERIDO',
].join('|'));

// Padrão de linha de continuação de descrição
// (apenas texto/números/pontuação simples, sem início de produto)
const PAD_CONT = /^[A-Za-z0-9\s\-,./'"()]+$/;

// ── Extração de produtos ──────────────────────────────────────────────────────
function extrairProdutos(linhas) {
  const raw = [];
  let i = 0;

  while (i < linhas.length) {
    const linha = linhas[i].trim();
    i++;

    if (!linha || PAD_IGNORAR.test(linha)) continue;

    // ── Padrão COM lote ──────────────────────────────────────────────────────
    let m = PAD_COM_LOTE.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw, lote, validade] = m;
      let desc = descRaw.trim();

      // Absorver linha de continuação da descrição (ex: "909 GR", "- 1,8KG")
      if (i < linhas.length) {
        const prox = linhas[i].trim();
        if (
          prox &&
          PAD_CONT.test(prox) &&
          !PAD_COM_LOTE.test(prox) &&
          !PAD_SEM_LOTE.test(prox) &&
          !PAD_IGNORAR.test(prox) &&
          !/^\d{6,}/.test(prox)   // não é um lote sozinho
        ) {
          desc += ' ' + prox;
          i++;
        }
      }

      raw.push({
        qtde:     parseInt(qtdeStr, 10),
        cod,
        descricao: limparDescricao(desc),
        lote,
        validade,
      });
      continue;
    }

    // ── Padrão SEM lote ──────────────────────────────────────────────────────
    m = PAD_SEM_LOTE.exec(linha);
    if (m) {
      const [, qtdeStr, cod, descRaw] = m;

      // Descarta se a "descrição" for só números (linha de subtotal)
      if (/^[\d.,*]+$/.test(descRaw.trim())) continue;

      let desc = descRaw.trim();

      if (i < linhas.length) {
        const prox = linhas[i].trim();
        if (
          prox &&
          PAD_CONT.test(prox) &&
          !PAD_COM_LOTE.test(prox) &&
          !PAD_SEM_LOTE.test(prox) &&
          !PAD_IGNORAR.test(prox) &&
          !/^\d+$/.test(prox)
        ) {
          desc += ' ' + prox;
          i++;
        }
      }

      raw.push({
        qtde:     parseInt(qtdeStr, 10),
        cod,
        descricao: limparDescricao(desc),
        lote:      '',
        validade:  '',
      });
    }
  }

  // ── Agrupar por código (mesmo SKU pode ter múltiplos lotes no pedido) ──────
  const grupos = new Map();

  for (const p of raw) {
    if (!grupos.has(p.cod)) {
      grupos.set(p.cod, { qtde: 0, lotes: [], validades: [], descricao: '' });
    }
    const g = grupos.get(p.cod);
    g.qtde      += p.qtde;
    g.descricao  = p.descricao; // última descrição encontrada (são idênticas)
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
/**
 * Lê um PDF de pedido TOTVS/SIGA e retorna cabeçalho + lista de produtos.
 *
 * @param {Buffer} buffer  Conteúdo binário do PDF
 * @returns {Promise<{ cab: object, produtos: Array }>}
 */
async function lerPdf(buffer) {
  const linhas = await extrairTexto(buffer);
  return {
    cab:      extrairCabecalho(linhas),
    produtos: extrairProdutos(linhas),
  };
}

module.exports = { lerPdf };

