'use strict';

/**
 * parsePdf.js  —  Parser de pedidos TOTVS/SIGA
 *
 * Dependência: pdfjs-dist  (npm install pdfjs-dist)
 * O pdf-parse usa o pdfjs-dist internamente; este módulo o chama diretamente
 * para ter acesso à API completa getTextContent() com coordenadas XY,
 * que é a mesma API que o pdf-parse expõe via pagerender.
 *
 * Compatível com a assinatura solicitada:
 *   const { lerPdf } = require('./parsePdf');
 *   const { cab, produtos } = await lerPdf(bufferOuCaminho);
 *
 * Problema resolvido:
 *   O TOTVS renderiza cada token PDF DUAS vezes na mesma posição XY.
 *   Alguns tokens contêm a descrição e o lote grudados (ex: "GR 2628021")
 *   porque a coluna de lote começa exatamente onde a descrição termina.
 *   O parser usa posições X absolutas (colunas) para separar os campos.
 *
 * Retorno:
 *   {
 *     cab: { pedido, data_emissao, cliente, cnpj, representante,
 *             tipo_frete, cond_pgt, volume, transportadora },
 *     produtos: [{ cod, descricao, lote, validade, qtde }]
 *   }
 *
 * Lote agrupado por SKU:  lote = "2628033 / 262803A"
 *                        validade = "02/28"
 */

// ─── Localização do pdfjs-dist ───────────────────────────────────────────────
// Tenta carregar em ordem: instalação local → global → path padrão npm
let PDFJS_PATH;
const CANDIDATES = [
  'pdfjs-dist/legacy/build/pdf.mjs',
  '/home/claude/.npm-global/lib/node_modules/pdfjs-dist/legacy/build/pdf.mjs',
  '/usr/lib/node_modules/pdfjs-dist/legacy/build/pdf.mjs',
];

// ─── Constantes de layout das colunas do PDF TOTVS/SIGA ─────────────────────
//
// Layout observado (coordenadas X em pt):
//   Qtde         x ≈  30–68   (número da quantidade)
//   Código PN    x ≈  57–87   (PNxxxx)
//   Descrição    x ≈ 107–330  (texto livre)
//   Zona perto   x ≈ 300–344  (pode ter final de desc OU lote grudado)
//   Lote         x ≈ 345–420  (código de lote)
//   Validade     x ≈ 414–470  (MM/AA)
//   VL.Un        x ≈ 457–488
//   V.Total      x ≈ 500–540

const COL_NEAR_LOTE = 300;  // tokens com x >= 300 entram na zona de checagem
const COL_LOTE_MIN  = 345;  // tokens com x >= 345 são tratados como lote/validade

// ─── Expressões regulares ────────────────────────────────────────────────────
const RE_PN      = /^PN\d+$/i;
const RE_MM_AA   = /^\d{2}\/\d{2}$/;
// Lote TOTVS: 6-8 chars alfanuméricos, começa com dígito, pode terminar com letra
const RE_LOTE    = /^\d[A-Z0-9]{5,7}[A-Z]?$/i;
const RE_MONEY   = /^\d{1,3}(?:\.\d{3})*,\d{2}$|^\*+,\*+$/;
const RE_FOLHA   = /Folha\.+\s*(\d+)/;

// ─── 1. Extração de tokens por página ────────────────────────────────────────

/**
 * Extrai tokens de uma página com suas coordenadas XY.
 * Remove duplicatas (TOTVS renderiza cada token 2x na mesma posição).
 * Agrupa por linha (Y com tolerância ±3 px) e ordena por X.
 *
 * @param {Object} page  Objeto de página do pdfjs-dist
 * @returns {Array<{y:number, toks:Array<{t:string,x:number,y:number}>}>}
 */
async function extrairPagina(page) {
  const content = await page.getTextContent({ normalizeWhitespace: false });

  // Remove tokens duplicados (mesma posição + mesmo texto)
  const seen = new Set();
  const tokens = [];
  for (const item of content.items) {
    if (!item.str?.trim()) continue;
    const x   = Math.round(item.transform[4]);
    const y   = Math.round(item.transform[5]);
    const key = `${x},${y},${item.str}`;
    if (seen.has(key)) continue;
    seen.add(key);
    tokens.push({ t: item.str, x, y });
  }

  // Agrupa tokens na mesma linha (tolerância ±3 px em Y)
  const map = new Map();
  for (const tok of tokens) {
    const existing = [...map.keys()].find(k => Math.abs(k - tok.y) <= 3);
    const key = existing ?? tok.y;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(tok);
  }

  // Retorna linhas ordenadas de cima para baixo (Y decrescente no PDF)
  return [...map.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([y, toks]) => {
      toks.sort((a, b) => a.x - b.x);
      return { y, toks };
    });
}

// ─── 2. Detecção de número de Folha ──────────────────────────────────────────

function getFolha(toks) {
  const m = toks.map(t => t.t).join(' ').match(RE_FOLHA);
  return m ? parseInt(m[1]) : null;
}

// ─── 3. Separação de "texto lote" grudado ────────────────────────────────────

/**
 * Detecta tokens onde a descrição e o lote foram renderizados juntos.
 * Ex: "GR 2628021"  → { before: "GR",   lote: "2628021" }
 *     "CREA 2426345"→ { before: "CREA",  lote: "2426345" }
 *     "60 2628145"  → { before: "60",    lote: "2628145" }
 */
function splitLoteGrudado(t) {
  const m = t.match(/^(.*?)\s+(\d[A-Z0-9]{5,7}[A-Z]?)$/i);
  if (m && RE_LOTE.test(m[2])) return { before: m[1], lote: m[2] };
  return null;
}

/**
 * Decide se um fragmento de continuação deve ser concatenado sem espaço
 * ao último token da descrição (ex: "G" + "R" → "GR").
 */
function deveConcat(ultimoStr, fragmento) {
  return /^[A-Za-z.,]+$/.test(fragmento) && /[A-Za-z]$/.test(ultimoStr);
}

// ─── 4. Parser de produtos ────────────────────────────────────────────────────

/**
 * Varre todas as linhas de tokens e extrai itens de produto.
 * Lida com:
 *   - Continuações de descrição em linhas seguintes
 *   - Lote grudado no final da descrição
 *   - Produtos sem lote (ex: PN8182)
 *   - Múltiplos lotes do mesmo SKU (agrupados com " / ")
 */
function parseProdutosLinhas(linhas) {
  const raw = [];
  let i = 0;

  while (i < linhas.length) {
    const { toks } = linhas[i];
    i++;
    if (!toks.length) continue;

    const t0 = toks[0];
    const t1 = toks[1];

    // Identifica linha de produto: número inteiro + PNxxxx
    if (!t0 || !/^\d+$/.test(t0.t)) continue;
    if (!t1 || !RE_PN.test(t1.t)) continue;

    const qtde = parseInt(t0.t);
    const cod  = t1.t.toUpperCase();
    let descParts = [], lote = '', validade = '';

    // Processa tokens da linha principal
    for (let j = 2; j < toks.length; j++) {
      const tok = toks[j];
      if (RE_MONEY.test(tok.t)) continue;  // ignora valores monetários

      if (tok.x >= COL_LOTE_MIN) {
        // Zona de lote / validade
        if (!lote && RE_LOTE.test(tok.t.trim())) {
          lote = tok.t.trim();
        } else if (!validade && RE_MM_AA.test(tok.t.trim())) {
          validade = tok.t.trim();
        }
      } else if (tok.x >= COL_NEAR_LOTE) {
        // Zona de transição: pode ser fim da descrição OU lote grudado
        const sg = splitLoteGrudado(tok.t);
        if (sg) {
          // Token contém "descrição lote" junto
          if (sg.before) descParts.push(sg.before);
          if (!lote) lote = sg.lote;
        } else if (!lote && RE_LOTE.test(tok.t.trim())) {
          // Token é apenas o lote
          lote = tok.t.trim();
        } else {
          // Não é lote → é descrição (ex: "1" de "150 TABL.", "G" de "GR")
          descParts.push(tok.t);
        }
      } else {
        // Zona de descrição pura
        descParts.push(tok.t);
      }
    }

    // Absorve linhas de continuação (descrição quebrada em linha seguinte)
    // Ex:  "GLYCOFUEL ... LARANJA  2628064  03/28"
    //      "909 GR"    ← continuação
    while (i < linhas.length) {
      const next = linhas[i];
      if (!next.toks.length) break;
      const ft = next.toks[0];

      // Para se a próxima linha é um novo produto
      if (/^\d+$/.test(ft.t) && next.toks[1] && RE_PN.test(next.toks[1].t)) break;
      if (ft.x < 80 && /^\d/.test(ft.t)) break;

      // Absorve se começa com indentação (x ≥ 100)
      if (ft.x >= 100) {
        for (const ct of next.toks) {
          if (ct.x >= COL_LOTE_MIN || RE_MONEY.test(ct.t)) break;
          // Concatena fragmentos de palavra sem espaço (ex: "G" + "R" → "GR")
          if (descParts.length > 0 && deveConcat(descParts[descParts.length - 1], ct.t)) {
            descParts[descParts.length - 1] += ct.t;
          } else {
            descParts.push(ct.t);
          }
        }
        i++;
        continue;
      }
      break;
    }

    const descricao = descParts.join(' ').replace(/\s+/g, ' ').trim();
    raw.push({ qtde, cod, descricao, lote, validade });
  }

  // Agrupa por código (mesmo SKU com lotes diferentes → soma qtde, concat lotes)
  const grupos = new Map();
  for (const p of raw) {
    if (!grupos.has(p.cod)) {
      grupos.set(p.cod, { qtde: 0, lotes: [], validades: [], descricao: '' });
    }
    const g = grupos.get(p.cod);
    g.qtde += p.qtde;
    // Prefere a descrição mais longa (geralmente a mais completa)
    if (p.descricao.length > g.descricao.length) g.descricao = p.descricao;
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

// ─── 5. Extração do cabeçalho ─────────────────────────────────────────────────

function extrairCabecalho(linhas) {
  const txt = linhas.map(l => l.toks.map(t => t.t).join(' ')).join('\n');
  const m = (pat) => {
    const r = new RegExp(pat).exec(txt);
    return r?.[1]?.trim() ?? '';
  };
  return {
    pedido:         m('Pedido de Venda\\s*-\\s*(\\d+)'),
    data_emissao:   m('Data Emiss[aã]o:\\s*(\\d{2}/\\d{2}/\\d{4})'),
    cliente:        m('Cliente\\s*:\\s*(.+?)\\s*\\(M\\d'),
    cnpj:           m('CNPJ/CPF\\s*:\\s*([\\d./\\-]+)'),
    representante:  m('Repres\\.\\s*:\\s*\\w+\\s*-\\s*(.+?)(?:\\s{2}|$)'),
    tipo_frete:     m('Tipo Frete:\\s*(\\w+)'),
    cond_pgt:       m('Cond\\.Pgt\\.:\\s*\\d+\\s*-\\s*([\\d/]+)'),
    volume:         m('Volume\\s*:\\s*(\\d+\\s*VOLUMES?)'),
    transportadora: m('Transp\\.:\\s*(.+?)(?:\n|$)'),
  };
}

// ─── 6. Função principal exportada ───────────────────────────────────────────

/**
 * Lê um PDF TOTVS/SIGA e retorna cabeçalho + lista de produtos.
 *
 * @param {string|Buffer} fonte  Caminho do arquivo PDF ou Buffer
 * @returns {Promise<{cab: Object, produtos: Array}>}
 */
async function lerPdf(fonte) {
  // Carrega pdfjs-dist dinamicamente (ESM)
  let getDocument;
  for (const candidate of CANDIDATES) {
    try {
      const mod = await import(candidate);
      getDocument = mod.getDocument;
      break;
    } catch {
      // tenta o próximo
    }
  }
  if (!getDocument) {
    throw new Error(
      'pdfjs-dist não encontrado. Instale com: npm install pdfjs-dist\n' +
      'Candidatos tentados:\n' + CANDIDATES.map(c => '  ' + c).join('\n')
    );
  }

  // Carrega o buffer
  let buffer;
  if (Buffer.isBuffer(fonte)) {
    buffer = fonte;
  } else {
    const fs = await import('fs');
    buffer = fs.readFileSync(fonte);
  }

  const doc = await getDocument({ data: new Uint8Array(buffer) }).promise;
  const folhasVistas = new Set();
  const todasLinhas  = [];

  for (let p = 1; p <= doc.numPages; p++) {
    const page   = await doc.getPage(p);
    const linhas = await extrairPagina(page);

    // Ignora páginas repetidas (TOTVS duplica o PDF inteiro em alguns casos)
    const folhaLine = linhas.find(l => getFolha(l.toks) !== null);
    const numFolha  = folhaLine ? getFolha(folhaLine.toks) : `_pag${p}`;
    if (folhasVistas.has(numFolha)) continue;
    folhasVistas.add(numFolha);

    todasLinhas.push(...linhas);
  }

  const cab      = extrairCabecalho(todasLinhas);
  const produtos = parseProdutosLinhas(todasLinhas);

  return { cab, produtos };
}

// ─── 7. Export ────────────────────────────────────────────────────────────────

module.exports = { lerPdf };

// ─── 8. CLI: node parsePdf.js <arquivo.pdf> ──────────────────────────────────

if (require.main === module) {
  const arquivo = process.argv[2];
  if (!arquivo) {
    console.error('Uso: node parsePdf.js <arquivo.pdf>');
    process.exit(1);
  }
  lerPdf(arquivo)
    .then(r => console.log(JSON.stringify(r, null, 2)))
    .catch(err => { console.error(err.message); process.exit(1); });
}

