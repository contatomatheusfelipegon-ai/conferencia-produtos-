const ExcelJS = require('exceljs');

// ── Paleta cinza (fiel à imagem) ─────────────────────────────────────────
const COR = {
  TITLE_BG:  'FF404040',  // cinza escuro — barra título
  TITLE_FG:  'FFFFFFFF',  // branco
  LABEL_BG:  'FFD8D8D8',  // cinza médio — células label
  LABEL_FG:  'FF222222',  // quase preto
  VAL_BG:    'FFFFFFFF',  // branco
  VAL_FG:    'FF111111',
  HDR_BG:    'FFB0B0B0',  // cinza — header tabela
  HDR_FG:    'FF111111',
  ROW_EVEN:  'FFFFFFFF',
  ROW_ODD:   'FFF5F5F5',  // cinza muito claro
  TOTAL_BG:  'FFD8D8D8',
  TOTAL_FG:  'FF111111',
  BORDER:    'FFB0B0B0',
};

function bordaFina() {
  const s = { style: 'thin', color: { argb: COR.BORDER } };
  return { top: s, left: s, bottom: s, right: s };
}

function setCell(cell, { value, font, fill, alignment, border }) {
  if (value !== undefined) cell.value = value;
  if (font)      cell.font      = font;
  if (fill)      cell.fill      = fill;
  if (alignment) cell.alignment = alignment;
  if (border)    cell.border    = border;
}

function makeFill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // ── Larguras A4 retrato (14 colunas com OBSERVACOES)
  // A   B     C     D     E     F    G    H    I    J    K    L     M    N
  [4, 8.5, 29, 12, 8.5, 7.5, 6, 7.5, 7.5, 7.5, 7.5, 8, 7.5, 18]
    .forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  const SZ = 7;
  const FONT = (bold = false, color = COR.VAL_FG) =>
    ({ name: 'Arial', bold, size: SZ, color: { argb: color } });
  const ALN = (horizontal = 'center', wrapText = false) =>
    ({ horizontal, vertical: 'middle', wrapText });

  // ── ROW 1: Título ────────────────────────────────────────────────────
  ws.mergeCells('A1:N1');
  setCell(ws.getCell('A1'), {
    value:     `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`,
    font:      { name: 'Arial', bold: true, size: 9, color: { argb: COR.TITLE_FG } },
    fill:      makeFill(COR.TITLE_BG),
    alignment: ALN('center'),
    border:    bordaFina(),
  });
  ws.getRow(1).height = 18;

  // ── ROWS 2–7: Info cabeçalho ─────────────────────────────────────────
  const LABEL_FONT  = FONT(true,  COR.LABEL_FG);
  const LABEL_FILL  = makeFill(COR.LABEL_BG);
  const VAL_FONT    = FONT(false, COR.VAL_FG);
  const WHITE_FILL  = makeFill(COR.VAL_BG);

  const campos = [
    ['Pedido No:',    cab.pedido,       'Cliente:',        cab.cliente],
    ['Data Emissao:', cab.data_emissao, 'Endereco:',       cab.endereco],
    ['Tipo Frete:',   cab.tipo_frete,   'Cond. Pgt:',      cab.cond_pgt],
    ['CNPJ/CPF:',     cab.cnpj,         'Representante:',  cab.representante],
    ['Volume:',       cab.volume,       'Transportadora:', cab.transportadora],
    ['NF Fiscal:',    '',               'Data Envio:',     ''],
  ];

  let row = 2;
  for (const [l1, v1, l2, v2] of campos) {
    ws.mergeCells(`A${row}:B${row}`);
    setCell(ws.getCell(`A${row}`), { value: l1, font: LABEL_FONT, fill: LABEL_FILL, alignment: ALN('left'), border: bordaFina() });

    ws.mergeCells(`C${row}:G${row}`);
    setCell(ws.getCell(`C${row}`), { value: v1, font: VAL_FONT, fill: WHITE_FILL, alignment: ALN('left'), border: bordaFina() });

    ws.mergeCells(`H${row}:I${row}`);
    setCell(ws.getCell(`H${row}`), { value: l2, font: LABEL_FONT, fill: LABEL_FILL, alignment: ALN('left'), border: bordaFina() });

    ws.mergeCells(`J${row}:N${row}`);
    setCell(ws.getCell(`J${row}`), { value: v2, font: VAL_FONT, fill: WHITE_FILL, alignment: ALN('left'), border: bordaFina() });

    ws.getRow(row).height = 13;
    row++;
  }

  row++; // espaço

  // ── Header tabela ────────────────────────────────────────────────────
  const HDR_FILL = makeFill(COR.HDR_BG);
  const hdrs = ['#','CODIGO','DESCRICAO','LOTE','VALID.','PICK.','OK',
                '1a R.','2a R.','3a R.','4a R.','ENVIADO','SALDO','OBSERVACOES'];
  const headerRow = row;

  hdrs.forEach((h, ci) => {
    const c = ws.getCell(row, ci + 1);
    setCell(c, {
      value:     h,
      font:      FONT(true, COR.HDR_FG),
      fill:      HDR_FILL,
      alignment: ALN('center', true),
      border:    bordaFina(),
    });
  });
  ws.getRow(row).height = 16;
  row++;

  // ── Linhas de produto ────────────────────────────────────────────────
  const dataStart = row;

  produtos.forEach((p, idx) => {
    const fundo = makeFill(idx % 2 === 0 ? COR.ROW_EVEN : COR.ROW_ODD);
    const baseFont = FONT(false, COR.VAL_FG);

    const vals = [idx + 1, p.cod, p.descricao, p.lote, p.validade, p.qtde, '', 0, 0, 0, 0, null, null, ''];

    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      if (val !== null) c.value = val;
      c.font      = baseFont;
      c.fill      = fundo;
      c.border    = bordaFina();
      c.alignment = ci === 2  ? { vertical: 'middle', wrapText: true }
                  : ci === 13 ? ALN('left')
                  : ALN('center');
    });

    // Fórmula ENVIADO (col L = 12)
    const ec = ws.getCell(row, 12);
    ec.value     = { formula: `SUM(H${row}:K${row})` };
    ec.font      = baseFont;
    ec.fill      = fundo;
    ec.alignment = ALN('center');
    ec.border    = bordaFina();

    // Fórmula SALDO (col M = 13)
    const sc = ws.getCell(row, 13);
    sc.value     = { formula: `F${row}-L${row}` };
    sc.font      = baseFont;
    sc.fill      = fundo;
    sc.alignment = ALN('center');
    sc.border    = bordaFina();

    ws.getRow(row).height = 13;
    row++;
  });

  const dataEnd = row - 1;

  // ── Linha TOTAL GERAL ────────────────────────────────────────────────
  const TOTAL_FILL = makeFill(COR.TOTAL_BG);
  const TOTAL_FONT = FONT(true, COR.TOTAL_FG);

  ws.mergeCells(`A${row}:E${row}`);
  setCell(ws.getCell(`A${row}`), {
    value: 'TOTAL GERAL',
    font: TOTAL_FONT, fill: TOTAL_FILL,
    alignment: ALN('center'), border: bordaFina(),
  });

  [[6,'F'],[8,'H'],[9,'I'],[10,'J'],[11,'K'],[12,'L'],[13,'M']].forEach(([ci, col]) => {
    const c = ws.getCell(row, ci);
    c.value     = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    c.font      = TOTAL_FONT;
    c.fill      = TOTAL_FILL;
    c.alignment = ALN('center');
    c.border    = bordaFina();
  });

  // células G e N do total (sem fórmula, só estilo)
  [7, 14].forEach(ci => {
    const c = ws.getCell(row, ci);
    c.font   = TOTAL_FONT;
    c.fill   = TOTAL_FILL;
    c.border = bordaFina();
  });

  ws.getRow(row).height = 14;

  // ── Formatação condicional SALDO (M) ────────────────────────────────
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      { type: 'cellIs', operator: 'equal',       formulae: ['0'], priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC8E6C9' }, fgColor: { argb: 'FFC8E6C9' } } } },
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 2,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFF3CC' }, fgColor: { argb: 'FFFFF3CC' } } } },
      { type: 'cellIs', operator: 'lessThan',    formulae: ['0'], priority: 3,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFCCCC' }, fgColor: { argb: 'FFFFCCCC' } } } },
    ],
  });

  // ── Formatação condicional ENVIADO (L) ──────────────────────────────
  ws.addConditionalFormatting({
    ref: `L${dataStart}:L${dataEnd}`,
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFE8F5E9' }, fgColor: { argb: 'FFE8F5E9' } } } },
    ],
  });

  // ── Freeze e impressão retrato A4 ────────────────────────────────────
  ws.views = [{ state: 'frozen', ySplit: dataStart - 1 }];
  ws.pageSetup = {
    orientation:   'portrait',
    paperSize:     9,
    fitToPage:     true,
    fitToWidth:    1,
    fitToHeight:   0,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };

