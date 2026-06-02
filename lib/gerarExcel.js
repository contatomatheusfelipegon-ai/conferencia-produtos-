const ExcelJS = require('exceljs');

const FILL_GRAY   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const FILL_WHITE  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
const FILL_STRIPE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F7FA' } };
const FILL_OBS_H  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };

const FONT_BASE  = { name: 'Arial', size: 7, color: { argb: 'FF000000' } };
const FONT_BOLD  = { name: 'Arial', size: 7, bold: true, color: { argb: 'FF000000' } };
const FONT_OBS_H = { name: 'Arial', size: 7, bold: true, color: { argb: 'FF111111' } };

function borda() {
  const s = { style: 'thin', color: { argb: 'FFB0B0B0' } };
  return { top: s, left: s, bottom: s, right: s };
}

function aplicarCelula(cell, { value, font, fill, halign, valign, wrap, border }) {
  if (value !== undefined) cell.value = value;
  if (font)   cell.font      = font;
  if (fill)   cell.fill      = fill;
  if (border) cell.border    = border;
  cell.alignment = {
    horizontal: halign ?? 'center',
    vertical:   valign ?? 'center',
    wrapText:   wrap   ?? true,
  };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // ── Larguras: A,B=8 | C=28 | D,E,F=8 | G,H,I,J,K,L,M=6
  const larguras = [8, 8, 28, 8, 8, 8, 6, 6, 6, 6, 6, 6, 6];
  larguras.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // Altura padrão de todas as linhas = 25
  ws.properties.defaultRowHeight = 25;

  let row = 1;

  // ════════════════════════════════════════════
  // Linha 1 — Título CONFERENCIA DE PRODUTOS
  // ════════════════════════════════════════════
  ws.mergeCells(`A${row}:M${row}`);
  aplicarCelula(ws.getCell(`A${row}`), {
    value:  `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`,
    font:   FONT_BOLD,
    fill:   FILL_GRAY,
    border: borda(),
  });
  ws.getRow(row).height = 25;
  row++;

  // ════════════════════════════════════════════
  // Linhas 2–7 — Info do cabeçalho
  // ════════════════════════════════════════════
  const campos = [
    ['Pedido No:',    cab.pedido,          'Cliente:',         cab.cliente       ],
    ['Data Emissao:', cab.data_emissao,    'Endereco:',        cab.endereco      ],
    ['Tipo Frete:',   cab.tipo_frete,      'Cond. Pagamento:', cab.cond_pgt      ],
    ['CNPJ/CPF:',     cab.cnpj,            'Representante:',   cab.representante ],
    ['Volume:',       cab.volume,          'Transportadora:',  cab.transportadora],
    ['NF Fiscal:',    '',                  'Data Envio:',      ''                ],
  ];

  for (const [l1, v1, l2, v2] of campos) {
    ws.mergeCells(`A${row}:B${row}`);
    aplicarCelula(ws.getCell(`A${row}`), { value: l1, font: FONT_BASE, fill: FILL_GRAY, border: borda() });

    ws.mergeCells(`C${row}:G${row}`);
    aplicarCelula(ws.getCell(`C${row}`), { value: v1, font: FONT_BASE, fill: FILL_WHITE, border: borda() });

    ws.mergeCells(`H${row}:I${row}`);
    aplicarCelula(ws.getCell(`H${row}`), { value: l2, font: FONT_BASE, fill: FILL_GRAY, border: borda() });

    ws.mergeCells(`J${row}:M${row}`);
    aplicarCelula(ws.getCell(`J${row}`), { value: v2, font: FONT_BASE, fill: FILL_WHITE, border: borda() });

    ws.getRow(row).height = 25;
    row++;
  }

  // ════════════════════════════════════════════
  // Linha 8 — Cabeçalho da tabela
  // ════════════════════════════════════════════
  const headerRow = row;
  const hdrs = ['#', 'CODIGO', 'DESCRICAO', 'LOTE', 'VALIDADE', 'PICKING', 'OK',
                '1a REM.', '2a REM.', '3a REM.', '4a REM.', 'ENVIADO', 'SALDO'];
  hdrs.forEach((h, ci) => {
    aplicarCelula(ws.getCell(row, ci + 1), {
      value:  h,
      font:   FONT_BASE,
      fill:   FILL_GRAY,
      border: borda(),
    });
  });
  ws.getRow(row).height = 25;
  row++;

  // ════════════════════════════════════════════
  // Linhas de dados
  // ════════════════════════════════════════════
  const dataStart = row;

  produtos.forEach((p, idx) => {
    const fill = idx % 2 === 0 ? FILL_WHITE : FILL_STRIPE;

    const fixos = [idx + 1, p.cod, p.descricao, p.lote, p.validade, p.qtde, ''];
    fixos.forEach((val, ci) => {
      aplicarCelula(ws.getCell(row, ci + 1), {
        value:  val,
        font:   FONT_BASE,
        fill,
        border: borda(),
      });
    });

    for (let ci = 7; ci <= 10; ci++) {
      aplicarCelula(ws.getCell(row, ci + 1), {
        value:  null,
        font:   FONT_BASE,
        fill,
        border: borda(),
      });
    }

    const envCell = ws.getCell(row, 12);
    envCell.value  = { formula: `SUM(H${row}:K${row})` };
    envCell.font   = FONT_BASE;
    envCell.fill   = fill;
    envCell.border = borda();
    envCell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };

    const salCell = ws.getCell(row, 13);
    salCell.value  = { formula: `F${row}-L${row}` };
    salCell.font   = FONT_BASE;
    salCell.fill   = fill;
    salCell.border = borda();
    salCell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };

    ws.getRow(row).height = 25;
    row++;
  });

  const dataEnd = row - 1;

  // ════════════════════════════════════════════
  // Linha TOTAL GERAL
  // ════════════════════════════════════════════
  ws.mergeCells(`A${row}:E${row}`);
  aplicarCelula(ws.getCell(`A${row}`), {
    value:  'TOTAL GERAL',
    font:   FONT_BOLD,
    fill:   FILL_GRAY,
    border: borda(),
  });

  const totaisCols = [
    [6, 'F'], [7, 'G'], [8, 'H'], [9, 'I'],
    [10, 'J'], [11, 'K'], [12, 'L'], [13, 'M'],
  ];
  for (const [ci, col] of totaisCols) {
    const c = ws.getCell(row, ci);
    if (col !== 'G') {
      c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    }
    c.font   = FONT_BOLD;
    c.fill   = FILL_GRAY;
    c.border = borda();
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  }
  ws.getRow(row).height = 25;
  row++;

  // ════════════════════════════════════════════
  // Linha OBSERVAÇÕES header + 10 linhas em branco
  // ════════════════════════════════════════════
  ws.mergeCells(`A${row}:M${row}`);
  aplicarCelula(ws.getCell(`A${row}`), {
    value:  'OBSERVAÇÕES',
    font:   FONT_OBS_H,
    fill:   FILL_OBS_H,
    border: borda(),
  });
  ws.getRow(row).height = 25;
  row++;

  for (let i = 0; i < 10; i++) {
    ws.mergeCells(`A${row}:M${row}`);
    const c = ws.getCell(`A${row}`);
    c.value     = null;
    c.font      = FONT_BASE;
    c.border    = borda();
    c.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
    ws.getRow(row).height = 25;
    row++;
  }

  // ════════════════════════════════════════════
  // Formatação condicional
  // ════════════════════════════════════════════
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      {
        type: 'cellIs', operator: 'equal',
        formulae: ['0'], priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC8E6C9' }, fgColor: { argb: 'FFC8E6C9' } } }
      },
      {
        type: 'cellIs', operator: 'greaterThan',
        formulae: ['0'], priority: 2,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFF3CC' }, fgColor: { argb: 'FFFFF3CC' } } }
      },
      {
        type: 'cellIs', operator: 'lessThan',
        formulae: ['0'], priority: 3,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFCCCC' }, fgColor: { argb: 'FFFFCCCC' } } }
      }
    ]
  });

  ws.addConditionalFormatting({
    ref: `L${dataStart}:L${dataEnd}`,
    rules: [
      {
        type: 'cellIs', operator: 'greaterThan',
        formulae: ['0'], priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFE8F5E9' }, fgColor: { argb: 'FFE8F5E9' } } }
      }
    ]
  });

  // ════════════════════════════════════════════
  // Freeze e impressão
  // ════════════════════════════════════════════
  ws.views = [{ state: 'frozen', ySplit: headerRow }];

  ws.pageSetup = {
    orientation:    'portrait',
    paperSize:      9,
    fitToPage:      true,
    fitToWidth:     1,
    fitToHeight:    1,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };
