const ExcelJS = require('exceljs');

function bordaFina() {
  const s = { style: 'thin', color: { argb: 'FFB0B0B0' } };
  return { top: s, left: s, bottom: s, right: s };
}

function bordaMedia() {
  const s = { style: 'medium', color: { argb: 'FF888888' } };
  return { top: s, left: s, bottom: s, right: s };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // Larguras para retrato A4 — colunas mais compactas
  const cols = [
    { key: 'A', width: 4  },  // #
    { key: 'B', width: 9  },  // Código
    { key: 'C', width: 30 },  // Descrição
    { key: 'D', width: 14 },  // Lote
    { key: 'E', width: 9  },  // Validade
    { key: 'F', width: 8  },  // Picking
    { key: 'G', width: 7  },  // OK
    { key: 'H', width: 8  },  // 1ª Rem
    { key: 'I', width: 8  },  // 2ª Rem
    { key: 'J', width: 8  },  // 3ª Rem
    { key: 'K', width: 8  },  // 4ª Rem
    { key: 'L', width: 9  },  // Enviado
    { key: 'M', width: 8  },  // Saldo
  ];
  cols.forEach(({ key, width }) => { ws.getColumn(key).width = width; });

  // ── Título
  ws.mergeCells('A1:M1');
  const titulo = ws.getCell('A1');
  titulo.value = `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`;
  titulo.font = { name: 'Arial', bold: true, size: 12, color: { argb: 'FF111111' } };
  titulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
  titulo.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo.border = bordaMedia();
  ws.getRow(1).height = 22;

  // ── Cabeçalho info — 2 colunas, fundo branco, texto preto, label cinza
  const campos = [
    ['Pedido No:', cab.pedido,        'Cliente:',        cab.cliente],
    ['Data Emissao:', cab.data_emissao, 'Endereco:',    cab.endereco],
    ['Tipo Frete:', cab.tipo_frete,   'Cond. Pgt:',     cab.cond_pgt],
    ['CNPJ/CPF:', cab.cnpj,           'Representante:', cab.representante],
    ['Volume:', cab.volume,           'Transportadora:', cab.transportadora],
    ['NF Fiscal:', '',                'Data Envio:', ''],
  ];

  const LABEL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
  const WHITE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  const LABEL_FONT = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF444444' } };
  const VAL_FONT   = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };

  let row = 2;
  const cabStart = row;
  for (const [l1, v1, l2, v2] of campos) {
    ws.mergeCells(`A${row}:B${row}`);
    const c1 = ws.getCell(`A${row}`);
    c1.value = l1; c1.font = LABEL_FONT; c1.fill = LABEL_FILL;
    c1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    c1.border = bordaFina();

    ws.mergeCells(`C${row}:G${row}`);
    const c2 = ws.getCell(`C${row}`);
    c2.value = v1; c2.font = VAL_FONT; c2.fill = WHITE_FILL;
    c2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    c2.border = bordaFina();

    ws.mergeCells(`H${row}:I${row}`);
    const c3 = ws.getCell(`H${row}`);
    c3.value = l2; c3.font = LABEL_FONT; c3.fill = LABEL_FILL;
    c3.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    c3.border = bordaFina();

    ws.mergeCells(`J${row}:M${row}`);
    const c4 = ws.getCell(`J${row}`);
    c4.value = v2; c4.font = VAL_FONT; c4.fill = WHITE_FILL;
    c4.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    c4.border = bordaFina();

    ws.getRow(row).height = 14;
    row++;
  }

  row++; // espaço

  // ── Header tabela
  const hdrs = [
    '#','CODIGO','DESCRICAO','LOTE','VALID.','PICK.',
    'OK','1a R.','2a R.','3a R.','4a R.','ENVIADO','SALDO'
  ];
  const headerRow = row;
  hdrs.forEach((h, ci) => {
    const c = ws.getCell(row, ci + 1);
    c.value = h;
    c.font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(row).height = 18;
  row++;

  // ── Linhas de produto
  const dataStart = row;
  produtos.forEach((p, idx) => {
    const corFundo = idx % 2 === 0 ? 'FFFFFFFF' : 'FFF7F7F7';
    const fundo = { type: 'pattern', pattern: 'solid', fgColor: { argb: corFundo } };
    const vals = [idx+1, p.cod, p.descricao, p.lote, p.validade, p.qtde, '', 0, 0, 0, 0, null, null];

    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      c.value = val;
      c.font = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };
      c.fill = fundo;
      c.border = bordaFina();
      if (ci === 2) {
        c.alignment = { vertical: 'middle', wrapText: true };
      } else {
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    // Fórmulas
    const env = ws.getCell(row, 12);
    env.value = { formula: `SUM(H${row}:K${row})` };
    env.font = { name: 'Arial', size: 8 };
    env.fill = fundo;
    env.alignment = { horizontal: 'center', vertical: 'middle' };
    env.border = bordaFina();

    const sal = ws.getCell(row, 13);
    sal.value = { formula: `F${row}-L${row}` };
    sal.font = { name: 'Arial', size: 8 };
    sal.fill = fundo;
    sal.alignment = { horizontal: 'center', vertical: 'middle' };
    sal.border = bordaFina();

    ws.getRow(row).height = 14;
    row++;
  });

  const dataEnd = row - 1;

  // ── Total
  ws.mergeCells(`A${row}:E${row}`);
  const tot = ws.getCell(`A${row}`);
  tot.value = 'TOTAL GERAL';
  tot.font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
  tot.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };
  tot.alignment = { horizontal: 'center', vertical: 'middle' };
  tot.border = bordaFina();

  [['F',6],['H',8],['I',9],['J',10],['K',11],['L',12],['M',13]].forEach(([col, ci]) => {
    const c = ws.getCell(row, ci);
    c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    c.font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = bordaFina();
  });

  for (const ci of [7]) {
    ws.getCell(row, ci).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };
    ws.getCell(row, ci).border = bordaFina();
  }
  ws.getRow(row).height = 16;

  // ── Formatação condicional SALDO (col M) — saldo = 0, > 0, < 0
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['0'], priority: 1,
        style: { font: { bold: true, color: { argb: 'FF1A5C1A' } },
                 fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC8E6C9' } } } },
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 2,
        style: { font: { bold: true, color: { argb: 'FF7B4F00' } },
                 fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFF3CC' } } } },
      { type: 'cellIs', operator: 'lessThan', formulae: ['0'], priority: 3,
        style: { font: { bold: true, color: { argb: 'FF8B0000' } },
                 fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFCCCC' } } } },
    ]
  });

  // ── Formatação condicional ENVIADO (col L) — > 0 destaca, = picking fica verde
  ws.addConditionalFormatting({
    ref: `L${dataStart}:L${dataEnd}`,
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1,
        style: { font: { bold: true, color: { argb: 'FF1A5C1A' } } } },
    ]
  });

  // ── Impressão retrato A4
  ws.views = [{ state: 'frozen', ySplit: dataStart - 1 }];
  ws.pageSetup = {
    orientation: 'portrait',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6 }
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };
