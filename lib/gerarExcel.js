const ExcelJS = require('exceljs');

function bordaFina() {
  const s = { style: 'thin', color: { argb: 'FFB0B0B0' } };
  return { top: s, left: s, bottom: s, right: s };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // Larguras para retrato A4
  const larguras = [4, 9, 30, 14, 9, 8, 7, 8, 8, 8, 8, 9, 8];
  larguras.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── Título
  ws.mergeCells('A1:M1');
  const titulo = ws.getCell('A1');
  titulo.value = `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`;
  titulo.font = { name: 'Arial', bold: true, size: 12, color: { argb: 'FF111111' } };
  titulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
  titulo.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo.border = bordaFina();
  ws.getRow(1).height = 22;

  // ── Cabeçalho info
  const LABEL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
  const WHITE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  const LABEL_FONT = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF444444' } };
  const VAL_FONT   = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };

  const campos = [
    ['Pedido No:', cab.pedido,         'Cliente:',         cab.cliente],
    ['Data Emissao:', cab.data_emissao,'Endereco:',        cab.endereco],
    ['Tipo Frete:', cab.tipo_frete,    'Cond. Pgt:',       cab.cond_pgt],
    ['CNPJ/CPF:', cab.cnpj,            'Representante:',   cab.representante],
    ['Volume:', cab.volume,            'Transportadora:',  cab.transportadora],
    ['NF Fiscal:', '',                 'Data Envio:',      ''],
  ];

  let row = 2;
  for (const [l1, v1, l2, v2] of campos) {
    ws.mergeCells(`A${row}:B${row}`);
    Object.assign(ws.getCell(`A${row}`), { value: l1, font: LABEL_FONT, fill: LABEL_FILL, border: bordaFina(), alignment: { horizontal: 'left', vertical: 'middle', indent: 1 } });

    ws.mergeCells(`C${row}:G${row}`);
    Object.assign(ws.getCell(`C${row}`), { value: v1, font: VAL_FONT, fill: WHITE_FILL, border: bordaFina(), alignment: { horizontal: 'left', vertical: 'middle', indent: 1 } });

    ws.mergeCells(`H${row}:I${row}`);
    Object.assign(ws.getCell(`H${row}`), { value: l2, font: LABEL_FONT, fill: LABEL_FILL, border: bordaFina(), alignment: { horizontal: 'left', vertical: 'middle', indent: 1 } });

    ws.mergeCells(`J${row}:M${row}`);
    Object.assign(ws.getCell(`J${row}`), { value: v2, font: VAL_FONT, fill: WHITE_FILL, border: bordaFina(), alignment: { horizontal: 'left', vertical: 'middle', indent: 1 } });

    ws.getRow(row).height = 14;
    row++;
  }

  row++; // espaço

  // ── Header tabela
  const CINZA_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD8D8D8' } };
  const hdrs = ['#','CODIGO','DESCRICAO','LOTE','VALID.','PICK.','OK','1a R.','2a R.','3a R.','4a R.','ENVIADO','SALDO'];
  const headerRow = row;
  hdrs.forEach((h, ci) => {
    const c = ws.getCell(row, ci + 1);
    c.value = h;
    c.font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
    c.fill = CINZA_FILL;
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(row).height = 18;
  row++;

  // ── Linhas de produto
  const dataStart = row;
  produtos.forEach((p, idx) => {
    const fgColor = { argb: idx % 2 === 0 ? 'FFFFFFFF' : 'FFF7F7F7' };
    const fundo = { type: 'pattern', pattern: 'solid', fgColor };

    const vals = [idx+1, p.cod, p.descricao, p.lote, p.validade, p.qtde, '', 0, 0, 0, 0, 0, 0];
    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      c.value = val;
      c.font = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };
      c.fill = fundo;
      c.border = bordaFina();
      c.alignment = ci === 2
        ? { vertical: 'middle', wrapText: true }
        : { horizontal: 'center', vertical: 'middle' };
    });

    // Fórmula Enviado
    ws.getCell(row, 12).value = { formula: `SUM(H${row}:K${row})` };
    ws.getCell(row, 12).font = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };
    ws.getCell(row, 12).fill = fundo;
    ws.getCell(row, 12).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 12).border = bordaFina();

    // Fórmula Saldo
    ws.getCell(row, 13).value = { formula: `F${row}-L${row}` };
    ws.getCell(row, 13).font = { name: 'Arial', size: 8, color: { argb: 'FF111111' } };
    ws.getCell(row, 13).fill = fundo;
    ws.getCell(row, 13).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 13).border = bordaFina();

    ws.getRow(row).height = 14;
    row++;
  });

  const dataEnd = row - 1;

  // ── Linha total
  ws.mergeCells(`A${row}:E${row}`);
  ws.getCell(`A${row}`).value = 'TOTAL GERAL';
  ws.getCell(`A${row}`).font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
  ws.getCell(`A${row}`).fill = CINZA_FILL;
  ws.getCell(`A${row}`).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(`A${row}`).border = bordaFina();

  [[6,'F'],[7,'G'],[8,'H'],[9,'I'],[10,'J'],[11,'K'],[12,'L'],[13,'M']].forEach(([ci, col]) => {
    const c = ws.getCell(row, ci);
    if (['F','H','I','J','K','L','M'].includes(col)) {
      c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    }
    c.font = { name: 'Arial', bold: true, size: 8, color: { argb: 'FF111111' } };
    c.fill = CINZA_FILL;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = bordaFina();
  });
  ws.getRow(row).height = 16;

  // ── Formatação condicional SALDO (M) — sem font, só fill
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

  // ── Formatação condicional ENVIADO (L) — destaca quando > 0
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

  // ── Freeze e impressão retrato
  ws.views = [{ state: 'frozen', ySplit: dataStart - 1 }];
  ws.pageSetup = {
    orientation: 'portrait',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };
