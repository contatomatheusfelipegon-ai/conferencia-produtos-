const ExcelJS = require('exceljs');

function fill(cor) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + cor } };
}

function bordaFina() {
  const s = { style: 'thin', color: { argb: 'FFD0D0D0' } };
  return { top: s, left: s, bottom: s, right: s };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  const larguras = [4,10,38,18,12,10,6,9,9,9,9,10,9,20];
  'ABCDEFGHIJKLMN'.split('').forEach((col, i) => {
    ws.getColumn(col).width = larguras[i];
  });

  // Título
  ws.mergeCells('A1:N1');
  const titulo = ws.getCell('A1');
  titulo.value = 'CONFERENCIA DE PRODUTOS';
  titulo.font = { name: 'Arial', bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  titulo.fill = fill('1A1A2E');
  titulo.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  // Cabeçalho info
  const LABEL = fill('0F3460');
  const VALOR = fill('FFFFFF');
  const campos = [
    ['Pedido No:', cab.pedido,       'Cliente:',         cab.cliente],
    ['Data Emissao:', cab.data_emissao, 'Endereco:',     cab.endereco],
    ['Tipo Frete:', cab.tipo_frete,  'Cond. Pagamento:', cab.cond_pgt],
    ['CNPJ/CPF:', cab.cnpj,          'Representante:',  cab.representante],
    ['Volume:', cab.volume,          'Transportadora:',  cab.transportadora],
    ['NF Fiscal:', '',               'Data Envio:',      ''],
  ];

  let row = 2;
  for (const [l1, v1, l2, v2] of campos) {
    ws.mergeCells(`A${row}:B${row}`);
    const c1 = ws.getCell(`A${row}`);
    c1.value = l1; c1.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    c1.fill = LABEL; c1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.mergeCells(`C${row}:G${row}`);
    const c2 = ws.getCell(`C${row}`);
    c2.value = v1; c2.font = { name: 'Arial', size: 9, color: { argb: 'FF1A1A2E' } };
    c2.fill = VALOR; c2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.mergeCells(`H${row}:I${row}`);
    const c3 = ws.getCell(`H${row}`);
    c3.value = l2; c3.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    c3.fill = LABEL; c3.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.mergeCells(`J${row}:N${row}`);
    const c4 = ws.getCell(`J${row}`);
    c4.value = v2; c4.font = { name: 'Arial', size: 9, color: { argb: 'FF1A1A2E' } };
    c4.fill = VALOR; c4.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

    ws.getRow(row).height = 16;
    row++;
  }

  row++; // espaço

  // Header tabela
  const HTAB = fill('0F3460');
  const hdrs = ['#','CODIGO','DESCRICAO','LOTE','VALIDADE','PICKING','OK',
                '1a REM.','2a REM.','3a REM.','4a REM.','ENVIADO','SALDO','OBSERVACOES'];
  const headerRow = row;
  hdrs.forEach((h, ci) => {
    const c = ws.getCell(row, ci + 1);
    c.value = h;
    c.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    c.fill = HTAB;
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(row).height = 22;
  row++;

  // Produtos
  const dataStart = row;
  produtos.forEach((p, idx) => {
    const fundo = (idx % 2 === 1) ? fill('F5F7FA') : fill('FFFFFF');
    const vals = [idx+1, p.cod, p.descricao, p.lote, p.validade, p.qtde, '', 0, 0, 0, 0, null, null, ''];

    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      c.value = val;
      c.font = { name: 'Arial', size: 9 };
      c.fill = fundo;
      c.border = bordaFina();
      if ([0,5,6,7,8,9,10,11,12].includes(ci)) {
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (ci === 2) {
        c.alignment = { vertical: 'middle', wrapText: true };
      } else {
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    // Fórmulas
    const env = ws.getCell(row, 12);
    env.value = { formula: `SUM(H${row}:K${row})` };
    env.font = { name: 'Arial', size: 9 }; env.fill = fundo;
    env.alignment = { horizontal: 'center', vertical: 'middle' };
    env.border = bordaFina();

    const sal = ws.getCell(row, 13);
    sal.value = { formula: `F${row}-L${row}` };
    sal.font = { name: 'Arial', size: 9 }; sal.fill = fundo;
    sal.alignment = { horizontal: 'center', vertical: 'middle' };
    sal.border = bordaFina();

    ws.getRow(row).height = 15;
    row++;
  });

  const dataEnd = row - 1;

  // Total
  ws.mergeCells(`A${row}:E${row}`);
  const tot = ws.getCell(`A${row}`);
  tot.value = 'TOTAL GERAL';
  tot.font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  tot.fill = HTAB; tot.alignment = { horizontal: 'center', vertical: 'middle' };

  [['F',6],['H',8],['I',9],['J',10],['K',11],['L',12]].forEach(([col, ci]) => {
    const c = ws.getCell(row, ci);
    c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    c.font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    c.fill = HTAB; c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  const ts = ws.getCell(row, 13);
  ts.value = { formula: `SUM(M${dataStart}:M${dataEnd})` };
  ts.font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  ts.fill = HTAB; ts.alignment = { horizontal: 'center', vertical: 'middle' };

  [7, 14].forEach(ci => { ws.getCell(row, ci).fill = HTAB; });
  ws.getRow(row).height = 20;

  // Formatação condicional saldo
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      { type: 'cellIs', operator: 'equal', formulae: ['0'], priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF92D050' } } } },
      { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 2,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFBA00' } } } },
      { type: 'cellIs', operator: 'lessThan', formulae: ['0'], priority: 3,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } },
    ]
  });

  // Impressão
  ws.views = [{ state: 'frozen', ySplit: dataStart - 1 }];
  ws.pageSetup = {
    orientation: 'landscape', paperSize: 9,
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 }
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };
