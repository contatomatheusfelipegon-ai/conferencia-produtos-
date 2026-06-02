const ExcelJS = require('exceljs');

// ── Cores exatas extraídas do modelo (tema Office com tints calculados)
const COR = {
  TITULO_FILL:  'FF737373', // theme2 tint -0.50 → cinza escuro (título e totais)
  HEADER_FILL:  'FFADACAC', // theme2 tint -0.25 → cinza médio (header tabela e labels info)
  BRANCO:       'FFFFFFFF', // linhas pares
  ZEBRA:        'FFF5F7FA', // linhas ímpares (azul muito claro)
  BORDA:        'FFE7E6E6', // theme2 sem tint → cinza claro
  FONTE_CLARA:  'FFFFFFFF', // texto em fundo escuro (título)
  FONTE_ESCURA: 'FF000000', // texto normal

  // Formatação condicional
  CF_SALDO_ZERO: { fg: 'FFC8E6C9', bg: 'FFC8E6C9' }, // verde claro
  CF_SALDO_POS:  { fg: 'FFFFF3CC', bg: 'FFFFF3CC' }, // amarelo claro
  CF_SALDO_NEG:  { fg: 'FFFFCCCC', bg: 'FFFFCCCC' }, // vermelho claro
  CF_ENVIADO:    { fg: 'FFE8F5E9', bg: 'FFE8F5E9' }, // verde muito claro
};

function bordaFina() {
  const s = { style: 'thin', color: { argb: COR.BORDA } };
  return { top: s, left: s, bottom: s, right: s };
}

function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function fonte(opts = {}) {
  return {
    name: 'Arial',
    size: 7,
    color: { argb: opts.color ?? COR.FONTE_ESCURA },
    bold: opts.bold ?? false,
  };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // ── Larguras exatas do modelo (colunas A-N = 14 colunas)
  const larguras = [5.6, 6.3, 19.4, 6.7, 8.3, 6.3, 2.7, 5.1, 5.6, 4.4, 4.4, 7.0, 5.6, 11.1];
  larguras.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── ROW 1: Título — merge A1:N1
  ws.mergeCells('A1:N1');
  const titulo = ws.getCell('A1');
  titulo.value = `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`;
  titulo.font = fonte({ bold: true, color: COR.FONTE_CLARA });
  titulo.fill = fill(COR.TITULO_FILL);
  titulo.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  titulo.border = bordaFina();
  ws.getRow(1).height = 18;

  // ── ROWS 2-7: Campos de cabeçalho
  // Estrutura do modelo: A:B=label, C:G=valor, H:I=label, J:N=valor
  const campos = [
    ['Pedido No:',    cab.pedido,        'Cliente:',        cab.cliente],
    ['Data Emissao:', cab.data_emissao,  'Endereco:',       cab.endereco],
    ['Tipo Frete:',   cab.tipo_frete,    'Cond. Pagamento:', cab.cond_pgt],
    ['CNPJ/CPF:',     cab.cnpj,          'Representante:',  cab.representante],
    ['Volume:',       cab.volume,        'Transportadora:', cab.transportadora],
    ['NF Fiscal:',    '',                'Data Envio:',     ''],
  ];

  for (let i = 0; i < campos.length; i++) {
    const row = i + 2;
    const [l1, v1, l2, v2] = campos[i];

    ws.mergeCells(`A${row}:B${row}`);
    Object.assign(ws.getCell(`A${row}`), {
      value: l1,
      font: fonte({ color: COR.FONTE_ESCURA }),
      fill: fill(COR.HEADER_FILL),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    ws.mergeCells(`C${row}:G${row}`);
    Object.assign(ws.getCell(`C${row}`), {
      value: v1,
      font: fonte(),
      fill: fill(COR.BRANCO),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    ws.mergeCells(`H${row}:I${row}`);
    Object.assign(ws.getCell(`H${row}`), {
      value: l2,
      font: fonte({ color: COR.FONTE_ESCURA }),
      fill: fill(COR.HEADER_FILL),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    ws.mergeCells(`J${row}:N${row}`);
    Object.assign(ws.getCell(`J${row}`), {
      value: v2,
      font: fonte(),
      fill: fill(COR.BRANCO),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    ws.getRow(row).height = 18;
  }

  // ── ROW 8: Header da tabela (14 colunas, igual ao modelo)
  const hdrs = ['#', 'CODIGO', 'DESCRICAO', 'LOTE', 'VALIDADE', 'PICKING', 'OK',
                '1a REM.', '2a REM.', '3a REM.', '4a REM.', 'ENVIADO', 'SALDO', 'OBSERVACOES'];
  const headerRow = 8;
  hdrs.forEach((h, ci) => {
    const c = ws.getCell(headerRow, ci + 1);
    c.value = h;
    c.font = fonte({ bold: false });
    c.fill = fill(COR.HEADER_FILL);
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(headerRow).height = 18;

  // ── ROWS 9+: Linhas de produto
  const dataStart = 9;
  produtos.forEach((p, idx) => {
    const row = dataStart + idx;
    const fgColor = idx % 2 === 0 ? COR.BRANCO : COR.ZEBRA;

    const vals = [idx + 1, p.cod, p.descricao, p.lote, p.validade, p.qtde, '', null, null, null, null, null, null, p.obs ?? ''];
    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      c.value = val;
      c.font = fonte();
      c.fill = fill(fgColor);
      c.border = bordaFina();
      c.alignment = ci === 2
        ? { vertical: 'center', wrapText: true }
        : { horizontal: 'center', vertical: 'center', wrapText: true };
    });

    // Fórmula ENVIADO (coluna L = 12)
    const cEnv = ws.getCell(row, 12);
    cEnv.value = { formula: `SUM(H${row}:K${row})` };
    cEnv.font = fonte();
    cEnv.fill = fill(fgColor);
    cEnv.alignment = { horizontal: 'center', vertical: 'center' };
    cEnv.border = bordaFina();

    // Fórmula SALDO (coluna M = 13)
    const cSal = ws.getCell(row, 13);
    cSal.value = { formula: `F${row}-L${row}` };
    cSal.font = fonte();
    cSal.fill = fill(fgColor);
    cSal.alignment = { horizontal: 'center', vertical: 'center' };
    cSal.border = bordaFina();

    ws.getRow(row).height = 18;
  });

  const dataEnd = dataStart + produtos.length - 1;
  const totalRow = dataEnd + 1;

  // ── Linha TOTAL — merge B:E igual ao modelo (B38:E38)
  ws.mergeCells(`A${totalRow}:E${totalRow}`);
  const cTotalLabel = ws.getCell(`A${totalRow}`);
  cTotalLabel.value = 'TOTAL GERAL';
  cTotalLabel.font = fonte({ bold: true });
  cTotalLabel.fill = fill(COR.HEADER_FILL);
  cTotalLabel.alignment = { horizontal: 'center', vertical: 'center' };
  cTotalLabel.border = bordaFina();

  ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach((col, i) => {
    const ci = 6 + i;
    const c = ws.getCell(totalRow, ci);
    if (['F', 'H', 'I', 'J', 'K', 'L', 'M'].includes(col)) {
      c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    }
    c.font = fonte({ bold: true });
    c.fill = fill(COR.HEADER_FILL);
    c.alignment = { horizontal: 'center', vertical: 'center' };
    c.border = bordaFina();
  });
  ws.getRow(totalRow).height = 18;

  // ── Formatação condicional SALDO (coluna M)
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      {
        type: 'cellIs', operator: 'equal',
        formulae: ['0'], priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_ZERO.bg },
                  fgColor: { argb: COR.CF_SALDO_ZERO.fg } }
        }
      },
      {
        type: 'cellIs', operator: 'greaterThan',
        formulae: ['0'], priority: 2,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_POS.bg },
                  fgColor: { argb: COR.CF_SALDO_POS.fg } }
        }
      },
      {
        type: 'cellIs', operator: 'lessThan',
        formulae: ['0'], priority: 3,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_NEG.bg },
                  fgColor: { argb: COR.CF_SALDO_NEG.fg } }
        }
      },
    ],
  });

  // ── Formatação condicional ENVIADO (coluna L)
  ws.addConditionalFormatting({
    ref: `L${dataStart}:L${dataEnd}`,
    rules: [
      {
        type: 'cellIs', operator: 'greaterThan',
        formulae: ['0'], priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_ENVIADO.bg },
                  fgColor: { argb: COR.CF_ENVIADO.fg } }
        }
      },
    ],
  });

  // ── Freeze nas linhas de cabeçalho (row 8)
  ws.views = [{ state: 'frozen', ySplit: headerRow }];

  // ── Configuração de impressão retrato A4
  ws.pageSetup = {
    orientation: 'portrait',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: `1:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
  };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };

