const ExcelJS = require('exceljs');

// ── Cores extraídas diretamente do MODELO_EXCEL_2_0.xlsx
const COR = {
  TITULO_FILL:  'FFE8E8E8', // cinza claro (título)
  HEADER_FILL:  'FFF0F0F0', // cinza muito claro (labels cabeçalho e header tabela)
  TOTAL_FILL:   'FFD8D8D8', // cinza médio (linha total geral)
  BRANCO:       'FFFFFFFF', // linhas pares
  ZEBRA:        'FFF7F7F7', // linhas ímpares
  BORDA:        'FFB0B0B0', // cinza médio (bordas)
  FONTE_TITULO: 'FF111111', // título (bold 12pt)
  FONTE_HEADER: 'FF444444', // labels/header tabela (bold 8pt)
  FONTE_NORMAL: 'FF111111', // dados normais (8pt)

  // Formatação condicional
  CF_SALDO_ZERO: 'FFC8E6C9', // verde claro (=0)
  CF_SALDO_POS:  'FFFFF3CC', // amarelo claro (>0)
  CF_SALDO_NEG:  'FFFFCCCC', // vermelho claro (<0)
  CF_ENVIADO:    'FFE8F5E9', // verde muito claro (>0)
};

function bordaFina() {
  const s = { style: 'thin', color: { argb: COR.BORDA } };
  return { top: s, left: s, bottom: s, right: s };
}

function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// Fonte Arial 8pt (tamanho extraído do modelo)
function fonte(opts = {}) {
  return {
    name: 'Arial',
    size: opts.size ?? 9,
    color: { argb: opts.color ?? COR.FONTE_NORMAL },
    bold: opts.bold ?? false,
  };
}

async function gerarExcel(cab, produtos) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`Pedido ${cab.pedido}`);

  // ── Larguras exatas do modelo 2.0 (colunas A-M = 13 colunas)
  // Col:  A     B     C      D    E     F    G    H     I     J     K     L     M
  const larguras = [2.71, 7.0, 38.86, 7.0, 6.14, 8.0, 3.14, 6.71, 6.71, 6.71, 6.71, 7.71, 6.29];
  larguras.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── ROW 1: Título — merge A1:M1
  ws.mergeCells('A1:M1');
  const titulo = ws.getCell('A1');
  titulo.value = `CONFERENCIA DE PRODUTOS  —  Pedido ${cab.pedido}`;
  titulo.font = fonte({ bold: true, size: 12, color: COR.FONTE_TITULO });
  titulo.fill = fill(COR.TITULO_FILL);
  titulo.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  titulo.border = bordaFina();
  ws.getRow(1).height = 20.25;

  // ── ROWS 2-7: Campos de cabeçalho
  // Estrutura do modelo 2.0: A:B=label, C:G=valor, H:I=label, J:M=valor
  const campos = [
    ['Pedido No:',     cab.pedido,        'Cliente:',         cab.cliente],
    ['Data Emissao:',  cab.data_emissao,  'Endereco:',        cab.endereco],
    ['Tipo Frete:',    cab.tipo_frete,    'Cond. Pgt:',       cab.cond_pgt],
    ['CNPJ/CPF:',      cab.cnpj,          'Representante:',   cab.representante],
    ['Volume:',        cab.volume ?? '',  'Transportadora:',  cab.transportadora],
    ['NF Fiscal:',     cab.nf ?? '',      'Data Envio:',      cab.data_envio ?? ''],
  ];

  for (let i = 0; i < campos.length; i++) {
    const row = i + 2;
    const [l1, v1, l2, v2] = campos[i];

    // Label esquerdo (A:B)
    ws.mergeCells(`A${row}:B${row}`);
    Object.assign(ws.getCell(`A${row}`), {
      value: l1,
      font: fonte({ bold: true, color: COR.FONTE_HEADER }),
      fill: fill(COR.HEADER_FILL),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    // Valor esquerdo (C:G)
    ws.mergeCells(`C${row}:G${row}`);
    Object.assign(ws.getCell(`C${row}`), {
      value: v1,
      font: fonte(),
      fill: fill(COR.BRANCO),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    // Label direito (H:I)
    ws.mergeCells(`H${row}:I${row}`);
    Object.assign(ws.getCell(`H${row}`), {
      value: l2,
      font: fonte({ bold: true, color: COR.FONTE_HEADER }),
      fill: fill(COR.HEADER_FILL),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    // Valor direito (J:M) — modelo 2.0 usa M como última coluna
    ws.mergeCells(`J${row}:M${row}`);
    Object.assign(ws.getCell(`J${row}`), {
      value: v2,
      font: fonte(),
      fill: fill(COR.BRANCO),
      border: bordaFina(),
      alignment: { horizontal: 'left', vertical: 'center', indent: 1, wrapText: true },
    });

    ws.getRow(row).height = 20.25;
  }

  // ── ROW 8: Header da tabela (13 colunas, igual ao modelo 2.0)
  // Nomes abreviados conforme modelo: VALID., PICK., 1a R., 2a R., etc.
  const hdrs = ['#', 'CODIGO', 'DESCRICAO', 'LOTE', 'VALIDADE', 'PICKING', 'OK',
                '1a R.', '2a R.', '3a R.', '4a R.', 'ENVIADO', 'SALDO'];
  const headerRow = 8;
  hdrs.forEach((h, ci) => {
    const c = ws.getCell(headerRow, ci + 1);
    c.value = h;
    c.font = fonte({ bold: true, color: COR.FONTE_HEADER });
    c.fill = fill(COR.TOTAL_FILL);
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(headerRow).height = 20.25;

  // ── ROWS 9+: Linhas de produto
  const dataStart = 9;
  produtos.forEach((p, idx) => {
    const row = dataStart + idx;
    const bgColor = idx % 2 === 0 ? COR.BRANCO : COR.ZEBRA;

    // Colunas A-K: #, codigo, descricao, lote, validade, qtde, OK, 1aR, 2aR, 3aR, 4aR
    const vals = [
      idx + 1, p.cod, p.descricao, p.lote, p.validade,
      p.qtde, '', null, null, null, null
    ];

    vals.forEach((val, ci) => {
      const c = ws.getCell(row, ci + 1);
      c.value = val;
      c.font = fonte();
      c.fill = fill(bgColor);
      c.border = bordaFina();
      // Coluna C (descrição) alinha à esquerda com wrap; demais centralizado
      c.alignment = ci === 2
        ? { vertical: 'center', wrapText: true }
        : { horizontal: 'center', vertical: 'center', wrapText: true };
    });

    // Coluna L (12): ENVIADO = SUM(H:K)
    const cEnv = ws.getCell(row, 12);
    cEnv.value = { formula: `SUM(H${row}:K${row})` };
    cEnv.font = fonte();
    cEnv.fill = fill(bgColor);
    cEnv.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    cEnv.border = bordaFina();

    // Coluna M (13): SALDO = PICK - ENVIADO
    const cSal = ws.getCell(row, 13);
    cSal.value = { formula: `F${row}-L${row}` };
    cSal.font = fonte();
    cSal.fill = fill(bgColor);
    cSal.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    cSal.border = bordaFina();

    // Altura automática (não fixada) para adaptar ao conteúdo com wrap
    ws.getRow(row).height = 20.25;
  });

  const dataEnd = dataStart + produtos.length - 1;
  const totalRow = dataEnd + 1;

  // ── Linha TOTAL GERAL — merge A:E, demais colunas com SUM
  ws.mergeCells(`A${totalRow}:E${totalRow}`);
  const cTotalLabel = ws.getCell(`A${totalRow}`);
  cTotalLabel.value = 'TOTAL GERAL';
  cTotalLabel.font = fonte({ bold: true, color: COR.FONTE_HEADER });
  cTotalLabel.fill = fill(COR.TOTAL_FILL);
  cTotalLabel.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  cTotalLabel.border = bordaFina();

  // Colunas F, H-M recebem SUM; G fica vazio
  const colsTotal = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
  const colsSumTotal = ['F', 'H', 'I', 'J', 'K', 'L', 'M'];
  colsTotal.forEach((col) => {
    const ci = col.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
    const c = ws.getCell(totalRow, ci);
    if (colsSumTotal.includes(col)) {
      c.value = { formula: `SUM(${col}${dataStart}:${col}${dataEnd})` };
    }
    c.font = fonte({ bold: true, color: COR.FONTE_HEADER });
    c.fill = fill(COR.TOTAL_FILL);
    c.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    c.border = bordaFina();
  });
  ws.getRow(totalRow).height = 20.25;

  // ── Área de OBSERVAÇÕES: 1 linha de título + 10 linhas vazias mergeadas A:M
  const obsHeaderRow = totalRow + 1;
  ws.mergeCells(`A${obsHeaderRow}:L${obsHeaderRow}`);
  const cObsLabel = ws.getCell(`A${obsHeaderRow}`);
  cObsLabel.value = 'OBSERVAÇÕES';
  cObsLabel.font = fonte({ bold: true, color: COR.FONTE_HEADER });
  cObsLabel.fill = fill(COR.TOTAL_FILL);
  cObsLabel.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  cObsLabel.border = {
    top: { style: 'thin', color: { argb: COR.BORDA } },
    left: { style: 'thin', color: { argb: COR.BORDA } },
    bottom: { style: 'thin', color: { argb: COR.BORDA } },
    right: { style: 'none' },
  };
  // Célula M do header de obs (borda direita)
  const cObsLabelRight = ws.getCell(`M${obsHeaderRow}`);
  cObsLabelRight.font = fonte({ bold: true, color: COR.FONTE_HEADER });
  cObsLabelRight.fill = fill(COR.TOTAL_FILL);
  cObsLabelRight.border = {
    top: { style: 'thin', color: { argb: COR.BORDA } },
    left: { style: 'none' },
    bottom: { style: 'thin', color: { argb: COR.BORDA } },
    right: { style: 'thin', color: { argb: COR.BORDA } },
  };
  ws.getRow(obsHeaderRow).height = 20.25;

  // 10 linhas de observações — cada linha mergeada A:M, borda apenas top/bottom/left/right do bloco
  for (let i = 0; i < 10; i++) {
    const r = obsHeaderRow + 1 + i;
    ws.mergeCells(`A${r}:M${r}`);
    const c = ws.getCell(`A${r}`);
    c.font = fonte();
    c.fill = fill(COR.BRANCO);
    c.alignment = { vertical: 'top', wrapText: true };
    // Borda fina em todas para manter linhas visíveis
    c.border = {
      top: { style: 'thin', color: { argb: COR.BORDA } },
      left: { style: 'thin', color: { argb: COR.BORDA } },
      bottom: { style: 'thin', color: { argb: COR.BORDA } },
      right: { style: 'thin', color: { argb: COR.BORDA } },
    };
    ws.getRow(r).height = 20.25;
  }

  // ── Formatação condicional SALDO (coluna M) — mesma do modelo 2.0
  ws.addConditionalFormatting({
    ref: `M${dataStart}:M${dataEnd}`,
    rules: [
      {
        type: 'cellIs', operator: 'equal',
        formulae: ['0'], priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_ZERO },
                  fgColor: { argb: COR.CF_SALDO_ZERO } }
        }
      },
      {
        type: 'cellIs', operator: 'greaterThan',
        formulae: ['0'], priority: 2,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_POS },
                  fgColor: { argb: COR.CF_SALDO_POS } }
        }
      },
      {
        type: 'cellIs', operator: 'lessThan',
        formulae: ['0'], priority: 3,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_SALDO_NEG },
                  fgColor: { argb: COR.CF_SALDO_NEG } }
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
        formulae: ['0'], priority: 4,
        style: {
          fill: { type: 'pattern', pattern: 'solid',
                  bgColor: { argb: COR.CF_ENVIADO },
                  fgColor: { argb: COR.CF_ENVIADO } }
        }
      },
    ],
  });

  // ── Freeze nas linhas de cabeçalho (linha 8, freeze abaixo dela)
  ws.views = [{ state: 'frozen', ySplit: headerRow }];

  // ── Configuração de impressão retrato A4
  // scale=83 é o valor exato do modelo 2.0 para caber em 1 folha
  ws.pageSetup = {
    orientation: 'portrait',
    paperSize: 9,           // A4
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,         // forçar 1 página de altura também
    scale: 100,              // escala exata do modelo (100%)
    printTitlesRow: `1:${headerRow}`,
    margins: {
      left: 0.4, right: 0.4,
      top: 0.6, bottom: 0.6,
      header: 0.3, footer: 0.3,
    },
  };

  // ── Quebra de texto automático: garantir que todas as linhas de dados
  // tenham wrapText=true já definido acima; altura 'auto' via defaultRowHeight
  ws.properties = { defaultRowHeight: 20.25 };

  return wb.xlsx.writeBuffer();
}

module.exports = { gerarExcel };

