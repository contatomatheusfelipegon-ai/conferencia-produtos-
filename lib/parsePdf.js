/**
 * lerPdf.js — Parser único Node.js para pedidos TOTVS/SIGA
 *
 * Embute o script Python inline e o executa via child_process,
 * sem precisar de um arquivo .py separado no projeto.
 *
 * Requer apenas: pip install pdfplumber
 */

const { execFileSync } = require('child_process');
const { writeFileSync, unlinkSync } = require('fs');
const { tmpdir } = require('os');
const path = require('path');

// ─── Script Python embutido ───────────────────────────────────────────────────
const PYTHON_SCRIPT = `
import sys, re, json

try:
    import pdfplumber
except ImportError:
    print(json.dumps({"error": "pdfplumber nao instalado. Execute: pip install pdfplumber"}))
    sys.exit(1)

def dedup_chars(texto):
    if not texto: return texto
    amostra = re.sub(r'\\s', '', texto)[:20]
    if len(amostra) < 6: return texto
    pares = sum(1 for i in range(0, len(amostra)-1, 2) if amostra[i] == amostra[i+1])
    if pares >= 3:
        return re.sub(r'(.)\\1', r'\\1', texto)
    return texto

def extrair_cabecalho(txt):
    def m(pat, flags=0):
        r = re.search(pat, txt, flags)
        return r.group(1).strip() if r else ''
    return {
        'pedido':         m(r'Pedido de Venda\\s*[-\\u2013]\\s*(\\d+)'),
        'data_emissao':   m(r'Data Emiss[a\\u00e3]o\\s*:\\s*(\\d{2}/\\d{2}/\\d{4})'),
        'cliente':        m(r'Cliente\\s*:\\s*(.+?)\\s*\\(M\\d+'),
        'endereco':       m(r'Endere[\\u00e7c]o\\s*:\\s*(.+?)(?=\\n)'),
        'tipo_frete':     m(r'Tipo Frete\\s*:\\s*(\\w+)'),
        'cond_pgt':       m(r'Cond\\.Pgt\\.\\s*:\\s*\\d+\\s*-\\s*(.+?)(?=\\n)'),
        'cnpj':           m(r'CNPJ/CPF\\s*:\\s*([\\d./\\-]+)'),
        'representante':  m(r'Repres\\.\\s*:\\s*\\w+\\s*-\\s*(.+?)(?=\\n)'),
        'volume':         m(r'Volume\\s*:\\s*(\\d+\\s*VOLUMES?)', re.I),
        'transportadora': m(r'Transp\\.\\s*:\\s*(.+?)(?=\\n)'),
    }

def limpar(desc):
    desc = re.sub(r'\\s+', ' ', desc).strip()
    desc = re.sub(r'(\\d)\\s+,(\\d)', r'\\1,\\2', desc)
    return desc

PAD      = re.compile(r'^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s+([A-Z0-9]{6,8}[A-Z]?)\\s+(\\d{2}/\\d{2})\\s+[\\d,]+\\s+[\\d*,]+')
PAD_SL   = re.compile(r'^(\\d+)\\s+(PN\\w+)\\s+(.+?)\\s+([\\d,]+)\\s+([\\d,]+)\\s*$')
PAD_IGN  = re.compile(r'^[-_]{5,}|^Qtde\\s|^Folha\\.|^SIGA\\s|^Hora\\.|^Data\\s|^Cliente\\s|^Endere|^Tipo\\s|^CNPJ|^Repres\\.|^Volume\\s|^Nat\\.|^Frete|^\\d+o\\s+Vencto|^Transp\\.|^Ender\\.|^NUM\\.CXA|\\(cid|^\\d{3,}\\s+[\\d.,]|^SANTO\\s|^RUA\\s|^GUARULHOS|^RODOVIA')
PAD_CONT = re.compile(r"^[A-Za-z0-9\\s\\-,./\\'()]+$")

def extrair_produtos(txt):
    linhas = txt.split('\\n')
    raw = []
    i = 0
    while i < len(linhas):
        ln = linhas[i].strip(); i += 1
        if not ln or PAD_IGN.search(ln): continue

        m = PAD.match(ln)
        if m:
            qtde, cod, desc, lote, val = int(m.group(1)), m.group(2), m.group(3).strip(), m.group(4), m.group(5)
            if i < len(linhas):
                nx = linhas[i].strip()
                if nx and PAD_CONT.match(nx) and not PAD.match(nx) and not PAD_IGN.search(nx):
                    desc += ' ' + nx; i += 1
            raw.append({'qtde': qtde, 'cod': cod, 'descricao': limpar(desc), 'lote': lote, 'validade': val})
            continue

        m2 = PAD_SL.match(ln)
        if m2:
            qtde, cod, desc = int(m2.group(1)), m2.group(2), m2.group(3).strip()
            if re.match(r'^[\\d.,]+$', desc): continue
            if i < len(linhas):
                nx = linhas[i].strip()
                if nx and PAD_CONT.match(nx) and not PAD.match(nx) and not PAD_IGN.search(nx) and not re.match(r'^\\d+$', nx):
                    desc += ' ' + nx; i += 1
            raw.append({'qtde': qtde, 'cod': cod, 'descricao': limpar(desc), 'lote': '', 'validade': ''})

    grupos = {}
    for p in raw:
        g = grupos.setdefault(p['cod'], {'qtde': 0, 'lotes': [], 'validades': [], 'descricao': ''})
        g['qtde'] += p['qtde']
        g['descricao'] = p['descricao']
        if p['lote']    and p['lote']    not in g['lotes']:    g['lotes'].append(p['lote'])
        if p['validade'] and p['validade'] not in g['validades']: g['validades'].append(p['validade'])

    return [{'cod': c, 'descricao': d['descricao'], 'lote': ' / '.join(d['lotes']),
             'validade': ' / '.join(d['validades']), 'qtde': d['qtde']} for c, d in grupos.items()]

def ler_pdf(caminho):
    with pdfplumber.open(caminho) as pdf:
        vistas, textos = set(), []
        for page in pdf.pages:
            txt = dedup_chars(page.extract_text() or '')
            fm = re.search(r'Folha\\.\\.: (\\d+)', txt)
            fid = fm.group(1) if fm else str(page.page_number)
            if fid in vistas: continue
            vistas.add(fid); textos.append(txt)
        txt = '\\n'.join(textos)
        return {'cab': extrair_cabecalho(txt), 'produtos': extrair_produtos(txt)}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Uso: python3 <script> <pdf>"})); sys.exit(1)
    try:
        print(json.dumps(ler_pdf(sys.argv[1]), ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)})); sys.exit(1)
`;

// ─── Função principal exportada ───────────────────────────────────────────────
async function lerPdf(buffer) {
  // Grava PDF temp
  const tmpPdf = path.join(tmpdir(), `ped_${Date.now()}.pdf`);
  // Grava script Python temp
  const tmpPy  = path.join(tmpdir(), `lerPdf_${Date.now()}.py`);

  try {
    writeFileSync(tmpPdf, buffer);
    writeFileSync(tmpPy,  PYTHON_SCRIPT, 'utf8');

    const stdout = execFileSync('python3', [tmpPy, tmpPdf], {
      timeout:   30000,
      maxBuffer: 10 * 1024 * 1024,
      encoding:  'utf8',
    });

    const resultado = JSON.parse(stdout);
    if (resultado.error) throw new Error(resultado.error);
    return resultado;

  } finally {
    try { unlinkSync(tmpPdf); } catch (_) {}
    try { unlinkSync(tmpPy);  } catch (_) {}
  }
}

module.exports = { lerPdf };
