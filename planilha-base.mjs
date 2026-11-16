/* ============================================================
   Gera a planilha base do casamento a partir do dados.json.

       node planilha-base.mjs                 -> base-de-dados.xlsx
       node planilha-base.mjs minha.xlsx      -> outro nome

   Tres abas: Fornecedores, Parcelas e Tarefas.

   A aba Fornecedores tem uma coluna "Confere" que refaz a regra de
   ouro dentro do proprio Excel: contratado menos pago menos a soma
   das parcelas lancadas tem de dar zero. Se um dia der diferente,
   a planilha acusa sozinha, sem depender de ninguem conferir.

   Nenhum total e gravado como numero. Todo total e formula SUM, e
   "A pagar" e formula tambem. A planilha original mostrava valores
   velhos em cache justamente por nao fazer isso.

   Este arquivo nao contem dado nenhum: tudo vem do dados.json, que
   fica fora do git.
   ============================================================ */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const SAIDA = path.join(AQUI, process.argv[2] || 'base-de-dados.xlsx');

// A planilha original e so referencia historica. Nunca e sobrescrita.
if (/casamento\.xlsx$/i.test(path.basename(SAIDA))) {
  console.error('Recusado: CASAMENTO.xlsx e a planilha original e nao pode ser sobrescrita.');
  process.exit(1);
}

const caminhoDados = path.join(AQUI, 'dados.json');
if (!fs.existsSync(caminhoDados)) {
  console.error('Falta o dados.json nesta pasta. Ele e a fonte da verdade e nao vai para o git.');
  process.exit(1);
}
const dados = JSON.parse(fs.readFileSync(caminhoDados, 'utf8'));

/* ---------- ZIP sem compressao, que o Excel aceita ---------- */
const TABELA_CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
  return t;
})();
const crc32 = b => {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < b.length; i++) c = TABELA_CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
};

function zipar(arquivos) {
  const agora = new Date();
  const horaDos = (agora.getHours() << 11) | (agora.getMinutes() << 5) | (agora.getSeconds() >> 1);
  const dataDos = ((agora.getFullYear() - 1980) << 9) | ((agora.getMonth() + 1) << 5) | agora.getDate();
  const corpo = [], central = [];
  let desloc = 0, totalCentral = 0;

  for (const a of arquivos) {
    const nome = Buffer.from(a.nome, 'utf8');
    const dd = Buffer.from(a.texto, 'utf8');
    const soma = crc32(dd);

    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(0, 8);
    lh.writeUInt16LE(horaDos, 10); lh.writeUInt16LE(dataDos, 12);
    lh.writeUInt32LE(soma, 14); lh.writeUInt32LE(dd.length, 18); lh.writeUInt32LE(dd.length, 22);
    lh.writeUInt16LE(nome.length, 26); lh.writeUInt16LE(0, 28);
    corpo.push(lh, nome, dd);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(0x0800, 8); cd.writeUInt16LE(0, 10);
    cd.writeUInt16LE(horaDos, 12); cd.writeUInt16LE(dataDos, 14);
    cd.writeUInt32LE(soma, 16); cd.writeUInt32LE(dd.length, 20); cd.writeUInt32LE(dd.length, 24);
    cd.writeUInt16LE(nome.length, 28); cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34); cd.writeUInt16LE(0, 36); cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(desloc, 42);
    central.push(cd, nome);

    desloc += 30 + nome.length + dd.length;
    totalCentral += 46 + nome.length;
  }

  const fim = Buffer.alloc(22);
  fim.writeUInt32LE(0x06054b50, 0);
  fim.writeUInt16LE(0, 4); fim.writeUInt16LE(0, 6);
  fim.writeUInt16LE(arquivos.length, 8); fim.writeUInt16LE(arquivos.length, 10);
  fim.writeUInt32LE(totalCentral, 12); fim.writeUInt32LE(desloc, 16);
  fim.writeUInt16LE(0, 20);

  return Buffer.concat([...corpo, ...central, fim]);
}

/* ---------- celulas ---------- */
const P = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const vazio = (ref, est) => '<c r="' + ref + '" s="' + est + '"/>';
// Texto vazio vira celula vazia de verdade. Um <is><t></t></is> em branco
// e aceito pelo esquema, mas e lixo que nao precisa existir.
const txt = (ref, est, v) => (v === '' || v == null)
  ? vazio(ref, est)
  : '<c r="' + ref + '" s="' + est + '" t="inlineStr"><is><t>' + esc(v) + '</t></is></c>';
const num = (ref, est, v) => '<c r="' + ref + '" s="' + est + '"><v>' + v + '</v></c>';
const frm = (ref, est, f) => '<c r="' + ref + '" s="' + est + '"><f>' + esc(f) + '</f></c>';
const vaz = vazio;

// Excel conta os dias desde 30/12/1899. Data vai como numero de verdade,
// e nao como texto, senao nao da para ordenar nem filtrar por periodo.
const serial = iso => {
  const [a, m, d] = iso.split('-').map(Number);
  return Math.round((Date.UTC(a, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
};

const ESTILOS = {
  normal: 0, titulo: 1, legenda: 2,
  cabLeft: 3, cabRight: 4, cabCentro: 5,
  celTexto: 6, celMoeda: 7, celData: 8, celCentro: 9, celNumero: 10,
  totTexto: 11, totMoeda: 12, totNumero: 13,
  aviso: 14,
};

/* ---------- monta uma aba ---------- */
function folha({ titulo, legenda, colunas, cabecalhos, linhas, totais, larguras, congelar = 4 }) {
  const L0 = 5;
  const L1 = L0 + linhas.length - 1;
  const LT = L1 + 1;
  const letra = i => String.fromCharCode(65 + i);
  const ultima = letra(colunas - 1);

  let sd = '';
  sd += '<row r="1" ht="26" customHeight="1">' + txt('A1', ESTILOS.titulo, titulo) + '</row>';
  sd += '<row r="2" ht="16" customHeight="1">' + txt('A2', ESTILOS.legenda, legenda) + '</row>';
  sd += '<row r="3"/>';
  sd += '<row r="4" ht="24" customHeight="1">' +
    cabecalhos.map((c, i) => txt(letra(i) + '4', c.est, c.rotulo)).join('') + '</row>';

  linhas.forEach((celulas, i) => {
    const r = L0 + i;
    sd += '<row r="' + r + '">' + celulas.map((c, j) => c(letra(j) + r, r)).join('') + '</row>';
  });

  if (totais) {
    sd += '<row r="' + LT + '" ht="24" customHeight="1">' +
      totais.map((t, j) => t(letra(j) + LT, letra(j), L0, L1)).join('') + '</row>';
  }

  const fim = totais ? LT : L1;

  /* A ordem destes elementos nao e livre. O CT_Worksheet do OOXML e uma
     xsd:sequence, entao sheetPr, dimension, sheetViews, sheetFormatPr, cols,
     sheetData, autoFilter, mergeCells, pageMargins e pageSetup tem de sair
     exatamente nesta ordem. O Excel recusa o arquivo se trocar, com a caixa
     de "encontramos um problema em um conteudo". Leitor de biblioteca releva
     e nao acusa, entao isso so aparece abrindo no Excel de verdade. */
  return P +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    '<dimension ref="A1:' + ultima + fim + '"/>' +
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0">' +
    '<pane ySplit="' + congelar + '" topLeftCell="A' + L0 + '" activePane="bottomLeft" state="frozen"/>' +
    '<selection pane="bottomLeft" activeCell="A' + L0 + '" sqref="A' + L0 + '"/>' +
    '</sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="16"/>' +
    '<cols>' + larguras.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>' +
    '<sheetData>' + sd + '</sheetData>' +
    '<autoFilter ref="A4:' + ultima + L1 + '"/>' +
    '<mergeCells count="2"><mergeCell ref="A1:' + ultima + '1"/><mergeCell ref="A2:' + ultima + '2"/></mergeCells>' +
    '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
    '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
    '</worksheet>';
}

/* ---------- dados ---------- */
const forn = dados.fornecedores;
const parcelas = forn.flatMap(f => (f.par || []).map(p => ({ forn: f.n, ...p })));
const hoje = new Date().toLocaleDateString('pt-BR');

/* aba 1: Fornecedores */
const abaFornecedores = folha({
  titulo: 'Fornecedores',
  legenda: 'Conferido em ' + dados.atualizado + '. Planilha gerada em ' + hoje +
           '. A coluna Confere tem de ficar zerada em todas as linhas.',
  colunas: 6,
  larguras: [34, 15, 15, 15, 17, 11],
  cabecalhos: [
    { rotulo: 'Fornecedor', est: ESTILOS.cabLeft },
    { rotulo: 'Contratado', est: ESTILOS.cabRight },
    { rotulo: 'Pago', est: ESTILOS.cabRight },
    { rotulo: 'A pagar', est: ESTILOS.cabRight },
    { rotulo: 'Parcelas lançadas', est: ESTILOS.cabRight },
    { rotulo: 'Confere', est: ESTILOS.cabCentro },
  ],
  linhas: forn.map(f => [
    (ref) => txt(ref, ESTILOS.celTexto, f.n),
    (ref) => num(ref, ESTILOS.celMoeda, f.c),
    (ref) => num(ref, ESTILOS.celMoeda, f.p),
    (ref, r) => frm(ref, ESTILOS.celMoeda, 'B' + r + '-C' + r),
    (ref, r) => frm(ref, ESTILOS.celMoeda, 'SUMIF(Parcelas!$A$5:$A$500,A' + r + ',Parcelas!$B$5:$B$500)'),
    (ref, r) => frm(ref, ESTILOS.celMoeda, 'D' + r + '-E' + r),
  ]),
  totais: [
    (ref) => txt(ref, ESTILOS.totTexto, 'Total'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
  ],
});

/* aba 2: Parcelas */
const ordenadas = [...parcelas].sort((a, b) => {
  if (!a.d && !b.d) return a.forn.localeCompare(b.forn, 'pt-BR');
  if (!a.d) return 1;
  if (!b.d) return -1;
  return a.d.localeCompare(b.d);
});
const abaParcelas = folha({
  titulo: 'Parcelas a pagar',
  legenda: ordenadas.length + ' parcelas, sendo ' + ordenadas.filter(p => !p.d).length +
           ' ainda sem data combinada. O nome do fornecedor tem de ser igual ao da aba Fornecedores.',
  colunas: 5,
  larguras: [34, 15, 15, 13, 30],
  cabecalhos: [
    { rotulo: 'Fornecedor', est: ESTILOS.cabLeft },
    { rotulo: 'Valor', est: ESTILOS.cabRight },
    { rotulo: 'Vencimento', est: ESTILOS.cabRight },
    { rotulo: 'Forma', est: ESTILOS.cabCentro },
    { rotulo: 'Observação', est: ESTILOS.cabLeft },
  ],
  linhas: ordenadas.map(p => [
    (ref) => txt(ref, ESTILOS.celTexto, p.forn),
    (ref) => num(ref, ESTILOS.celMoeda, p.v),
    (ref) => p.d ? num(ref, ESTILOS.celData, serial(p.d)) : vaz(ref, ESTILOS.celData),
    (ref) => txt(ref, ESTILOS.celCentro, p.cartao ? 'Cartão' : 'À vista'),
    (ref) => txt(ref, p.d ? ESTILOS.celTexto : ESTILOS.aviso, p.d ? '' : 'Data a combinar'),
  ]),
  totais: [
    (ref) => txt(ref, ESTILOS.totTexto, 'Total'),
    (ref, c, a, b) => frm(ref, ESTILOS.totMoeda, 'SUM(' + c + a + ':' + c + b + ')'),
    (ref) => vaz(ref, ESTILOS.totTexto),
    (ref) => vaz(ref, ESTILOS.totTexto),
    (ref) => vaz(ref, ESTILOS.totTexto),
  ],
});

/* aba 3: Tarefas */
const tarefas = [...dados.tarefas].sort((a, b) => {
  if (!a.q && !b.q) return 0;
  if (!a.q) return 1;
  if (!b.q) return -1;
  return a.q.localeCompare(b.q);
});
const abaTarefas = folha({
  titulo: 'Tarefas',
  legenda: tarefas.length + ' tarefas. Marque Feito com um x quando resolver.',
  colunas: 3,
  larguras: [70, 15, 10],
  cabecalhos: [
    { rotulo: 'Tarefa', est: ESTILOS.cabLeft },
    { rotulo: 'Prazo', est: ESTILOS.cabRight },
    { rotulo: 'Feito', est: ESTILOS.cabCentro },
  ],
  linhas: tarefas.map(t => [
    (ref) => txt(ref, ESTILOS.celTexto, t.t),
    (ref) => t.q ? num(ref, ESTILOS.celData, serial(t.q)) : vaz(ref, ESTILOS.celData),
    (ref) => vaz(ref, ESTILOS.celCentro),
  ]),
});

/* ---------- estilos ---------- */
const VINHO = 'FF7A1236', CREME = 'FFF7EDE7', LINHA = 'FFE6D5CA', TINTA = 'FF3A3338', CINZA = 'FF8A7F86', AMBAR = 'FF9A5B1E';

const estilos = P +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="2">' +
  '<numFmt numFmtId="164" formatCode="&quot;R$&quot;\\ #,##0.00"/>' +
  '<numFmt numFmtId="165" formatCode="dd/mm/yyyy"/>' +
  '</numFmts>' +
  '<fonts count="6">' +
  '<font><sz val="11"/><color rgb="' + TINTA + '"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="11"/><color rgb="' + VINHO + '"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="16"/><color rgb="' + VINHO + '"/><name val="Calibri"/></font>' +
  '<font><i/><sz val="9"/><color rgb="' + CINZA + '"/><name val="Calibri"/></font>' +
  '<font><sz val="10"/><color rgb="' + AMBAR + '"/><name val="Calibri"/></font>' +
  '</fonts>' +
  '<fills count="4">' +
  '<fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="' + VINHO + '"/><bgColor indexed="64"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="' + CREME + '"/><bgColor indexed="64"/></patternFill></fill>' +
  '</fills>' +
  '<borders count="3">' +
  '<border><left/><right/><top/><bottom/><diagonal/></border>' +
  '<border><left/><right/><top/><bottom style="thin"><color rgb="' + LINHA + '"/></bottom><diagonal/></border>' +
  '<border><left/><right/><top style="thin"><color rgb="' + VINHO + '"/></top><bottom/><diagonal/></border>' +
  '</borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="15">' +
  /*  0 normal        */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  /*  1 titulo        */ '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
  /*  2 legenda       */ '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
  /*  3 cab esquerda  */ '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
  /*  4 cab direita   */ '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>' +
  /*  5 cab centro    */ '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
  /*  6 texto         */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
  /*  7 moeda         */ '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
  /*  8 data          */ '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>' +
  /*  9 centro        */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
  /* 10 numero        */ '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
  /* 11 total texto   */ '<xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
  /* 12 total moeda   */ '<xf numFmtId="164" fontId="2" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
  /* 13 total numero  */ '<xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
  /* 14 aviso         */ '<xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
  '</cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>';

/* ---------- empacota ---------- */
const abas = [
  { nome: 'Fornecedores', xml: abaFornecedores },
  { nome: 'Parcelas', xml: abaParcelas },
  { nome: 'Tarefas', xml: abaTarefas },
];

const arquivos = [
  { nome: '[Content_Types].xml', texto: P +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    abas.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>' },
  { nome: '_rels/.rels', texto: P +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>' },
  { nome: 'xl/workbook.xml', texto: P +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets>' + abas.map((a, i) => '<sheet name="' + esc(a.nome) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets>' +
    // Sem isto o Excel recalcula so o que ele julga necessario, e formula
    // recem-escrita por script as vezes aparece vazia na primeira abertura.
    '<calcPr calcId="191029" fullCalcOnLoad="1"/>' +
    '</workbook>' },
  { nome: 'xl/_rels/workbook.xml.rels', texto: P +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    abas.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
    '<Relationship Id="rId' + (abas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>' },
  { nome: 'xl/styles.xml', texto: estilos },
  ...abas.map((a, i) => ({ nome: 'xl/worksheets/sheet' + (i + 1) + '.xml', texto: a.xml })),
];

/* Guarda o arquivo anterior antes de sobrescrever.
   Quem edita a planilha pode ter formatado, somado uma coluna ou deixado
   uma anotacao, e gerar de novo apaga tudo isso sem avisar. Uma copia
   custa nada e ja evitou perda uma vez. */
const bytes = zipar(arquivos);
if (fs.existsSync(SAIDA)) {
  const anterior = SAIDA.replace(/\.xlsx$/i, '.anterior.xlsx');
  fs.copyFileSync(SAIDA, anterior);
  console.log('Copia do arquivo anterior em ' + path.basename(anterior));
}
fs.writeFileSync(SAIDA, bytes);

/* ---------- confere o que acabou de gravar ---------- */
const soma = forn.reduce((a, f) => ({ c: a.c + f.c, p: a.p + f.p }), { c: 0, p: 0 });
const somaParcelas = parcelas.reduce((a, p) => a + p.v, 0);
const brl = v => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

console.log('Planilha gravada: ' + path.basename(SAIDA) + '  (' + bytes.length.toLocaleString('pt-BR') + ' bytes, ' + arquivos.length + ' partes)');
console.log('  Fornecedores: ' + forn.length + '   Parcelas: ' + parcelas.length + '   Tarefas: ' + dados.tarefas.length);
console.log('  Contratado ' + brl(soma.c) + ' | Pago ' + brl(soma.p) + ' | A pagar ' + brl(soma.c - soma.p));
if (somaParcelas !== soma.c - soma.p) {
  console.error('  ATENCAO: as parcelas somam ' + brl(somaParcelas) + ', diferente do a pagar. A coluna Confere vai acusar.');
  process.exitCode = 1;
} else {
  console.log('  Regra de ouro: as parcelas somam exatamente o a pagar. A coluna Confere fica zerada.');
}
