/* ============================================================
   Confere se um .xlsx e valido do jeito que o EXCEL cobra.

       node verifica-xlsx.mjs base-de-dados.xlsx

   Existe por um motivo especifico. O openpyxl, o navegador e o
   descompactador do Windows abriram um arquivo que o Excel recusou
   com "encontramos um problema em um conteudo". A causa era ordem
   de elemento: no esquema do OOXML varios elementos sao xsd:sequence,
   ou seja, a ordem e obrigatoria, e so o Excel cobra isso. Leitor de
   biblioteca reordena calado e nao acusa nada.

   Entao aqui a ordem e conferida contra a sequencia oficial.
   ============================================================ */
import fs from 'node:fs';
import zlib from 'node:zlib';
import path from 'node:path';

const alvo = process.argv[2];
if (!alvo) { console.error('Uso: node verifica-xlsx.mjs <arquivo.xlsx>'); process.exit(1); }
if (!fs.existsSync(alvo)) { console.error('Nao achei ' + alvo); process.exit(1); }

const falhas = [];
const ok = (c, m, extra = '') => {
  console.log((c ? '  [ok]  ' : '  [FALHA] ') + m + (extra ? '  ' + extra : ''));
  if (!c) falhas.push(m);
};

/* ---------- abre o ZIP ---------- */
function abrirZip(buf) {
  // Le pelo diretorio central, que e a fonte de verdade de um ZIP.
  let fim = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 65557; i--)
    if (buf.readUInt32LE(i) === 0x06054b50) { fim = i; break; }
  if (fim < 0) throw new Error('nao achei o fim do diretorio central');
  const total = buf.readUInt16LE(fim + 10);
  let p = buf.readUInt32LE(fim + 16);
  const partes = new Map();
  for (let n = 0; n < total; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('cabecalho central ' + n + ' corrompido');
    const metodo = buf.readUInt16LE(p + 10);
    const crcEsperado = buf.readUInt32LE(p + 16);
    const comp = buf.readUInt32LE(p + 20);
    const nomeLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const comentLen = buf.readUInt16LE(p + 32);
    const desloc = buf.readUInt32LE(p + 42);
    const nome = buf.toString('utf8', p + 46, p + 46 + nomeLen);

    if (buf.readUInt32LE(desloc) !== 0x04034b50) throw new Error(nome + ': cabecalho local corrompido');
    const lNome = buf.readUInt16LE(desloc + 26);
    const lExtra = buf.readUInt16LE(desloc + 28);
    const ini = desloc + 30 + lNome + lExtra;
    const cru = buf.subarray(ini, ini + comp);
    const dados = metodo === 0 ? cru : zlib.inflateRawSync(cru);
    partes.set(nome, { texto: dados.toString('utf8'), bytes: dados.length, crcEsperado, dados });
    p += 46 + nomeLen + extraLen + comentLen;
  }
  return partes;
}

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

/* ---------- filhos diretos de um elemento ---------- */
function filhos(xml, raiz) {
  const abre = new RegExp('<' + raiz + '[\\s>]');
  const i = xml.search(abre);
  if (i < 0) return null;
  const saida = [];
  let prof = 0;
  const marcas = /<\/?([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  marcas.lastIndex = i;
  let m;
  while ((m = marcas.exec(xml))) {
    const nome = m[1], fecha = m[0].startsWith('</'), vazio = m[3] === '/';
    if (fecha) { prof--; if (prof === 0) break; continue; }
    if (prof === 1) saida.push(nome);
    if (!vazio) prof++;
  }
  return saida;
}

/* ---------- sequencias oficiais do ECMA-376 ---------- */
const SEQUENCIAS = {
  worksheet: ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr',
    'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate',
    'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks',
    'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties',
    'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'drawingHF', 'picture', 'oleObjects', 'controls',
    'webPublishItems', 'tableParts', 'extLst'],
  workbook: ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets',
    'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews',
    'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'],
  styleSheet: ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs',
    'tableStyles', 'colors', 'extLst'],
  sheetView: ['pane', 'selection', 'pivotSelection', 'extLst'],
  font: ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color',
    'name', 'family', 'charset', 'scheme'],
  c: ['f', 'v', 'is', 'extLst'],
};

function conferirOrdem(xml, raiz, rotulo) {
  const seq = SEQUENCIAS[raiz];
  const achados = filhos(xml, raiz);
  if (!achados) { ok(false, rotulo + ': nao achei o elemento <' + raiz + '>'); return; }
  const desconhecidos = achados.filter(n => !seq.includes(n));
  if (desconhecidos.length) { ok(false, rotulo + ': elemento fora do esquema', desconhecidos.join(', ')); return; }
  let ultimo = -1, erro = null;
  for (const n of achados) {
    const i = seq.indexOf(n);
    if (i < ultimo) { erro = n; break; }
    ultimo = i;
  }
  ok(!erro, rotulo + ': ordem dos elementos de <' + raiz + '>',
     erro ? '<' + erro + '> veio tarde demais. Ordem encontrada: ' + achados.join(' ') : achados.join(' '));
}

/* ---------- roda ---------- */
console.log('Conferindo ' + path.basename(alvo) + '\n');
const buf = fs.readFileSync(alvo);
let partes;
try { partes = abrirZip(buf); ok(true, 'O ZIP abriu pelo diretorio central', partes.size + ' partes'); }
catch (e) { ok(false, 'O ZIP nao abriu: ' + e.message); console.log('\nFALHOU'); process.exit(1); }

let crcRuim = 0;
for (const [nome, p] of partes) if (crc32(p.dados) !== p.crcEsperado) { crcRuim++; console.log('         CRC errado em ' + nome); }
ok(crcRuim === 0, 'CRC de todas as partes confere');

for (const obrigatoria of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'])
  ok(partes.has(obrigatoria), 'Parte obrigatoria presente: ' + obrigatoria);

// XML bem formado: toda abertura tem fechamento na ordem certa.
for (const [nome, p] of partes) {
  const pilha = [];
  let quebrou = null;
  const marcas = /<\/?([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = marcas.exec(p.texto))) {
    if (m[0].startsWith('<?') || m[0].startsWith('<!')) continue;
    if (m[0].startsWith('</')) { if (pilha.pop() !== m[1]) { quebrou = m[1]; break; } }
    else if (m[3] !== '/') pilha.push(m[1]);
  }
  if (quebrou || pilha.length) { ok(false, 'XML mal formado em ' + nome, quebrou || 'sobrou ' + pilha.join(',')); }
}
ok(!falhas.some(f => f.startsWith('XML mal formado')), 'Todo XML bem formado');

console.log('');
conferirOrdem(partes.get('xl/workbook.xml').texto, 'workbook', 'workbook.xml');
if (partes.has('xl/styles.xml')) conferirOrdem(partes.get('xl/styles.xml').texto, 'styleSheet', 'styles.xml');

const abas = [...partes.keys()].filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
for (const nome of abas) {
  const xml = partes.get(nome).texto;
  conferirOrdem(xml, 'worksheet', path.basename(nome));
  const sv = xml.match(/<sheetView[\s>][\s\S]*?<\/sheetView>/);
  if (sv) conferirOrdem(sv[0], 'sheetView', path.basename(nome) + ' sheetView');
}

console.log('');
/* contagens declaradas batem com a realidade */
if (partes.has('xl/styles.xml')) {
  const est = partes.get('xl/styles.xml').texto;
  for (const [tag, filho] of [['fonts', 'font'], ['fills', 'fill'], ['borders', 'border'], ['cellXfs', 'xf'], ['numFmts', 'numFmt'], ['cellStyleXfs', 'xf']]) {
    const bloco = est.match(new RegExp('<' + tag + '[^>]*count="(\\d+)"[^>]*>([\\s\\S]*?)</' + tag + '>'));
    if (!bloco) continue;
    const reais = (bloco[2].match(new RegExp('<' + filho + '[\\s/>]', 'g')) || []).length;
    ok(Number(bloco[1]) === reais, 'styles.xml: count de <' + tag + '> bate', bloco[1] + ' declarado, ' + reais + ' presentes');
  }
  const nEstilos = Number((est.match(/<cellXfs[^>]*count="(\d+)"/) || [])[1] || 0);
  let maior = -1;
  for (const nome of abas)
    for (const m of partes.get(nome).texto.matchAll(/<c [^>]*s="(\d+)"/g)) maior = Math.max(maior, Number(m[1]));
  ok(maior < nEstilos, 'Todo estilo usado nas celulas existe no styles.xml',
     'maior indice usado ' + maior + ', existem ' + nEstilos);
}

/* cada aba declarada no workbook tem relacionamento e arquivo */
const rels = partes.get('xl/_rels/workbook.xml.rels').texto;
const wb = partes.get('xl/workbook.xml').texto;
const declaradas = [...wb.matchAll(/<sheet [^>]*r:id="([^"]+)"/g)].map(m => m[1]);
let faltando = [];
for (const rid of declaradas) {
  const r = rels.match(new RegExp('<Relationship Id="' + rid + '"[^>]*Target="([^"]+)"'));
  if (!r) { faltando.push(rid + ' sem relacionamento'); continue; }
  const destino = 'xl/' + r[1].replace(/^\//, '');
  if (!partes.has(destino)) faltando.push(destino + ' nao existe no pacote');
}
ok(faltando.length === 0, declaradas.length + ' abas apontam para arquivo existente', faltando.join('; '));

/* o Content_Types cobre tudo */
const ct = partes.get('[Content_Types].xml').texto;
const semTipo = [...partes.keys()].filter(n => n.endsWith('.xml') && n !== '[Content_Types].xml')
  .filter(n => !ct.includes('/' + n) && !/Default Extension="xml"/.test(ct));
ok(semTipo.length === 0, 'Content_Types cobre todas as partes XML', semTipo.join(', '));

/* linhas e celulas em ordem crescente, que o Excel tambem cobra */
for (const nome of abas) {
  const xml = partes.get(nome).texto;
  const linhas = [...xml.matchAll(/<row r="(\d+)"/g)].map(m => Number(m[1]));
  const crescente = linhas.every((v, i) => i === 0 || v > linhas[i - 1]);
  ok(crescente, path.basename(nome) + ': linhas em ordem crescente', linhas.length + ' linhas');
  let celulasOk = true;
  for (const bloco of xml.matchAll(/<row [^>]*>([\s\S]*?)<\/row>/g)) {
    const cols = [...bloco[1].matchAll(/<c r="([A-Z]+)\d+"/g)]
      .map(m => [...m[1]].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0));
    if (!cols.every((v, i) => i === 0 || v > cols[i - 1])) celulasOk = false;
  }
  ok(celulasOk, path.basename(nome) + ': celulas em ordem dentro de cada linha');
}

console.log('\n' + (falhas.length
  ? falhas.length + ' PROBLEMA' + (falhas.length > 1 ? 'S' : '') + '. O Excel provavelmente vai recusar.'
  : 'ESTRUTURA VALIDA PARA O EXCEL.'));
process.exitCode = falhas.length ? 1 : 0;
