/* ============================================================
   Le a planilha base editada no Excel e devolve para o dados.json.

       node ler-planilha.mjs            -> so mostra o que mudou
       node ler-planilha.mjs --gravar   -> mostra e grava o dados.json

   Sem --gravar nao escreve nada. Esse e o ponto: primeiro se olha a
   diferenca, depois se grava.

   As colunas sao achadas pelo NOME do cabecalho, nao pela posicao.
   Assim dá para mover coluna, mudar largura, cor e fonte sem quebrar
   a leitura. O que nao pode e renomear cabecalho.

   Antes de gravar, confere a regra de ouro: para cada fornecedor, a
   soma das parcelas tem de ser igual a contratado menos pago. Se
   algum nao fechar, nao grava nada e diz qual.
   ============================================================ */
import fs from 'node:fs';
import zlib from 'node:zlib';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const GRAVAR = process.argv.includes('--gravar');
const XLSX = path.join(AQUI, process.argv.find(a => a.endsWith('.xlsx')) || 'base-de-dados.xlsx');
const DADOS = path.join(AQUI, 'dados.json');

for (const [c, m] of [[XLSX, 'a planilha'], [DADOS, 'o dados.json']])
  if (!fs.existsSync(c)) { console.error('Nao achei ' + m + ': ' + c); process.exit(1); }

/* ---------- ZIP ---------- */
function abrirZip(buf) {
  let fim = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 65557; i--)
    if (buf.readUInt32LE(i) === 0x06054b50) { fim = i; break; }
  if (fim < 0) throw new Error('nao achei o fim do diretorio central');
  const total = buf.readUInt16LE(fim + 10);
  let p = buf.readUInt32LE(fim + 16);
  const partes = new Map();
  for (let n = 0; n < total; n++) {
    const metodo = buf.readUInt16LE(p + 10);
    const comp = buf.readUInt32LE(p + 20);
    const nomeLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const comentLen = buf.readUInt16LE(p + 32);
    const desloc = buf.readUInt32LE(p + 42);
    const nome = buf.toString('utf8', p + 46, p + 46 + nomeLen);
    const ini = desloc + 30 + buf.readUInt16LE(desloc + 26) + buf.readUInt16LE(desloc + 28);
    const cru = buf.subarray(ini, ini + comp);
    partes.set(nome, (metodo === 0 ? cru : zlib.inflateRawSync(cru)).toString('utf8'));
    p += 46 + nomeLen + extraLen + comentLen;
  }
  return partes;
}

const desescapar = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
  .replace(/&amp;/g, '&');

/* ---------- planilha para grade ---------- */
const partes = abrirZip(fs.readFileSync(XLSX));

// O Excel troca texto embutido por um indice no sharedStrings quando salva.
const compartilhadas = [];
if (partes.has('xl/sharedStrings.xml'))
  for (const si of partes.get('xl/sharedStrings.xml').matchAll(/<si>([\s\S]*?)<\/si>/g))
    compartilhadas.push([...si[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => desescapar(t[1])).join(''));

const nomesAbas = [...partes.get('xl/workbook.xml').matchAll(/<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]+)"/g)]
  .map(m => ({ nome: desescapar(m[1]), rid: m[2] }));
const relsWb = partes.get('xl/_rels/workbook.xml.rels');

function grade(nomeAba) {
  const aba = nomesAbas.find(a => a.nome === nomeAba);
  if (!aba) throw new Error('a planilha nao tem a aba "' + nomeAba + '"');
  const alvo = relsWb.match(new RegExp('<Relationship Id="' + aba.rid + '"[^>]*Target="([^"]+)"'))[1];
  const xml = partes.get('xl/' + alvo.replace(/^\//, ''));
  const linhas = new Map();
  for (const mr of xml.matchAll(/<row [^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const r = Number(mr[1]);
    const celulas = new Map();
    for (const mc of mr[2].matchAll(/<c ([^>]*?)\/?>(?:([\s\S]*?)<\/c>)?/g)) {
      const attrs = mc[1], corpo = mc[2] || '';
      const col = (attrs.match(/r="([A-Z]+)\d+"/) || [])[1];
      if (!col) continue;
      const tipo = (attrs.match(/t="([^"]+)"/) || [])[1] || 'n';
      let v;
      if (tipo === 'inlineStr') {
        v = [...corpo.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => desescapar(t[1])).join('');
      } else {
        // Celula com formula traz o resultado em <v>. E esse que interessa.
        const mv = corpo.match(/<v>([\s\S]*?)<\/v>/);
        if (!mv) continue;
        const bruto = desescapar(mv[1]);
        v = tipo === 's' ? compartilhadas[Number(bruto)] : (tipo === 'str' || tipo === 'e') ? bruto : Number(bruto);
      }
      if (v !== undefined && v !== '') celulas.set(col, v);
    }
    if (celulas.size) linhas.set(r, celulas);
  }
  return linhas;
}

/* ---------- acha o cabecalho pelo nome ---------- */
function tabela(nomeAba, rotulos) {
  const linhas = grade(nomeAba);
  let cab = null, linhaCab = null;
  for (let r = 1; r <= 12 && !cab; r++) {
    const l = linhas.get(r);
    if (!l) continue;
    const mapa = {};
    for (const [col, v] of l) {
      const achado = rotulos.find(x => String(v).trim().toLowerCase() === x.toLowerCase());
      if (achado) mapa[achado] = col;
    }
    if (rotulos.every(x => mapa[x])) { cab = mapa; linhaCab = r; }
  }
  if (!cab) throw new Error('na aba "' + nomeAba + '" nao achei o cabecalho com as colunas: ' + rotulos.join(', '));

  const saida = [];
  const ultima = Math.max(...linhas.keys());
  for (let r = linhaCab + 1; r <= ultima; r++) {
    const l = linhas.get(r);
    if (!l) continue;
    const primeira = l.get(cab[rotulos[0]]);
    if (primeira === undefined) continue;
    if (String(primeira).trim().toLowerCase() === 'total') continue;
    const reg = { _linha: r };
    for (const x of rotulos) reg[x] = l.get(cab[x]);
    saida.push(reg);
  }
  return saida;
}

// O Excel guarda data como dias desde 30/12/1899.
const paraISO = v => {
  if (v === undefined || v === null || v === '') return null;
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000).toISOString().slice(0, 10);
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  if (/^\d{4}-\d{2}-\d{2}/.test(String(v).trim())) return String(v).trim().slice(0, 10);
  return String(v).trim();
};
const numero = v => (v === undefined || v === null || v === '') ? null : Number(v);

/* ---------- le as tres abas ---------- */
let fornecedores, parcelas, tarefas;
try {
  fornecedores = tabela('Fornecedores', ['Fornecedor', 'Contratado', 'Pago'])
    .map(x => ({ linha: x._linha, n: String(x.Fornecedor).trim(), c: numero(x.Contratado), p: numero(x.Pago) }));
  parcelas = tabela('Parcelas', ['Fornecedor', 'Valor', 'Vencimento', 'Forma'])
    .filter(x => x.Valor !== undefined)
    .map(x => ({
      linha: x._linha, forn: String(x.Fornecedor).trim(), v: numero(x.Valor), d: paraISO(x.Vencimento),
      cartao: String(x.Forma || '').trim().toLowerCase().startsWith('cart'),
    }));
  tarefas = tabela('Tarefas', ['Tarefa', 'Prazo', 'Feito'])
    .map(x => ({ linha: x._linha, t: String(x.Tarefa).trim(), q: paraISO(x.Prazo), feito: String(x.Feito || '').trim() }));
} catch (e) {
  console.error('Nao consegui ler a planilha: ' + e.message);
  process.exit(1);
}

/* ---------- confere ---------- */
const problemas = [];
const vistos = new Set();
for (const f of fornecedores) {
  if (!f.n) problemas.push('linha ' + f.linha + ' de Fornecedores sem nome');
  if (vistos.has(f.n)) problemas.push('fornecedor repetido: "' + f.n + '"');
  vistos.add(f.n);
  if (!Number.isFinite(f.c) || !Number.isFinite(f.p))
    problemas.push('"' + f.n + '": contratado ou pago nao e numero');
  else if (f.p > f.c)
    problemas.push('"' + f.n + '": pago ' + f.p + ' maior que o contratado ' + f.c);
}
for (const p of parcelas) {
  if (!vistos.has(p.forn))
    problemas.push('linha ' + p.linha + ' de Parcelas: "' + p.forn + '" nao existe na aba Fornecedores');
  if (!Number.isFinite(p.v) || p.v <= 0)
    problemas.push('linha ' + p.linha + ' de Parcelas: valor invalido');
}
for (const f of fornecedores) {
  if (!Number.isFinite(f.c) || !Number.isFinite(f.p)) continue;
  const soma = parcelas.filter(p => p.forn === f.n).reduce((a, p) => a + p.v, 0);
  const falta = f.c - f.p;
  if (soma !== falta)
    problemas.push('"' + f.n + '": falta pagar ' + falta + ', mas as parcelas somam ' + soma +
                   '. Diferenca de ' + (falta - soma) + '.');
}

/* ---------- compara com o dados.json ---------- */
const velho = JSON.parse(fs.readFileSync(DADOS, 'utf8'));
const brl = v => Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const porNome = Object.fromEntries(velho.fornecedores.map(f => [f.n, f]));
const mudou = [];

for (const f of fornecedores) {
  const v = porNome[f.n];
  if (!v) { mudou.push('  + Novo fornecedor ' + f.n + ': contratado ' + brl(f.c) + ', pago ' + brl(f.p)); continue; }
  if (v.c !== f.c) mudou.push('  ~ ' + f.n + ': contratado ' + brl(v.c) + ' para ' + brl(f.c));
  if (v.p !== f.p) mudou.push('  ~ ' + f.n + ': pago ' + brl(v.p) + ' para ' + brl(f.p));
}
for (const v of velho.fornecedores)
  if (!fornecedores.some(f => f.n === v.n)) mudou.push('  - Sumiu o fornecedor ' + v.n);

const velhasPar = velho.fornecedores.flatMap(f => (f.par || []).map(p => ({ forn: f.n, ...p })));
const sobra = [...velhasPar];
for (const p of parcelas) {
  const i = sobra.findIndex(x => x.forn === p.forn && x.v === p.v);
  if (i < 0) { mudou.push('  + Nova parcela de ' + p.forn + ': ' + brl(p.v) + ' em ' + (p.d || 'sem data')); continue; }
  const x = sobra.splice(i, 1)[0];
  if ((x.d || null) !== p.d) mudou.push('  ~ Parcela de ' + p.forn + ' ' + brl(p.v) + ': vencimento ' + (x.d || 'sem data') + ' para ' + (p.d || 'sem data'));
  if (!!x.cartao !== p.cartao) mudou.push('  ~ Parcela de ' + p.forn + ' ' + brl(p.v) + ': forma mudou');
}
for (const x of sobra) mudou.push('  - Parcela de ' + x.forn + ' ' + brl(x.v) + ' saiu da lista, foi paga');

const velhasTar = Object.fromEntries(velho.tarefas.map(t => [t.t, t]));
const feitas = tarefas.filter(t => t.feito);
for (const t of tarefas) {
  const v = velhasTar[t.t];
  if (!v) { mudou.push('  + Nova tarefa: ' + t.t); continue; }
  if ((v.q || null) !== t.q) mudou.push('  ~ Tarefa "' + t.t.slice(0, 40) + '": prazo ' + (v.q || 'sem prazo') + ' para ' + (t.q || 'sem prazo'));
}
for (const v of velho.tarefas)
  if (!tarefas.some(t => t.t === v.t)) mudou.push('  - Sumiu a tarefa: ' + v.t.slice(0, 50));
for (const t of feitas) mudou.push('  v Tarefa marcada como feita, vai sair da lista: ' + t.t.slice(0, 50));

/* ---------- mostra ---------- */
const tc = fornecedores.reduce((a, f) => a + (f.c || 0), 0);
const tp = fornecedores.reduce((a, f) => a + (f.p || 0), 0);
const vc = velho.fornecedores.reduce((a, f) => a + f.c, 0);
const vp = velho.fornecedores.reduce((a, f) => a + f.p, 0);

console.log('Lido de ' + path.basename(XLSX) + ': ' + fornecedores.length + ' fornecedores, ' +
            parcelas.length + ' parcelas, ' + tarefas.length + ' tarefas.\n');
console.log('O QUE MUDOU');
console.log(mudou.length ? mudou.join('\n') : '  Nada. A planilha esta igual ao dados.json.');
console.log('\nTOTAIS');
console.log('  Antes:  contratado ' + brl(vc) + ' | pago ' + brl(vp) + ' | a pagar ' + brl(vc - vp));
console.log('  Agora:  contratado ' + brl(tc) + ' | pago ' + brl(tp) + ' | a pagar ' + brl(tc - tp));

if (problemas.length) {
  console.error('\nNAO GRAVEI. ' + problemas.length + (problemas.length === 1 ? ' problema:' : ' problemas:'));
  problemas.forEach(p => console.error('  - ' + p));
  console.error('\nArrume a planilha e rode de novo.');
  process.exit(1);
}
console.log('\n  [ok]  Regra de ouro: fecha em todos os ' + fornecedores.length + ' fornecedores.');

if (!GRAVAR) {
  console.log('\nNada foi gravado. Para aplicar: node ler-planilha.mjs --gravar');
  process.exit(0);
}

/* ---------- grava ---------- */
const hoje = new Date().toLocaleDateString('pt-BR');
const novo = {
  atualizado: hoje,
  senha: velho.senha,
  fornecedores: fornecedores.map(f => ({
    n: f.n, c: f.c, p: f.p,
    par: parcelas.filter(p => p.forn === f.n)
      .sort((a, b) => (a.d || '9999').localeCompare(b.d || '9999') || a.v - b.v)
      .map(p => p.cartao ? { v: p.v, d: p.d, cartao: true } : { v: p.v, d: p.d }),
  })),
  // Tarefa marcada como feita sai da lista. E isso que marcar significa.
  tarefas: tarefas.filter(t => !t.feito).map(t => t.q ? { t: t.t, q: t.q } : { t: t.t }),
};

fs.writeFileSync(DADOS, JSON.stringify(novo, null, 2) + '\n', 'utf8');
console.log('\nGravado o dados.json. Atualizado em ' + hoje + '.');
console.log('Agora rode: node construir.mjs');
