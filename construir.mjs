/* ============================================================
   Le o dados.json, cifra o conteudo com a senha e grava o bloco
   COFRE dentro do index.html.

   Uso:  node construir.mjs
         node construir.mjs dados.exemplo.json     (para testar)

   O index.html resultante pode ir para um repositorio publico:
   nenhum valor aparece em texto aberto dentro dele.
   ============================================================ */
import fs from 'node:fs';
import path from 'node:path';
import { webcrypto as crypto } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const ITERACOES = 250000;

const arquivoDados = process.argv[2] || 'dados.json';
const caminhoDados = path.join(AQUI, arquivoDados);
const caminhoHtml  = path.join(AQUI, 'index.html');

if (!fs.existsSync(caminhoDados)) {
  console.error('Nao achei ' + arquivoDados + '. Se voce acabou de clonar o repositorio,');
  console.error('o arquivo com os valores reais nao vem junto, de proposito.');
  console.error('Para testar: node construir.mjs dados.exemplo.json');
  process.exit(1);
}

const bruto = JSON.parse(fs.readFileSync(caminhoDados, 'utf8'));

/* ---- conferencias antes de cifrar, para nao publicar numero quebrado ---- */
const problemas = [];
if (!bruto.senha || bruto.senha.length < 8) problemas.push('A senha precisa de pelo menos 8 caracteres.');
if (!Array.isArray(bruto.fornecedores) || !bruto.fornecedores.length) problemas.push('Nenhum fornecedor.');
if (!bruto.atualizado) problemas.push('Falta a data em "atualizado".');

let totC = 0, totP = 0, totAberto = 0;
for (const f of bruto.fornecedores || []) {
  if (typeof f.c !== 'number' || typeof f.p !== 'number') { problemas.push(f.n + ': contratado ou pago nao e numero.'); continue; }
  if (f.p > f.c + 0.005) problemas.push(f.n + ': pago (' + f.p + ') maior que o contratado (' + f.c + ').');
  const aberto = (f.par || []).reduce((s, p) => s + p.v, 0);
  const falta = f.c - f.p;
  if (Math.abs(aberto - falta) >= 0.005)
    problemas.push(f.n + ': as parcelas somam ' + aberto + ' mas falta pagar ' + falta + '.');
  totC += f.c; totP += f.p; totAberto += aberto;
}
if (Math.abs(totP + totAberto - totC) >= 0.005)
  problemas.push('O total nao fecha: pago ' + totP + ' mais aberto ' + totAberto + ' difere de ' + totC + '.');

if (problemas.length) {
  console.error('\nNao gerei nada. Corrija o ' + arquivoDados + ':');
  problemas.forEach(p => console.error('  - ' + p));
  process.exit(1);
}

/* ---- cifrar ---- */
const claro = JSON.stringify({
  atualizado: bruto.atualizado,
  fornecedores: bruto.fornecedores,
  tarefas: bruto.tarefas || [],
});

const sal = crypto.getRandomValues(new Uint8Array(16));
const iv  = crypto.getRandomValues(new Uint8Array(12));
// Normalizar: o teclado do celular poe maiuscula na primeira letra sozinho.
// A pagina normaliza igual, senao a senha certa seria recusada.
const senhaNormal = bruto.senha.trim().toLowerCase();
const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(senhaNormal), 'PBKDF2', false, ['deriveKey']);
const chave = await crypto.subtle.deriveKey(
  { name: 'PBKDF2', salt: sal, iterations: ITERACOES, hash: 'SHA-256' },
  base, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
const cifrado = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, chave, new TextEncoder().encode(claro)));

const b64 = u => Buffer.from(u).toString('base64');
const cofre = JSON.stringify({ sal: b64(sal), iv: b64(iv), dados: b64(cifrado), iteracoes: ITERACOES });

/* ---- gravar no html ---- */
const html = fs.readFileSync(caminhoHtml, 'utf8');
const linha = /^const COFRE = .*; \/\* COFRE \*\/$/m;
if (!linha.test(html)) {
  console.error('Nao achei a linha do COFRE no index.html. Ela precisa terminar com  /* COFRE */');
  process.exit(1);
}
fs.writeFileSync(caminhoHtml, html.replace(linha, 'const COFRE = ' + cofre + '; /* COFRE */'), 'utf8');

/* ---- conferir que nada real sobrou em texto aberto ---- */
const final = fs.readFileSync(caminhoHtml, 'utf8');
const vazou = [];
for (const f of bruto.fornecedores) if (final.includes(f.n)) vazou.push('nome "' + f.n + '"');
for (const t of bruto.tarefas || []) if (final.includes(t.t)) vazou.push('tarefa "' + t.t.slice(0, 30) + '"');
if (final.includes(bruto.senha)) vazou.push('a senha');
// Canarios: so valores que nao poderiam estar no CSS por acaso.
// Numeros redondos (200, 300, 2000) e pequenos aparecem como medida, fonte e milissegundo,
// entao dariam falso positivo. Multiplo de 100 fora, abaixo de 900 fora.
const valores = new Set(
  bruto.fornecedores
    .flatMap(f => [f.c, f.p, ...(f.par || []).map(p => p.v)])
    .filter(v => v >= 900 && v % 100 !== 0));
for (const v of valores) if (new RegExp('(^|[^\\d.])' + v + '($|[^\\d])').test(final)) vazou.push('o valor ' + v);
if (!valores.size) console.warn('Aviso: nenhum valor distintivo para usar como canario.');

if (vazou.length) {
  console.error('\nPARE. Sobrou dado em texto aberto no index.html: ' + [...new Set(vazou)].join(', '));
  process.exit(1);
}

const brl = n => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
console.log('\nindex.html gerado a partir de ' + arquivoDados + '.');
console.log('  Fornecedores: ' + bruto.fornecedores.length + '   Tarefas: ' + (bruto.tarefas || []).length);
console.log('  Contratado ' + brl(totC) + '  |  Pago ' + brl(totP) + '  |  A pagar ' + brl(totAberto));
console.log('  Cofre: ' + cifrado.length + ' bytes cifrados, PBKDF2 de ' + ITERACOES.toLocaleString('pt-BR') + ' voltas.');
console.log('  Nenhum nome, valor ou senha aparece em texto aberto no arquivo.');
