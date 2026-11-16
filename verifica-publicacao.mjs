/* ============================================================
   Barreira de publicacao (item 12 do checklist do Mateus).

   Varre o que esta prestes a entrar no commit e BLOQUEIA se achar
   dado sensivel em texto aberto. Roda sozinho antes de cada commit,
   pelo hook do git, e pode rodar a mao:

       node verifica-publicacao.mjs

   Sai com codigo 1 quando acha problema, e ai o commit nao passa.

   Os termos proibidos NAO ficam escritos aqui. Este arquivo vai para
   um repositorio publico, entao uma lista com os nomes dos fornecedores
   dentro dele vazaria justamente o que ela deveria proteger. Em vez
   disso, a lista e montada na hora a partir do dados.json, que fica
   so no computador do Mateus.
   ============================================================ */
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const git = cmd => execSync('git ' + cmd, { cwd: AQUI, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });

/* ---- 1. Arquivos que nunca podem ser versionados ----
   Imagem entra aqui porque nao da para varrer imagem por texto, e o convite
   do casamento tem o CPF impresso nele. A pagina nao usa arquivo de imagem:
   todo o desenho e SVG dentro do proprio HTML. Se um dia precisar de uma,
   tire a linha das imagens aqui de proposito, depois de olhar a imagem. */
const PROIBIDOS = [
  /^dados\.json$/i,
  /\.xlsx?$/i, /\.xlsm$/i,
  /\.(jpe?g|png|gif|webp|bmp|tiff?|pdf)$/i,
  /^~\$/,
];

/* ---- 2. Padroes que nunca podem aparecer em texto aberto ---- */
const PADROES = [
  { nome: 'CPF',                 re: /\b\d{3}\.\d{3}\.\d{3}-\d{2}\b/ },
  { nome: 'CNPJ',                re: /\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/ },
  { nome: 'e-mail',              re: /\b[\w.+-]+@[\w-]+\.[\w.]{2,}\b/ },
  { nome: 'telefone brasileiro', re: /\b\(?\d{2}\)?\s?9\d{4}[- ]?\d{4}\b/ },
  { nome: 'chave de API',        re: /\b(gh[pousr]_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9]{20,}|AIza[A-Za-z0-9_-]{30,})\b/ },
];

/* ---- 3. Termos deste projeto, montados a partir do dados.json ---- */
const termos = new Set();
const caminhoDados = path.join(AQUI, 'dados.json');
let temDados = false;
if (fs.existsSync(caminhoDados)) {
  temDados = true;
  const d = JSON.parse(fs.readFileSync(caminhoDados, 'utf8'));
  if (d.senha) termos.add(d.senha);
  for (const f of d.fornecedores || []) {
    termos.add(f.n);
    // So valores distintivos: redondo e pequeno aparece no CSS por acaso.
    for (const v of [f.c, f.p, ...(f.par || []).map(p => p.v)])
      if (v >= 900 && v % 100 !== 0) termos.add(String(v));
  }
  for (const t of d.tarefas || []) {
    // Nomes proprios dentro das tarefas, que identificam gente de verdade.
    for (const palavra of String(t.t).match(/\b[A-ZÁÉÍÓÚÂÊÔÃÕÇ][a-záéíóúâêôãõç]{3,}\b/g) || []) termos.add(palavra);
  }
  const soma = (d.fornecedores || []).reduce((a, f) => ({ c: a.c + f.c, p: a.p + f.p }), { c: 0, p: 0 });
  for (const v of [soma.c, soma.p, soma.c - soma.p]) termos.add(String(v));

  // Tambem na forma escrita em portugues, com ponto de milhar.
  // Sem isto um texto com os totais formatados passaria batido.
  for (const t of [...termos]) {
    if (!/^\d{4,}$/.test(t)) continue;
    termos.add(Number(t).toLocaleString('pt-BR'));
  }
}
// Palavras comuns que o apanhador de nomes proprios pega sem querer.
for (const falso of ['Tarefa','Caixas','Playlist','Programa','Marcar','Fechar','Combinar','Prova','Autorizacao','Autorização'])
  termos.delete(falso);

/* ---- quais arquivos olhar ---- */
let alvos, ondeVeio;
try {
  alvos = git('diff --cached --name-only --diff-filter=ACM').split('\n').map(s => s.trim()).filter(Boolean);
  ondeVeio = 'no stage do commit';
  if (!alvos.length) {
    alvos = git('ls-files').split('\n').map(s => s.trim()).filter(Boolean);
    ondeVeio = 'ja versionados (nada no stage)';
  }
} catch (e) {
  // Fora de um repositorio nada esta sendo publicado, entao os arquivos
  // que o .gitignore cobre ficam de fora da varredura.
  alvos = fs.readdirSync(AQUI)
    .filter(f => fs.statSync(path.join(AQUI, f)).isFile())
    .filter(f => !PROIBIDOS.some(re => re.test(f)));
  ondeVeio = 'a pasta, fora de um repositorio git';
}

const falhas = [];
const BINARIO = /\.(png|jpe?g|gif|webp|ico|pdf|zip|xlsx?|xlsm|woff2?|ttf)$/i;

for (const rel of alvos) {
  if (PROIBIDOS.some(re => re.test(path.basename(rel)) || re.test(rel)))
    falhas.push(rel + ': este arquivo nao pode ser versionado.');

  const abs = path.join(AQUI, rel);
  if (!fs.existsSync(abs) || BINARIO.test(rel)) continue;

  const linhas = fs.readFileSync(abs, 'utf8').split('\n');
  for (const { nome, re } of PADROES)
    linhas.forEach((l, i) => { if (re.test(l)) falhas.push(rel + ':' + (i + 1) + ': ' + nome + ' em texto aberto.'); });
  for (const t of termos)
    linhas.forEach((l, i) => { if (l.includes(t)) falhas.push(rel + ':' + (i + 1) + ': "' + t + '" so pode existir cifrado.'); });
}

/* ---- o cofre precisa estar la, e fechado ---- */
const html = path.join(AQUI, 'index.html');
if (fs.existsSync(html)) {
  const t = fs.readFileSync(html, 'utf8');
  if (!/"dados":"[A-Za-z0-9+/=]{400,}"/.test(t))
    falhas.push('index.html: o COFRE esta vazio ou pequeno demais. Rode node construir.mjs antes de commitar.');
  if (/const FORNECEDORES = \[\s*\{/.test(t))
    falhas.push('index.html: existe uma lista de fornecedores em texto aberto no arquivo.');
}

if (falhas.length) {
  console.error('\nPUBLICACAO BLOQUEADA. ' + falhas.length + (falhas.length === 1 ? ' problema:' : ' problemas:'));
  [...new Set(falhas)].forEach(f => console.error('  - ' + f));
  console.error('\nCorrija antes de commitar. Para entender cada item, veja o CONTEXTO.md.');
  process.exit(1);
}

console.log('Barreira de publicacao: ' + alvos.length + ' arquivos conferidos (' + ondeVeio + ').');
console.log(temDados
  ? '  ' + termos.size + ' termos do dados.json conferidos, nenhum aparece em texto aberto.'
  : '  Aviso: dados.json nao existe aqui, entao conferi so os padroes gerais (CPF, e-mail, chaves).');
