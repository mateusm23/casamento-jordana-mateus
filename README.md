# Jordana e Mateus

Página de uma tela só para um casamento, com contagem regressiva aberta e um controle de pagamentos protegido por senha.

Não usa framework, não usa biblioteca e não tem servidor. É um arquivo HTML.

## Como os valores ficam protegidos num repositório público

O repositório é público, porque o GitHub Pages em conta gratuita exige isso. Qualquer pessoa pode baixar o `index.html` e ler o código.

Por isso os valores não estão aqui em texto aberto. Eles ficam no objeto `COFRE`, cifrado em **AES-GCM de 256 bits**, com a chave derivada da senha por **PBKDF2-SHA256 de 250 mil iterações** e sal aleatório de 16 bytes. A página só decifra no navegador, depois que alguém digita a senha certa. O próprio AES-GCM verifica a senha: com a chave errada a autenticação falha e nada é revelado.

A senha não está escrita em lugar nenhum do repositório. Quem gera o cofre é o `construir.mjs`, a partir de um arquivo de dados que fica fora do git.

## Os arquivos

| Arquivo | O que faz |
|---|---|
| `index.html` | A página inteira, com o cofre cifrado dentro |
| `construir.mjs` | Lê os dados, confere as contas, cifra e grava o cofre no HTML |
| `dados.exemplo.json` | Dados fictícios, para rodar o build sem ter os reais |
| `verifica-publicacao.mjs` | Barreira que bloqueia commit com dado sensível em texto aberto |

## Rodando

```
node construir.mjs dados.exemplo.json
node verifica-publicacao.mjs
```

Depois é só abrir o `index.html` no navegador. A senha do exemplo está no `dados.exemplo.json`.

## O que a página tem

- Capa em tela cheia, pensada primeiro para o celular, com a coroa de louro se desenhando, os nomes escritos da esquerda para a direita por uma máscara animada e um relógio regressivo grande.
- Controle de pagamentos atrás da senha: contratado, pago e a pagar, calendário por data de vencimento, tabela de fornecedores, lista de tarefas e um bloco de conferência que refaz as somas a cada abertura.
- Exportação para `.xlsx`, montada no próprio navegador sem biblioteca nenhuma. O total vai como fórmula `SUM`, não como número gravado, para a planilha nunca mostrar um valor velho.

## Licença

Uso pessoal.
