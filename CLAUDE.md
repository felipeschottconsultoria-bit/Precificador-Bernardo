# Precificador Reforma

Ferramenta interna, separada dos outros apps (CBS SIMPLES, Simples Lite, Diniz,
Pactus, Planalto). Feita para um cliente de comércio (marca "Vest", produtos de
vestuário/lingerie vendidos em kits) que quer entender: quanto vai pagar de
Simples hoje x na Reforma Tributária (CBS/IBS), e quanto precisa reajustar
preço pra manter a margem — de forma manual (sem importar XML) e visual.

Stack: Vite + React 18 + Tailwind, `xlsx` (SheetJS) pra importar planilha de
produtos, `lucide-react` pros ícones, `firebase` (Firestore) pra persistir os
produtos num banco de verdade (com fallback pro localStorage do navegador).

## Como rodar

```
npm install
npm run dev
```

O Vite tenta a porta 5183 mas quase sempre sobe em outra (tem várias portas
antigas ocupadas nesta máquina — 5200+ — de sessões anteriores do Claude Code
que não foram encerradas limpo). Confira o output do terminal pra pegar a
porta certa.

## As 5 abas

1. **Simples Nacional** — RBT12 + faturamento do mês → identifica a faixa do
   Anexo I (comércio) e a alíquota efetiva do DAS de **hoje (2026)**.
   Faixas/nominal/dedução são as mesmas de 2027+ (a reforma não muda isso, só
   a composição interna do DAS) — por isso o motor reaproveita a tabela de
   2027 pra calcular "hoje" também (ver `simplesEngine.js`).

2. **CBS e IBS** — simula a reforma a partir de 2027. Regime **Por Dentro**
   (DAS reconstituído, CBS/IBS embutidos) x **Por Fora** (regime regular:
   DAS residual + CBS + IBS separados). Alíquota de CBS é **editável**
   (some com o nominal da LC 214/2025 quando muda o ano). Tem bloco de
   **créditos manuais** (CBS/IBS pago nas compras, abate do devido no mês).
   Mostra o detalhamento por tributo (IRPJ/CSLL/CPP/ICMS no residual,
   CBS/IBS por fora) e os subtotais de alíquota efetiva de cada grupo.

3. **Produtos** — cadastro dos produtos em lista clicável (linha por linha).
   Clicar abre um modal onde dá pra editar o preço de compra base e **adicionar quantos
   custos extras quiser** (frete, embalagem, taxa de marketplace...) — o
   custo total (base + extras) é o que entra em todos os cálculos de margem.
   Tem "Importar planilha" (reconhece colunas por nome, tolera cabeçalhos
   fora de ordem) e "Adicionar produto manualmente".

4. **Precificação** — tabela comparando, produto a produto: margem líquida
   hoje, preço sugerido e aumento necessário pra manter a margem em **Por
   Dentro** e em **Por Fora**, no ano escolhido.

5. **Repasse** — pra quando o RBT12 real já é maior que o que os preços
   atuais consideram (a empresa cresceu de faixa e o preço não acompanhou
   ainda). Você informa a alíquota que os preços de hoje embutem e em
   quantos meses quer fechar esse gap; mostra um gráfico (margem preservada
   com repasse x margem se o preço não mudar) e uma tabela por produto,
   navegável por mês (botões "Hoje", "Mês 1"...).

## Arquivos principais

- `src/App.jsx` — todos os componentes de aba (arquivo único, de propósito —
  é um app pequeno, sem necessidade de roteamento/múltiplos arquivos ainda).
- `src/simplesEngine.js` — `SIMPLES_DB` (tabela do Anexo I, 2027-2033),
  `REFORM_SCHEDULE` (nominal de CBS/IBS por ano), `calcDasHoje`,
  `calcDasPorDentroRate`, `calcDasPorForaRate` (todas com `detalhe` por
  tributo pra UI itemizada).
- `src/produtosStore.js` — toda a persistência de produtos. Usa Firestore
  quando configurado (`.env`), cai pro localStorage senão. Cada produto é um
  documento na coleção `produtos`, com `custosExtras` embutido como array no
  próprio documento (sem subcoleção). **Importante:** `carregarOuSemear`
  memoiza a carga inicial num promise de módulo — sem isso, o duplo-efeito
  do React StrictMode (dev) semeia a planilha-base duas vezes no banco (isso
  já aconteceu uma vez no Supabase, 624 produtos ao invés de 312 — corrigido,
  e a proteção continua valendo pro Firestore).
- `src/firebaseClient.js` — inicializa o app do Firebase e o Firestore só se
  `VITE_FIREBASE_API_KEY` e `VITE_FIREBASE_PROJECT_ID` estiverem no `.env`.
- `firestore.rules` — regra aberta pra coleção `produtos` (ferramenta interna
  sem login) — colar no console do Firebase (Firestore Database > Regras).
- `src/assets/produtos-base.xlsx` — planilha real do cliente, usada como
  semente inicial do banco/localStorage na primeiríssima carga.

## Firebase — status

Projeto criado no console (plano Spark/gratuito), Firestore em
`southamerica-east1`. Migrado do Supabase (usado antes) porque o usuário
preferiu Firebase. `.env` local (gitignored) precisa das 6 chaves do app Web
registrado no console — `.env.example` no repo mostra os nomes das
variáveis, sem valores.

## Estado do Git

Só existe **um commit** até agora (`f6e9972`, snapshot inicial). Desde então
rolou bastante trabalho ainda **não commitado**: a integração com Supabase e
depois a migração pra Firebase, o detalhamento por tributo na aba CBS e IBS,
a separação da aba Produtos (antes fazia parte de Precificação, virou lista
em vez de cards), a aba Repasse, e os ajustes finos de UI. Repositório não
tem remoto no GitHub ainda — só existe local
(`C:\Users\felip\Downloads\Analise-precifica-o--main\Analise-precifica-o--main`).

## Próximos passos em aberto

- Perguntado ao usuário se quer subir pro GitHub (como os outros 4 apps) —
  ainda sem resposta definitiva.
- Commit do trabalho recente (migração pra Firebase + ajustes de UI) ainda
  pendente — só fazer quando o usuário pedir explicitamente.
