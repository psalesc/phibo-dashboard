# Painel de Vendas Phibo — contexto para o Claude Code

Este repo existia antes só como uma sequência de rodadas manuais numa conversa
de Cowork/claude.ai: Pedro dizia "atualizar", e o Claude daquela sessão
baixava 5 planilhas do Google Drive, rodava um script Python, e republicava
um HTML num link de Artifact (`https://claude.ai/artifact/5N8gL4zd6BLnqeAWGMSTmL`,
hoje na Version 43). Essa pasta é a tentativa de tornar isso um pipeline de
verdade — arquivos num repo, histórico no git, rodando por aqui (Claude Code)
em vez de numa conversa.

**Antes de tocar em qualquer coisa, leia este arquivo inteiro.** Ele carrega
18 rodadas de bugs já encontrados e corrigidos — reverter uma dessas
correções sem saber por que ela existe é o erro mais provável de repetir.

## O que é isso, de verdade

Phibo é um ERP B2B SaaS para varejo de moda. Pedro é o Gerente Comercial.
O time de vendas novas é Samuel, Marina e Bruno (cada um com meta/comissão);
Maria é renovação/retenção, com modelo de meta diferente, tratada em aba
separada do painel. O painel mostra: KPIs de faturamento/meta/fechamentos,
tendência diária, receita diária, evolução acumulada vs. meta, taxa de
conversão (leads → vendas, ideal ≥25%), vendas diárias por vendedor, vendas
por plano, uma tabela de transações pesquisável, e uma aba de renovações.

## Arquitetura — como os dados fluem

```
Google Drive (5 planilhas, veja "Fontes de dados" abaixo)
        │
        ▼
  pipeline/fetch_sheets.py   (baixa as 5 como .xlsx — ver "Acesso aos dados")
        │
        ▼
  pipeline/build_dashboard.py   (lê os .xlsx, monta o JSON de dados)
        │  usa dashboard_template.html como esqueleto,
        │  injeta o JSON no marcador /*__DATA_JSON__*/
        ▼
  pipeline/dashboard.html   (saída final — é isso que vai pro ar)
```

`pipeline/historical/*.json` é um cache por mês fechado (gerado uma vez por
`build_historical.py`, direto das planilhas de vendedor) — guarda o seletor
de mês/histórico do painel. O mês corrente nunca vem desse cache; só os
meses já encerrados.

## Fontes de dados (Google Drive)

Todas de propriedade de `pedro@phibo.com.br`. Os IDs abaixo já estão
hardcoded em `pipeline/fetch_sheets.py`.

| Arquivo | Nome no Drive | ID |
|---|---|---|
| `master_latest.xlsx` | Dashboard de Vendas — Phibo (ao vivo) | `1YFK6s_fJw_VIVeQWJ-p9qtqtY39L2Qakx7jN6WvXCXg` |
| `controle_vendas.xlsx` | Controle Vendas | `15jgTNMrhClhy4h7c-rURpVg6IfJd4uPAU9Qe8_k2GoE` |
| `samuel.xlsx` | Samuel - Planilha | `1z78YkW9VIvSLEnJ2iTGI9CEXNeMYbXDHPnlxJk6MizA` |
| `marina.xlsx` | Marina - Planilha | `1Y_tPhMDUiYJD6zRlnZLIEvYvkVPNlCTyyaH2DMgH1H4` |
| `bruno.xlsx` | Bruno - Planilha | `1JYUSwj97Y2huGglji619t1UkoXaKCoOoUbsIzRYPTEQ` |

- **Master** ("Dashboard de Vendas — Phibo (ao vivo)") é alimentada por um
  Google Apps Script (`Dashboard_AoVivo.gs`, roda a cada 10 min, já instalado
  na planilha — o código-fonte dele NÃO está neste repo, só existe dentro do
  Apps Script do Google Sheets do Pedro). Ele lê Samuel/Marina/Bruno/Maria e
  escreve nas abas "Resumo Vendedores" e "Vendas Detalhadas" do master —
  `build_dashboard.py` lê essas duas abas. **"Vendas Detalhadas" só guarda o
  mês corrente** — é sobrescrita a cada virada de mês, por isso meses
  fechados precisam ser capturados antes (ver "Virada de mês" abaixo).
- **Controle Vendas** é preenchida manualmente pelo Pedro, uma aba por mês
  desde nov/2025. Ela é a **fonte de verdade para Faturamento e Meta** (ver
  "Regras de reconciliação" abaixo) e tem dados que o master não tem: vendas
  diárias por vendedor, leads/conversão, bloco de renovações da Maria.
- **Samuel/Marina/Bruno** (planilhas individuais) são usadas só pra
  reconstruir um mês fechado (`build_data_from_vendors()`/
  `build_historical.py`) ou, se precisar, pra montar o mês corrente quando o
  master ainda não virou (ver "Virada de mês").

## Acesso aos dados

**Estado atual (15/10/2026): Workload Identity Federation, sem chave.** A
organização Google da Phibo bloqueia a criação de chaves de service account
(política `iam.disableServiceAccountKeyCreation`, padrão de segurança do
Workspace), então a opção "chave JSON" abaixo não está disponível. O que foi
configurado no lugar:

- Projeto Google Cloud `phibo-dashboard-511718` (número `554455998023`),
  Google Drive API ativada.
- Service account `dashboard-reader@phibo-dashboard-511718.iam.gserviceaccount.com`,
  compartilhada como **Leitor** nas 5 planilhas.
- Pool `github` / provider `phibo-dashboard` (OIDC do GitHub), restrito ao
  repo `psalesc/phibo-dashboard`, com `roles/iam.workloadIdentityUser` na
  service account. Ou seja: **só o GitHub Actions desse repo** consegue
  baixar as planilhas. Se o repo mudar de nome/dono, é preciso recriar o
  binding e a condição do provider.
- `fetch_sheets.py` usa `service_account.json` se existir; senão,
  `google.auth.default()` (a credencial que o passo
  `google-github-actions/auth` cria no Actions).
- Fora do Actions (Mac do Pedro, sessões do Claude) não há credencial — e
  as sessões do Claude também não alcançam `googleapis.com` pela rede.
  Nesses casos, use a opção manual abaixo (links de export direto:
  `https://docs.google.com/spreadsheets/d/<ID>/export?format=xlsx`).

Duas opções originais (referência) — os scripts não se importam de onde o .xlsx veio, só esperam
o caminho do arquivo como argumento:

**Opção recomendada — `pipeline/fetch_sheets.py` (service account do
Google).** Configura uma vez (~15 min: criar projeto no Google Cloud, ativar
a API do Drive, criar uma service account, gerar uma chave JSON, compartilhar
as 5 planilhas com o e-mail dela) e depois disso o script baixa os 5
arquivos sozinho a cada rodada — sem exportar nada manualmente. O passo a
passo completo está no cabeçalho do próprio `fetch_sheets.py`. **A chave
JSON (`service_account.json`) NUNCA pode ir pro git** — já está no
`.gitignore`.

**Opção manual (fallback, ou pra começar rápido antes de configurar a
service account).** Baixar os 5 arquivos do Drive na mão (Arquivo > Fazer
download > Microsoft Excel) e salvar na pasta `pipeline/` com os mesmos
nomes que `fetch_sheets.py` geraria (`master_latest.xlsx`,
`controle_vendas.xlsx`, `samuel.xlsx`, `marina.xlsx`, `bruno.xlsx`).

## Rodando o pipeline

```bash
cd pipeline
pip install -r requirements.txt

# 1. buscar os dados (service account já configurada) ou baixar manualmente
python3 fetch_sheets.py

# 2. gerar o dashboard.html (mês corrente detectado automaticamente)
python3 build_dashboard.py master_latest.xlsx dashboard.html controle_vendas.xlsx

# (opcional, no início de cada mês — ver "Virada de mês" — forçar um mês
# específico em vez do detectado pela data real do sistema)
python3 build_dashboard.py master_latest.xlsx dashboard.html controle_vendas.xlsx Setembro
```

`dashboard.html` é o artefato final — é esse arquivo que vai pro GitHub
Pages (ver "Publicação").

## IMPORTANTE: primeira rodada neste repo — reconstruir o histórico

Este repo começa sem `pipeline/historical/*.json` (não foram incluídos no
pacote inicial — são fáceis de regerar e continham dados de cliente). Antes
da primeira build "de verdade", rode, com os 5 .xlsx já baixados:

```bash
python3 build_historical.py Agosto 2026 08 samuel.xlsx marina.xlsx bruno.xlsx controle_vendas.xlsx "Agosto/2026 (fechado)"
python3 build_historical.py Setembro 2026 09 samuel.xlsx marina.xlsx bruno.xlsx controle_vendas.xlsx "Setembro/2026 (fechado)"
```

**Valores esperados pra validar que deu certo** (números finais, já
reconciliados pela Controle Vendas — confira antes de seguir):
- **Agosto 2026:** faturamento R$ 537.617,54 · 620 fechamentos · meta
  R$ 512.000.
- **Setembro 2026:** faturamento R$ 533.168,43 · 637 fechamentos · meta
  R$ 482.000 (110,6% da meta).

**Resultado da regeneração em 15/10/2026:** Agosto bateu nos 3 números.
Setembro bateu faturamento e meta, mas deu **638** fechamentos (1 a mais) —
nenhuma duplicata; provável lançamento/edição na planilha de vendedor depois
de 05/10. Candidatos na aba de setembro do Bruno: "Vinicius" (Prata Anual,
R$ 1.428) **sem data** e "Thays" (Migração, R$ 50) com data "28/98". Pedro
optou por seguir com 638. Há também 4 vendas de setembro sem valor
(Marilene e Daniele/Samuel; Thais e Fábio Arruda/Marina) — entram na
contagem, não afetam o faturamento (que vem da Controle Vendas).

Se esses números não baterem, o problema mais provável é planilha de
vendedor desatualizada — re-baixe os 5 arquivos (ver "Lições de processo"
abaixo, rodada 14) antes de desconfiar do código.

## Regras de reconciliação de dados — NÃO REVERTER sem entender o porquê

Estas três correções existem porque os números "nativos" de cada planilha
de vendedor já se mostraram errados, mais de uma vez, de formas diferentes.
Todas estão implementadas em `build_dashboard.py` e se aplicam tanto ao mês
corrente quanto ao histórico:

1. **Faturamento e Meta vêm da Controle Vendas, não da planilha do
   vendedor** (`reconcile_kpis_with_cv()`). Achado 14/09: o painel mostrava
   R$478k de faturamento em agosto; o real (Controle Vendas) era R$537k —
   a planilha do Samuel e do Bruno tinham fórmulas de "Faturamento"
   desatualizadas/quebradas (a da Marina batia certo). Achado 25/09: Meta
   total mostrava R$480.000 em vez de R$482.000 — a planilha do Samuel
   tinha a meta antiga, a Controle Vendas já tinha a corrigida. A Controle
   Vendas é preenchida manualmente pelo Pedro e é o que ele realmente
   confere — por isso prevalece, sempre que disponível.
2. **Fechamentos vêm da contagem real de `records`, não de uma fórmula de
   intervalo fixo na planilha do vendedor** (`recompute_fechamentos_from_records()`).
   Achado 15/09: o KPI "Fechamentos" não batia com o número de linhas na
   aba Transações do mesmo painel (Samuel 130 vs 133 reais, Bruno com a
   fórmula contando 1 linha que não existe mais). Contar `records` de
   verdade elimina a dependência de qualquer fórmula interna que fica pra
   trás quando alguém lança venda depois do intervalo ter sido definido.
3. **Mensais somam ×4 sobre o valor da transação.** Não é bug — é lógica
   comercial da Phibo pra planos mensais, confirmada pelo Pedro desde a
   primeira versão. Por isso o "Faturamento Total" do painel nunca vai
   bater com a soma bruta da aba Transações (que mostra valor real por
   venda) — isso já está explicado no rodapé do próprio painel.

Se algum número parecer errado de novo, **o primeiro passo é sempre conferir
se as 5 planilhas foram baixadas de fato atualizadas** (checar
`modifiedTime` no Drive), não assumir que é bug de código — já aconteceu
mais de uma vez (ver "Lições de processo").

## Virada de mês — o ponto que mais quebra

O `Dashboard_AoVivo.gs` (na planilha master) e a aba nova em Controle
Vendas **não viram o mês automaticamente** — dependem do Pedro trocar um
CONFIG manualmente. Isso cria uma janela, todo início de mês, em que o
relógio do sistema já virou mas os dados de origem ainda não.

`build_dashboard.py` detecta o mês a processar pela **data real do
sistema** por padrão (não pelo conteúdo das planilhas) — então, nos
primeiros dias de qualquer mês novo, **confira antes de rodar sem
argumento**:

1. A aba "Vendas Detalhadas" do master já tem registros do mês novo?
2. Já existe a aba do mês novo na Controle Vendas, com dados?

Se qualquer uma das duas for "não", passe o **mês anterior** explicitamente
como 4º argumento de `build_dashboard.py`, em vez de deixar a detecção
automática usar a data do calendário — senão o build degrada silenciosamente
(a Controle Vendas é pulada com um aviso no log, perdendo as correções de
Faturamento/Meta acima) ou mistura dados de dois meses.

**Modo oficial desde 15/10/2026:** `build_dashboard.py` agora detecta
sozinho quando o master está num mês diferente do processado (mês mais
frequente nas datas de "Vendas Detalhadas") e monta o mês direto das
planilhas de vendedor (`build_data_from_vendors()`), desde que
`samuel/marina/bruno.xlsx` estejam na mesma pasta do master — o
`fetch_sheets.py` já baixa as 5. Sem elas, o build **aborta com exit 2** em
vez de publicar dados de um mês com rótulo de outro. Flag `--from-vendors`
força esse modo. Em 15/10 o master ainda estava preso em setembro (CONFIG
do `.gs` não virado) e o painel de outubro saiu por esse caminho:
R$ 186.769,42 · 221 vendas · meta R$ 588.000.

Histórico do problema — **se os dados de origem (vendedores + Controle Vendas) já viraram o mês mas
o master ainda não** (aconteceu em 05/10: master preso em setembro, mas as 3
planilhas de vendedor e a Controle Vendas já tinham outubro real) — isso é
um caso à parte, sem suporte no `main()` padrão. A solução usada foi um
script ad-hoc que chama `build_data_from_vendors()` (normalmente só usada
pra histórico) para o **mês corrente**, em vez de `build_data()`. Esse
script não ficou salvo como arquivo reutilizável na sessão original; se essa
situação se repetir, a forma mais limpa de resolver é adicionar essa opção
como um modo oficial em `build_dashboard.py` (ex. uma flag
`--from-vendors`), não reescrever o ad-hoc toda vez.

Os passos manuais de virada de mês, do lado do Pedro, continuam sendo: (a)
trocar `mm`/`yyyy`/`mesNomePT` no CONFIG do `Dashboard_AoVivo.gs`; (b) criar
a aba do mês novo em Controle Vendas, duplicando o layout da anterior; (c)
preencher a Meta de cada vendedor pra esse mês na Controle Vendas (bloco-
resumo, coluna "Meta") — sem isso o painel mostra Meta R$0/0% nesse mês
(comportamento esperado, não é bug).

## Publicação — GitHub Pages

Decisão desta migração: o painel passa a ser publicado em
`psalesc.github.io/phibo-dashboard` (repo `psalesc/phibo-dashboard`, que o
Pedro já tinha — era atualizado colando o HTML manualmente no GitHub). O
Claude Code (CLI) não tem acesso à ferramenta de Artifact do claude.ai —
só quem conversa no Cowork/claude.ai consegue publicar nesse link — então
continuar publicando ali exigiria voltar pra lá a cada atualização, o que
anula o ganho de automatizar isso aqui.

Passos (uma vez configurado o pipeline):
```bash
cp pipeline/dashboard.html /caminho/do/repo/phibo-dashboard/index.html
cd /caminho/do/repo/phibo-dashboard
git add index.html
git commit -m "Atualiza painel — <mês/ano>"
git push
```
(ou integrar isso como último passo de um script único — ver "Próximos
passos" abaixo.)

**Nota sobre o botão "Exportar CSV"** da aba Transações: ele usa
`window.claude.use('downloads')`, uma capacidade que só existe dentro de um
Artifact publicado no claude.ai. No GitHub Pages esse objeto não existe, e
o template já trata isso com elegância — o botão simplesmente fica
escondido (`if (!downloadsCap) btn.style.display = 'none'`), sem erro. Se
quiser o botão funcionando no GitHub Pages, é preciso trocar essa parte do
`dashboard_template.html` por um download via blob/`<a download>` nativo do
navegador (poucas linhas) — ainda não feito nesta migração.

**Nota de privacidade:** a tabela de Transações (e os exports) mostra nome
e telefone de cliente. Se `psalesc/phibo-dashboard` for um repo público (e
o GitHub Pages padrão também é público, a não ser que o repo seja privado
num plano GitHub pago), qualquer pessoa com o link vê esses dados — igual
já era verdade no link de Artifact (privado até ser compartilhado, mas sem
controle de acesso além disso). Vale o Pedro confirmar a visibilidade do
repo antes do primeiro push, se isso for uma mudança em relação ao que ele
já aceitava hoje.

## Automação — GitHub Actions

Implementado em `.github/workflows/atualizar-painel.yml` (15/10/2026): de
hora em hora, 8h–20h de Brasília, seg–sáb, mais o botão "Run workflow".
Passos: auth no Google via WIF → `fetch_sheets.py` → `build_dashboard.py`
(com `TZ=America/Sao_Paulo`) → copia pra `index.html` na raiz → commit/push
se mudou. Como o "Atualizado em" muda a cada rodada, praticamente toda
rodada gera um commit — esperado.

**Decisão de privacidade (Pedro, 15/10/2026):** o repo `psalesc/phibo-dashboard`
é **público** e continua público, ciente de que `index.html` e
`pipeline/historical/*.json` expõem nome e telefone de clientes.

Texto original da recomendação:

O que falta pra isso rodar sozinho, sem o Pedro (ou alguém) disparar
manualmente: um GitHub Actions (ou cron local) que, numa cadência (ex.
diária, ou algumas vezes ao dia): (1) roda `fetch_sheets.py` com a
`service_account.json` guardada como GitHub Secret; (2) roda
`build_dashboard.py`; (3) copia o resultado pra `index.html` e faz commit/push
se houver mudança. Isso fecha o circuito — é o que faz valer a pena ter
migrado pra fora do Cowork.

## Lições de processo (não repetir)

- **Sempre re-baixar as 5 planilhas antes de investigar uma divergência de
  números** — inclusive as de vendedor, não só master/Controle Vendas.
  Achado 14-15/09: uma investigação de "agosto tem 558 vendas ou 620?" ficou
  presa porque `samuel.xlsx`/`bruno.xlsx` estavam com 17 dias de atraso (só
  master e Controle Vendas eram rebaixadas a cada rodada). Checar
  `modifiedTime` no Drive é mais rápido que descartar hipóteses por dedução.
- **`build_controle_vendas()` localiza os blocos por busca de texto, não por
  posição fixa**, porque o layout da Controle Vendas já deslocou de posição
  entre meses (ex. o bloco "RENOVAÇÕES"). Se um aviso de "não encontrado"
  aparecer no log, é mais provável ser isso do que um bug — confira o layout
  real da aba do mês antes de alterar código.
- **Nunca reverter as 3 regras de reconciliação** sem reler o achado que
  gerou cada uma (seção acima) — os números "nativos" de cada planilha de
  vendedor já se mostraram não-confiáveis de formas diferentes e recorrentes.

## O que falta fazer (pendências)

- [x] Service account configurada (via WIF, sem chave — ver "Acesso aos
      dados"). Falta o primeiro teste real, que só acontece rodando o
      workflow no GitHub.
- [x] Regenerados `agosto_2026.json` e `setembro_2026.json` (setembro com
      638 fechamentos, ver nota em "Primeira rodada").
- [x] Workflow do GitHub Actions criado. Falta subir os arquivos pro repo e
      rodar a primeira vez.
- [ ] Virar o CONFIG do `Dashboard_AoVivo.gs` para outubro (master ainda em
      setembro em 15/10 — o build contorna, mas o master fica errado).
- [ ] Decidir se o botão "Exportar CSV" deve ser reescrito para
      funcionar fora do claude.ai (ver nota em "Publicação").
- [x] Visibilidade do repo: público, decisão do Pedro (ver "Automação").
- **Do lado do Pedro, sempre que o mês virar:** trocar o CONFIG do
  `Dashboard_AoVivo.gs`, criar a aba nova em Controle Vendas, preencher a
  Meta de cada vendedor nela. Em outubro/2026 especificamente, essas duas
  coisas ainda não tinham sido feitas na última rodada manual (05/10) — a
  Meta de outubro estava zerada na Controle Vendas. Em 15/10 a Controle
  Vendas já tinha outubro com meta (R$ 588k); o master ainda não.
