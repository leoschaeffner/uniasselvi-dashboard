# Plano de migração — VinciLab → Supabase + hospedagem

Documento de **planejamento**, não de implementação. Nenhuma conta foi criada,
nenhuma linha de código de migração foi escrita. Decisão final é do Leo —
este documento é o mapa de opções com uma recomendação.

Relaciona-se com o mapa de evolução de 4 níveis já desenhado anteriormente
(trocar onde roda → banco de dados → API+login → tempo real). O que está
aqui é essencialmente o **Nível 2 (banco de dados) + Nível 3 (API+login)**
feitos ao mesmo tempo, porque é isso que resolve o problema real que motivou
a pergunta: segregar acesso por pessoa/papel.

---

## 0. Por que isso importa (o problema real, não o modismo)

Hoje, com os 4 portais (dashboard, coordenadores, gestor, laboratórios):

- A "segurança" é uma **senha única compartilhada por portal**, igual pra
  todo mundo que a recebe.
- Depois que o navegador decifra o JSON, **o dado inteiro está lá** — um
  coordenador que só deveria ver o curso dele tecnicamente tem acesso a
  todos os outros cursos; o corte por curso é feito em JavaScript, não é
  controle de acesso de verdade. Qualquer pessoa com um pouco de DevTools
  vê tudo.
- Não há como saber **quem** acessou o quê, nem revogar o acesso de uma
  pessoa específica sem trocar a senha de todo mundo.

Isso é uma limitação estrutural da arquitetura atual (estático + senha +
decifra no cliente), não um bug pontual. Resolver de verdade exige banco +
autenticação individual + controle de acesso no servidor (RLS) — exatamente
a motivação do Supabase. **Vercel é uma questão separada** (só hospedagem) —
ver §2 sobre por que não recomendo trocar a hospedagem junto.

---

## 1. Pesquisa: limites reais do tier gratuito (confirmados em supabase.com/pricing e vercel.com/pricing, 2026-10-09)

### Supabase Free
| Limite | Valor confirmado |
|---|---|
| Armazenamento de banco (Postgres) | 500 MB por projeto |
| Projetos ativos simultâneos | **2** (não é "ilimitado dev+staging+prod") |
| Egress/bandwidth do banco | 5 GB + 5 GB cacheado |
| Usuários do Auth (MAU) | 50.000/mês |
| Storage de arquivos | 1 GB |
| Invocações de Edge Functions | 500.000/mês |
| **Pausa por inatividade** | projeto free **pausa após 1 semana sem uso** (dado não é perdido, mas fica inacessível até reativar manualmente ou via API) |

Não há limite documentado de **número de linhas** — o limite é por **tamanho
em MB**, o que muda a forma de pensar o volume: o JSON decifrado hoje passa
de 40-50 MB, mas isso inclui **múltiplas visões pré-agregadas do mesmo dado**
(por curso, por ordem, por tutor, por polo...). Num banco relacional normal
você guarda o fato uma vez só e calcula os agregados via query/view — o
tamanho real em disco tende a ser bem menor que 40-50 MB.

**Medido na Fase 0 (09/10/2026, a partir do `saida/dashboard.html` real,
rodada de 02/10/2026, decifrado — ver `supabase/schema_fase1_dashboard.sql`
para o detalhe por tabela):**

| Entidade | Volume real |
|---|---|
| Tutores ativos | 621 |
| Tutores desligados | 5 |
| Polos distintos | ~235–251 (duas fontes do dict normalizam o nome de polo de forma levemente diferente — mesmo problema do PATCH 149; o schema novo unifica isso numa única tabela `polos`) |
| Categorias amplas (CAT_MAP) | 6 |
| Práticas no catálogo oficial | 303 |
| Submissões de portfólio (snapshot atual, tutores ativos) | 5.746 |
| Ofertas de gerenciamento (GIOCONDA), semestre ativo 2026/2 | 24.531 |
| Ofertas de gerenciamento, semestre anterior 2026/1 | 14.796 |
| Ofertas de gerenciamento, soma dos 2 semestres disponíveis | 39.327 |
| Matrículas distintas (hub de alunos) | 74.636 (só agregado por polo/categoria — nunca matrícula/nome individual, ver LGPD) |
| Vagas de RH (Lotação) | 399 |

O JSON decifrado inteiro tem **40,9 MB** nesta rodada — a estimativa grosseira
(linhas × tamanho médio de linha, sem as visões pré-agregadas repetidas) para
o schema relacional normalizado da Fase 1 (tutores + ofertas de gerenciamento,
que é de longe a maior tabela + submissões + dimensões de polo/categoria/
prática + vagas RH) fica na faixa de **~15–30 MB mesmo com folga generosa e
índices** — abaixo de 10% do limite de 500 MB do Supabase Free. **Confirmado:
cabe folgado**, inclusive projetando alguns semestres de crescimento acumulado
antes de se aproximar do limite.

### Vercel Hobby (free) — achado mais importante da pesquisa
| Limite | Valor confirmado |
|---|---|
| Transferência de dados | 100 GB/mês |
| Invocações de função | 1.000.000/mês |
| **Uso permitido** | citação literal da página oficial: **"Our Hobby plan is for personal, non-commercial use"** |
| Equipe | 1 assento de desenvolvedor — não dá pra adicionar colaboradores pagos nem "viewer" |

**Isso é uma letra miúda que pega de surpresa de verdade**: o VinciLab é uma
ferramenta de gestão usada pela instituição (UNIASSELVI) para operação real
de tutoria/laboratórios — isso é uso institucional/organizacional, não
"pessoal". A Vercel reserva o direito de suspender contas Hobby usadas fora
desse escopo. Não encontrei um caso documentado de suspensão exatamente
deste tipo de projeto (não é possível confirmar isso por pesquisa — é uma
política de Termos de Uso, não uma trava técnica automática), mas a
exposição de risco é real: se a Vercel decidir aplicar a regra, o site cai
sem aviso, e a única saída seria pagar o Pro (US$20/mês/assento) às pressas.

**Recomendação derivada disso:** não troque a hospedagem para Vercel agora.
Use o Supabase só como **banco + autenticação + RLS**, e mantenha a
hospedagem do HTML/JS onde já está validado e sem essa cláusula de risco:
- **GitHub Pages** (como hoje, zero custo, sem essa restrição, já funciona),
  ou
- o **servidor próprio com acesso root que já existe** (Nível 1 do mapa de
  evolução) — também remove essa dependência de terceiro de uma vez.

Isso também é mais **reversível**: trocar banco sem trocar hospedagem é uma
mudança por vez; se o Supabase não servir, você desfaz só uma peça, não
duas. Se mais adiante for necessário rodar lógica server-side (ex.: um
endpoint que verifica login antes de servir dado — ver Fase 3/4), aí sim
vale reavaliar Vercel (ou Edge Functions do próprio Supabase, que já
resolvem boa parte disso sem precisar de outro provedor).

---

## 2. Modelo de autorização: dos 4 portais pra RLS

Hoje: 1 senha fixa por portal = 1 "papel" compartilhado. O Supabase permite
granularidade real via **Row Level Security** (regra de acesso presa à
linha do banco, avaliada no servidor, não no navegador).

Proposta de mapeamento (login individual por e-mail, via Supabase Auth):

| Portal atual | Papel no novo modelo | Regra de RLS |
|---|---|---|
| `coordenadores.html` | `coordenador` + `curso_id` (coluna de metadado no usuário) | Só enxerga linhas de tutores/ofertas/alunos onde `curso_id` bate com o dele — **enforced no banco**, não mais em JS. Isso fecha o vazamento estrutural descrito no §0. |
| `gestor.html` | `gestor` | Vê agregados de todos os cursos (visão executiva) — pode continuar sem granularidade fina, é o papel que já é "global" por natureza. |
| `laboratorios.html` | `laboratorios` | Vê dados de laboratório/vínculo/vistoria de todos os polos (hoje também já é visão operacional ampla). |
| `dashboard.html` (`index.html`) | `admin` / `dono` | Acesso total — hoje é literalmente isso, só que com senha compartilhada em vez de contas nomeadas. |

Ganhos concretos sobre o modelo atual:
- Dá pra **revogar o acesso de uma pessoa só** (hoje: só trocando a senha de
  todo mundo).
- Dá pra **saber quem acessou o quê** (auditoria — hoje: zero rastro).
- O corte por curso do coordenador passa a ser **garantia do banco**, não
  "o JS não mostra, mas o dado chegou no navegador" — relevante pro padrão
  de LGPD que o projeto já segue em outras partes (ex.: candidatos de RH só
  agregados).

Ponto de atenção sério de segurança: a diferença entre a **anon key** (chave
pública, segura de expor no cliente, só funciona dentro do que a RLS
permitir) e a **service_role key** (ignora toda RLS, acesso total ao banco).
A `service_role key` só pode existir no `processar.py`/ETL (ambiente do
GitHub Actions, como secret) — se ela vazar pro frontend por engano, é
estritamente pior que a senha única de hoje, porque dá acesso de escrita
também, não só leitura. **Esse risco é detalhado e mitigado no §6.**

---

## 3. Plano de fases (incremental, sem quebrar o pipeline atual)

Princípio: GitHub Actions + Pages **continuam rodando em paralelo** até cada
fase estar validada. Nada é desligado até o substituto provar que funciona
igual ou melhor, por pelo menos 1-2 semanas de uso real.

**Ordem escolhida pelo Leo: Dashboard → Coordenadores → Gestor →
Laboratórios** — do portal mais crítico/mais usado pro de menor uso. Isso
inverte a ordem que eu recomendaria por padrão (menor uso primeiro, pra
"esquentar" o processo num ambiente de baixo risco antes de mexer no que
todo mundo usa todo dia). Registro o trade-off explicitamente, porque é
real:

- **O que se perde começando pelo mais crítico primeiro:** não há fase de
  aquecimento. Os primeiros erros de modelagem de RLS, de schema, de
  performance de query — o tipo de erro que é normal cometer numa primeira
  tentativa — vão acontecer direto no portal com mais gente olhando e mais
  dado sensível, não num portal de baixíssima exposição como Laboratórios.
  Um erro de RLS mal configurada aqui (ex.: policy que libera mais do que
  deveria) tem o maior raio de impacto possível dentre as 4 opções.
- **Mitigação que eu recomendo adotar nesta ordem, não opcional:**
  1. Rodar o Dashboard em paralelo (Supabase + sistema antigo lado a lado)
     por **mais tempo que o mínimo das outras fases** — sugiro 3-4 semanas
     de uso real sem incidente antes de considerar desligar qualquer parte
     do pipeline antigo, não as 1-2 semanas que bastariam pra um portal de
     menor risco.
  2. Validar com um **usuário de teste próprio do Leo** (não um usuário
     real da instituição) durante a maior parte dessa janela, só expondo
     pra uso real depois de uma primeira rodada de confiança.
  3. **Não apagar nem desligar o pipeline antigo do Dashboard em nenhuma
     circunstância** até a Fase 4 (Laboratórios) também estar migrada — ou
     seja, o `index.html` cifrado continua sendo gerado pelo GitHub Actions
     como rede de segurança durante **todo** o projeto, não só durante a
     fase do próprio Dashboard. Isso custa zero (o pipeline já roda) e é a
     forma mais barata de manter a reversibilidade mesmo começando pelo
     caminho mais arriscado.

### Fase 0 — Preparação (não migra nada ainda)
- **Esforço:** 3-5 dias.
- **Pré-requisitos:** nenhum além de decidir seguir em frente.
- **O que fazer:** criar o projeto Supabase (grátis), desenhar o schema
  relacional a partir do dict que `processar.py` já produz (tutores, polos,
  ofertas, cursos, usuários/papéis), e **medir o volume real** (quantas
  linhas em cada tabela hoje) pra confirmar que cabe nos 500 MB antes de
  prometer isso como certo.
- **Risco principal:** nenhum — é só desenho, roda isolado.
- **O que destrava:** visibilidade real de "cabe ou não cabe" antes de
  qualquer compromisso.
- **Reversão:** trivial — é só apagar o projeto Supabase, nada em produção
  foi tocado.

### Fase 1 — Dashboard principal (o mais crítico, migra primeiro)
- **Esforço:** 4-6 semanas — é o portal "fonte da verdade" pros outros, com
  mais dado, mais visões agregadas e mais gente de confiança olhando. Sem a
  fase de aquecimento que a ordem alternativa daria, é realista prever mais
  tempo aqui do que se este fosse o último portal a migrar (iterações de
  RLS/schema que normalmente já estariam resolvidas nos portais anteriores
  vão acontecer aqui pela primeira vez).
- **Pré-requisitos:** Fase 0 concluída; schema completo das tabelas usadas
  pelo Dashboard (tutores, ofertas, cursos, polos, agregados); login
  individual configurado pra quem usa este portal hoje (inclusive o próprio
  Leo).
- **O que muda:** `index.html` passa a buscar dado do Supabase (anon key +
  RLS) em vez de decifrar um JSON embutido. Pode continuar hospedado no
  GitHub Pages — é só JS fazendo fetch pra outro lugar.
- **Principal risco:** técnico — é a maior superfície de schema/RLS do
  projeto pra acertar de uma vez, sem ter validado o modelo em nada antes;
  e de manutenção — qualquer decisão errada de schema aqui é a mais cara de
  desfazer depois, porque os outros 3 portais provavelmente vão herdar
  convenções deste. Mitigação: ver as 3 medidas no início da §3 (paralelo
  mais longo, usuário de teste antes de uso real, pipeline antigo
  permanece ativo até o fim do projeto todo).
- **O que isso destrava:** valida o modelo inteiro (Auth + RLS + fetch no
  cliente) direto no portal que mais importa, sem depender de migrar os
  outros 3 primeiro — se a prioridade do Leo é ver o ganho de segurança
  rápido onde ele mais importa, é aqui que ele aparece primeiro.
- **Reversão:** o `index.html` atual continua sendo gerado pelo pipeline
  antigo em paralelo durante toda a janela de validação (e, por decisão
  deste plano, até o fim do projeto inteiro — ver mitigação acima); reverter
  é apontar de volta pra essa versão.

### Fase 2 — Portal Coordenadores (aqui está o ganho real de segurança)
- **Esforço:** 3-4 semanas — é o mais trabalhoso em granularidade de RLS,
  porque é onde o corte por curso precisa funcionar perfeitamente — é
  literalmente o problema que motivou a migração (§0).
- **Pré-requisitos:** Fase 1 validada por pelo menos 3-4 semanas de uso real
  sem incidente; mapeamento `curso_id` por usuário revisado com cuidado
  (reaproveita `CURSO_PARA_CATEGORIA` já existente no código, mas agora
  vira metadado de usuário, não filtro client-side).
- **Principal risco:** manutenção — qualquer curso mal mapeado deixa um
  coordenador sem ver o próprio curso (visível, reclamado rápido) **ou**
  vendo curso de outro (não visível, silencioso — o pior caso). Recomendo
  um período de auditoria ativa (como já foi feito em 17/09 pro filtro
  atual) logo após o go-live desta fase.
- **O que isso destrava:** fecha de vez o vazamento estrutural do §0.
- **Reversão:** `coordenadores.html` antigo continua em paralelo até
  confiança.

### Fase 3 — Portal Gestor
- **Esforço:** 1-2 semanas.
- **Pré-requisitos:** Fase 2 validada por pelo menos 1-2 semanas de uso real
  sem incidente.
- **Risco principal:** este portal tem mais cards/fontes (ocorrências,
  turnover, engajamento) — mais superfície de RLS pra acertar, mas menor que
  Coordenadores (sem a granularidade por curso).
- **Reversão:** mesma lógica das fases anteriores — paralelo até confiança.

### Fase 4 — Portal Laboratórios (menor uso, migra por último)
- **Esforço:** 1-2 semanas — a essa altura o modelo (Auth + RLS + fetch) já
  está validado 3 vezes, então é a fase mais mecânica do projeto.
- **Pré-requisitos:** Fase 3 validada; schema das tabelas de
  laboratório/vínculo/vistoria criado; login individual configurado pra
  quem usa esse portal hoje.
- **O que muda:** `laboratorios.html` passa a buscar dado do Supabase
  (anon key + RLS) em vez de decifrar um JSON embutido.
- **Principal risco:** baixo a essa altura — o erro clássico (policy
  `USING (true)` esquecida) já deveria ter sido pego nas 3 fases
  anteriores. Testar mesmo assim com usuário de teste antes de liberar.
- **O que isso destrava:** fecha o último portal, viabilizando a Fase 5.
- **Reversão:** `laboratorios.html` atual continua existindo e sendo
  gerado pelo pipeline antigo em paralelo até confiança.

### Fase 5 (opcional, só depois de tudo validado) — Aposentar o pipeline antigo
- Desligar a geração do JSON cifrado e o GitHub Pages só depois de TODOS os
  portais terem rodado em paralelo com sucesso por um tempo. Não há pressa
  nem necessidade de fazer isso — pode conviver indefinidamente se for mais
  confortável manter os dois.

**Estimativa total, calendário real (não dias de trabalho contínuo):** dado
que é o Leo com apoio de IA e não uma equipe dedicada, e que o projeto
continua recebendo pedidos novos em paralelo (ver histórico de PATCHes),
uma estimativa honesta de calendário é **4 a 6 meses** do início da Fase 0
até o portal Laboratórios (Fase 4) estar validado — mais do que os "3 a 5
meses" que esta mesma soma de esforços daria na ordem alternativa (menor
risco primeiro), porque começar pelo portal mais crítico tende a gastar
mais tempo bruto ali (ver Fase 1) e também costuma gerar mais interrupções
reais (um incidente no portal mais usado puxa atenção/prioridade de volta
pra ele, atrasando o resto da fila) do que um incidente equivalente num
portal de baixo uso.

---

## 4. O que muda no `processar.py`

Hoje `processar()` e as funções irmãs (`processar_vagas_rh`,
`processar_ocorrencias`, `processar_vistoria` etc., ver
`MAPA_DO_CODIGO.md` §5-6) terminam gerando um **dict Python gigante**, que
vira JSON, que é cifrado e embutido no HTML.

Com Supabase, o fluxo passa a ser: ETL lê as mesmas planilhas → em vez de
montar o dict final pra um JSON, faz **upsert nas tabelas do Postgres** via
chave de escrita (ver §6 sobre qual chave usar e como protegê-la — GitHub
Actions continua sendo o lugar certo pra essa chave ficar, como secret,
nunca no cliente).

Reaproveitável (~70-80% do código, estimativa, não medido linha a linha):
- Toda a leitura/parsing das planilhas (`verificar_e_localizar`,
  `carregar_lotacao`, `processar_vagas`, `processar_vagas_rh`,
  `processar_ocorrencias`, `processar_vistoria`, `carregar_alunos_hub`).
- Toda a lógica de cruzamento/matching (nomes de tutor, polos, categorias,
  dedup de alunos, GIOCONDA) — isso é regra de negócio pura, independe de
  onde o resultado vai parar.

Precisa reescrever/descartar:
- As funções que hoje montam **visões pré-agregadas pro frontend consumir
  pronto** (`_stats_semestre`, os vários `ger_*`, `dados_por_semestre`) —
  num banco relacional, essas visões idealmente viram **queries/views SQL**
  calculadas sob demanda (ou materializadas), não dicts pré-calculados em
  Python. Isso é a parte mais trabalhosa da reescrita, não a leitura do dado.
- `cifrar_dados`, `gerar_html`, `gerar_html_coordenadores` — deixam de fazer
  sentido como estão (não tem mais JSON cifrado embutido); viram, na
  prática, "upsert no Supabase" em vez de "gerar HTML".
- Qualquer lógica de "esse campo não entra no dict final" (ex.: PII de
  candidatos de RH, hoje resolvido excluindo campos do dict) precisa virar
  uma decisão **por tabela/coluna** no schema do banco — mais explícito, mas
  precisa ser revisto com cuidado pra não esquecer nenhuma regra de LGPD
  que hoje é implícita no "não incluí esse campo no retorno".

---

## 5. Pontos de atenção de segurança/LGPD na transição

- **Nunca** colocar a `service_role key` em código de frontend/commit público
  — ela ignora toda RLS. É o equivalente a publicar a senha-mestra de tudo.
- Enquanto um portal roda nos dois sistemas em paralelo (fase de validação),
  **o dado continua tão exposto quanto hoje no sistema antigo** — a
  migração não piora nada, mas também não resolve nada até a fase daquele
  portal específico estar completa. Não anunciar "mais seguro" antes da
  hora.
- Regras de LGPD já aplicadas hoje (ex.: candidatos de RH só agregados,
  nunca nominais) precisam ser **replicadas explicitamente no schema**
  (ex.: nem ter coluna de nome/telefone/e-mail de candidato na tabela, não
  só "RLS esconde") — manter o princípio de "não existe o dado onde não
  precisa existir", que é mais forte que só controlar quem vê.
- Testar RLS com usuário de teste em **cada papel antes de liberar**, em
  cada fase — é o tipo de erro que fica invisível até alguém notar (ou um
  coordenador notar que vê o curso errado, o que já é tarde).

---

## 6. Segurança da chave de escrita (GitHub Actions → Supabase) — resposta direta à pergunta "isso não seria uma brecha pro meu banco?"

Resposta honesta: **sim, é um risco real e novo**, e precisa ser tratado como
tal, não varrido pra baixo do tapete porque "Supabase é mais moderno que
senha fixa". Vale deixar isso bem explícito antes de qualquer outra coisa:

> **Esta migração troca um tipo de risco por outro, não troca "risco" por
> "sem risco".** Hoje, se um secret vazar, o pior caso é um link de planilha
> do OneDrive exposto — chato, mas contido, e sem acesso de escrita a nada.
> Com Supabase, o pipeline (GitHub Actions) passa a precisar de uma chave
> capaz de **escrever** no banco de produção. Se essa chave vazar e for a
> `service_role key` (a opção mais simples de configurar), o vazamento dá
> acesso de **leitura e escrita a todo o banco**, ignorando toda RLS — pior
> do que qualquer coisa possível hoje. O ganho líquido de segurança da
> migração como um todo (RLS, login individual, revogação por pessoa)
> **só se realiza se este risco novo for mitigado de verdade** — caso
> contrário, a migração pode deixar o projeto com uma superfície de ataque
> pior, não melhor, nesse ponto específico.

### 6.1 Nunca logar a chave em nenhum step do workflow
- O secret do GitHub Actions (`SUPABASE_SERVICE_ROLE_KEY` ou equivalente)
  nunca deve aparecer em `print`/log de debug — isso já é prática do
  projeto hoje (ver regras de log do `processar.py`).
- **Cuidado específico que não existia antes:** bibliotecas cliente do
  Supabase (e de Postgres em geral) às vezes **ecoam a request completa**
  em mensagens de erro quando uma chamada falha (timeout, erro de rede,
  payload inválido) — isso pode incluir o header `Authorization` com a
  chave dentro, mesmo sem nenhum `print` explícito da chave em si. Qualquer
  `try/except` que capture e logue o erro da chamada ao Supabase precisa
  **sanitizar a mensagem antes de logar** (ex.: remover qualquer coisa que
  pareça um JWT/Bearer token) em vez de assumir que "eu não imprimi a chave,
  então estou seguro".
- GitHub Actions já mascara automaticamente o valor literal de um secret
  nos logs — mas essa máscara só funciona se o valor aparecer **exatamente
  igual** ao secret configurado; não protege contra o erro acima.

### 6.2 Usar uma chave mais restrita em vez da `service_role` completa — pesquisado, não presumido
Pesquisei se o Supabase oferece algo entre a `anon key` (só leitura, sujeita
a RLS) e a `service_role key` (acesso total, ignora RLS). **Confirmei que
sim, existe meio-termo**, mas não é uma "chave" no mesmo sentido da anon/
service_role — é um **role customizado do Postgres**:

- O Supabase expõe o banco como Postgres de verdade, então é possível criar
  um role de login próprio (`create role app_etl login password '...'`)
  com `GRANT` apenas de `INSERT`/`UPDATE` nas tabelas específicas que o
  pipeline escreve, e nada de `DELETE`/`DROP`/acesso às demais tabelas.
  Essa é uma feature padrão de roles do Postgres, documentada no guia de
  roles do Supabase.
- Essa opção exige conectar **direto no Postgres via connection string**
  (não pela API REST com `anon`/`service_role` key) — é uma forma de acesso
  diferente da que o projeto usaria pro frontend, mas é exatamente o padrão
  certo pra um job de ETL como o `processar.py`/GitHub Actions, que não
  precisa passar pela API pública.
- **Não encontrei confirmação explícita** de que esse tipo de role
  customizado tem alguma limitação no plano Free (a documentação de roles
  não menciona diferenças por plano) — a criação de roles e grants é SQL
  padrão do Postgres, então é razoável esperar que funcione igual, mas isso
  deveria ser testado na Fase 0 antes de depender disso como premissa.

**Recomendação derivada:** usar um role customizado (`app_etl` ou nome
similar) com permissão só de `INSERT`/`UPDATE` nas tabelas que o pipeline
escreve, conectando via connection string direta, **em vez de** colocar a
`service_role key` no GitHub Actions. Isso limita o estrago de um vazamento
a "alguém pode inserir/alterar linhas nessas tabelas específicas", não
"alguém tem o banco inteiro". Se isso não se confirmar viável na prática
(Fase 0), a segunda melhor opção é a `service_role key` mesmo, mas aí as
mitigações dos itens 6.1, 6.3 e 6.4 passam de "boa prática" para
"obrigatório".

### 6.3 Plano de rotação da chave em caso de suspeita de vazamento
"Rotacione a chave" sozinho não é um plano — o processo precisa estar
escrito antes de precisar dele, não inventado sob pressão:

1. **Gerar a nova chave/credencial primeiro**, sem desativar a antiga ainda
   (Supabase permite ter a chave antiga ativa enquanto gera substituta, no
   caso da `service_role`; pra um role customizado, é trocar a senha do
   role via `ALTER ROLE ... PASSWORD`).
2. **Atualizar o secret no GitHub Actions** (`Settings > Secrets and
   variables > Actions`) com o novo valor.
3. **Rodar o workflow manualmente uma vez** (`workflow_dispatch`, via
   `git-github-vinci`) pra confirmar que o pipeline ainda escreve com a
   credencial nova antes de revogar a antiga.
4. **Revogar/apagar a credencial antiga** só depois do passo 3 confirmado —
   nunca revogar antes de validar a nova, senão o pipeline para de escrever
   sem aviso no próximo cron.
5. **Auditar o que foi escrito/lido com a chave suspeita** no intervalo
   entre o vazamento suspeito e a rotação — o Supabase mantém logs de API
   (Free tier tem retenção limitada, confirmar o período exato na conta no
   momento do incidente) que ajudam a entender o alcance real do incidente.
6. Se o vazamento veio de um commit público (ex.: secret commitado por
   engano), tratar o commit como comprometido permanentemente mesmo após
   remoção — reescrever histórico do Git não garante que a chave não foi
   vista por scrapers automáticos de secrets do GitHub, que existem e
   varrem repositórios públicos ativamente. Rotacionar é obrigatório nesse
   caso, não opcional "se desconfiar".

### 6.4 O que isto significa em termos do risco líquido do projeto
Deixando explícito, porque é fácil deslizar pra "Supabase = mais seguro em
tudo" sem perceber a troca: **este é um risco que o projeto não tinha
antes**. Hoje, o pior cenário de vazamento de secret é "uma planilha fica
visível pra quem não deveria". Depois da migração, o pior cenário possível
(se as mitigações acima não forem seguidas) é "o banco inteiro de produção
pode ser lido e escrito por quem tiver a chave". O ganho de segurança do
lado do usuário final (RLS, login individual, revogação por pessoa) é real
e vale a pena — mas ele não compensa automaticamente esse novo risco do
lado do pipeline; os dois precisam ser avaliados separadamente, e o segundo
só fica "resolvido" se as mitigações de 6.1-6.3 forem de fato implementadas,
não apenas planejadas aqui.

---

## 7. Quando o tier gratuito provavelmente deixa de servir

- **Supabase 500 MB:** não é um risco imediato pra este volume de dado
  (cadastro + ofertas da UNIASSELVI, sem vídeo/imagem) — mas só será
  confirmado com a medição da Fase 0. Se algum dia passar disso, o Pro
  (US$25/mês) resolve com folga (8 GB inclusos).
- **Pausa semanal por inatividade:** mitigável com um "keepalive" simples
  (o projeto já faz algo equivalente hoje, ver commit `6d1cd8b keepalive
  2026-10-05` no histórico do Actions) — um ping programado evita a pausa
  sem custo.
- **50.000 MAU de Auth / 500.000 invocações de Edge Function:** folga
  enorme pro número de pessoas que realmente usam os 4 portais hoje — não é
  um ponto de atenção previsível neste projeto.
- **2 projetos simultâneos (limite real, não óbvio):** não dá pra ter
  dev + staging + produção como 3 projetos separados grátis. Planejar
  testar com cuidado dentro do mesmo projeto (schemas/RLS de teste) em vez
  de assumir que dá pra espelhar um ambiente de homologação completo.
- **Vercel Hobby "non-commercial":** não é sobre volume, é sobre política de
  uso — por isso a recomendação de não usar Vercel pra isto (§1), e não um
  "vai custar quando crescer".

---

## 8. Recomendação (decisão é do Leo)

1. **Sim** ao Supabase como banco + Auth + RLS — resolve o problema real
   (segregação de acesso) que a arquitetura atual estruturalmente não
   resolve, e cabe no free tier pra este volume (a confirmar na Fase 0).
2. **Não** migrar a hospedagem pra Vercel agora — a cláusula "personal,
   non-commercial use" do Hobby é um risco real pra um sistema de gestão
   institucional. Manter GitHub Pages (como hoje) ou usar o servidor próprio
   com root já disponível (que também elimina outra dependência de
   terceiro, alinhado ao Nível 1 do mapa de evolução já desenhado).
3. Migrar na ordem escolhida pelo Leo — **Dashboard → Coordenadores →
   Gestor → Laboratórios** —, sempre em paralelo com o pipeline atual até
   validar, nunca como "big bang". Registrado o trade-off: essa ordem
   começa pelo portal de maior exposição sem fase de aquecimento, por isso
   a janela de validação da Fase 1 (Dashboard) deve ser mais longa que as
   demais (3-4 semanas, não 1-2) e o pipeline antigo do Dashboard deve
   continuar ativo até o projeto inteiro estar concluído, não só até a
   própria Fase 1 ser validada.
4. Tratar a Fase 2 (Coordenadores) como a entrega que fecha o objetivo
   original — é onde o vazamento estrutural de hoje é fechado de verdade.
5. **Não** usar a `service_role key` completa no GitHub Actions sem antes
   testar, na Fase 0, se um role customizado do Postgres (só `INSERT`/
   `UPDATE` nas tabelas do pipeline) é viável — ver §6.2. Se não for viável,
   seguir com `service_role key` só com as mitigações do §6 (nunca logar,
   plano de rotação escrito, risco líquido reconhecido) tratadas como
   obrigatórias, não opcionais.
