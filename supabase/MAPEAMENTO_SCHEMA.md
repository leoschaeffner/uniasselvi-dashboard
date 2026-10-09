# Mapeamento campo antigo (dict/JSON) → tabela.coluna nova (Supabase)

Gerado na Fase 0 da migração (ver `PLANO_MIGRACAO_SUPABASE_VERCEL.md` e
`supabase/schema_fase1_dashboard.sql`). Cobre as visões principais do
Dashboard (`template_dashboard.html`) — KPIs do topo, Gerenciamento por
semestre, Polos. Objetivo: evitar que o frontend quebre por engano ao trocar
a fonte de dado. **Atualizar este arquivo conforme mais seções forem
migradas** — não é definitivo, é o ponto de partida da Fase 1.

Convenção: `dict['campo']` é o campo de hoje (JSON cifrado/`DB` no JS);
`tabela.coluna` é o novo lugar. `(view)` indica que o campo vira uma coluna
calculada por uma VIEW SQL, não uma coluna armazenada.

---

## 1. KPIs do topo (`dict['kpis']`)

| Campo antigo | Novo | Observação |
|---|---|---|
| `kpis.total` | `v_kpis_tutores.total` (view) | `count(*) where ativo` |
| `kpis.enviaram` | `v_kpis_tutores.enviaram` (view) | tutor ativo com ≥1 linha em `submissoes_portfolio` |
| `kpis.pendentes` | `v_kpis_tutores.pendentes` (view) | `total - enviaram` |
| `kpis.atrasados` | **ainda não modelado como view** | Depende de `config_semestres_ordens.prazo` cruzado com `submissoes_portfolio` por Ordem — a lógica de "atrasado" (que Ordem o tutor deveria ter enviado e não enviou, considerando o prazo daquela Ordem) precisa ser replicada da função `_stats_semestre` (`processar.py` ~L1927) antes de virar SQL. Não presumir a regra — conferir com o Leo/`processar.py` linha a linha quando essa parte entrar em implementação. |
| `kpis.urgentes` | **idem acima** | Mesma dependência de prazo por Ordem. |
| `kpis.total_alunos` | `sum(matriculas_agregadas_hub.total_matriculas_distintas)` (view a criar) | **Atenção:** no dict de hoje este campo já vem substituído pela contagem do hub (`carregar_alunos_hub`), não é soma de `submissoes_portfolio.alunos_informados` — não confundir as duas fontes. |
| `kpis.total_polos` | `count(distinct polo_id)` sobre `tutores` (view a criar) | |
| `kpis.polos_ok` | **ainda não modelado** | Regra de "polo OK" (cobertura de categoria) precisa ser confirmada no código antes de virar view — não presumir. |

## 2. Gerenciamento por semestre (`dict['gerenciamento_por_semestre'][sem]`)

Fato bruto: tabela `ofertas_gerenciamento` (coluna `semestre` filtra o
equivalente a `dict['gerenciamento_por_semestre']['2026/2']` etc. — hoje
`dict['ger_*']` na raiz é sempre uma cópia do semestre ativo, isso deixa de
existir como "cópia": vira sempre `WHERE semestre = SEMESTRE_ATUAL`).

| Campo antigo | Novo | Observação |
|---|---|---|
| `ger_ofertas` (lista) | `select * from ofertas_gerenciamento where semestre = :sem` | Fato cru, 1 linha por oferta. 24.531 linhas medidas em 2026/2. |
| `ger_ofertas[].polo` | `ofertas_gerenciamento.polo_id` → join `polos.nome` | |
| `ger_ofertas[].categoria` | `ofertas_gerenciamento.categoria_id` → join `categorias.nome_exibicao` | |
| `ger_ofertas[].ordem` | `ofertas_gerenciamento.ordem` | |
| `ger_ofertas[].pratica` | `ofertas_gerenciamento.pratica_nome` | |
| `ger_ofertas[].tutor` | `ofertas_gerenciamento.tutor_id` → join `tutores.nome` | |
| `ger_ofertas[].tem_tutor` / `tem_agenda` / `gerenciado` | colunas booleanas homônimas | |
| `ger_ofertas[].alunos_mat` / `alunos_agend` | `alunos_matriculados` / `alunos_agendados` | |
| `ger_ofertas[].dt_agenda` / `hr_agenda` / `dia_semana` / `turno` | colunas homônimas | |
| `ger_ofertas[]._anomalia_ordem` / `_anomalia_ordem_futura` | `anomalia_ordem` / `anomalia_ordem_futura` | |
| `ger_kpis` (dict agregado) | `v_ger_kpis_por_semestre` (view, `where semestre = :sem`) | `pct_capacidade`/`total_capacidade` do dict (hoje sempre 0 nesta base) não modelado ainda — não há coluna de capacidade na fonte atual. |
| `ger_polo` (lista) | `v_ger_polo_por_semestre` (view) | |
| `ger_cat` (lista) | `v_ger_cat_por_semestre` (view) | |
| `ger_ordem` (lista) | `v_ger_ordem_por_semestre` (view) | |
| `ger_agendas` (lista, por polo) | **view a criar** (`group by semestre, polo_id`, `array_agg(data_agenda)` para `datas_agenda`) | Não incluída ainda no `.sql` — padrão igual às outras `v_ger_*`, adicionar quando for implementar de verdade. |
| `ger_contratacao` (lista, por polo×categoria) | **view a criar** (`group by semestre, polo_id, categoria_id`) | Campo `tutores: []` do dict vira `array_agg(tutor_id)` ou um join separado — decidir quando implementar (pode ter 0, 1 ou N tutores por combinação). |
| `ger_anomalias_ordem` | `select * from ofertas_gerenciamento where anomalia_ordem or anomalia_ordem_futura` | |
| `tem_gerenciamento` (bool) | deixa de fazer sentido como campo — vira `exists(select 1 from ofertas_gerenciamento where semestre = :sem)` na app | |

## 3. Polos (`dict['polo_stats']`)

| Campo antigo | Novo | Observação |
|---|---|---|
| `polo_stats[].POLO` / `.polo` / `.n` | `polos.nome` | As 3 chaves do dict são o MESMO valor (redundância do JSON) — vira 1 coluna só. |
| `polo_stats[].total` / `.t` | `v_kpis_tutores_por_polo.total` (view a criar, `group by polo_id`) | |
| `polo_stats[].enviaram` / `.e` | idem, `count(*) filter (where tutor tem submissão)` | |
| `polo_stats[].atrasados` | **mesma pendência do item 1** (`kpis.atrasados`) — depende de prazo por Ordem | |
| `polo_stats[].alunos` / `.a` | `matriculas_agregadas_hub` somado por `polo_id` (não `submissoes_portfolio` — são fontes diferentes, não confundir) | |
| `polo_stats[].pend` | `total - enviaram` | |
| `polo_stats[].pct` | `round(100.0 * enviaram / total, 0)` | |
| `polo_stats[].envios` | `count(*)` em `submissoes_portfolio` por polo (via tutor) | |
| `polo_stats[].contatos[]` | `contatos_polo` (tabela própria, FK `polo_id`) | |

## 4. Observações gerais de reconciliação

- **Contagem de polos diverge entre fontes hoje** (235 em `polo_stats` vs 251
  em `ger_polo`, medido na mesma rodada) — isso é um sintoma direto do
  PATCH 149 (6+ implementações locais de normalização de nome de polo no
  `processar.py`). O schema novo **resolve isso por construção**: existe
  UMA tabela `polos` com UMA coluna `nome_normalizado`, e tudo mais referencia
  por FK — não há mais "visão A conta 235, visão B conta 251" possível,
  porque as duas passam a vir da mesma tabela. **Atenção na migração de
  dado**: ao popular `polos` pela primeira vez, decidir qual das duas
  normalizações vira a canônica é uma decisão de regra de negócio, não
  trivial — não resolver por suposição, confirmar com dado real/Leo antes
  de escrever o script de carga definitivo (mesma regra 1 do `dados-etl`).
- **`kpis.total_alunos` do topo ≠ soma de `submissoes_portfolio`** — são
  fontes diferentes (hub de matrículas vs. portfólio). Ver item 1. Isso já é
  verdade no sistema atual (regra de negócio documentada no
  `MAPA_DO_CODIGO.md` §7, "quando o hub CSV está disponível, ele é a fonte
  de verdade") — só está sendo reafirmado aqui para não se perder na
  tradução pro schema novo.
- Views marcadas "**ainda não modelado**"/"**a criar**" não devem ser
  inventadas por suposição — são exatamente os pontos onde a regra de
  negócio (prazo por Ordem, "polo OK", capacidade) precisa ser lida direto
  do `processar.py` (`_stats_semestre`, `_injetar_tutores_sem_oferta` etc.)
  linha a linha antes de virar SQL, não estimada.
