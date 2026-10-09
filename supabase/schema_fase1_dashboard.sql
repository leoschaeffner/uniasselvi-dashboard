-- ============================================================================
-- VinciLab — Schema Fase 1 (Dashboard principal)
-- ============================================================================
-- Gerado na Fase 0 da migração Supabase (ver PLANO_MIGRACAO_SUPABASE_VERCEL.md).
-- Fonte: dict real produzido por processar.py, medido a partir do
-- saida/dashboard.html decifrado (rodada 02/10/2026, 621 tutores ativos,
-- 39.327 ofertas somando os 2 semestres, 5.746 submissões, 74.636 matrículas
-- distintas — ver números completos no PLANO_MIGRACAO_SUPABASE_VERCEL.md §1).
--
-- ESCOPO: só tabelas de DADO (fato + dimensão) do Dashboard principal.
-- NÃO inclui (de propósito, ver regra 4 do plano / regra 1-7 do agente
-- supabase-vinci):
--   - Tabelas de usuário/papel/RLS (depende do projeto Supabase existir e de
--     decisão de auth — fica para quando a Fase 1 entrar em implementação de
--     verdade, depois do projeto criado).
--   - Qualquer coluna de PII de candidato de RH (telefone/e-mail/nome de
--     candidato) — o dict de hoje já NUNCA inclui isso (processar_vagas_rh
--     só expõe agregados); replicado aqui como AUSÊNCIA de coluna, não como
--     RLS escondendo.
--   - Qualquer coluna de matrícula/nome de aluno individual — o dict de hoje
--     só expõe contagem agregada de matrículas distintas por polo/categoria
--     (carregar_alunos_hub conta nunique(MATRICULA) e descarta a matrícula
--     em si); replicado aqui como tabela agregada, nunca fato por aluno.
--   - laboratorios.ricos / laboratorios.simples (indicadores aninhados de
--     forma irregular, vindos de um JSON pré-processado fora do
--     processar.py, estrutura ainda não confirmada em detalhe com o Leo) —
--     deixados fora deste schema de propósito; só laboratorios_pendencias
--     (estrutura fixa e já bem conhecida) entra aqui. Ver nota no final do
--     arquivo.
--   - turnover: não vira tabela própria — é 100% derivável de
--     tutores.data_contratacao / tutores.data_desligamento via VIEW (ver
--     final do arquivo), para não duplicar fato.
--
-- Convenção: chaves primárias `id bigint generated always as identity` (ou
-- `uuid` só onde fizer sentido ter ID opaco) para não travar com decisão de
-- auth ainda não tomada. snake_case em tudo.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1. DIMENSÕES
-- ----------------------------------------------------------------------------

-- Polo (LAP). Única fonte de verdade de normalização de nome de polo — hoje
-- o processar.py reimplementa essa normalização em ~6 lugares diferentes
-- (_norm_polo_bf, _norm_polo_ger, _norm_polo_inj, _norm_polo_cruzamento,
-- _norm_polo_contato, _norm_polo_labs — ver MAPA_DO_CODIGO.md §7/Armadilhas,
-- PATCH 149) e isso já causou contagens de polo divergentes entre visões
-- (235 em polo_stats vs 251 em ger_polo, medido na mesma rodada). O schema
-- relacional resolve isso na raiz: 1 tabela, 1 normalização, todo o resto
-- referencia por FK.
create table polos (
    id              bigint generated always as identity primary key,
    nome            text not null unique,        -- nome de exibição, ex: "Abaetetuba/PA"
    nome_normalizado text not null unique,         -- chave de match (sem "LAP -", sem acento, lowercase)
    uf              text,                          -- extraído do nome (ex: "PA")
    created_at      timestamptz not null default now(),
    updated_at      timestamptz not null default now()
);
create index idx_polos_normalizado on polos (nome_normalizado);

-- Categoria ampla (CAT_MAP) — as 6 categorias usadas em cat_stats /
-- pratica_stats / alunos_por_curso. Ex: "Multidisciplinar II - Enfermagem e
-- Instrumentação Cirúrgica".
create table categorias (
    id              bigint generated always as identity primary key,
    codigo_interno  text not null unique,   -- rótulo curto interno (chave do CAT_MAP), ex: 'ENF-INS (Multidisciplinar II)'
    nome_exibicao   text not null,          -- nome longo de exibição
    grupo_amplo     text,                   -- um dos 5 GRUPOS_GER usados no frontend (Gerenciamento → Detalhe de Ofertas)
    created_at      timestamptz not null default now()
);

-- Curso específico (código fino, ex: BFI, BTO, COS-TIP, GPI, ECE...) — vem
-- da coluna "curso específico" do CONTROLE e da Lotação. N:1 com categorias
-- (vários cursos finos caem na mesma categoria ampla, ex: Multi III agrega
-- Fisioterapia/T.Ocupacional/Estética).
create table cursos (
    id              bigint generated always as identity primary key,
    codigo          text not null unique,   -- ex: 'BFI'
    nome            text,
    categoria_id    bigint references categorias(id),
    created_at      timestamptz not null default now()
);

-- Catálogo oficial de práticas previstas (catalogo_oficial.json) — 303
-- nomes distintos medidos, distribuídos entre as 6 categorias.
create table praticas_catalogo (
    id              bigint generated always as identity primary key,
    nome            text not null,
    categoria_id    bigint not null references categorias(id),
    unique (nome, categoria_id)
);

-- Config de prazos/períodos por Ordem e por semestre (hoje em
-- config_semestre.json, editado à mão pelo Leo pra virar o semestre). Vira
-- tabela editável em vez de arquivo — é o insumo que falta pras views de
-- "atrasado"/"urgente"/"status_ordem" funcionarem via SQL puro (ver
-- MAPEAMENTO_SCHEMA.md, kpis.atrasados / kpis.urgentes).
create table config_semestres_ordens (
    semestre        text not null,           -- ex: '2026/2'
    ordem           smallint not null,       -- 1..5 (ger_ordem tem 5 linhas hoje)
    prazo           date,
    periodo_inicio  date,
    periodo_fim     date,
    disciplina      text,
    primary key (semestre, ordem)
);

-- ----------------------------------------------------------------------------
-- 2. TUTORES (dimensão com atributos de RH — é "fato" no sentido de que
--    muda ao longo do tempo via contratação/desligamento, mas é tratado
--    como dimensão porque ofertas/submissões referenciam ele por FK)
-- ----------------------------------------------------------------------------

-- 626 linhas medidas (621 ativos + 5 desligados na rodada 02/10/2026).
-- IMPORTANTE (LGPD): nome/e-mail/whatsapp/chapa JÁ estão no dict de hoje
-- (usados na "Ficha do Tutor" operacional) — não são PII excluída por
-- design, ao contrário de aluno/candidato RH. Ficam aqui como colunas
-- normais; o controle de quem pode LER essas colunas é trabalho de RLS,
-- fora do escopo desta Fase 0 (ver plano §2/§3 Fase 1).
create table tutores (
    id                      bigint generated always as identity primary key,
    nome                    text not null,
    email                   text,
    chapa                   text,                 -- variante com chapa é preferida (PATCH 26)
    whatsapp                text,
    polo_id                 bigint references polos(id),
    categoria_id            bigint references categorias(id),
    curso_id                bigint references cursos(id),   -- curso específico do CONTROLE
    ch_semanal              numeric(5,2),         -- _parse_ch: HH:MM ou decimal → horas
    ch_contratada           numeric(5,2),         -- vem da Lotação (enriquecer_tutores)
    ch_ideal                numeric(5,2),
    data_contratacao        date,                 -- _interpretar_data_contratacao (BR/US tolerante)
    data_desligamento       date,                 -- null se ativo
    ativo                   boolean not null default true,
    titulacao               text,
    graduacao               text,
    especializacao          text,
    mestrado                text,
    doutorado               text,
    lattes_url              text,
    lattes_id               text,
    exp_fora_meses          integer,
    exp_tutor_uni_meses     integer,
    lab_pendencia           boolean default false,  -- flag "lab em obras" (PATCH 178)
    onboarding_trilha       boolean,
    onboarding_checklist    boolean,
    onboarding_um_a_um      boolean,
    created_at              timestamptz not null default now(),
    updated_at              timestamptz not null default now()
);
create index idx_tutores_polo on tutores (polo_id);
create index idx_tutores_categoria on tutores (categoria_id);
create index idx_tutores_ativo on tutores (ativo);
create index idx_tutores_data_contratacao on tutores (data_contratacao);
create index idx_tutores_data_desligamento on tutores (data_desligamento);

-- Contatos do polo (hoje em contatos_por_polo.json) — N:1 com polos.
create table contatos_polo (
    id          bigint generated always as identity primary key,
    polo_id     bigint not null references polos(id),
    nome        text,
    cargo       text,
    telefone    text,
    email       text
);
create index idx_contatos_polo on contatos_polo (polo_id);

-- ----------------------------------------------------------------------------
-- 3. SUBMISSÕES DE PORTFÓLIO (fato)
-- ----------------------------------------------------------------------------
-- 5.746 linhas medidas (snapshot atual, só tutores ativos — cresce a cada
-- semestre). Cada linha = 1 "rodada" de prática reportada (até 7 rodadas por
-- submissão do formulário 2026/2 viram 1 linha cada — PATCH 151, nunca
-- cópia redundante).
create table submissoes_portfolio (
    id                  bigint generated always as identity primary key,
    tutor_id            bigint not null references tutores(id),
    pratica_catalogo_id bigint references praticas_catalogo(id),  -- null se não casou com o catálogo oficial
    pratica_nome_bruto  text not null,           -- texto como veio do formulário, sempre preservado
    semestre_origem     text not null,           -- '2026/1' / '2026/2' — da ORIGEM DO ARQUIVO, não da data (PATCH 127)
    ordem               smallint,
    data_submissao      date,
    alunos_informados   integer,
    created_at          timestamptz not null default now()
);
create index idx_submissoes_tutor on submissoes_portfolio (tutor_id);
create index idx_submissoes_semestre on submissoes_portfolio (semestre_origem);
create index idx_submissoes_ordem on submissoes_portfolio (semestre_origem, ordem);

-- ----------------------------------------------------------------------------
-- 4. OFERTAS DE GERENCIAMENTO — GIOCONDA (fato principal, maior tabela)
-- ----------------------------------------------------------------------------
-- 39.327 linhas medidas somando os 2 semestres disponíveis hoje (24.531 em
-- 2026/2 + 14.796 em 2026/1) — cresce ~15-25k por semestre novo. Por trás de
-- TODAS as visões pré-agregadas ger_kpis/ger_polo/ger_cat/ger_ordem/
-- ger_agendas/ger_contratacao de hoje (que viram VIEWs, ver final do
-- arquivo) está esta mesma tabela de fato.
create table ofertas_gerenciamento (
    id                      bigint generated always as identity primary key,
    semestre                text not null,
    ordem                   smallint,
    polo_id                 bigint references polos(id),
    categoria_id            bigint references categorias(id),
    curso_id                bigint references cursos(id),       -- curso (campo 'curso' do dict)
    subcurso_id             bigint references cursos(id),       -- subcurso (ex: Multi III → Fisio/TO/Estética)
    tutor_id                bigint references tutores(id),      -- null = oferta sem tutor casado
    pratica_nome            text,
    tem_tutor               boolean not null default false,
    tem_agenda              boolean not null default false,
    gerenciado              boolean not null default false,
    alunos_matriculados     integer,
    alunos_agendados        integer,
    data_agenda             date,
    hora_agenda             time,
    dia_semana              text,
    turno                   text,
    horario_incomum         boolean default false,
    sem_alunos              boolean default false,
    anomalia_ordem          boolean default false,   -- _anomalia_ordem (PATCH 97)
    anomalia_ordem_futura   boolean default false,   -- _anomalia_ordem_futura (PATCH 130)
    created_at              timestamptz not null default now()
);
create index idx_ofertas_semestre on ofertas_gerenciamento (semestre);
create index idx_ofertas_polo on ofertas_gerenciamento (polo_id);
create index idx_ofertas_categoria on ofertas_gerenciamento (categoria_id);
create index idx_ofertas_tutor on ofertas_gerenciamento (tutor_id);
create index idx_ofertas_ordem on ofertas_gerenciamento (semestre, ordem);
create index idx_ofertas_gerenciado on ofertas_gerenciamento (semestre, gerenciado);

-- ----------------------------------------------------------------------------
-- 5. MATRÍCULAS DISTINTAS (fato agregado — NUNCA aluno individual)
-- ----------------------------------------------------------------------------
-- A fonte (Relatorio_alunos_por_hub.csv) tem MATRICULA por linha, mas
-- carregar_alunos_hub() só usa isso pra fazer nunique() e descarta a
-- matrícula em si — o dict de hoje nunca expõe o aluno nominal, só a
-- contagem. Replicado aqui com a MESMA granularidade do dict (por polo x
-- categoria x semestre), nunca por aluno — ver regra 5 do agente
-- supabase-vinci ("não existe PII onde não precisa existir" > RLS escondendo).
create table matriculas_agregadas_hub (
    id                          bigint generated always as identity primary key,
    semestre                    text not null,
    polo_id                     bigint references polos(id),
    categoria_id                bigint references categorias(id),
    total_matriculas_distintas  integer not null,
    unique (semestre, polo_id, categoria_id)
);
create index idx_matriculas_semestre on matriculas_agregadas_hub (semestre);

-- ----------------------------------------------------------------------------
-- 6. VAGAS RH (fato, origem Lotação — processar_vagas, NÃO o funil de
--    candidatos de processar_vagas_rh, que já é agregado por natureza e nem
--    apareceu nesta medição — não presente no dict desta rodada)
-- ----------------------------------------------------------------------------
-- 399 linhas medidas. Sem PII de candidato (candidato é outra fonte, p7,
-- que já só entra como agregado quando presente — fora do escopo desta
-- tabela).
create table vagas_rh (
    id                  bigint generated always as identity primary key,
    polo_id             bigint references polos(id),
    curso_id            bigint references cursos(id),
    perfil              text,
    status              text,              -- 'Aumento de Quadro' / 'Substituição'
    contratacao_liberada boolean,
    tutor_atual         text,              -- nome do tutor atual na vaga (substituição) — não é PII de terceiro externo
    chamado_sydle       text,
    status_chamado      text,
    ch_semanal          numeric(5,2),
    ch_ideal            numeric(5,2),
    prioridade          text,
    autorizado          text,
    alunos_polo         integer,
    created_at          timestamptz not null default now()
);
create index idx_vagas_rh_polo on vagas_rh (polo_id);

-- ----------------------------------------------------------------------------
-- 7. LABORATÓRIOS — PENDÊNCIAS (fato curado externamente, estrutura fixa
--    e já conhecida — labs_pendencias.json)
-- ----------------------------------------------------------------------------
-- 51 linhas medidas. NÃO inclui laboratorios.ricos / laboratorios.simples
-- (estrutura de indicadores aninhada e variável por categoria — ex:
-- "indicadores_gerais": [{k, v}, ...] — normalizar isso direito exige uma
-- conversa com o Leo sobre o que cada indicador realmente significa e se
-- todos são genuinamente necessários na Fase 1; deixado fora de propósito
-- pra não supor estrutura, ver regra 1 do processar-etl-vinci/orientação
-- geral do projeto de nunca resolver ambiguidade por suposição).
create table laboratorios_pendencias (
    id                  bigint generated always as identity primary key,
    polo_id             bigint references polos(id),
    categoria           text,
    status              text,       -- 'Apto' / 'Não Apto'
    motivo              text,
    empresa              text,
    total_alunos        integer,
    tutor_contratado     boolean,
    created_at          timestamptz not null default now()
);
create index idx_lab_pendencias_polo on laboratorios_pendencias (polo_id);

-- ============================================================================
-- VIEWS — substituem as visões pré-agregadas que hoje são dict pré-calculado
-- em Python. Ilustrativas (ver MAPEAMENTO_SCHEMA.md para o mapeamento campo a
-- campo completo) — ajustar quando o projeto Supabase existir e puder-se
-- testar de verdade contra o schema.
-- ============================================================================

-- kpis.total / kpis.enviaram / kpis.pendentes (sem o que depende de
-- prazo/ordem — ver MAPEAMENTO_SCHEMA.md sobre atrasados/urgentes)
create view v_kpis_tutores as
select
    count(*) filter (where ativo) as total,
    count(*) filter (where ativo and id in (select tutor_id from submissoes_portfolio)) as enviaram,
    count(*) filter (where ativo and id not in (select tutor_id from submissoes_portfolio)) as pendentes
from tutores;

-- ger_kpis por semestre (equivalente a dict['gerenciamento_por_semestre'][sem]['ger_kpis'])
create view v_ger_kpis_por_semestre as
select
    semestre,
    count(*) as total_ofertas,
    count(*) filter (where gerenciado) as ofertas_gerenciadas,
    count(*) filter (where not gerenciado) as ofertas_nao_gerenciadas,
    round(100.0 * count(*) filter (where gerenciado) / nullif(count(*), 0), 1) as pct_gerenciado,
    count(*) filter (where tem_tutor) as ofertas_com_tutor,
    count(*) filter (where not tem_tutor) as ofertas_sem_tutor,
    round(100.0 * count(*) filter (where tem_tutor) / nullif(count(*), 0), 1) as pct_com_tutor,
    count(*) filter (where tem_agenda) as ofertas_com_agenda,
    sum(alunos_matriculados) as total_alunos_matriculados,
    sum(alunos_agendados) as total_alunos_agendados,
    count(distinct polo_id) as polos_total,
    count(distinct polo_id) filter (where not tem_tutor) as polos_sem_tutor
from ofertas_gerenciamento
group by semestre;

-- ger_polo por semestre
create view v_ger_polo_por_semestre as
select
    o.semestre,
    p.id as polo_id,
    p.nome as polo_nome,
    count(*) as total,
    count(*) filter (where o.gerenciado) as gerenciadas,
    count(*) filter (where o.tem_agenda) as com_agenda
from ofertas_gerenciamento o
join polos p on p.id = o.polo_id
group by o.semestre, p.id, p.nome;

-- ger_cat por semestre
create view v_ger_cat_por_semestre as
select
    o.semestre,
    c.id as categoria_id,
    c.nome_exibicao,
    count(*) as total,
    count(*) filter (where o.gerenciado) as gerenciadas
from ofertas_gerenciamento o
join categorias c on c.id = o.categoria_id
group by o.semestre, c.id, c.nome_exibicao;

-- ger_ordem por semestre
create view v_ger_ordem_por_semestre as
select
    semestre,
    ordem,
    count(*) as total,
    count(*) filter (where gerenciado) as gerenciadas
from ofertas_gerenciamento
where ordem is not null
group by semestre, ordem;

-- turnover: 100% derivável de tutores, sem precisar de tabela própria
-- (equivalente a dict['turnover'] — janelas semana/mês/semestre calculadas
-- em cima de data_contratacao/data_desligamento; o corte exato de "semana
-- atual"/"mês atual"/"semestre ativo" deve ser parametrizado na query, não
-- fixo na view, já que "hoje" muda todo dia).
create view v_turnover_base as
select id, nome, polo_id, categoria_id, data_contratacao, data_desligamento
from tutores;

-- ============================================================================
-- NOTA FINAL: tabelas de usuário/papel/RLS ficam para depois do projeto
-- Supabase existir (ver TAREFA 3 / COMO_CRIAR_PROJETO.md) — este arquivo
-- cobre só dado, como pedido.
-- ============================================================================
