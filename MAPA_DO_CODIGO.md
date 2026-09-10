# Mapa do Código — VinciLab

Última atualização: 2026-09-09
Commit de referência: `341316c` (Adiciona 10 subagentes especializados do VinciLab)
+ mudanças locais não commitadas: **PATCH 157** (Recrutamento & Seleção na seção Vagas).

> Primeiro arquivo a ler antes de mexer no VinciLab. Responde "onde fica X" e
> "por que Y foi feito assim" sem precisar carregar `processar.py` (~4.600 linhas)
> nem os dois templates (~5.100 linhas cada) no contexto.
>
> **Regra de ouro do projeto:** quase toda mudança de frontend precisa ser
> replicada nos DOIS templates (`template_dashboard.html` e
> `template_coordenadores.html`). Esquecer um dos dois é a fonte nº 1 de bugs aqui.

---

## Índice

1. [Visão geral do sistema](#1-visão-geral-do-sistema)
2. [Estrutura de arquivos](#2-estrutura-de-arquivos)
3. [Planilhas de entrada](#3-planilhas-de-entrada)
4. [Arquivos JSON do repositório](#4-arquivos-json-do-repositório)
5. [processar.py — fluxo geral](#5-processarpy--fluxo-geral)
6. [processar.py — índice de funções](#6-processarpy--índice-de-funções)
7. [Regras de negócio não óbvias](#7-regras-de-negócio-não-óbvias)
8. [Os dois portais (templates HTML/JS)](#8-os-dois-portais-templates-htmljs)
9. [template_dashboard.html — índice de funções JS](#9-template_dashboardhtml--índice-de-funções-js)
10. [template_coordenadores.html — o que muda](#10-template_coordenadoreshtml--o-que-muda)
11. [Build e deploy (GitHub Actions)](#11-build-e-deploy-github-actions)
12. [Glossário de convenções](#12-glossário-de-convenções)
13. [Histórico de PATCHes (resumido)](#13-histórico-de-patches-resumido)
14. [Armadilhas conhecidas](#14-armadilhas-conhecidas)

---

## 1. Visão geral do sistema

VinciLab é um dashboard de gestão de práticas/tutoria (laboratórios da saúde e
engenharias) da UNIASSELVI. O pipeline:

```
Planilhas SharePoint/OneDrive  →  processar.py  →  saida/dashboard.html
      (CONTROLE, PORTIFOLIO,          (Python/pandas)      saida/coordenadores.html
       GIOCONDA, LOTAÇÃO, hub CSV)                         saida/lookup.json
                                                                │
                                              GitHub Actions copia p/ raiz:
                                              index.html, coordenadores.html, lookup.json
                                                                │
                                                    GitHub Pages (estático)
```

- **Backend:** `processar.py` — lê as planilhas, cruza tudo por polo/tutor/
  categoria/prática, calcula agregados e KPIs, e injeta um JSON **cifrado
  (AES-256-GCM)** dentro dos templates HTML.
- **Frontend:** dois HTML de página única, com todo o JS embutido. O JSON
  cifrado é decifrado no browser com a senha `uniasselvi2026`
  (`SENHA_DASHBOARD`, `processar.py:3905`).
- **Deploy:** GitHub Actions roda de 2 em 2 horas (cron), baixa as planilhas via
  secrets, roda `processar.py --sem-browser`, e commita `index.html` /
  `coordenadores.html` / `lookup.json`.
- **Site publicado:** `https://leoschaeffner.github.io/uniasselvi-dashboard/`

---

## 2. Estrutura de arquivos

| Arquivo | Papel |
|---|---|
| `processar.py` | **Todo o ETL + geração de saída.** Ponto de entrada `__main__` (~linha 3986). |
| `template_dashboard.html` | Template do portal principal (VinciLab). Placeholder `'DATA_GOES_HERE'` e `TIMESTAMP_GOES_HERE`. |
| `template_coordenadores.html` | Template do portal de coordenadores (versão travada por curso, simplificada). Placeholder `'DATA_GOES_HERE'`. |
| `portfolio_form.html` | Formulário público de envio de portfólio; autopreenche via `lookup.json` e redireciona pra uma lista do SharePoint (não passa pelo `processar.py`). |
| `index.html` / `coordenadores.html` / `lookup.json` | **Saída gerada.** Desde o PATCH 160 **NÃO são mais commitados** — o workflow monta `_site/` e publica via artifact do GitHub Pages. Estão no `.gitignore`. NÃO editar à mão. |
| `config_semestre.json` | Config editável de prazos/períodos das Ordens por semestre. Única coisa que se edita pra virar o semestre. |
| `catalogo_oficial.json` | Catálogo oficial de práticas (formulário de portfólio 2026/2). |
| `categoria_para_curso.json` | Texto do formulário de portfólio → código fino de curso (ex: `"Multidisciplinar III - Fisioterapia" → "BFI"`). |
| `categoria_para_curso.json` / `id_to_perfil.json` / `nome_to_perfil.json` | Mapas de reconciliação prática↔curso (ver [Glossário](#12-glossário-de-convenções)). |
| `mec_cache.json` | Cache de dados MEC/perfil por e-mail do tutor. |
| `laboratorios_data.json` | Dados pré-processados da seção "Laboratórios" (gerados fora do `processar.py`). |
| `labs_pendencias.json` | Laboratórios em obras / não aptos (ícone de "obras" nos tutores). |
| `contatos_por_polo.json` | Lista Oficial de Contatos por polo (anexada à ficha do tutor). |
| `tutores_comunicacao.json` | Relatório manual de acompanhamento de comunicação com tutores. |
| `insumos_estudo.json` | Estudo estático "Racional de Insumos por Experimento" (v10_FINAL). **Não é auto-atualizado** — substituir à mão se vier versão nova. NÃO está commitado no momento; a aba fica vazia se ausente. |
| `config_links.json` | (Local, não commitado) caminhos absolutos das planilhas pra rodar local sem baixar. |
| `snapshot_manifest.json` | Snapshot da rodada anterior (colunas detectadas + contagens) pra detecção de regressão. Gerado/lido pelo próprio `processar.py`. |
| `.github/workflows/publicar.yml` | Pipeline de build/deploy. |
| `.claude/agents/*.md` | 10 subagentes especializados (este arquivo é mantido pelo `guia-codigo`). |
| `robots.txt`, `files (1).zip` | Irrelevantes pro código. |

---

## 3. Planilhas de entrada

Baixadas pelo Actions via secrets (`URL_*`) para `planilhas/`. Localizadas por
`achar_arquivo()` / `verificar_e_localizar()` (`processar.py:363`). Ordem de
retorno: `p1, p2, tmpl, p3, p3b, p4, p5, p6, p7` (PATCH 157 acrescentou `p7`).

| Var | Arquivo | Secret | Obrigatória | Conteúdo |
|---|---|---|---|---|
| `p1` | `01_CONTROLE_TUTORIA.xlsx` | `URL_CONTROLE` | **Sim** | **CONTROLE** — cadastro-mestre de tutores ativos: nome, e-mail, polo, categoria, curso específico, CH SEMANAL, data de contratação, desligamento. |
| `p2` | `PORTIFOLIO_TUTOR.xlsx` (com "I") | `URL_PORTFOLIO` | **Sim** | **Portfólio 2026/1** — submissões de portfólio (uma linha por envio), com protocolos de prática, data, nº de alunos. |
| `p2b` | `PORTIFOLIO_TUTOR_2026_2.xlsx` | `URL_PORTFOLIO_2026_2` | Não | **Portfólio 2026/2** — formulário customizado, schema próprio, com múltiplas "rodadas" por linha (ver PATCH 146/151). |
| `p3` | `REL_GERAL_DE_GERENCIAMENTO.xlsx` | `URL_GERENCIAL` | Não | **GIOCONDA** (gerenciamento de ofertas) — formato antigo (.xlsx). |
| `p3b` | `REL_GERAL_DE_GERENCIAMENTO_26_02.csv` | `URL_GERENCIAL_26_02` | Não | **GIOCONDA** — export novo (CSV), tem coluna `SEMESTRE` por linha. |
| `p4` | `LOTACAO_TUTORES.xlsm` / `.xlsx` | `URL_LOTACAO` | Não | **Lotação** — CH contratada por tutor, cursos do polo, total de alunos por polo. Base de RH/Vagas e "alunos por curso" (fallback). |
| `p5` | `Relatorio_alunos_por_hub.csv` | `URL_ALUNOS_HUB` | Não | **Hub de alunos** — matrículas; usado pra contar alunos DISTINTOS (substitui a contagem inflada do GIOCONDA). Dois esquemas de coluna auto-detectados (PATCH 43). |
| `p6` | `Acompanhamento_Onboarding.xlsx` (aba `Onboarding Tutores`) | `URL_ONBOARDING_TUTORES` (secret ainda não configurado; download nem está no workflow ainda) | Não | **Onboarding** — flags Trilha/Checklist/1:1 por tutor. O `processar.py` também **gera/atualiza** essa planilha localmente (`gerar_onboarding_atualizado`, roda sempre). |
| `p7` | `VAGAS_RH.xlsx` | `URL_VAGAS_RH` | Não | **Recrutamento & Seleção (RH)** — 4 abas: `Status- Anotações` (legenda de status + tabela salário×CH), `Agendamento de entrevista` (1 linha por candidato — só o **agregado** entra no JSON, sem nome/telefone/e-mail), `Aumento` e `Substituição` (vagas no INHIRE: ticket, link, "Vaga Fechada", Etapa). PATCH 157. |
| — | `REL_DETALHADO.csv` | `URL_REL_NOVO` | Não | Baixado pelo Actions mas **não consumido** por nenhuma função ativa (ver Armadilhas). |

---

## 4. Arquivos JSON do repositório

Carregados por `processar.py` (a maioria dentro de `processar()`):

- `config_semestre.json` (`_carregar_semestres`, ~L55) → `ALL_SEMESTRES`,
  `SEMESTRE_ATUAL` (= o mais recente por ordem alfabética das chaves), `PRAZOS_ORDENS`, `PERIODOS_ORDENS`.
- `catalogo_oficial.json` (~L785) — catálogo de práticas do formulário.
- `id_to_perfil.json` (~L798) — id da prática → código de perfil (ex: `'206' → 'EMF-ISN'`).
- `nome_to_perfil.json` (~L811, também ~L3465) — nome da prática → perfil correto
  (só pra códigos ambíguos BFI/BTO/COS-TIP). Também usado pra separar catálogo por curso.
- `categoria_para_curso.json` (~L858, ~L4237) — texto do formulário → código fino.
- `mec_cache.json` (~L777).
- `labs_pendencias.json` (~L1782), `tutores_comunicacao.json` (~L1822),
  `contatos_por_polo.json` (~L1878), `laboratorios_data.json` (~L2127).
- `insumos_estudo.json` (~L4082) — opcional.
- `snapshot_manifest.json` (~L325) — regressão.

---

## 5. processar.py — fluxo geral

Ponto de entrada: `if __name__ == '__main__'` (~L3986).

```
1. verificar_e_localizar()                → resolve p1..p6 + template
2. dados = processar(p1, p2)              → NÚCLEO: CONTROLE + Portfólio (ambos semestres)
                                             → tutores[], polo_stats, kpis, cat_stats,
                                               pratica_stats, catalogo, por_ordem, por_mes,
                                               dados_por_semestre, portfolio_alunos_dedup*
3. se p4:  lotacao = carregar_lotacao(p4)
           dados = enriquecer_tutores(dados, lotacao)   → CH contratada, alunos_por_curso,
                                                           laboratórios agregados
4. se p6:  lê "Onboarding Tutores" e aplica flags nos tutores (por chapa, senão nome+polo)
5. se p4:  dados['vagas'] = processar_vagas(p4)          → posições em aberto (RH/Vagas)
6.         gerar_onboarding_atualizado(p1, p6, ...)      → (re)gera Acompanhamento_Onboarding.xlsx
7.         dados['tem_lotacao'] = (CH>0 em ≥1 tutor)
8.         dados['insumos'] = insumos_estudo.json (ou None)
9. se p3 ou p3b:  bloco GIOCONDA (grande, ~L4090-4345)
      a. monta controle_tutor_lookup {(polo_norm, cat|curso): nome}  (backfill de tutor)
      b. ger_por_semestre = processar_gerenciamento_semestres([(p3, sem_antigo), (p3b, '2026/2')], lookup)
      c. PATCH 118: remove do GIOCONDA linhas de tutores sem vínculo ativo / "Aviso de Portfólio"
      d. _injetar_tutores_sem_oferta(): garante que todo tutor ativo aparece no Detalhe
      e. _detectar_gerenciamento_fora_ordem(): sinaliza anomalias de ordem
      f. dados['gerenciamento_por_semestre'] = ger_por_semestre
      g. dados.update(ger_dados do SEMESTRE_ATUAL) + tem_gerenciamento=True
      h. cruzamento Portfólio×Agendado por semestre (_cruza) → cruzamento_portfolio_agendado
10. se p5: alunos_hub = carregar_alunos_hub(p5)
      → substitui KPI "Alunos" e "alunos_por_curso" por matrículas DISTINTAS
11. html = gerar_html(dados)                 → saida/dashboard.html + saida/lookup.json
12. gerar_html_coordenadores(dados)          → saida/coordenadores.html
13. (sem --sem-browser) abre o navegador; (WATCH_MODE) entra em modo_watch
```

`processar()` internamente (ordem aproximada):
1. Lê CONTROLE → `df_at` (ativos) + `tutores_desligados` (PATCH 105).
2. Lê Portfólio 2026/1 (`df_p`) e 2026/2 (`df_p` estendido, schema próprio) →
   marca `_SEMESTRE_ORIGEM` pela **origem do arquivo**, não pela data (PATCH 127).
3. `soma_estudantes()` — soma alunos de TODAS as colunas de "rodada" (PATCH 122/151).
4. `_calc_portfolio_dedup()` — alunos registrados no portfólio, deduplicados por
   polo×categoria (PATCH 123, mesmo problema do PATCH 38).
5. Para cada tutor ativo: casa submissões por `_encontrar_por_polo` (nome/e-mail/
   subsequência), monta `hist[]`, calcula situação por Ordem contra `PERIODOS_ORDENS`.
6. `_catalogo_por_curso` — filtra o catálogo de "práticas previstas" do tutor pelo curso específico.
7. `_stats_semestre` — recalcula KPIs/pratica_stats/prazos por semestre → `dados_por_semestre`.
8. `_verificar_snapshot_regressao` — compara com a rodada anterior (só log).
9. `return limpar({...})` — dict gigante (ver `processar.py:2094`).

---

## 6. processar.py — índice de funções

### Nível de módulo / helpers globais

| Função | Linha ~ | O que faz |
|---|---|---|
| `_carregar_semestres()` | 55 | Lê `config_semestre.json`, monta `ALL_SEMESTRES` (dict de prazos/períodos por semestre) + disciplinas por ordem. |
| `_data_para_semestre(data_str)` | 84 | Data → chave de semestre (`'2026/1'` etc.), pelos intervalos de `ALL_SEMESTRES`. |
| `_ordem_relativa(ordem_forms, semestre)` | 100 | Ajusta o número da Ordem conforme o semestre. |
| `_parse_ch(v)` | 115 | CH SEMANAL (`HH:MM` ou decimal) → float horas (PATCH 1). |
| `_dia_semana_pt(iso_str)` | 136 | ISO → dia da semana em PT (Análise de Agendas, PATCH 30). |
| `_turno_de_horario(hr_str)` | 146 | Hora → Manhã/Tarde/Noite (PATCH 30). |
| `achar_pasta_script()` | 163 | Diretório do script → `SCRIPT_DIR`. |
| `ler_url_file(path_url)` | 186 | Lê `.url` do Windows e extrai a URL. |
| `forcar_download_onedrive(path_url_file, destino, label)` | 195 | Baixa planilha do OneDrive/SharePoint a partir de um `.url`. |
| `_bate(caminho_arq, padrao)` / `achar_arquivo(pasta, padrao)` | 226 / 232 | Match de nome de arquivo (exato → aproximado → OneDrive local). |
| `_normaliza_categoria_bio_duplicado(s)` | 294 | Corrige rótulo com prefixo "BIO-" duplicado na FONTE (PATCH 82). |
| `ts()` | 303 | Timestamp pros logs. |
| `limpar(obj)` | 308 | Sanitiza dict/list pra JSON (NaN→None, numpy→nativo). |
| `_verificar_snapshot_regressao(colunas, contagens)` | 324 | Compara colunas detectadas + contagens com `snapshot_manifest.json` (só log). |
| `verificar_e_localizar()` | 363 | Resolve `p1..p6` + template. Inclui download do hub CSV (`_build_dl_urls`, `_build_dl_urls_onb`). |
| `ler_excel(path, **kwargs)` | 526 | Wrapper de `pd.read_excel` com detecção de engine. |
| `_ler_arquivo_gerenciamento(path)` | 536 | Lê GIOCONDA como `.xlsx` OU `.csv` (PATCH 18). |

### `processar(p1, p2)` — L551 (função-núcleo, ~1550 linhas)

Helpers internos (só existem dentro de `processar()`):

| Helper | Linha ~ | O que faz |
|---|---|---|
| `_interpretar_data_contratacao(valor)` | 558 | Resolve datas BR (DD/MM) vs US (MM/DD) na MESMA coluna: BR primeiro, troca pra US se mês>12 ou data futura (PATCH 155/156). |
| `col(df, *partes)` | 715 | Busca flexível de coluna por pedaços do nome. |
| `soma_estudantes(df)` | 733 | Soma nº de alunos de TODAS as colunas de "rodada"/estudantes (PATCH 122). |
| `_norm_proto(s)` | 805 | Normaliza nome de prática/protocolo pra chave. |
| `_col_rodada_2b(prefixo, n, ...)` / `_norm_espacos(s)` | 868 / 897 | Parsing das 7 rodadas do formulário 2026/2 (PATCH 146/151). |
| `_calc_portfolio_dedup(df_sub)` | 957 | "Alunos registrados no portfólio" dedup por polo×categoria (PATCH 123). |
| `_norm_nome_pratica(s)` | 1043 | Normaliza nome de prática. |
| `_norm_nome_match / _eh_subsequencia_nome_match / _nomes_batem_match` | 1117-1137 | Match de nome de tutor por igualdade / subsequência de palavras / primeiro+último (PATCH 83). |
| `_dividir_codigo_composto(chave)` | 1175 | Divide código de curso composto (ex: `ECE-ENM-...`). |
| `_encontrar_por_polo(chave, cat_subm='')` | 1188 | **Casa uma submissão de portfólio ao tutor certo** (por e-mail do remetente → PATCH 23; senão nome; senão subsequência). Núcleo do cruzamento. |
| `_catalogo_por_curso(praticas_full, cursos_t)` | 1420 | Filtra "práticas previstas" do tutor pelo curso específico via `nome_to_perfil.json` (PATCH 24). Se o mapa não cobrir nada, mantém tudo. |
| `_norm_polo_labs(s)` | 1787 | Normaliza polo pra casar com `laboratorios_data.json`. |
| `_norm_nome(s)` | 1819 | Normaliza nome (genérico). |
| `_norm_polo_contato(s)` | 1874 | Normaliza polo pra `contatos_por_polo.json`. |
| `_stats_semestre(sem_key, tutores_list, catalogo_dict, prazos_dict, periodos_dict)` | 1927 | Recalcula KPIs/pratica_stats/praticas/prazos de UM semestre → `dados_por_semestre[sem]` (PATCH 108). |

### Fora de `processar()`

| Função | Linha ~ | O que faz |
|---|---|---|
| `_carregar_laboratorios()` | 2124 | Carrega `laboratorios_data.json`. |
| `_detectar_e_corrigir_base64(p4)` | 2137 | Detecta Lotação salva em base64 e corrige. |
| `_ler_lotacao_xlsx / _xls / _pandas(p4)` | 2158-2181 | Leitores da planilha de Lotação por formato. |
| `carregar_lotacao(p4)` | 2187 | Lê Lotação → dict `{nome_lower: {cursos, polo_hub, total_alunos, ch...}}`. |
| `processar_vagas(p4)` | 2257 | Extrai posições em aberto (Aumento de Quadro / Substituição) da Lotação → seção RH/Vagas (PATCH 29). `_gv` interno lê célula por índice. |
| `processar_vagas_rh(p7)` | ~2470 | **PATCH 157.** Lê `VAGAS_RH.xlsx` (4 abas, localizador de coluna tolerante). Retorna `{reqs, funil, kpis, ref_salario}`: `reqs` = vagas no INHIRE (Aumento+Substituição); `funil` = candidatos AGREGADOS (`por_status`/`por_mes`/`por_curso`/`por_curso_status`/`por_uf`/`pcd`) **sem nome/telefone/e-mail**; `kpis` = derivados. Helpers de módulo: `_vrh_norm`, `_vrh_cidade_uf`; constantes `_VRH_STATUS_GRUPO`/`_VRH_STATUS_ORDEM`/`_VRH_REF_SALARIO`. |
| `_cruzar_vagas_recrutamento(vagas_lotacao, reqs)` | ~2640 | **PATCH 157.** Anota `rec_etapa`/`rec_ticket`/`rec_link`/`rec_fechada` em cada vaga da Lotação que casa com um `req` do INHIRE (mesmo tipo + curso + cidade/polo normalizados). Conservador: só grava com polo E curso batendo. |
| `gerar_onboarding_atualizado(p1, p6, destino)` | 2344 | (Re)gera `Acompanhamento_Onboarding.xlsx`: preserva flags de quem já é acompanhado, adiciona tutor novo, remove quem virou apto (PATCH 89). `_categoria_exibicao` interno. |
| `enriquecer_tutores(dados, lotacao)` | 2469 | Casa Lotação↔tutores (por nome, com fallback fuzzy `_nomes_batem_ch`): CH contratada, `alunos_por_curso`, alunos por laboratório. `LAB_PARA_CAT` mapeia string de cursos da Lotação → nome de categoria. |
| `processar_gerenciamento_csv(p5)` | 2667 | **DEAD CODE** — nunca chamado (ver Armadilhas). Processava um CSV detalhado de gerenciamento. |
| `_processar_gerenciamento_novo(df_g)` | 2835 | Processa GIOCONDA formato NOVO (LABORATORIO/NOME_EXPERIMENTO). Expande linhas por `_PESO_OFERTA`, deriva `_GERENCIADO`, dia/turno, dedup por polo×categoria antes de somar alunos (PATCH 38). Retorna `ger_kpis/ger_polo/ger_cat/ger_ordem/ger_contratacao/ger_agendas/ger_ofertas`. `gc`, `extrair_ordem_exp`, `_corrige_prefixo_bio_duplicado`, `_extrair_curso`, `to_iso` internos. |
| `_recalcular_agregados_de_ofertas(ofertas)` | 3105 | Recalcula todos os `ger_*` a partir da lista `ger_ofertas` (usado depois de injetar ofertas sintéticas). |
| `_detectar_gerenciamento_fora_ordem(ofertas, periodos)` | 3240 | Marca `_anomalia_ordem` (geriu ordem avançada durante período de ordem anterior — PATCH 97) e `_anomalia_ordem_futura` (ordem ainda não começou — PATCH 130). |
| `_injetar_tutores_sem_oferta(ger_dados, tutores_ativos)` | 3306 | Injeta oferta-placeholder pra todo tutor ativo sem oferta no GIOCONDA (PATCH 32). `_categorias_validas_para` define as categorias reais válidas (PATCH 41). Exclui pseudo-tutores "Aviso de Portfólio" (PATCH 40). |
| `processar_gerenciamento_semestres(arquivos, controle_tutor_lookup=None)` | 3433 | **Orquestrador do GIOCONDA.** Lê 1+ arquivos, separa por coluna `SEMESTRE` quando existe, faz backfill de TUTOR via `controle_tutor_lookup` (PATCH 21/32). Retorna `{semestre: ger_dados}`. `_norm_polo_ger`, `_norm_proto_bf`, `_pratica_de_experimento_bf` internos. |
| `processar_gerenciamento(p3)` | 3550 | Processa GIOCONDA formato ANTIGO (.xlsx). Se detectar schema novo, delega pra `_processar_gerenciamento_novo`. `gcol`, `extrair_ordem` internos. |
| `carregar_alunos_hub(path_csv)` | 3755 | Lê o hub CSV → matrículas DISTINTAS por polo e por categoria. Auto-detecta esquema ANTIGO (`POLO_HUB`) vs NOVO (`POLO`/`CATEGORIA_LABORATORIO`) (PATCH 43). Rejeita HTML (download falho). `_norm`, `_grupo_para_cat`, `_classif_disc`, `_norm_tutor` internos. |
| `cifrar_dados(dados_json_str, senha)` | 3909 | AES-256-GCM (chave = SHA-256 da senha), retorna `"iv_b64:ct_b64"` (PATCH 8). |
| `gerar_html_coordenadores(dados)` | 3918 | Injeta JSON cifrado em `template_coordenadores.html` → `saida/coordenadores.html` (PATCH 110). |
| `gerar_html(dados)` | 3939 | Injeta JSON cifrado + timestamp em `template_dashboard.html` → `saida/dashboard.html`; também escreve `saida/lookup.json` (PATCH 9). |
| `modo_watch(p1, p2)` | 3966 | Modo `watch`: reprocessa a cada 30s quando as planilhas mudam. |

---

## 7. Regras de negócio não óbvias

### Semestre
- `SEMESTRE_ATUAL` = maior chave de `ALL_SEMESTRES` por ordenação de string.
  Virar o semestre = editar `config_semestre.json`, **não** o código.
- **`_SEMESTRE_ORIGEM` do portfólio vem da ORIGEM DO ARQUIVO, não da data da
  submissão** (PATCH 127). Portfólio 2026/1 sempre `'2026/1'`, o arquivo `_2026_2`
  sempre `'2026/2'`.
- Todo cruzamento Portfólio×Agendado é calculado **por semestre separadamente** —
  comparar histórico somado vs. um semestre dava >100% impossível (PATCH 123/127).

### Categorias
- **`CAT_MAP`** (`processar.py:267`): rótulo curto/interno → nome de exibição longo.
  Ex: `'BIO-FISIO-EST-TO (Multidisciplinar III)'` → `'Multidisciplinar III - ...'`.
- Rótulo com prefixo "BIO-" duplicado (`BIO-BIO-FISIO-...`) é normalizado **na
  fonte** por `_normaliza_categoria_bio_duplicado` (PATCH 82) — antes virava
  "categoria fantasma" em qualquer agrupamento por valor cru.
- **Reconciliação de vocabulário** entre fontes: o Portfólio guarda o texto do
  formulário (`"Multidisciplinar III - Fisioterapia"`), o GIOCONDA guarda o
  rótulo amplo do `CAT_MAP`. Ponte: `categoria_para_curso.json` (texto → código
  fino, ex `BFI`) + `_FINO_PARA_AMPLO` (`processar.py:4230`, código fino → rótulo amplo).
- Filtros de Gerenciamento no frontend usam **5 grupos de área** (`GRUPOS_GER`),
  não as categorias cruas (PATCH 47). Ver [Glossário](#12-glossário-de-convenções).

### Deduplicação de alunos
- **GIOCONDA infla contagem de alunos**: o mesmo aluno aparece 1x por prática. A
  regra correta é dedup por polo×categoria dentro de uma ordem **antes** de somar
  (PATCH 38 no backend, PATCH 48/114 no frontend).
- Quando o hub CSV (`p5`) está disponível, ele é a **fonte de verdade** de
  "alunos matriculados" (matrículas distintas) e substitui a contagem do GIOCONDA
  e da Lotação (PATCH 131/135/136).
- `_OFE` × `_PESO_OFERTA`: uma linha do GIOCONDA representa N vagas; expande-se a
  linha por `_PESO_OFERTA` (= `_OFE` clip≥1), **mas deduplicando por prática
  primeiro** senão multiplica duplo (PATCH 150).

### Nomes de tutor
- Match em cascata, sempre nesta lógica: nome exato normalizado → primeiro+último
  nome → subsequência de palavras completas (nunca substring — PATCH 98).
  Reimplementado localmente em MUITOS lugares (`_nomes_batem_*`) — ver Armadilhas.
- Variante do nome COM chapa `(12345)` é preferida (mais rastreável — PATCH 26).
- Datas de contratação misturam BR e US na mesma coluna
  (`_interpretar_data_contratacao`, PATCH 155/156).

### Polos
- Normalização de polo: remove prefixo `LAP -`, remove parênteses, remove
  acentos, colapsa espaços, lowercase. Reimplementada localmente em vários pontos
  (`_norm_polo_bf`, `_norm_polo_ger`, `_norm_polo_inj`, `_norm_polo_cruzamento`,
  `_norm_polo_contato`, `_norm_polo_labs`, `normPoloExp` no JS...). **Não são a
  mesma função** — ver Armadilhas (PATCH 149).
- UF extraída do polo por regex `/\/([A-Z]{2})(?:\s*-|$)/` (frontend).

### Tutores
- **Ativos** = linha ativa no CONTROLE. **Desligados** = lista separada
  (`tutores_desligados`, PATCH 105), com data de desligamento.
- Um tutor desligado que reapareceu ativo NÃO é excluído do gerenciamento
  (casos reais citados no PATCH 118).
- Práticas enviadas ANTES da data de admissão do tutor atual não contam
  (PATCH 16).
- **"Em treinamento"**: ≤60 dias de casa **E** ainda não gerenciou nada
  (PATCH 134/137).
- **"Aviso de Portfólio"**: pseudo-tutores criados de submissões que não casaram
  com nenhum tutor real — excluídos de KPIs, gerenciamento e `lookup.json`.

---

## 8. Os dois portais (templates HTML/JS)

Ambos são página única, JS inline, dados via `DB` (JSON decifrado). Estrutura de
seções quase idêntica; a lógica de render é **replicada** (mesmos nomes de
função) — mudança no dashboard geralmente tem que ir no coordenadores também.

### Seções / navegação

| Seção (`navTo`) | Dashboard | Coordenadores | Conteúdo |
|---|:---:|:---:|---|
| `port-visao` | ✅ | ✅ | Visão Geral: KPIs de portfólio, cruzamento Portfólio×Agendado, ofensores, por mês, overview de categoria. |
| `port-ordens` | ✅ | ❌ | Tabela de Ordens (envios/alunos/status por Ordem). |
| `port-praticas` | ✅ | ✅ | Ranking de práticas (pendentes/enviadas), barras por categoria, filtro por curso. |
| `port-polos` | ✅ | ✅ | Tabela de polos + mapa do Brasil (D3) + modais por estado. |
| `port-tutores` | ✅ | ❌ | Tabela de tutores (portfólios enviados, % conclusão, situação). |
| `tutores-mec` | ✅ | ✅ (como "Ficha dos Tutores") | Fichas de tutor: MEC, admissão, CH, laboratório, badges (treinamento/obras), consolidado de gerenciamento. |
| `ger-visao` | ✅ | ❌ | Gerenciamento — Visão Geral (KPIs de ofertas). |
| `ger-ofertas` | ✅ | ❌ | Gerenciamento — Ofertas por polo/categoria/ordem. |
| `ger-agendas` | ✅ | ✅ | Gerenciamento — Agendas (ofertas com/sem agenda por polo). |
| `agendas-estudo` | ✅ | ✅ | Horários e Engajamento: dia/turno/região, evolução, melhores horários. |
| `ger-detalhe` | ✅ | ✅ | Detalhe por tutor: % gerenciado, pizza, "em treinamento", "em obras", modal por tutor. Filtro por `GRUPOS_GER`. |
| `vagas` | ✅ | ✅ | RH/Vagas: posições em aberto da Lotação **+ (PATCH 157)** funil de recrutamento agregado, vagas no INHIRE (só dashboard) e cruzamento Lotação↔INHIRE. Coordenadores vê o funil filtrado pelo curso. |
| `insumos` | ✅ | ❌ | Racional de Insumos por Experimento (estudo estático). |

### Só no `template_coordenadores.html`
- **Seletor / painel de curso** (`_montarPainelCurso`, `abrirPainelCurso`,
  `_aplicarCursosSelecionados`, `_filtrarDBPorCursos`, ~L4787+): trava o portal
  a um ou mais cursos; filtra TODAS as coleções de `DB` e recalcula KPIs
  (PATCH 111/116/135/145). `CURSO_PARA_CATEGORIA` (~L4787) mapeia curso → categoria.
- `_iniciarPortalCoordenador`, `_clonarDB`, `_kpisParaSemestre`,
  `_recalcularPoloStatsDeTutores`, `_categoriaBate` — só existem aqui.
- Versões simplificadas de KPI (`_renderVagasKPIs_ORIGINAL_NAO_USAR` etc. são
  restos mortos — não usar).

### Recursos comuns notáveis
- Login: `initSenha` / `verificarSenha` / `_decifrarDados` (Web Crypto, AES-GCM).
- Seletor de semestre: `initSemestres` / `selecionarSemestre` / `_getSem*()`
  (troca o recorte de `DB` pro semestre da aba). Gerenciamento tem par próprio:
  `_gerSelecionarSemestre` / `_gerMergeAllSemestres` / `_gerAggregateFromOfertas`.
- Mapa do Brasil: `renderMapaBrasil` → `_renderMapaD3` / `_renderMapaD3Alt` /
  `_renderMapaFallback` (3 níveis de degradação).
- Dark mode, sidebar recolhível, exports CSV (`exportCSV` e vários `export*`).
- `_ativarMoria` — easter egg do login.

---

## 9. template_dashboard.html — índice de funções JS

Números de linha aproximados (arquivo ~5.111 linhas).

### Infra / login / util
`_sha256Buf` 801 · `_hexToBytes` 804 · `_bytesToHex` 809 · `_decifrarDados` 812 ·
`verificarSenha` 819 · `initSenha` 838 · `_ativarMoria` 859 ·
`gPN/gPT/gPE/gPA/gTS` 879-883 (getters tolerantes a chave curta/longa) ·
`getPraticas` 884 · `_getSemPraticas` 893 · `_getSemPraticaStats` 898 ·
`getOE/getOA/getOS` 903-905 (envios/alunos/status por ordem) · `navTo` 919 ·
`esc` 936 · `pctColor` 937 · `pctBadge` 938 · `progCell` 939 · `fmtNum` 940 ·
`ordemNum` 941.

### Portfólio (Visão Geral / Ordens / Práticas / Polos / Tutores)
`renderKPIs` 944 · `renderCruzamentoPortfolioAgendado` 964 ·
`switchCruzamentoTab` 990 · `renderOfensores` 996 · `renderMes` 1004 ·
`renderCatOverview` 1010 · `renderOrdens` 1027 · `filterOrdemTbl` 1039 ·
`sortOrdem` 1044 · `renderOrdemTbl` 1045 · `renderPracKPIs` 1053 ·
`renderCatPracBars` 1062 · `switchRankTab` 1098 · `renderTopPrac` 1099 ·
`setFilterPrac` 1108 · `filterPrac` 1114 · `sortPrac` 1115 · `renderPracTbl` 1116 ·
`filterPolos` 1123 · `sortPolo` 1124 · `renderPoloTbl` 1125 · `sortTutor` 1139 ·
`renderTutorTbl` 1140 · `openListModal` 1160 · `openTutorModal` 1181 ·
`openPoloModal` 1225 · `closeModal` 1245 · `openAlunosModal` 1249.

### Gerenciamento
`renderGerenciamento` 1318 (helpers `_poloKeyCanon`, `_semChapaCanon`,
`_nomesBatemCanon`, `_chapaDe`, `_poloKeySimples` — só locais) ·
`openGerPoloModal` 1540 · `renderGerPoloTbl` 1563 · `filterGerPolo` 1567 ·
`renderGerContrTbl` 1568 · `toggleContrGrp` 1575 · `filterGerContr` 1576 ·
`renderAgendaKPIs` 1577 · `renderGerAgendaTbl` 1583 · `filterGerAgenda` 1632 ·
`sortDetalhe` 1633 · `_grupoById` 1673 · `_semAcento` 1674 ·
`_grupoBaseDaCategoria` 1691 · `_subcursoDeCodigo` 1720 · `_rowMatchesGrupo` 1726 ·
`_tutorMatchesGrupo` 1747 · `_updateGerSubfiltro` 1764 · `filterGerDetalhe` 1787
(helper `_baseKeyOferta` 1824 + `_normPoloComparavel`, `_primeiroUltimo`,
`_nomesBatem` locais) · `renderDetPizza` 2220 · `renderDetTreinamento` 2273 ·
`renderDetObras` 2324 · `exportRelatorioPorCurso` 2366 · `_openDetTutorModal` 2386 ·
`toggleDetGrp` 2413 · `abrirSemGerenciamento` 2892 · `_renderSemGer` 2901 ·
`exportarSemGer` 2951.

### Mapa / estados
`extrairUF` 2417 · `agruparPorEstado` 2418 · `renderEstadosInline` 2420 ·
`filterGerEstadosInline` 2426 · `openEstadoModal` 2427 · `_corEstado` 2498 ·
`renderMapaBrasil` 2507 · `_corEstadoD3` 2604 · `_renderMapaD3` 2614 ·
`_renderMapaD3Alt` 2731 · `_renderMapaFallback` 2753 · `_corLight` 2762 · `_gradId` 2771.

### Filtros globais / tema
`populateFilters` 2773 · `toggleDark` 2803 · `initDark` 2804 · `toggleSidebar` 2805 ·
`initSidebar` 2811.

### Exports
`exportEnviados` 2817 · `exportSemTutor` 2818 · `exportTutores` 2819 ·
`exportPolosSemGerenc` 2820 · `exportDetalhe` 2858 · `exportarSemGer` 2951 ·
`exportTutoresAtivosPorCurso` 4909 · `exportTutoresDesligados` 4946 ·
`exportAnaliseEngajamento` 3776 · `exportAgendas*` 3860-3866 · `exportVagas` 3960.

### Alunos por curso / categoria / horários
`openAlunosCursoModal` 2959 · `_ordensParaCategoria` 3015 · `catParaCurso` 3020 ·
`cursoLabel` 3029 · `getCatColor` 3037 · `openCategoriaHorariosModal` 3041 ·
`openDiaSemanaModal` 3091 · `openTurnoModal` 3171.

### Agendas / engajamento
`_getOfertasAgenda` 3228 · `_regiaoDoPolo` 3245 · `_dedupAlunos` 3251 ·
`_renderAgendasKpisDiasTurnos` 3272 · `_nomeCursoOferta` 3340 ·
`_popularFiltroCurso` 3361 · `_popularFiltroOrdem` 3385 · `_popularFiltroTutor` 3392 ·
`filterAgendasPoloTutor` 3413 · `renderAgendasEstudo` 3432 ·
`_popularFiltroGrupoEngajamento` 3583 · `_ordemNumOrd` 3590 · `renderEvolucaoGeral` 3592 ·
`renderMelhorEngajamento` 3647 · `renderEngajamentoPorOrdem` 3678 ·
`_labelGrupoOferta` 3772 · `filterAgendasIncomuns/SemAluno/Pratica` 3802-3839.

### Vagas (RH)
`_getVagas` 3873 · `populateVagasFilters` 3874 · `renderVagasKPIs` 3885 ·
`filterVagas` 3922 · `sortVagas` 3937 · `renderVagasTbl` 3938 · `exportVagas` 3960.

### Ficha dos tutores (MEC) + badges
`_horasAdminPorCH` 3981 (PATCH 27) · `_capOrdemPorNomeTutor` 3989 · `_normTutorKey` 4021 ·
`_labPendenciaDoTutor` 4033 · `_badgeLabPendencia` 4043 · `_diasDesdeAdmissao` 4049 ·
`_estaEmTreinamento` 4073 · `_badgeTreinamento` 4078 · `formatTutorShort` 4083 ·
`getTutorCatAbrev` 4106 · `renderAvisosPortfolio` 4190 · `renderTutoresMec` 4626 ·
`_filtraTutoresPorCategoriaSel` 4735 · `filterTutoresMec` 4751 · `_tutMecSort` 4970.

### Semestre / consolidação
`_gerSelecionarSemestre` 4226 · `_gerMergeAllSemestres` 4236 ·
`_gerAggregateFromOfertas` 4248 · `initSemestres` 4367 · `selecionarSemestre` 4401 ·
`_mergeAllSemestres` 4429 · `_getSemKpis/PortfolioAlunos/Cruzamento/PoloStats/PorOrdem/Prazos/StatusOrdem/Tutores` 4469-4523 ·
`_renderComSemestre` 4549 · `_construirIndiceGerPorOrdem` 4572 · `_gerConsolidadoDoTutor` 4597.

### Bootstrap / insumos
`_iniciarDashboard` 4976 · `initInsumos` 5008 · `switchInsumosTab` 5038 ·
`filterInsumosConsulta` 5044 · `showInsumosDetalhe` 5057 · `renderInsumosCobertura` 5073 ·
`renderInsumosLacunas` 5084 · `filterInsumosLacunas` 5095.

### Constantes JS
`GRUPOS_GER` 1654 · `CAT_CURSO` 3019 · `CAT_COLORS` 3036 · (`UF_NOMES`, `_CAT_NORM_MAP`
existem também — grep).

---

## 10. template_coordenadores.html — o que muda

Estrutura de funções ~90% idêntica ao dashboard (mesmos nomes, offsets
diferentes — arquivo ~5.174 linhas). Diferenças relevantes:

- **Não tem:** `port-ordens`, `port-tutores`, `ger-visao`, `ger-ofertas`,
  `insumos`. Não tem `_popularFiltroCurso`/`_popularFiltroOrdem` em Agendas
  (só `_popularFiltroTutor`), nem `renderEngajamentoPorOrdem` separado da mesma forma.
- **Tem a mais (exclusivo):** painel/seletor de curso —
  `_iniciarPortalCoordenador` 4821 · `_clonarDB` 4842 · `_montarPainelCurso` 4845 ·
  `abrirPainelCurso` 4865 · `_fecharPainelCurso` 4871 · `_aplicarCursosSelecionados` 4874 ·
  `_aplicarMarcaCursos` 4883 · `_filtrarDBPorCursos` 4909 (helper `_filtraPraticasArr`) ·
  `_kpisParaSemestre` 5099 · `_normTxtCoord` 5128 · `_recalcularPoloStatsDeTutores` 5129 ·
  `_categoriaBate` 4804 · `const CURSO_PARA_CATEGORIA` 4787 · `PATCH 116` banner ~L700.
- `renderVagasKPIs` / `renderVagasTbl` são versões simplificadas; existem
  `_renderVagasKPIs_ORIGINAL_NAO_USAR` (3675) e `_renderVagasTbl_ORIGINAL_NAO_USAR`
  (3753) — **restos mortos, não usar**.
- `renderGerenciamento` (1230) tem versões simplificadas de KPI para coordenador
  (PATCH 111).
- `GRUPOS_GER` 1564 · `CAT_CURSO` 2932 · `CAT_COLORS` 2949.

**Ao alterar qualquer coisa de portfólio/gerenciamento/agendas/mapa/vagas no
dashboard, procure a função de mesmo nome aqui e replique.**

---

## 11. Build e deploy (GitHub Actions)

`.github/workflows/publicar.yml` — jobs `build` (que também faz o deploy) em `ubuntu-latest`.

**Gatilhos:** push em `main`, `workflow_dispatch`, cron `0 */2 * * *` (a cada 2h).
`concurrency: dashboard-deploy` (não cancela em andamento).

**Passos (PATCH 160 — publica via artifact do Pages, não commita mais HTML):**
1. `checkout` + `setup-python@3.11` + `pip install pandas openpyxl xlrd requests cryptography`.
2. **Baixar planilhas** — script Python inline. `sp_download_url()` converte link
   de compartilhamento SharePoint em `_layouts/15/download.aspx?share=TOKEN`
   (formato que funciona no tenant UNIASSELVI), com 2 fallbacks. Rejeita resposta
   HTML (página de login). Secrets: `URL_CONTROLE`, `URL_PORTFOLIO`,
   `URL_PORTFOLIO_2026_2`, `URL_GERENCIAL`, `URL_GERENCIAL_26_02`, `URL_LOTACAO`,
   `URL_REL_NOVO`, `URL_ALUNOS_HUB`, `URL_VAGAS_RH` (PATCH 157). Só CONTROLE e
   PORTFOLIO são obrigatórios.
3. `python processar.py --sem-browser`.
4. **Montar site:** `rsync` do checkout pra `_site/` **excluindo** `.git`,
   `.github`, `planilhas/`, `saida/`, `processar.py`, os `template_*.html`,
   `*.md`, `*.sh` (o site nunca precisou desses; `processar.py` ainda vazava a
   senha `uniasselvi2026`). Depois copia `saida/dashboard.html → _site/index.html`,
   `saida/coordenadores.html → _site/coordenadores.html`,
   `saida/lookup.json → _site/lookup.json`.
5. `configure-pages` + `upload-pages-artifact (path: _site)` + `deploy-pages`.
   **Nada é commitado no git** — o `.git` parou de crescer ~1 GB/dia.
6. **Keepalive:** às segundas, commita `.github/last_run.txt` (único commit que o
   workflow ainda faz — evita o GitHub desativar o cron por inatividade).

**GitHub Pages** (source = "GitHub Actions", `build_type: workflow`) serve o
conteúdo de `_site/`: `index.html`, `coordenadores.html`, `lookup.json`
(consumido pelo `portfolio_form.html`), `portfolio_form.html`, `robots.txt`, os
`.json` de dados. Mesmo endereço de sempre.

`insumos_estudo.json` e `catalogo_oficial.json` são **estáticos no repo** — pra
atualizar, substituir o arquivo e deixar o pipeline rodar (não mexe em código).

---

## 12. Glossário de convenções

| Termo | Onde | Significado |
|---|---|---|
| `SEMESTRE_ATUAL` | `processar.py:108` | Maior chave de `ALL_SEMESTRES` (string sort). O semestre "ativo" do dashboard. |
| `ALL_SEMESTRES` | `processar.py:82` | `{ '2026/1': {prazos, periodos}, ... }` de `config_semestre.json`. |
| `CAT_MAP` | `processar.py:267` | Rótulo curto de categoria → nome de exibição longo. |
| `CAT_CURSO` / `CAT_COLORS` | templates | Categoria → nome curto do curso / cor do gráfico. |
| `CURSO_PARA_CATEGORIA` | `template_coordenadores.html:4787` | Curso selecionado no painel → categoria(s) correspondente(s). |
| `GRUPOS_GER` | templates (`:1654` dash / `:1564` coord) | Os **5 grupos de área** dos filtros de Gerenciamento: Enfermagem / Fisioterapia e T.O. / Biomedicina-Farmácia-Estética / Nutrição / Exatas (com subfiltro Engenharias vs Química-Física). Substitui as categorias cruas (PATCH 47). |
| `_OFE` | `processar.py:2895` | Nº de ofertas/vagas cadastradas que uma linha do GIOCONDA representa (coluna `OFERTAS_CADASTRADAS`). |
| `_PESO_OFERTA` | `processar.py:2944` | `_OFE` com mínimo 1; a linha é replicada esse nº de vezes pra virar "ofertas reais" — mas só depois de deduplicar por prática (PATCH 150). |
| `_GERENCIADO` | `processar.py:2922` | Booleano: tem tutor **E** tem data de gerenciamento (`DT_GERENCIADA`) (PATCH 22). |
| `_baseKeyOferta(r)` | templates (JS) | Chave de agrupamento de uma oferta = tutor+polo+subcurso, usada no Detalhe de Gerenciamento pra somar geridas/total por tutor (PATCH 48/53). |
| `_SEMESTRE_ORIGEM` | `processar.py:712` | Semestre de uma submissão de portfólio = origem do arquivo, não a data (PATCH 127). |
| `controle_tutor_lookup` | `processar.py:4114` | `{(polo_norm, categoria|curso): nome}` — backfill do TUTOR quando o GIOCONDA ainda não preencheu (PATCH 21/32). |
| `_FINO_PARA_AMPLO` | `processar.py:4230` | Código fino de curso (BFI/BTO/...) → rótulo amplo do `CAT_MAP`. |
| "Aviso de Portfólio" | vários | Pseudo-tutor de submissão sem tutor real correspondente; excluído de KPIs/lookup/gerenciamento. |
| `tutores_desligados` | `processar.py:2102` | Lista separada de tutores inativos + data de desligamento (PATCH 105). |
| `apto` / `onboarding_apto` | `processar.py:4043` | Tutor com Trilha + Checklist + 1:1 todos "Sim" → sai da lista de treinamento. |
| `_anomalia_ordem` / `_anomalia_ordem_futura` | `processar.py:3240` | Geriu ordem avançada durante período de ordem anterior (PATCH 97) / ordem ainda não começou = erro de preenchimento (PATCH 130). |
| `DATA_GOES_HERE` / `TIMESTAMP_GOES_HERE` | templates | Placeholders substituídos por `gerar_html*`. |
| `p1..p6` | `processar.py` | CONTROLE / PORTFÓLIO / (template) / GIOCONDA antigo / GIOCONDA CSV / LOTAÇÃO / hub CSV / ONBOARDING. (`p2b`/`p3b` = variantes 2026/2). |

---

## 13. Histórico de PATCHes (resumido)

Os comentários numerados `PATCH N` no código são o changelog real do projeto.
Resumo de uma linha por patch (ver o comentário no código pro detalhe completo).
Onde a origem é evidente: **[py]** = `processar.py`, **[dash]** / **[coord]** =
templates, **[both]** = replicado nos dois.

| # | Resumo |
|---|---|
| 1 | [py] Detecta e converte coluna "CH SEMANAL" (`HH:MM` ou decimal) para float. |
| 2 | [py] `tem_lotacao` passa a ser baseado em dado real (CH>0 em ≥1 tutor). |
| 3 | [py] Situação por Ordem usa datas de início reais de `PERIODOS_ORDENS`. |
| 4 | [py] Corrige situação quando não há nenhuma ordem vencida. |
| 5 | [py] Reavalia situação e sincroniza `sit`/`situacao`. |
| 6 | [py] Remove prints duplicados de p1/p2. |
| 7 | [py] Estrutura completa `datas_por_tutor` no gerenciamento. |
| 8 | [py] Cifra o JSON (AES-256-GCM) antes de injetar no HTML. |
| 9 | [py] Gera `lookup.json` público (sem cifra) pro `portfolio_form.html`. |
| 10 | [py] Suporte à planilha nova de portfólio 2026/2 (schema próprio) + `id_to_perfil`. |
| 11 | [both] `initSemestres` limpa abas antes de recriar (evita botões duplicados). |
| 12 | [`portfolio_form.html`] Chamada direta à API do SharePoint é bloqueada por CORS — redireciona pro `listform.aspx` em vez de POST via fetch. |
| 13 | [py/both] Seção "Laboratórios" + bloco de acompanhamento de comunicação na ficha do tutor. |
| 14 | [py] Nomes incompletos no relatório de comunicação casados por primeiro+último. |
| 15 | [py] `nome_to_perfil.json` — nome da prática → perfil pra códigos ambíguos (BFI/BTO/COS-TIP). |
| 16 | [py] Práticas enviadas antes da admissão do tutor não contam pra ele. |
| 17 | [py] Agrupa por curso específico (não categoria ampla) no catálogo/agregado por polo. |
| 18 | [py/both] Gerenciamento separado por semestre; GIOCONDA lido como `.xlsx` ou `.csv`; declarações JS que faltavam. |
| 19 | [both] Tabela Detalhe ordenável (Polo/Tutor/% Ger.); exibe curso específico; `_sortState`. |
| 20 | [both] Declara `_sortState` (faltava — quebrava `sortTable()` em silêncio). |
| 21 | [py] Backfill de TUTOR a partir do CONTROLE quando o GIOCONDA não preencheu. |
| 22 | [py] `GERENCIADO` = tem tutor **e** tem data de gerenciamento. |
| 23 | [py] E-mail do remetente identifica a tutora de forma inequívoca no cruzamento de portfólio. |
| 24 | [py] `nome_to_perfil` também separa o catálogo de "práticas previstas" por curso. |
| 25 / 25a / 25b / 25c | [py] Validação obrigatória de colunas críticas; detecção de e-mail duplicado no CONTROLE; snapshot/regressão de colunas + contagens. |
| 26 | [both] Normalização canônica de nome de tutor; preferir sempre a variante com chapa. |
| 27 | [both] Horas administrativas por faixa de CH — correção do percentual fixo de 25% do PATCH 26. |
| 28 | [both] Só ENGMAKER e Química-Física/Agronomia têm Ordem 5. |
| 29 / 29a | [py] Página de Vagas (RH) da Lotação; "Aumento de Quadro" só conta como vaga real sob condição. |
| 30 | [py] Helpers de dia-da-semana e turno (Análise de Agendas). |
| 32 / 33 | [py/both] Garante que todo tutor ativo apareça no gerenciamento mesmo sem oferta no GIOCONDA; corrige falso-positivo "Sem oferta". |
| 34 | [both] Resumo "Gerenciamento de Ofertas" (pedido Anderson/Leo). |
| 35 | [both] Semestre padrão = sempre o semestre ativo do `config`. |
| 37 | [both] Sem dado de oferta no GIOCONDA ≠ "sem capacidade". |
| 38 | [py] **Dedup por polo×categoria dentro da ordem antes de somar alunos** (GIOCONDA infla). |
| 39 | [both] Guarda a lista filtrada atual pra exportação. |
| 40 | [py] Exclui pseudo-tutores "Aviso de Portfólio" da injeção de ofertas. |
| 41 | [py] Lista fixa de categorias reais válidas pra injeção de ofertas. |
| 42 | [py] Curso específico (BFI/BTO/COS-TIP/...) no gerenciamento novo. |
| 43 | [py] Auto-detecção de esquema ANTIGO vs NOVO do hub CSV. |
| 44 | [both] Deduplicar por NOME de tutor (não por grupo tutor+polo+categoria). |
| 45 | [py] `_norm_polo_*` também remove parênteses e acentos, não só prefixo "LAP -". |
| 47 | [both] Filtros de Gerenciamento reagrupados nos 5 `GRUPOS_GER` por área real. |
| 48 | [both] Dedup do PATCH 38 replicado no frontend (polo soma o MAIOR valor visto por prática). |
| 49 | [both] Resolve chapa/homônimo antes de agregar; remove card "Presença Lançada". |
| 50 | [both] Tabela filtrada de verdade por ordem/status. |
| 51 / 54 / 59 | [both] Recalcula `ger_ok/ger_total/ger_pct` por tutor+polo no frontend; subfiltro de grupo. |
| 52 | [both] Corrige tutores aparecendo no grupo de área errado (comparação de categoria). |
| 53 | [both] `baseKey` já inclui subcurso → cada linha recebe um subcurso só. |
| 55 | [both] Só separa linha por subcurso quando ele é especialidade real. |
| 56 | [both] Corrige tutores de Estética/Bio-Far classificados errado. |
| 57 | [both] Match de categoria sem exigir ")" logo após o numeral. |
| 58 | [both] Resolve especialidade por nome+polo+categoria (não só nome). |
| 60 | [both] Placeholder da opção "todos" não fica preso ao rótulo antigo. |
| 61 / 63 / 64 / 65 | [both] "Total real" ≠ capacidade causava bugs; PATCH 65 confirma com planilha real que o certo é a capacidade contratada (4h/8 ordens etc.). |
| 67 | [both] Farmácia e Biomedicina = mesma vaga/categoria (Bio-Far). |
| 68 | [both] Comparação de capacidade agregada por tutor+ordem (não por linha). |
| 69 | [both] "Não gerenciadas" = só quem tem zero geridas. |
| 70 / 71 | [both] Guarda lista completa (antes do filtro de status) pro gráfico pizza. |
| 72 | [both] Lista "sem oferta no GIOCONDA" sempre preenchida (pro pizza). |
| 73 | [both] Tutor com ativa mas nenhuma oferta real no GIOCONDA tratado à parte. |
| 74 / 75 / 76 / 77 | [both] Match de nome com parte a mais/menos; bloqueio de re-injeção de "sem oferta"; caso "Renata Souza da/De Silva" ainda duplicada. |
| 78 / 79 | [both] "Ver 1 por 1" filtra KPIs + gráficos por polo/tutor; seletores dependem do recorte. |
| 80 | [both] Modais de dia/turno respeitam o recorte filtrado atual. |
| 81 | [both] Filtro de Ordem não restringe mais os dados usados pra montar a lista; Ordem + Status decidem juntos. |
| 82 | [py] Normaliza rótulo com prefixo "BIO-" duplicado na fonte. |
| 83 | [py] Match de nome por subsequência de palavras (fallback). |
| 84 | [both] KPI de prazos usa o semestre selecionado, não `DB.prazos` global. |
| 85 (P7) | [both] Tooltip de agenda inclui o horário de cada agendamento. |
| 86 (P6) | [py/both] Conta tutores distintos que gerenciaram qualquer coisa por ordem. |
| 87 | [both] Exporta os tutores ativos atualmente exibidos na tela. |
| 88 / 89 | [py] Lê planilha de Onboarding e aplica flags; (re)gera `Acompanhamento_Onboarding.xlsx` automaticamente. |
| 90 | [py] Campo só de exibição com curso específico (Fisioterapia/T.O./...). |
| 91 | [both] Trata grupo combinado (`grupo:bio-far-est`) no consolidado. |
| 92 | [both] Sem ordem específica selecionada → "Todas as ordens". |
| 93 / 94 | [both] Segundo gráfico lado a lado; universo de tutores não filtra por ordem. |
| 95 / 96 | [py] Fallback de match por subsequência (nome com parte a mais/menos). |
| 97 | [py/both] GIOCONDA passou a permitir gerenciar qualquer ordem; detecta/sinaliza gerenciamento fora do período da ordem. |
| 98 | [both] "Match 2/3" por subsequência de palavras completas (não substring). |
| 99 / 100 | [both] Match de tutor em 2/3 passadas, da mais confiável pra menos. |
| 102 | [py] Nome real do arquivo é "PORTIFOLIO" (com I). |
| 103 | [both] Ficha do tutor inclui data de admissão (busca em `DB.tutores` por nome). |
| 104 | [both] Filtro extra pedido pra Enfermagem (genérico pra qualquer curso). |
| 105 | [py] Captura tutores desligados numa lista separada (`tutores_desligados`). |
| 106 | [py] Coluna DESLIGAMENTO vem como data/hora de verdade. |
| 107 | [py] `processar_vagas(p4)` estava aninhado dentro de `if p6:` — nunca rodava; agora roda sempre que houver Lotação. |
| 108 | [py/both] Stats/pratica_stats por semestre; frontend lê o recorte do semestre, não `DB.praticas` global. |
| 109 | [both] Popula filtro de curso específico (Biomedicina, Farmácia, ...). |
| 110 | [py] Gera `coordenadores.html` — segundo portal, mesma cifra/senha. |
| 111 | [coord] Versões simplificadas de KPI pro coordenador; categoria no formato curto. |
| 112 / 113 | [both] Mapa código de curso → nome legível; curso específico → laboratório/categoria ampla. |
| 114 | [both] Dedup por polo+categoria nos modais de turno/dia-semana. |
| 115 | [py] GIOCONDA passou a trazer nome do tutor com sufixo de chapa — limpar. |
| 116 | [coord] Seleção de curso deixou de ser tela cheia bloqueante; agrupada por laboratório. |
| 117 | [both] KPI "Total Tutores" já excluía entradas anônimas/"Aviso de Portfólio". |
| 118 | [py] **Movido do JS pro backend**: remove do GIOCONDA linhas de tutores sem vínculo ativo hoje (355 no gráfico vs 348 no KPI). Cuidado: não excluir quem voltou/trocou de polo. |
| 119 | [both] Visão "Matriculados × Agendados por Prática". |
| 120 | [py] Dois casos confirmados como erro de preenchimento de ordem na origem. |
| 121 | [py/dash] Aba "Racional de Insumos por Experimento" — estudo estático (`insumos_estudo.json`). |
| 122 | [py] Soma alunos de TODAS as colunas de estudante (não só "a de maior sufixo"). |
| 123 | [py/both] "Alunos registrados no Portfólio" dedup por polo×categoria (mesmo bug do PATCH 38); cruzamento Portfólio×Agendado nos 3 níveis. |
| 124 | [py] Contatos por polo (Lista Oficial de Contatos) na ficha do tutor. |
| 125 | [both] "Engajamento por Ordem" agrupado pelos mesmos 5 grupos (revisado 2x). |
| 126 | [py/both] Base de alunos do polo em cada vaga. |
| 127 | [py/both] Semestre do portfólio vem da origem do arquivo; cruzamento e "Alunos Registrados" por semestre selecionado. |
| 128 | *(sem entrada no código — número pulado)* |
| 129 | *(sem entrada no código — número pulado)* |
| 130 | [py/both] Segunda checagem de anomalia: gerenciamento marcado numa ordem que ainda não começou; exclui essas ofertas dos agregados. |
| 131 | [both] "Alunos matriculados" usa `DB.alunos_por_curso` (hub CSV), prioridade invertida. |
| 132 | [both] UF → macrorregião; "Evolução do Engajamento" e "Horários com Maior Engajamento" por curso e região. |
| 133 | [py] Cruzamento Portfólio×Agendado: direção estava invertida — AGENDADO é a base, taxa = registrado/agendado. |
| 134 | [both] Sinalização de tutor novo (≤60 dias) / filtro independente por status de treinamento. |
| 135 | [both, coord principalmente] Filtro de curso checa também `subcurso`; recalcula `DB.kpis` e KPIs por semestre a partir dos tutores filtrados. |
| 136 | [coord] "Alunos Agendados" (Visão Geral) lê direto de `DB` recalculado, não de valor pré-calculado. |
| 137 | [both] Gráfico "Em Treinamento"; "em treinamento" exige também que o tutor ainda não tenha gerenciado. |
| 138 | [both] Gráfico "Em Treinamento" ignorava o filtro — corrigido (dash e coord). |
| 139 | [py] Checagem "essa prática já é conhecida oficialmente sob outra categoria" (nomes tipo "ENGMAKER - Análise estrutural de uma viga"). |
| 140 / 141 / 141b / 141c / 141d | [py] Diagnóstico + correção do "vazamento de categoria": bucket anônimo "Tutor desligado" em Manaus com práticas de outra categoria; resolução de categoria por votação; diagnóstico final de divergência. |
| 142 | [both] `gerDetData` setado e usado logo no início da render (não mais no fim). |
| 145 | [coord] Bug grave: checagem de subcurso no filtro de curso; ambiguidade julgada pelo total global (revisado 2x). |
| 146 | [py] Parsing do portfólio 2026/2 lia só 1 das ~7 "rodadas" do formulário. |
| 147 | [py] Cruzamento por polo comparava nomes de polo sem normalizar. |
| 148 / 148b | [py/both] Ícone "em obras / não apto" (`labs_pendencias.json`); gráfico "Em Obras/Não Apto" (coord). |
| 149 | [py] **URGENTE**: PATCH 147 chamava `_norm_polo_ger` (que só existe dentro de outra função) → `NameError` → `except` genérico marcava `tem_gerenciamento=False` e escondia a aba inteira. Corrigido com função local `_norm_polo_cruzamento`. |
| 150 | [py] `_OFE` × `_PESO_OFERTA`: expandir cada linha pelo próprio peso multiplicava em dobro — deduplicar por prática antes. |
| 151 | [py] Achado GRAVE: PATCH 146 lia só uma coluna de rodada; ler todas. |
| 152 | [both] Revertido: gráficos "Em Treinamento"/"Em Obras" usam sempre o universo completo (filosofia do PATCH 94). |
| 153 | [py] (superado pelo 155) Assumia sempre data de contratação brasileira. |
| 154 | [py] Diagnóstico: tutores com início ≤60 dias vs. sem nenhuma data utilizável. |
| 155 | [py] `_interpretar_data_contratacao` — resolve datas BR/US misturadas na mesma coluna. |
| 156 | [py] Mesmo uma data nativa do Excel pode estar ambígua BR/US — tratar também. |
| 166 | [both] **Vagas/RH — limpeza visual do PATCH 165.** Fileira de KPIs 12→4 cards; impacto vira 1 faixa "O que decidir agora" (não mais cards); rodapé do funil: sai "por mês", dificuldade por curso em largura cheia, salário em 1 linha. |
| 165 | [py/both] **Vagas/RH expandida pra decisão de coordenador/diretor.** `processar_vagas_rh`: `reqs[].situacao` ∈ aberta/fechada/congelada/cancelada (antes só bool `fechada`) + `travada` (etapa "sem currículos"); `funil.por_curso` com `taxa` de conversão (agregado em curso CANÔNICO via `_vrh_curso_canon` — o campo é texto livre caótico); ciclo mediano contato→entrevista; `salarios_vagas_abertas`. Nova `_analisar_vagas_criticas` → `dados['vagas']['criticas']` (vagas sem tutor priorizadas por alunos+travas), `polos_dificeis` (>200 alunos sem tutor), `dificuldade_por_curso`, e KPIs `alunos_sem_tutor`/`polos_sem_tutor`. Roda DEPOIS do enriquecimento de `alunos_polo` pelo hub. Frontend: banner de impacto + cards "Vagas que precisam de decisão" e "Polos difíceis"; funil ganhou "dificuldade por curso" e "salário oferecido". Coordenadores: tudo filtrado pelo curso no `_filtrarDBPorCursos`. |
| 164 | [py] **Engajamento da "próxima ordem" não é mais escondido.** PATCH 130 marcava `_anomalia_ordem_futura` pra qualquer ordem cujo `período.início` (config = data das PRÁTICAS) está no futuro, e o frontend esconde essas ofertas do Engajamento por Ordem. Mas o agendamento do aluno / montagem de agenda abre antes (quando a ordem anterior entra em prática). Agora só marca anomalia quando nem a ordem ANTERIOR começou. Caso: 1786 gerenciamentos da Ordem 3 (2026/2, 247 tutores) sumiam. `_detectar_gerenciamento_fora_ordem` (`processar.py:~3673`). |
| 163 | [both] **Card "Ocupação" removido** da Visão Geral de Gerenciamento — mostrava "0%" fixo (`total_capacidade` nunca é calculada; Agendados÷Matriculados dá >100% porque as bases são incompatíveis). |
| 162 | [py] **`polo_stats` conta só tutores reais** (mesma lógica do 158 pro headcount). Baldes-fantasma criavam ~26 linhas de polo e infhavam `total_polos` (235→209) e a soma de tutor/polo (396→356). Alunos/envios de prática órfã ainda somam num polo existente. Raiz + `_stats_semestre`. |
| 161 | [py] **Correção manual de aviso não vale pra desligado.** `_CORRECOES_MANUAIS_AVISO` (PATCH 120) só se aplica se a pessoa NÃO está em `tutores_desligados`. Caso: Ingrid Schroeder Pineiro (correção de 21/08, desligada 28/07) contava como ativa (357 vs 356). + status `Encaminhar Positiva`/`Coordenação` no funil. |
| 160 | [infra] **Publica via artifact do GitHub Pages, não commita mais os HTML gerados.** `.git` crescia ~1 GB/dia (rodada de 2h × ~86 MB de `index.html`+`coordenadores.html` não-comprimíveis). Workflow monta `_site/` e usa `deploy-pages`; Pages passou pra `build_type: workflow`. `index.html`/`coordenadores.html`/`lookup.json` foram pro `.gitignore` e removidos da history com `git filter-repo` (force-push no `main`). **Todos os SHAs de commit mudaram — quem tinha clone teve que reclonar.** |
| 159 | [py/both] **Recrutamento: funil simplificado + coerência dos "Aviso de Portfólio".** Funil vira 3 blocos (Em andamento / Contratados / Não seguiram) + top-3 motivos de saída, no lugar de 13 barras. KPIs de recrutamento saem da fileira de Vagas da Lotação. "Aviso de Portfólio" de ex-tutor casa com `tutores_desligados` → polo limpo + rótulo "ex-tutor desligado em DD/MM" (`c` continua 'Aviso de Portfólio' pra não inflar headcount; categoria real em `categoria_ex_tutor`). Remetente sem nome → "Remetente não identificado". Statuses novos da planilha real no `_VRH_STATUS_GRUPO`. Rótulo de status com acento (`st_label`). |
| 158 | [py] **KPI headcount de tutores excluía os fantasmas.** `total`/`enviaram`/`pendentes`/`urgentes`/`atrasados` (`processar.py:~1911` e dentro de `_stats_semestre` `:~2098`) contavam os baldes `{'n':'Tutor desligado','_anonimo':True}` e os pseudo-tutores `'Aviso de Portfólio'`. Como demitido sai do CONTROLE e vira fantasma, o total ficava travado. Novo helper `_eh_tutor_fantasma`; os fantasmas seguem em `tutores_out` (seção Aviso de Portfólio / agregados de prática dependem deles), só não contam no headcount. `total_alunos`, `polo_stats`, `pratica_stats` e `por_ordem` **não** mudaram (atividade de portfólio órfã é real). Frontend já filtrava no `renderKPIs`, mas lia `k.urgentes`/`k.atrasados` crus do backend. |
| 157 | [py/both] **Recrutamento & Seleção na seção Vagas.** `p7`/`URL_VAGAS_RH` → `processar_vagas_rh` + `_cruzar_vagas_recrutamento`, anexados em `dados['vagas']` (`recrutamento.reqs`, `funil`, `kpis` estendido). Funil de candidatos **agregado, sem PII** (LGPD — candidato externo). Frontend: KPIs de recrutamento + card "Funil de Recrutamento" + tabela "Vagas no INHIRE" (só dashboard) + colunas `rec_*` na tabela da Lotação; coordenadores tem versão enxuta filtrada por curso via `por_curso_status` no `_filtrarDBPorCursos`. |

> **Números sem comentário `PATCH N` em nenhum arquivo:** 31, 36, 46, 62, 66,
> 101, 128, 129, 143, 144. Podem ter sido patches de discussão sem comentário,
> renumerados, ou aplicados só no changelog externo. Não assuma que não existiram.

---

## 14. Armadilhas conhecidas

1. **Replicação nos dois templates.** `template_dashboard.html` e
   `template_coordenadores.html` têm as MESMAS funções com os MESMOS nomes.
   Quase toda correção de frontend precisa ir nos dois. Bugs recorrentes:
   PATCH 138 (corrigido só num), PATCH 118 (por isso foi movido pro backend).

2. **Funções auxiliares definidas DENTRO de uma função Python não existem fora
   dela.** `_norm_polo_ger` só existe dentro de
   `processar_gerenciamento_semestres`. O PATCH 147 chamou ela de dentro de
   `processar()` → `NameError` → caiu no `except` genérico do bloco de
   gerenciamento → `tem_gerenciamento=False` → **a aba inteira de Gerenciamento
   sumiu** (PATCH 149). Sempre confira o escopo antes de reusar um helper.

3. **Não existe UMA função de normalizar polo.** Há pelo menos 8 reimplementações
   quase-iguais (`_norm_polo_bf`, `_norm_polo_ger`, `_norm_polo_inj`,
   `_norm_polo_cruzamento`, `_norm_polo_contato`, `_norm_polo_labs`,
   `_norm_polo_hub_main`, `normPoloExp*` no JS). Elas divergem em detalhes (remove
   parênteses? remove chapa? remove acento?). Ao mexer numa, verifique se o bug
   não está em outra. O "padrão canônico" é: tira `LAP -`, tira `(...)`, tira
   acento, colapsa espaço, lowercase (PATCH 45).

4. **Idem para "os nomes de tutor batem?"** — `_nomes_batem_match`,
   `_nomes_batem_ch`, `_nomes_batem_inj`, `_nomes_batem_ger2`, `_nomesBatemCanon`
   (JS), `_nomesBatem` (JS)... cascata igual (exato → primeiro+último →
   subsequência de palavras), implementações separadas.

5. **`processar_gerenciamento_csv(p5)` (`processar.py:2667`) é código morto** —
   nunca é chamado. O parâmetro se chama `p5` mas o `p5` real do pipeline é o
   **hub CSV de alunos** (`carregar_alunos_hub`), não gerenciamento. Não
   confunda. `REL_DETALHADO.csv` (`URL_REL_NOVO`) é baixado pelo Actions e também
   não é consumido.

6. **GIOCONDA infla contagem de alunos** — o mesmo aluno aparece 1x por prática.
   Qualquer soma de alunos vinda do GIOCONDA precisa deduplicar por polo×categoria
   dentro da ordem primeiro (PATCH 38 backend, PATCH 48/114 frontend). Quando o
   hub CSV existe, ele é a fonte de verdade e substitui essas contagens.

7. **`_OFE` / `_PESO_OFERTA`**: expandir linhas pela capacidade multiplica em
   dobro se a mesma prática aparecer em várias linhas — deduplicar por prática
   antes de expandir (PATCH 150).

8. **Duplicação de linha nos exports do GIOCONDA**: o export pode trazer o nome do
   tutor com e sem sufixo de chapa `(12345)` na mesma pessoa (PATCH 26/115) —
   sempre normalizar removendo `\s*\(\d+\)\s*$` e preferir a variante COM chapa.

9. **`_SEMESTRE_ORIGEM` ≠ data da submissão.** É a origem do arquivo (PATCH 127).
   Uma submissão de 2026/2 feita com data de janeiro continua sendo `'2026/2'`.

10. **Cruzamento Portfólio×Agendado é sempre por semestre.** Comparar histórico
    somado (2026/1+2026/2) contra agendados de um semestre dá percentuais
    impossíveis (>100%, >600%) (PATCH 123/133).

11. **`insumos_estudo.json` não está no repo agora.** A aba Insumos fica vazia até
    alguém commitar o arquivo. É estático por decisão — não tem fonte que atualiza.

12. **`index.html` / `coordenadores.html` na raiz são SAÍDA gerada.** Editar essas
    (~43 MB cada!) à mão é inútil — o próximo ciclo do Actions sobrescreve. Edite
    `template_*.html`.

13. **Restos mortos no coordenadores**: `_renderVagasKPIs_ORIGINAL_NAO_USAR`,
    `_renderVagasTbl_ORIGINAL_NAO_USAR`. Não são chamados; não replique bugfix neles.

14. **`except` genéricos engolem erro e escondem seção.** Vários blocos de
    `processar()` (`if p4:`, `if p6:`, `if p3 or p3b:`) capturam qualquer exceção
    e seguem com a seção vazia/desativada. Um erro que "some com uma aba" no
    dashboard quase sempre é uma exceção silenciada aqui — rode local e olhe os
    prints `[AVISO]` / `[ERRO]`.
