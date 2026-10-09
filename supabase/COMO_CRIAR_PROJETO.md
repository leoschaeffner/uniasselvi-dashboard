# Como criar o projeto Supabase — passo a passo pro Leo

Isto é só o que falta pra destravar a Fase 1 de verdade (schema já desenhado
em `supabase/schema_fase1_dashboard.sql`, mapeamento em
`supabase/MAPEAMENTO_SCHEMA.md`). Nenhuma conta foi criada, nenhuma
credencial foi gerada ou pedida ainda — isso é só o roteiro pra você seguir
quando quiser.

## 1. Criar a conta / projeto

1. Acesse `https://supabase.com` e crie uma conta grátis (pode ser com login
   do GitHub, mais rápido).
2. Clique em **"New project"**.
3. **Nome sugerido do projeto:** `vincilab-dashboard` (ou `vincilab` se você
   preferir um projeto só que depois recebe os outros portais também —
   lembrete: o free tier só permite **2 projetos ativos simultâneos**, então
   vale decidir desde já se é "1 projeto pra tudo" ou "1 por portal"; a
   recomendação do plano é 1 projeto só, já que os 4 portais vão compartilhar
   o mesmo banco com RLS separando o acesso, não bancos separados).
4. **Organização:** crie uma organização nova ou use a pessoal — tanto faz
   pra este estágio.
5. **Senha do banco de dados (database password):** o Supabase pede uma
   senha mestra na criação do projeto (é a senha do usuário `postgres`
   nativo, diferente da chave de API). Gere uma senha forte (pode usar o
   gerador que o próprio formulário do Supabase oferece) e **guarde num
   gerenciador de senhas seu** — não vai precisar me passar essa agora, só
   mais adiante se formos usar o role customizado do §6.2 do plano.
6. **Região:** escolha **South America (São Paulo)** se aparecer disponível
   no plano free (o Supabase nem sempre libera todas as regiões no free tier
   — se São Paulo não aparecer na lista, a segunda melhor opção é a região
   dos EUA mais ao sul disponível, tipicamente `us-east-1`; evite Europa/Ásia
   pela latência com o Brasil).
7. **Plano:** Free.
8. Clique em criar — leva 1-2 minutos pro projeto provisionar.

## 2. O que fazer logo depois de criado (evitar a pausa por inatividade)

O plano free pausa o projeto após 1 semana sem uso. O projeto já tem um
padrão de "keepalive" pro GitHub Actions (commit `6d1cd8b`) — quando a Fase 1
entrar em implementação de verdade, replicamos esse mesmo padrão pra fazer um
ping programado no Supabase. Não precisa fazer nada manual agora, só não se
assustar se o projeto pausar antes disso estar configurado (reativar é
simples, só demora alguns minutos).

## 3. Credenciais que vou precisar depois (quando a Fase 1 entrar em código)

**Não me cole nenhuma delas direto no chat.** O caminho é sempre: você
configura como **GitHub Secret** no repositório, e só me avisa o **nome** do
secret (eu leio o valor de dentro do workflow, nunca preciso ver o valor em
texto no chat).

No painel do projeto Supabase, em **Project Settings**:

### a) Project URL e anon key (settings → API)
- **Project URL** (algo como `https://xxxxxxxx.supabase.co`) — essa pode
  inclusive ficar direto no código do frontend (`template_dashboard.html`),
  não é secreta.
- **anon key** (chave pública, começa com `eyJ...`) — também pode ficar no
  frontend. Ela só funciona dentro do que a RLS permitir, então não precisa
  virar GitHub Secret, mas se preferir centralizar, pode configurar como
  secret mesmo assim (ex: nome `SUPABASE_ANON_KEY`).
- **NUNCA** me passe nem configure em lugar nenhum a **service_role key**
  dessa mesma tela sem antes a gente decidir junto se vai usá-la (ver
  próximo item) — ela ignora toda RLS.

### b) Connection string do Postgres (pra testar o role customizado do §6.2 do plano)
1. Em **Project Settings → Database**, copie a **Connection string** no
   formato "URI" (tem um seletor de modo — use o modo **Session** ou
   **Transaction** pooler, não o direto, pra funcionar bem com scripts
   curtos tipo o `processar.py`).
2. **Antes de me passar qualquer coisa**, crie um GitHub Secret no
   repositório (`Settings > Secrets and variables > Actions > New repository
   secret`) com essa connection string (ela já inclui a senha do passo 1.5
   acima).
3. **Nome sugerido do secret:** `SUPABASE_DB_URL` (ou outro nome de sua
   preferência — só me avise qual você escolheu).
4. Quando formos testar o role customizado (`app_etl`, só INSERT/UPDATE nas
   tabelas do pipeline, sem acesso a mais nada — ver §6.2 do plano), vou
   pedir pra você rodar um `CREATE ROLE`/`GRANT` especifico via SQL Editor do
   próprio painel Supabase (interface web, sem precisar expor connection
   string nenhuma pra mim) — e aí sim criamos um SEGUNDO secret só com a
   connection string **desse role restrito**, que é o que realmente vai pro
   GitHub Actions. A connection string do usuário `postgres` master (passo
   acima) fica só pra você usar manualmente, nunca em automação.

### c) Se o role customizado não for viável (fallback do plano, §6.2)
Só se confirmarmos que o role restrito não funciona no free tier: aí sim
seria necessário usar a `service_role key` (em **Project Settings → API**,
seção "service_role secret") como GitHub Secret (nome sugerido:
`SUPABASE_SERVICE_ROLE_KEY`). **Isso só deve ser feito com todas as
mitigações do §6 do plano aplicadas** (nunca logar, plano de rotação
escrito) — não é o caminho padrão, é o plano B.

## 4. O que me avisar quando terminar

Só preciso que você me diga:
1. Que o projeto foi criado (nome e região escolhida).
2. O **nome do GitHub Secret** onde você guardou a connection string
   (ex: "configurei como `SUPABASE_DB_URL`").
3. Se quiser, o Project URL e a anon key direto no chat (essas não são
   secretas) — ou também como secret, se preferir manter tudo no mesmo
   lugar.

Não preciso ver senha nem connection string em texto em nenhum momento.
