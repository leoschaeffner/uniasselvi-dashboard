# VinciLab — guia de eficiência para o Claude Code

Leia isto antes de qualquer tarefa. Objetivo: gastar menos tokens sem perder qualidade.

## Ordem de leitura (nunca ler processar.py "a frio")
1. **`MAPA_DO_CODIGO.md`** primeiro — sempre. Responde "onde fica X" sem carregar
   `processar.py` (~5000 linhas) ou os templates (~5100 linhas cada) inteiros.
2. Se a dúvida for "onde fica X / por que Y foi feito assim", chame o agente
   **`guia-codigo`** em vez de grepar por conta própria.
3. Só leia o arquivo grande com `offset`/`limit` na região relevante (já
   localizada pelo MAPA ou pelo `guia-codigo`), nunca o arquivo inteiro.

## Agentes — use-os, não reinvente o que eles já sabem
- **`git-github-vinci`**: qualquer pull/push/PR/deploy/disparo de workflow. Ele
  já sabe que `git push`/`fetch` podem travar neste repo, o padrão de PR via
  API do GitHub, e o id do workflow. Não redescubra isso investigando do zero.
- **`investigador-bugs`**: antes de tentar corrigir algo "estranho" no dado,
  isole a causa primeiro (evita 2-3 rodadas de tentativa e erro).
- **`dados-etl`** / **`frontend-vinci`**: mudança em `processar.py` ou nos
  templates. Eles já commitam e abrem PR — não repita esse fluxo manualmente.
- **`qa-vinci`**: antes de qualquer commit em `processar.py`/templates.

## Auditoria de dado publicado sem planilha em mãos
Não decifre `index.html` (43 MB) e faça `print`/dump gigante no terminal.
Use `action: "read"` só se necessário, ou o script já documentado na memória
**`auditar-db-publicado`** — decifra pra um arquivo e faça as consultas em
Python com `json.load` + `print` de **campos específicos**, nunca do dict
inteiro. Isso sozinho é a maior fonte de gasto de token deste projeto.

## Rodar o pipeline local
`python processar.py --sem-browser` demora ~1min e imprime muito log. Filtre
sempre: `| grep -E "Tutores:|Vagas RH:|Concluído|ERRO|Traceback"` (ou o
subconjunto relevante à mudança) em vez de deixar o log inteiro no contexto.
Não rode de novo "só para conferir" — bateje várias mudanças de código e rode
uma vez só antes de comitar.

## Scratchpad
Arquivos de teste/debug vão em `scratchpad/`, nunca na raiz do repo. Limpe
arquivos grandes (`*.json`/`*.html` de debug) depois de usar — já aconteceu de
acumular ~2 GB de dumps de depuração numa sessão. `rm -rf` está bloqueado no
projeto — apague arquivo por arquivo ou peça pro usuário.

## Permissões
`.claude/settings.local.json` já libera `Bash` com `defaultMode: acceptEdits`
— não é preciso confirmar cada comando. `git push --force` e `rm -rf` seguem
bloqueados de propósito (usar a API do GitHub pra isso, via `git-github-vinci`).

## Regra de ouro do frontend
Quase toda mudança visual precisa ir nos DOIS templates
(`template_dashboard.html` e `template_coordenadores.html`). Esquecer um dos
dois é a causa nº1 de bug aqui — sempre grepe o nome da função/elemento nos
dois antes de considerar terminado.

## Trabalho autônomo
O Leo prefere que eu **execute e mostre o resultado**, não que eu pare pra
perguntar a cada passo (ver memória `feedback-autonomia`). Reservar pergunta
só pra decisão genuinamente irreversível e ambígua.
