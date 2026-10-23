# Ferramentas de IA avaliadas para o VinciLab

Referência: post do LinkedIn de Lívio M., "10 repositórios do GitHub para
otimizar o uso da IA no trabalho". A lista foi conferida nos repositórios
oficiais em 23/10/2026.

Este arquivo não transforma dependências de desenvolvimento em dependências do
dashboard. Nenhuma destas ferramentas deve entrar no `processar.py`, no HTML
publicado ou no GitHub Actions sem uma necessidade concreta e uma revisão de
segurança.

## Adotadas no ambiente do Codex

| Ferramenta | Uso no VinciLab | Regra prática |
|---|---|---|
| [Ponytail](https://github.com/DietrichGebert/ponytail) | Reduzir mudanças, dependências e abstrações desnecessárias | Usar como princípio em correções e refatorações; nunca cortar validação, segurança, acessibilidade ou testes essenciais |
| [Impeccable](https://github.com/pbakaus/impeccable) | Auditar e refinar os portais existentes | Usar junto do `frontend-vinci`; preservar identidade, conteúdo e regras do dashboard; validar os dois templates e o render final |
| [Humanizer](https://github.com/blader/humanizer) | Melhorar textos, tooltips, rótulos e documentação | Não alterar números, regras de negócio ou significado durante a reescrita |
| [Diagram Design](https://github.com/cathrynlavery/diagram-design) | Produzir diagramas de fluxo, arquitetura e dados | Usar para documentação técnica; diagramas não substituem testes nem o `MAPA_DO_CODIGO.md` |

As quatro skills acima ficam no ambiente do Codex, fora do artefato publicado.
Elas não adicionam JavaScript, Python ou chamadas externas ao dashboard.

## Mantidas como referência

| Ferramenta | O que faz | Decisão para este projeto |
|---|---|---|
| [REA](https://github.com/morluto/rea) | Engenharia reversa de comportamento, aplicações e binários | Não instalar por padrão. Só usar quando houver autorização e uma necessidade real de analisar um sistema externo; não é ferramenta de manutenção rotineira do VinciLab |
| [Hallmark](https://github.com/Nutlope/hallmark) | Evita padrões visuais genéricos em páginas novas | Não instalar agora: sobrepõe grande parte do Impeccable e é orientado principalmente a páginas novas. Pode ser reavaliado num redesign explícito |
| [Onyx](https://github.com/onyx-dot-app/onyx) | Plataforma completa de busca e chat com IA sobre fontes corporativas | Não incorporar. É uma aplicação/infraestrutura própria, não uma skill leve, e não resolve um problema atual do pipeline estático |
| [UI/UX Pro Max](https://github.com/nextlevelbuilder/ui-ux-pro-max-skill) | Biblioteca de inteligência e padrões de UI/UX | Não instalar agora: há forte sobreposição com Impeccable. Consultar como referência quando faltar um padrão específico |
| [Last 30 Days](https://github.com/mvanhorn/last30days-skill) | Pesquisa recente em Reddit, X, YouTube, Hacker News e web | Não ativar no pipeline. Exige Python 3.12+ e algumas fontes dependem de tokens/APIs; usar somente em pesquisa pontual com credenciais fora do repositório |
| [Scrapling](https://github.com/D4Vinci/Scrapling) | Extração e crawling de páginas, inclusive com navegador | Não instalar por padrão. Exige Python/browser e amplia a superfície de segurança. Usar apenas quando uma fonte autorizada não tiver API ou exportação; nunca para contornar acesso ou termos de uso |

## Fluxo recomendado

1. Investigar a causa com os agentes e as fontes reais do projeto.
2. Aplicar o menor ajuste completo (Ponytail), respeitando as regras do
   `AGENTS.md` e do `MAPA_DO_CODIGO.md`.
3. Em alterações visuais, usar Impeccable como auditoria complementar e
   replicar a mudança nos dois templates quando aplicável.
4. Passar pelo `qa-vinci` antes do commit.
5. Usar Humanizer só na revisão da comunicação e Diagram Design quando um
   fluxo complexo realmente se beneficiar de um visual.

## Segurança e manutenção

- Skills e MCPs executam instruções e podem chamar programas locais. Atualize
  apenas a partir dos repositórios oficiais e revise mudanças antes de subir de
  versão.
- O Impeccable inclui um lançador que pode baixar um binário no primeiro uso.
  Esse download deve ser explícito e revisto; a instalação da skill, sozinha,
  não executa o lançador.
- Nunca gravar chaves, cookies ou tokens no repositório, em planilhas, logs ou
  HTML gerado.
- Não adicionar uma ferramenta ao GitHub Actions só para torná-la disponível ao
  agente local.
- Se duas skills cobrirem a mesma tarefa, usar uma só. Para este projeto,
  Impeccable é a referência visual principal.
- Ferramentas externas não prevalecem sobre as regras de negócio, testes e
  agentes especializados do VinciLab.
