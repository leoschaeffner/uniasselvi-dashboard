#!/usr/bin/env python3
"""Gera vinculo_laboratorios_real.json a partir das bases de BI reais em
`planilhas/Arquivos BI/` (Alunos3399_Geral, Dim_Polo, cidades_geocoded,
Fato_Distancias_4_mais_proximas, Laboratorios_Tratado).

Uso:
  python tests/tools/gerar_vinculo_laboratorios.py [pasta_arquivos_bi] [--saida arq.json] [--dry-run]

Default da pasta: planilhas/Arquivos BI (relativo à raiz do repo).

NUNCA lê as colunas ALUNO / MATRICULA / "TUTOR DE PRÁTICA" (PII) — nem no
`usecols` do Alunos3399_Geral.

Regras de negócio (decisões do Leo, confirmadas no pedido, NÃO inferidas):
- "Polo Dependente" conta como HUB resolvido (junto com "Polo HUB"). Só
  "Sem HUB" cai em `pendentes[]`.
- `urgencia = alunos_sem_pratica*3 + alunos_com_pratica*1` (SEM bônus de
  recorrência — a etiqueta `recorrente` é só informativa).
- Candidato HUB com `LABORATORIO_IMPLANTADO=="Não"` continua sendo oferecido
  (não é excluído), só leva a flag `lab_implantado=False` pro frontend avisar.

Agrupamento: por (CODIGO_POLO, CURSO) — CURSO aqui é a coluna real de
Alunos3399_Geral/Laboratorios_Tratado (nome do curso/graduação, ex.
"Biomedicina", "Engenharia Elétrica"), NÃO o rótulo de laboratório amplo
(Multidisciplinar I/II/III/IV) usado em labs_pendencias.json — são fontes e
domínios diferentes; confirmado lendo as duas planilhas reais (ambas usam
CURSO = nome de graduação).
"""
import argparse
import json
import math
import os
import sys
from collections import Counter

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
RAIO_VIAVEL_KM = 150  # ver relatório final para distribuição medida de km

# Colunas de Alunos3399_Geral que entram em memória. NUNCA inclua
# ALUNO / MATRICULA / "TUTOR DE PRÁTICA" aqui.
_COLS_ALUNOS = ['CODIGO_POLO', 'CURSO', 'SEMESTRE', 'SITUACAO_PRATICA', 'STATUS_HUB']


def _log(msg):
    print(msg, flush=True)


def _achar_pasta_bi(arg):
    if arg:
        return arg
    return os.path.join(RAIZ, 'planilhas', 'Arquivos BI')


def _ler_alunos(pasta):
    import pandas as pd
    caminho = os.path.join(pasta, 'Alunos3399_Geral.xlsx')
    df = pd.read_excel(caminho, usecols=_COLS_ALUNOS)
    # Sanidade: garante que nenhuma coluna de PII foi lida por engano.
    proibidas = {'ALUNO', 'MATRICULA', 'TUTOR DE PRÁTICA', 'TUTOR DE PR�TICA'}
    if proibidas & set(df.columns):
        raise SystemExit(f'ERRO FATAL: coluna de PII detectada em memória: {proibidas & set(df.columns)}')
    return df


def _ler_dim_polo(pasta):
    import pandas as pd
    df = pd.read_excel(os.path.join(pasta, 'Dim_Polo.xlsx'))
    return {int(r.CODIGO_POLO): {
        'nome_polo': r.NOME_POLO, 'cidade': r.CIDADE, 'uf': r.UF,
        'cidade_uf_key': r.CIDADE_UF_KEY,
    } for r in df.itertuples()}


def _ler_geocode(pasta):
    import pandas as pd
    df = pd.read_csv(os.path.join(pasta, 'cidades_geocoded.csv'))
    return {r.CIDADE_UF_KEY: (float(r.LATITUDE), float(r.LONGITUDE)) for r in df.itertuples()}


def _ler_distancias(pasta):
    import pandas as pd
    df = pd.read_excel(os.path.join(pasta, 'Fato_Distancias_4_mais_proximas.xlsx'))
    out = {}
    for r in df.itertuples():
        out[(r.CIDADE_ORIGEM_KEY, r.CIDADE_DESTINO_KEY)] = (float(r.DISTANCIA_RODOVIARIA_KM), float(r.DURACAO_MIN))
    return out


def _ler_laboratorios_tratado(pasta):
    import pandas as pd
    df = pd.read_excel(os.path.join(pasta, 'Laboratorios_Tratado.xlsx'))
    return df


def _haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def _moda_status(serie):
    c = Counter(serie)
    (valor, _n), = c.most_common(1)
    return valor, len(c) > 1


def _agrupar_semestre(df, semestre):
    """Agrupa (CODIGO_POLO, CURSO) dentro de um semestre. Retorna dict
    {(cod_polo, curso): {...}} + contador de grupos com STATUS_HUB misto."""
    sub = df[df['SEMESTRE'] == semestre]
    grupos = {}
    mistos = 0
    # groupby é bem mais rápido que iterrows em 240k linhas.
    g = sub.groupby(['CODIGO_POLO', 'CURSO'], sort=False)
    for (cod_polo, curso), gdf in g:
        status_moda, misto = _moda_status(gdf['STATUS_HUB'])
        if misto:
            mistos += 1
        sem_pratica = int((gdf['SITUACAO_PRATICA'] == 'Sem Prática').sum())
        total = int(len(gdf))
        grupos[(int(cod_polo), curso)] = {
            'total_alunos': total,
            'alunos_sem_pratica': sem_pratica,
            'alunos_com_pratica': total - sem_pratica,
            'status_hub': status_moda,
        }
    return grupos, mistos


def _candidatos_por_curso(df_lab, semestre):
    """{curso: [{cod_polo, lab_implantado}]} só HUBs (STATUS_HUB=='Polo HUB')
    do semestre pedido. Dedup por (cod_polo, curso); lab_implantado = True só
    se TODAS as linhas LABORATORIO daquele polo+curso estiverem implantadas
    (decisão de engenharia: um polo pode ter múltiplos LABORATORIO por curso,
    ex. ENGMAKER + QUÍMICA E FÍSICA — reporta pendência se qualquer um faltar)."""
    sub = df_lab[(df_lab['SEMESTRE'] == semestre) & (df_lab['STATUS_HUB'] == 'Polo HUB')]
    por_curso = {}
    for curso, gdf in sub.groupby('CURSO', sort=False):
        agregado = {}
        for cod_polo, gg in gdf.groupby('CODIGO_POLO'):
            implantado = bool((gg['LABORATORIO_IMPLANTADO'] == 'Sim').all())
            agregado[int(cod_polo)] = implantado
        por_curso[curso] = agregado
    return por_curso


def gerar(pasta, log=_log):
    diag = {}
    log('Lendo Alunos3399_Geral.xlsx (usecols restrito, sem PII)...')
    df = _ler_alunos(pasta)
    log(f'  {len(df)} linhas lidas.')

    semestres = sorted(df['SEMESTRE'].unique())
    if len(semestres) < 2:
        semestre_atual = semestres[-1]
        semestre_anterior = None
        log(f'AVISO: só 1 semestre encontrado ({semestre_atual}) — recorrência sempre False.')
    else:
        semestre_atual = max(semestres)
        semestre_anterior = [s for s in semestres if s != semestre_atual][-1]
    log(f'semestre_atual={semestre_atual!r} semestre_anterior={semestre_anterior!r}')

    log('Agrupando semestre atual por (CODIGO_POLO, CURSO)...')
    grupos_atual, mistos_atual = _agrupar_semestre(df, semestre_atual)
    diag['grupos_totais'] = len(grupos_atual)
    diag['grupos_status_misto'] = mistos_atual
    if mistos_atual:
        log(f'AVISO: {mistos_atual} grupo(s) com STATUS_HUB misto no semestre atual — resolvido por moda.')

    grupos_pend_ant = set()
    if semestre_anterior:
        log('Agrupando semestre anterior (só para recorrência)...')
        grupos_ant, mistos_ant = _agrupar_semestre(df, semestre_anterior)
        if mistos_ant:
            log(f'AVISO: {mistos_ant} grupo(s) com STATUS_HUB misto no semestre anterior — resolvido por moda.')
        grupos_pend_ant = {k for k, v in grupos_ant.items() if v['status_hub'] == 'Sem HUB'}

    del df  # libera memória cedo — planilha grande

    log('Lendo Dim_Polo.xlsx / cidades_geocoded.csv / Fato_Distancias / Laboratorios_Tratado...')
    dim_polo = _ler_dim_polo(pasta)
    geocode = _ler_geocode(pasta)
    distancias = _ler_distancias(pasta)
    df_lab = _ler_laboratorios_tratado(pasta)
    candidatos_curso = _candidatos_por_curso(df_lab, semestre_atual)

    def _enriquecer(cod_polo):
        info = dim_polo.get(cod_polo)
        if not info:
            return None
        lat_lon = geocode.get(info['cidade_uf_key'])
        return {
            'nome_polo': info['nome_polo'], 'cidade': info['cidade'], 'uf': info['uf'],
            'cidade_uf_key': info['cidade_uf_key'],
            'lat': lat_lon[0] if lat_lon else None,
            'lon': lat_lon[1] if lat_lon else None,
        }

    polos_sem_dim = set()
    cidades_sem_geo = set()

    pendentes, resolvidos = [], []
    for (cod_polo, curso), g in grupos_atual.items():
        info = _enriquecer(cod_polo)
        if info is None:
            polos_sem_dim.add(cod_polo)
            continue
        if info['lat'] is None:
            cidades_sem_geo.add(info['cidade_uf_key'])

        if g['status_hub'] == 'Sem HUB':
            recorrente = (cod_polo, curso) in grupos_pend_ant
            urgencia = g['alunos_sem_pratica'] * 3 + g['alunos_com_pratica'] * 1

            # candidatos: HUBs do mesmo curso
            cands_raw = []
            for cand_polo, lab_implantado in (candidatos_curso.get(curso) or {}).items():
                if cand_polo == cod_polo:
                    continue
                cinfo = _enriquecer(cand_polo)
                if cinfo is None:
                    continue
                km = mn = None
                fonte = None
                if info['cidade_uf_key'] == cinfo['cidade_uf_key']:
                    km, mn, fonte = 0.0, 0.0, 'mesma_cidade'
                elif (info['cidade_uf_key'], cinfo['cidade_uf_key']) in distancias:
                    km, mn = distancias[(info['cidade_uf_key'], cinfo['cidade_uf_key'])]
                    fonte = 'base_4_proximas'
                elif info['lat'] is not None and cinfo['lat'] is not None:
                    km = _haversine_km(info['lat'], info['lon'], cinfo['lat'], cinfo['lon']) * 1.25
                    mn = km / (70 / 60)  # 70km/h
                    fonte = 'estimado_linha_reta'
                if km is None:
                    continue  # sem coordenada em nenhum dos dois lados — não dá pra estimar
                cands_raw.append({
                    'cod_polo': f'POLO{cand_polo}', 'nome_polo': cinfo['nome_polo'],
                    'cidade': cinfo['cidade'], 'uf': cinfo['uf'],
                    'lat': cinfo['lat'], 'lon': cinfo['lon'],
                    'km': round(km, 1), 'min': round(mn, 0),
                    'fonte_distancia': fonte, 'lab_implantado': lab_implantado,
                })
            cands_raw.sort(key=lambda c: c['km'])
            candidatos = cands_raw[:3]
            tem_candidato_viavel = any(c['km'] <= RAIO_VIAVEL_KM for c in candidatos)

            pendentes.append({
                'cod_polo': f'POLO{cod_polo}', 'nome_polo': info['nome_polo'], 'curso': curso,
                'cidade': info['cidade'], 'uf': info['uf'], 'lat': info['lat'], 'lon': info['lon'],
                'total_alunos': g['total_alunos'], 'alunos_sem_pratica': g['alunos_sem_pratica'],
                'alunos_com_pratica': g['alunos_com_pratica'], 'urgencia': urgencia,
                'recorrente': recorrente, 'tem_candidato_viavel': tem_candidato_viavel,
                'candidatos': candidatos,
            })
        else:
            resolvidos.append({
                'cod_polo': f'POLO{cod_polo}', 'nome_polo': info['nome_polo'], 'curso': curso,
                'cidade': info['cidade'], 'uf': info['uf'], 'total_alunos': g['total_alunos'],
            })

    pendentes.sort(key=lambda p: -p['urgencia'])

    # KPIs — nomes de campo EXATAMENTE como template_laboratorios.html já lê
    # (montarVincKPIs): polos_sem_hub (na verdade combinações polo+curso, o
    # rótulo do frontend já deixa isso explícito), matriculas_afetadas,
    # matriculas_sem_pratica, sem_candidato_viavel, total_hubs_ativos,
    # total_cursos_com_hub.
    polos_unicos_pend = {p['cod_polo'] for p in pendentes}
    cursos_pend = {p['curso'] for p in pendentes}
    # PATCH (correção de bug do protótipo): total_hubs_ativos antes contava
    # combinações polo+curso com STATUS_HUB=='Polo HUB' (inflava o número —
    # um polo HUB de várias práticas era contado várias vezes). Aqui conta
    # POLOS ÚNICOS com pelo menos 1 STATUS_HUB=='Polo HUB' no semestre atual.
    polos_hub_ativos = {cod_polo for (cod_polo, _curso), g in grupos_atual.items() if g['status_hub'] == 'Polo HUB'}
    cursos_com_hub_ativo = {curso for (_cod_polo, curso), g in grupos_atual.items() if g['status_hub'] == 'Polo HUB'}

    kpis = {
        'polos_sem_hub': len(pendentes),
        'polos_unicos_sem_hub': len(polos_unicos_pend),
        'matriculas_afetadas': sum(p['total_alunos'] for p in pendentes),
        'matriculas_sem_pratica': sum(p['alunos_sem_pratica'] for p in pendentes),
        'sem_candidato_viavel': sum(1 for p in pendentes if not p['tem_candidato_viavel']),
        'total_hubs_ativos': len(polos_hub_ativos),
        'total_cursos_com_hub': len(cursos_com_hub_ativo),
        'cursos': len(cursos_pend),
    }

    # Diagnóstico de km dos candidatos (pro threshold de 150km — não decide, só mede)
    todos_km = [c['km'] for p in pendentes for c in p['candidatos']]
    diag['candidatos_km'] = todos_km
    diag['polos_sem_dim_polo'] = polos_sem_dim
    diag['cidades_sem_geocode'] = cidades_sem_geo
    diag['recorrentes'] = sum(1 for p in pendentes if p['recorrente'])
    diag['pendentes_n'] = len(pendentes)
    diag['resolvidos_n'] = len(resolvidos)
    diag['polos_unicos_pend'] = len(polos_unicos_pend)

    saida = {
        'gerado_em': __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M'),
        'semestre_atual': semestre_atual,
        'semestre_anterior': semestre_anterior,
        'raio_viavel_km': RAIO_VIAVEL_KM,
        'kpis': kpis,
        'pendentes': pendentes,
        'resolvidos': resolvidos,
        'demo': False,
        'recorrencia_real': True,
        'fonte': 'Arquivos BI — Alunos3399_Geral/Dim_Polo/Laboratorios_Tratado/Fato_Distancias/cidades_geocoded',
    }
    return saida, diag


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pasta', nargs='?')
    ap.add_argument('--saida', default=os.path.join(RAIZ, 'vinculo_laboratorios_real.json'))
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    pasta = _achar_pasta_bi(a.pasta)
    if not os.path.isdir(pasta):
        raise SystemExit(f'ERRO: pasta não encontrada: {pasta}')

    saida, diag = gerar(pasta)

    log = _log
    log('')
    log('=== RELATÓRIO ===')
    log(f"Combinações polo+curso pendentes (Sem HUB): {diag['pendentes_n']}")
    log(f"Polos únicos pendentes: {diag['polos_unicos_pend']}")
    log(f"Resolvidos (Polo HUB + Polo Dependente): {diag['resolvidos_n']}")
    log(f"Grupos com STATUS_HUB misto (resolvido por moda): {diag['grupos_status_misto']}")
    pct_rec = 100.0 * diag['recorrentes'] / diag['pendentes_n'] if diag['pendentes_n'] else 0.0
    log(f"Recorrência real: {diag['recorrentes']}/{diag['pendentes_n']} = {pct_rec:.1f}%")
    log(f"Polos sem Dim_Polo (excluídos): {len(diag['polos_sem_dim_polo'])} {sorted(diag['polos_sem_dim_polo'])[:10]}")
    log(f"Cidades sem geocode: {len(diag['cidades_sem_geocode'])}")
    kms = sorted(diag['candidatos_km'])
    if kms:
        n = len(kms)
        mediana = kms[n // 2] if n % 2 else (kms[n // 2 - 1] + kms[n // 2]) / 2
        pct_ate_150 = 100.0 * sum(1 for k in kms if k <= RAIO_VIAVEL_KM) / n
        log(f"Distribuição de km dos candidatos (n={n}): mediana={mediana:.1f}km · "
            f"<=150km: {pct_ate_150:.1f}% · >150km: {100 - pct_ate_150:.1f}% · "
            f"min={kms[0]:.1f} max={kms[-1]:.1f}")
    log(f"KPIs: {json.dumps(saida['kpis'], ensure_ascii=False)}")

    if a.dry_run:
        log('--dry-run: nada gravado.')
        return
    with open(a.saida, 'w', encoding='utf-8') as f:
        json.dump(saida, f, ensure_ascii=False)
    log(f'Gravado: {a.saida}')


if __name__ == '__main__':
    main()
