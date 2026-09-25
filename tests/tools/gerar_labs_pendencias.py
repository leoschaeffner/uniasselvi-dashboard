#!/usr/bin/env python3
"""Gera labs_pendencias.json a partir da planilha mensal "PENDENCIAS DE IMPLANTACAO".

Uso:
  python tests/tools/gerar_labs_pendencias.py [planilha.xlsx] [--saida arq.json] [--dry-run]

Default da planilha: glob planilhas/PEND*IMPLANTA*.xlsx. Nunca grava o nome do tutor
(so o booleano `tutor_contratado`). Falha em cabecalho alterado ou HUB desconhecido.
"""
import argparse, glob, json, os, re, sys, unicodedata
from collections import Counter, OrderedDict

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
CABECALHOS = ['PARCEIRO', 'POLO', 'HUB', 'STATUS', 'OBS.:', 'TUTOR CONTRATADO', 'TOTAL ALUNOS']

# HUB da planilha -> rotulo de categoria usado em t['c'] dos tutores.
HUB_PARA_CATEGORIA = {
    'Multidisciplinar I': 'BIO-FAR (Multidisciplinar I)',
    'Multidisciplinar II': 'ENF-INS (Multidisciplinar II)',
    'Multidisciplinar III': 'BIO-FISIO-EST-TO (Multidisciplinar III)',
    'Multidisciplinar IV': 'NUTRI (Multidisciplinar IV)',
    'Multi IV': 'NUTRI (Multidisciplinar IV)',  # grafia inconsistente da planilha
    'ENGMAKER': 'ENGMAKER',
    'ENGMAKER+QUÍMICA E FÍSICA': 'ENGMAKER+QUÍMICA E FÍSICA',
    'EXATAS': 'QUÍMICA E FÍSICA',
}

# Override de categoria (decisao do Leo, 25/09/2026). Chave: (polo_norm, HUB original da planilha).
# BH Shopping Estacao: a planilha traz "Multi IV" com OBS "Recebeu o lab e nao implantou, estao vendo
# nova sala. No sistema ja nao consta mais como HUB Multi IV" -> vale para o polo inteiro (categoria vazia),
# como na base antiga, em vez de NUTRI (Multidisciplinar IV).
CATEGORIA_OVERRIDE = {
    ('belo horizonte/mg - shopping estacao bh', 'Multi IV'): '',
}

# Substituicoes de texto no motivo (anonimiza nome de terceiro), aplicadas apos o corte.
SUBSTITUICOES_MOTIVO = [
    ('Bruno passou a informação pelo teams', 'informação passada pelo Teams'),
]


def norm_polo(s):
    """Copia fiel de _norm_polo_labs (funcao aninhada em processar() de processar.py)."""
    s = str(s or '').replace('\xa0', ' ').strip()
    s = re.sub(r'^LAP\s*[-–]\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\([^)]*\)', '', s)
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip().lower()


CONTADORES = Counter()

_RE_CORTE = re.compile(r'\s*[-–]*\s*[úu]ltima\s+atualiza[çc][ãa]o.*$', re.IGNORECASE | re.DOTALL)

_RE_VISITA = re.compile(r'\s*-+\s*Visita sendo agendada pelo analista.*$', re.IGNORECASE | re.DOTALL)


def limpar_motivo(obs):
    s = str(obs or '').replace('\xa0', ' ').strip()
    s = _RE_CORTE.sub('', s).strip()
    s = _RE_VISITA.sub('', s).strip()
    for de, para in SUBSTITUICOES_MOTIVO:
        if de in s:
            s = s.replace(de, para)
            CONTADORES['substituicoes'] += 1
    return s


def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ''
    return str(v).replace('\xa0', ' ').strip()


def ler_planilha(caminho):
    import pandas as pd
    df = pd.read_excel(caminho, dtype=object)
    cols = [str(c).strip() for c in df.columns]
    if cols != CABECALHOS:
        raise SystemExit(f'ERRO: cabecalhos da planilha mudaram.\n  esperado: {CABECALHOS}\n  encontrado: {cols}')
    df.columns = cols
    return df


def gerar(df):
    saida, vistos, dups, hubs = [], set(), 0, Counter()
    lidas = 0
    for _, r in df.iterrows():
        polo = _txt(r['POLO'])
        if not polo:
            continue
        lidas += 1
        hub = _txt(r['HUB'])
        if hub not in HUB_PARA_CATEGORIA:
            raise SystemExit(f'ERRO: HUB desconhecido {hub!r} (polo {polo!r}). Atualize HUB_PARA_CATEGORIA.')
        cat = HUB_PARA_CATEGORIA[hub]
        pn = norm_polo(polo)
        if (pn, hub) in CATEGORIA_OVERRIDE:
            cat = CATEGORIA_OVERRIDE[(pn, hub)]
            CONTADORES['overrides'] += 1
            print(f'  override de categoria aplicado: {pn!r} HUB {hub!r} -> {cat!r}')
        hubs[cat] += 1
        tutor = _txt(r['TUTOR CONTRATADO'])
        total = r['TOTAL ALUNOS']
        try:
            total = None if total is None or total != total or _txt(total) == '' else int(float(total))
        except (ValueError, TypeError):
            total = None
        reg = OrderedDict([
            ('polo_norm', pn),
            ('polo_original', polo),
            ('categoria', cat),
            ('status', _txt(r['STATUS'])),
            ('motivo', limpar_motivo(r['OBS.:'])),
            ('empresa', _txt(r['PARCEIRO'])),
            ('total_alunos', total),
            ('tutor_contratado', bool(tutor) and not tutor.lower().startswith('sem tutor')),
        ])
        chave = (reg['polo_norm'], reg['categoria'], reg['status'], reg['motivo'])
        if chave in vistos:
            dups += 1
            continue
        vistos.add(chave)
        saida.append(reg)
    return saida, {'lidas': lidas, 'escritas': len(saida), 'dups': dups,
                   'polos': len({x['polo_norm'] for x in saida}), 'hubs': dict(hubs)}


def serializar(regs):
    # Formato identico ao arquivo historico: indent=2, ensure_ascii=False, CRLF, newline final.
    return json.dumps(regs, ensure_ascii=False, indent=2).replace('\n', '\r\n') + '\r\n'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('planilha', nargs='?')
    ap.add_argument('--saida', default=os.path.join(RAIZ, 'labs_pendencias.json'))
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    caminho = a.planilha
    if not caminho:
        achados = sorted(glob.glob(os.path.join(RAIZ, 'planilhas', 'PEND*IMPLANTA*.xlsx')))
        if not achados:
            raise SystemExit('ERRO: nenhuma planilha planilhas/PEND*IMPLANTA*.xlsx encontrada.')
        caminho = achados[0]
    regs, res = gerar(ler_planilha(caminho))
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    print(f"Linhas lidas: {res['lidas']} | escritas: {res['escritas']} | dedups: {res['dups']} | polos: {res['polos']}")
    print(f"Overrides de categoria aplicados: {CONTADORES['overrides']} | substituicoes de motivo aplicadas: {CONTADORES['substituicoes']}")
    for k, v in sorted(res['hubs'].items()):
        print(f'  {k}: {v}')
    if a.dry_run:
        print('--dry-run: nada gravado.')
        return
    with open(a.saida, 'w', encoding='utf-8', newline='') as f:
        f.write(serializar(regs))
    print(f'Gravado: {a.saida}')


if __name__ == '__main__':
    main()
