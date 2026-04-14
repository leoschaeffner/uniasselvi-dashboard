"""
UNIASSELVI - Dashboard de Portfolios
Le as planilhas da pasta planilhas/ e gera saida/dashboard.html
"""

import pandas as pd
import json, os, sys, math, webbrowser, time, threading, glob
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# ── Localiza a pasta do script de 3 formas diferentes (a que funcionar serve) ─
def achar_pasta_script():
    candidatos = []

    # Forma 1: via sys.argv[0]
    try:
        p = os.path.dirname(os.path.abspath(sys.argv[0]))
        if os.path.isdir(p):
            candidatos.append(p)
    except:
        pass

    # Forma 2: via __file__
    try:
        p = os.path.dirname(os.path.abspath(__file__))
        if os.path.isdir(p):
            candidatos.append(p)
    except:
        pass

    # Forma 3: diretorio de trabalho atual
    try:
        p = os.getcwd()
        if os.path.isdir(p):
            candidatos.append(p)
    except:
        pass

    # Retorna o primeiro candidato que tenha a pasta planilhas ou o proprio script
    for p in candidatos:
        if os.path.isdir(os.path.join(p, "planilhas")):
            return p
        if os.path.isfile(os.path.join(p, "processar.py")):
            return p

    # Fallback: primeiro candidato
    return candidatos[0] if candidatos else os.getcwd()


SCRIPT_DIR = achar_pasta_script()

# ── Usar glob para achar os arquivos (ignora problemas de encoding/acento) ────
def ler_url_file(path_url):
    """Le um arquivo .url do Windows e extrai a URL do SharePoint."""
    try:
        with open(path_url, encoding='utf-8', errors='replace') as f:
            for line in f:
                if line.upper().startswith('URL='):
                    return line[4:].strip()
    except:
        pass
    return None


def forcar_download_onedrive(path_url_file, destino, label):
    """
    Forca o OneDrive a materializar um arquivo Files On-Demand (.url)
    como arquivo real (.xlsx) usando o comando attrib do Windows.
    Depois copia para o destino.
    """
    import subprocess, shutil, time

    # Caminho esperado do arquivo real (mesmo nome sem .url)
    path_xlsx = path_url_file.replace('.url', '').replace('.URL', '')

    # Força download via attrib (pina o arquivo no dispositivo)
    try:
        subprocess.run(
            ['attrib', '-P', '+U', path_url_file],
            capture_output=True, timeout=10
        )
    except Exception:
        pass

    # Aguarda até 30s para o OneDrive sincronizar
    for _ in range(6):
        if os.path.isfile(path_xlsx):
            with open(path_xlsx, 'rb') as f:
                header = f.read(4)
            if header == b'PK\x03\x04':
                shutil.copy2(path_xlsx, destino)
                print(f"  [OneDrive] Sincronizado: {label}")
                return destino
        time.sleep(5)
        print(f"  [OneDrive] Aguardando sync do OneDrive para {label}...")

    print(f"  [OneDrive] Timeout aguardando {label}. Verifique a conexao.")
    return None


# Palavras-chave para identificar cada planilha pelo nome
_KEYWORDS = {
    '01_CONTROLE_TUTORIA.xlsx': ['CONTROLE'],
    'PORTFOLIO_TUTOR.xlsx':     ['PORTFOLIO', 'PORTIFOLIO', 'PORTF'],
}

# Nomes reais no OneDrive (para busca por .url)
_ONEDRIVE_NAMES = {
    '01_CONTROLE_TUTORIA.xlsx': ['CONTROLE'],
    'PORTFOLIO_TUTOR.xlsx':     ['PORTF', 'PORTFOLIO'],
}


def _bate(caminho_arq, padrao):
    """Verifica se o arquivo corresponde ao padrao por palavras-chave."""
    bn  = os.path.basename(caminho_arq).upper()
    kws = _KEYWORDS.get(padrao, [os.path.splitext(padrao)[0].upper()])
    return any(kw in bn for kw in kws)


def achar_arquivo(pasta, padrao):
    """
    Localiza o arquivo Excel por:
    1. Nome exato na pasta planilhas
    2. Palavra-chave na pasta planilhas (.xlsx real)
    3. Arquivo .url na pasta planilhas -> baixa via PowerShell
    4. Arquivo .url no OneDrive local  -> baixa via PowerShell
    5. .xlsx real no OneDrive local
    """
    pasta_planilhas = os.path.join(pasta, "planilhas")

    # 1. Nome exato
    direto = os.path.join(pasta_planilhas, padrao)
    if os.path.isfile(direto):
        return direto

    # 2. .xlsx real por palavra-chave na pasta planilhas
    for arq in glob.glob(os.path.join(pasta_planilhas, "*.xls")) + glob.glob(os.path.join(pasta_planilhas, "*.xlsx")):
        if _bate(arq, padrao):
            return arq

    # 3. .url na pasta planilhas -> baixar via PowerShell
    for arq in glob.glob(os.path.join(pasta_planilhas, "*.url")) + glob.glob(os.path.join(pasta_planilhas, "*.xlsx.url")) + glob.glob(os.path.join(pasta_planilhas, "*.xls.url")):
        if _bate(arq, padrao):
            url = ler_url_file(arq)
            if url:
                destino = os.path.join(pasta_planilhas, padrao)
                resultado = forcar_download_onedrive(arq, destino, padrao)
                if resultado:
                    return resultado

    # 4 e 5. Busca no OneDrive local
    usuario = os.environ.get('USERNAME', os.environ.get('USER', 'leona'))
    for base in [
        f"C:\\Users\\{usuario}\\OneDrive - Uniasselvi",
        f"C:\\Users\\{usuario}\\OneDrive - UNIASSELVI",
        f"C:\\Users\\{usuario}\\OneDrive - Grupo Uniasselvi",
        f"C:\\Users\\{usuario}\\OneDrive",
    ]:
        if not os.path.isdir(base):
            continue

        # .url no OneDrive -> baixar
        for arq in glob.glob(os.path.join(base, "*.url")):
            if _bate(arq, padrao):
                url = ler_url_file(arq)
                if url:
                    destino = os.path.join(pasta_planilhas, padrao)
                    resultado = forcar_download_onedrive(arq, destino, padrao)
                    if resultado:
                        return resultado

        # .xlsx real no OneDrive
        for arq in glob.glob(os.path.join(base, "**", "*.xls"), recursive=True) + glob.glob(os.path.join(base, "**", "*.xlsx"), recursive=True):
            if _bate(arq, padrao):
                return arq

    return None

WATCH_MODE = len(sys.argv) > 1 and sys.argv[1].lower() == "watch"

CAT_MAP = {
    'ENF-INS (Multidisciplinar II)':
        'Multidisciplinar II - Enfermagem e Instrumentação Cirúrgica',
    'BIO-FISIO-EST-TO (Multidisciplinar III)':
        'Multidisciplinar III - Biomedicina Estética, Fisioterapia, Terapia Ocupacional e Estética e Cosmética',
    'BIO-BIO-FISIO-EST-TO (Multidisciplinar III)':
        'Multidisciplinar III - Biomedicina Estética, Fisioterapia, Terapia Ocupacional e Estética e Cosmética',
    'BIO-FAR (Multidisciplinar I)':
        'Multidisciplinar I - Biomedicina e Farmácia',
    'NUTRI (Multidisciplinar IV)':
        'Multidisciplinar IV - Nutrição',
    'QUÍMICA E FÍSICA':
        'Química e Física - Agronomia',
    'ENGMAKER':
        'EngeMaker | Química e Física - Engenharias e Licenciaturas',
    'ENGMAKER+QUÍMICA E FÍSICA':
        'EngeMaker | Química e Física - Engenharias e Licenciaturas',
}


def ts():
    return datetime.now().strftime('%H:%M:%S')


def limpar(obj):
    if isinstance(obj, dict):   return {k: limpar(v) for k, v in obj.items()}
    if isinstance(obj, list):   return [limpar(v) for v in obj]
    if isinstance(obj, float) and math.isnan(obj): return None
    return obj


def verificar_e_localizar():
    pasta_planilhas = os.path.join(SCRIPT_DIR, "planilhas")
    os.makedirs(pasta_planilhas, exist_ok=True)

    print(f"  Script em : {SCRIPT_DIR}")
    print(f"  Planilhas : {pasta_planilhas}")
    print()

    # Carrega config para usar caminhos diretos
    cfg = {}
    cfg_file = os.path.join(SCRIPT_DIR, "config_links.json")
    if os.path.isfile(cfg_file):
        try:
            with open(cfg_file, encoding="utf-8") as f:
                cfg = json.load(f)
        except: pass

    cam_t = cfg.get("caminho_planilha_tutores", "").strip().strip('"')
    cam_p = cfg.get("caminho_planilha_portfolio", "").strip().strip('"')

    # Usa caminho direto do config se existir
    if cam_t and os.path.isfile(cam_t):
        p1 = cam_t
        print(f"  [OK] {os.path.basename(p1)}")
    else:
        p1 = achar_arquivo(SCRIPT_DIR, "01_CONTROLE_TUTORIA.xlsx")
        if p1: print(f"  [OK] {os.path.basename(p1)}")
        else:  print(f"  [FALTA] 01_CONTROLE_TUTORIA.xlsx")

    if cam_p and os.path.isfile(cam_p):
        p2 = cam_p
        print(f"  [OK] {os.path.basename(p2)}")
    else:
        p2 = achar_arquivo(SCRIPT_DIR, "PORTFOLIO_TUTOR.xlsx")
        if p2: print(f"  [OK] {os.path.basename(p2)}")
        else:  print(f"  [FALTA] PORTFOLIO_TUTOR.xlsx")

    if p1: print(f"  [OK] {os.path.basename(p1)}")
    else:  print(f"  [FALTA] 01_CONTROLE_TUTORIA.xlsx")

    if p2: print(f"  [OK] {os.path.basename(p2)}")
    else:  print(f"  [FALTA] PORTFOLIO_TUTOR.xlsx")

    tmpl = os.path.join(SCRIPT_DIR, "template_dashboard.html")
    if os.path.isfile(tmpl): print(f"  [OK] template_dashboard.html")
    else:                    print(f"  [FALTA] template_dashboard.html")

    return p1, p2, tmpl


def processar(p1, p2):
    print(f"[{ts()}] Lendo tutores...")
    engine_t = 'xlrd' if str(p1).lower().endswith('.xls') and not str(p1).lower().endswith('.xlsx') else None
    df_t = pd.read_excel(p1, sheet_name='Base de Tutores', header=1, **({'engine': engine_t} if engine_t else {}))

    col_sit  = next((c for c in df_t.columns if 'SITUA' in str(c).upper()), None)
    col_nome = next((c for c in df_t.columns if 'NOME'  in str(c).upper() and 'TUTOR' in str(c).upper()), None)
    col_polo = next((c for c in df_t.columns if c == 'POLO'), 'POLO')
    col_cur  = next((c for c in df_t.columns if c == 'CURSOS'), 'CURSOS')
    col_cat  = next((c for c in df_t.columns if 'CATEGORIA' in str(c).upper()), None)

    df_at = df_t[df_t[col_sit].astype(str).str.strip() == 'Ativo'].copy() if col_sit else df_t.copy()
    df_at['_CHAVE'] = df_at[col_polo].astype(str).str.strip() + df_at[col_cur].astype(str).str.strip()

    print(f"[{ts()}] Lendo portfolios...")
    df_p = pd.read_excel(p2, sheet_name='Sheet1')

    # Localizar colunas pelo conteudo do nome
    def col(df, *partes):
        for c in df.columns:
            cu = str(c).upper()
            if all(p.upper() in cu for p in partes):
                return c
        return None

    c_chave = col(df_p, 'CHAVE', 'LINK')
    c_proto = col(df_p, 'PROTOCOLOS', 'ATIVIDADES') 
    # Pega o :7 especificamente
    proto_cols = [c for c in df_p.columns if 'PROTOCOLOS' in str(c).upper() and str(c).endswith(':7')]
    if proto_cols: c_proto = proto_cols[0]
    
    data_cols = [c for c in df_p.columns if 'DATA DA APLICA' in str(c).upper() and str(c).endswith(':7')]
    c_data = data_cols[0] if data_cols else None

    aluno_cols = [c for c in df_p.columns if 'ESTUDANTES' in str(c).upper() and str(c).endswith('72')]
    c_aluno = aluno_cols[0] if aluno_cols else None

    cat_cols = [c for c in df_p.columns if 'CATEGORIA' in str(c).upper() and 'PONTOS' not in str(c).upper() and 'COMENT' not in str(c).upper()]
    c_cat = cat_cols[0] if cat_cols else None

    print(f"[{ts()}] Colunas: chave={c_chave}, proto={c_proto}, data={c_data}, alunos={c_aluno}, cat={c_cat}")
    


    # Detecta coluna de ordem
    c_ordem_cols = [c for c in df_p.columns if 'ORDEM' in str(c).upper() and 'PONTOS' not in str(c).upper() and 'COMENT' not in str(c).upper()]
    c_ordem = c_ordem_cols[0] if c_ordem_cols else None
    print(f"[{ts()}] Coluna ordem: {c_ordem}")

    df_p['_CHAVE']  = df_p[c_chave].astype(str).str.strip() if c_chave else ''
    df_p['_PROTO']  = df_p[c_proto].astype(str).str.strip() if c_proto else ''
    df_p['_DATA']   = pd.to_datetime(df_p[c_data],  errors='coerce') if c_data  else pd.NaT
    df_p['_ALUNOS'] = pd.to_numeric(df_p[c_aluno],  errors='coerce').fillna(0).astype(int) if c_aluno else 0
    df_p['_CAT']    = df_p[c_cat].astype(str).str.strip() if c_cat else ''
    df_p['_ORDEM']  = df_p[c_ordem].astype(str).str.strip() if c_ordem else 'Ordem 1'

    # ── Catalogo: carrega oficial + suplementa com dados reais ─────────────────
    catalogo_oficial = {}
    cat_file = os.path.join(SCRIPT_DIR, 'catalogo_oficial.json')
    if os.path.isfile(cat_file):
        with open(cat_file, encoding='utf-8') as f:
            raw = json.load(f)
        # Converte: {categoria: [{nome, ...}]} -> {categoria: [nome, ...]}
        for cat_nome, praticas in raw.items():
            if isinstance(praticas, list) and praticas:
                if isinstance(praticas[0], dict):
                    catalogo_oficial[cat_nome] = sorted(set(p['nome'] for p in praticas))
                else:
                    catalogo_oficial[cat_nome] = sorted(set(praticas))
        print(f"[{ts()}] Catalogo oficial: {len(catalogo_oficial)} categorias")

    # Mapa de chave -> cat_raw do tutor (para enriquecer portfolio)
    chave_to_cat_raw = {}
    chave_to_cf = {}
    chave_alias = {}  # portfolio_chave -> tutor_chave canônica

    # Primeiro passo: coleta todos os cursos por polo para BIO-FAR
    polo_biofar_cursos = {}  # polo -> set de cursos BBI/BFR
    for _, t in df_at.iterrows():
        polo_   = str(t.get(col_polo, '') or '').strip()
        cursos_ = str(t.get(col_cur,  '') or '').strip()
        cat_    = str(t.get(col_cat,  '') or '').strip() if col_cat else ''
        if cursos_ in ('BBI', 'BFR') and 'BIO-FAR' in cat_.upper():
            if polo_ not in polo_biofar_cursos:
                polo_biofar_cursos[polo_] = set()
            polo_biofar_cursos[polo_].add(cursos_)

    for _, t in df_at.iterrows():
        polo   = str(t.get(col_polo, '') or '').strip()
        cursos = str(t.get(col_cur,  '') or '').strip()
        cat_raw = str(t.get(col_cat, '') or '').strip() if col_cat else ''
        cf = CAT_MAP.get(cat_raw, cat_raw)
        chave = polo + cursos
        if chave and cat_raw:
            chave_to_cat_raw[chave] = cat_raw
            chave_to_cf[chave] = cf

            # Registra variantes BFI para BBI/BFR
            if cursos in ('BBI', 'BFR'):
                outros = polo_biofar_cursos.get(polo, set()) - {cursos}
                variantes = [polo + 'BFI']
                for outro in outros:
                    variantes += [
                        polo + cursos + '-' + outro,  # ex: BFR-BBI
                        polo + outro + '-' + cursos,  # ex: BBI-BFR
                        polo + cursos + outro,         # ex: BFRBBI
                        polo + outro + cursos,         # ex: BBIBFR
                    ]
                for v in variantes:
                    chave_to_cf.setdefault(v, cf)
                    # Mapa reverso: variante portfolio -> chave canônica do tutor
                    chave_alias.setdefault(v, chave)

    # Catalogo real: aprende com portfolios enviados
    catalogo_real = defaultdict(set)
    for _, r in df_p.iterrows():
        chave = str(r.get('_CHAVE', '') or '').strip()
        chave = chave_alias.get(chave, chave)  # normalize
        proto = r['_PROTO']
        if not chave or chave == 'nan' or not proto or proto == 'nan': continue
        cf = chave_to_cf.get(chave, '')
        if not cf: continue
        for p in proto.split(';'):
            p = p.strip()
            if p: catalogo_real[cf].add(p)

    # Merge: oficial como base, real adiciona novas práticas
    catalogo = {}
    all_cats = set(list(catalogo_oficial.keys()) + list(catalogo_real.keys()))
    for cat in all_cats:
        base = set(catalogo_oficial.get(cat, []))
        real = catalogo_real.get(cat, set())
        catalogo[cat] = sorted(base | real)

    print(f"[{ts()}] Catalogo final: {len(catalogo)} cats, {sum(len(v) for v in catalogo.values())} praticas")

    # Email do tutor no portfolio -> cf lookup (fallback para quando chave não casa)
    email_to_cf = {}
    email_to_chave_tutor = {}
    col_email_t = next((c for c in df_t.columns if 'E-MAIL' in str(c).upper() or 'EMAIL' in str(c).upper()), None)
    if col_email_t:
        for _, t in df_at.iterrows():
            em = str(t.get(col_email_t, '') or '').strip().lower()
            chave_t = t['_CHAVE']
            cat_raw_ = str(t.get(col_cat, '') or '').strip() if col_cat else ''
            cf_ = CAT_MAP.get(cat_raw_, cat_raw_)
            if em and em != 'nan':
                email_to_cf[em] = cf_
                email_to_chave_tutor[em] = chave_t

    # Email col in portfolio
    col_email_p = next((c for c in df_p.columns if c.upper() in ('EMAIL', 'E-MAIL')), None)

    # Historico
    enviados = defaultdict(list)
    for _, r in df_p.iterrows():
        chave = r['_CHAVE']; proto = r['_PROTO']
        if not chave or chave == 'nan' or not proto or proto == 'nan': continue
        
        # Normaliza chave: portfolio variant -> chave canônica do tutor
        chave = chave_alias.get(chave, chave)
        
        # Fallback por email se chave ainda não encontra tutor
        if chave not in chave_to_cf and col_email_p:
            em_p = str(r.get(col_email_p, '') or '').strip().lower()
            if em_p in email_to_chave_tutor:
                chave = email_to_chave_tutor[em_p]
        data  = r['_DATA']; aluno = int(r['_ALUNOS'])
        for p in proto.split(';'):
            p = p.strip()
            if p:
                ordem_val = str(r.get('_ORDEM', 'Ordem 1') or 'Ordem 1').strip()
                if not any(o in ordem_val for o in ['Ordem 1','Ordem 2','Ordem 3','Ordem 4','Ordem 5']):
                    ordem_val = 'Ordem 1'
                enviados[chave].append({
                    'p': p[:80],
                    'd': str(data)[:10] if pd.notna(data) else None,
                    'a': aluno,
                    'o': ordem_val,
                })

    # Tutores
    tutores = []
    for _, t in df_at.iterrows():
        chave    = t['_CHAVE']
        cat_raw  = str(t.get(col_cat, '') or '').strip() if col_cat else ''
        cat_form = CAT_MAP.get(cat_raw, cat_raw)
        praticas = catalogo.get(cat_form, catalogo.get(cat_raw, []))
        hist     = enviados.get(chave, [])
        reais    = set(h['p'] for h in hist)
        pend     = [p for p in praticas if p not in reais]
        te = len(reais); tp = len(praticas)
        tutores.append({
            'n': str(t.get(col_nome, '') or ''),
            'p': str(t.get(col_polo, '') or ''),
            'c': cat_raw,
            'cf': cat_form or 'Sem mapeamento',
            'tp': tp, 'te': te,
            'pend': pend, 'real': sorted(reais), 'hist': hist,
            'pct': round(te / tp * 100, 1) if tp else 0,
        })

    # ── Deduplicação por (polo, nome): mesmo tutor com BBI e BFR = 1 entrada ──
    seen = {}
    tutores_dedup = []
    for t in tutores:
        key = (t.get('p',''), t.get('n','').strip().lower())
        if key in seen:
            # Merge: soma te, pend, hist, real do duplicado
            existing = seen[key]
            existing['te'] = max(existing['te'], t['te'])
            existing['hist'] = existing['hist'] + [h for h in t['hist'] if h not in existing['hist']]
            existing['real'] = sorted(set(existing['real']) | set(t['real']))
            existing['pend'] = [p for p in existing['pend'] if p not in existing['real']]
            existing['tp'] = max(existing['tp'], t['tp'])
            existing['pct'] = round(existing['te'] / existing['tp'] * 100, 1) if existing['tp'] else 0
        else:
            seen[key] = t
            tutores_dedup.append(t)
    tutores = tutores_dedup

    # Stats pratica
    ps = defaultdict(lambda: {'enviou': 0, 'nao_enviou': 0, 'categoria': ''})
    for t in tutores:
        for p in t['real']: ps[p]['enviou'] += 1; ps[p]['categoria'] = t['cf']
        for p in t['pend']: ps[p]['nao_enviou'] += 1; ps[p]['categoria'] = t['cf']
    ps_list = sorted([{'nome': k, **v} for k, v in ps.items()], key=lambda x: -x['nao_enviou'])[:30]

    # Stats categoria
    cs = defaultdict(lambda: {'total_tutores': 0, 'com_100pct': 0, 'total_previstas': 0, 'total_enviadas': 0})
    for t in tutores:
        if not t['tp']: continue
        c = t['cf']
        cs[c]['total_tutores'] += 1
        if t['pct'] == 100: cs[c]['com_100pct'] += 1
        cs[c]['total_previstas'] += t['tp']
        cs[c]['total_enviadas']  += t['te']

    print(f"[{ts()}] {len(tutores)} tutores, {sum(len(v) for v in catalogo.values())} praticas")

    # ── Normaliza campos para compatibilidade com template ────────────────────
    prazos = {
        'Ordem 1': '14/03/2026', 'Ordem 2': '11/04/2026',
        'Ordem 3': '09/05/2026', 'Ordem 4': '06/06/2026', 'Ordem 5': '04/07/2026',
    }
    status_ordem = {
        'Ordem 1': 'VENCIDO', 'Ordem 2': 'ABERTA',
        'Ordem 3': 'FUTURA',  'Ordem 4': 'FUTURA',  'Ordem 5': 'FUTURA',
    }

    tutores_out = []
    for t in tutores:
        por_ordem = {}
        for h in t['hist']:
            o = h.get('o', 'Ordem 1') or 'Ordem 1'
            por_ordem[o] = por_ordem.get(o, 0) + 1
        enviou_o1 = por_ordem.get('Ordem 1', 0) > 0
        enviou_o2 = por_ordem.get('Ordem 2', 0) > 0
        sit = 'ok' if enviou_o1 else ('urgente' if not enviou_o2 else 'atrasado')
        tutores_out.append({
            **t,
            'nome':      t.get('n', ''),
            'polo':      t.get('p', ''),
            'cat':       t.get('c', ''),
            'n':         t.get('n', ''),
            'p':         t.get('p', ''),
            'c':         t.get('c', ''),
            'cf':        t.get('cf', 'Sem mapeamento'),
            'por_ordem': por_ordem,
            'porOrdem':  por_ordem,
            'situacao':  sit,
        })

    # ── Deduplica tutores pelo email (mesmo tutor com múltiplos cursos) ───────
    col_email_key = next((c for c in df_t.columns if 'E-MAIL' in str(c).upper()), None)
    # Build email lookup from original df_at
    nome_to_email = {}
    if col_email_key:
        for _, row in df_at.iterrows():
            nome = str(row.get(col_nome, '') or '').strip()
            email = str(row.get(col_email_key, '') or '').strip().lower()
            if nome and email and email != 'nan':
                nome_to_email[nome] = email

    seen = {}
    tutores_dedup = []
    for t in tutores_out:
        nome = t['n']
        polo = t['p']
        email = nome_to_email.get(nome, '')
        key = email if email else (nome + '|' + polo).lower()

        if key not in seen:
            seen[key] = len(tutores_dedup)
            tutores_dedup.append(dict(t))
        else:
            ex = tutores_dedup[seen[key]]
            # Merge histórico e práticas realizadas
            ex['hist'] = ex['hist'] + t['hist']
            merged_real = sorted(set(ex['real']) | set(t['real']))
            ex['real'] = merged_real
            ex['te']   = len(merged_real)
            # Merge por_ordem
            for o, cnt in t['por_ordem'].items():
                ex['por_ordem'][o] = ex['por_ordem'].get(o, 0) + cnt
                ex['porOrdem'][o]  = ex['porOrdem'].get(o, 0) + cnt
            # Recalcula pend e pct
            real_set = set(merged_real)
            ex['pend'] = [p for p in ex['pend'] if p not in real_set]
            if ex['tp'] > 0:
                ex['pct'] = round(ex['te'] / ex['tp'] * 100, 1)
            # Reavalia situação
            enviou_o1 = ex['por_ordem'].get('Ordem 1', 0) > 0
            enviou_o2 = ex['por_ordem'].get('Ordem 2', 0) > 0
            ex['situacao'] = 'ok' if enviou_o1 else ('urgente' if not enviou_o2 else 'atrasado')

    tutores_out = tutores_dedup
    print(f"[{ts()}] Após deduplicação: {len(tutores_out)} tutores únicos")


    total     = len(tutores_out)
    enviaram  = sum(1 for t in tutores_out if t['te'] > 0)
    atrasados = sum(1 for t in tutores_out if t['situacao'] == 'atrasado')
    urgentes  = sum(1 for t in tutores_out if t['situacao'] == 'urgente')
    total_alunos = sum(h['a'] for t in tutores_out for h in t['hist'])

    # ── Polo stats ────────────────────────────────────────────────────────────
    polo_map = {}
    for t in tutores_out:
        p = t['polo']
        if p not in polo_map:
            polo_map[p] = {'POLO': p, 'polo': p, 'total': 0, 'enviaram': 0, 'atrasados': 0, 'alunos': 0}
        polo_map[p]['total'] += 1
        if t['te'] > 0: polo_map[p]['enviaram'] += 1
        if t['situacao'] == 'atrasado': polo_map[p]['atrasados'] += 1
        polo_map[p]['alunos'] += sum(h['a'] for h in t['hist'])
    # Add envios count per polo
    polo_envios = {}
    for t in tutores_out:
        p = t['polo']
        polo_envios[p] = polo_envios.get(p, 0) + len(t.get('hist', []))

    polo_stats = sorted(polo_map.values(), key=lambda x: -x['atrasados'])
    for p in polo_stats:
        p['pend']   = p['total'] - p['enviaram']
        p['pct']    = round(p['enviaram'] / p['total'] * 100) if p['total'] else 0
        p['envios'] = polo_envios.get(p['POLO'], 0)

    # ── Por ordem ─────────────────────────────────────────────────────────────
    ordem_map = {o: {'envios': 0, 'alunos': 0} for o in prazos}
    for t in tutores_out:
        for h in t['hist']:
            o = h.get('o', 'Ordem 1') or 'Ordem 1'
            if o in ordem_map:
                ordem_map[o]['envios'] += 1
                ordem_map[o]['alunos'] += h['a']
    por_ordem = [
        {'ordem': o, 'prazo': prazos[o], 'status': status_ordem[o],
         'envios': ordem_map[o]['envios'], 'alunos': ordem_map[o]['alunos']}
        for o in prazos
    ]

    # ── Por mês ───────────────────────────────────────────────────────────────
    mes_map = {}
    for t in tutores_out:
        for h in t['hist']:
            d = h.get('d') or ''
            mes = d[:7] if d and len(d) >= 7 else 'Sem data'
            if mes not in mes_map: mes_map[mes] = {'MES': mes, 'mes': mes, 'envios': 0, 'alunos': 0}
            mes_map[mes]['envios'] += 1
            mes_map[mes]['alunos'] += h.get('a', 0)
    por_mes = sorted(mes_map.values(), key=lambda x: x['mes'])

    gerado = datetime.now().strftime('%d/%m/%Y %H:%M')
    print(f"[{ts()}] {total} tutores · {enviaram} enviaram · {atrasados} atrasados · {urgentes} urgentes")

    return limpar({
        'kpis': {
            'total': total, 'enviaram': enviaram, 'pendentes': total - enviaram,
            'atrasados': atrasados, 'urgentes': urgentes,
            'total_alunos': total_alunos,
            'total_polos': len(polo_map),
            'polos_ok': sum(1 for p in polo_stats if p['enviaram'] > 0),
        },
        'tutores':      tutores_out,
        'polo_stats':   polo_stats,
        'por_ordem':    por_ordem,
        'cat_stats':    [{'categoria': k, **v} for k, v in cs.items()],
        'pratica_stats': ps_list,
        'catalogo':     catalogo,
        'prazos':       prazos,
        'por_mes':      por_mes,
        'gerado_em':    gerado,
    })


def gerar_html(dados):
    saida = os.path.join(SCRIPT_DIR, "saida")
    os.makedirs(saida, exist_ok=True)
    output = os.path.join(saida, "dashboard.html")
    tmpl   = os.path.join(SCRIPT_DIR, "template_dashboard.html")

    with open(tmpl, encoding='utf-8') as f:
        html = f.read()

    html = html.replace("'DATA_GOES_HERE'", json.dumps(dados, ensure_ascii=False))
    html = html.replace("TIMESTAMP_GOES_HERE", dados['gerado_em'])

    with open(output, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"[{ts()}] Salvo: {output}")
    return output


def modo_watch(p1, p2):
    print(f"[{ts()}] Monitorando a cada {30}s — feche a janela para parar")
    mods = {p1: 0.0, p2: 0.0}
    def loop():
        while True:
            try:
                mudou = any(
                    os.path.getmtime(a) != mods[a]
                    for a in mods if os.path.isfile(a)
                )
                if mudou:
                    for a in mods:
                        if os.path.isfile(a): mods[a] = os.path.getmtime(a)
                    print(f"[{ts()}] Mudanca detectada, atualizando...")
                    gerar_html(processar(p1, p2))
            except Exception as e:
                print(f"[{ts()}] Erro: {e}")
            time.sleep(30)
    threading.Thread(target=loop, daemon=True).start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        print(f"\n[{ts()}] Encerrado.")


if __name__ == '__main__':
    print()
    print(" Verificando arquivos...")
    print()

    p1, p2, tmpl = verificar_e_localizar()

    if not p1 or not p2 or not os.path.isfile(tmpl):
        print()
        print(" Coloque as planilhas na pasta planilhas\\")
        print(" e tente novamente.")
        print()
        input(" Pressione Enter para sair...")
        sys.exit(1)

    print()
    dados  = processar(p1, p2)
    html   = gerar_html(dados)

    print(f"[{ts()}] Abrindo navegador...")
    webbrowser.open(Path(html).as_uri())

    if WATCH_MODE:
        modo_watch(p1, p2)
    else:
        print(f"[{ts()}] Concluido!")
