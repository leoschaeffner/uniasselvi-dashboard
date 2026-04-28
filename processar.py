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
    'REL_GERAL_DE_GERENCIAMENTO.xlsx': ['GERENCIAMENTO', 'REL_GERAL'],
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

    # Planilha de gerenciamento (opcional)
    p3 = achar_arquivo(SCRIPT_DIR, "REL_GERAL_DE_GERENCIAMENTO.xlsx")
    if p3: print(f"  [OK] {os.path.basename(p3)}")
    else:  print(f"  [INFO] REL_GERAL_DE_GERENCIAMENTO.xlsx não encontrada (módulo gerenciamento desativado)")

    return p1, p2, tmpl, p3


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
        print(f"[{ts()}] Catalogo oficial (JSON): {len(catalogo_oficial)} categorias")

    # Fallback: carrega catálogo de planilha Excel se disponível
    if not catalogo_oficial:
        cat_xlsx = achar_arquivo(SCRIPT_DIR, 'CATALOGO_EXPERIMENTOS.xlsx')
        if not cat_xlsx:
            # Tenta nome alternativo
            for f in os.listdir(os.path.join(SCRIPT_DIR, 'planilhas')) if os.path.isdir(os.path.join(SCRIPT_DIR, 'planilhas')) else []:
                fu = f.upper()
                if ('RELAT' in fu and 'EXPER' in fu) or ('CATALOGO' in fu and 'EXPER' in fu):
                    cat_xlsx = os.path.join(SCRIPT_DIR, 'planilhas', f)
                    break
        if cat_xlsx and os.path.isfile(cat_xlsx):
            try:
                df_cat = pd.read_excel(cat_xlsx)
                c_cat_nome = next((c for c in df_cat.columns if 'CATEGORIA' in str(c).upper()), None)
                c_exp_nome = next((c for c in df_cat.columns if 'EXPERIMENTO' in str(c).upper() or 'NOME' in str(c).upper()), None)
                c_sit = next((c for c in df_cat.columns if 'SITUA' in str(c).upper()), None)
                if c_cat_nome and c_exp_nome:
                    if c_sit:
                        df_cat = df_cat[df_cat[c_sit].astype(str).str.strip().str.upper() == 'ATIVO']
                    for cat_val, grp in df_cat.groupby(c_cat_nome):
                        cat_str = str(cat_val).strip()
                        if cat_str and cat_str != 'nan':
                            nomes = sorted(set(str(n).strip() for n in grp[c_exp_nome].dropna() if str(n).strip() and str(n).strip() != 'nan'))
                            if nomes:
                                catalogo_oficial[cat_str] = nomes
                    print(f"[{ts()}] Catalogo oficial (Excel): {len(catalogo_oficial)} categorias, {sum(len(v) for v in catalogo_oficial.values())} práticas")
            except Exception as e:
                print(f"[{ts()}] AVISO: Erro ao ler catálogo Excel: {e}")

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
    ps_all = sorted([{'nome': k, **v} for k, v in ps.items()], key=lambda x: -x['nao_enviou'])
    ps_list = ps_all[:30]  # top 30 for ranking

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
    # Calcula status dinamicamente baseado na data atual
    hoje = datetime.now()
    status_ordem = {}
    for ordem, prazo_str in prazos.items():
        prazo_date = datetime.strptime(prazo_str, '%d/%m/%Y')
        if hoje > prazo_date:
            status_ordem[ordem] = 'VENCIDO'
        elif hoje >= prazo_date.replace(day=1):
            status_ordem[ordem] = 'ABERTA'
        else:
            status_ordem[ordem] = 'FUTURA'

    tutores_out = []
    for t in tutores:
        por_ordem = {}
        for h in t['hist']:
            o = h.get('o', 'Ordem 1') or 'Ordem 1'
            por_ordem[o] = por_ordem.get(o, 0) + 1
        
        # Calcula situação dinâmica: quantas ordens vencidas o tutor enviou?
        ordens_vencidas = [o for o, s in status_ordem.items() if s == 'VENCIDO']
        enviou_todas_vencidas = all(por_ordem.get(o, 0) > 0 for o in ordens_vencidas) if ordens_vencidas else False
        enviou_alguma = any(por_ordem.get(o, 0) > 0 for o in ordens_vencidas) if ordens_vencidas else False
        
        if enviou_todas_vencidas:
            sit = 'ok'
        elif enviou_alguma:
            sit = 'atrasado'  # enviou algumas mas não todas
        else:
            sit = 'urgente'   # não enviou nenhuma ordem vencida
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
            polo_map[p] = {'POLO': p, 'polo': p, 'n': p, 'total': 0, 'enviaram': 0, 'atrasados': 0, 'alunos': 0}
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
        p['n']      = p.get('polo', p.get('POLO', ''))
        p['t']      = p['total']
        p['e']      = p['enviaram']
        p['a']      = p['alunos']
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

    # ── Convenience dicts for template ────────────────────────────────────────
    por_ordem_dict = {o: ordem_map[o]['envios'] for o in prazos}
    alunos_por_ordem = {o: ordem_map[o]['alunos'] for o in prazos}

    # Add short keys to polo_stats for template compatibility
    for p in polo_stats:
        p['n'] = p.get('polo', p.get('POLO', ''))
        p['t'] = p.get('total', 0)
        p['e'] = p.get('enviaram', 0)
        p['a'] = p.get('alunos', 0)

    # Add sit shortcut and al (total alunos) to tutores
    for t in tutores_out:
        t['sit'] = t.get('situacao', 'urgente')
        t['al'] = sum(h.get('a', 0) for h in t.get('hist', []))
        t['email'] = nome_to_email.get(t.get('n', ''), '')

    # pratica_stats with template-compatible keys (ALL practices)
    praticas_template = []
    for p in ps_all:
        total_p = p['enviou'] + p['nao_enviou']
        praticas_template.append({
            'n': p['nome'], 'c': p['categoria'],
            'env_n': p['enviou'], 'pend_n': p['nao_enviou'],
            'pct': round(p['enviou'] / total_p * 100, 1) if total_p else 0,
            'nome': p['nome'], 'enviou': p['enviou'], 'nao_enviou': p['nao_enviou'], 'categoria': p['categoria'],
        })

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
        'por_ordem':    por_ordem_dict,
        'por_ordem_lista': por_ordem,
        'alunos_por_ordem': alunos_por_ordem,
        'status_ordem': status_ordem,
        'cat_stats':    [{'categoria': k, **v} for k, v in cs.items()],
        'pratica_stats': ps_list,
        'praticas':     praticas_template,
        'catalogo':     catalogo,
        'prazos':       prazos,
        'por_mes':      por_mes,
        'gerado_em':    gerado,
    })


def processar_gerenciamento(p3):
    """Processa a planilha REL_GERAL_DE_GERENCIAMENTO.xlsx e retorna dados de gerenciamento."""
    print(f"[{ts()}] Lendo gerenciamento...")
    df_g = pd.read_excel(p3)
    print(f"[{ts()}] Gerenciamento: {len(df_g)} linhas, {len(df_g.columns)} colunas")

    # ── Identificar colunas ───────────────────────────────────────────────────
    def gcol(df, *partes):
        for c in df.columns:
            cu = str(c).upper()
            if all(p.upper() in cu for p in partes):
                return c
        return None

    c_polo     = gcol(df_g, 'CEEM', 'RSOC') or 'CEEM_RSOC'
    c_cat      = gcol(df_g, 'CATP', 'NOME') or 'CATP_NOME'
    c_lab      = gcol(df_g, 'LABE', 'NOME') or 'LABE_NOME'
    c_curso    = gcol(df_g, 'NOME', 'CURS') or 'NOME_CURS'
    c_situ     = gcol(df_g, 'SITU') or 'SITU'
    c_alunos   = gcol(df_g, 'ALUNOS', 'MATRIC') or 'ALUNOS_MATRICULADOS'
    c_capa     = 'CAPA' if 'CAPA' in df_g.columns else gcol(df_g, 'CAPA')
    c_capa_exp = gcol(df_g, 'CAPA', 'EXP') or 'CAPA_EXP'
    c_ofe_cad  = gcol(df_g, 'OFE', 'CAD') or 'OFE_CAD'
    c_qtd_alun = gcol(df_g, 'QTD', 'ALUN') or 'QTD_ALUN'
    c_tutor    = gcol(df_g, 'TUTOR') or 'TUTOR'
    c_dt_ger   = gcol(df_g, 'DT', 'GERENCIAMENTO') or 'DT_GERENCIAMENTO'
    c_dt_agenda = gcol(df_g, 'DT', 'GERENCIADA') or 'DT_GERENCIADA'
    c_hr_agenda = gcol(df_g, 'HR', 'GERENCIADA') or 'HR_GERENCIADA'
    c_ofex_dtin = gcol(df_g, 'OFEX', 'DTIN') or 'OFEX_DTIN'
    c_ofex_dtfi = gcol(df_g, 'OFEX', 'DTFI') or 'OFEX_DTFI'

    # ── Filtrar apenas ativos ─────────────────────────────────────────────────
    if c_situ in df_g.columns:
        df_g = df_g[df_g[c_situ].astype(str).str.strip().str.upper() == 'ATIVO'].copy()
    print(f"[{ts()}] Gerenciamento após filtro ativos: {len(df_g)} linhas")

    # ── Extrair ordem e nome da prática do LABE_NOME ──────────────────────────
    # Formato: "O.1: Nome da Prática" ou "O.2: Nome da Prática"
    df_g['_ORDEM_G'] = ''
    df_g['_PRATICA_G'] = ''
    if c_lab in df_g.columns:
        import re
        def extrair_ordem(val):
            val = str(val or '')
            m = re.match(r'O\.(\d+):\s*(.*)', val)
            if m:
                return f'Ordem {m.group(1)}', m.group(2).strip()
            return '', val.strip()
        parsed = df_g[c_lab].apply(extrair_ordem)
        df_g['_ORDEM_G'] = parsed.apply(lambda x: x[0])
        df_g['_PRATICA_G'] = parsed.apply(lambda x: x[1])

    # ── Campos calculados ─────────────────────────────────────────────────────
    df_g['_GERENCIADO'] = pd.to_numeric(df_g.get(c_ofe_cad, 0), errors='coerce').fillna(0) > 0
    df_g['_TEM_TUTOR'] = df_g[c_tutor].notna() & (df_g[c_tutor].astype(str).str.strip() != '') & (df_g[c_tutor].astype(str).str.strip().str.upper() != 'NAN')
    df_g['_TEM_AGENDA'] = df_g.get(c_dt_agenda, pd.Series(dtype='object')).notna()
    df_g['_ALUNOS_MAT'] = pd.to_numeric(df_g.get(c_alunos, 0), errors='coerce').fillna(0).astype(int)
    df_g['_QTD_ALUN'] = pd.to_numeric(df_g.get(c_qtd_alun, 0), errors='coerce').fillna(0).astype(int)
    df_g['_CAPA'] = pd.to_numeric(df_g.get(c_capa, 0), errors='coerce').fillna(0).astype(int)

    # ── KPIs Globais de Gerenciamento ─────────────────────────────────────────
    total_ofertas = len(df_g)
    ofertas_gerenciadas = int(df_g['_GERENCIADO'].sum())
    ofertas_com_tutor = int(df_g['_TEM_TUTOR'].sum())
    ofertas_sem_tutor = total_ofertas - ofertas_com_tutor
    ofertas_com_agenda = int(df_g['_TEM_AGENDA'].sum())
    total_alunos_mat = int(df_g['_ALUNOS_MAT'].sum())
    total_alunos_agend = int(df_g['_QTD_ALUN'].sum())
    total_capacidade = int(df_g['_CAPA'].sum())
    polos_total = df_g[c_polo].nunique() if c_polo in df_g.columns else 0
    polos_sem_tutor_count = 0
    if c_polo in df_g.columns:
        polos_sem_tutor_count = int(df_g[~df_g['_TEM_TUTOR']].groupby(c_polo).ngroups)

    ger_kpis = {
        'total_ofertas': total_ofertas,
        'ofertas_gerenciadas': ofertas_gerenciadas,
        'ofertas_nao_gerenciadas': total_ofertas - ofertas_gerenciadas,
        'pct_gerenciado': round(ofertas_gerenciadas / total_ofertas * 100, 1) if total_ofertas else 0,
        'ofertas_com_tutor': ofertas_com_tutor,
        'ofertas_sem_tutor': ofertas_sem_tutor,
        'pct_com_tutor': round(ofertas_com_tutor / total_ofertas * 100, 1) if total_ofertas else 0,
        'ofertas_com_agenda': ofertas_com_agenda,
        'total_alunos_matriculados': total_alunos_mat,
        'total_alunos_agendados': total_alunos_agend,
        'total_capacidade': total_capacidade,
        'pct_ocupacao': round(total_alunos_agend / total_capacidade * 100, 1) if total_capacidade else 0,
        'polos_total': polos_total,
        'polos_sem_tutor': polos_sem_tutor_count,
    }

    print(f"[{ts()}] Gerenciamento: {total_ofertas} ofertas, {ofertas_gerenciadas} gerenciadas, {ofertas_sem_tutor} sem tutor")

    # ── Stats por Polo (gerenciamento) ────────────────────────────────────────
    ger_polo = []
    if c_polo in df_g.columns:
        for polo, grp in df_g.groupby(c_polo):
            ger_polo.append({
                'polo': str(polo),
                'total_ofertas': len(grp),
                'gerenciadas': int(grp['_GERENCIADO'].sum()),
                'pct_gerenciado': round(grp['_GERENCIADO'].sum() / len(grp) * 100, 1) if len(grp) else 0,
                'com_tutor': int(grp['_TEM_TUTOR'].sum()),
                'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
                'com_agenda': int(grp['_TEM_AGENDA'].sum()),
                'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()),
                'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
                'capacidade': int(grp['_CAPA'].sum()),
                'tutores_unicos': list(grp[grp['_TEM_TUTOR']][c_tutor].dropna().unique()),
            })
        ger_polo.sort(key=lambda x: -x['sem_tutor'])

    # ── Stats por Categoria (gerenciamento) ───────────────────────────────────
    ger_cat = []
    if c_cat in df_g.columns:
        for cat, grp in df_g.groupby(c_cat):
            ger_cat.append({
                'categoria': str(cat),
                'total_ofertas': len(grp),
                'gerenciadas': int(grp['_GERENCIADO'].sum()),
                'pct_gerenciado': round(grp['_GERENCIADO'].sum() / len(grp) * 100, 1) if len(grp) else 0,
                'com_tutor': int(grp['_TEM_TUTOR'].sum()),
                'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
                'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()),
                'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
            })
        ger_cat.sort(key=lambda x: -x['total_ofertas'])

    # ── Stats por Ordem (gerenciamento) ───────────────────────────────────────
    ger_ordem = []
    ordens_encontradas = df_g['_ORDEM_G'].unique()
    ordens_validas = [o for o in sorted(ordens_encontradas) if o and 'Ordem' in str(o)]
    for ordem in ordens_validas:
        grp = df_g[df_g['_ORDEM_G'] == ordem]
        # Datas da oferta
        datas_inicio = pd.to_datetime(grp.get(c_ofex_dtin, pd.Series(dtype='object')), errors='coerce').dropna()
        datas_fim = pd.to_datetime(grp.get(c_ofex_dtfi, pd.Series(dtype='object')), errors='coerce').dropna()
        dt_inicio = datas_inicio.min().strftime('%d/%m/%Y') if len(datas_inicio) > 0 else ''
        dt_fim = datas_fim.max().strftime('%d/%m/%Y') if len(datas_fim) > 0 else ''

        ger_ordem.append({
            'ordem': ordem,
            'total_ofertas': len(grp),
            'gerenciadas': int(grp['_GERENCIADO'].sum()),
            'pct_gerenciado': round(grp['_GERENCIADO'].sum() / len(grp) * 100, 1) if len(grp) else 0,
            'com_tutor': int(grp['_TEM_TUTOR'].sum()),
            'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()),
            'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
            'dt_inicio': dt_inicio,
            'dt_fim': dt_fim,
        })

    # ── Contratação: tutores por polo e categoria ─────────────────────────────
    ger_contratacao = []
    if c_polo in df_g.columns and c_cat in df_g.columns:
        for (polo, cat), grp in df_g.groupby([c_polo, c_cat]):
            tutores_list = list(grp[grp['_TEM_TUTOR']][c_tutor].dropna().unique())
            tem_tutor = len(tutores_list) > 0
            ger_contratacao.append({
                'polo': str(polo),
                'categoria': str(cat),
                'total_ofertas': len(grp),
                'tem_tutor': tem_tutor,
                'tutores': [str(t) for t in tutores_list],
                'status': 'Contratado' if tem_tutor else 'Sem tutor',
            })
        ger_contratacao.sort(key=lambda x: (0 if x['tem_tutor'] else 1, x['polo']))

    # ── Agendas: ofertas com e sem agenda ─────────────────────────────────────
    ger_agendas = []
    if c_polo in df_g.columns:
        for polo, grp in df_g.groupby(c_polo):
            total = len(grp)
            com_agenda = int(grp['_TEM_AGENDA'].sum())
            sem_agenda = total - com_agenda
            ger_agendas.append({
                'polo': str(polo),
                'total': total,
                'com_agenda': com_agenda,
                'sem_agenda': sem_agenda,
                'pct_agendado': round(com_agenda / total * 100, 1) if total else 0,
            })
        ger_agendas.sort(key=lambda x: -x['sem_agenda'])

    # ── Tabela detalhada de ofertas (top 500 para não sobrecarregar o HTML) ───
    ger_ofertas_detalhe = []
    cols_detalhe = [c_polo, c_cat, '_ORDEM_G', '_PRATICA_G', c_curso, c_tutor,
                    '_GERENCIADO', '_TEM_AGENDA', '_ALUNOS_MAT', '_QTD_ALUN', '_CAPA']
    for _, row in df_g.head(5000).iterrows():
        ger_ofertas_detalhe.append({
            'polo': str(row.get(c_polo, '')),
            'categoria': str(row.get(c_cat, '')),
            'ordem': str(row.get('_ORDEM_G', '')),
            'pratica': str(row.get('_PRATICA_G', '')),
            'curso': str(row.get(c_curso, '')),
            'tutor': str(row.get(c_tutor, '')) if pd.notna(row.get(c_tutor)) else '',
            'gerenciado': bool(row.get('_GERENCIADO', False)),
            'tem_agenda': bool(row.get('_TEM_AGENDA', False)),
            'alunos_mat': int(row.get('_ALUNOS_MAT', 0)),
            'alunos_agend': int(row.get('_QTD_ALUN', 0)),
            'capacidade': int(row.get('_CAPA', 0)),
        })

    resultado = {
        'ger_kpis': ger_kpis,
        'ger_polo': ger_polo,
        'ger_cat': ger_cat,
        'ger_ordem': ger_ordem,
        'ger_contratacao': ger_contratacao,
        'ger_agendas': ger_agendas,
        'ger_ofertas': ger_ofertas_detalhe,
    }

    print(f"[{ts()}] Gerenciamento processado: {len(ger_polo)} polos, {len(ger_cat)} categorias, {len(ger_ordem)} ordens")
    return resultado


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

    p1, p2, tmpl, p3 = verificar_e_localizar()

    if not p1 or not p2 or not os.path.isfile(tmpl):
        print()
        print(" Coloque as planilhas na pasta planilhas\\")
        print(" e tente novamente.")
        print()
        if '--sem-browser' not in sys.argv:
            input(" Pressione Enter para sair...")
        sys.exit(1)

    print()
    dados  = processar(p1, p2)

    # Integra dados de gerenciamento se a planilha existir
    if p3:
        try:
            ger_dados = processar_gerenciamento(p3)
            dados.update(ger_dados)
            dados['tem_gerenciamento'] = True
        except Exception as e:
            print(f"[{ts()}] AVISO: Erro ao processar gerenciamento: {e}")
            import traceback; traceback.print_exc()
            dados['tem_gerenciamento'] = False
    else:
        dados['tem_gerenciamento'] = False

    html   = gerar_html(dados)

    if '--sem-browser' not in sys.argv:
        print(f"[{ts()}] Abrindo navegador...")
        webbrowser.open(Path(html).as_uri())

    if WATCH_MODE:
        modo_watch(p1, p2)
    else:
        print(f"[{ts()}] Concluido!")
