"""
UNIASSELVI - Dashboard de Portfolios
Le as planilhas da pasta planilhas/ e gera saida/dashboard.html
"""

import pandas as pd
import json, os, sys, math, webbrowser, time, threading, glob
from pathlib import Path
from datetime import datetime, timezone, timedelta
from collections import defaultdict

# Prazos semestrais — edite aqui a cada semestre
PRAZOS_ORDENS = {
    'Ordem 1': '14/03/2026',
    'Ordem 2': '11/04/2026',
    'Ordem 3': '09/05/2026',
    'Ordem 4': '06/06/2026',
    'Ordem 5': '04/07/2026',
}

# Datas de início e fim da realização das práticas (período de oferta)
# Fonte: tabela do Help Tutor 2026/1
PERIODOS_ORDENS = {
    'Ordem 1': {'inicio': '16/02/2026', 'fim': '14/03/2026', 'semanas': 4},
    'Ordem 2': {'inicio': '16/03/2026', 'fim': '11/04/2026', 'semanas': 4},
    'Ordem 3': {'inicio': '13/04/2026', 'fim': '09/05/2026', 'semanas': 4},
    'Ordem 4': {'inicio': '11/05/2026', 'fim': '06/06/2026', 'semanas': 4},
    'Ordem 5': {'inicio': '08/06/2026', 'fim': '04/07/2026', 'semanas': 4},
}

# Constantes de CH
CH_ADMIN_FATOR   = 0.25   # 1h a cada 4h = 25% para administrativo
CH_PRATICA_DURAC = 1.5    # cada prática dura 1:30h

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
    BRT = timezone(timedelta(hours=-3))
    return datetime.now(BRT).strftime('%H:%M:%S')


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

    # Planilha de gerenciamento (novo formato ou antigo — detectado automaticamente)
    p3 = achar_arquivo(SCRIPT_DIR, "REL_GERAL_DE_GERENCIAMENTO.xlsx")
    if p3: print(f"  [OK] {os.path.basename(p3)}")
    else:  print(f"  [INFO] REL_GERAL_DE_GERENCIAMENTO.xlsx não encontrada (módulo desativado)")

    # Lotação de tutores (enriquecimento)
    p4 = achar_arquivo(SCRIPT_DIR, "LOTACAO_TUTORES.xlsx")
    if p4: print(f"  [OK] {os.path.basename(p4)}")
    else:  print(f"  [INFO] LOTACAO_TUTORES.xlsx não encontrada (dados de perfil indisponíveis)")

    return p1, p2, tmpl, p3, p4


def ler_excel(path, **kwargs):
    """Lê Excel tentando múltiplos engines — robusto para arquivos do SharePoint."""
    for engine in ('openpyxl', 'xlrd', None):
        try:
            kw = dict(kwargs)
            if engine:
                kw['engine'] = engine
            return pd.read_excel(path, **kw)
        except Exception:
            continue
    raise ValueError(f"Não foi possível ler {path} com nenhum engine disponível")


def processar(p1, p2):
    print(f"[{ts()}] Lendo tutores...")
    # Validar que o arquivo é realmente um Excel (magic bytes PK)
    with open(p1, 'rb') as _f:
        _magic = _f.read(8)
        _preview = _f.read(200)
    print(f"  [DEBUG] Magic bytes de {p1}: {_magic.hex()}")
    if _magic[:2] == b'PK':
        print(f"  [DEBUG] Formato ZIP/XLSX confirmado")
    elif _magic[:2] in (b'\xd0\xcf', b'\xCF\xD0'):
        print(f"  [DEBUG] Formato XLS (OLE2) confirmado")
    else:
        print(f"  [ERRO] Arquivo não é Excel válido. Conteúdo inicial:")
        print((_magic + _preview).decode('utf-8', errors='replace')[:300])
        raise ValueError(f"Arquivo {p1} não é um Excel válido — SharePoint pode ter retornado HTML de erro")
    df_t = ler_excel(p1, sheet_name='Base de Tutores', header=1)

    col_sit  = next((c for c in df_t.columns if 'SITUA' in str(c).upper()), None)
    col_nome = next((c for c in df_t.columns if 'NOME'  in str(c).upper() and 'TUTOR' in str(c).upper()), None)
    col_polo = next((c for c in df_t.columns if c == 'POLO'), 'POLO')
    col_cur  = next((c for c in df_t.columns if c == 'CURSOS'), 'CURSOS')
    col_cat  = next((c for c in df_t.columns if 'CATEGORIA' in str(c).upper()), None)

    df_at = df_t[df_t[col_sit].astype(str).str.strip() == 'Ativo'].copy() if col_sit else df_t.copy()
    df_at['_CHAVE'] = df_at[col_polo].astype(str).str.strip() + df_at[col_cur].astype(str).str.strip()

    print(f"[{ts()}] Lendo portfolios...")
    df_p = ler_excel(p2, sheet_name='Sheet1')

    # Localizar colunas pelo conteudo do nome
    def col(df, *partes):
        for c in df.columns:
            cu = str(c).upper()
            if all(p.upper() in cu for p in partes):
                return c
        return None

    c_chave = col(df_p, 'CHAVE', 'LINK')
    c_proto = col(df_p, 'PROTOCOLOS', 'ATIVIDADES') 
    # Tenta sufixo :7, depois :8, depois qualquer PROTOCOLOS
    for sfx in (':7', ':8', ':6', ':9'):
        proto_sfx = [c for c in df_p.columns if 'PROTOCOLOS' in str(c).upper() and str(c).endswith(sfx)]
        if proto_sfx: c_proto = proto_sfx[0]; break
    else:
        proto_any = [c for c in df_p.columns if 'PROTOCOLOS' in str(c).upper() or 'ATIVIDADE' in str(c).upper()]
        if proto_any: c_proto = proto_any[0]
    
    # Data: tenta sufixo :7 primeiro, depois qualquer sufixo
    data_cols = [c for c in df_p.columns if 'DATA DA APLICA' in str(c).upper() and str(c).endswith(':7')]
    if not data_cols:
        data_cols = [c for c in df_p.columns if 'DATA DA APLICA' in str(c).upper()]
    if not data_cols:
        data_cols = [c for c in df_p.columns if 'DATA' in str(c).upper() and 'APLICA' in str(c).upper()]
    c_data = data_cols[0] if data_cols else None

    # Detecção robusta: tenta várias estratégias em ordem de confiança
    def find_aluno_col(df):
        cols = df.columns.tolist()
        # 1. Contém ESTUDANTES e termina em número (sufixo do Forms)
        for suffix_end in ('72', '73', '74', '70', '71', '75'):
            for c in cols:
                if 'ESTUDANTES' in str(c).upper() and str(c).endswith(suffix_end):
                    return c
        # 2. Contém ESTUDANTES (qualquer sufixo)
        for c in cols:
            if 'ESTUDANTES' in str(c).upper() and 'PONTOS' not in str(c).upper():
                return c
        # 3. Contém ALUNO ou ALUNOS
        for c in cols:
            cu = str(c).upper()
            if ('ALUNO' in cu or 'ALUNOS' in cu) and 'PONTOS' not in cu and 'COMENT' not in cu:
                return c
        # 4. Contém QUANTIDADE ou QTD próximo de ALUNO
        for c in cols:
            cu = str(c).upper()
            if 'QUANTIDADE' in cu or 'QTD' in cu:
                return c
        return None
    c_aluno = find_aluno_col(df_p)

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

    # Mapa rápido: prática -> categoria oficial (para blindar catalogo_real)
    oficial_p_to_cat = {}
    for cat, pracs in catalogo_oficial.items():
        for p in pracs:
            oficial_p_to_cat.setdefault(p, cat)

    # Catalogo real: aprende com portfolios enviados.
    # Só adiciona prática a uma categoria se ela NÃO pertencer já ao catálogo
    # oficial de OUTRA categoria — evita que fisioterapia contamine bio-far e vice-versa.
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
            if not p: continue
            cat_oficial = oficial_p_to_cat.get(p)
            # Prática já no catálogo oficial de OUTRA categoria → não contaminar
            if cat_oficial and cat_oficial != cf:
                continue
            catalogo_real[cf].add(p)

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

    # Mapa prática -> categoria a partir do catálogo (fonte de verdade)
    # Garante que Fisioterapia não apareça como BIO-FAR por causa
    # do loop sobrescrever a categoria com o último tutor processado.
    p_to_cat = {}
    for cat, pracs in catalogo.items():
        for p in pracs:
            if p not in p_to_cat:          # primeiro catálogo que contém a prática vence
                p_to_cat[p] = cat

    # Stats pratica — conta por tutor, mas categoria vem do catálogo
    ps = defaultdict(lambda: {'enviou': 0, 'nao_enviou': 0, 'categoria': ''})
    for t in tutores:
        for p in t['real']:  ps[p]['enviou']    += 1
        for p in t['pend']:  ps[p]['nao_enviou'] += 1
    # Atribui categoria do catálogo; se a prática não estiver no catálogo,
    # usa o cf do primeiro tutor que a enviou (compatibilidade retroativa)
    _p_fallback = {}
    for t in tutores:
        for p in t['real'] + t['pend']:
            _p_fallback.setdefault(p, t['cf'])
    for p in ps:
        ps[p]['categoria'] = p_to_cat.get(p, _p_fallback.get(p, ''))

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
    prazos = PRAZOS_ORDENS.copy()
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

    BRT = timezone(timedelta(hours=-3))
    gerado = datetime.now(BRT).strftime('%d/%m/%Y %H:%M')
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



def carregar_lotacao(p4):
    """Carrega LOTACAO_TUTORES.xlsx e retorna mapa nome_lower -> dados do tutor."""
    from openpyxl import load_workbook as _lwb
    print(f"[{ts()}] Lendo lotação de tutores...")
    _wb = _lwb(str(p4), read_only=True, data_only=True)
    _ws = _wb['Quadro Geral de Lotação'] if 'Quadro Geral de Lotação' in _wb.sheetnames else list(_wb.worksheets)[0]
    _rows = list(_ws.iter_rows(values_only=True))

    def parse_ch(v):
        """Converte HH:MM ou HH:MM:SS para horas decimais."""
        try:
            v = str(v or '').strip()
            if ':' in v:
                parts = v.split(':')
                return float(parts[0]) + float(parts[1])/60
            return float(v)
        except: return 0.0

    lotacao = {}
    for r in _rows[2:]:
        if not r[8] or str(r[8]).strip() in ('', '-', 'None', 'nan'):
            continue
        nome_raw = str(r[8]).strip()
        nome_lower = nome_raw.lower()
        lotacao[nome_lower] = {
            'nome_oficial': nome_raw,
            'perfil':       str(r[13] or '').strip(),
            'cursos':       str(r[0]  or '').strip(),
            'ch_semanal':   parse_ch(r[14]),
            'ch_ideal':     parse_ch(r[15]),
            'contratacao':  str(r[7]  or '').strip(),
            'polo_hub':     str(r[4]  or '').strip(),
            'categoria_gio': str(r[29] or '').strip(),
        }
    print(f"[{ts()}] Lotação: {len(lotacao)} tutores mapeados")
    return lotacao


def enriquecer_tutores(dados, lotacao):
    """Enriquece tutores com perfil, cursos e CH da planilha de lotação."""
    tutores = dados.get('tutores', [])
    matched = 0
    for t in tutores:
        nome_lower = str(t.get('n', '')).lower()
        # Tentar match exato; depois match parcial
        info = lotacao.get(nome_lower)
        if not info:
            for k, v in lotacao.items():
                if nome_lower in k or k in nome_lower:
                    info = v
                    break
        if info:
            t['perfil']     = info['perfil']
            t['cursos']     = info['cursos']
            t['ch_semanal'] = info['ch_semanal']
            t['ch_ideal']   = info['ch_ideal']
            t['contratacao_lot'] = info['contratacao']
            matched += 1
    print(f"[{ts()}] Enriquecimento: {matched}/{len(tutores)} tutores com perfil/CH")
    dados['tutores'] = tutores
    return dados


def processar_gerenciamento_csv(p5):
    """Processa o novo CSV detalhado de gerenciamento (REL_DETALHADO.csv)."""
    import csv, re as _re
    print(f"[{ts()}] Lendo gerenciamento (novo CSV)...")

    # Detectar encoding
    for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
        try:
            with open(str(p5), 'r', encoding=enc, errors='replace') as f:
                rows = list(csv.reader(f, delimiter=';'))
            break
        except: continue

    header = rows[0]
    data   = rows[1:]

    # Mapeamento de colunas pelo nome
    col = {h.strip().upper(): i for i, h in enumerate(header)}
    def gc(name): return col.get(name.upper())

    ci_polo   = gc('LABORATORIO')
    ci_cat    = gc('CATEGORIA')
    ci_exp    = gc('NOME_EXPERIMENTO')
    ci_tutor  = gc('TUTOR')
    ci_mat    = gc('ALUNOS_MATRICULADOS')
    ci_agend  = gc('ALUNOS_AGENDADOS')
    ci_pend   = gc('PENDENCIA_AGENDAMENTOS')
    ci_capa   = gc('CAPACIDADE_TOTAL')
    ci_ofe    = gc('OFERTAS_CADASTRADAS')
    ci_situ   = gc('SITU_OFERTA')
    ci_dt_ger = gc('DT_GERENCIAMENTO')
    ci_dt_ag  = gc('DT_GERENCIADA')
    ci_hr_ag  = gc('HR_GERENCIADA')
    ci_sem    = gc('SEMESTRE')

    def gv(row, ci, default=''):
        try: return str(row[ci]).strip() if ci is not None and ci < len(row) else default
        except: return default

    def gn(row, ci):
        try: return float(str(row[ci]).replace(',','.').strip()) if ci is not None and ci < len(row) and row[ci] else 0
        except: return 0

    print(f"[{ts()}] Gerenciamento CSV: {len(data)} linhas, {len(header)} colunas")

    # ── Extrair ordem e prática do NOME_EXPERIMENTO ───────────────────────────
    def extrair_ordem_exp(val):
        m = _re.match(r'O\.(\d+):\s*(.*)', str(val or ''))
        if m: return f'Ordem {m.group(1)}', m.group(2).strip()
        return '', str(val or '').strip()

    # Construir lista de registros
    registros = []
    for r in data:
        polo   = gv(r, ci_polo)
        cat    = gv(r, ci_cat)
        exp    = gv(r, ci_exp)
        tutor  = gv(r, ci_tutor)
        situ   = gv(r, ci_situ)
        if not polo: continue

        ordem, pratica = extrair_ordem_exp(exp)
        mat   = int(gn(r, ci_mat))
        agend = int(gn(r, ci_agend))
        capa  = int(gn(r, ci_capa))
        ofe   = int(gn(r, ci_ofe))
        dt_ag = gv(r, ci_dt_ag)
        hr_ag = gv(r, ci_hr_ag)

        # Converter data DD/MM/AAAA para AAAA-MM-DD
        dt_ag_iso = ''
        if dt_ag and '/' in dt_ag:
            try:
                parts = dt_ag.split('/')
                dt_ag_iso = f"{parts[2]}-{parts[1]}-{parts[0]}"
            except: pass

        tem_tutor  = bool(tutor)
        tem_agenda = bool(dt_ag_iso)
        gerenciado = ofe > 0

        registros.append({
            'polo': polo, 'categoria': cat, 'pratica': pratica,
            'ordem': ordem, 'tutor': tutor if tem_tutor else '',
            'tem_tutor': tem_tutor, 'tem_agenda': tem_agenda,
            'gerenciado': gerenciado, 'situ': situ,
            'alunos_mat': mat, 'alunos_agend': agend,
            'capacidade': capa, 'ofertas_cad': ofe,
            'dt_agenda_iso': dt_ag_iso, 'hr_agenda': hr_ag,
        })

    df_r = pd.DataFrame(registros)
    if df_r.empty:
        return {'ger_kpis': {}, 'ger_polo': [], 'ger_cat': [], 'ger_ordem': [],
                'ger_contratacao': [], 'ger_agendas': [], 'ger_ofertas': []}

    # ── KPIs Globais ──────────────────────────────────────────────────────────
    total       = len(df_r)
    com_tutor   = int(df_r['tem_tutor'].sum())
    sem_tutor   = total - com_tutor
    gerenciadas = int(df_r['gerenciado'].sum())
    com_agenda  = int(df_r['tem_agenda'].sum())
    tot_mat     = int(df_r['alunos_mat'].sum())
    tot_agend   = int(df_r['alunos_agend'].sum())
    tot_capa    = int(df_r['capacidade'].sum())
    ordens_u    = df_r['ordem'].nunique()
    print(f"[{ts()}] Gerenciamento: {total} ofertas, {gerenciadas} gerenciadas, {sem_tutor} sem tutor")
    print(f"[{ts()}] Gerenciamento processado: {df_r['polo'].nunique()} polos, {df_r['categoria'].nunique()} categorias, {ordens_u} ordens")

    ger_kpis = {
        'total_ofertas': total, 'ofertas_gerenciadas': gerenciadas,
        'ofertas_nao_gerenciadas': total - gerenciadas,
        'pct_gerenciado': round(gerenciadas/total*100,1) if total else 0,
        'ofertas_com_tutor': com_tutor, 'ofertas_sem_tutor': sem_tutor,
        'pct_com_tutor': round(com_tutor/total*100,1) if total else 0,
        'ofertas_com_agenda': com_agenda,
        'total_alunos_matriculados': tot_mat,
        'total_alunos_agendados': tot_agend,
        'total_capacidade': tot_capa,
        'pct_ocupacao': round(tot_agend/tot_capa*100,1) if tot_capa else 0,
        'polos_total': df_r['polo'].nunique(),
        'polos_sem_tutor': int(df_r[~df_r['tem_tutor']].groupby('polo').ngroups),
    }

    # ── Por Polo ──────────────────────────────────────────────────────────────
    ger_polo = []
    for polo, grp in df_r.groupby('polo'):
        tutores_unicos = list(grp[grp['tem_tutor']]['tutor'].dropna().unique())
        ger_polo.append({
            'polo': str(polo),
            'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()),
            'sem_tutor': int((~grp['tem_tutor']).sum()),
            'com_agenda': int(grp['tem_agenda'].sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()),
            'alunos_agendados': int(grp['alunos_agend'].sum()),
            'capacidade': int(grp['capacidade'].sum()),
            'tutores_unicos': [str(t) for t in tutores_unicos],
        })
    ger_polo.sort(key=lambda x: -x['sem_tutor'])

    # ── Por Categoria ─────────────────────────────────────────────────────────
    ger_cat = []
    for cat, grp in df_r.groupby('categoria'):
        ger_cat.append({
            'categoria': str(cat),
            'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()),
            'sem_tutor': int((~grp['tem_tutor']).sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()),
            'alunos_agendados': int(grp['alunos_agend'].sum()),
        })
    ger_cat.sort(key=lambda x: -x['total_ofertas'])

    # ── Por Ordem ─────────────────────────────────────────────────────────────
    ger_ordem = []
    ordem_map = {'Ordem 1':1,'Ordem 2':2,'Ordem 3':3,'Ordem 4':4,'Ordem 5':5}
    for ordem in sorted(df_r['ordem'].unique(), key=lambda x: ordem_map.get(x,9)):
        if not ordem: continue
        grp = df_r[df_r['ordem']==ordem]
        # Datas da oferta: usar DATA_OFERTADA e DATA_EXPIRACAO do CSV se disponíveis
        # Por ora usar prazos configurados
        dt_ini = ''
        dt_fim = PRAZOS_ORDENS.get(ordem, '')
        ger_ordem.append({
            'ordem': ordem,
            'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()),
            'alunos_agendados': int(grp['alunos_agend'].sum()),
            'dt_inicio': dt_ini,
            'dt_fim': dt_fim,
        })

    # ── Contratação ───────────────────────────────────────────────────────────
    ger_contratacao = []
    for (polo, cat), grp in df_r.groupby(['polo','categoria']):
        tutores_list = list(grp[grp['tem_tutor']]['tutor'].dropna().unique())
        tem_tutor = len(tutores_list) > 0
        ger_contratacao.append({
            'polo': str(polo), 'categoria': str(cat),
            'total_ofertas': len(grp),
            'tem_tutor': tem_tutor,
            'tutores': [str(t) for t in tutores_list],
            'status': 'Contratado' if tem_tutor else 'Sem tutor',
        })

    # ── Agendas por Polo ──────────────────────────────────────────────────────
    ger_agendas = []
    for polo, grp in df_r.groupby('polo'):
        total_p = len(grp)
        com_ag  = int(grp['tem_agenda'].sum())
        sem_ag  = total_p - com_ag

        # Datas por categoria
        datas_por_cat = {}
        for _, row in grp[grp['tem_agenda']].iterrows():
            d = row['dt_agenda_iso']
            c = row['categoria']
            t = row['tutor'] or ''
            if d:
                if d not in datas_por_cat:
                    datas_por_cat[d] = {'cats': [], 'tutores': []}
                if c and c not in datas_por_cat[d]['cats']:
                    datas_por_cat[d]['cats'].append(c)
                if t and t not in datas_por_cat[d]['tutores']:
                    datas_por_cat[d]['tutores'].append(t)

        ger_agendas.append({
            'polo': str(polo),
            'total': total_p,
            'com_agenda': com_ag,
            'sem_agenda': sem_ag,
            'pct_agendado': round(com_ag/total_p*100, 1) if total_p else 0,
            'datas_agenda': sorted(datas_por_cat.keys()),
            'datas_por_cat': {d: v['cats'] for d, v in datas_por_cat.items()},
            'datas_por_tutor': {d: v['tutores'] for d, v in datas_por_cat.items()},
        })
    ger_agendas.sort(key=lambda x: -x['sem_agenda'])

    # ── Detalhe de Ofertas (todas) ────────────────────────────────────────────
    ger_ofertas_detalhe = []
    for _, row in df_r.iterrows():
        ger_ofertas_detalhe.append({
            'polo': row['polo'], 'categoria': row['categoria'],
            'ordem': row['ordem'], 'pratica': row['pratica'],
            'tutor': row['tutor'], 'tem_tutor': row['tem_tutor'],
            'tem_agenda': row['tem_agenda'], 'gerenciado': row['gerenciado'],
            'alunos_mat': row['alunos_mat'], 'alunos_agend': row['alunos_agend'],
            'dt_agenda': row['dt_agenda_iso'], 'hr_agenda': row['hr_agenda'],
        })

    return {
        'ger_kpis': ger_kpis, 'ger_polo': ger_polo,
        'ger_cat': ger_cat, 'ger_ordem': ger_ordem,
        'ger_contratacao': ger_contratacao,
        'ger_agendas': ger_agendas,
        'ger_ofertas': ger_ofertas_detalhe,
    }


def _processar_gerenciamento_novo(df_g):
    """Processa o novo formato de gerenciamento (DataFrame já lido)."""
    import re as _re

    # Mapeamento de colunas pelo nome exato
    col = {str(c).strip().upper(): c for c in df_g.columns}
    def gc(name): return col.get(name.upper())

    c_polo  = gc('LABORATORIO')
    c_cat   = gc('CATEGORIA')
    c_exp   = gc('NOME_EXPERIMENTO')
    c_tutor = gc('TUTOR')
    c_mat   = gc('ALUNOS_MATRICULADOS')
    c_agend = gc('ALUNOS_AGENDADOS')
    c_capa  = gc('CAPACIDADE_TOTAL')
    c_ofe   = gc('OFERTAS_CADASTRADAS')
    c_situ  = gc('SITU_OFERTA')
    c_dt_ag = gc('DT_GERENCIADA')
    c_hr_ag = gc('HR_GERENCIADA')

    def extrair_ordem_exp(val):
        m = _re.match(r'O\.(\d+):\s*(.*)', str(val or ''))
        if m: return f'Ordem {m.group(1)}', m.group(2).strip()
        return '', str(val or '').strip()

    def safe_int(val):
        try: return int(float(str(val).replace(',','.').strip() or 0))
        except: return 0

    # Construir campos calculados
    df = df_g.copy()
    df['_POLO']    = df[c_polo].astype(str).str.strip() if c_polo else ''
    df['_CAT']     = df[c_cat].astype(str).str.strip()  if c_cat  else ''
    df['_TUTOR']   = df[c_tutor].fillna('').astype(str).str.strip().replace('nan','') if c_tutor else ''
    df['_MAT']     = pd.to_numeric(df[c_mat],  errors='coerce').fillna(0).astype(int) if c_mat  else 0
    df['_AGEND']   = pd.to_numeric(df[c_agend],errors='coerce').fillna(0).astype(int) if c_agend else 0
    df['_CAPA']    = pd.to_numeric(df[c_capa], errors='coerce').fillna(0).astype(int) if c_capa  else 0
    df['_OFE']     = pd.to_numeric(df[c_ofe],  errors='coerce').fillna(0).astype(int) if c_ofe   else 0
    df['_TEM_TUTOR'] = df['_TUTOR'].str.len() > 0
    # Gerenciado = tem oferta cadastrada OU situação Concluído
    _situ_col = df[c_situ].fillna('').astype(str).str.strip() if c_situ else pd.Series([''] * len(df))
    df['_GERENCIADO'] = (df['_OFE'] > 0) | _situ_col.str.upper().str.contains('CONCLU', na=False)

    # Data de agenda
    dt_col = df[c_dt_ag] if c_dt_ag else pd.Series([''] * len(df))  # manter tipo nativo (datetime)
    def to_iso(v):
        if v is None: return ''
        # Objeto datetime/date nativo (xlsx entrega assim)
        try:
            import datetime as _dt
            if isinstance(v, (_dt.datetime, _dt.date)):
                return v.strftime('%Y-%m-%d')
        except: pass
        sv = str(v).strip()
        if not sv or sv == 'nan': return ''
        # String DD/MM/AAAA
        if '/' in sv:
            try:
                parts = sv.split('/')
                if len(parts) == 3: return f'{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}'
            except: pass
        # String AAAA-MM-DD já no formato correto
        if '-' in sv and len(sv) >= 10:
            return sv[:10]
        # Número serial do Excel
        try:
            n = float(sv)
            import datetime as _dt
            base = _dt.date(1899, 12, 30)
            return (base + _dt.timedelta(days=int(n))).strftime('%Y-%m-%d')
        except: pass
        return ''
    df['_DT_AG_ISO'] = dt_col.apply(to_iso)
    df['_TEM_AGENDA'] = df['_DT_AG_ISO'].str.len() > 0
    df['_HR_AG'] = df[c_hr_ag].fillna('').astype(str).str.strip().replace('nan','').replace('NaT','') if c_hr_ag else ''

    # Ordem e prática
    parsed = (df[c_exp] if c_exp else pd.Series([''] * len(df))).apply(extrair_ordem_exp)
    df['_ORDEM']  = parsed.apply(lambda x: x[0])
    df['_PRATICA'] = parsed.apply(lambda x: x[1])

    # Filtrar linhas sem polo
    df = df[df['_POLO'].str.len() > 0].copy()

    total       = len(df)
    com_tutor   = int(df['_TEM_TUTOR'].sum())
    gerenciadas = int(df['_GERENCIADO'].sum())
    com_agenda  = int(df['_TEM_AGENDA'].sum())
    tot_mat     = int(df['_MAT'].sum())
    tot_agend   = int(df['_AGEND'].sum())
    tot_capa    = int(df['_CAPA'].sum())

    print(f"[{ts()}] Gerenciamento: {total} ofertas, {gerenciadas} gerenciadas, {total-com_tutor} sem tutor")
    print(f"[{ts()}] Agendas: {com_agenda} com data · amostra datas: {sorted(df[df['_TEM_AGENDA']]['_DT_AG_ISO'].head(3).tolist())}")
    print(f"[{ts()}] Gerenciamento processado: {df['_POLO'].nunique()} polos, {df['_CAT'].nunique()} categorias, {df['_ORDEM'].nunique()} ordens")

    ger_kpis = {
        'total_ofertas': total, 'ofertas_gerenciadas': gerenciadas,
        'ofertas_nao_gerenciadas': total - gerenciadas,
        'pct_gerenciado': round(gerenciadas/total*100,1) if total else 0,
        'ofertas_com_tutor': com_tutor, 'ofertas_sem_tutor': total-com_tutor,
        'pct_com_tutor': round(com_tutor/total*100,1) if total else 0,
        'ofertas_com_agenda': com_agenda,
        'total_alunos_matriculados': tot_mat, 'total_alunos_agendados': tot_agend,
        'total_capacidade': tot_capa,
        'pct_ocupacao': round(tot_agend/tot_capa*100,1) if tot_capa else 0,
        'polos_total': df['_POLO'].nunique(),
        'polos_sem_tutor': int(df[~df['_TEM_TUTOR']].groupby('_POLO').ngroups),
    }

    # Por Polo
    ger_polo = []
    for polo, grp in df.groupby('_POLO'):
        tuts = list(grp[grp['_TEM_TUTOR']]['_TUTOR'].dropna().unique())
        ger_polo.append({
            'polo': str(polo), 'total_ofertas': len(grp),
            'gerenciadas': int(grp['_GERENCIADO'].sum()),
            'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['_TEM_TUTOR'].sum()), 'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
            'com_agenda': int(grp['_TEM_AGENDA'].sum()),
            'alunos_matriculados': int(grp['_MAT'].sum()), 'alunos_agendados': int(grp['_AGEND'].sum()),
            'capacidade': int(grp['_CAPA'].sum()), 'tutores_unicos': [str(t) for t in tuts],
        })
    ger_polo.sort(key=lambda x: -x['sem_tutor'])

    # Por Categoria
    ger_cat = []
    for cat, grp in df.groupby('_CAT'):
        ger_cat.append({
            'categoria': str(cat), 'total_ofertas': len(grp),
            'gerenciadas': int(grp['_GERENCIADO'].sum()),
            'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['_TEM_TUTOR'].sum()), 'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
            'alunos_matriculados': int(grp['_MAT'].sum()), 'alunos_agendados': int(grp['_AGEND'].sum()),
        })
    ger_cat.sort(key=lambda x: -x['total_ofertas'])

    # Por Ordem
    ger_ordem = []
    ordem_sort = {'Ordem 1':1,'Ordem 2':2,'Ordem 3':3,'Ordem 4':4,'Ordem 5':5}
    for ordem in sorted(df['_ORDEM'].unique(), key=lambda x: ordem_sort.get(x,9)):
        if not ordem: continue
        grp = df[df['_ORDEM']==ordem]
        ger_ordem.append({
            'ordem': ordem, 'total_ofertas': len(grp),
            'gerenciadas': int(grp['_GERENCIADO'].sum()),
            'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['_TEM_TUTOR'].sum()),
            'alunos_matriculados': int(grp['_MAT'].sum()), 'alunos_agendados': int(grp['_AGEND'].sum()),
            'dt_inicio': '', 'dt_fim': PRAZOS_ORDENS.get(ordem,''),
        })

    # Contratação
    ger_contratacao = []
    for (polo, cat), grp in df.groupby(['_POLO','_CAT']):
        tuts = list(grp[grp['_TEM_TUTOR']]['_TUTOR'].dropna().unique())
        ger_contratacao.append({
            'polo': str(polo), 'categoria': str(cat), 'total_ofertas': len(grp),
            'tem_tutor': len(tuts)>0, 'tutores': [str(t) for t in tuts],
            'status': 'Contratado' if len(tuts)>0 else 'Sem tutor',
        })

    # Agendas por Polo
    ger_agendas = []
    for polo, grp in df.groupby('_POLO'):
        total_p = len(grp); com_ag = int(grp['_TEM_AGENDA'].sum())
        datas_por_cat = {}; datas_por_tutor = {}
        for _, row in grp[grp['_TEM_AGENDA']].iterrows():
            d = row['_DT_AG_ISO']; c = row['_CAT']; t = row['_TUTOR']
            if d:
                if d not in datas_por_cat: datas_por_cat[d]=[]
                if c and c not in datas_por_cat[d]: datas_por_cat[d].append(c)
                if d not in datas_por_tutor: datas_por_tutor[d]=[]
                if t and t not in datas_por_tutor[d]: datas_por_tutor[d].append(t)
        ger_agendas.append({
            'polo': str(polo), 'total': total_p, 'com_agenda': com_ag,
            'sem_agenda': total_p-com_ag,
            'pct_agendado': round(com_ag/total_p*100,1) if total_p else 0,
            'datas_agenda': sorted(datas_por_cat.keys()),
            'datas_por_cat': datas_por_cat, 'datas_por_tutor': datas_por_tutor,
        })
    ger_agendas.sort(key=lambda x: -x['sem_agenda'])

    # Detalhe de Ofertas
    ger_ofertas = []
    for _, row in df.iterrows():
        ger_ofertas.append({
            'polo': row['_POLO'], 'categoria': row['_CAT'],
            'ordem': row['_ORDEM'], 'pratica': row['_PRATICA'],
            'tutor': row['_TUTOR'], 'tem_tutor': bool(row['_TEM_TUTOR']),
            'tem_agenda': bool(row['_TEM_AGENDA']), 'gerenciado': bool(row['_GERENCIADO']),
            'alunos_mat': int(row['_MAT']), 'alunos_agend': int(row['_AGEND']),
            'dt_agenda': row['_DT_AG_ISO'], 'hr_agenda': row['_HR_AG'],
        })

    return {
        'ger_kpis': ger_kpis, 'ger_polo': ger_polo, 'ger_cat': ger_cat,
        'ger_ordem': ger_ordem, 'ger_contratacao': ger_contratacao,
        'ger_agendas': ger_agendas, 'ger_ofertas': ger_ofertas,
    }

def processar_gerenciamento(p3):
    """Processa REL_GERAL_DE_GERENCIAMENTO.xlsx — detecta formato novo ou antigo."""
    print(f"[{ts()}] Lendo gerenciamento...")
    df_g = ler_excel(p3)
    print(f"[{ts()}] Gerenciamento: {len(df_g)} linhas, {len(df_g.columns)} colunas")

    # ── Detectar formato: novo (LABORATORIO, NOME_EXPERIMENTO) ou antigo ─────
    cols_upper = [str(c).upper() for c in df_g.columns]
    is_novo = 'LABORATORIO' in cols_upper and 'NOME_EXPERIMENTO' in cols_upper
    if is_novo:
        print(f"[{ts()}] Formato detectado: NOVO (relatório detalhado)")
        return _processar_gerenciamento_novo(df_g)
    else:
        print(f"[{ts()}] Formato detectado: ANTIGO (GIOCONDA)")

    # ── Identificar colunas (formato antigo) ─────────────────────────────────
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
            # Extrair datas por categoria para o calendário com cores
            datas = []
            datas_por_cat = {}
            if c_dt_agenda and c_dt_agenda in grp.columns:
                for _, ag_row in grp[grp['_TEM_AGENDA']].iterrows():
                    dt_val = pd.to_datetime(ag_row.get(c_dt_agenda), errors='coerce')
                    if pd.notna(dt_val):
                        dt_str = dt_val.strftime('%Y-%m-%d')
                        cat_val = str(ag_row.get(c_cat, '') or '')
                        if dt_str not in datas:
                            datas.append(dt_str)
                        if cat_val:
                            if dt_str not in datas_por_cat:
                                datas_por_cat[dt_str] = []
                            if cat_val not in datas_por_cat[dt_str]:
                                datas_por_cat[dt_str].append(cat_val)
                datas = sorted(set(datas))
            ger_agendas.append({
                'polo': str(polo),
                'total': total,
                'com_agenda': com_agenda,
                'sem_agenda': sem_agenda,
                'pct_agendado': round(com_agenda / total * 100, 1) if total else 0,
                'datas_agenda': datas,
                'datas_por_cat': datas_por_cat,
            })
        ger_agendas.sort(key=lambda x: -x['sem_agenda'])

    # ── Tabela detalhada de ofertas (top 500 para não sobrecarregar o HTML) ───
    ger_ofertas_detalhe = []
    cols_detalhe = [c_polo, c_cat, '_ORDEM_G', '_PRATICA_G', c_curso, c_tutor,
                    '_GERENCIADO', '_TEM_AGENDA', '_ALUNOS_MAT', '_QTD_ALUN', '_CAPA']
    # Sem limite — o template agrupa e pagina automaticamente
    for _, row in df_g.iterrows():
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

    p1, p2, tmpl, p3, p4 = verificar_e_localizar()

    if not p1 or not p2 or not os.path.isfile(tmpl):
        print()
        print(" Coloque as planilhas na pasta planilhas\\")
        print(" e tente novamente.")
        print()
        if '--sem-browser' not in sys.argv:
            input(" Pressione Enter para sair...")
        sys.exit(1)

    print()
    dados = processar(p1, p2)

    # Lotação de tutores — enriquece dados com perfil, cursos, CH
    if p4:
        try:
            lotacao = carregar_lotacao(p4)
            dados = enriquecer_tutores(dados, lotacao)
            dados['tem_lotacao'] = True
        except Exception as e:
            print(f"[{ts()}] AVISO: Erro ao processar lotação: {e}")
            dados['tem_lotacao'] = False
    else:
        dados['tem_lotacao'] = False

        # Gerenciamento — detecta formato automaticamente pelo cabeçalho
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
