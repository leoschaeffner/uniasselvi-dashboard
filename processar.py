"""
UNIASSELVI - Dashboard de Portfolios
Le as planilhas da pasta planilhas/ e gera saida/dashboard.html

PATCHES v2 aplicados:
  1. CH SEMANAL lida de 01_CONTROLE_TUTORIA.xlsx (não depende de LOTACAO)
  2. tem_lotacao = True somente quando há ch_semanal > 0 nos tutores
  3. status_ordem usa PERIODOS_ORDENS para datas de início corretas
  4. Situação do tutor corrigida quando nenhuma ordem está vencida ainda
  5. Campo 'sit' sincronizado após deduplicação por email
  6. Prints duplicados removidos de verificar_e_localizar()
  7. datas_por_tutor adicionado no formato antigo de agendas
"""

import pandas as pd
import json, os, sys, math, webbrowser, time, threading, glob, hashlib, base64, unicodedata
from pathlib import Path
from datetime import datetime, timezone, timedelta
from collections import defaultdict
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# ── Configuração multi-semestre ─────────────────────────────────────────────
# Lida de config_semestre.json. Fallback hardcoded para 2026/1.
_SEMESTRES_DEFAULT = {
    '2026/1': {
        'prazos': {
            'Ordem 1': '14/03/2026', 'Ordem 2': '11/04/2026',
            'Ordem 3': '09/05/2026', 'Ordem 4': '06/06/2026', 'Ordem 5': '04/07/2026',
        },
        'periodos': {
            'Ordem 1': {'inicio': '16/02/2026', 'fim': '14/03/2026', 'semanas': 4},
            'Ordem 2': {'inicio': '16/03/2026', 'fim': '11/04/2026', 'semanas': 4},
            'Ordem 3': {'inicio': '13/04/2026', 'fim': '09/05/2026', 'semanas': 4},
            'Ordem 4': {'inicio': '11/05/2026', 'fim': '06/06/2026', 'semanas': 4},
            'Ordem 5': {'inicio': '08/06/2026', 'fim': '04/07/2026', 'semanas': 4},
        },
    },
    '2026/2': {
        'prazos': {
            'Ordem 1': '22/08/2026', 'Ordem 2': '19/09/2026',
            'Ordem 3': '17/10/2026', 'Ordem 4': '14/11/2026', 'Ordem 5': '12/12/2026',
        },
        'periodos': {
            'Ordem 1': {'inicio': '27/07/2026', 'fim': '22/08/2026', 'semanas': 4},
            'Ordem 2': {'inicio': '24/08/2026', 'fim': '19/09/2026', 'semanas': 4},
            'Ordem 3': {'inicio': '21/09/2026', 'fim': '17/10/2026', 'semanas': 4},
            'Ordem 4': {'inicio': '19/10/2026', 'fim': '14/11/2026', 'semanas': 4},
            'Ordem 5': {'inicio': '16/11/2026', 'fim': '12/12/2026', 'semanas': 4},
        },
    },
}

_DISCIPLINAS_POR_ORDEM_GLOBAL = {}
def _carregar_semestres():
    import os as _os, json as _json
    _cfg_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'config_semestre.json')
    _sems = dict(_SEMESTRES_DEFAULT)
    if _os.path.isfile(_cfg_path):
        try:
            with open(_cfg_path, encoding='utf-8') as _f: _cfg = _json.load(_f)
            # config pode ter 'semestres' (dict) ou o formato antigo (semestre único)
            if 'semestres' in _cfg:
                _sems.update(_cfg['semestres'])
            elif 'semestre' in _cfg:
                _sem_key = _cfg['semestre']
                _sems[_sem_key] = {'prazos': _cfg.get('prazos', {}), 'periodos': _cfg.get('periodos', {})}
            print(f"  [CONFIG] Semestres carregados: {sorted(_sems.keys())}")
            # Carregar mapeamento de disciplinas por ordem (opcional)
            _disc_file = _cfg.get('disciplinas_por_ordem')
            if _disc_file:
                _disc_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), _disc_file)
                if _os.path.isfile(_disc_path):
                    with open(_disc_path, encoding='utf-8') as _fd:
                        _DISCIPLINAS_POR_ORDEM_GLOBAL.update(_json.load(_fd))
                    _total_disc = sum(len(v) for ordens in _DISCIPLINAS_POR_ORDEM_GLOBAL.values() for v in ordens.values())
                    print(f"  [CONFIG] Disciplinas por ordem: {_total_disc} total")
        except Exception as _e:
            print(f"  [AVISO] Erro ao ler config_semestre.json: {_e} — usando padrão")
    return _sems

ALL_SEMESTRES = _carregar_semestres()

def _data_para_semestre(data_str):
    """Dado '2026-03-15', retorna '2026/1' ou '2026/2' ou None."""
    if not data_str or data_str == 'None': return None
    try:
        from datetime import datetime as _dt
        d = _dt.strptime(str(data_str)[:10], '%Y-%m-%d').date()
        for sem, cfg in sorted(ALL_SEMESTRES.items()):
            for ord_cfg in cfg.get('periodos', {}).values():
                try:
                    ini = _dt.strptime(ord_cfg['inicio'], '%d/%m/%Y').date()
                    fim = _dt.strptime(ord_cfg['fim'],    '%d/%m/%Y').date()
                    if ini <= d <= fim: return sem
                except: pass
    except: pass
    return None

def _ordem_relativa(ordem_forms, semestre):
    """Ordem 1 no Forms sempre = Ordem 1 do semestre em questão."""
    return ordem_forms  # As ordens são relativas dentro de cada semestre

# Para compatibilidade com o resto do código — usa o semestre mais recente como padrão
_sem_atual = sorted(ALL_SEMESTRES.keys())[-1]
PRAZOS_ORDENS   = ALL_SEMESTRES[_sem_atual]['prazos']
PERIODOS_ORDENS = ALL_SEMESTRES[_sem_atual]['periodos']
SEMESTRE_ATUAL  = _sem_atual
print(f"  [CONFIG] Semestre ativo (dashboard): {SEMESTRE_ATUAL}")
# ── Fim configuração multi-semestre ──────────────────────────────────────────
CH_ADMIN_FATOR   = 0.25
CH_PRATICA_DURAC = 1.5


def _parse_ch(v):
    """PATCH 1 helper — converte CH SEMANAL (HH:MM ou decimal) para float horas."""
    if v is None: return None
    sv = str(v).strip()
    if sv in ('', 'nan', 'NaN', 'None', '0', '0.0'): return None
    try:
        if ':' in sv:
            parts = sv.split(':')
            result = float(parts[0]) + float(parts[1]) / 60
        else:
            result = float(sv.replace(',', '.'))
        return result if result > 0 else None
    except (ValueError, TypeError):
        return None


def achar_pasta_script():
    candidatos = []
    try:
        p = os.path.dirname(os.path.abspath(sys.argv[0]))
        if os.path.isdir(p): candidatos.append(p)
    except: pass
    try:
        p = os.path.dirname(os.path.abspath(__file__))
        if os.path.isdir(p): candidatos.append(p)
    except: pass
    try:
        p = os.getcwd()
        if os.path.isdir(p): candidatos.append(p)
    except: pass
    for p in candidatos:
        if os.path.isdir(os.path.join(p, "planilhas")): return p
        if os.path.isfile(os.path.join(p, "processar.py")): return p
    return candidatos[0] if candidatos else os.getcwd()


SCRIPT_DIR = achar_pasta_script()


def ler_url_file(path_url):
    try:
        with open(path_url, encoding='utf-8', errors='replace') as f:
            for line in f:
                if line.upper().startswith('URL='): return line[4:].strip()
    except: pass
    return None


def forcar_download_onedrive(path_url_file, destino, label):
    import subprocess, shutil, time
    path_xlsx = path_url_file.replace('.url', '').replace('.URL', '')
    try:
        subprocess.run(['attrib', '-P', '+U', path_url_file], capture_output=True, timeout=10)
    except Exception: pass
    for _ in range(6):
        if os.path.isfile(path_xlsx):
            with open(path_xlsx, 'rb') as f:
                header = f.read(4)
            if header == b'PK\x03\x04':
                shutil.copy2(path_xlsx, destino)
                print(f"  [OneDrive] Sincronizado: {label}")
                return destino
        time.sleep(5)
        print(f"  [OneDrive] Aguardando sync para {label}...")
    print(f"  [OneDrive] Timeout aguardando {label}.")
    return None


_KEYWORDS = {
    '01_CONTROLE_TUTORIA.xlsx': ['CONTROLE'],
    'PORTFOLIO_TUTOR.xlsx':     ['PORTFOLIO', 'PORTIFOLIO', 'PORTF'],
    'REL_GERAL_DE_GERENCIAMENTO.xlsx': ['GERENCIAMENTO', 'REL_GERAL'],
}
_ONEDRIVE_NAMES = {
    '01_CONTROLE_TUTORIA.xlsx': ['CONTROLE'],
    'PORTFOLIO_TUTOR.xlsx':     ['PORTF', 'PORTFOLIO'],
}


def _bate(caminho_arq, padrao):
    bn  = os.path.basename(caminho_arq).upper()
    kws = _KEYWORDS.get(padrao, [os.path.splitext(padrao)[0].upper()])
    return any(kw in bn for kw in kws)


def achar_arquivo(pasta, padrao):
    pasta_planilhas = os.path.join(pasta, "planilhas")
    direto = os.path.join(pasta_planilhas, padrao)
    if os.path.isfile(direto): return direto
    for arq in glob.glob(os.path.join(pasta_planilhas, "*.xls")) + glob.glob(os.path.join(pasta_planilhas, "*.xlsx")):
        if _bate(arq, padrao): return arq
    for arq in glob.glob(os.path.join(pasta_planilhas, "*.url")) + glob.glob(os.path.join(pasta_planilhas, "*.xlsx.url")) + glob.glob(os.path.join(pasta_planilhas, "*.xls.url")):
        if _bate(arq, padrao):
            url = ler_url_file(arq)
            if url:
                destino = os.path.join(pasta_planilhas, padrao)
                resultado = forcar_download_onedrive(arq, destino, padrao)
                if resultado: return resultado
    usuario = os.environ.get('USERNAME', os.environ.get('USER', 'leona'))
    for base in [
        f"C:\\Users\\{usuario}\\OneDrive - Uniasselvi",
        f"C:\\Users\\{usuario}\\OneDrive - UNIASSELVI",
        f"C:\\Users\\{usuario}\\OneDrive - Grupo Uniasselvi",
        f"C:\\Users\\{usuario}\\OneDrive",
    ]:
        if not os.path.isdir(base): continue
        for arq in glob.glob(os.path.join(base, "*.url")):
            if _bate(arq, padrao):
                url = ler_url_file(arq)
                if url:
                    destino = os.path.join(pasta_planilhas, padrao)
                    resultado = forcar_download_onedrive(arq, destino, padrao)
                    if resultado: return resultado
        for arq in glob.glob(os.path.join(base, "**", "*.xls"), recursive=True) + glob.glob(os.path.join(base, "**", "*.xlsx"), recursive=True):
            if _bate(arq, padrao): return arq
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
    cfg = {}
    cfg_file = os.path.join(SCRIPT_DIR, "config_links.json")
    if os.path.isfile(cfg_file):
        try:
            with open(cfg_file, encoding="utf-8") as f: cfg = json.load(f)
        except: pass
    cam_t = cfg.get("caminho_planilha_tutores", "").strip().strip('"')
    cam_p = cfg.get("caminho_planilha_portfolio", "").strip().strip('"')
    if cam_t and os.path.isfile(cam_t):
        p1 = cam_t; print(f"  [OK] {os.path.basename(p1)}")
    else:
        p1 = achar_arquivo(SCRIPT_DIR, "01_CONTROLE_TUTORIA.xlsx")
        if p1: print(f"  [OK] {os.path.basename(p1)}")
        else:  print(f"  [FALTA] 01_CONTROLE_TUTORIA.xlsx")
    if cam_p and os.path.isfile(cam_p):
        p2 = cam_p; print(f"  [OK] {os.path.basename(p2)}")
    else:
        p2 = achar_arquivo(SCRIPT_DIR, "PORTFOLIO_TUTOR.xlsx")
        if p2: print(f"  [OK] {os.path.basename(p2)}")
        else:  print(f"  [FALTA] PORTFOLIO_TUTOR.xlsx")
    # PATCH 10: planilha nova de portfólios 2026/2 (formulário customizado, schema próprio)
    p2b = achar_arquivo(SCRIPT_DIR, "PORTIFOLIO_TUTOR_2026_2.xlsx")
    if p2b: print(f"  [OK] {os.path.basename(p2b)}")
    else:   print(f"  [INFO] PORTIFOLIO_TUTOR_2026_2.xlsx não encontrada (ainda sem envios 2026/2?)")
    # PATCH 6: prints duplicados de p1/p2 removidos aqui
    tmpl = os.path.join(SCRIPT_DIR, "template_dashboard.html")
    if os.path.isfile(tmpl): print(f"  [OK] template_dashboard.html")
    else:                    print(f"  [FALTA] template_dashboard.html")
    p3 = achar_arquivo(SCRIPT_DIR, "REL_GERAL_DE_GERENCIAMENTO.xlsx")
    if p3: print(f"  [OK] {os.path.basename(p3)}")
    else:  print(f"  [INFO] REL_GERAL_DE_GERENCIAMENTO.xlsx não encontrada (módulo desativado)")
    # PATCH 18: planilha de gerenciamento específica de 2026/2 (export novo, CSV)
    p3b = achar_arquivo(SCRIPT_DIR, "REL_GERAL_DE_GERENCIAMENTO_26_02.csv")
    if p3b: print(f"  [OK] {os.path.basename(p3b)}")
    else:   print(f"  [INFO] REL_GERAL_DE_GERENCIAMENTO_26_02.csv não encontrada")
    p4 = achar_arquivo(SCRIPT_DIR, "LOTACAO_TUTORES.xlsm") or achar_arquivo(SCRIPT_DIR, "LOTACAO_TUTORES.xlsx")
    if p4: print(f"  [OK] {os.path.basename(p4)}")
    else:  print(f"  [INFO] LOTACAO_TUTORES não encontrada (.xlsx/.xlsm)")
    # ── CSV de alunos por hub (igual aos outros arquivos: URL no Secret/env) ──
    p5 = None
    # 1. Tentar achar na pasta planilhas/ (já baixado anteriormente)
    p5 = achar_arquivo(SCRIPT_DIR, "Relatorio_alunos_por_hub.csv")
    if p5:
        print(f"  [OK] {os.path.basename(p5)}")
    else:
        # 2. Tentar baixar via variável de ambiente URL_ALUNOS_HUB (Secret GitHub)
        import re
        url_hub = os.environ.get("URL_ALUNOS_HUB", "").strip()
        if url_hub:
            print(f"  [Baixando] Relatorio_alunos_por_hub.csv via URL_ALUNOS_HUB...")
            try:
                import urllib.request
                # Converter link SharePoint/OneDrive para download direto
                # Tentar múltiplos formatos de URL
                def _build_dl_urls(url):
                    urls = []
                    if 'sharepoint.com' in url:
                        # Formato 1: adicionar &download=1
                        sep = '&' if '?' in url else '?'
                        urls.append(url + sep + 'download=1')
                        # Formato 2: download.aspx com token
                        m = re.search(r'/([A-Za-z0-9_-]{20,})[?]', url)
                        if m:
                            base = re.match(r'(https://[^/]+)', url).group(1)
                            user = re.search(r'/personal/([^/]+)/', url)
                            if user:
                                urls.append(f"{base}/personal/{user.group(1)}/_layouts/15/download.aspx?share={m.group(1)}")
                    elif '1drv.ms' in url:
                        sep = '&' if '?' in url else '?'
                        urls.append(url + sep + 'download=1')
                    urls.append(url)  # URL original como último recurso
                    return urls

                dest = os.path.join(pasta_planilhas, "Relatorio_alunos_por_hub.csv")
                downloaded = False
                for url_dl in _build_dl_urls(url_hub):
                    try:
                        req = urllib.request.Request(url_dl, headers={
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
                        with urllib.request.urlopen(req, timeout=120) as r:
                            data = r.read()
                        if len(data) > 10000 and b'<!DOCTYPE' not in data[:500]:
                            with open(dest, 'wb') as f_out: f_out.write(data)
                            p5 = dest
                            print(f"  [OK] Relatorio_alunos_por_hub.csv ({len(data):,} bytes)")
                            downloaded = True
                            break
                        else:
                            print(f"  [AVISO] URL retornou conteúdo inválido ({len(data)} bytes): {url_dl[:80]}")
                    except Exception as ex:
                        print(f"  [AVISO] Erro ao baixar: {ex} | URL: {url_dl[:80]}")
                if not downloaded:
                    print(f"  [ERRO] Não foi possível baixar o CSV de alunos — verifique URL_ALUNOS_HUB")
            except Exception as e:
                print(f"  [ERRO] Não foi possível baixar CSV de alunos: {e}")
        else:
            print(f"  [INFO] Relatorio_alunos_por_hub.csv não encontrado (defina URL_ALUNOS_HUB)")
    return p1, p2, tmpl, p3, p3b, p4, p5


def ler_excel(path, **kwargs):
    for engine in ('openpyxl', 'xlrd', None):
        try:
            kw = dict(kwargs)
            if engine: kw['engine'] = engine
            return pd.read_excel(path, **kw)
        except Exception: continue
    raise ValueError(f"Não foi possível ler {path} com nenhum engine disponível")


def _ler_arquivo_gerenciamento(path):
    # PATCH 18: REL_GERAL_DE_GERENCIAMENTO pode vir como .xlsx (export antigo) ou
    # .csv (export novo, ISO-8859-1, separador ';', tudo entre aspas)
    if str(path).lower().endswith('.csv'):
        last_err = None
        for enc in ('latin-1', 'utf-8', 'cp1252'):
            try:
                df = pd.read_csv(path, sep=';', encoding=enc, dtype=str)
                if len(df.columns) > 1: return df
            except Exception as e:
                last_err = e
        raise ValueError(f"Não foi possível ler CSV de gerenciamento ({path}): {last_err}")
    return ler_excel(path)


def processar(p1, p2):
    print(f"[{ts()}] Lendo tutores...")
    with open(p1, 'rb') as _f:
        _magic = _f.read(8); _preview = _f.read(200)
    print(f"  [DEBUG] Magic bytes de {p1}: {_magic.hex()}")
    if _magic[:2] == b'PK': print(f"  [DEBUG] Formato ZIP/XLSX confirmado")
    elif _magic[:2] in (b'\xd0\xcf', b'\xCF\xD0'): print(f"  [DEBUG] Formato XLS (OLE2) confirmado")
    else:
        print(f"  [ERRO] Arquivo não é Excel válido. Conteúdo inicial:")
        print((_magic + _preview).decode('utf-8', errors='replace')[:300])
        raise ValueError(f"Arquivo {p1} não é um Excel válido — SharePoint pode ter retornado HTML de erro")
    df_t = ler_excel(p1, sheet_name='Base de Tutores', header=1)
    col_sit  = next((c for c in df_t.columns if 'SITUA' in str(c).upper()), None)
    col_nome = next((c for c in df_t.columns if 'NOME'  in str(c).upper() and 'TUTOR' in str(c).upper()), None)
    # Busca flexível de colunas no CONTROLE_TUTORIA
    col_polo = next((c for c in df_t.columns if str(c).strip().upper() == 'POLO'), None) or                next((c for c in df_t.columns if 'POLO' in str(c).upper() and 'HUB' not in str(c).upper()), None) or                next((c for c in df_t.columns if 'POLO' in str(c).upper()), None) or 'POLO'
    col_cur  = next((c for c in df_t.columns if str(c).strip().upper() == 'CURSOS'), None) or                next((c for c in df_t.columns if 'CURSO' in str(c).upper() and 'VINC' not in str(c).upper()), None) or                next((c for c in df_t.columns if 'CURSO' in str(c).upper()), None) or 'CURSOS'
    col_email= next((c for c in df_t.columns if 'E-MAIL' in str(c).upper() or 'EMAIL' in str(c).upper()), None)
    print(f"[{ts()}] CONTROLE colunas detectadas: polo='{col_polo}' cursos='{col_cur}' email='{col_email}'")
    print(f"[{ts()}] CONTROLE todas colunas: {list(df_t.columns[:20])}")
    col_cat    = next((c for c in df_t.columns if 'CATEGORIA' in str(c).upper()), None)
    col_inicio = next((c for c in df_t.columns if str(c).upper().strip() in ('INÍCIO','INICIO')), None)
    col_whats  = next((c for c in df_t.columns if 'WHATSAPP' in str(c).upper()), None)
    col_chapa  = next((c for c in df_t.columns if 'CHAPA' in str(c).upper()), None)
    # PATCH 1: Detectar coluna CH SEMANAL na planilha de controle
    col_ch = next((c for c in df_t.columns if str(c).upper().strip() == 'CH SEMANAL' or
                   ('CH' in str(c).upper() and 'SEMAL' in str(c).upper())), None)
    if col_ch:
        print(f"[{ts()}] CH SEMANAL encontrada: '{col_ch}'")
        ch_vals = df_t[col_ch].dropna()
        print(f"[{ts()}] Amostra CH SEMANAL: {list(ch_vals.head(5))}")
    else:
        print(f"[{ts()}] CH SEMANAL não encontrada — colunas CH disponíveis: {[c for c in df_t.columns if 'CH' in str(c).upper()]}")
    # Filtrar tutores ativos: incluir Ativo + afastamentos temporários (Licença Maternidade, etc.)
    # Excluir apenas Inativo e Desligado explicitamente
    _SITUACOES_EXCLUIR = {'inativo', 'desligado', 'rescindido', 'demitido', 'encerrado',
                           'admissão prox.mês', 'admissao prox.mes', 'em admissão',
                           'pendente', 'aguardando'}
    if col_sit:
        _sit_norm = df_t[col_sit].astype(str).str.strip().str.lower()
        df_at = df_t[~_sit_norm.isin(_SITUACOES_EXCLUIR)].copy()
        _contagem = df_at[col_sit].value_counts().to_dict()
        print(f"[{ts()}] Situações incluídas: {_contagem}")
    else:
        df_at = df_t.copy()
    # Usar CHAVE LOTAÇÃO do CONTROLE diretamente (mesma chave usada pelo Forms)
    col_chave_lot = next((c for c in df_t.columns if 'CHAVE' in str(c).upper() and 'LOTA' in str(c).upper()
                          and 'CLASSIF' not in str(c).upper()), None)
    if col_chave_lot:
        df_at['_CHAVE'] = df_at[col_chave_lot].astype(str).str.strip()
        print(f"[{ts()}] CONTROLE usando coluna '{col_chave_lot}' como chave")
    else:
        df_at['_CHAVE'] = df_at[col_polo].astype(str).str.strip() + df_at[col_cur].astype(str).str.strip()
        print(f"[{ts()}] CONTROLE chave construída de POLO + CURSOS")
    _sample_chaves = df_at['_CHAVE'].dropna().head(8).tolist()
    print(f"[{ts()}] CONTROLE chaves (amostra): {_sample_chaves}")
    print(f"[{ts()}] Lendo portfolios...")
    df_p = ler_excel(p2, sheet_name='Sheet1')


    def col(df, *partes):
        for c in df.columns:
            cu = str(c).upper()
            if all(p.upper() in cu for p in partes): return c
        return None
    c_chave = col(df_p, 'CHAVE', 'LINK')
    c_proto = col(df_p, 'PROTOCOLOS', 'ATIVIDADES')
    for sfx in (':7', ':8', ':6', ':9'):
        proto_sfx = [c for c in df_p.columns if 'PROTOCOLOS' in str(c).upper() and str(c).endswith(sfx)]
        if proto_sfx: c_proto = proto_sfx[0]; break
    else:
        proto_any = [c for c in df_p.columns if 'PROTOCOLOS' in str(c).upper() or 'ATIVIDADE' in str(c).upper()]
        if proto_any: c_proto = proto_any[0]
    data_cols = [c for c in df_p.columns if 'DATA DA APLICA' in str(c).upper() and str(c).endswith(':7')]
    if not data_cols: data_cols = [c for c in df_p.columns if 'DATA DA APLICA' in str(c).upper()]
    if not data_cols: data_cols = [c for c in df_p.columns if 'DATA' in str(c).upper() and 'APLICA' in str(c).upper()]
    c_data = data_cols[0] if data_cols else None
    def find_aluno_col(df):
        cols = df.columns.tolist()
        for suffix_end in ('72', '73', '74', '70', '71', '75'):
            for c in cols:
                if 'ESTUDANTES' in str(c).upper() and str(c).endswith(suffix_end): return c
        for c in cols:
            if 'ESTUDANTES' in str(c).upper() and 'PONTOS' not in str(c).upper(): return c
        for c in cols:
            cu = str(c).upper()
            if ('ALUNO' in cu or 'ALUNOS' in cu) and 'PONTOS' not in cu and 'COMENT' not in cu: return c
        for c in cols:
            cu = str(c).upper()
            if 'QUANTIDADE' in cu or 'QTD' in cu: return c
        return None
    c_aluno = find_aluno_col(df_p)
    cat_cols = [c for c in df_p.columns if 'CATEGORIA' in str(c).upper() and 'PONTOS' not in str(c).upper() and 'COMENT' not in str(c).upper()]
    c_cat = cat_cols[0] if cat_cols else None
    print(f"[{ts()}] Colunas: chave={c_chave}, proto={c_proto}, data={c_data}, alunos={c_aluno}, cat={c_cat}")
    c_ordem_cols = [c for c in df_p.columns if 'ORDEM' in str(c).upper() and 'PONTOS' not in str(c).upper() and 'COMENT' not in str(c).upper()]
    c_ordem = c_ordem_cols[0] if c_ordem_cols else None
    print(f"[{ts()}] Coluna ordem: {c_ordem}")
    df_p['_CHAVE']  = df_p[c_chave].astype(str).str.strip() if c_chave else ''
    df_p['_PROTO']  = df_p[c_proto].astype(str).str.strip() if c_proto else ''
    df_p['_DATA']   = pd.to_datetime(df_p[c_data], errors='coerce') if c_data else pd.NaT
    df_p['_ALUNOS'] = pd.to_numeric(df_p[c_aluno], errors='coerce').fillna(0).astype(int) if c_aluno else 0
    df_p['_CAT']    = df_p[c_cat].astype(str).str.strip() if c_cat else ''
    df_p['_ORDEM']  = df_p[c_ordem].astype(str).str.strip() if c_ordem else 'Ordem 1'
    # ── MEC Cache ────────────────────────────────────────────────────────────
    mec_cache = {}
    mec_file = os.path.join(SCRIPT_DIR, 'mec_cache.json')
    if os.path.isfile(mec_file):
        with open(mec_file, encoding='utf-8') as f: mec_cache = json.load(f)
        print(f"[{ts()}] MEC cache: {len(mec_cache)} tutores")
    # ── Fim MEC Cache ─────────────────────────────────────────────────────────

    catalogo_oficial = {}
    id_to_perfil = {}  # PATCH 10: id da prática -> código de perfil (ex: '206'->'EMF-ISN')
    cat_file = os.path.join(SCRIPT_DIR, 'catalogo_oficial.json')
    if os.path.isfile(cat_file):
        with open(cat_file, encoding='utf-8') as f: raw = json.load(f)
        for cat_nome, praticas in raw.items():
            if isinstance(praticas, list) and praticas:
                if isinstance(praticas[0], dict):
                    catalogo_oficial[cat_nome] = sorted(set(p['nome'] for p in praticas))
                    for p in praticas:
                        if p.get('id') and p.get('perfil'): id_to_perfil.setdefault(str(p['id']), p['perfil'])
                else: catalogo_oficial[cat_nome] = sorted(set(praticas))
        print(f"[{ts()}] Catalogo oficial (JSON): {len(catalogo_oficial)} categorias")
    # PATCH 10: reforça id_to_perfil com mapa verificado direto do catálogo do formulário
    # (garante a correspondência mesmo se catalogo_oficial.json não tiver id/perfil)
    idp_file = os.path.join(SCRIPT_DIR, 'id_to_perfil.json')
    if os.path.isfile(idp_file):
        with open(idp_file, encoding='utf-8') as f: idp_extra = json.load(f)
        for k, v in idp_extra.items(): id_to_perfil.setdefault(k, v)
        print(f"[{ts()}] id_to_perfil: {len(id_to_perfil)} práticas mapeadas")

    # PATCH 15: nome da prática -> perfil correto (só pra códigos ambíguos BFI/BTO/COS-TIP)
    def _norm_proto(s):
        # normaliza unicode (resolve variantes de hífen/travessão), colapsa espaços
        s = unicodedata.normalize('NFKC', str(s or ''))
        s = s.replace('–', '-').replace('—', '-')  # en-dash/em-dash -> hífen comum
        return ' '.join(s.split()).strip()
    NOME_TO_PERFIL = {}
    nomep_file = os.path.join(SCRIPT_DIR, 'nome_to_perfil.json')
    if os.path.isfile(nomep_file):
        with open(nomep_file, encoding='utf-8') as f: _ntp_raw = json.load(f)
        NOME_TO_PERFIL = {_norm_proto(k): v for k, v in _ntp_raw.items()}
        print(f"[{ts()}] nome_to_perfil: {len(NOME_TO_PERFIL)} práticas (correção BFI/BTO/COS-TIP)")
    if not catalogo_oficial:
        cat_xlsx = achar_arquivo(SCRIPT_DIR, 'CATALOGO_EXPERIMENTOS.xlsx')
        if not cat_xlsx:
            for f in os.listdir(os.path.join(SCRIPT_DIR, 'planilhas')) if os.path.isdir(os.path.join(SCRIPT_DIR, 'planilhas')) else []:
                fu = f.upper()
                if ('RELAT' in fu and 'EXPER' in fu) or ('CATALOGO' in fu and 'EXPER' in fu):
                    cat_xlsx = os.path.join(SCRIPT_DIR, 'planilhas', f); break
        if cat_xlsx and os.path.isfile(cat_xlsx):
            try:
                df_cat = pd.read_excel(cat_xlsx)
                c_cat_nome = next((c for c in df_cat.columns if 'CATEGORIA' in str(c).upper()), None)
                c_exp_nome = next((c for c in df_cat.columns if 'EXPERIMENTO' in str(c).upper() or 'NOME' in str(c).upper()), None)
                c_sit = next((c for c in df_cat.columns if 'SITUA' in str(c).upper()), None)
                if c_cat_nome and c_exp_nome:
                    if c_sit: df_cat = df_cat[df_cat[c_sit].astype(str).str.strip().str.upper() == 'ATIVO']
                    for cat_val, grp in df_cat.groupby(c_cat_nome):
                        cat_str = str(cat_val).strip()
                        if cat_str and cat_str != 'nan':
                            nomes = sorted(set(str(n).strip() for n in grp[c_exp_nome].dropna() if str(n).strip() and str(n).strip() != 'nan'))
                            if nomes: catalogo_oficial[cat_str] = nomes
                    print(f"[{ts()}] Catalogo oficial (Excel): {len(catalogo_oficial)} categorias, {sum(len(v) for v in catalogo_oficial.values())} práticas")
            except Exception as e: print(f"[{ts()}] AVISO: Erro ao ler catálogo Excel: {e}")

    # PATCH 10: ler e mesclar PORTIFOLIO_TUTOR_2026_2.xlsx (formulário novo, schema próprio)
    p2b = achar_arquivo(SCRIPT_DIR, "PORTIFOLIO_TUTOR_2026_2.xlsx")
    if p2b:
        try:
            df_novo = ler_excel(p2b, sheet_name='PORTIFOLIOS')
        except Exception:
            df_novo = ler_excel(p2b, sheet_name=0)
        df_novo.columns = [str(c).strip().upper() for c in df_novo.columns]
        if len(df_novo):
            def _g(col): return df_novo[col] if col in df_novo.columns else ''
            _protoid = _g('PROTOCOLO_ID').astype(str).str.strip()
            df_novo['_CHAVE']  = _g('POLO').astype(str).str.strip() + _protoid.map(id_to_perfil).fillna('')
            df_novo['_PROTO']  = _g('PROTOCOLO_NOME').astype(str).str.strip()
            df_novo['_DATA']   = pd.to_datetime(_g('DATA_APLICACAO'), errors='coerce')
            df_novo['_ALUNOS'] = pd.to_numeric(_g('QTD_ESTUDANTES'), errors='coerce').fillna(0).astype(int)
            df_novo['_CAT']    = _g('CATEGORIA_LAB').astype(str).str.strip()
            df_novo['_ORDEM']  = _g('ORDEM_DISCIPLINA').astype(str).str.strip().replace('', 'Ordem 1')
            df_novo['EMAIL']      = _g('EMAIL_TUTOR').astype(str).str.strip()  # nome exato p/ busca col_email_p
            df_novo['NOME_TUTOR'] = _g('NOME_TUTOR').astype(str).str.strip()   # contém NOME+TUTOR p/ busca col_nome_tutor_p
            _sem_perfil = int(_protoid.map(id_to_perfil).isna().sum())
            if _sem_perfil: print(f"[{ts()}] AVISO: {_sem_perfil} envios em PORTIFOLIO_TUTOR_2026_2 sem PROTOCOLO_ID mapeado em id_to_perfil")
            df_p = pd.concat([df_p, df_novo], ignore_index=True)
            print(f"[{ts()}] PORTIFOLIO_TUTOR_2026_2: {len(df_novo)} envios mesclados (2026/2)")
        else:
            print(f"[{ts()}] PORTIFOLIO_TUTOR_2026_2: 0 envios ainda")
    else:
        print(f"[{ts()}] PORTIFOLIO_TUTOR_2026_2.xlsx não encontrada — só 2026/1 nesta rodada")

    chave_to_cat_raw = {}; chave_to_cf = {}; chave_alias = {}
    polo_biofar_cursos = {}
    for _, t in df_at.iterrows():
        polo_   = str(t.get(col_polo, '') or '').strip()
        cursos_ = str(t.get(col_cur,  '') or '').strip()
        cat_    = str(t.get(col_cat,  '') or '').strip() if col_cat else ''
        if cursos_ in ('BBI', 'BFR') and 'BIO-FAR' in cat_.upper():
            if polo_ not in polo_biofar_cursos: polo_biofar_cursos[polo_] = set()
            polo_biofar_cursos[polo_].add(cursos_)
    # Pré-calcular polos que têm tutor BFI legítimo (para não criar alias conflitante)
    _polos_com_bfi = set()
    for _, t in df_at.iterrows():
        _cur_t = str(t.get(col_cur, '') or '').strip()
        _pol_t = str(t.get(col_polo, '') or '').strip()
        if _cur_t == 'BFI':
            _polos_com_bfi.add(_pol_t)
    for _, t in df_at.iterrows():
        polo = str(t.get(col_polo, '') or '').strip()
        cursos = str(t.get(col_cur, '') or '').strip()
        cat_raw = str(t.get(col_cat, '') or '').strip() if col_cat else ''
        cf = CAT_MAP.get(cat_raw, cat_raw)
        chave = polo + cursos
        if chave and cat_raw:
            chave_to_cat_raw[chave] = cat_raw
            chave_to_cf[chave] = cf
            if cursos in ('BBI', 'BFR'):
                outros = polo_biofar_cursos.get(polo, set()) - {cursos}
                # Só adicionar alias BFI se NÃO existe tutor BFI legítimo neste polo
                # (evita colisão que faz portfólios BFI serem atribuídos a BIO-FAR)
                variantes = [] if polo in _polos_com_bfi else [polo + 'BFI']
                for outro in outros:
                    variantes += [polo+cursos+'-'+outro, polo+outro+'-'+cursos, polo+cursos+outro, polo+outro+cursos]
                for v in variantes:
                    chave_to_cf.setdefault(v, cf)
                    chave_alias.setdefault(v, chave)


    oficial_p_to_cat = {}
    for cat, pracs in catalogo_oficial.items():
        for p in pracs: oficial_p_to_cat.setdefault(p, cat)
    catalogo_real = defaultdict(set)
    for _, r in df_p.iterrows():
        chave = str(r.get('_CHAVE', '') or '').strip()
        chave = chave_alias.get(chave, chave)
        proto = r['_PROTO']
        if not chave or chave == 'nan' or not proto or proto == 'nan': continue
        cf = chave_to_cf.get(chave, '')
        if not cf: continue
        for p in proto.split(';'):
            p = p.strip()
            if not p: continue
            cat_oficial = oficial_p_to_cat.get(p)
            if cat_oficial and cat_oficial != cf: continue
            catalogo_real[cf].add(p)
    catalogo = {}
    all_cats = set(list(catalogo_oficial.keys()) + list(catalogo_real.keys()))
    for cat in all_cats:
        base = set(catalogo_oficial.get(cat, [])); real = catalogo_real.get(cat, set())
        catalogo[cat] = sorted(base | real)
    print(f"[{ts()}] Catalogo final: {len(catalogo)} cats, {sum(len(v) for v in catalogo.values())} praticas")
    email_to_cf = {}; email_to_chave_tutor = {}
    col_email_t = next((c for c in df_t.columns if 'E-MAIL' in str(c).upper() or 'EMAIL' in str(c).upper()), None)
    if col_email_t:
        for _, t in df_at.iterrows():
            em = str(t.get(col_email_t, '') or '').strip().lower()
            chave_t = t['_CHAVE']
            cat_raw_ = str(t.get(col_cat, '') or '').strip() if col_cat else ''
            cf_ = CAT_MAP.get(cat_raw_, cat_raw_)
            if em and em != 'nan':
                email_to_cf[em] = cf_; email_to_chave_tutor[em] = chave_t
    col_email_p = next((c for c in df_p.columns if c.upper() in ('EMAIL', 'E-MAIL')), None)
    # ── Mapeamentos de fallback adicionais ──────────────────────────────────
    # Fallback 3: por nome do tutor (coluna "Nome do tutor" no Forms)
    col_nome_tutor_p = None
    for c in df_p.columns:
        cu = str(c).upper()
        if 'NOME' in cu and 'TUTOR' in cu: col_nome_tutor_p = c; break
    if not col_nome_tutor_p:
        for c in df_p.columns:
            cu = str(c).upper()
            if 'TUTOR' in cu and 'PONTOS' not in cu and 'COMENT' not in cu:
                col_nome_tutor_p = c; break

    def _norm_nome_match(s):
        import unicodedata
        s = str(s or '').strip().lower()
        s = unicodedata.normalize('NFD', s)
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
        return s

    # Mapear: nome_normalizado → chave CONTROLE
    nome_to_chave_tutor = {}
    for _, t in df_at.iterrows():
        nome_t = str(t.get(col_nome, '') or '').strip()
        chave_t = t['_CHAVE']
        if nome_t and chave_t:
            nome_to_chave_tutor[_norm_nome_match(nome_t)] = chave_t
            # Também mapear primeiro+último nome
            parts = _norm_nome_match(nome_t).split()
            if len(parts) >= 2:
                fl = parts[0] + ' ' + parts[-1]
                nome_to_chave_tutor.setdefault(fl, chave_t)

    sem_match = 0; com_match = 0; match_por_email = 0; match_por_nome = 0

    # ── Constantes de tratamento de portfólios ───────────────────────────────
    # Emails a ignorar completamente (gestão — não são tutores de prática)
    EMAILS_IGNORAR = {'elienai.cesar@uniasselvi.com.br'}

    # Domínio válido para tutores de prática
    DOMINIO_TUTORIA = 'tutorpratica.uniasselvi.com.br'

    # Avisos de portfólio (email incorreto, preenchimento incorreto, etc.)
    _avisos_raw = {}  # key → {email, nome, polo, tipo, msg, count}

    def _dividir_codigo_composto(chave):
        """Divide chave com código composto (BFR-BBI) em sub-chaves individuais."""
        import re as _re2
        m = _re2.search(r'([A-Z]{2,4}-[A-Z]{2,4})$', chave)
        if m:
            polo_base = chave[:-len(m.group(1))]
            partes = m.group(1).split('-')
            candidatos = [polo_base + p for p in partes]
            encontrados = [c for c in candidatos if c in chave_to_cf]
            if encontrados:
                return encontrados
        return []

    def _encontrar_por_polo(chave, cat_subm=''):
        """Opção B: vincula submissão ao tutor ativo no mesmo polo.
        Remove sufixos progressivos da chave para encontrar o polo base."""
        cat_up = str(cat_subm).upper()
        # Palavras-chave por categoria para priorizar match
        _cat_kw = {
            'BIO': ['BFR','BBI','BIO'],
            'FAR': ['BFR','BBI','BIO'],
            'ENF': ['ENF','ISN','INS'],
            'NTR': ['NTR','NUTRI'],
            'ENG': ['EMF','ENG','MEC'],
            'FIS': ['BFI','FIS','FISIO'],
        }
        # Determinar categoria da submissão para priorizar candidatos
        _pref_codigos = []
        for kw, codigos in _cat_kw.items():
            if kw in cat_up:
                _pref_codigos = codigos; break

        for code_len in range(3, min(len(chave), 12)):
            polo_prefix = chave[:-code_len]
            if len(polo_prefix) < 6: break
            candidatos = [k for k in chave_to_cf if k.startswith(polo_prefix)]
            if candidatos:
                # Priorizar por categoria
                if _pref_codigos:
                    pref = [c for c in candidatos if any(cod in c for cod in _pref_codigos)]
                    if pref: return pref[0]
                return candidatos[0]
        return None

    enviados = defaultdict(list)
    polo_sem_tutor = defaultdict(list)  # polo+cat → práticas de tutores desligados
    _correcoes_perfil = 0
    _correcoes_perfil_falhou = set()
    for _, r in df_p.iterrows():
        chave = r['_CHAVE']; proto = r['_PROTO']
        if not chave or chave == 'nan' or not proto or proto == 'nan': continue

        # Ignorar emails de gestão
        _email_subm = str(r.get(col_email_p, '') or '').strip().lower() if col_email_p else ''
        if _email_subm in EMAILS_IGNORAR:
            continue

        chave = chave_alias.get(chave, chave)

        # PATCH 15: a planilha antiga (Forms) tem uma "CHAVE LINK POLO" pré-calculada
        # que não distingue Fisio (BFI) / T.O. (BTO) / Estética (COS-TIP) corretamente
        # quando os 3 cursos compartilham o mesmo polo/laboratório — corrige usando o
        # nome da prática (inequívoco), que é mais confiável que a chave pronta.
        if NOME_TO_PERFIL and _norm_proto(proto) in NOME_TO_PERFIL:
            perfil_certo = NOME_TO_PERFIL[_norm_proto(proto)]
            for _cod_errado in ('BFI', 'BTO', 'COS-TIP', 'TIP-COS'):
                if chave.endswith(_cod_errado) and _cod_errado != perfil_certo:
                    chave_corrigida = chave[:-len(_cod_errado)] + perfil_certo
                    # Aplica sempre, mesmo sem tutor ativo no destino — é melhor cair
                    # no fluxo de "sem match" (vira anônimo no polo) do que ficar
                    # silenciosamente contado pro curso errado.
                    if chave_corrigida in chave_to_cf:
                        _correcoes_perfil += 1
                    else:
                        _correcoes_perfil_falhou.add((chave, chave_corrigida, proto))
                    chave = chave_corrigida
                    break

        if chave not in chave_to_cf:
            # Fallback 1: por email
            if col_email_p:
                em_p = str(r.get(col_email_p, '') or '').strip().lower()
                if em_p in email_to_chave_tutor:
                    chave = email_to_chave_tutor[em_p]
                    match_por_email += 1
                    # Sinalizar envio por e-mail de Regente de Polo (informativo)
                    if '@regentedepolo.' in em_p:
                        _nome_rp = str(r.get(col_nome_tutor_p, '') or '') if col_nome_tutor_p else ''
                        _aviso_key_rp = f"{em_p}||regente_de_polo"
                        if _aviso_key_rp not in _avisos_raw:
                            _avisos_raw[_aviso_key_rp] = {
                                'email': em_p, 'nome': _nome_rp, 'chave': chave,
                                'tipo': 'regente_de_polo',
                                'msg': 'Envio realizado pelo e-mail de Regente de Polo — tutor também é regente',
                                'count': 0
                            }
                        _avisos_raw[_aviso_key_rp]['count'] += 1

        if chave not in chave_to_cf:
            # Fallback 2: por nome do tutor na coluna específica
            if col_nome_tutor_p:
                nome_p = _norm_nome_match(str(r.get(col_nome_tutor_p, '') or ''))
                if nome_p in nome_to_chave_tutor:
                    chave = nome_to_chave_tutor[nome_p]
                    match_por_nome += 1
                else:
                    # Tentar primeiro+último nome
                    parts_p = nome_p.split()
                    if len(parts_p) >= 2:
                        fl_p = parts_p[0] + ' ' + parts_p[-1]
                        if fl_p in nome_to_chave_tutor:
                            chave = nome_to_chave_tutor[fl_p]
                            match_por_nome += 1

        if chave not in chave_to_cf:
            # Fallback 3: normalizar a própria chave (diferenças de espaços/acentos)
            chave_norm = _norm_nome_match(chave.replace(' ', ''))
            for k in chave_to_cf:
                if _norm_nome_match(k.replace(' ', '')) == chave_norm:
                    chave = k; match_por_nome += 1; break

        # Fallback 4: código composto (ex: BFR-BBI → BFR e BBI)
        if chave not in chave_to_cf:
            _sub_chaves = _dividir_codigo_composto(chave)
            if _sub_chaves:
                chave = _sub_chaves[0]  # usar primeira sub-chave
                match_por_nome += 1
                # Se houver mais de uma sub-chave, adicionar às demais depois
                for _sc in _sub_chaves[1:]:
                    if _sc in chave_to_cf:
                        _extra_proto = str(r.get('_PROTO', '') or '').strip()
                        _extra_data  = r['_DATA']
                        _extra_aluno = int(r['_ALUNOS'])
                        _extra_ordem = str(r.get('_ORDEM', 'Ordem 1') or 'Ordem 1').strip()
                        for _p in _extra_proto.split(';'):
                            _p = _p.strip()
                            if _p: enviados[_sc].append({'p':_p,'d':str(_extra_data)[:10] if pd.notna(_extra_data) else None,'a':_extra_aluno,'o':_extra_ordem})

        if chave in chave_to_cf:
            com_match += 1
        else:
            sem_match += 1
            # Criar aviso de portfólio
            _nome_subm  = str(r.get(col_nome_tutor_p, '') or '-') if col_nome_tutor_p else '-'
            _polo_subm  = chave  # chave contém polo+código
            if '@regentedepolo.' in _email_subm:
                _tipo = 'regente_de_polo'
                _msg  = 'Envio por e-mail de Regente de Polo sem correspondência no CONTROLE_TUTORIA'
            elif _email_subm and not _email_subm.endswith('@' + DOMINIO_TUTORIA):
                _dom = _email_subm.split('@')[1] if '@' in _email_subm else 'desconhecido'
                _tipo = 'email_incorreto'
                _msg  = f'E-mail incorreto: @{_dom} ao invés de @{DOMINIO_TUTORIA}'
            else:
                _tipo = 'chave_invalida'
                _msg  = f'Polo/categoria não encontrado no CONTROLE: {chave}'
            _aviso_key = f"{_email_subm}||{_tipo}"
            if _aviso_key not in _avisos_raw:
                _avisos_raw[_aviso_key] = {'email':_email_subm,'nome':_nome_subm,'chave':chave,'tipo':_tipo,'msg':_msg,'count':0}
            _avisos_raw[_aviso_key]['count'] += 1
            # Guardar práticas em polo_sem_tutor para entry anônimo no polo
            _polo_part = chave  # chave = polo+código sem separador
            for _pfb in str(r.get('_PROTO','') or '').split(';'):
                _pfb = _pfb.strip()
                if _pfb:
                    _ordem_pfb = str(r.get('_ORDEM','Ordem 1') or 'Ordem 1').strip()
                    polo_sem_tutor[_polo_part].append({'p':_pfb[:80],'d':None,'a':0,'o':_ordem_pfb})

        data = r['_DATA']; aluno = int(r['_ALUNOS'])
        for p in proto.split(';'):
            p = p.strip()
            if p:
                ordem_val = str(r.get('_ORDEM', 'Ordem 1') or 'Ordem 1').strip()
                if not any(o in ordem_val for o in ['Ordem 1','Ordem 2','Ordem 3','Ordem 4','Ordem 5']):
                    ordem_val = 'Ordem 1'
                _data_str = str(data)[:10] if pd.notna(data) else None
                _sem_envio = _data_para_semestre(_data_str)
                if _sem_envio is None:
                    # Data fora de qualquer janela → semestre mais antigo
                    _sem_envio = sorted(ALL_SEMESTRES.keys())[0]
                    # Log apenas se data estranha (debug)
                    if _data_str and _data_str < '2026-02-01':
                        pass  # silencioso — serão agrupados no 2026/1
                enviados[chave].append({'p': p, 'd': _data_str, 'a': aluno, 'o': ordem_val, 's': _sem_envio})
    # Finalizar avisos
    avisos_portfolio = sorted(_avisos_raw.values(), key=lambda x: -x['count'])
    print(f"[{ts()}] Matching submissões: {com_match} com chave, {match_por_email} por email, {match_por_nome} por nome/código, {sem_match} sem match")
    if _correcoes_perfil:
        print(f"[{ts()}] Correções BFI/BTO/COS-TIP por nome de prática: {_correcoes_perfil}")
    if _correcoes_perfil_falhou:
        print(f"[{ts()}] AVISO: {len(_correcoes_perfil_falhou)} correções de perfil sem tutor ativo no destino (foram para anônimo/polo):")
        for _ch, _chc, _pr in list(_correcoes_perfil_falhou)[:10]:
            print(f"    {_ch!r} -> {_chc!r} (não encontrado) | prática: {_pr!r}")
    if sem_match > 0:
        print(f"[{ts()}] Avisos de portfólio gerados: {len(avisos_portfolio)}")
        for av in avisos_portfolio:
            print(f"[{ts()}]   {av['tipo'].upper()} ({av['count']}x): {av['email']} | {av['nome']} | {av['msg']}")
    tutores = []
    # Pré-computar: para cada polo+CURSO ESPECÍFICO, agregar TODOS os envios
    # (resolve o caso de múltiplos tutores por polo que compartilham o mesmo curso)
    # PATCH 17: agrupar por curso específico (col_cur), não pela categoria ampla —
    # BFI/BTO/COS-TIP têm a mesma categoria ampla (Multidisciplinar III) mas são
    # cursos diferentes; agrupar pela categoria ampla juntava os 3 indevidamente.
    _polo_cat_enviados = {}  # (polo_str, curso_especifico) → lista merged de hist
    for _, t in df_at.iterrows():
        _ch = t['_CHAVE']
        _cursos_t = str(t.get(col_cur, '') or '').strip()
        _polo_str = str(t.get(col_polo, '') or '').strip()
        _key_pc = (_polo_str, _cursos_t)
        if _key_pc not in _polo_cat_enviados:
            # Buscar enviados pela chave canônica E por variantes de chave do mesmo
            # polo que tenham o MESMO curso específico (não só a mesma categoria ampla)
            _hist_merged = list(enviados.get(_ch, []))
            for _k, _h in enviados.items():
                if _k == _ch: continue
                if _k.startswith(_polo_str) and _k[len(_polo_str):] == _cursos_t:
                    for _item in _h:
                        if _item not in _hist_merged:
                            _hist_merged.append(_item)
            _polo_cat_enviados[_key_pc] = _hist_merged
    print(f"[{ts()}] Polo×curso com envios: {len([v for v in _polo_cat_enviados.values() if v])}")

    _hist_pre_admissao = 0
    for _, t in df_at.iterrows():
        chave    = t['_CHAVE']
        cat_raw  = str(t.get(col_cat, '') or '').strip() if col_cat else ''
        cat_form = CAT_MAP.get(cat_raw, cat_raw)
        praticas = catalogo.get(cat_form, catalogo.get(cat_raw, []))
        polo_str = str(t.get(col_polo, '') or '').strip()
        cursos_t = str(t.get(col_cur, '') or '').strip()
        # Enriquecimento MEC / data de admissão (calculado antes do filtro de hist)
        _email_t = str(t.get(col_email, '') or '').strip().lower() if col_email else ''
        _mec = mec_cache.get(_email_t, {})
        _inicio_ctrl = t.get(col_inicio) if col_inicio else None
        _inicio_str = None
        if _inicio_ctrl and str(_inicio_ctrl) not in ('nan','NaT','None',''):
            try:
                _inicio_str = _inicio_ctrl.strftime('%Y-%m-%d') if hasattr(_inicio_ctrl,'strftime') else str(_inicio_ctrl)[:10]
            except: pass
        if not _inicio_str: _inicio_str = _mec.get('admissao')

        # PATCH 17: agregado por polo+CURSO ESPECÍFICO (não categoria ampla) — cobre
        # múltiplos tutores do mesmo curso no mesmo polo, sem juntar BFI/BTO/COS-TIP
        hist_bruto = _polo_cat_enviados.get((polo_str, cursos_t), enviados.get(chave, []))
        # PATCH 16: práticas enviadas ANTES da admissão do tutor atual não são dele —
        # provavelmente foram enviadas por quem ocupava essa vaga antes (desligado).
        # Vão pro bucket anônimo do polo em vez de ficarem com o tutor novo.
        if _inicio_str:
            hist = []
            for h in hist_bruto:
                if h.get('d') and h['d'] < _inicio_str:
                    polo_sem_tutor[chave].append(h)
                    _hist_pre_admissao += 1
                else:
                    hist.append(h)
        else:
            hist = hist_bruto
        reais    = set(h['p'] for h in hist)
        pend     = [p for p in praticas if p not in reais]
        te = len(reais); tp = len(praticas)

        tutores.append({
            'n': str(t.get(col_nome, '') or ''),
            'p': str(t.get(col_polo, '') or ''),
            'c': cat_raw, 'cf': cat_form or 'Sem mapeamento',
            'tp': tp, 'te': te,
            'pend': pend, 'real': sorted(reais), 'hist': hist,
            'pct': round(te / tp * 100, 1) if tp else 0,
            'ch_semanal': _parse_ch(t.get(col_ch)) if col_ch else None,
            # Campos MEC / CONTROLE
            'inicio': _inicio_str,
            'lattes_url': _mec.get('lattes_url'),
            'lattes_id': _mec.get('lattes_id'),
            'titulacao': _mec.get('titulacao'),
            'graduacao': _mec.get('graduacao'),
            'especializacao': _mec.get('especializacao'),
            'mestrado': _mec.get('mestrado'),
            'doutorado': _mec.get('doutorado'),
            'exp_fora_meses': _mec.get('exp_fora_meses'),
            'exp_tutor_uni_meses': _mec.get('exp_tutor_uni_meses'),
            'whatsapp': str(t.get(col_whats,'') or '') if col_whats else None,
            'chapa': str(t.get(col_chapa,'') or '') if col_chapa else None,
        })
    if _hist_pre_admissao:
        print(f"[{ts()}] Práticas pré-admissão removidas do tutor atual (vão pro polo): {_hist_pre_admissao}")
    seen = {}; tutores_dedup = []
    for t in tutores:
        key = (t.get('p',''), t.get('n','').strip().lower())
        if key in seen:
            existing = seen[key]
            existing['te'] = max(existing['te'], t['te'])
            existing['hist'] = existing['hist'] + [h for h in t['hist'] if h not in existing['hist']]
            existing['real'] = sorted(set(existing['real']) | set(t['real']))
            existing['pend'] = [p for p in existing['pend'] if p not in existing['real']]
            existing['tp'] = max(existing['tp'], t['tp'])
            existing['pct'] = round(existing['te'] / existing['tp'] * 100, 1) if existing['tp'] else 0
            if t.get('ch_semanal') and not existing.get('ch_semanal'):
                existing['ch_semanal'] = t['ch_semanal']
        else:
            seen[key] = t; tutores_dedup.append(t)
    tutores = tutores_dedup

    # Criar entries anônimos para práticas de tutores desligados (vinculado ao polo)
    for chave_polo, hist_polo in polo_sem_tutor.items():
        if not hist_polo: continue
        _reais_polo = set(h['p'] for h in hist_polo)
        _cf_polo = ''
        for _p_polo in _reais_polo:
            _cf_polo = oficial_p_to_cat.get(_p_polo, '')
            if _cf_polo: break
        # Normalizar nome da categoria para o padrão do dashboard (usando CAT_MAP invertido)
        _CAT_MAP_INV = {v: k for k, v in CAT_MAP.items()}
        # Também mapear nomes parciais
        _cf_polo_norm = _cf_polo
        for _full, _short in _CAT_MAP_INV.items():
            if _full.lower() in _cf_polo.lower() or _cf_polo.lower() in _full.lower():
                _cf_polo_norm = _short; break
        # Verificar se já está no formato correto (chave do CAT_MAP)
        if _cf_polo_norm not in CAT_MAP:
            for _short in CAT_MAP:
                if _cf_polo.lower() in CAT_MAP[_short].lower():
                    _cf_polo_norm = _short; break
        if _cf_polo_norm in CAT_MAP:
            _cf_polo = _cf_polo_norm
        _praticas_polo = catalogo.get(_cf_polo, [])
        _pend_polo = [p for p in _praticas_polo if p not in _reais_polo]
        # Extrair nome do polo (remover código de curso do fim da chave)
        import re as _re_anon
        _polo_limpo = _re_anon.sub(r'[A-Z]{2,6}(-[A-Z]{2,6})?$', '', chave_polo).strip()
        if not _polo_limpo: _polo_limpo = chave_polo
        tutores.append({
            'n': 'Tutor desligado', 'p': _polo_limpo, 'c': _cf_polo, 'cf': _cf_polo or 'Sem categoria',
            'tp': len(_praticas_polo), 'te': len(_reais_polo),
            'pend': _pend_polo, 'real': sorted(_reais_polo), 'hist': hist_polo,
            'pct': round(len(_reais_polo)/len(_praticas_polo)*100,1) if _praticas_polo else 0,
            'ch_semanal': None, '_anonimo': True,
        })
    if polo_sem_tutor:
        print(f"[{ts()}] Entries anônimos criados: {len(polo_sem_tutor)} polos com práticas de tutores desligados")

    p_to_cat = {}
    for cat, pracs in catalogo.items():
        for p in pracs:
            if p not in p_to_cat: p_to_cat[p] = cat
    ps = defaultdict(lambda: {'enviou': 0, 'nao_enviou': 0, 'categoria': ''})
    for t in tutores:
        for p in t['real']:  ps[p]['enviou']    += 1
        for p in t['pend']:  ps[p]['nao_enviou'] += 1
    _p_fallback = {}
    for t in tutores:
        for p in t['real'] + t['pend']: _p_fallback.setdefault(p, t['cf'])
    for p in ps: ps[p]['categoria'] = p_to_cat.get(p, _p_fallback.get(p, ''))
    ps_all = sorted([{'nome': k, **v} for k, v in ps.items()], key=lambda x: -x['nao_enviou'])
    ps_list = ps_all[:30]
    cs = defaultdict(lambda: {'total_tutores': 0, 'com_100pct': 0, 'total_previstas': 0, 'total_enviadas': 0})
    for t in tutores:
        if not t['tp']: continue
        c = t['cf']
        cs[c]['total_tutores'] += 1
        if t['pct'] == 100: cs[c]['com_100pct'] += 1
        cs[c]['total_previstas'] += t['tp']; cs[c]['total_enviadas'] += t['te']
    print(f"[{ts()}] {len(tutores)} tutores, {sum(len(v) for v in catalogo.values())} praticas")
    prazos = PRAZOS_ORDENS.copy()
    hoje = datetime.now()
    status_ordem = {}
    # PATCH 3: usar datas de início reais de PERIODOS_ORDENS
    for ordem, prazo_str in prazos.items():
        prazo_date = datetime.strptime(prazo_str, '%d/%m/%Y')
        periodo = PERIODOS_ORDENS.get(ordem, {})
        inicio_str = periodo.get('inicio', '')
        if inicio_str:
            try: inicio_date = datetime.strptime(inicio_str, '%d/%m/%Y')
            except ValueError: inicio_date = prazo_date.replace(day=1)
        else: inicio_date = prazo_date.replace(day=1)
        if hoje > prazo_date: status_ordem[ordem] = 'VENCIDO'
        elif hoje >= inicio_date: status_ordem[ordem] = 'ABERTA'
        else: status_ordem[ordem] = 'FUTURA'
    tutores_out = []
    for t in tutores:
        por_ordem = {}
        for h in t['hist']:
            o = h.get('o', 'Ordem 1') or 'Ordem 1'
            por_ordem[o] = por_ordem.get(o, 0) + 1
        # PATCH 4: situação corrigida quando não há ordens vencidas
        ordens_vencidas = [o for o, s in status_ordem.items() if s == 'VENCIDO']
        ordens_abertas  = [o for o, s in status_ordem.items() if s == 'ABERTA']
        if not ordens_vencidas:
            if any(por_ordem.get(o, 0) > 0 for o in ordens_abertas) if ordens_abertas else False:
                sit = 'ok'
            elif ordens_abertas: sit = 'atrasado'
            else: sit = 'ok'
        else:
            if all(por_ordem.get(o, 0) > 0 for o in ordens_vencidas): sit = 'ok'
            elif any(por_ordem.get(o, 0) > 0 for o in ordens_vencidas): sit = 'atrasado'
            else: sit = 'urgente'
        tutores_out.append({
            **t,
            'nome': t.get('n',''), 'polo': t.get('p',''), 'cat': t.get('c',''),
            'n': t.get('n',''), 'p': t.get('p',''), 'c': t.get('c',''),
            'cf': t.get('cf','Sem mapeamento'),
            'por_ordem': por_ordem, 'porOrdem': por_ordem, 'situacao': sit,
        })
    col_email_key = next((c for c in df_t.columns if 'E-MAIL' in str(c).upper()), None)
    nome_to_email = {}
    if col_email_key:
        for _, row in df_at.iterrows():
            nome = str(row.get(col_nome, '') or '').strip()
            email = str(row.get(col_email_key, '') or '').strip().lower()
            if nome and email and email != 'nan': nome_to_email[nome] = email
    seen = {}; tutores_dedup = []
    for t in tutores_out:
        nome = t['n']; polo = t['p']
        email = nome_to_email.get(nome, '')
        key = email if email else (nome + '|' + polo).lower()
        if key not in seen:
            seen[key] = len(tutores_dedup); tutores_dedup.append(dict(t))
        else:
            ex = tutores_dedup[seen[key]]
            ex['hist'] = ex['hist'] + t['hist']
            merged_real = sorted(set(ex['real']) | set(t['real']))
            ex['real'] = merged_real; ex['te'] = len(merged_real)
            for o, cnt in t['por_ordem'].items():
                ex['por_ordem'][o] = ex['por_ordem'].get(o, 0) + cnt
                ex['porOrdem'][o]  = ex['porOrdem'].get(o, 0) + cnt
            real_set = set(merged_real)
            ex['pend'] = [p for p in ex['pend'] if p not in real_set]
            if ex['tp'] > 0: ex['pct'] = round(ex['te'] / ex['tp'] * 100, 1)
            # PATCH 5: reavalia situação + sincroniza sit
            _orv = [o for o, s in status_ordem.items() if s == 'VENCIDO']
            if not _orv: ex['situacao'] = 'ok'
            elif all(ex['por_ordem'].get(o, 0) > 0 for o in _orv): ex['situacao'] = 'ok'
            elif any(ex['por_ordem'].get(o, 0) > 0 for o in _orv): ex['situacao'] = 'atrasado'
            else: ex['situacao'] = 'urgente'
            ex['sit'] = ex['situacao']  # PATCH 5: sync shortcut
            if t.get('ch_semanal') and not ex.get('ch_semanal'): ex['ch_semanal'] = t['ch_semanal']
    tutores_out = tutores_dedup
    print(f"[{ts()}] Após deduplicação: {len(tutores_out)} tutores únicos")

    # PATCH 13: anexa o relatório de acompanhamento/comunicação na ficha do tutor
    def _norm_nome(s):
        s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode('ascii')
        return ' '.join(s.upper().split())
    com_file = os.path.join(SCRIPT_DIR, 'tutores_comunicacao.json')
    if os.path.isfile(com_file):
        with open(com_file, encoding='utf-8') as f: com_lista = json.load(f)
        # PATCH 14: nomes do relatório costumam vir incompletos (ex: "Gabriella Ribeiro"
        # vs "Gabriella Ribeiro Sousa" no cadastro) — casa por prefixo de tokens, não
        # só por igualdade exata
        com_tokens = [(tuple(_norm_nome(c['nome']).split()), c) for c in com_lista]
        _matches = 0
        for t in tutores_out:
            t_tokens = tuple(_norm_nome(t.get('n','')).split())
            achou = None
            for ctoks, c in com_tokens:
                n = min(len(ctoks), len(t_tokens))
                if n >= 2 and ctoks[:n] == t_tokens[:n]:
                    achou = c; break
            if achou:
                t['comunicacao'] = achou
                _matches += 1
        print(f"[{ts()}] Comunicação de tutores: {_matches}/{len(com_lista)} vinculados por nome")

    total     = len(tutores_out)
    enviaram  = sum(1 for t in tutores_out if t['te'] > 0)
    atrasados = sum(1 for t in tutores_out if t['situacao'] == 'atrasado')
    urgentes  = sum(1 for t in tutores_out if t['situacao'] == 'urgente')
    total_alunos = sum(h['a'] for t in tutores_out for h in t['hist'])
    polo_map = {}
    for t in tutores_out:
        p = t['polo']
        if p not in polo_map:
            polo_map[p] = {'POLO': p, 'polo': p, 'n': p, 'total': 0, 'enviaram': 0, 'atrasados': 0, 'alunos': 0}
        polo_map[p]['total'] += 1
        if t['te'] > 0: polo_map[p]['enviaram'] += 1
        if t['situacao'] == 'atrasado': polo_map[p]['atrasados'] += 1
        polo_map[p]['alunos'] += sum(h['a'] for h in t['hist'])
    polo_envios = {}
    for t in tutores_out:
        p = t['polo']
        polo_envios[p] = polo_envios.get(p, 0) + len(t.get('hist', []))
    polo_stats = sorted(polo_map.values(), key=lambda x: -x['atrasados'])
    for p in polo_stats:
        p['n'] = p.get('polo', p.get('POLO', ''))
        p['t'] = p['total']; p['e'] = p['enviaram']; p['a'] = p['alunos']
        p['pend'] = p['total'] - p['enviaram']
        p['pct'] = round(p['enviaram'] / p['total'] * 100) if p['total'] else 0
        p['envios'] = polo_envios.get(p['POLO'], 0)
    ordem_map = {o: {'envios': 0, 'alunos': 0} for o in prazos}
    for t in tutores_out:
        for h in t['hist']:
            o = h.get('o', 'Ordem 1') or 'Ordem 1'
            if o in ordem_map: ordem_map[o]['envios'] += 1; ordem_map[o]['alunos'] += h['a']
    por_ordem = [
        {'ordem': o, 'prazo': prazos[o], 'status': status_ordem[o],
         'envios': ordem_map[o]['envios'], 'alunos': ordem_map[o]['alunos']}
        for o in prazos
    ]
    mes_map = {}
    for t in tutores_out:
        for h in t['hist']:
            d = h.get('d') or ''
            mes = d[:7] if d and len(d) >= 7 else 'Sem data'
            if mes not in mes_map: mes_map[mes] = {'MES': mes, 'mes': mes, 'envios': 0, 'alunos': 0}
            mes_map[mes]['envios'] += 1; mes_map[mes]['alunos'] += h.get('a', 0)
    por_mes = sorted(mes_map.values(), key=lambda x: x['mes'])
    por_ordem_dict = {o: ordem_map[o]['envios'] for o in prazos}
    alunos_por_ordem = {o: ordem_map[o]['alunos'] for o in prazos}
    for p in polo_stats:
        p['n'] = p.get('polo', p.get('POLO', ''))
        p['t'] = p.get('total', 0); p['e'] = p.get('enviaram', 0); p['a'] = p.get('alunos', 0)
    for t in tutores_out:
        t['sit'] = t.get('situacao', 'urgente')
        t['al'] = sum(h.get('a', 0) for h in t.get('hist', []))
        t['email'] = nome_to_email.get(t.get('n', ''), '')
    praticas_template = []
    for p in ps_all:
        total_p = p['enviou'] + p['nao_enviou']
        praticas_template.append({
            'n': p['nome'], 'c': p['categoria'],
            'env_n': p['enviou'], 'pend_n': p['nao_enviou'],
            'pct': round(p['enviou'] / total_p * 100, 1) if total_p else 0,
            'nome': p['nome'], 'enviou': p['enviou'], 'nao_enviou': p['nao_enviou'], 'categoria': p['categoria'],
        })
    # ── Estatísticas por semestre ────────────────────────────────────────────
    def _stats_semestre(sem_key, tutores_list, catalogo_dict, prazos_dict, periodos_dict):
        """Gera o mesmo bloco de dados que processar() retorna, mas filtrado por semestre."""
        from datetime import datetime as _dt2
        _hoje = datetime.now()
        _status_ord = {}
        for _o, _pz in prazos_dict.items():
            try:
                _pz_d = _dt2.strptime(_pz, '%d/%m/%Y')
                _ini  = _dt2.strptime(periodos_dict.get(_o,{}).get('inicio',_pz), '%d/%m/%Y')
                if _hoje > _pz_d: _status_ord[_o] = 'VENCIDO'
                elif _hoje >= _ini: _status_ord[_o] = 'ABERTA'
                else: _status_ord[_o] = 'FUTURA'
            except: _status_ord[_o] = 'FUTURA'

        _por_ordem = {}; _alunos_por_ordem = {}
        _polo_map = {}; _cat_stats = {}

        for _t in tutores_list:
            _sem_antigo = sorted(ALL_SEMESTRES.keys())[0]
            _hist_sem = [h for h in _t.get('hist', []) if h.get('s', _sem_antigo) == sem_key]
            _reais_sem = set(h['p'] for h in _hist_sem)
            _te_sem = len(_reais_sem)
            _tp = _t.get('tp', 0)
            _pct_sem = round(_te_sem / _tp * 100, 1) if _tp else 0

            # por_ordem deste semestre
            _po = {}
            for _h in _hist_sem:
                _o = _h.get('o','Ordem 1') or 'Ordem 1'
                _po[_o] = _po.get(_o, 0) + 1
                _por_ordem[_o] = _por_ordem.get(_o, 0) + 1
                _alunos_por_ordem[_o] = _alunos_por_ordem.get(_o, 0) + _h.get('a', 0)

            # situação neste semestre
            _orv = [_o for _o, _s in _status_ord.items() if _s == 'VENCIDO']
            if not _orv: _sit = 'ok' if _te_sem > 0 else 'atrasado'
            elif all(_po.get(_o,0) > 0 for _o in _orv): _sit = 'ok'
            elif any(_po.get(_o,0) > 0 for _o in _orv): _sit = 'atrasado'
            else: _sit = 'urgente'

            # polo
            _p = _t.get('p','')
            if _p not in _polo_map:
                _polo_map[_p] = {'n':_p,'polo':_p,'POLO':_p,'total':0,'enviaram':0,'atrasados':0,'alunos':0,'pend':0,'pct':0,'envios':0,'t':0,'e':0,'a':0}
            _polo_map[_p]['total'] += 1; _polo_map[_p]['t'] += 1
            if _te_sem > 0: _polo_map[_p]['enviaram'] += 1; _polo_map[_p]['e'] += 1
            if _sit == 'atrasado': _polo_map[_p]['atrasados'] += 1
            _polo_map[_p]['alunos'] += sum(h.get('a',0) for h in _hist_sem)
            _polo_map[_p]['a'] = _polo_map[_p]['alunos']

            # cat_stats
            _cf = _t.get('cf','')
            if _cf not in _cat_stats:
                _cat_stats[_cf] = {'total_tutores':0,'com_100pct':0,'total_previstas':0,'total_enviadas':0}
            if _tp:
                _cat_stats[_cf]['total_tutores'] += 1
                if _pct_sem == 100: _cat_stats[_cf]['com_100pct'] += 1
                _cat_stats[_cf]['total_previstas'] += _tp
                _cat_stats[_cf]['total_enviadas'] += _te_sem

        # Calcular pct e pend nos polos
        for _ps in _polo_map.values():
            _ps['pend'] = _ps['total'] - _ps['enviaram']
            _ps['pct']  = round(_ps['enviaram']/_ps['total']*100) if _ps['total'] else 0

        _total = len(tutores_list)
        _sem_ant = sorted(ALL_SEMESTRES.keys())[0]
        _enviaram = sum(1 for _t in tutores_list if any(h.get('s',_sem_ant)==sem_key and h.get('p') for h in _t.get('hist',[])))
        # Calcular urgentes e atrasados corretamente
        _orv = [_o for _o, _s in _status_ord.items() if _s == 'VENCIDO']
        _urgentes = 0; _atrasados = 0
        for _t in tutores_list:
            _sem_ant2 = sorted(ALL_SEMESTRES.keys())[0]
            _h = [h for h in _t.get('hist', []) if h.get('s', _sem_ant2) == sem_key]
            _po = {}
            for _hh in _h:
                _o = _hh.get('o', 'Ordem 1') or 'Ordem 1'
                _po[_o] = _po.get(_o, 0) + 1
            if _orv:
                _venc_ok = [_o for _o in _orv if _po.get(_o, 0) > 0]
                if len(_venc_ok) == 0:
                    _urgentes += 1
                elif len(_venc_ok) < len(_orv):
                    _atrasados += 1

        return {
            'semestre': sem_key,
            'kpis': {
                'total': _total, 'enviaram': _enviaram, 'pendentes': _total - _enviaram,
                'urgentes': _urgentes, 'atrasados': _atrasados,
                'total_polos': len(_polo_map),
                'polos_ok': sum(1 for _ps in _polo_map.values() if _ps['pend']==0),
            },
            'polo_stats': sorted(_polo_map.values(), key=lambda x: -x.get('pend',0)),
            'por_ordem': _por_ordem,
            'alunos_por_ordem': _alunos_por_ordem,
            'status_ordem': _status_ord,
            'prazos': prazos_dict,
            'periodos': periodos_dict,
            'cat_stats': [{'categoria':k,**v} for k,v in _cat_stats.items()],
        }

    _dados_por_semestre = {}
    for _sem_k, _sem_cfg in ALL_SEMESTRES.items():
        _dados_por_semestre[_sem_k] = _stats_semestre(
            _sem_k, tutores_out,
            catalogo, _sem_cfg['prazos'], _sem_cfg['periodos']
        )
        _env_s = _dados_por_semestre[_sem_k]['kpis']['enviaram']
        print(f"[{ts()}] Semestre {_sem_k}: {_env_s} tutores com envios")
    # ── Fim estatísticas por semestre ─────────────────────────────────────────

    BRT = timezone(timedelta(hours=-3))
    gerado = datetime.now(BRT).strftime('%d/%m/%Y %H:%M')
    ch_ok = sum(1 for t in tutores_out if t.get('ch_semanal'))
    print(f"[{ts()}] {total} tutores · {enviaram} enviaram · {atrasados} atrasados · {urgentes} urgentes")
    print(f"[{ts()}] CH SEMANAL preenchida: {ch_ok}/{total} tutores")
    return limpar({
        'kpis': {
            'total': total, 'enviaram': enviaram, 'pendentes': total - enviaram,
            'atrasados': atrasados, 'urgentes': urgentes,
            'total_alunos': total_alunos, 'total_polos': len(polo_map),
            'polos_ok': sum(1 for p in polo_stats if p['enviaram'] > 0),
        },
        'tutores': tutores_out, 'polo_stats': polo_stats,
        'por_ordem': por_ordem_dict, 'por_ordem_lista': por_ordem,
        'alunos_por_ordem': alunos_por_ordem, 'status_ordem': status_ordem,
        'cat_stats': [{'categoria': k, **v} for k, v in cs.items()],
        'pratica_stats': ps_list, 'praticas': praticas_template,
        'catalogo': catalogo, 'prazos': prazos,
        'por_mes': por_mes, 'gerado_em': gerado,
        'avisos_portfolio': avisos_portfolio,
        'semestre': SEMESTRE_ATUAL,
        'todos_semestres': sorted(ALL_SEMESTRES.keys()),
        'dados_por_semestre': _dados_por_semestre,
        'disciplinas_por_ordem': _DISCIPLINAS_POR_ORDEM_GLOBAL,
        'laboratorios': _carregar_laboratorios(),
    })


def _carregar_laboratorios():
    # PATCH 13: dados pré-processados da seção "Laboratórios" (gerados fora do
    # pipeline, a partir das bases da Juliana + do dataset de práticas 2026/1)
    path = os.path.join(SCRIPT_DIR, 'laboratorios_data.json')
    if not os.path.isfile(path):
        print(f"[{ts()}] laboratorios_data.json não encontrado — seção Laboratórios fica vazia")
        return {}
    with open(path, encoding='utf-8') as f:
        lab = json.load(f)
    print(f"[{ts()}] Laboratórios: {len(lab.get('ricos',{}))} categorias ricas, {len(lab.get('simples',{}))} categorias simples")
    return lab


def _detectar_e_corrigir_base64(p4):
    import base64 as _b64
    try:
        with open(str(p4), 'rb') as f: raw = f.read(16)
        if raw[:4] == b'PK\x03\x04': return
        with open(str(p4), 'rb') as f: full = f.read()
        for padded in [full.strip(), full.strip() + b'==']:
            try:
                decoded = _b64.b64decode(padded)
                if decoded[:4] == b'PK\x03\x04':
                    with open(str(p4), 'wb') as fw: fw.write(decoded)
                    print(f"  [FIX] Base64 detectado e corrigido ({len(full)}->{len(decoded)} bytes)")
                    return
            except Exception: continue
        if raw[:5] in (b'\r\n<!D', b'<!DOC', b'<html'):
            print(f"  [ERRO] Arquivo é uma página HTML (login Microsoft)")
        else:
            print(f"  [INFO] Arquivo não é ZIP nem base64 ({raw[:4].hex()})")
    except Exception as e: print(f"  [AVISO] Verificação base64 falhou: {e}")


def _ler_lotacao_xlsx(p4):
    from openpyxl import load_workbook as _lwb
    wb = _lwb(str(p4), read_only=True, data_only=True, keep_vba=False)
    ws = wb['Quadro Geral de Lotação'] if 'Quadro Geral de Lotação' in wb.sheetnames else list(wb.worksheets)[0]
    return list(ws.iter_rows(values_only=True))

def _ler_lotacao_xls(p4):
    import xlrd
    wb = xlrd.open_workbook(str(p4))
    try: ws = wb.sheet_by_name('Quadro Geral de Lotação')
    except xlrd.XLRDError: ws = wb.sheet_by_index(0)
    rows = []
    for i in range(ws.nrows):
        row = []
        for j in range(ws.ncols):
            cell = ws.cell(i, j)
            if cell.ctype == xlrd.XL_CELL_DATE:
                import xlrd.xldate
                row.append(xlrd.xldate.xldate_as_datetime(cell.value, wb.datemode))
            else: row.append(cell.value if cell.ctype != xlrd.XL_CELL_EMPTY else None)
        rows.append(tuple(row))
    return rows

def _ler_lotacao_pandas(p4):
    p = str(p4)
    try: df = pd.read_excel(p, sheet_name='Quadro Geral de Lotação', header=None)
    except Exception: df = pd.read_excel(p, sheet_name=0, header=None)
    return [tuple(row) for row in df.fillna('').values.tolist()]

def carregar_lotacao(p4):
    fname = os.path.basename(str(p4))
    print(f"[{ts()}] Lendo lotação de tutores ({fname})...")
    _detectar_e_corrigir_base64(p4)
    _rows = None
    for estrategia, fn in [('openpyxl', _ler_lotacao_xlsx), ('xlrd', _ler_lotacao_xls), ('pandas', _ler_lotacao_pandas)]:
        try:
            _rows = fn(p4)
            print(f"[{ts()}] Lotação lida via {estrategia}: {len(_rows)} linhas")
            break
        except Exception as e: print(f"[{ts()}] Tentativa {estrategia}: {e}")
    if not _rows: raise RuntimeError(f"Não foi possível ler {fname}")
    lotacao = {}
    # Diagnóstico: mostrar cabeçalhos (linha 0) e linha 2 para verificar estrutura
    if _rows:
        _hdrs = [str(c or '').strip()[:20] for c in _rows[0]] if _rows[0] else []
        print(f"[{ts()}] Lotação colunas (linha 1): {_hdrs[:35]}")
        if len(_rows) > 1:
            _r2 = [str(c or '')[:15] for c in _rows[1]]
            print(f"[{ts()}] Lotação linha 2: {_r2[:35]}")
        # Detectar coluna de total_alunos automaticamente
        _col_alunos = 26  # default
        for _ci, _h in enumerate(_hdrs):
            _hu = _h.upper()
            if any(k in _hu for k in ['TOTAL', 'ALUNOS', 'MATR']):
                _col_alunos = _ci
                print(f"[{ts()}] Coluna alunos detectada: {_ci} = '{_h}'")
                break
    else:
        _col_alunos = 26

    for r in _rows[2:]:
        if not r[8] or str(r[8]).strip() in ('', '-', 'None', 'nan'): continue
        nome_raw = str(r[8]).strip(); nome_lower = nome_raw.lower()
        try: total_al = int(float(str(r[27] if len(r) > 27 else 0) or 0))
        except: total_al = 0
        # Colunas detectadas da planilha Lotação de Tutores_2026_2:
        # [5]=CURSOS, [7]=CONTRATAÇÃO, [8]=TUTOR, [14]=PERFIL, [15]=CH SEMANAL
        # [16]=CH IDEAL, [27]=TOTAL ALUNOS, [30]=CATEGORIA GIOCONDA
        lotacao[nome_lower] = {
            'nome_oficial': nome_raw,
            'perfil':       str(r[14] if len(r) > 14 else '') or '',
            'cursos':       str(r[5]  if len(r) > 5  else '') or '',
            'ch_semanal':   _parse_ch(r[15] if len(r) > 15 else None),
            # CH IDEAL: usar r[16] se preenchido, senão CH PROPOSTA (r[13])
            '_ch_ideal_raw': _parse_ch(r[16] if len(r) > 16 else None) or 0.0,
            '_ch_prop_raw':  _parse_ch(r[13] if len(r) > 13 else None) or 0.0,
            'ch_ideal': (_parse_ch(r[16] if len(r) > 16 else None) or
                         _parse_ch(r[13] if len(r) > 13 else None) or 0.0),
            'contratacao':  str(r[7]  if len(r) > 7  else '') or '',
            'polo_hub':     str(r[4]  if len(r) > 4  else '') or '',
            'categoria_gio':str(r[30] if len(r) > 30 else '') or '',
            'total_alunos': total_al,
        }
        # Indexar também por nome primeiro+último para match mais abrangente
        _parts = nome_lower.split()
        if len(_parts) >= 2:
            _nfl = _parts[0] + ' ' + _parts[-1]
            if _nfl not in lotacao:
                lotacao[_nfl] = lotacao[nome_lower]
    print(f"[{ts()}] Lotação: {len(lotacao)} tutores mapeados")
    return lotacao


CURSOS_NOMES = {
    'EMF-ISN': 'Enfermagem e Instrumentação Cirúrgica', 'EMF-ISN2': 'Enfermagem e Instrumentação Cirúrgica',
    'BFR': 'Farmácia', 'BBI': 'Biomedicina', 'BFI': 'Fisioterapia', 'BTO': 'T. Ocupacional',
    'COS-TIP': 'Estética e Cosmética', 'NTR': 'Nutrição', 'AGM': 'Agronomia',
    'BAU': 'Arquitetura e Urbanismo', 'ECE-ENM-ENS-ENG-EEA-GPI-CDE-OBR-SAN-TER-FSA-SLF-QUI': 'Engenharias e Licenciaturas',
    'BIOMEDICINA': 'Biomedicina', 'FARMÁCIA': 'Farmácia', 'FISIOTERAPIA': 'Fisioterapia',
    'TERAPIA OCUPACIONAL': 'T. Ocupacional', 'NUTRIÇÃO': 'Nutrição', 'AGRONOMIA': 'Agronomia',
    'ARQUITETURA E URBANISMO': 'Arquitetura e Urbanismo',
}

def enriquecer_tutores(dados, lotacao):
    tutores = dados.get('tutores', [])
    matched = 0
    LAB_PARA_CAT = {
        'ENFERMAGEM,INSTRUMENTAÇÃO CIRÚRGICA': 'Enfermagem e Instrumentação Cirúrgica',
        'ENFERMAGEM,INSTRUMENTAÇÃO CIRÚRGICA2': 'Enfermagem e Instrumentação Cirúrgica',
        'BIOMEDICINA': 'Biomedicina', 'FARMÁCIA': 'Farmácia', 'FISIOTERAPIA': 'Fisioterapia',
        'TERAPIA OCUPACIONAL': 'T. Ocupacional',
        'TECNOLOGIA EM ESTÉTICA E COSMÉTICA,ESTÉTICA E IMAGEM PESSOAL': 'Estética e Cosmética',
        'NUTRIÇÃO': 'Nutrição', 'AGRONOMIA': 'Agronomia', 'ARQUITETURA E URBANISMO': 'Arquitetura e Urbanismo',
        'CONSTRUÇÃO DE EDIFÍCIOS,ENGENHARIA CIVIL,ENGENHARIA ELÉTRICA,ENGENHARIA DE PRODUÇÃO,ENGENHARIA MECÂNICA,ENGENHARIA AMBIENTAL E SANITÁRIA,FORMAÇÃO PEDAGÓGICA EM FÍSICA,FÍSICA,GESTÃO DA PRODUÇÃO INDUSTRIAL,CONTROLE DE OBRAS,QUÍMICA,SANEAMENTO AMBIENTAL,SEGUNDA LICENCIATURA EM FÍSICA,TECNOLOGIA EM ENERGIAS RENOVÁVEIS': 'Engenharias e Licenciaturas',
        'ENGENHARIA CIVIL,ENGENHARIA ELÉTRICA,ENGENHARIA DE PRODUÇÃO,ENGENHARIA MECÂNICA': 'Engenharias (Civil/Elét./Prod./Mec.)',
        'EMF-ISN': 'Enfermagem e Instrumentação Cirúrgica', 'BFR': 'Farmácia',
        'BBI': 'Biomedicina', 'BFI': 'Fisioterapia', 'BTO': 'T. Ocupacional',
        'COS-TIP': 'Estética e Cosmética', 'NTR': 'Nutrição', 'AGM': 'Agronomia',
        'BAU': 'Arquitetura e Urbanismo',
    }
    polo_lab_seen = set(); alunos_por_lab_raw = {}
    for nome_lower, info in lotacao.items():
        cursos_raw = info.get('cursos', '').strip().upper()
        polo_hub   = info.get('polo_hub', '').strip()
        total_al   = info.get('total_alunos', 0)
        if not cursos_raw or not total_al: continue
        chave_polo_lab = f"{polo_hub}||{cursos_raw}"
        if chave_polo_lab in polo_lab_seen: continue
        polo_lab_seen.add(chave_polo_lab)
        sep = '+' if '+' in cursos_raw else ','
        componentes = sorted([c.strip() for c in cursos_raw.split(sep)])
        lab_key = (','.join(componentes) if any(len(c) > 8 for c in componentes) else '+'.join(componentes))
        alunos_por_lab_raw[lab_key] = alunos_por_lab_raw.get(lab_key, 0) + total_al
    def _norm_lab_key(k):
        sep = '+' if '+' in k else ','
        partes = sorted([p.strip().upper() for p in k.split(sep)])
        return (','.join(partes) if any(len(p) > 8 for p in partes) else '+'.join(partes))
    lab_cat_norm = {_norm_lab_key(k): v for k, v in LAB_PARA_CAT.items()}
    alunos_por_curso = []
    if alunos_por_lab_raw:
        # Fonte 1: TOTAL ALUNOS da lotação (quando preenchido)
        for lab_key, total in sorted(alunos_por_lab_raw.items(), key=lambda x: -x[1]):
            nome = lab_cat_norm.get(lab_key)
            if not nome:
                primeiro = lab_key.split(',')[0].split('+')[0].strip()
                nome = CURSOS_NOMES.get(primeiro, primeiro.title())
            alunos_por_curso.append({'sigla': lab_key, 'curso': nome, 'alunos': total})
    else:
        # Fonte 2 (fallback): agrupar matrículas do hub CSV por categoria
        # Usa alunos_hub_por_grupo gerado no processar_alunos_hub()
        # Fallback: usar por_cat do hub CSV (matrículas únicas por categoria)
        # Disponível em dados['hub']['por_cat'] após processar_alunos_hub()
        _hub_por_cat = (dados.get('hub') or {}).get('por_cat', {})
        _CAT_NOME = {
            'ENF-INS (Multidisciplinar II)':        'Enfermagem e Instrumentação Cirúrgica',
            'BIO-FAR (Multidisciplinar I)':         'Biomedicina e Farmácia',
            'BIO-FISIO-EST-TO (Multidisciplinar III)': 'Fisioterapia, T.Ocupacional e Estética',
            'NUTRI (Multidisciplinar IV)':          'Nutrição',
            'ENGMAKER':                             'Engenharias e Licenciaturas',
            'QUÍMICA E FÍSICA':                     'Química e Física',
        }
        if _hub_por_cat:
            for sigla, total in sorted(_hub_por_cat.items(), key=lambda x: -x[1]):
                if total > 0:
                    alunos_por_curso.append({
                        'sigla': sigla,
                        'curso': _CAT_NOME.get(sigla, sigla),
                        'alunos': int(total)
                    })
            print(f"[{ts()}] Alunos por curso (hub CSV): {len(alunos_por_curso)} categorias, total {sum(x['alunos'] for x in alunos_por_curso)}")
        else:
            print(f"[{ts()}] Alunos por curso: sem dados disponíveis")
    dados['alunos_por_curso'] = alunos_por_curso
    total_al_sum = sum(x['alunos'] for x in alunos_por_curso)
    print(f"[{ts()}] Alunos por lab: {len(alunos_por_curso)} labs, total {total_al_sum:,}")

    for t in tutores:
        nome_lower = str(t.get('n', '')).lower()
        info = lotacao.get(nome_lower)
        if not info:
            for k, v in lotacao.items():
                if nome_lower in k or k in nome_lower: info = v; break
        if info:
            t['perfil'] = info['perfil']; t['cursos'] = info['cursos']
            # Lotação tem prioridade sobre CH da planilha de controle
            if info.get('ch_semanal') and info['ch_semanal'] > 0:
                t['ch_semanal'] = info['ch_semanal']
            t['ch_ideal'] = info.get('ch_ideal', 0)
            t['contratacao_lot'] = info['contratacao']
            t['lab'] = info.get('cursos', '')  # curso da planilha de lotação (para Multi 3)
            t['polo_hub_lot'] = info.get('polo_hub', '')
            matched += 1
    print(f"[{ts()}] Enriquecimento: {matched}/{len(tutores)} tutores com perfil/CH")
    # ── Adicionar tutores sintéticos para avisos (aparecem na aba Tutores) ────
    _avisos_enr = dados.get('avisos_portfolio', [])
    if _avisos_enr:
        for av in _avisos_enr:
            if av['nome'] and av['nome'] not in ('nan', '-', ''):
                nome_display = av['nome']
            else:
                # Extrair nome do email
                local = av['email'].split('@')[0] if '@' in av['email'] else av['email']
                nome_display = local.replace('.', ' ').replace('_', ' ').title()
            tutores.append({
                'n': nome_display,
                'p': av.get('chave', '').replace('BFR-BBI','').replace('EMF-ISN','').replace('NTR','').strip(),
                'c': 'Aviso de Portfólio', 'cf': 'Aviso de Portfólio',
                'tp': 0, 'te': av['count'],
                'pend': [], 'real': [], 'hist': [],
                'pct': 0,
                'ch_semanal': None,
                'aviso_tipo': av['tipo'],
                'aviso_msg': av['msg'],
                'aviso_email': av['email'],
                'aviso_count': av['count'],
            })
    dados['tutores'] = tutores

    # avisos_portfolio já está em dados (vindo de processar()) — não sobrescrever com []
    if 'avisos_portfolio' not in dados:
        dados['avisos_portfolio'] = []
    return dados


def processar_gerenciamento_csv(p5):
    """Processa CSV detalhado de gerenciamento (REL_DETALHADO.csv)."""
    import csv, re as _re
    print(f"[{ts()}] Lendo gerenciamento (CSV)...")
    for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
        try:
            with open(str(p5), 'r', encoding=enc, errors='replace') as f:
                rows = list(csv.reader(f, delimiter=';'))
            break
        except: continue
    header = rows[0]; data = rows[1:]
    col = {h.strip().upper(): i for i, h in enumerate(header)}
    def gc(name): return col.get(name.upper())
    ci_polo = gc('LABORATORIO'); ci_cat = gc('CATEGORIA'); ci_exp = gc('NOME_EXPERIMENTO')
    ci_tutor = gc('TUTOR'); ci_mat = gc('ALUNOS_MATRICULADOS'); ci_agend = gc('ALUNOS_AGENDADOS')
    ci_capa = gc('CAPACIDADE_TOTAL'); ci_ofe = gc('OFERTAS_CADASTRADAS')
    ci_situ = gc('SITU_OFERTA'); ci_dt_ag = gc('DT_GERENCIADA'); ci_hr_ag = gc('HR_GERENCIADA')
    def gv(row, ci, default=''):
        try: return str(row[ci]).strip() if ci is not None and ci < len(row) else default
        except: return default
    def gn(row, ci):
        try: return float(str(row[ci]).replace(',','.').strip()) if ci is not None and ci < len(row) and row[ci] else 0
        except: return 0
    print(f"[{ts()}] Gerenciamento CSV: {len(data)} linhas, {len(header)} colunas")
    def extrair_ordem_exp(val):
        m = _re.match(r'O\.(\d+):\s*(.*)', str(val or ''))
        if m: return f'Ordem {m.group(1)}', m.group(2).strip()
        return '', str(val or '').strip()
    registros = []
    for r in data:
        polo = gv(r, ci_polo)
        if not polo: continue
        cat = gv(r, ci_cat); exp = gv(r, ci_exp); tutor = gv(r, ci_tutor); situ = gv(r, ci_situ)
        ordem, pratica = extrair_ordem_exp(exp)
        mat = int(gn(r, ci_mat)); agend = int(gn(r, ci_agend))
        capa = int(gn(r, ci_capa)); ofe = int(gn(r, ci_ofe))
        dt_ag = gv(r, ci_dt_ag); hr_ag = gv(r, ci_hr_ag)
        dt_ag_iso = ''
        if dt_ag and '/' in dt_ag:
            try:
                parts = dt_ag.split('/')
                dt_ag_iso = f"{parts[2]}-{parts[1]}-{parts[0]}"
            except: pass
        registros.append({
            'polo': polo, 'categoria': cat, 'pratica': pratica, 'ordem': ordem,
            'tutor': tutor if bool(tutor) else '',
            'tem_tutor': bool(tutor), 'tem_agenda': bool(dt_ag_iso),
            'gerenciado': ofe > 0, 'situ': situ,
            'alunos_mat': mat, 'alunos_agend': agend, 'capacidade': capa,
            'dt_agenda_iso': dt_ag_iso, 'hr_agenda': hr_ag,
        })
    df_r = pd.DataFrame(registros)
    if df_r.empty:
        return {'ger_kpis': {}, 'ger_polo': [], 'ger_cat': [], 'ger_ordem': [],
                'ger_contratacao': [], 'ger_agendas': [], 'ger_ofertas': []}
    polo_cat_tem_tutor = (
        df_r[df_r['tem_tutor']].groupby(['polo','categoria']).size()
        .reset_index(name='_qt').assign(_tem=True)
        .set_index(['polo','categoria'])['_tem']
    )
    def _fix_tem_tutor(row):
        return polo_cat_tem_tutor.get((row['polo'], row['categoria']), row['tem_tutor'])
    df_r['tem_tutor'] = df_r.apply(_fix_tem_tutor, axis=1)
    total = len(df_r); com_tutor = int(df_r['tem_tutor'].sum()); sem_tutor = total - com_tutor
    gerenciadas = int(df_r['gerenciado'].sum()); com_agenda = int(df_r['tem_agenda'].sum())
    tot_mat = int(df_r['alunos_mat'].sum()); tot_agend = int(df_r['alunos_agend'].sum())
    tot_capa = int(df_r['capacidade'].sum())
    print(f"[{ts()}] Gerenciamento: {total} ofertas, {gerenciadas} ger., {sem_tutor} sem tutor")
    ger_kpis = {
        'total_ofertas': total, 'ofertas_gerenciadas': gerenciadas,
        'ofertas_nao_gerenciadas': total - gerenciadas,
        'pct_gerenciado': round(gerenciadas/total*100,1) if total else 0,
        'ofertas_com_tutor': com_tutor, 'ofertas_sem_tutor': sem_tutor,
        'pct_com_tutor': round(com_tutor/total*100,1) if total else 0,
        'ofertas_com_agenda': com_agenda, 'total_alunos_matriculados': tot_mat,
        'total_alunos_agendados': tot_agend, 'total_capacidade': tot_capa,
        'pct_ocupacao': round(tot_agend/tot_capa*100,1) if tot_capa else 0,
        'polos_total': df_r['polo'].nunique(),
        'polos_sem_tutor': int(df_r[~df_r['tem_tutor']].groupby('polo').ngroups),
    }
    ger_polo = []
    for polo, grp in df_r.groupby('polo'):
        tuts = list(grp[grp['tem_tutor']]['tutor'].dropna().unique())
        ger_polo.append({
            'polo': str(polo), 'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()), 'sem_tutor': int((~grp['tem_tutor']).sum()),
            'com_agenda': int(grp['tem_agenda'].sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()), 'alunos_agendados': int(grp['alunos_agend'].sum()),
            'capacidade': int(grp['capacidade'].sum()), 'tutores_unicos': [str(t) for t in tuts],
        })
    ger_polo.sort(key=lambda x: -x['sem_tutor'])
    ger_cat = []
    for cat, grp in df_r.groupby('categoria'):
        ger_cat.append({
            'categoria': str(cat), 'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()), 'sem_tutor': int((~grp['tem_tutor']).sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()), 'alunos_agendados': int(grp['alunos_agend'].sum()),
        })
    ger_cat.sort(key=lambda x: -x['total_ofertas'])
    ger_ordem = []
    ordem_sort = {'Ordem 1':1,'Ordem 2':2,'Ordem 3':3,'Ordem 4':4,'Ordem 5':5}
    for ordem in sorted(df_r['ordem'].unique(), key=lambda x: ordem_sort.get(x,9)):
        if not ordem: continue
        grp = df_r[df_r['ordem']==ordem]
        ger_ordem.append({
            'ordem': ordem, 'total_ofertas': len(grp),
            'gerenciadas': int(grp['gerenciado'].sum()),
            'pct_gerenciado': round(grp['gerenciado'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['tem_tutor'].sum()),
            'alunos_matriculados': int(grp['alunos_mat'].sum()), 'alunos_agendados': int(grp['alunos_agend'].sum()),
            'dt_inicio': '', 'dt_fim': PRAZOS_ORDENS.get(ordem,''),
        })
    ger_contratacao = []
    for (polo, cat), grp in df_r.groupby(['polo','categoria']):
        tuts = list(grp[grp['tem_tutor']]['tutor'].dropna().unique())
        ger_contratacao.append({
            'polo': str(polo), 'categoria': str(cat), 'total_ofertas': len(grp),
            'tem_tutor': len(tuts)>0, 'tutores': [str(t) for t in tuts],
            'status': 'Contratado' if len(tuts)>0 else 'Sem tutor',
        })
    ger_agendas = []
    for polo, grp in df_r.groupby('polo'):
        total_p = len(grp); com_ag = int(grp['tem_agenda'].sum()); sem_ag = total_p - com_ag
        datas_por_cat = {}; datas_por_tutor = {}
        for _, row in grp[grp['tem_agenda']].iterrows():
            d = row['dt_agenda_iso']; c = row['categoria']; t = row['tutor'] or ''
            if d:
                if d not in datas_por_cat: datas_por_cat[d] = {'cats': [], 'tutores': []}
                if c and c not in datas_por_cat[d]['cats']: datas_por_cat[d]['cats'].append(c)
                if t and t not in datas_por_cat[d]['tutores']: datas_por_cat[d]['tutores'].append(t)
        # PATCH 7: estrutura completa de datas_por_tutor
        for d, v in datas_por_cat.items():
            datas_por_tutor[d] = v['tutores']
        ger_agendas.append({
            'polo': str(polo), 'total': total_p, 'com_agenda': com_ag, 'sem_agenda': sem_ag,
            'pct_agendado': round(com_ag/total_p*100, 1) if total_p else 0,
            'datas_agenda': sorted(datas_por_cat.keys()),
            'datas_por_cat': {d: v['cats'] for d, v in datas_por_cat.items()},
            'datas_por_tutor': datas_por_tutor,  # PATCH 7
        })
    ger_agendas.sort(key=lambda x: -x['sem_agenda'])
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
        'ger_kpis': ger_kpis, 'ger_polo': ger_polo, 'ger_cat': ger_cat,
        'ger_ordem': ger_ordem, 'ger_contratacao': ger_contratacao,
        'ger_agendas': ger_agendas, 'ger_ofertas': ger_ofertas_detalhe,
    }


def _processar_gerenciamento_novo(df_g):
    import re as _re
    col = {str(c).strip().upper(): c for c in df_g.columns}
    def gc(name): return col.get(name.upper())
    c_polo = gc('LABORATORIO'); c_cat = gc('CATEGORIA'); c_exp = gc('NOME_EXPERIMENTO')
    c_tutor = gc('TUTOR'); c_mat = gc('ALUNOS_MATRICULADOS'); c_agend = gc('ALUNOS_AGENDADOS')
    c_capa = gc('CAPACIDADE_TOTAL'); c_ofe = gc('OFERTAS_CADASTRADAS'); c_situ = gc('SITU_OFERTA')
    c_dt_ag = gc('DT_GERENCIADA'); c_hr_ag = gc('HR_GERENCIADA')
    def extrair_ordem_exp(val):
        m = _re.match(r'O\.(\d+):\s*(.*)', str(val or ''))
        if m: return f'Ordem {m.group(1)}', m.group(2).strip()
        return '', str(val or '').strip()
    df = df_g.copy()
    df['_POLO']  = df[c_polo].astype(str).str.strip() if c_polo else ''
    # PATCH 19: o export novo do gerenciamento (CSV 2026/2) usa um rótulo diferente
    # do arquivo antigo pra mesma categoria (Fisio/T.O./Estética) — normaliza pra
    # não virar "categoria fantasma" duplicada no seletor/agregações
    _CAT_RAW_NORM = {'FISIO-TO-EST-BIO (Multidisciplinar III)': 'BIO-FISIO-EST-TO (Multidisciplinar III)'}
    df['_CAT']   = df[c_cat].astype(str).str.strip().replace(_CAT_RAW_NORM) if c_cat  else ''
    df['_TUTOR'] = df[c_tutor].fillna('').astype(str).str.strip().replace('nan','') if c_tutor else ''
    df['_MAT']   = pd.to_numeric(df[c_mat],  errors='coerce').fillna(0).astype(int) if c_mat  else 0
    df['_AGEND'] = pd.to_numeric(df[c_agend],errors='coerce').fillna(0).astype(int) if c_agend else 0
    df['_CAPA']  = pd.to_numeric(df[c_capa], errors='coerce').fillna(0).astype(int) if c_capa  else 0
    df['_OFE']   = pd.to_numeric(df[c_ofe],  errors='coerce').fillna(0).astype(int) if c_ofe   else 0
    df['_TEM_TUTOR'] = df['_TUTOR'].str.len() > 0
    _situ_col = df[c_situ].fillna('').astype(str).str.strip() if c_situ else pd.Series([''] * len(df))
    # GERENCIADO = tem tutor E (tem ofertas cadastradas OU status CONCLUÍDO)
    df['_GERENCIADO'] = df['_TEM_TUTOR'] & ((df['_OFE'] > 0) | _situ_col.str.upper().str.contains('CONCLU', na=False))
    dt_col = df[c_dt_ag] if c_dt_ag else pd.Series([''] * len(df))
    def to_iso(v):
        if v is None: return ''
        try:
            import datetime as _dt
            if isinstance(v, (_dt.datetime, _dt.date)): return v.strftime('%Y-%m-%d')
        except: pass
        sv = str(v).strip()
        if not sv or sv == 'nan': return ''
        if '/' in sv:
            try:
                parts = sv.split('/')
                if len(parts) == 3: return f'{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}'
            except: pass
        if '-' in sv and len(sv) >= 10: return sv[:10]
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
    parsed = (df[c_exp] if c_exp else pd.Series([''] * len(df))).apply(extrair_ordem_exp)
    df['_ORDEM'] = parsed.apply(lambda x: x[0])
    df['_PRATICA'] = parsed.apply(lambda x: x[1])
    df = df[df['_POLO'].str.len() > 0].copy()
    total = len(df); com_tutor = int(df['_TEM_TUTOR'].sum()); gerenciadas = int(df['_GERENCIADO'].sum())
    com_agenda = int(df['_TEM_AGENDA'].sum())
    # FIX: Alunos Matriculados — deduplicar por polo×categoria (remove contagem múltipla por ordem)
    _mat_col = '_MAT'; _agend_col = '_AGEND'; _capa_col = '_CAPA'
    _raw_mat = int(df[_mat_col].sum()) if _mat_col in df.columns else 0
    _grp_cols_ok = ['_POLO','_CAT']
    if all(c in df.columns for c in _grp_cols_ok + [_mat_col, _agend_col, _capa_col]):
        _dedup = df.groupby(_grp_cols_ok)[[_mat_col, _agend_col, _capa_col]].max()
        tot_mat   = int(_dedup[_mat_col].sum())
        tot_agend = int(_dedup[_agend_col].sum())
        tot_capa  = int(_dedup[_capa_col].sum())
        print(f"[{ts()}] Alunos DEDUPLICADOS por polo×cat: {tot_mat:,} (bruto era {_raw_mat:,}, redução: {_raw_mat-tot_mat:,})")
    else:
        tot_mat   = _raw_mat
        tot_agend = int(df[_agend_col].sum()) if _agend_col in df.columns else 0
        tot_capa  = int(df[_capa_col].sum())  if _capa_col  in df.columns else 0
        print(f"[{ts()}] Alunos sem dedup: {tot_mat:,}")
    print(f"[{ts()}] Gerenciamento: {total} ofertas, {gerenciadas} ger., {total-com_tutor} sem tutor")
    print(f"[{ts()}] Agendas: {com_agenda} · datas: {sorted(df[df['_TEM_AGENDA']]['_DT_AG_ISO'].head(3).tolist())}")
    print(f"[{ts()}] {df['_POLO'].nunique()} polos, {df['_CAT'].nunique()} cats, {df['_ORDEM'].nunique()} ordens")
    ger_kpis = {
        'total_ofertas': total, 'ofertas_gerenciadas': gerenciadas,
        'ofertas_nao_gerenciadas': total - gerenciadas,
        'pct_gerenciado': round(gerenciadas/total*100,1) if total else 0,
        'ofertas_com_tutor': com_tutor, 'ofertas_sem_tutor': total-com_tutor,
        'pct_com_tutor': round(com_tutor/total*100,1) if total else 0,
        'ofertas_com_agenda': com_agenda, 'total_alunos_matriculados': tot_mat,
        'total_alunos_agendados': tot_agend, 'total_capacidade': tot_capa,
        'pct_ocupacao': round(tot_agend/tot_capa*100,1) if tot_capa else 0,
        'polos_total': df['_POLO'].nunique(),
        'polos_sem_tutor': int(df[~df['_TEM_TUTOR']].groupby('_POLO').ngroups),
    }
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
    ger_ordem = []; ordem_sort = {'Ordem 1':1,'Ordem 2':2,'Ordem 3':3,'Ordem 4':4,'Ordem 5':5}
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
    ger_contratacao = []
    for (polo, cat), grp in df.groupby(['_POLO','_CAT']):
        tuts = list(grp[grp['_TEM_TUTOR']]['_TUTOR'].dropna().unique())
        ger_contratacao.append({
            'polo': str(polo), 'categoria': str(cat), 'total_ofertas': len(grp),
            'tem_tutor': len(tuts)>0, 'tutores': [str(t) for t in tuts],
            'status': 'Contratado' if len(tuts)>0 else 'Sem tutor',
        })
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
            'datas_por_cat': datas_por_cat,
            'datas_por_tutor': datas_por_tutor,  # PATCH 7: preservado
        })
    ger_agendas.sort(key=lambda x: -x['sem_agenda'])
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


def processar_gerenciamento_semestres(arquivos):
    """
    PATCH 18: lê 1+ arquivos de gerenciamento (cada um com um semestre padrão de
    fallback) e devolve {semestre: ger_dados_dict}. Quando o arquivo tem coluna
    SEMESTRE (export novo), usa o valor da própria linha como fonte de verdade —
    não confia só em "qual arquivo é qual semestre".
    arquivos: lista de (path, semestre_fallback)
    """
    frames_novo = []
    resultado = {}
    for path, fallback_sem in arquivos:
        if not path or not os.path.isfile(path):
            continue
        try:
            df = _ler_arquivo_gerenciamento(path)
        except Exception as e:
            print(f"[{ts()}] ERRO ao ler {os.path.basename(path)}: {e}")
            continue
        cols_upper = [str(c).upper() for c in df.columns]
        is_novo = 'LABORATORIO' in cols_upper and 'NOME_EXPERIMENTO' in cols_upper
        if not is_novo:
            print(f"[{ts()}] {os.path.basename(path)}: formato ANTIGO — todo o arquivo tratado como {fallback_sem}")
            resultado[fallback_sem] = processar_gerenciamento(path)
            continue
        sem_col = next((c for c in df.columns if str(c).upper() == 'SEMESTRE'), None)
        if sem_col:
            df['_SEM_ROW'] = df[sem_col].astype(str).str.strip()
            _fora = ~df['_SEM_ROW'].isin(ALL_SEMESTRES.keys())
            if _fora.any():
                print(f"[{ts()}] {os.path.basename(path)}: {int(_fora.sum())} linhas com SEMESTRE não reconhecido — usando fallback {fallback_sem}")
            df.loc[_fora, '_SEM_ROW'] = fallback_sem
        else:
            df['_SEM_ROW'] = fallback_sem
        frames_novo.append(df)
    if frames_novo:
        df_all = pd.concat(frames_novo, ignore_index=True)
        for sem, grp in df_all.groupby('_SEM_ROW'):
            grp2 = grp.drop(columns=['_SEM_ROW'])
            print(f"[{ts()}] Gerenciamento {sem}: {len(grp2)} linhas")
            resultado[sem] = _processar_gerenciamento_novo(grp2)
    return resultado


def processar_gerenciamento(p3):
    print(f"[{ts()}] Lendo gerenciamento...")
    df_g = ler_excel(p3)
    print(f"[{ts()}] Gerenciamento: {len(df_g)} linhas, {len(df_g.columns)} colunas")
    cols_upper = [str(c).upper() for c in df_g.columns]
    is_novo = 'LABORATORIO' in cols_upper and 'NOME_EXPERIMENTO' in cols_upper
    if is_novo:
        print(f"[{ts()}] Formato: NOVO (relatório detalhado)")
        return _processar_gerenciamento_novo(df_g)
    print(f"[{ts()}] Formato: ANTIGO (GIOCONDA)")
    def gcol(df, *partes):
        for c in df.columns:
            cu = str(c).upper()
            if all(p.upper() in cu for p in partes): return c
        return None
    c_polo = gcol(df_g, 'CEEM', 'RSOC') or 'CEEM_RSOC'
    c_cat  = gcol(df_g, 'CATP', 'NOME') or 'CATP_NOME'
    c_lab  = gcol(df_g, 'LABE', 'NOME') or 'LABE_NOME'
    c_curso = gcol(df_g, 'NOME', 'CURS') or 'NOME_CURS'
    c_situ = gcol(df_g, 'SITU') or 'SITU'
    c_alunos = gcol(df_g, 'ALUNOS', 'MATRIC') or 'ALUNOS_MATRICULADOS'
    c_capa_exp = gcol(df_g, 'CAPA', 'EXP') or 'CAPA_EXP'
    c_ofe_cad = gcol(df_g, 'OFE', 'CAD') or 'OFE_CAD'
    c_qtd_alun = gcol(df_g, 'QTD', 'ALUN') or 'QTD_ALUN'
    c_tutor = gcol(df_g, 'TUTOR') or 'TUTOR'
    c_dt_agenda = gcol(df_g, 'DT', 'GERENCIADA') or 'DT_GERENCIADA'
    c_hr_agenda = gcol(df_g, 'HR', 'GERENCIADA') or 'HR_GERENCIADA'
    c_ofex_dtin = gcol(df_g, 'OFEX', 'DTIN') or 'OFEX_DTIN'
    c_ofex_dtfi = gcol(df_g, 'OFEX', 'DTFI') or 'OFEX_DTFI'
    if c_situ in df_g.columns:
        df_g = df_g[df_g[c_situ].astype(str).str.strip().str.upper() == 'ATIVO'].copy()
    print(f"[{ts()}] Gerenciamento após filtro ativos: {len(df_g)} linhas")
    df_g['_ORDEM_G'] = ''; df_g['_PRATICA_G'] = ''
    if c_lab in df_g.columns:
        import re
        def extrair_ordem(val):
            val = str(val or '')
            m = re.match(r'O\.(\d+):\s*(.*)', val)
            if m: return f'Ordem {m.group(1)}', m.group(2).strip()
            return '', val.strip()
        parsed = df_g[c_lab].apply(extrair_ordem)
        df_g['_ORDEM_G'] = parsed.apply(lambda x: x[0])
        df_g['_PRATICA_G'] = parsed.apply(lambda x: x[1])
    # FIX BUG 1: _TEM_TUTOR deve ser definido ANTES de _GERENCIADO
    df_g['_TEM_TUTOR'] = df_g[c_tutor].notna() & (df_g[c_tutor].astype(str).str.strip() != '') & (df_g[c_tutor].astype(str).str.strip().str.upper() != 'NAN')
    # GERENCIADO = tem tutor E (tem ofertas cadastradas OU status CONCLUÍDO)
    _situ_g = df_g[c_situ].fillna('').astype(str).str.strip() if c_situ and c_situ in df_g.columns else pd.Series([''] * len(df_g))
    df_g['_GERENCIADO'] = df_g['_TEM_TUTOR'] & ((pd.to_numeric(df_g.get(c_ofe_cad, 0), errors='coerce').fillna(0) > 0) | _situ_g.str.upper().str.contains('CONCLU', na=False))
    df_g['_TEM_AGENDA'] = df_g.get(c_dt_agenda, pd.Series(dtype='object')).notna()
    df_g['_ALUNOS_MAT'] = pd.to_numeric(df_g.get(c_alunos, 0), errors='coerce').fillna(0).astype(int)
    df_g['_QTD_ALUN'] = pd.to_numeric(df_g.get(c_qtd_alun, 0), errors='coerce').fillna(0).astype(int)
    df_g['_CAPA'] = pd.to_numeric(df_g.get(c_capa_exp, 0), errors='coerce').fillna(0).astype(int)
    total_ofertas = len(df_g); gerenciadas = int(df_g['_GERENCIADO'].sum())
    com_tutor = int(df_g['_TEM_TUTOR'].sum()); sem_tutor = total_ofertas - com_tutor
    # FIX: Alunos Matriculados — deduplicar por polo×categoria (soma bruta conta os mesmos alunos por ordem)
    # Usar apenas colunas que REALMENTE existem (não fallbacks)
    _c_polo_real = c_polo if (c_polo and c_polo in df_g.columns) else None
    _c_cat_real  = c_cat  if (c_cat  and c_cat  in df_g.columns) else None
    # Se nenhuma das buscas primárias funcionou, tentar qualquer coluna polo/cat
    if not _c_polo_real:
        _c_polo_real = next((c for c in df_g.columns if 'POLO' in str(c).upper() or 'CEEM' in str(c).upper()), None)
    if not _c_cat_real:
        _c_cat_real = next((c for c in df_g.columns if 'CATEG' in str(c).upper() or 'CATP' in str(c).upper()), None)
    _raw_mat = int(df_g['_ALUNOS_MAT'].sum())
    if _c_polo_real and _c_cat_real:
        _dedup_g = df_g.groupby([_c_polo_real, _c_cat_real])[['_ALUNOS_MAT','_QTD_ALUN','_CAPA']].max()
        tot_mat   = int(_dedup_g['_ALUNOS_MAT'].sum())
        tot_agend = int(_dedup_g['_QTD_ALUN'].sum())
        tot_capa  = int(_dedup_g['_CAPA'].sum())
        print(f"[{ts()}] Alunos DEDUPLICADOS por polo×cat: {tot_mat:,} (bruto era {_raw_mat:,}, redução: {_raw_mat-tot_mat:,})")
    else:
        tot_mat = _raw_mat
        tot_agend = int(df_g['_QTD_ALUN'].sum())
        tot_capa  = int(df_g['_CAPA'].sum())
        print(f"[{ts()}] Alunos sem dedup (colunas polo/cat não encontradas): {tot_mat:,}")
    polos_total = df_g[c_polo].nunique() if c_polo in df_g.columns else 0
    polos_sem_tutor_count = int(df_g[~df_g['_TEM_TUTOR']].groupby(c_polo).ngroups) if c_polo in df_g.columns else 0
    ger_kpis = {
        'total_ofertas': total_ofertas, 'ofertas_gerenciadas': gerenciadas,
        'ofertas_nao_gerenciadas': total_ofertas - gerenciadas,
        'pct_gerenciado': round(gerenciadas/total_ofertas*100,1) if total_ofertas else 0,
        'ofertas_com_tutor': com_tutor, 'ofertas_sem_tutor': sem_tutor,
        'pct_com_tutor': round(com_tutor/total_ofertas*100,1) if total_ofertas else 0,
        'ofertas_com_agenda': int(df_g['_TEM_AGENDA'].sum()),
        'total_alunos_matriculados': tot_mat, 'total_alunos_agendados': tot_agend,
        'total_capacidade': tot_capa,
        'pct_ocupacao': round(tot_agend/tot_capa*100,1) if tot_capa else 0,
        'polos_total': polos_total, 'polos_sem_tutor': polos_sem_tutor_count,
    }
    print(f"[{ts()}] Gerenciamento: {total_ofertas} ofertas, {gerenciadas} ger., {sem_tutor} sem tutor")
    ger_polo = []
    if c_polo in df_g.columns:
        for polo, grp in df_g.groupby(c_polo):
            ger_polo.append({
                'polo': str(polo), 'total_ofertas': len(grp),
                'gerenciadas': int(grp['_GERENCIADO'].sum()),
                'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
                'com_tutor': int(grp['_TEM_TUTOR'].sum()), 'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
                'com_agenda': int(grp['_TEM_AGENDA'].sum()),
                'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()), 'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
                'capacidade': int(grp['_CAPA'].sum()),
                'tutores_unicos': list(grp[grp['_TEM_TUTOR']][c_tutor].dropna().unique()),
            })
        ger_polo.sort(key=lambda x: -x['sem_tutor'])
    ger_cat = []
    if c_cat in df_g.columns:
        for cat, grp in df_g.groupby(c_cat):
            ger_cat.append({
                'categoria': str(cat), 'total_ofertas': len(grp),
                'gerenciadas': int(grp['_GERENCIADO'].sum()),
                'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
                'com_tutor': int(grp['_TEM_TUTOR'].sum()), 'sem_tutor': int((~grp['_TEM_TUTOR']).sum()),
                'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()), 'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
            })
        ger_cat.sort(key=lambda x: -x['total_ofertas'])
    ger_ordem = []
    ordens_validas = [o for o in sorted(df_g['_ORDEM_G'].unique()) if o and 'Ordem' in str(o)]
    for ordem in ordens_validas:
        grp = df_g[df_g['_ORDEM_G'] == ordem]
        datas_inicio = pd.to_datetime(grp.get(c_ofex_dtin, pd.Series(dtype='object')), errors='coerce').dropna()
        datas_fim = pd.to_datetime(grp.get(c_ofex_dtfi, pd.Series(dtype='object')), errors='coerce').dropna()
        dt_inicio = datas_inicio.min().strftime('%d/%m/%Y') if len(datas_inicio) > 0 else ''
        dt_fim = datas_fim.max().strftime('%d/%m/%Y') if len(datas_fim) > 0 else ''
        ger_ordem.append({
            'ordem': ordem, 'total_ofertas': len(grp),
            'gerenciadas': int(grp['_GERENCIADO'].sum()),
            'pct_gerenciado': round(grp['_GERENCIADO'].sum()/len(grp)*100,1) if len(grp) else 0,
            'com_tutor': int(grp['_TEM_TUTOR'].sum()),
            'alunos_matriculados': int(grp['_ALUNOS_MAT'].sum()), 'alunos_agendados': int(grp['_QTD_ALUN'].sum()),
            'dt_inicio': dt_inicio, 'dt_fim': dt_fim,
        })
    ger_contratacao = []
    if c_polo in df_g.columns and c_cat in df_g.columns:
        for (polo, cat), grp in df_g.groupby([c_polo, c_cat]):
            tutores_list = list(grp[grp['_TEM_TUTOR']][c_tutor].dropna().unique())
            ger_contratacao.append({
                'polo': str(polo), 'categoria': str(cat), 'total_ofertas': len(grp),
                'tem_tutor': len(tutores_list)>0, 'tutores': [str(t) for t in tutores_list],
                'status': 'Contratado' if len(tutores_list)>0 else 'Sem tutor',
            })
        ger_contratacao.sort(key=lambda x: (0 if x['tem_tutor'] else 1, x['polo']))
    ger_agendas = []
    if c_polo in df_g.columns:
        for polo, grp in df_g.groupby(c_polo):
            total = len(grp); com_agenda = int(grp['_TEM_AGENDA'].sum()); sem_agenda = total - com_agenda
            datas = []; datas_por_cat = {}; datas_por_tutor = {}
            if c_dt_agenda and c_dt_agenda in grp.columns:
                for _, ag_row in grp[grp['_TEM_AGENDA']].iterrows():
                    dt_val = pd.to_datetime(ag_row.get(c_dt_agenda), errors='coerce')
                    if pd.notna(dt_val):
                        dt_str = dt_val.strftime('%Y-%m-%d')
                        cat_val = str(ag_row.get(c_cat, '') or '')
                        tutor_val = str(ag_row.get(c_tutor, '') or '')
                        if dt_str not in datas: datas.append(dt_str)
                        if cat_val:
                            if dt_str not in datas_por_cat: datas_por_cat[dt_str] = []
                            if cat_val not in datas_por_cat[dt_str]: datas_por_cat[dt_str].append(cat_val)
                        if tutor_val and tutor_val != 'nan':
                            if dt_str not in datas_por_tutor: datas_por_tutor[dt_str] = []
                            if tutor_val not in datas_por_tutor[dt_str]: datas_por_tutor[dt_str].append(tutor_val)
                datas = sorted(set(datas))
            ger_agendas.append({
                'polo': str(polo), 'total': total, 'com_agenda': com_agenda, 'sem_agenda': sem_agenda,
                'pct_agendado': round(com_agenda/total*100, 1) if total else 0,
                'datas_agenda': datas, 'datas_por_cat': datas_por_cat,
                'datas_por_tutor': datas_por_tutor,  # PATCH 7: preservado
            })
        ger_agendas.sort(key=lambda x: -x['sem_agenda'])
    ger_ofertas_detalhe = []
    for _, row in df_g.iterrows():
        ger_ofertas_detalhe.append({
            'polo': str(row.get(c_polo, '')), 'categoria': str(row.get(c_cat, '')),
            'ordem': str(row.get('_ORDEM_G', '')), 'pratica': str(row.get('_PRATICA_G', '')),
            'curso': str(row.get(c_curso, '')),
            'tutor': str(row.get(c_tutor, '')) if pd.notna(row.get(c_tutor)) else '',
            'gerenciado': bool(row.get('_GERENCIADO', False)),
            'tem_agenda': bool(row.get('_TEM_AGENDA', False)),
            'alunos_mat': int(row.get('_ALUNOS_MAT', 0)), 'alunos_agend': int(row.get('_QTD_ALUN', 0)),
            'capacidade': int(row.get('_CAPA', 0)),
        })
    print(f"[{ts()}] Gerenciamento: {len(ger_polo)} polos, {len(ger_cat)} cats, {len(ger_ordem)} ordens")
    return {
        'ger_kpis': ger_kpis, 'ger_polo': ger_polo, 'ger_cat': ger_cat,
        'ger_ordem': ger_ordem, 'ger_contratacao': ger_contratacao,
        'ger_agendas': ger_agendas, 'ger_ofertas': ger_ofertas_detalhe,
    }





def carregar_alunos_hub(path_csv):
    """
    Lê Relatorio_alunos_por_hub.csv e retorna dict com matrículas distintas
    por polo e por categoria — substitui a contagem inflacionada do GIOCONDA.
    """
    import unicodedata as _ud, re as _re
    if not path_csv or not os.path.isfile(path_csv):
        print(f"[{ts()}] Alunos hub: arquivo não encontrado ({path_csv})")
        return None
    print(f"[{ts()}] Lendo alunos por hub: {os.path.basename(path_csv)}")
    # Verificar se o arquivo é HTML (download falhou) e não CSV real
    try:
        with open(path_csv, 'rb') as _f: _head = _f.read(500)
        if b'<!DOCTYPE' in _head or b'<html' in _head.lower() or b'<HTML' in _head:
            sz = os.path.getsize(path_csv)
            print(f"[{ts()}] ERRO: Relatorio_alunos_por_hub.csv é HTML ({sz} bytes) — o download falhou")
            print(f"[{ts()}] SOLUÇÃO: Atualize o secret URL_ALUNOS_HUB para o formato download.aspx:")
            print(f"[{ts()}]   https://uniasselvi01-my.sharepoint.com/personal/[USUARIO]/_layouts/15/download.aspx?share=[TOKEN]")
            return None
    except: pass
    for enc in ['latin-1', 'utf-8', 'cp1252']:
        try:
            df = pd.read_csv(path_csv, sep=';', encoding=enc, dtype=str)
            if 'MATRICULA' in df.columns: break
        except: continue
    else:
        print(f"[{ts()}] ERRO: não foi possível ler {path_csv}")
        return None

    # Apenas matrículas confirmadas
    if 'SITUACAO_SEMESTRE' in df.columns:
        df = df[df['SITUACAO_SEMESTRE'].str.strip() == 'Matrícula Confirmada'].copy()

    def _norm(s):
        s = _ud.normalize('NFD', str(s or '').upper().strip())
        s = ''.join(c for c in s if _ud.category(c) != 'Mn')
        s = _re.sub(r'^LAP\s*[-–]\s*', '', s).strip()
        return _re.sub(r'\s+', ' ', s)

    # Mapear GRUPO_HUB → nossas categorias
    GRUPO_CAT = {
        'MULTIDISCIPLINAR II':        'ENF-INS (Multidisciplinar II)',
        'MULTIDISCIPLINAR I':         'BIO-FAR (Multidisciplinar I)',
        'MULTIDISCIPLINAR III':       'BIO-FISIO-EST-TO (Multidisciplinar III)',
        'ENGMAKER+QUIMICA E FISICA':  'QUÍMICA E FÍSICA',
        'ENGMAKER':                   'ENGMAKER',
        'MULTIDISCIPLINAR IV':        'NUTRI (Multidisciplinar IV)',
    }
    def _grupo_para_cat(g):
        gn = _norm(g)
        # Match EXATO primeiro (evita 'MULTIDISCIPLINAR I' casar com 'MULTIDISCIPLINAR II')
        for k, v in GRUPO_CAT.items():
            if _norm(k) == gn: return v
        # Fallback: contém (só para casos como 'ENGMAKER+...' vs 'ENGMAKER')
        for k, v in GRUPO_CAT.items():
            kn = _norm(k)
            if kn in gn and len(kn) > 8: return v
        return g

    df['_POLO_NORM'] = df['POLO_HUB'].apply(_norm)
    df['_CAT']       = df['GRUPO_HUB'].apply(_grupo_para_cat)

    total_distintos = df['MATRICULA'].nunique()
    print(f"[{ts()}] Matrículas DISTINTAS (ativos): {total_distintos:,}")

    # Por polo (chave normalizada)
    por_polo = (df.groupby('_POLO_NORM')['MATRICULA']
                  .nunique().to_dict())

    # Por polo × categoria
    por_polo_cat = {}
    for (polo, cat), grp in df.groupby(['_POLO_NORM', '_CAT']):
        por_polo_cat[f"{polo}||{cat}"] = int(grp['MATRICULA'].nunique())

    # Por categoria (totais)
    por_cat = (df.groupby('_CAT')['MATRICULA']
                 .nunique().to_dict())

    # ── Mapear TUTOR_PRATICA → subcurso para Multi 3 ────────────────────
    tutor_subcurso = {}  # nome_norm → 'Fisio'/'T.Oc'/'Est'
    if 'TUTOR_PRATICA' in df.columns and 'DISCIPLINA' in df.columns and 'GRUPO_HUB' in df.columns:
        import re as _re
        from collections import Counter as _Counter
        _FISIO = ['FISIOTERAPIA','CINESIOTERAPIA','ELETROTERM','CARDIORRESPIR',
                  'PROTESE','ORTESE','RECURSOS TERAPEUTICOS','MOVIMENTO FUNCIONAL',
                  'AVALIACAO FISICO','REABILITACAO','NEUROFUNC','ORTOPEDIC','RESPIRATORIA']
        _TO    = ['TERAPIA OCUPACIONAL','PSICOMOTRICIDADE','INTEGRACAO SENSORIAL',
                  'TRANSTORNOS MENTAIS','COMPORTAMENTO HUMANO','VIDA DIARIA','TRABALHO EM GRUPO']
        _EST   = ['ESTETICA','COSMETOLOGIA','BIOMEDICINA ESTETICA','PIGMENTAC',
                  'DEPILAC','FACIAL CORPORAL','MICROAGULH']
        def _classif_disc(d):
            d2 = _norm(d) if d else ''
            if any(k in d2 for k in _FISIO): return 'Fisio'
            if any(k in d2 for k in _TO):    return 'T.Oc'
            if any(k in d2 for k in _EST):   return 'Est'
            return None
        def _norm_tutor(s):
            s = _re.sub(r'\s*\(\d+\)\s*$', '', str(s or '')).strip()
            return _norm(s)
        df3 = df[df['GRUPO_HUB'].str.upper().str.contains('MULTIDISCIPLINAR III|MULTI.*3|BIO-FISIO', na=False)].copy()
        df3 = df3[df3['TUTOR_PRATICA'].notna() & (df3['TUTOR_PRATICA'].astype(str).str.strip().str.upper() != 'NAN')]
        df3['_sub'] = df3['DISCIPLINA'].apply(_classif_disc)
        df3['_tnorm'] = df3['TUTOR_PRATICA'].apply(_norm_tutor)
        for tutor, grp in df3[df3['_sub'].notna()].groupby('_tnorm'):
            subs = list(grp['_sub'])
            if subs:
                # Guardar em MINÚSCULAS para o JS (que usa normN = toLowerCase)
                tutor_lower = tutor.lower()
                tutor_subcurso[tutor_lower] = _Counter(subs).most_common(1)[0][0]
                # Também guardar primeiro+último nome (fallback)
                parts = tutor_lower.split()
                if len(parts) >= 2:
                    fl = parts[0] + ' ' + parts[-1]
                    if fl not in tutor_subcurso:
                        tutor_subcurso[fl] = tutor_subcurso[tutor_lower]
        print(f"[{ts()}] Subcursos Multi 3 mapeados: {len(tutor_subcurso)} tutores")

    return {
        'total_distintos': int(total_distintos),
        'por_polo': {k: int(v) for k, v in por_polo.items()},
        'por_polo_cat': por_polo_cat,
        'por_cat': {k: int(v) for k, v in por_cat.items()},
        'tutor_subcurso': tutor_subcurso,  # Multi 3: nome_tutor → Fisio/T.Oc/Est
    }

# Senha de acesso ao dashboard (mesma da tela de login)
SENHA_DASHBOARD = "uniasselvi2026"

# PATCH 8: cifra o JSON antes de injetar no HTML — sem isso, dava pra ver
# tudo no Ctrl+U mesmo sem digitar a senha
def cifrar_dados(dados_json_str, senha):
    chave = hashlib.sha256(senha.encode('utf-8')).digest()  # 32 bytes → AES-256
    aesgcm = AESGCM(chave)
    iv = os.urandom(12)
    ct = aesgcm.encrypt(iv, dados_json_str.encode('utf-8'), None)
    iv_b64 = base64.b64encode(iv).decode('ascii')
    ct_b64 = base64.b64encode(ct).decode('ascii')
    return f"{iv_b64}:{ct_b64}"

def gerar_html(dados):
    saida = os.path.join(SCRIPT_DIR, "saida")
    os.makedirs(saida, exist_ok=True)
    output = os.path.join(saida, "dashboard.html")
    tmpl   = os.path.join(SCRIPT_DIR, "template_dashboard.html")
    with open(tmpl, encoding='utf-8') as f: html = f.read()
    json_str = json.dumps(dados, ensure_ascii=False)
    payload_cifrado = cifrar_dados(json_str, SENHA_DASHBOARD)
    html = html.replace("'DATA_GOES_HERE'", json.dumps(payload_cifrado))
    html = html.replace("TIMESTAMP_GOES_HERE", dados['gerado_em'])
    with open(output, 'w', encoding='utf-8') as f: f.write(html)
    print(f"[{ts()}] Salvo: {output} (JSON cifrado com AES-256-GCM, {len(payload_cifrado)} chars)")

    # PATCH 9: lookup público (não cifrado) só com email/nome/polo/categoria,
    # pro portfolio_form.html autopreencher sem precisar da senha do dashboard
    lookup = [
        {'email': t.get('email',''), 'n': t.get('n',''), 'p': t.get('p',''), 'c': t.get('c','')}
        for t in dados.get('tutores', [])
        if t.get('email') and not t.get('_anonimo') and t.get('c') != 'Aviso de Portfólio'
    ]
    lookup_path = os.path.join(saida, "lookup.json")
    with open(lookup_path, 'w', encoding='utf-8') as f:
        json.dump(lookup, f, ensure_ascii=False)
    print(f"[{ts()}] Salvo: {lookup_path} ({len(lookup)} tutores, sem cifra)")
    return output


def modo_watch(p1, p2):
    print(f"[{ts()}] Monitorando a cada 30s — feche a janela para parar")
    mods = {p1: 0.0, p2: 0.0}
    def loop():
        while True:
            try:
                mudou = any(os.path.getmtime(a) != mods[a] for a in mods if os.path.isfile(a))
                if mudou:
                    for a in mods:
                        if os.path.isfile(a): mods[a] = os.path.getmtime(a)
                    print(f"[{ts()}] Mudança detectada, atualizando...")
                    gerar_html(processar(p1, p2))
            except Exception as e: print(f"[{ts()}] Erro: {e}")
            time.sleep(30)
    threading.Thread(target=loop, daemon=True).start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt: print(f"\n[{ts()}] Encerrado.")


if __name__ == '__main__':
    print()
    print(" Verificando arquivos...")
    print()
    p1, p2, tmpl, p3, p3b, p4, p5 = verificar_e_localizar()
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
    if p4:
        try:
            lotacao = carregar_lotacao(p4)
            dados = enriquecer_tutores(dados, lotacao)
        except Exception as e:
            print(f"[{ts()}] AVISO: Erro ao processar lotação: {e}")
            dados['alunos_por_curso'] = []
    else:
        dados['alunos_por_curso'] = []
    # PATCH 2: tem_lotacao baseado em dados reais (CH > 0 em pelo menos 1 tutor)
    _ch_ok = sum(1 for t in dados.get('tutores', []) if t.get('ch_semanal') and t['ch_semanal'] > 0)
    dados['tem_lotacao'] = _ch_ok > 0
    print(f"[{ts()}] tem_lotacao={dados['tem_lotacao']} ({_ch_ok} tutores com CH SEMANAL)")
    if p3 or p3b:
        try:
            # PATCH 18: cada arquivo tem um semestre de fallback (usado só quando a
            # linha não tem coluna SEMESTRE reconhecível) — arquivo antigo -> mais
            # antigo dos semestres carregados; arquivo "_26_02" -> 2026/2 explícito
            _sem_mais_antigo = sorted(ALL_SEMESTRES.keys())[0]
            ger_por_semestre = processar_gerenciamento_semestres([
                (p3,  _sem_mais_antigo),
                (p3b, '2026/2' if '2026/2' in ALL_SEMESTRES else sorted(ALL_SEMESTRES.keys())[-1]),
            ])
            dados['gerenciamento_por_semestre'] = ger_por_semestre
            for _sk, _sv in ger_por_semestre.items():
                print(f"[{ts()}] Gerenciamento {_sk}: {_sv['ger_kpis']['total_ofertas']} ofertas, {_sv['ger_kpis']['ofertas_gerenciadas']} ger.")
            # dados['ger_*'] no nível raiz = semestre ativo do dashboard (compat
            # com todo o código de enriquecimento abaixo, que sempre operou em
            # cima de um único conjunto de ofertas)
            ger_dados = ger_por_semestre.get(SEMESTRE_ATUAL) or next(iter(ger_por_semestre.values()), {})
            dados.update(ger_dados)
            dados['tem_gerenciamento'] = True
            # ── Injetar gerenciamento nos tutores (ger_pct, ger_ok, ger_total) ──────
            import unicodedata as _ud5, re as _re6
            def _norm_ger2(s):
                s = _ud5.normalize('NFD', str(s or '').lower().strip())
                s = ''.join(ch for ch in s if _ud5.category(ch) != 'Mn')
                s = _re6.sub(r'\s*\(\d+\)\s*$', '', s).strip()
                return _re6.sub(r'\s+', ' ', s)
            def _fl_ger2(s):
                pts = _norm_ger2(s).split()
                return f"{pts[0]} {pts[-1]}" if len(pts) >= 2 else _norm_ger2(s)
            _ger_idx2 = {}
            for _g2 in dados.get('ger_ofertas', []):
                _gn2 = (_g2.get('tutor') or '').strip()
                if not _gn2: continue
                for _k2 in [_norm_ger2(_gn2), _fl_ger2(_gn2)]:
                    if _k2 not in _ger_idx2:
                        _ger_idx2[_k2] = {'ger': 0, 'total': 0}
                    _ger_idx2[_k2]['total'] += 1
                    if _g2.get('gerenciado'): _ger_idx2[_k2]['ger'] += 1
            _ger_matched2 = 0
            for _t2 in dados.get('tutores', []):
                _tn2 = _t2.get('n', '')
                _gd2 = _ger_idx2.get(_norm_ger2(_tn2)) or _ger_idx2.get(_fl_ger2(_tn2))
                if _gd2 and _gd2['total'] > 0:
                    _t2['ger_total'] = _gd2['total']
                    _t2['ger_ok']    = _gd2['ger']
                    _t2['ger_pct']   = round(_gd2['ger'] / _gd2['total'] * 100)
                    _ger_matched2 += 1
                else:
                    _t2['ger_total'] = 0
                    _t2['ger_ok']    = 0
                    _t2['ger_pct']   = None
            print(f"[{ts()}] Gerenciamento injetado nos tutores: {_ger_matched2}/{len(dados.get('tutores',[]))} matches")
            # ────────────────────────────────────────────────────────────────────────
            # Enriquecer ger_ofertas com ch_semanal (join por nome normalizado)
            def _norm_nome(s):
                import unicodedata
                s = str(s or '').lower().split('(')[0].strip()
                s = unicodedata.normalize('NFD', s)
                s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
                return ' '.join(s.split())
            def _nome_fl(s):
                pts = _norm_nome(s).split()
                return (pts[0] + ' ' + pts[-1]) if len(pts) >= 2 else _norm_nome(s)
            # Mapear ch_semanal por nome completo E por primeiro+último nome
            _ch_map = {}; _ch_map_fl = {}
            # Fonte 1: portfólio tutores (com CH já enriquecida pela lotação)
            for t in dados.get('tutores', []):
                if t.get('ch_semanal') and t.get('n'):
                    _ch_map[_norm_nome(t['n'])] = t['ch_semanal']
                    _ch_map_fl[_nome_fl(t['n'])] = t['ch_semanal']
            # Fonte 2: lotação DIRETAMENTE (589 tutores vs 298 do portfólio)
            # Fix principal: tutores que estão no GIOCONDA mas não no portfólio
            _lotacao_safe = lotacao if 'lotacao' in dir() and lotacao else {}
            if _lotacao_safe:
                for lot_nome, lot_info in lotacao.items():
                    lot_ch = lot_info.get('ch_semanal', 0) if isinstance(lot_info, dict) else 0
                    if lot_ch:
                        if lot_nome not in _ch_map:
                            _ch_map[lot_nome] = lot_ch
                        lot_fl = _nome_fl(lot_nome)
                        if lot_fl not in _ch_map_fl:
                            _ch_map_fl[lot_fl] = lot_ch
            print(f"[{ts()}] CH map Gerenciamento: {len(_ch_map)} entradas")
            # Injetar ch_semanal em cada oferta
            enr = 0
            # Pré-computar lista de (nome_normalizado, nome_fl, ch) para lookup rápido
            _lot_list = [(k, _nome_fl(k), v) for k, v in _ch_map.items()]

            for oferta in dados.get('ger_ofertas', []):
                tutor = oferta.get('tutor', '')
                if not tutor or oferta.get('ch_semanal'): continue
                tn = _norm_nome(tutor); tfl = _nome_fl(tutor)

                # Match 1: exato ou FL
                ch = _ch_map.get(tn) or _ch_map.get(tfl) or _ch_map_fl.get(tn) or _ch_map_fl.get(tfl)

                # Match 2: tokens do GIOCONDA presentes no nome da lotação
                if not ch:
                    _tokens = [t for t in tfl.split() if len(t) > 2]
                    if len(_tokens) >= 2:
                        for lot_n, lot_fl, lot_ch in _lot_list:
                            if all(tok in lot_n for tok in _tokens) or all(tok in lot_fl for tok in _tokens):
                                ch = lot_ch; break

                # Match 3: tokens da LOTAÇÃO presentes no nome do GIOCONDA (inverso)
                if not ch:
                    for lot_n, lot_fl, lot_ch in _lot_list:
                        lot_tokens = [t for t in lot_fl.split() if len(t) > 2]
                        if len(lot_tokens) >= 2 and all(tok in tn for tok in lot_tokens):
                            ch = lot_ch; break

                if ch:
                    oferta['ch_semanal'] = ch; enr += 1
            print(f"[{ts()}] CH enriquecida: {enr}/{len(dados.get('ger_ofertas',[]))} ofertas")
        except Exception as e:
            print(f"[{ts()}] AVISO: Erro ao processar gerenciamento: {e}")
            import traceback; traceback.print_exc()
            dados['tem_gerenciamento'] = False
    else:
        dados['tem_gerenciamento'] = False
    # ── ALUNOS HUB: matrículas distintas ──────────────────────────────────────
    if p5:
        try:
            alunos_hub = carregar_alunos_hub(p5)
            if alunos_hub:
                dados['alunos_hub'] = alunos_hub
                # Sobrescrever total_alunos_matriculados nos ger_kpis
                if 'ger_kpis' in dados:
                    dados['ger_kpis']['total_alunos_matriculados'] = alunos_hub['total_distintos']
                    dados['ger_kpis']['alunos_mat_fonte'] = 'hub_csv'
                    # Atualizar também DB.kpis.total_alunos com o valor correto do hub
                    if 'kpis' in dados:
                        dados['kpis']['total_alunos'] = alunos_hub['total_distintos']
                    print(f"[{ts()}] KPI alunos substituído: {alunos_hub['total_distintos']:,} (matrículas distintas)")
                # BUG 3 FIX: enriquecer alunos por polo usando hub CSV (por_polo normalizado)
                # Cobre polos com alunos=0 porque TOTAL_ALUNOS está zerado na lotação 2026_2
                import unicodedata as _ud3, re as _re4
                def _norm_polo_hub_main(s):
                    s = _ud3.normalize('NFD', str(s or '').upper().strip())
                    s = ''.join(c for c in s if _ud3.category(c) != 'Mn')
                    s = _re4.sub(r'^LAP\s*[-–]\s*', '', s).strip()
                    return _re4.sub(r'\s+', ' ', s)
                _hub_por_polo = alunos_hub.get('por_polo', {})
                _enr_polo = 0
                for _ps in dados.get('polo_stats', []):
                    if _ps.get('a', _ps.get('alunos', 0)) == 0:
                        _pn = _norm_polo_hub_main(_ps.get('n', _ps.get('polo', _ps.get('POLO', ''))))
                        _al_hub = _hub_por_polo.get(_pn, 0)
                        if _al_hub:
                            _ps['a'] = int(_al_hub)
                            _ps['alunos'] = int(_al_hub)
                            _enr_polo += 1
                print(f"[{ts()}] Polos enriquecidos com alunos (hub CSV): {_enr_polo}")
        except Exception as e:
            print(f"[{ts()}] AVISO: erro ao ler alunos hub: {e}")
    else:
        print(f"[{ts()}] INFO: Relatorio_alunos_por_hub.csv não encontrado — usando contagem GIOCONDA")

    # Preencher alunos_por_curso com hub CSV se ainda vazio (lotação sem TOTAL ALUNOS)
    if not dados.get('alunos_por_curso') and 'alunos_hub' in dir():
        _por_cat = alunos_hub.get('por_cat', {}) if alunos_hub else {}
        _CAT_NOME = {
            'ENF-INS (Multidisciplinar II)':          'Enfermagem e Instrumentação Cirúrgica',
            'BIO-FAR (Multidisciplinar I)':           'Biomedicina e Farmácia',
            'BIO-FISIO-EST-TO (Multidisciplinar III)':'Fisioterapia, T.Ocupacional e Estética',
            'NUTRI (Multidisciplinar IV)':            'Nutrição',
            'ENGMAKER':                               'Engenharias e Licenciaturas',
            'QUÍMICA E FÍSICA':                       'Química e Física',
        }
        if _por_cat:
            dados['alunos_por_curso'] = [
                {'sigla': k, 'curso': _CAT_NOME.get(k, k), 'alunos': int(v)}
                for k, v in sorted(_por_cat.items(), key=lambda x: -x[1])
                if v > 0
            ]
            _tot = sum(x['alunos'] for x in dados['alunos_por_curso'])
            print(f"[{ts()}] Alunos por curso (hub CSV): {len(dados['alunos_por_curso'])} categorias, total {_tot:,}")

    html = gerar_html(dados)
    if '--sem-browser' not in sys.argv:
        print(f"[{ts()}] Abrindo navegador...")
        webbrowser.open(Path(html).as_uri())
    if WATCH_MODE: modo_watch(p1, p2)
    else: print(f"[{ts()}] Concluído!")
