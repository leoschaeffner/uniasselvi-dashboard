"""Testes de processar_remessas / _remessas_demo / montar_dados_lab (PATCH 178).
Sem framework: python tests/test_remessas.py  (exit != 0 se algo falhar)."""
import os, sys, json, tempfile, shutil
from datetime import date, datetime

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
import openpyxl
import processar as P

HOJE = date(2026, 9, 25)
falhas = []


def check(cond, msg):
    if not cond:
        falhas.append(msg)
        print("  FALHOU:", msg)


TMP = tempfile.mkdtemp()


def xlsx(nome, abas):
    """abas: {nome_aba: [linhas]} (1ª linha = cabeçalho)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for n, linhas in abas.items():
        ws = wb.create_sheet(n)
        for l in linhas:
            ws.append(l)
    p = os.path.join(TMP, nome)
    wb.save(p)
    return p


TRANSP = P._carregar_transportadoras()
check(any(t['id'] == 'correios' and t['url'] for t in TRANSP), "transportadoras_rastreio.json sem Correios com url")

# ── 1. planilha principal: colunas embaralhadas, acentos, PII presente ──────
cab = ['Destinatário', 'Código de Rastreio', 'Situação', 'Previsão de Entrega', 'Data de Envio', 'Polo',
       'Telefone', 'Data de Entrega', 'Transportadora', 'Categoria de Laboratório', 'Kit', 'Qtd',
       'E-mail', 'Endereço', 'Observação']
def L(**k):
    return [k.get(c, None) for c in ['dest', 'cod', 'st', 'prev', 'env', 'polo', 'tel', 'ent', 'transp', 'cat', 'kit', 'qtd', 'mail', 'end', 'obs']]
PII = dict(dest='Maria Segredo', tel='(48) 99999-0000', mail='segredo@x.com', end='Rua Secreta 1')
linhas = [cab,
    L(polo='Blumenau/SC — João da Silva', cat='enfermagem', kit='Kit A', qtd=2, env=datetime(2026, 9, 10),
      prev=datetime(2026, 9, 20), transp='sedex', cod=' ab 123456789 br ', **PII),                     # r1
    L(polo='Curitiba/PR - X', cat='EngeMaker', kit='Kit B', qtd='3', env=datetime(2026, 9, 1), ent='05/09/2026',
      prev=datetime(2026, 9, 4), transp='Correios', cod='ZZ111111111BR', **PII),                        # r2
    L(polo='Recife/PE', env='01/12/2026', transp='Jadlog', cod=123456, st='em transito',
      prev='30/09/2026'),                                                                              # r3
    L(polo='Natal/RN', transp='Transp XYZ', cod='xx1', env=datetime(2026, 9, 20), prev=datetime(2026, 9, 30)),  # r4
    L(),                                                                                               # linha vazia
    L(polo='Lima/PE', st='Preparando', cat='Terapia Ocupacional', kit='Kit C', qtd=1),                 # r6
    L(polo='Exemplo/XX', obs='EXEMPLO — apagar', kit='Kit'),                                           # exemplo
    L(polo='Sul/SC', st='Problema', env=datetime(2026, 9, 12), prev=datetime(2026, 9, 24), transp='Correios',
      cod='PP222222222BR', obs='extraviado'),                                                          # r8 atrasado
    L(polo='Dev/SC', st='Devolvido', env=datetime(2026, 8, 1), prev=datetime(2026, 8, 10)),            # r9 devolvido nao atrasa
    L(polo='Amb/SC', env=datetime(2026, 9, 20), prev='10/01/2026', transp='Correios', cod='QQ333333333BR'),  # r10 previsao ambigua -> US
    L(polo='Amb2/SC', env=datetime(2026, 9, 20), prev='01/10/2026'),                                   # r11 previsao ambigua -> BR
    L(polo='Amb3/SC', env='09/25/2026'),                                                               # r12 texto so-US valido
]
p = xlsx('rem.xlsx', {'Registro': linhas, 'Listas': [['STATUS'], ['Enviado']]})
res = P.processar_remessas(p, hoje=HOJE, transportadoras=TRANSP)
R = {r['polo']: r for r in res['registros']}
check(len(res['registros']) == 10, f"esperava 10 registros (vazia e exemplo ignoradas), veio {len(res['registros'])}")
check('Exemplo/XX' not in R, "linha EXEMPLO não deveria entrar")

r1 = R['Blumenau/SC']
check(r1['status'] == 'Enviado' and r1['status_derivado'], "r1 status derivado Enviado")
check(r1['codigo_rastreio'] == 'AB123456789BR' and r1['tem_codigo'], "r1 código normalizado")
check(r1['atrasado'] and r1['dias_atraso'] == 5 and r1['dias_desde_envio'] == 15, f"r1 atraso/dias: {r1['dias_atraso']}/{r1['dias_desde_envio']}")
check(r1['transportadora_id'] == 'correios' and r1['link_rastreio'] and 'AB123456789BR' in r1['link_rastreio'], "r1 link Correios")
check(r1['codigo_formato_ok'] is True, "r1 formato Correios ok")
check(r1['categoria'] == 'Enfermagem' and r1['quantidade'] == 2 and r1['item'] == 'Kit A', "r1 categoria canônica/qtd/item")
check(r1['data_envio'] == '2026-09-10' and r1['previsao_entrega'] == '2026-09-20', "r1 datas ISO")
check('tutor' not in r1 and 'João' not in json.dumps(res, ensure_ascii=False), "tutor do polo combinado descartado")

r2 = R['Curitiba/PR - X']
check(r2['status'] == 'Entregue' and r2['data_entrega'] == '2026-09-05' and not r2['atrasado'], "r2 entregue derivado, não atrasado")
check(r2['quantidade'] == 3 and r2['dias_ate_entrega'] == 4, "r2 qtd texto e dias_ate_entrega")

r3 = R['Recife/PE']
check(r3['status'] == 'Em trânsito' and not r3['status_derivado'], "r3 status texto sem acento")
check(r3['data_envio'] == '2026-01-12', f"r3 envio 01/12/2026 (BR=futuro) deve virar 12/jan (US): {r3['data_envio']}")
check(r3['codigo_rastreio'] == '123456', "r3 código numérico sem .0")
check(r3['transportadora_id'] == 'jadlog' and r3['link_rastreio'] is None, "r3 Jadlog casada mas sem URL -> sem link")

r4 = R['Natal/RN']
check(r4['transportadora_id'] is None and r4['link_rastreio'] is None and r4['tem_codigo'], "r4 transportadora desconhecida sem link")
r6 = R['Lima/PE']
check(r6['status'] == 'Preparando' and not r6['tem_codigo'] and not r6['atrasado'] and r6['categoria'] == 'Terapia Ocupacional', "r6 preparando")
r8 = R['Sul/SC']
check(r8['status'] == 'Problema' and r8['atrasado'] and r8['dias_atraso'] == 1, "r8 problema atrasado")
check(not R['Dev/SC']['atrasado'], "devolvido nunca atrasado")
check(R['Amb/SC']['previsao_entrega'] == '2026-10-01', f"previsão ambígua -> leitura na janela (US): {R['Amb/SC']['previsao_entrega']}")
check(R['Amb2/SC']['previsao_entrega'] == '2026-10-01', f"previsão ambígua -> BR: {R['Amb2/SC']['previsao_entrega']}")
check(R['Amb3/SC']['data_envio'] == '2026-09-25', "texto MM/DD só válido como US")

k = res['kpis']
check(k['total'] == 10 and k['preparando'] == 1 and k['entregues'] == 1 and k['problemas'] == 1 and k['devolvidos'] == 1
      and k['em_transito'] == 1, f"kpis por status: {k}")
check(k['enviado'] == 5, f"kpis.enviado (r1,r4,r10,r11,r12) = {k['enviado']}")
check(k['atrasados'] == 2, f"atrasados (r1,r8) = {k['atrasados']}")
check(k['sem_codigo'] == 2, f"sem_codigo (despachadas sem código: r11,r12) = {k['sem_codigo']}")
check(k['polos_atendidos'] == 10 and k['ultima_atualizacao'] == '2026-09-25', f"polos/última: {k['polos_atendidos']}/{k['ultima_atualizacao']}")
check([s['status'] for s in res['por_status']] == P._REM_STATUS, "por_status com os 6 status em ordem")
check(sum(c['total'] for c in res['por_categoria']) == 10, "por_categoria soma")
check(any(t['transportadora'] == 'Correios' for t in res['por_transportadora']), "por_transportadora nome canônico")
check(all('mes' in m for m in res['por_mes_envio']) and sum(m['total'] for m in res['por_mes_envio']) == 9, "por_mes_envio")

blob = json.dumps(res, ensure_ascii=False)
for pii in ['Maria Segredo', '99999-0000', 'segredo@x.com', 'Rua Secreta']:
    check(pii not in blob, f"PII vazou no JSON: {pii}")
for chave in ('destinatario', 'telefone', 'email', 'endereco'):
    check(chave not in blob.lower(), f"chave de PII no JSON: {chave}")

# ── 2. só cabeçalho / aba ausente / arquivo inválido ────────────────────────
p2 = xlsx('so_cab.xlsx', {'Registro': [cab]})
r = P.processar_remessas(p2, hoje=HOJE, transportadoras=TRANSP)
check(r is not None and r['kpis']['total'] == 0 and r['registros'] == [] and r['kpis']['ultima_atualizacao'] is None, "só cabeçalho -> total 0")
p3 = xlsx('outra_aba.xlsx', {'Planilha1': [['Polo', 'Status'], ['Zeta/SC', 'Entregue']], 'Listas': [['x']]})
r = P.processar_remessas(p3, hoje=HOJE, transportadoras=TRANSP)
check(r['kpis']['total'] == 1 and r['registros'][0]['status'] == 'Entregue', "aba 'Registro' ausente -> 1ª aba")
p4 = xlsx('sem_colunas.xlsx', {'Registro': [['Foo', 'Bar'], [1, 2]]})
r = P.processar_remessas(p4, hoje=HOJE, transportadoras=TRANSP)
check(r['kpis']['total'] == 0, "colunas irreconhecíveis -> nenhuma linha (sem quebrar)")
bad = os.path.join(TMP, 'lixo.xlsx')
open(bad, 'wb').write(b'nao e xlsx')
check(P.processar_remessas(bad, hoje=HOJE, transportadoras=TRANSP) is None, "arquivo inválido -> None")
p5 = xlsx('prev_vs_entrega.xlsx', {'Registro': [['Polo', 'Previsão de Entrega', 'Data de Entrega'],
                                                ['P/SC', datetime(2026, 9, 30), datetime(2026, 9, 12)]]})
r = P.processar_remessas(p5, hoje=HOJE, transportadoras=TRANSP)['registros'][0]
check(r['previsao_entrega'] == '2026-09-30' and r['data_entrega'] == '2026-09-12',
      "'Previsão de Entrega' não pode ser confundida com 'Data de Entrega'")

# ── 3. demo determinística ─────────────────────────────────────────────────
d1, d2 = P._remessas_demo(HOJE), P._remessas_demo(HOJE)
check(json.dumps(d1, sort_keys=True) == json.dumps(d2, sort_keys=True), "demo determinística")
check(len(d1['registros']) == 80, f"demo ~80 registros: {len(d1['registros'])}")
check({r['status'] for r in d1['registros']} == set(P._REM_STATUS), "demo com todos os status")
check(all(r['demo'] is True for r in d1['registros']), "todo registro demo:true")
check(all(r['codigo_rastreio'].startswith('DEMO') or not r['codigo_rastreio'] for r in d1['registros']), "códigos fictícios")
pct = d1['kpis']['atrasados'] / len(d1['registros'])
check(0.05 <= pct <= 0.15, f"~10% atrasadas: {pct:.2%}")
check({'Correios', 'Jadlog', 'Total Express'} <= {t['transportadora'] for t in d1['por_transportadora']}, "3 transportadoras na demo")
check(any(r['link_rastreio'] for r in d1['registros']) and any(r['transportadora_id'] and not r['link_rastreio'] and r['tem_codigo'] for r in d1['registros']),
      "demo com e sem link")
d3 = P._remessas_demo(date(2026, 12, 1))
check(d3['kpis']['atrasados'] == d1['kpis']['atrasados'], "datas relativas a hoje: mesmo nº de atrasadas em outra data")

# ── 4. montar_dados_lab / gerar_html_laboratorios ──────────────────────────
dl = P.montar_dados_lab({'gerado_em': 'x', 'laboratorios': {'vistorias': None}}, None, HOJE)
check(set(dl) == {'gerado_em', 'demo', 'pendencias', 'vistorias', 'remessas', 'vinculo'}, f"chaves dados_lab: {sorted(dl)}")
check(dl['demo'] == {'vinculo': True, 'remessas': True}, "demo flags sem p10")
check(len(dl['pendencias']) == 51 and len(dl['pendencias'][0]) == 8, "51 pendências com 8 chaves")
check(dl['vinculo']['demo'] is True and dl['vinculo']['recorrencia_real'] is False, "vinculo marcado demo")
dl2 = P.montar_dados_lab({}, res, HOJE)
check(dl2['demo']['remessas'] is False and dl2['remessas'] is res, "com p10 real: demo.remessas False")
blob = json.dumps(dl, ensure_ascii=False).lower()
check('"tutores"' not in blob and '@' not in blob, "dados_lab sem tutores/e-mail")

orig = P.SCRIPT_DIR
try:
    P.SCRIPT_DIR = TMP
    P.gerar_html_laboratorios(dl)
    check(not os.path.exists(os.path.join(TMP, 'saida', 'laboratorios.html')), "sem template -> pula sem gerar")
    open(os.path.join(TMP, 'template_laboratorios.html'), 'w', encoding='utf-8').write("<script>const D='DATA_GOES_HERE';</script>")
    P.gerar_html_laboratorios(dl)
    h = open(os.path.join(TMP, 'saida', 'laboratorios.html'), encoding='utf-8').read()
    check('DATA_GOES_HERE' not in h and ':' in h, "template preenchido com payload cifrado")
finally:
    P.SCRIPT_DIR = orig

shutil.rmtree(TMP, ignore_errors=True)
if falhas:
    print(f"\n{len(falhas)} FALHA(S)")
    sys.exit(1)
print("test_remessas: OK")
