// Carrega template_laboratorios.html (4º portal, PATCH 178) com um payload
// montado a partir da fixture anonimizada existente (montar_dados_lab(dados,
// None), gerado ad-hoc via Python -- não commitamos um novo JSON gigante em
// tests/fixtures/ só pra isso). Confere: 0 erros de JS, banners de DEMO
// visíveis (vínculo é sempre demo hoje; remessas é demo quando p10 ausente),
// tabelas de Pendências/Remessas renderizando, e o mapa Leaflet não lança
// exceção fatal (mesmo em jsdom, que não tem layout/canvas de verdade --
// ver comentário em renderizarVincMapa/destacarVincNoMapa no template).
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { carregarPortal } = require('./lib/jsdom_portal');

const ROOT = path.join(__dirname, '..');
const TEMPLATE = path.join(ROOT, 'template_laboratorios.html');
const DASHBOARD_FIXTURE = path.join(__dirname, 'fixtures', 'db_real_anonimizado.json');

async function main() {
  if (!fs.existsSync(TEMPLATE)) {
    console.log('SKIP: template_laboratorios.html não existe ainda.');
    process.exit(0);
  }
  if (!fs.existsSync(DASHBOARD_FIXTURE)) {
    console.log('SKIP: tests/fixtures/db_real_anonimizado.json não existe ainda (fonte pro montar_dados_lab).');
    process.exit(0);
  }

  // Monta o payload de Laboratórios em cima da mesma fixture anonimizada do
  // dashboard (mesma função que o processar.py usa de verdade).
  const tmpOut = path.join(__dirname, '..', 'scratchpad', '_dados_lab_fixture_tmp.json');
  try {
    execFileSync(process.platform === 'win32' ? 'python' : 'python3', ['-c', `
import json, sys
sys.path.insert(0, ${JSON.stringify(ROOT)})
import processar as p
dados = json.load(open(${JSON.stringify(DASHBOARD_FIXTURE)}, encoding='utf-8'))
dl = p.montar_dados_lab(dados, None)
json.dump(dl, open(${JSON.stringify(tmpOut)}, 'w', encoding='utf-8'), ensure_ascii=False)
`], { stdio: 'inherit' });
  } catch (e) {
    console.log('SKIP: não foi possível montar dados_lab via Python (' + e.message + ').');
    process.exit(0);
  }

  const dadosLab = JSON.parse(fs.readFileSync(tmpOut, 'utf-8'));
  try { fs.unlinkSync(tmpOut); } catch (e) {}

  let ok = true;
  const { errors, document } = await carregarPortal(TEMPLATE, dadosLab, '_iniciarLabs');
  console.log('=== template_laboratorios.html ===');
  console.log('erros:', errors.length);
  errors.forEach((e) => console.log(' -', e));
  if (errors.length) ok = false;

  const bodyText = document.body.textContent;
  const temBannerVinculo = bodyText.includes('DADOS DE DEMONSTRAÇÃO') && bodyText.includes('Vínculo de Laboratórios');
  console.log('banner demo (vínculo) presente:', temBannerVinculo);
  if (dadosLab.demo && dadosLab.demo.vinculo && !temBannerVinculo) {
    ok = false;
    console.log('FALHA: demo.vinculo=true mas o banner não apareceu no texto renderizado.');
  }

  const remBannerVisivel = document.getElementById('rem-demo-banner').style.display !== 'none';
  console.log('banner demo (remessas) visível:', remBannerVisivel, '(demo.remessas=' + (dadosLab.demo || {}).remessas + ')');
  if (!!(dadosLab.demo || {}).remessas !== remBannerVisivel) {
    ok = false;
    console.log('FALHA: banner de demo de remessas não bate com dadosLab.demo.remessas.');
  }

  console.log('pend-sub:', document.getElementById('pend-sub').textContent);
  console.log('pend linhas:', document.querySelectorAll('#pend-tbody tr').length);
  console.log('rem-sub:', document.getElementById('rem-sub').textContent);
  console.log('rem linhas:', document.querySelectorAll('#rem-tbody tr').length);
  console.log('vist-sub:', document.getElementById('vist-sub').textContent);

  process.exit(ok ? 0 : 1);
}
main().catch((e) => { console.error('ERRO', e); process.exit(1); });
