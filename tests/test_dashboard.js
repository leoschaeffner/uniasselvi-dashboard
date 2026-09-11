// Carrega template_dashboard.html com a fixture de referência, navega pelas
// páginas principais do menu e confere que nada lança erro.
const fs = require('fs');
const path = require('path');
const { carregarPortal } = require('./lib/jsdom_portal');

const ROOT = path.join(__dirname, '..');
const TEMPLATE = path.join(ROOT, 'template_dashboard.html');
const FIXTURE = process.env.VINCILAB_FIXTURE || path.join(__dirname, 'fixtures', 'db_real_anonimizado.json');

const PAGINAS = [
  'port-visao', 'port-ordens', 'port-praticas', 'port-polos', 'port-tutores',
  'tutores-mec', 'ger-visao', 'ger-ofertas', 'ger-agendas', 'agendas-estudo',
  'ger-detalhe', 'vagas', 'insumos',
];

async function main() {
  if (!fs.existsSync(FIXTURE)) {
    console.log('SKIP: tests/fixtures/db_real_anonimizado.json não existe ainda.');
    console.log('Gere com tests/tools/decifrar_payload.js + tests/tools/anonimizar_db.js');
    console.log('a partir de uma saida/dashboard.html rodada localmente, e confira manualmente');
    console.log('por PII residual antes de colocar em tests/fixtures/.');
    process.exit(0);
  }
  const db = fs.readFileSync(FIXTURE, 'utf-8');
  const { errors } = await carregarPortal(TEMPLATE, db, '_iniciarDashboard', PAGINAS);

  console.log('=== template_dashboard.html ===');
  console.log('erros capturados:', errors.length);
  errors.forEach((e) => console.log(' -', e));

  process.exit(errors.length ? 1 : 0);
}
main();
