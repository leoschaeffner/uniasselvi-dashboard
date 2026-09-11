// Carrega template_coordenadores.html com a fixture de referência, navega
// pelas páginas principais e depois aplica CADA filtro de curso disponível
// no painel, conferindo que a contagem de tutores nunca some
// inesperadamente (fica em 0) e que nada lança erro.
const fs = require('fs');
const path = require('path');
const { carregarPortal } = require('./lib/jsdom_portal');

const ROOT = path.join(__dirname, '..');
const TEMPLATE = path.join(ROOT, 'template_coordenadores.html');
const FIXTURE = process.env.VINCILAB_FIXTURE || path.join(__dirname, 'fixtures', 'db_real_anonimizado.json');

const PAGINAS = ['port-visao', 'port-praticas', 'port-polos', 'tutores-mec', 'ger-agendas', 'agendas-estudo', 'ger-detalhe', 'vagas'];

async function main() {
  if (!fs.existsSync(FIXTURE)) {
    console.log('SKIP: tests/fixtures/db_real_anonimizado.json não existe ainda.');
    process.exit(0);
  }
  const db = fs.readFileSync(FIXTURE, 'utf-8');
  const { errors, window, document } = await carregarPortal(TEMPLATE, db, '_iniciarPortalCoordenador', PAGINAS);

  console.log('=== template_coordenadores.html (navegação, sem filtro de curso) ===');
  console.log('erros capturados:', errors.length);
  errors.forEach((e) => console.log(' -', e));

  // Filtros de curso: um de cada vez
  let falhasFiltro = 0;
  const cursos = Object.keys(window.eval('CURSO_PARA_CATEGORIA'));
  console.log(`\n=== Filtros de curso (${cursos.length} cursos) ===`);
  for (const curso of cursos) {
    document.querySelectorAll('.curso-check').forEach((chk) => { chk.checked = chk.value === curso; });
    let throwMsg = null;
    try {
      window.eval('_aplicarCursosSelecionados();');
    } catch (e) {
      throwMsg = e.message;
    }
    const nTutores = window.eval('DB.tutores.length');
    const ok = !throwMsg && nTutores > 0;
    if (!ok) {
      falhasFiltro++;
      console.log(`[${curso}] FALHOU — tutores=${nTutores}${throwMsg ? ' throw=' + throwMsg : ''}`);
    }
  }
  console.log(`Filtros com problema: ${falhasFiltro} de ${cursos.length}`);

  const ok = errors.length === 0 && falhasFiltro === 0;
  process.exit(ok ? 0 : 1);
}
main();
