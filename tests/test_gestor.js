// Carrega template_gestor.html duas vezes: uma com a fixture completa (tem
// DB.ocorrencias populado) e outra com uma cópia sem a chave 'ocorrencias'
// (simula p8 ausente, que é o estado real de produção hoje) — confere que
// os dois casos rendem sem erro e que o card degrada certo quando ausente.
const fs = require('fs');
const path = require('path');
const { carregarPortal } = require('./lib/jsdom_portal');

const ROOT = path.join(__dirname, '..');
const TEMPLATE = path.join(ROOT, 'template_gestor.html');
const FIXTURE = process.env.VINCILAB_FIXTURE || path.join(__dirname, 'fixtures', 'db_real_anonimizado.json');

async function main() {
  if (!fs.existsSync(FIXTURE)) {
    console.log('SKIP: tests/fixtures/db_real_anonimizado.json não existe ainda.');
    process.exit(0);
  }
  const dbComOcorrencias = JSON.parse(fs.readFileSync(FIXTURE, 'utf-8'));
  const dbSemOcorrencias = { ...dbComOcorrencias };
  delete dbSemOcorrencias.ocorrencias;

  let ok = true;

  {
    const { errors, document } = await carregarPortal(TEMPLATE, dbComOcorrencias, '_iniciarGestor');
    console.log('=== gestor.html (com ocorrencias) ===');
    console.log('erros:', errors.length);
    errors.forEach((e) => console.log(' -', e));
    console.log('ocor-sub:', document.getElementById('ocor-sub').textContent);
    console.log('lab-sub:', document.getElementById('lab-sub').textContent);
    if (errors.length) ok = false;
  }

  {
    const { errors, document } = await carregarPortal(TEMPLATE, dbSemOcorrencias, '_iniciarGestor');
    console.log('\n=== gestor.html (SEM ocorrencias, simula p8 ausente) ===');
    console.log('erros:', errors.length);
    errors.forEach((e) => console.log(' -', e));
    const ocorSub = document.getElementById('ocor-sub').textContent;
    console.log('ocor-sub:', ocorSub);
    if (errors.length || ocorSub !== 'Sem dados') {
      ok = false;
      console.log('FALHA: esperado "Sem dados" com p8 ausente, veio:', ocorSub);
    }
  }

  process.exit(ok ? 0 : 1);
}
main();
