// Harness reutilizável pra testar os três portais (dashboard/coordenadores/
// gestor) com jsdom: carrega o template, injeta uma DB (JSON puro, já
// decifrada) direto na variável de topo `DB`, chama a função de bootstrap
// do portal, navega pelas páginas indicadas e reporta qualquer erro
// (exceção lançada ou console.error) capturado durante o processo.
//
// Não decifra nada aqui de propósito — quem chama passa o DB já como JSON.
// Pra decifrar um HTML gerado localmente (dashboard.html/gestor.html) use
// tests/tools/decifrar_payload.js.
const fs = require('fs');
const { JSDOM, VirtualConsole } = require('jsdom');

/**
 * @param {string} templatePath   caminho do template_*.html
 * @param {object|string} db      objeto DB ou string JSON já serializada
 * @param {string} bootstrapFn    nome da função global que inicializa o portal
 *                                (ex: '_iniciarDashboard', '_iniciarPortalCoordenador', '_iniciarGestor')
 * @param {string[]} paginas      nomes de página pra passar em navTo(), na ordem
 * @returns {Promise<{errors:string[], window: any, document: any}>}
 */
async function carregarPortal(templatePath, db, bootstrapFn, paginas = []) {
  const html = fs.readFileSync(templatePath, 'utf-8');
  const dbJson = typeof db === 'string' ? db : JSON.stringify(db);
  const errors = [];
  const vc = new VirtualConsole();
  vc.on('jsdomError', (e) => errors.push('[jsdomError] ' + e.message));
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    resources: 'usable',
    url: 'http://localhost/',
    virtualConsole: vc,
  });
  const { window } = dom;
  window.console.error = (...a) => errors.push('[console.error] ' + a.map(String).join(' ').slice(0, 400));
  await new Promise((r) => setTimeout(r, 250));

  window.eval('DB = ' + dbJson + ';');
  try {
    window.eval(bootstrapFn + '();');
  } catch (e) {
    errors.push('[bootstrap throw] ' + e.stack);
  }
  await new Promise((r) => setTimeout(r, 150));

  for (const pg of paginas) {
    try {
      window.eval(`navTo('${pg}');`);
    } catch (e) {
      errors.push(`[navTo(${pg}) throw] ` + e.stack);
    }
  }
  await new Promise((r) => setTimeout(r, 150));

  return { errors, window, document: window.document };
}

module.exports = { carregarPortal };
