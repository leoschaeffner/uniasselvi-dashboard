// Roda toda a suíte de regressão de frontend em sequência e falha se
// qualquer um dos scripts falhar. Uso: node tests/run_all.js
const { execFileSync } = require('child_process');
const path = require('path');

const scripts = ['test_dashboard.js', 'test_coordenadores.js', 'test_gestor.js'];
let falhou = false;

for (const s of scripts) {
  console.log(`\n########## ${s} ##########`);
  try {
    execFileSync(process.execPath, [path.join(__dirname, s)], { stdio: 'inherit' });
  } catch (e) {
    falhou = true;
  }
}

process.exit(falhou ? 1 : 0);
