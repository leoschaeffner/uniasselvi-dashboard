// Decifra o payload AES-256-GCM embutido num HTML gerado pelo processar.py
// (saida/dashboard.html, saida/coordenadores.html ou saida/gestor.html).
// Não contém nenhum dado real — é só a mesma lógica de decifragem que já
// está pública no próprio JS do cliente (_decifrarDados em template_*.html).
//
// Uso: node decifrar_payload.js <caminho_html> <senha> <saida_json>
//   senha: uniasselvi2026 (dashboard/coordenadores) | vincilab_gestor_2026 (gestor)
//
// IMPORTANTE: a saída deste script contém dado real (PII de tutores) se
// rodado contra uma saida/*.html gerada com planilhas reais. NUNCA commitar
// o resultado direto — passe por um anonimizador antes de qualquer coisa
// entrar em tests/fixtures/.
const fs = require('fs');
const { webcrypto } = require('crypto');

async function main() {
  const [, , htmlPath, senha, outPath] = process.argv;
  if (!htmlPath || !senha || !outPath) {
    console.error('Uso: node decifrar_payload.js <html> <senha> <saida.json>');
    process.exit(1);
  }
  const html = fs.readFileSync(htmlPath, 'utf-8');
  const marker = 'const ENCRYPTED_PAYLOAD = "';
  const start = html.indexOf(marker);
  if (start === -1) { console.error('Marker ENCRYPTED_PAYLOAD não encontrado em', htmlPath); process.exit(1); }
  const contentStart = start + marker.length;
  const end = html.indexOf('";', contentStart);
  if (end === -1) { console.error('Fim do payload não encontrado'); process.exit(1); }
  const payload = html.slice(contentStart, end); // "iv_b64:ct_b64", sem escapes especiais
  const [ivB64, ctB64] = payload.split(':');
  const iv = Uint8Array.from(Buffer.from(ivB64, 'base64'));
  const ct = Uint8Array.from(Buffer.from(ctB64, 'base64'));
  const chaveBuf = await webcrypto.subtle.digest('SHA-256', new TextEncoder().encode(senha));
  const cryptoKey = await webcrypto.subtle.importKey('raw', chaveBuf, { name: 'AES-GCM' }, false, ['decrypt']);
  const plainBuf = await webcrypto.subtle.decrypt({ name: 'AES-GCM', iv }, cryptoKey, ct);
  const json = new TextDecoder().decode(plainBuf);
  fs.writeFileSync(outPath, json, 'utf-8');
  console.log('OK ->', outPath, '(' + (json.length / 1024 / 1024).toFixed(1) + ' MB)');
}
main().catch((e) => { console.error('ERRO', e); process.exit(1); });
