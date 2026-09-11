// Anonimiza uma DB decifrada (JSON, ver decifrar_payload.js) pra virar
// fixture de teste versionável: troca nomes, e-mails, whatsapp, chapa e
// lattes por valores sintéticos consistentes (o mesmo valor original sempre
// vira o mesmo valor sintético, então cruzamentos entre coleções continuam
// batendo — ex: mesmo tutor em `tutores[]` e em `ger_ofertas[].tutor`).
// Também trunca coleções gigantes e repetitivas (ger_ofertas /
// gerenciamento_por_semestre) pra um tamanho razoável de fixture em git.
//
// Uso: node anonimizar_db.js <entrada.json> <saida.json> [limiteArraysGrandes=2000]
//
// Depois de gerar, sempre confira manualmente por PII residual antes de
// commitar (grep por domínios de e-mail reais, nomes conhecidos etc.) — este
// script cobre os campos conhecidos, mas não é uma garantia formal de
// anonimização completa.
const fs = require('fs');

const nameMap = new Map();
let nameCounter = 0;
function anonName(v) {
  if (!v || typeof v !== 'string' || !v.trim()) return v;
  if (!nameMap.has(v)) {
    nameCounter++;
    nameMap.set(v, `Tutor Sintetico ${String(nameCounter).padStart(4, '0')}`);
  }
  return nameMap.get(v);
}

const emailMap = new Map();
let emailCounter = 0;
function anonEmail(v) {
  if (!v || typeof v !== 'string' || !v.trim()) return v;
  if (!emailMap.has(v)) {
    emailCounter++;
    emailMap.set(v, `tutor.sintetico.${emailCounter}@exemplo.test`);
  }
  return emailMap.get(v);
}

const phoneMap = new Map();
let phoneCounter = 0;
function anonPhone(v) {
  if (!v || typeof v !== 'string' || !v.trim()) return v;
  if (!phoneMap.has(v)) {
    phoneCounter++;
    phoneMap.set(v, `+55 99 9${String(phoneCounter).padStart(4, '0')}-0000`);
  }
  return phoneMap.get(v);
}

const chapaMap = new Map();
let chapaCounter = 0;
function anonChapa(v) {
  if (v === null || v === undefined || v === '') return v;
  const key = String(v);
  if (!chapaMap.has(key)) {
    chapaCounter++;
    chapaMap.set(key, String(90000 + chapaCounter));
  }
  return chapaMap.get(key);
}

const NAME_KEYS = new Set(['n', 'nome', 'tutor', 'tutor_atual', 'multiplicador', 'responsavel', 'responsavel_tratativa', 'nome_tutor', 'nome_subm']);
const EMAIL_KEYS = new Set(['email', 'aviso_email', 'email_subm']);
const PHONE_KEYS = new Set(['whatsapp']);
const CHAPA_KEYS = new Set(['chapa']);
const NULL_KEYS = new Set(['lattes_url', 'lattes_id']); // identificador pessoal direto — não precisa pro teste

function walk(node) {
  if (Array.isArray(node)) {
    for (let i = 0; i < node.length; i++) node[i] = walk(node[i]);
    return node;
  }
  if (node && typeof node === 'object') {
    for (const k of Object.keys(node)) {
      const v = node[k];
      if (NULL_KEYS.has(k)) { node[k] = null; continue; }
      if (NAME_KEYS.has(k) && typeof v === 'string') { node[k] = anonName(v); continue; }
      if (EMAIL_KEYS.has(k) && typeof v === 'string') { node[k] = anonEmail(v); continue; }
      if (PHONE_KEYS.has(k) && typeof v === 'string') { node[k] = anonPhone(v); continue; }
      if (CHAPA_KEYS.has(k)) { node[k] = anonChapa(v); continue; }
      node[k] = walk(v);
    }
    return node;
  }
  return node;
}

function truncar(db, limite) {
  if (Array.isArray(db.ger_ofertas) && db.ger_ofertas.length > limite) {
    db.ger_ofertas = db.ger_ofertas.slice(0, limite);
  }
  if (db.gerenciamento_por_semestre && typeof db.gerenciamento_por_semestre === 'object') {
    for (const sem of Object.keys(db.gerenciamento_por_semestre)) {
      const bloco = db.gerenciamento_por_semestre[sem];
      if (bloco && Array.isArray(bloco.ger_ofertas) && bloco.ger_ofertas.length > limite) {
        bloco.ger_ofertas = bloco.ger_ofertas.slice(0, limite);
      }
    }
  }
  return db;
}

function main() {
  const [, , inPath, outPath, limiteArg] = process.argv;
  if (!inPath || !outPath) {
    console.error('Uso: node anonimizar_db.js <entrada.json> <saida.json> [limite]');
    process.exit(1);
  }
  const limite = limiteArg ? parseInt(limiteArg, 10) : 2000;
  const db = JSON.parse(fs.readFileSync(inPath, 'utf-8'));
  truncar(db, limite);
  walk(db);
  let jsonStr = JSON.stringify(db);

  // Varredura final por e-mail: pega qualquer e-mail real que tenha escapado
  // do walk por key-name (texto livre, chave não mapeada etc.). Preserva os
  // já anonimizados (@exemplo.test).
  const emailRe = /[A-Za-z0-9._%+-]+@(?!exemplo\.test)[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
  let strayEmails = 0;
  jsonStr = jsonStr.replace(emailRe, (m) => { strayEmails++; return anonEmail(m); });

  // Varredura final por nome: o mesmo tutor aparece em MUITOS lugares fora
  // das NAME_KEYS conhecidas — listas soltas de nome (`tutores_unicos`,
  // `ger_contratacao[].tutores`), dict data->[nomes] (`datas_por_tutor`), e
  // até string combinada ("Fulano — 19:00 - 20:30" em `datas_por_horario`).
  // Em vez de tentar mapear key por key (frágil, sempre aparece um lugar
  // novo), troca pelo texto inteiro todo nome real já descoberto no walk()
  // por key conhecida (nameMap) — mais comprido primeiro, pra um nome nunca
  // comer parte de um nome mais longo que o contém.
  let strayNames = 0;
  const nomesReais = [...nameMap.keys()].sort((a, b) => b.length - a.length);
  for (const nomeReal of nomesReais) {
    const sintetico = nameMap.get(nomeReal);
    // nome real pode ter aspas/backslash escapados dentro do JSON serializado
    const escapado = JSON.stringify(nomeReal).slice(1, -1);
    if (escapado !== nomeReal && jsonStr.includes(escapado)) {
      const re = new RegExp(escapado.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      const n = (jsonStr.match(re) || []).length;
      if (n) { strayNames += n; jsonStr = jsonStr.replace(re, sintetico); }
    }
    if (jsonStr.includes(nomeReal)) {
      const re = new RegExp(nomeReal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      const n = (jsonStr.match(re) || []).length;
      if (n) { strayNames += n; jsonStr = jsonStr.replace(re, sintetico); }
    }
  }

  fs.writeFileSync(outPath, jsonStr, 'utf-8');
  const stat = fs.statSync(outPath);
  console.log('OK ->', outPath, (stat.size / 1024 / 1024).toFixed(1) + ' MB');
  console.log('nomes:', nameCounter, '| emails por chave:', emailCounter - strayEmails, '| emails via varredura de texto:', strayEmails, '| whatsapp:', phoneCounter, '| chapas:', chapaCounter, '| ocorrências de nome fora das NAME_KEYS (varredura de texto):', strayNames);
  console.log('Confira manualmente por PII residual antes de commitar (grep por domínios reais, nomes etc).');
}
main();
