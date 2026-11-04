'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadDashboard, demoPassword } = require('../demo/runtime');
const { createRuntime } = require('../demo/gas_mock');
const path = require('node:path');

const NOW = '2026-10-30T15:00:00-03:00';
const plain = (x) => JSON.parse(JSON.stringify(x));
const DEVICE = 'device-token-0001';

function fresh() { return loadDashboard({ now: NOW }); }
function login(rt, user, device = DEVICE) {
  const res = rt.context.verificarLogin(user, demoPassword(user), device);
  assert.equal(res.ok, true, 'login should succeed for ' + user);
  return res.token;
}

test('convertMonetaryToFloat_ understands Brazilian and US formats', () => {
  const rt = fresh();
  const f = (v) => rt.context.convertMonetaryToFloat_(v);
  assert.equal(f('R$ 1.234,56'), 1234.56);
  assert.equal(f('1,234.56'), 1234.56);
  assert.equal(f('1500'), 1500);
  assert.equal(f(12.5), 12.5);
  assert.ok(Number.isNaN(f('')));
  assert.ok(Number.isNaN(f(null)));
});

test('telefonePadrao10_ makes DDI and ninth-digit variants comparable', () => {
  const rt = fresh();
  const f = (v) => rt.context.telefonePadrao10_(v);
  const base10 = '0000000001';
  assert.equal(f('(00) 0000-0001'), base10);
  assert.equal(f('(00) 90000-0001'), base10);
  assert.equal(f('00' + '00' + '0000' + '0001'), base10);
  assert.equal(f('00' + '00' + '90000' + '0001'), base10);
});

test('parseDataHoraProposta_ accepts ISO text, BR text, Date objects and day fractions', () => {
  const rt = fresh();
  const p = (d, h) => rt.context.parseDataHoraProposta_(d, h);
  const iso = p('2026-10-29 00:00:00', '14:35:10');
  assert.deepEqual([iso.getFullYear(), iso.getMonth(), iso.getDate(), iso.getHours(), iso.getMinutes(), iso.getSeconds()],
    [2026, 9, 29, 14, 35, 10]);
  const br = p('29/10/2026', '09:05');
  assert.deepEqual([br.getDate(), br.getMonth(), br.getHours(), br.getMinutes()], [29, 9, 9, 5]);
  assert.equal(p('29/10/2026', 0.5).getHours(), 12);
  assert.equal(p('not a date', '10:00'), null);
});

test('Brazilian business-day calendar for a month with national holidays', () => {
  const rt = fresh();
  assert.ok(rt.context.feriadosNacionaisMeta_(2026).has('2026-04-03'), 'Good Friday 2026');
  const dias = plain(rt.context.diasUteisMeta_('2026-11'));
  assert.equal(dias.total, 19);
  assert.throws(() => rt.context.diasUteisMeta_('2026-13'), /inválida/);
});

test('cache helpers split payloads above the 100KB limit and restore them', () => {
  const rt = fresh();
  const cache = rt.context.CacheService.getScriptCache();
  const big = { rows: Array.from({ length: 12000 }, (_, i) => 'row-' + i + '-xxxxxxxxxxxxxxxx') };
  assert.ok(JSON.stringify(big).length > 200 * 1024);
  rt.context.cachePutJson_(cache, 'k', big, 60);
  assert.deepEqual(plain(rt.context.cacheGetJson_(cache, 'k')), big);
  assert.equal(rt.context.cacheGetJson_(cache, 'missing'), null);
});

test('login: right password works, wrong one gets the same generic message', () => {
  const rt = fresh();
  assert.equal(rt.context.verificarLogin('admin.demo', demoPassword('admin.demo'), DEVICE).papel, 'Admin');
  const wrongPwd = plain(rt.context.verificarLogin('admin.demo', 'nope-nope', DEVICE));
  const wrongUser = plain(rt.context.verificarLogin('ghost.user', 'nope-nope', DEVICE));
  assert.equal(wrongPwd.ok, false);
  assert.equal(wrongPwd.mensagem, wrongUser.mensagem, 'no user enumeration');
});

test('login input validation rejects emoji, markup characters and short secrets', () => {
  const rt = fresh();
  const v = (l, s) => rt.context.entradaValida_(l, s);
  assert.equal(v('good.login', 'secret1'), true);
  assert.equal(v('bad login', 'secret1'), false);
  assert.equal(v('good.login', 'ab'), false);
  assert.equal(v('good.login', 'sec<ret'), false);
  assert.equal(v('good.login', 'secret\u{1F600}'), false);
  assert.equal(v(null, 'secret1'), false);
});

test('progressive lockout: 5 failures lock for 2 minutes, then the clock releases it', () => {
  const rt = fresh();
  for (let i = 0; i < 4; i++) assert.equal(rt.context.verificarLogin('operator.demo', 'wrong-pass', DEVICE).bloqueado, undefined);
  const fifth = plain(rt.context.verificarLogin('operator.demo', 'wrong-pass', DEVICE));
  assert.equal(fifth.bloqueado, true);
  assert.equal(rt.context.verificarLogin('operator.demo', demoPassword('operator.demo'), DEVICE).bloqueado, true);
  rt.setNow(new Date(Date.parse(NOW) + 3 * 60 * 1000));
  assert.equal(rt.context.verificarLogin('operator.demo', demoPassword('operator.demo'), DEVICE).ok, true);
  assert.equal(rt.evaluate('duracaoBloqueioMinutos_(20)'), 48 * 60);
  assert.equal(rt.evaluate('duracaoBloqueioMinutos_(4)'), 0);
});

test('session token: tampering and expiry are rejected, roles come from the signed token', () => {
  const rt = fresh();
  const token = login(rt, 'operator.demo');
  assert.equal(rt.context.validarTokenSessao(token).ok, true);
  const [payload, sig] = token.split('.');
  const forged = Buffer.from('operator.demo|Admin|9999999999999').toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  assert.equal(rt.context.validarTokenSessao(forged + '.' + sig).ok, false);
  assert.equal(rt.context.validarTokenSessao(payload).ok, false);
  rt.setNow(new Date(Date.parse(NOW) + 9 * 3600 * 1000));
  assert.equal(plain(rt.context.validarTokenSessao(token)).expirado, true);
});

test('admin-only and master-only endpoints enforce the role', () => {
  const rt = fresh();
  const op = login(rt, 'operator.demo');
  const admin = login(rt, 'admin.demo');
  const master = login(rt, 'Master');
  assert.throws(() => rt.context.getGerencialData(op, '2026-10-01', '2026-10-30'), /administradores/);
  assert.doesNotThrow(() => rt.context.getGerencialData(admin, '2026-10-01', '2026-10-30'));
  assert.throws(() => rt.context.getMetaBootstrap(admin), /Master/);
  assert.equal(plain(rt.context.getMetaBootstrap(master)).modo, 'demonstracao');
  assert.throws(() => rt.context.getDashboardData('garbage', null, null), /Sessão inválida/);
});

test('no credentials live in the source: seeds without properties create no accounts', () => {
  const rt = createRuntime({ now: NOW, files: [path.join(__dirname, '..', 'apps-script', 'Code.gs')], htmlFiles: {} });
  rt.evaluate('inicializarPlanilha()');
  const users = rt.workbooks.get(rt.properties.ASSINA_AI_PLANILHA_ID).getSheetByName('Usuarios').data;
  assert.equal(users.length, 1, 'only the header row');
});

test('Meta settings can only be saved by Master and are validated', () => {
  const rt = fresh();
  const admin = login(rt, 'admin.demo');
  const master = login(rt, 'Master');
  const input = {
    inss: { auditoria: 80, mediaDiaria: 0.5, ticketMedio: 6000, pausa: 16, valorPago: 60000 },
    atendimento: { auditoria: 80, tmaLigacao: 5, tmaChat: 18, atendimentos: 250 },
    clt: { auditoria: 80, pausa: 16, valorPago: 500000 },
  };
  assert.throws(() => rt.context.salvarConfiguracaoMeta(admin, input), /Master/);
  const saved = plain(rt.context.salvarConfiguracaoMeta(master, input));
  assert.equal(saved.confirmada, true);
  assert.equal(saved.clt.valorPago, 500000);
  assert.ok(saved.clt.mediaDiaria > 0);
  const bad = JSON.parse(JSON.stringify(input)); bad.inss.auditoria = 140;
  assert.throws(() => rt.context.salvarConfiguracaoMeta(master, bad), /inválido/);
});

test('rows from unknown correspondents and human analysts on the API desk are ignored', () => {
  const rt = fresh();
  const { rows } = plain(rt.context.getRawRows_());
  const corr = new Set(rows.map((r) => r['Correspondente']));
  assert.deepEqual([...corr].sort(), ['CORRESPONDENT A - 0001', 'CORRESPONDENT B - 0002']);
  const humanOnApiDesk = rows.filter((r) => r['Correspondente'] === 'CORRESPONDENT A - 0001' && !/API/.test(r['Analista']));
  assert.equal(humanOnApiDesk.length, 0);
  const sheetRows = rt.workbook.getSheetByName('Raw').getLastRow() - 1;
  assert.ok(rows.length < sheetRows, 'the synthetic contamination rows were dropped');
});

test('period filter returns only proposals inside the range', () => {
  const rt = fresh();
  const { rows } = plain(rt.context.getRawRows_('2026-10-28', '2026-10-29'));
  assert.ok(rows.length > 0);
  rows.forEach((r) => assert.match(r['Solic. Proposta'], /^2026-10-(28|29)/));
});

test('stalled cases: only the stalled stages, one row per customer (latest proposal)', () => {
  const rt = fresh();
  const { rows } = plain(rt.context.getRawRows_());
  const latest = plain(rt.context.agruparMaisRecentePorCliente_(rows));
  const ids = latest.map((r) => r._idCliente);
  assert.equal(new Set(ids).size, ids.length, 'one entry per customer');

  const sup = plain(rt.context.getSupervisoresMap_());
  const result = plain(rt.context.computeNaoAssinados_(latest.map((r) => Object.assign(r, { _dataHora: new Date(r._dataHora) })), sup, {}, {}));
  const allowed = plain(rt.evaluate('STATUS_NAO_ASSINADOS_ORDEM'));
  assert.deepEqual(allowed, ['NÃO REDIGITOU', 'PAROU NA BIO', 'ASSINADO', 'CONF DE FRAUDE', 'CONFERÊNCIA',
    'PROCESSANDO PAGAMENTO AUTOMÁTICO', 'PENDENTE REGULARIZAÇÃO'], 'the business definition of a stalled case');
  result.detalhe.forEach((d) => assert.ok(allowed.includes(d.validacao), d.validacao + ' is not a stalled stage'));
  assert.ok(!('PAGO' in result.pivotCount) && !('CANCELADO' in result.pivotCount));

  const labels = { 'não assinado': 'NÃO REDIGITOU', 'assinado o primeiro': 'PAROU NA BIO', 'assinado': 'ASSINADO',
    'conferência de fraude': 'CONF DE FRAUDE', 'conferência': 'CONFERÊNCIA',
    'processando pagamento automático': 'PROCESSANDO PAGAMENTO AUTOMÁTICO', 'pendente regularização': 'PENDENTE REGULARIZAÇÃO' };
  const newest = {};
  rows.forEach((r) => {
    const k = r['ID Cliente'];
    const when = r['Solic. Proposta'] + ' ' + r['Hora Solic. Proposta'];
    if (!newest[k] || when > newest[k].when) newest[k] = { when, status: String(r['Status Admin.']).toLowerCase() };
  });
  const expected = Object.values(newest).filter((n) => labels[n.status]).length;
  assert.equal(result.detalhe.length, expected);
  const pivotTotal = Object.values(result.pivotCount).reduce((a, b) => a + b, 0);
  assert.equal(pivotTotal, result.detalhe.length, 'pivot and detail agree');
  assert.ok(result.detalhe.length > 0);
});

test('funnel and stalled table are computed from the same deduplicated list and agree', () => {
  const rt = fresh();
  const token = login(rt, 'admin.demo');
  const data = plain(rt.context.getDashboardData(token, '2026-10-01', '2026-10-30', true));
  const funnelTotal = Object.values(data.funil.Todos.contagem).reduce((a, b) => a + b, 0);
  const funnelByStage = data.funil.Todos.contagem;
  Object.keys(data.naoAssinados.pivotCount).forEach((stage) => {
    assert.equal(data.naoAssinados.pivotCount[stage], funnelByStage[stage], 'stage ' + stage);
  });
  assert.equal(funnelTotal, Object.values(data.funil.Callcenter.contagem).reduce((a, b) => a + b, 0)
    + Object.values(data.funil.API.contagem).reduce((a, b) => a + b, 0));
});

test('return-to-customer counts only outbound calls made AFTER the proposal', () => {
  const rt = fresh();
  const proposalAt = new Date(2026, 9, 29, 10, 0, 0).getTime();
  const phone = '(00) 90000-0001';
  const key = rt.context.telefonePadrao10_(phone);
  const callLog = {};
  callLog[key] = [
    { dataHora: proposalAt - 3600 * 1000, segundos: 120 },
    { dataHora: proposalAt + 3600 * 1000, segundos: 1 },
    { dataHora: proposalAt + 7200 * 1000, segundos: 90 },
  ];
  const r = plain(rt.context.calcularRetornoCliente_(phone, new Date(proposalAt), callLog));
  assert.deepEqual(r, { retornouCliente: true, tentativasRetorno: 2, tentativasBemSucedidas: 1 });
  const none = plain(rt.context.calcularRetornoCliente_('(00) 90000-0999', new Date(proposalAt), callLog));
  assert.equal(none.retornouCliente, false);
});

test('conversion joins Raw and paid sheets by proposal number, not by customer', () => {
  const rt = createRuntime({ now: NOW, files: [path.join(__dirname, '..', 'apps-script', 'Code.gs')], htmlFiles: {},
    properties: { ASSINA_AI_SEED_ADMIN_DEMO: 'x' } });
  rt.evaluate('inicializarPlanilha()');
  const wb = rt.workbooks.get(rt.properties.ASSINA_AI_PLANILHA_ID);
  const H = ['Solic. Proposta', 'Hora Solic. Proposta', 'N° Proposta', 'ID Cliente', 'Analista', 'Status Admin.', 'Correspondente', 'Val. Solic.'];
  const A = 'CORRESPONDENT B - 0002';
  const raw = wb.getSheetByName('Raw');
  raw.getRange(1, 1, 3, H.length).setValues([
    H,
    ['2026-10-20 09:00:00', '09:00:00', '1001', 'C1', 'Analyst 01', 'Pago', A, '1.000,00'],
    ['2026-10-29 09:00:00', '09:00:00', '1002', 'C1', 'Analyst 01', 'Cancelado', A, '2.000,00'],
  ]);
  wb.getSheetByName('Efetivados') || wb.insertSheet('Efetivados');
  wb.getSheetByName('Efetivados').getRange(1, 1, 2, H.length).setValues([H, ['2026-10-20 09:00:00', '09:00:00', '1001', 'C1', 'Analyst 01', 'Pago', A, '1.000,00']]);
  wb.insertSheet('EfetivadosAPI').getRange(1, 1, 1, H.length).setValues([H]);
  const pagas = plain(rt.context.mapaPropostasPagas_());
  assert.deepEqual(pagas, { 1001: 1000 });
  const clientes = rt.context.agruparMaisRecentePorCliente_(plain(rt.context.getRawRows_()).rows);
  assert.equal(clientes.length, 1);
  assert.equal(clientes[0]['N° Proposta'], '1002', 'the newest proposal represents the customer');
  const serie = plain(rt.context.agruparSerieConversao_(clientes, pagas));
  assert.equal(serie.Todos['2026-10-29'].pago, 0, 'the newest proposal is NOT converted just because the customer paid before');
});

test('cancelled: registry-return reasons are shortened, "#N/D" is kept out of the pivot but not out of the detail', () => {
  const rt = fresh();
  const token = login(rt, 'admin.demo');
  const data = plain(rt.context.getDashboardData(token, '2026-10-01', '2026-10-30', true));
  const reasons = Object.keys(data.cancelados.pivot);
  assert.ok(!reasons.includes('#N/D') && !reasons.includes(''));
  assert.ok(!reasons.some((r) => r.includes('TN - Monthly interest')), 'long TN reason normalised to the short form');
  assert.ok(reasons.includes('REGISTRY RETURN FAILURE: CET RATE 1.00%'));
  const cancelledTotal = data.cancelados.detalhe.length;
  const pivotTotal = Object.values(data.cancelados.pivot).reduce((a, b) => a + b.qtd, 0);
  assert.ok(cancelledTotal > pivotTotal, 'detail also holds cancelled cases without a named reason');
});

test('management view lists the five channels and counts paid proposals per channel', () => {
  const rt = fresh();
  const admin = login(rt, 'admin.demo');
  const g = plain(rt.context.getGerencialData(admin, '2026-10-01', '2026-10-30'));
  assert.deepEqual(g.canais.map((c) => c.canal), ['Channel A', 'Channel B', 'App', 'Landing Page', 'Call Center']);
  g.canais.forEach((c) => {
    assert.equal(c.comSeguro + c.semSeguro, c.totalPropostas);
    assert.ok(c.qtdeEfetivados <= c.totalPropostas + 5);
  });
});

test('stalled block export (CSV) matches the on-screen detail and is labelled "Casos Parados"', () => {
  const rt = fresh();
  const token = login(rt, 'admin.demo');
  const data = plain(rt.context.getDashboardData(token, '2026-10-01', '2026-10-30', true));
  const file = plain(rt.context.exportarBloco(token, 'casos', 'csv', '2026-10-01', '2026-10-30'));
  assert.match(file.nome, /Casos Parados.*\.csv$/);
  const text = Buffer.from(file.base64, 'base64').toString('utf8').replace(/^﻿/, '');
  const lines = text.split('\n');
  assert.equal(lines[0].split(';')[1], 'Validação');
  assert.equal(lines.length - 1, data.naoAssinados.detalhe.length);
});

test('exports without a session are refused', () => {
  const rt = fresh();
  assert.throws(() => rt.context.exportarComFiltros('bad-token', {}), /Sessão inválida/);
});
