'use strict';
const fs = require('fs');
const path = require('path');
const { createRuntime } = require('./gas_mock');

const ROOT = path.resolve(__dirname, '..');
const DEMO_LOGINS = ['admin.demo', 'Master', 'supervisor.demo', 'operator.demo'];

const demoPassword = (login) => 'demo-' + login.toLowerCase();

function loadDashboard(options = {}) {
  const properties = {};
  DEMO_LOGINS.forEach((login) => { properties['ASSINA_AI_SEED_' + login.toUpperCase()] = demoPassword(login); });

  const rt = createRuntime({
    files: [path.join(ROOT, 'apps-script', 'Code.gs')],
    htmlFiles: { Index: path.join(ROOT, 'apps-script', 'Index.html') },
    now: options.now,
    properties,
  });

  rt.evaluate('inicializarPlanilha()');
  const id = rt.properties['ASSINA_AI_PLANILHA_ID'];
  const workbook = rt.workbooks.get(id);
  const data = JSON.parse(fs.readFileSync(options.dataPath || process.env.DEMO_DATA || path.join(__dirname, 'demo_data.json'), 'utf8'));
  Object.keys(data.sheets).forEach((name) => {
    let sheet = workbook.getSheetByName(name) || workbook.insertSheet(name);
    sheet.clearContents();
    const { headers, rows } = data.sheets[name];
    const all = (headers.length ? [headers] : []).concat(rows);
    if (all.length) {
      const width = Math.max(...all.map((l) => l.length));
      sheet.getRange(1, 1, all.length, width).setValues(all.map((l) => l.concat(Array(width - l.length).fill(''))));
    }
  });
  rt.cache.remove('config_mapa_v1');
  return Object.assign(rt, { workbook, data });
}

module.exports = { loadDashboard, demoPassword, DEMO_LOGINS, ROOT };
