const SHEET_RAW = 'Raw';
const SHEET_GER_TOTAL = 'TotalPropostas';
const SHEET_GER_EFETIVADOS = 'Efetivados';
const SHEET_GER_EFETIVADOS_API = 'EfetivadosAPI';
const SHEET_SUPERVISORES = 'Supervisores';
const SHEET_CONFIG = 'Config';
const SHEET_SNAPSHOT = 'Snapshot_Diario';
const SHEET_ATRIBUICOES = 'Atribuicoes';
const SHEET_USUARIOS = 'Usuarios';
const SHEET_TENTATIVAS = 'TentativasLogin';
const SHEET_LOG_ACESSOS = 'LogAcessos';
const PLANILHA_ID_PROP = 'ASSINA_AI_PLANILHA_ID';
let PLANILHA_ABERTA_ = null;

function planilha_() {
  if (PLANILHA_ABERTA_) return PLANILHA_ABERTA_;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PLANILHA_ID_PROP);
  if (id) {
    try {
      PLANILHA_ABERTA_ = SpreadsheetApp.openById(id);
      return PLANILHA_ABERTA_;
    } catch (e) {
    }
  }
  const ss = SpreadsheetApp.create('Assina AI : Database');
  props.setProperty(PLANILHA_ID_PROP, ss.getId());
  Logger.log('Planilha criada: ' + ss.getUrl());
  PLANILHA_ABERTA_ = ss;
  return PLANILHA_ABERTA_;
}

const CORRESPONDENTES_VALIDOS = ['CORRESPONDENT A - 0001', 'CORRESPONDENT B - 0002'];
const CORRESPONDENTE_API = 'CORRESPONDENT A - 0001';

const STATUS_LABEL = {
  'cancelado': 'CANCELADO',
  'banco': 'BANCO',
  'assinado o primeiro': 'PAROU NA BIO',
  'assinado': 'ASSINADO',
  'conferência': 'CONFERÊNCIA',
  'conferencia': 'CONFERÊNCIA',
  'conferência de fraude': 'CONF DE FRAUDE',
  'conferencia de fraude': 'CONF DE FRAUDE',
  'pago': 'PAGO',
  'conferido': 'CONFERIDO',
  'pendente regularização': 'PENDENTE REGULARIZAÇÃO',
  'liberado': 'LIBERADO',
  'pendente': 'PENDENTE',
  'pendente antifraude': 'PENDENTE ANTIFRAUDE',
  'pendente regularização multi transações': 'PENDENTE REGULARIZAÇÃO MULTI TRANSAÇÕES',
  'pendente retorno averbacao': 'PENDENTE RETORNO AVERBACAO',
  'processando pagamento automático': 'PROCESSANDO PAGAMENTO AUTOMÁTICO',
  'recompra': 'RECOMPRA',
  'reenviando': 'REENVIANDO',
  'reliberado': 'RELIBERADO',
  'conferencia documento': 'CONFERENCIA DOCUMENTO',
  'conferido cdc': 'CONFERIDO CDC',
  'corrigido': 'CORRIGIDO',
  'não assinado': 'NÃO REDIGITOU'
};

const ESTAGIOS_ESTEIRA = [
  'NÃO REDIGITOU', 'PAROU NA BIO', 'CONF DE FRAUDE', 'CONFERÊNCIA',
  'ASSINADO', 'CONFERIDO', 'LIBERADO', 'PENDENTE', 'PENDENTE ANTIFRAUDE',
  'PENDENTE REGULARIZAÇÃO', 'PENDENTE REGULARIZAÇÃO MULTI TRANSAÇÕES',
  'PENDENTE RETORNO AVERBACAO', 'PROCESSANDO PAGAMENTO AUTOMÁTICO',
  'RECOMPRA', 'REENVIANDO', 'RELIBERADO', 'CONFERENCIA DOCUMENTO',
  'CONFERIDO CDC', 'CORRIGIDO', 'BANCO'
];
const ESTAGIO_PAGO = 'PAGO';
const ESTAGIO_CANCELADO = 'CANCELADO';

const MOTIVO_TN_ORIGINAL = 'REGISTRY RETURN FAILURE: TN - Monthly interest rate is invalid: it must be lower than or equal to the CET rate, with a spread of at most 1.00%';
const MOTIVO_TN_CURTO = 'REGISTRY RETURN FAILURE: CET RATE 1.00%';

const APP_GROUPS = [
  ['API Consignado - Landing Page', 'LANDING'],
  ['API Consignado App', 'APP'],
  ['API Consignado - Leader', 'LEADER'],
  ['API Consignado Partner Desk', 'PARTNER DESK'],
  ['API Consignado Channel A', 'CHANNEL A'],
  ['API Consignado Channel B', 'CHANNEL B']
];

function normalizeCpf_(cpf) {
  return String(cpf || '').replace(/\D/g, '');
}

function cleanText_(text) {
  if (text === null || text === undefined) return '';
  return String(text).normalize('NFKC').replace(/[\x00-\x1f\x7f-\x9f]/g, ' ').replace(/\s+/g, ' ').trim();
}

function convertMonetaryToFloat_(value) {
  if (value === null || value === undefined || value === '') return NaN;
  if (typeof value === 'number') return value;
  let clean = String(value).replace('R$', '').replace(/\s/g, '').trim();
  if (!clean) return NaN;
  const lastComma = clean.lastIndexOf(',');
  const lastDot = clean.lastIndexOf('.');
  const lastSep = Math.max(lastComma, lastDot);
  if (lastSep === -1) {
    const n = parseFloat(clean);
    return isNaN(n) ? NaN : n;
  }
  const before = clean.slice(0, lastSep).replace(/\./g, '').replace(/,/g, '');
  const after = clean.slice(lastSep + 1);
  const n = parseFloat(before + '.' + after);
  return isNaN(n) ? NaN : n;
}

function cachePutJson_(cache, key, valor, ttlSegundos) {
  const str = JSON.stringify(valor);
  const TAM = 90000;
  const total = Math.max(1, Math.ceil(str.length / TAM));
  try {
    const entradas = {};
    for (let i = 0; i < total; i++) entradas[key + '_' + i] = str.substr(i * TAM, TAM);
    entradas[key + '_n'] = String(total);
    cache.putAll(entradas, ttlSegundos);
  } catch (e) {  }
}
function cacheGetJson_(cache, key) {
  const totalStr = cache.get(key + '_n');
  if (!totalStr) return null;
  const total = Number(totalStr);
  const chaves = [];
  for (let i = 0; i < total; i++) chaves.push(key + '_' + i);
  const partes = cache.getAll(chaves);
  let str = '';
  for (let i = 0; i < total; i++) {
    const parte = partes[key + '_' + i];
    if (parte === undefined || parte === null) return null;
    str += parte;
  }
  try { return JSON.parse(str); } catch (e) { return null; }
}

function classificarSetor_(analista) {
  const a = cleanText_(analista).toUpperCase();
  return a.indexOf('API ') === 0 ? 'API' : 'Callcenter';
}

function classificarEstagioAtual_(statusAdmin) {
  const norm = cleanText_(statusAdmin).toLowerCase();
  if (!norm) return 'DESCONHECIDO';
  return STATUS_LABEL[norm] || norm.toUpperCase();
}

const CHAVE_SESSAO_PROP = 'ASSINA_AI_CHAVE_SESSAO';
const MENSAGEM_ERRO_LOGIN = 'Login ou senha incorretos. Tente novamente.';
const REGEX_LOGIN_VALIDO = /^[A-Za-z0-9._-]{3,40}$/;
const REGEX_EMOJI = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{2190}-\u{21FF}\u{2B00}-\u{2BFF}]/u;

const LIMITES_BLOQUEIO = [
  { minimo: 20, minutos: 48 * 60 },
  { minimo: 15, minutos: 6 * 60 },
  { minimo: 10, minutos: 60 },
  { minimo: 5, minutos: 2 }
];
function duracaoBloqueioMinutos_(tentativas) {
  for (let i = 0; i < LIMITES_BLOQUEIO.length; i++) {
    if (tentativas >= LIMITES_BLOQUEIO[i].minimo) return LIMITES_BLOQUEIO[i].minutos;
  }
  return 0;
}

function hashSenha_(senha, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + senha);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function entradaValida_(login, senha) {
  if (typeof login !== 'string' || typeof senha !== 'string') return false;
  if (!REGEX_LOGIN_VALIDO.test(login)) return false;
  if (senha.length < 3 || senha.length > 100) return false;
  if (REGEX_EMOJI.test(login) || REGEX_EMOJI.test(senha)) return false;
  if (/[<>`]/.test(login) || /[<>`]/.test(senha)) return false;
  return true;
}

function lerTentativas_() {
  const aba = planilha_().getSheetByName(SHEET_TENTATIVAS);
  const dados = aba.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < dados.length; i++) {
    mapa[dados[i][0]] = { linha: i + 1, tentativas: Number(dados[i][1] || 0), bloqueadoAte: dados[i][3] || null };
  }
  return { aba, mapa };
}

function bloqueioDoRegistro_(reg) {
  if (!reg || !reg.bloqueadoAte) return null;
  const ate = new Date(reg.bloqueadoAte);
  return ate.getTime() > Date.now() ? ate : null;
}

function registrarFalhaEm_(estado, chave) {
  const reg = estado.mapa[chave];
  const tentativas = (reg ? reg.tentativas : 0) + 1;
  const agora = new Date();
  const minutos = duracaoBloqueioMinutos_(tentativas);
  const ate = minutos > 0 ? new Date(Date.now() + minutos * 60000) : '';
  if (reg) {
    estado.aba.getRange(reg.linha, 2, 1, 3).setValues([[tentativas, agora, ate]]);
  } else {
    estado.aba.appendRow([chave, tentativas, agora, ate]);
  }
  return ate || null;
}

function limparTentativasEm_(estado, chave) {
  const reg = estado.mapa[chave];
  if (!reg) return;
  estado.aba.getRange(reg.linha, 2, 1, 3).setValues([[0, new Date(), '']]);
}

function maisTarde_(a, b) {
  if (a && b) return a.getTime() > b.getTime() ? a : b;
  return a || b || null;
}

function mensagemBloqueio_(ate) {
  const min = Math.max(1, Math.ceil((ate.getTime() - Date.now()) / 60000));
  if (min < 60) return 'Muitas tentativas. Tente novamente em ' + min + ' minuto' + (min > 1 ? 's' : '') + '.';
  const horas = Math.ceil(min / 60);
  return 'Muitas tentativas. Tente novamente em ' + horas + ' hora' + (horas > 1 ? 's' : '') + '.';
}

function lerUsuarios_() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get('usuarios_v1');
  if (cacheado) return JSON.parse(cacheado);
  const dados = planilha_().getSheetByName(SHEET_USUARIOS).getDataRange().getValues();
  try { cache.put('usuarios_v1', JSON.stringify(dados), 300); } catch (e) {  }
  return dados;
}

function verificarLogin(login, senha, deviceToken) {
  login = String(login == null ? '' : login).trim();
  senha = String(senha == null ? '' : senha);
  deviceToken = String(deviceToken == null ? '' : deviceToken).trim();
  if (!/^[A-Za-z0-9-]{8,64}$/.test(deviceToken)) deviceToken = 'sem-token';

  if (!entradaValida_(login, senha)) return { ok: false, mensagem: MENSAGEM_ERRO_LOGIN };

  const chaveLogin = 'login:' + login.toLowerCase();
  const chaveDispositivo = 'device:' + deviceToken;

  const estado = lerTentativas_();
  const bloqueioAtivo = maisTarde_(bloqueioDoRegistro_(estado.mapa[chaveLogin]), bloqueioDoRegistro_(estado.mapa[chaveDispositivo]));
  if (bloqueioAtivo) {
    return { ok: false, mensagem: mensagemBloqueio_(bloqueioAtivo), bloqueado: true, bloqueadoAte: bloqueioAtivo.getTime() };
  }

  const dados = lerUsuarios_();
  let usuario = null;
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === login.toLowerCase() && dados[i][4] === true) {
      usuario = { login: dados[i][0], senhaHash: dados[i][1], salt: dados[i][2], papel: dados[i][3] };
      break;
    }
  }

  const sucesso = !!usuario && usuario.senhaHash === hashSenha_(senha, usuario.salt);

  if (!sucesso) {
    const ateFalha = maisTarde_(registrarFalhaEm_(estado, chaveLogin), registrarFalhaEm_(estado, chaveDispositivo));
    registrarAcessoAsync(login, '', false, 'login-invalido');
    if (ateFalha) return { ok: false, mensagem: mensagemBloqueio_(ateFalha), bloqueado: true, bloqueadoAte: ateFalha.getTime() };
    return { ok: false, mensagem: MENSAGEM_ERRO_LOGIN };
  }

  limparTentativasEm_(estado, chaveLogin);
  limparTentativasEm_(estado, chaveDispositivo);
  registrarAcessoAsync(usuario.login, usuario.papel, true, 'login');
  return {
    ok: true,
    login: usuario.login,
    papel: usuario.papel,
    podeAcessarMeta: String(usuario.login || '').toLowerCase() === 'master',
    token: gerarTokenSessao_(usuario.login, usuario.papel)
  };
}

function verificarSenhaAtual(tokenSessao, senha) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) return { ok: false, mensagem: 'Sessão inválida ou expirada. Faça login novamente.' };
  const dados = lerUsuarios_();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === sessao.login.toLowerCase()) {
      const ok = dados[i][1] === hashSenha_(String(senha || ''), dados[i][2]);
      return { ok: ok, mensagem: ok ? '' : 'Senha incorreta.' };
    }
  }
  return { ok: false, mensagem: 'Usuário não encontrado.' };
}

function registrarAcessoAsync(login, papel, sucesso, detalhe) {
  try {
    planilha_().getSheetByName(SHEET_LOG_ACESSOS).appendRow([new Date(), login, papel, sucesso, detalhe || '']);
  } catch (e) {  }
}

function chaveAssinatura_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty(CHAVE_SESSAO_PROP);
  if (!k) { k = Utilities.getUuid(); props.setProperty(CHAVE_SESSAO_PROP, k); }
  return k;
}
function assinar_(payload) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload + chaveAssinatura_());
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function gerarTokenSessao_(login, papel) {
  const expira = Date.now() + 8 * 60 * 60 * 1000;
  const payload = login + '|' + papel + '|' + expira;
  return Utilities.base64EncodeWebSafe(payload) + '.' + assinar_(payload);
}
function validarTokenSessao(token) {
  try {
    const partes = String(token || '').split('.');
    if (partes.length !== 2) return { ok: false };
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString();
    if (assinar_(payload) !== partes[1]) return { ok: false };
    const campos = payload.split('|');
    const expira = Number(campos[2]);
    if (!expira || Date.now() > expira) return { ok: false, expirado: true };
    return {
      ok: true,
      login: campos[0],
      papel: campos[1],
      podeAcessarMeta: String(campos[0] || '').toLowerCase() === 'master'
    };
  } catch (e) {
    return { ok: false };
  }
}
function exigirSessaoValida_(tokenSessao) {
  const v = validarTokenSessao(tokenSessao);
  return v.ok ? v : null;
}
function exigirAdmin_(tokenSessao) {
  const v = exigirSessaoValida_(tokenSessao);
  return (v && v.papel === 'Admin') ? v : null;
}
function exigirMaster_(tokenSessao) {
  const v = exigirAdmin_(tokenSessao);
  return (v && String(v.login || '').toLowerCase() === 'master') ? v : null;
}

const META_PROP_PREFIXO = 'ASSINA_AI_META_CONFIG_';
const META_TIMEZONE = 'America/Sao_Paulo';

function competenciaMetaAtual_() {
  return Utilities.formatDate(new Date(), META_TIMEZONE, 'yyyy-MM');
}

function dataLocalMeta_(ano, mes, dia) {
  return new Date(ano, mes - 1, dia, 12, 0, 0, 0);
}

function pascoaMeta_(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return dataLocalMeta_(ano, mes, dia);
}

function feriadosNacionaisMeta_(ano) {
  const fixos = [[1,1],[4,21],[5,1],[9,7],[10,12],[11,2],[11,15],[11,20],[12,25]];
  const datas = fixos.map(([mes, dia]) => Utilities.formatDate(dataLocalMeta_(ano, mes, dia), META_TIMEZONE, 'yyyy-MM-dd'));
  const sextaPaixao = pascoaMeta_(ano);
  sextaPaixao.setDate(sextaPaixao.getDate() - 2);
  datas.push(Utilities.formatDate(sextaPaixao, META_TIMEZONE, 'yyyy-MM-dd'));
  return new Set(datas);
}

function diasUteisMeta_(competencia) {
  const partes = String(competencia || '').split('-').map(Number);
  const ano = partes[0], mes = partes[1];
  if (!ano || !mes || mes < 1 || mes > 12) throw new Error('Competência inválida.');
  const hojeIso = Utilities.formatDate(new Date(), META_TIMEZONE, 'yyyy-MM-dd');
  const ultimoDia = new Date(ano, mes, 0).getDate();
  const feriados = feriadosNacionaisMeta_(ano);
  let total = 0, decorridos = 0;
  for (let dia = 1; dia <= ultimoDia; dia++) {
    const data = dataLocalMeta_(ano, mes, dia);
    const iso = Utilities.formatDate(data, META_TIMEZONE, 'yyyy-MM-dd');
    const semana = data.getDay();
    if (semana === 0 || semana === 6 || feriados.has(iso)) continue;
    total++;
    if (iso <= hojeIso) decorridos++;
  }
  return { total: total, decorridos: Math.min(total, decorridos), restantes: Math.max(0, total - decorridos) };
}

function configuracaoMetaPadrao_(competencia) {
  const dias = diasUteisMeta_(competencia).total || 1;
  return {
    competencia: competencia,
    confirmada: false,
    inss: { auditoria: 80, mediaDiaria: 0.5, ticketMedio: 6000, pausa: 16, valorPago: 60000 },
    atendimento: { auditoria: 80, tmaLigacao: 5, tmaChat: 18, atendimentos: 250 },
    clt: { auditoria: 80, mediaDiaria: Math.round((500000 / dias) * 100) / 100, pausa: 16, valorPago: 500000 },
    faixas: [
      { eficiencia: 80, bonus: 300 },
      { eficiencia: 85, bonus: 450 },
      { eficiencia: 90, bonus: 600 },
      { eficiencia: 95, bonus: 750 }
    ],
    absenteismo: [
      { maximo: 4.17, fator: 1 },
      { maximo: 8, fator: 0.8 },
      { maximo: 12, fator: 0.5 },
      { maximo: null, fator: 0 }
    ]
  };
}

function lerConfiguracaoMeta_(competencia) {
  const padrao = configuracaoMetaPadrao_(competencia);
  const bruto = PropertiesService.getScriptProperties().getProperty(META_PROP_PREFIXO + competencia);
  if (!bruto) return padrao;
  try {
    const salvo = JSON.parse(bruto);
    return Object.assign({}, padrao, salvo, {
      inss: Object.assign({}, padrao.inss, salvo.inss || {}),
      atendimento: Object.assign({}, padrao.atendimento, salvo.atendimento || {}),
      clt: Object.assign({}, padrao.clt, salvo.clt || {})
    });
  } catch (e) {
    return padrao;
  }
}

function numeroMeta_(valor, minimo, maximo, campo) {
  const n = Number(valor);
  if (!isFinite(n) || n < minimo || n > maximo) throw new Error('Valor inválido em ' + campo + '.');
  return Math.round(n * 100) / 100;
}

function salvarConfiguracaoMeta(tokenSessao, entrada) {
  const sessao = exigirMaster_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito ao login Master.');
  const competencia = competenciaMetaAtual_();
  const dias = diasUteisMeta_(competencia).total || 1;
  entrada = entrada || {};
  const inss = entrada.inss || {}, atendimento = entrada.atendimento || {}, clt = entrada.clt || {};
  const config = configuracaoMetaPadrao_(competencia);
  config.confirmada = true;
  config.atualizadaEm = new Date().toISOString();
  config.inss = {
    auditoria: numeroMeta_(inss.auditoria, 0, 100, 'Auditoria INSS'),
    mediaDiaria: numeroMeta_(inss.mediaDiaria, 0.01, 1000, 'Média diária INSS'),
    ticketMedio: numeroMeta_(inss.ticketMedio, 1, 10000000, 'Ticket médio INSS'),
    pausa: numeroMeta_(inss.pausa, 0, 100, 'Pausa INSS'),
    valorPago: numeroMeta_(inss.valorPago, 1, 1000000000, 'Valor pago INSS')
  };
  config.atendimento = {
    auditoria: numeroMeta_(atendimento.auditoria, 0, 100, 'Auditoria Atendimento'),
    tmaLigacao: numeroMeta_(atendimento.tmaLigacao, 0.1, 1440, 'TMA ligação'),
    tmaChat: numeroMeta_(atendimento.tmaChat, 0.1, 1440, 'TMA chat'),
    atendimentos: numeroMeta_(atendimento.atendimentos, 1, 10000000, 'Atendimentos')
  };
  config.clt = {
    auditoria: numeroMeta_(clt.auditoria, 0, 100, 'Auditoria CLT'),
    mediaDiaria: Math.round((numeroMeta_(clt.valorPago, 1, 1000000000, 'Valor pago CLT') / dias) * 100) / 100,
    pausa: numeroMeta_(clt.pausa, 0, 100, 'Pausa CLT'),
    valorPago: numeroMeta_(clt.valorPago, 1, 1000000000, 'Valor pago CLT')
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    PropertiesService.getScriptProperties().setProperty(META_PROP_PREFIXO + competencia, JSON.stringify(config));
  } finally {
    lock.releaseLock();
  }
  registrarAcessoAsync(sessao.login, sessao.papel, true, 'meta-config:' + competencia);
  return config;
}

function getMetaBootstrap(tokenSessao) {
  const sessao = exigirMaster_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito ao login Master.');
  const competencia = competenciaMetaAtual_();
  return {
    modo: 'demonstracao',
    aviso: 'Dados demonstrativos para validação da interface. Nenhuma base produtiva está conectada.',
    competencia: competencia,
    diasUteis: diasUteisMeta_(competencia),
    configuracao: lerConfiguracaoMeta_(competencia),
    dados: {
      absenteismo: 8.6,
      operacoes: [
        { id: 'inss', nome: 'INSS / FGTS', pessoas: 12, metricas: [
          { id: 'auditoria', nome: 'Auditoria', atual: 93.4, peso: 15, sentido: 'maior', formato: 'percentual' },
          { id: 'mediaDiaria', nome: 'Produção diária', atual: 0.29, peso: 20, sentido: 'maior', formato: 'decimal' },
          { id: 'ticketMedio', nome: 'Ticket médio', atual: 4261, peso: 15, sentido: 'maior', formato: 'moeda' },
          { id: 'pausa', nome: 'Pausa', atual: 9.3, peso: 15, sentido: 'menor', formato: 'percentual' },
          { id: 'valorPago', nome: 'Valor pago', atual: 13800, peso: 35, sentido: 'maior', formato: 'moeda' }
        ]},
        { id: 'atendimento', nome: 'Atendimento', pessoas: 9, metricas: [
          { id: 'auditoria', nome: 'Auditoria', atual: 92.3, peso: 25, sentido: 'maior', formato: 'percentual' },
          { id: 'tmaLigacao', nome: 'TMA ligação', atual: 6.72, peso: 25, sentido: 'menor', formato: 'minutos' },
          { id: 'tmaChat', nome: 'TMA chat', atual: 218, peso: 25, sentido: 'menor', formato: 'minutos' },
          { id: 'atendimentos', nome: 'Atendimentos', atual: 80.5, peso: 25, sentido: 'maior', formato: 'numero' }
        ]},
        { id: 'clt', nome: 'Consignado CLT', pessoas: 28, metricas: [
          { id: 'auditoria', nome: 'Auditoria', atual: 93.6, peso: 25, sentido: 'maior', formato: 'percentual' },
          { id: 'mediaDiaria', nome: 'Média diária', atual: 17616, peso: 25, sentido: 'maior', formato: 'moeda' },
          { id: 'pausa', nome: 'Pausa', atual: 11.9, peso: 25, sentido: 'menor', formato: 'percentual' },
          { id: 'valorPago', nome: 'Valor pago', atual: 118000, peso: 25, sentido: 'maior', formato: 'moeda' }
        ]}
      ],
      distribuicao: [
        { rotulo: '< 80', quantidade: 31 },
        { rotulo: '80 a 84,9', quantidade: 9 },
        { rotulo: '85 a 89,9', quantidade: 5 },
        { rotulo: '90 a 94,9', quantidade: 2 },
        { rotulo: '≥ 95', quantidade: 1 }
      ]
    }
  };
}

function criarUsuario(login, senha, papel) {
  const aba = planilha_().getSheetByName(SHEET_USUARIOS);
  const salt = Utilities.getUuid();
  aba.appendRow([login, hashSenha_(senha, salt), salt, papel || 'Admin', true, new Date()]);
  try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {}
  Logger.log('Usuário criado: ' + login + ' (' + (papel || 'Admin') + ')');
}

function redefinirSenha(login, novaSenha) {
  const aba = planilha_().getSheetByName(SHEET_USUARIOS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === login.toLowerCase()) {
      const salt = Utilities.getUuid();
      aba.getRange(i + 1, 2, 1, 2).setValues([[hashSenha_(novaSenha, salt), salt]]);
      try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {}
      return true;
    }
  }
  return false;
}

const SHEET_CALL_LOG = 'CallLog';
const DURACAO_MINIMA_SUCESSO_SEGUNDOS = 3;

function normalizeTelefone_(v) {
  return String(v || '').replace(/\D/g, '');
}

function telefonePadrao10_(v) {
  let d = normalizeTelefone_(v);
  if (d.length === 12 || d.length === 13) d = d.slice(2);
  if (d.length === 11) d = d.slice(0, 2) + d.slice(3);
  return d;
}

function parseDataHoraProposta_(dataStr, horaStr) {
  let ano, mes, dia;
  if (dataStr instanceof Date && !isNaN(dataStr.getTime())) {
    ano = dataStr.getFullYear(); mes = dataStr.getMonth() + 1; dia = dataStr.getDate();
  } else {
    const d = cleanText_(dataStr);
    const mIso = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    const mBr = d.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (mIso) { ano = mIso[1]; mes = mIso[2]; dia = mIso[3]; }
    else if (mBr) { dia = mBr[1]; mes = mBr[2]; ano = mBr[3]; }
    else return null;
  }

  let h = 0, mi = 0, s = 0;
  if (horaStr instanceof Date && !isNaN(horaStr.getTime())) {
    h = horaStr.getHours(); mi = horaStr.getMinutes(); s = horaStr.getSeconds();
  } else if (typeof horaStr === 'number' && isFinite(horaStr)) {
    const segundos = Math.round((((horaStr % 1) + 1) % 1) * 86400) % 86400;
    h = Math.floor(segundos / 3600); mi = Math.floor((segundos % 3600) / 60); s = segundos % 60;
  } else {
    const hm = cleanText_(horaStr).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (hm) { h = Number(hm[1]); mi = Number(hm[2]); s = Number(hm[3] || 0); }
  }
  const dt = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(mi), Number(s));
  return isNaN(dt.getTime()) ? null : dt;
}

function parseDataHoraBr_(str) {
  const m = cleanText_(str).match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5]), Number(m[6]));
}

function getCallLogPorTelefone_(telefonesNecessarios) {
  const cache = CacheService.getScriptCache();
  const cacheado = cacheGetJson_(cache, 'call_log_v1');
  if (cacheado) return cacheado;
  const sh = planilha_().getSheetByName(SHEET_CALL_LOG);
  if (!sh) return {};
  const ultimaLinha = sh.getLastRow();
  if (ultimaLinha < 2) return {};
  const values = sh.getRange(2, 1, ultimaLinha - 1, Math.min(3, sh.getLastColumn())).getValues();
  const mapa = {};
  for (let r = 0; r < values.length; r++) {
    const celular = telefonePadrao10_(values[r][0]);
    if (telefonesNecessarios && !telefonesNecessarios.has(celular)) continue;
    const dataHora = parseDataHoraBr_(values[r][1]);
    if (!celular || !dataHora) continue;
    if (!mapa[celular]) mapa[celular] = [];
    mapa[celular].push({ dataHora: dataHora.getTime(), segundos: Number(values[r][2]) || 0 });
  }
  if (!telefonesNecessarios) cachePutJson_(cache, 'call_log_v1', mapa, 600);
  return mapa;
}

function calcularRetornoCliente_(celularRow, dataHoraProposta, callLogMap) {
  if (!dataHoraProposta) return { retornouCliente: false, tentativasRetorno: 0, tentativasBemSucedidas: 0 };
  const instanteProposta = dataHoraProposta.getTime();
  const chamadas = (callLogMap[telefonePadrao10_(celularRow)] || [])
    .filter(c => Number(c.dataHora) > instanteProposta);
  const bemSucedidas = chamadas.filter(c => c.segundos >= DURACAO_MINIMA_SUCESSO_SEGUNDOS).length;
  return { retornouCliente: chamadas.length > 0, tentativasRetorno: chamadas.length, tentativasBemSucedidas: bemSucedidas };
}

const TABULACOES_ID_PROP = 'ASSINA_AI_TABULACOES_PLANILHA_ID';
const ABAS_TABULACOES_IGNORAR_ = new Set([
  'MENU', 'CONSOLIDADO', 'BASE DE DADOS', 'BASE DE DADOS - CANCELADOS',
  'RELATÓRIO - NÃO ASSINADOS', 'RELATÓRIO - CANCELADOS', 'DETALHAMENTO'
]);

function getMapaTabulacoes_() {
  const cache = CacheService.getScriptCache();
  const cacheado = cacheGetJson_(cache, 'tabulacoes_v2');
  if (cacheado) return cacheado;

  const mapa = {};
  let abasEncontradas = 0;
  const idTabulacoes = PropertiesService.getScriptProperties().getProperty(TABULACOES_ID_PROP);
  if (!idTabulacoes) return mapa;
  try {
    const ssTab = SpreadsheetApp.openById(idTabulacoes);
    ssTab.getSheets().forEach(aba => {
      const nome = aba.getName().trim().toUpperCase();
      if (ABAS_TABULACOES_IGNORAR_.has(nome)) return;
      const ultimaLinha = aba.getLastRow();
      const ultimaColuna = aba.getLastColumn();
      if (ultimaLinha < 4 || ultimaColuna < 1) return;

      const cabecalho = aba.getRange(3, 1, 1, ultimaColuna).getValues()[0]
        .map(h => String(h || '').trim().toUpperCase());
      const colProposta = cabecalho.indexOf('Nº PROPOSTA');
      const colT1 = cabecalho.indexOf('RETORNO TENTATIVA 1º');
      const colT2 = cabecalho.indexOf('RETORNO TENTATIVA 2º');
      if (colProposta === -1 || (colT1 === -1 && colT2 === -1)) return;
      abasEncontradas++;

      const valores = aba.getRange(4, 1, ultimaLinha - 3, ultimaColuna).getValues();
      valores.forEach(linha => {
        const proposta = cleanText_(linha[colProposta]);
        if (!proposta) return;
        const t1 = colT1 !== -1 ? cleanText_(linha[colT1]) : '';
        const t2 = colT2 !== -1 ? cleanText_(linha[colT2]) : '';
        if (!t1 && !t2) return;
        mapa[proposta] = { t1: t1, t2: t2 };
      });
    });
  } catch (e) {
    console.error('getMapaTabulacoes_ falhou: ' + e);
    return mapa;
  }

  if (abasEncontradas === 0) {
    console.warn('getMapaTabulacoes_: planilha de tabulações abriu, mas nenhuma aba teve o cabeçalho esperado (Nº PROPOSTA / RETORNO TENTATIVA).');
    return mapa;
  }

  console.log('getMapaTabulacoes_ OK: ' + abasEncontradas + ' aba(s) reconhecida(s), ' + Object.keys(mapa).length + ' proposta(s) no mapa.');
  cachePutJson_(cache, 'tabulacoes_v2', mapa, 600);
  return mapa;
}

function lerAbaProfostas_(nomeAba, dataInicio, dataFim) {
  const sh = planilha_().getSheetByName(nomeAba);
  if (!sh) return { headers: [], rows: [] };
  const ultimaLinha = sh.getLastRow();
  const ultimaColuna = sh.getLastColumn();
  if (ultimaLinha < 1 || ultimaColuna < 1) return { headers: [], rows: [] };
  const headers = sh.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(h => String(h).trim());
  if (ultimaLinha < 2) return { headers, rows: [] };

  let values;
  const dataIdx = headers.indexOf('Solic. Proposta');
  if ((dataInicio || dataFim) && dataIdx !== -1) {
    const datas = sh.getRange(2, dataIdx + 1, ultimaLinha - 1, 1).getValues();
    const limiteInicioBusca = dataInicio ? new Date(dataInicio + 'T00:00:00') : null;
    const limiteFimBusca = dataFim ? new Date(dataFim + 'T23:59:59') : null;
    let primeira = -1, ultima = -1;
    for (let i = 0; i < datas.length; i++) {
      const bruto = datas[i][0];
      const d = bruto instanceof Date ? bruto : parseDataHoraProposta_(bruto, '00:00:00');
      if (!d || (limiteInicioBusca && d < limiteInicioBusca) || (limiteFimBusca && d > limiteFimBusca)) continue;
      if (primeira === -1) primeira = i + 2;
      ultima = i + 2;
    }
    if (primeira === -1) return { headers, rows: [] };
    values = [headers].concat(sh.getRange(primeira, 1, ultima - primeira + 1, ultimaColuna).getValues());
  } else {
    values = sh.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
  }
  const corrIdx = headers.indexOf('Correspondente');
  const analistaIdx = headers.indexOf('Analista');

  const limiteInicio = dataInicio ? new Date(dataInicio + 'T00:00:00') : null;
  const limiteFim = dataFim ? new Date(dataFim + 'T23:59:59') : null;

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (corrIdx !== -1) {
      const corr = String(row[corrIdx] || '').trim().toUpperCase();
      if (CORRESPONDENTES_VALIDOS.indexOf(corr) === -1) continue;
      if (nomeAba === SHEET_RAW && corr === CORRESPONDENTE_API && analistaIdx !== -1) {
        const analista = String(row[analistaIdx] || '').toUpperCase();
        if (analista.indexOf('API') === -1) continue;
      }
    }
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);

    if (limiteInicio || limiteFim) {
      const dh = parseDataHoraProposta_(obj['Solic. Proposta'], obj['Hora Solic. Proposta']);
      if (limiteInicio && (!dh || dh < limiteInicio)) continue;
      if (limiteFim && (!dh || dh > limiteFim)) continue;
    }
    rows.push(obj);
  }
  return { headers, rows };
}

function getRawRows_(dataInicio, dataFim) {
  return lerAbaProfostas_(SHEET_RAW, dataInicio, dataFim);
}

function getGerTotalRows_(dataInicio, dataFim) {
  return lerAbaProfostas_(SHEET_GER_TOTAL, dataInicio, dataFim);
}

function getGerEfetivadosRows_(dataInicio, dataFim) {
  return lerAbaProfostas_(SHEET_GER_EFETIVADOS, dataInicio, dataFim);
}

function getGerEfetivadosApiRows_(dataInicio, dataFim) {
  return lerAbaProfostas_(SHEET_GER_EFETIVADOS_API, dataInicio, dataFim);
}

function getSupervisoresMap_() {
  const cache = CacheService.getScriptCache();
  const cacheado = cacheGetJson_(cache, 'supervisores_v1');
  if (cacheado) return cacheado;
  const sh = planilha_().getSheetByName(SHEET_SUPERVISORES);
  if (!sh) return {};
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < values.length; r++) {
    const analista = cleanText_(values[r][0]);
    if (analista) map[analista] = values[r][1] || '';
  }
  cachePutJson_(cache, 'supervisores_v1', map, 600);
  return map;
}

function getConfigMapa_() {
  const cache = CacheService.getScriptCache();
  const cacheado = cacheGetJson_(cache, 'config_mapa_v1');
  if (cacheado) return cacheado;
  const sh = planilha_().getSheetByName(SHEET_CONFIG);
  const values = sh.getDataRange().getValues();
  const cfg = {};
  for (let r = 1; r < values.length; r++) {
    const chave = values[r][0];
    if (chave) cfg[String(chave).trim()] = values[r][1];
  }
  cachePutJson_(cache, 'config_mapa_v1', cfg, 600);
  return cfg;
}

function getConfig_() {
  const cfg = getConfigMapa_();
  return {
    metaConversaoDiaria: Number(cfg['Meta_Conversao_Diaria']) || 0,
    limiarAlertaVermelho: Number(cfg['Limiar_Alerta_Vermelho']) || 0.15,
    diasMediaMovel: Number(cfg['Dias_Media_Movel']) || 7,
    horarioCorteDia: cfg['Horario_Corte_Dia'] || '23:59'
  };
}

function agruparMaisRecentePorCliente_(rows) {
  const maisRecentePorCliente = {};
  rows.forEach(row => {
    const idCliente = cleanText_(row['ID Cliente']);
    if (!idCliente) return;
    const dataHora = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (!dataHora) return;
    const atual = maisRecentePorCliente[idCliente];
    if (!atual || dataHora > atual._dataHora) {
      maisRecentePorCliente[idCliente] = Object.assign({}, row, { _dataHora: dataHora, _idCliente: idCliente });
    }
  });
  return Object.values(maisRecentePorCliente);
}

function computeFunil_(clientesMaisRecentes) {
  const resultado = {
    Todos: { contagem: {}, valores: {} },
    Callcenter: { contagem: {}, valores: {} },
    API: { contagem: {}, valores: {} }
  };
  clientesMaisRecentes.forEach(row => {
    const setor = classificarSetor_(row['Analista']);
    const estagio = classificarEstagioAtual_(row['Status Admin.']);
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    [resultado.Todos, resultado[setor]].forEach(alvo => {
      alvo.contagem[estagio] = (alvo.contagem[estagio] || 0) + 1;
      alvo.valores[estagio] = (alvo.valores[estagio] || 0) + (isNaN(valor) ? 0 : valor);
    });
  });
  return resultado;
}

function normalizarDataCelula_(valor) {
  if (valor instanceof Date) return Utilities.formatDate(valor, 'America/Sao_Paulo', 'yyyy-MM-dd');
  return String(valor);
}

function syncSnapshotDiario_() {
  const { rows } = getRawRows_();
  const clientesMaisRecentes = agruparMaisRecentePorCliente_(rows);
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');

  const mapaPagos = mapaPropostasPagas_();

  const agregadosHoje = {};
  const pagoPorDiaSetor = {};

  clientesMaisRecentes.forEach(row => {
    const setor = classificarSetor_(row['Analista']);
    const v = convertMonetaryToFloat_(row['Val. Solic.']);
    const valor = isNaN(v) ? 0 : v;

    if (cleanText_(row['N° Proposta']) in mapaPagos) {
      const dia = Utilities.formatDate(row._dataHora, 'America/Sao_Paulo', 'yyyy-MM-dd');
      const key = dia + '|' + setor;
      if (!pagoPorDiaSetor[key]) pagoPorDiaSetor[key] = { qtd: 0, valor: 0 };
      pagoPorDiaSetor[key].qtd += 1;
      pagoPorDiaSetor[key].valor += valor;
      return;
    }
    const estagio = classificarEstagioAtual_(row['Status Admin.']);
    const key = setor + '|' + estagio;
    if (!agregadosHoje[key]) agregadosHoje[key] = { qtd: 0, valor: 0 };
    agregadosHoje[key].qtd += 1;
    agregadosHoje[key].valor += valor;
  });

  const sh = planilha_().getSheetByName(SHEET_SNAPSHOT);
  const data = sh.getDataRange().getValues();
  const headerRow = data.length ? data[0] : ['Data', 'Setor', 'Estagio', 'Qtd', 'Valor'];
  const dateIdx = headerRow.indexOf('Data');
  const estagioIdx = headerRow.indexOf('Estagio');

  const linhasParaManter = [headerRow];
  for (let r = 1; r < data.length; r++) {
    const ehHoje = normalizarDataCelula_(data[r][dateIdx]) === hoje;
    const ehPago = data[r][estagioIdx] === ESTAGIO_PAGO;
    if (!ehHoje && !ehPago) linhasParaManter.push(data[r]);
  }
  Object.keys(agregadosHoje).forEach(key => {
    const [setor, estagio] = key.split('|');
    linhasParaManter.push([hoje, setor, estagio, agregadosHoje[key].qtd, agregadosHoje[key].valor]);
  });
  Object.keys(pagoPorDiaSetor).forEach(key => {
    const [dia, setor] = key.split('|');
    linhasParaManter.push([dia, setor, ESTAGIO_PAGO, pagoPorDiaSetor[key].qtd, pagoPorDiaSetor[key].valor]);
  });

  sh.clearContents();
  sh.getRange(1, 1, linhasParaManter.length, headerRow.length).setValues(linhasParaManter);
}

function diasNoIntervalo_(dataInicio, dataFim) {
  const dias = [];
  const d = new Date(dataInicio + 'T00:00:00');
  const fim = new Date(dataFim + 'T00:00:00');
  while (d <= fim) {
    dias.push(Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }
  return dias;
}

function mapaPropostasPagas_(dataInicio, dataFim) {
  const { rows: efetivados } = getGerEfetivadosRows_(dataInicio, dataFim);
  const { rows: efetivadosApi } = getGerEfetivadosApiRows_(dataInicio, dataFim);
  const mapa = {};
  [efetivados, efetivadosApi].forEach(lista => {
    lista.forEach(r => {
      const proposta = cleanText_(r['N° Proposta']);
      if (!proposta) return;
      const v = convertMonetaryToFloat_(r['Val. Solic.']);
      mapa[proposta] = isNaN(v) ? 0 : v;
    });
  });
  return mapa;
}

function agruparSerieConversao_(clientes, mapaPagos) {
  const porSetor = { Todos: {}, Callcenter: {}, API: {} };
  clientes.forEach(row => {
    const dia = Utilities.formatDate(row._dataHora, 'America/Sao_Paulo', 'yyyy-MM-dd');
    const setor = classificarSetor_(row['Analista']);
    const proposta = cleanText_(row['N° Proposta']);
    const convertido = proposta in mapaPagos;
    [porSetor.Todos, porSetor[setor]].forEach(m => {
      if (!m[dia]) m[dia] = { total: 0, pago: 0, valorPago: 0 };
      m[dia].total++;
      if (convertido) { m[dia].pago++; m[dia].valorPago += mapaPagos[proposta]; }
    });
  });
  return porSetor;
}

function computeSerieConversao_(rowsFornecidas) {
  const cache = CacheService.getScriptCache();
  const key = 'serie_conversao_v4';
  const cached = cacheGetJson_(cache, key);
  if (cached) return cached;

  const rows = rowsFornecidas || getRawRows_().rows;
  const clientes = agruparMaisRecentePorCliente_(rows);
  const mapaPagos = mapaPropostasPagas_();
  const porSetor = agruparSerieConversao_(clientes, mapaPagos);
  cachePutJson_(cache, key, porSetor, 600);
  return porSetor;
}

function computeSerieConversaoFaixa_(dataInicio, dataFim) {
  const cache = CacheService.getScriptCache();
  const key = 'serie_conversao_faixa_v1_' + dataInicio.replace(/-/g, '') + '_' + dataFim.replace(/-/g, '');
  const cached = cacheGetJson_(cache, key);
  if (cached) return cached;
  const { rows } = getRawRows_(dataInicio, dataFim);
  const clientes = agruparMaisRecentePorCliente_(rows);
  const resultado = agruparSerieConversao_(clientes, mapaPropostasPagas_(dataInicio, dataFim));
  cachePutJson_(cache, key, resultado, 600);
  return resultado;
}

function taxaMediaConversao_(buckets) {
  let soma = 0, n = 0;
  buckets.forEach(b => { if (b && b.total > 0) { soma += b.pago / b.total; n++; } });
  return n ? soma / n : 0;
}

function computeProjecao_(porDia, cfg, dataInicio, dataFim) {
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  const inicio = dataInicio || hoje;
  const dias = Object.keys(porDia).sort();
  const historico = dias.filter(d => d < inicio).slice(-cfg.diasMediaMovel);
  const taxaHistorica = taxaMediaConversao_(historico.map(d => porDia[d]));

  const selecionados = (dataInicio && dataFim) ? diasNoIntervalo_(dataInicio, dataFim) : [hoje];
  let totalPeriodo = 0, pagoHoje = 0, valorPagoHoje = 0;
  selecionados.forEach(d => {
    if (!porDia[d]) return;
    totalPeriodo += porDia[d].total;
    pagoHoje += porDia[d].pago;
    valorPagoHoje += porDia[d].valorPago;
  });

  return {
    taxaHistorica,
    convertivelHoje: totalPeriodo,
    pagoHoje,
    valorPagoHoje,
    projecaoPago: Math.round(totalPeriodo * taxaHistorica),
    alerta: historico.length > 0 && taxaHistorica < cfg.limiarAlertaVermelho
  };
}

function getUltimaSincronizacao_() {
  const bruto = getConfigMapa_()['Ultima_Sincronizacao'];
  if (!bruto) return null;
  if (bruto instanceof Date) return bruto.toISOString();
  const texto = String(bruto);
  if (/^\d{4}-\d{2}-\d{2}T/.test(texto)) return texto;
  const m = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) {
    const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5]), Number(m[6]));
    return dt.toISOString();
  }
  return texto;
}

function montarSerieTemporal_(porSetor, cfg) {
  const resultado = {};
  Object.keys(porSetor).forEach(setor => {
    const porDia = porSetor[setor];
    const dias = Object.keys(porDia).sort();
    resultado[setor] = dias.map((d, i) => {
      const anteriores = dias.slice(Math.max(0, i - cfg.diasMediaMovel), i).map(x => porDia[x]);
      const taxa = taxaMediaConversao_(anteriores);
      return { data: d, pago: porDia[d].pago, valorPago: porDia[d].valorPago, total: porDia[d].total, projecao: Math.round(porDia[d].total * taxa) };
    });
  });
  return resultado;
}

function getSerieTemporal(tokenSessao) {
  if (!exigirSessaoValida_(tokenSessao)) throw new Error('Sessão inválida ou expirada. Faça login novamente.');
  return montarSerieTemporal_(computeSerieConversao_(), getConfig_());
}

const STATUS_NAO_ASSINADOS_ORDEM = [
  'NÃO REDIGITOU', 'PAROU NA BIO', 'ASSINADO', 'CONF DE FRAUDE', 'CONFERÊNCIA',
  'PROCESSANDO PAGAMENTO AUTOMÁTICO', 'PENDENTE REGULARIZAÇÃO'
];

function computeNaoAssinados_(clientesMaisRecentes, supervisores, callLogMap, tabulacoesMap) {
  const pivotCount = {};
  const detalhe = [];
  const tabulacoes = tabulacoesMap || {};

  clientesMaisRecentes.forEach(row => {
    const estagio = classificarEstagioAtual_(row['Status Admin.']);
    if (STATUS_NAO_ASSINADOS_ORDEM.indexOf(estagio) === -1) return;
    pivotCount[estagio] = (pivotCount[estagio] || 0) + 1;

    const analista = cleanText_(row['Analista']);
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    const retorno = calcularRetornoCliente_(row['Celular'], row._dataHora, callLogMap);
    const tab = tabulacoes[cleanText_(row['N° Proposta'])] || {};
    detalhe.push({
      dataProposta: Utilities.formatDate(row._dataHora, 'America/Sao_Paulo', 'dd/MM/yyyy'),
      validacao: estagio,
      proposta: row['N° Proposta'],
      cliente: row['Cliente'],
      celular: row['Celular'],
      analista: analista,
      correspondente: row['Correspondente'],
      valor: isNaN(valor) ? null : valor,
      supervisor: supervisores[analista] || '',
      retornouCliente: retorno.retornouCliente,
      tentativasRetorno: retorno.tentativasRetorno,
      tentativasBemSucedidas: retorno.tentativasBemSucedidas,
      tabulacaoTentativa1: tab.t1 || '',
      tabulacaoTentativa2: tab.t2 || ''
    });
  });

  return { pivotCount, detalhe };
}

function normalizarMotivo_(motivo) {
  const clean = cleanText_(motivo);
  return clean === MOTIVO_TN_ORIGINAL ? MOTIVO_TN_CURTO : clean;
}

function computeCancelados_(rows, supervisores, callLogMap, tabulacoesMap) {
  const canceladas = rows.filter(r => cleanText_(r['Status Admin.']).toLowerCase() === 'cancelado');
  const pivot = {};
  const detalhe = [];
  const tabulacoes = tabulacoesMap || {};

  canceladas.forEach(row => {
    const motivoBruto = cleanText_(row['Motivo_Cancelamento']);
    const temMotivo = motivoBruto && motivoBruto.toUpperCase() !== '#N/D';
    const motivo = temMotivo ? normalizarMotivo_(motivoBruto) : '';
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);

    if (temMotivo) {
      if (!pivot[motivo]) pivot[motivo] = { qtd: 0, valor: 0 };
      pivot[motivo].qtd += 1;
      pivot[motivo].valor += isNaN(valor) ? 0 : valor;
    }

    const analista = cleanText_(row['Analista']);
    const dataHoraProposta = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    const retorno = calcularRetornoCliente_(row['Celular'], dataHoraProposta, callLogMap);
    const tab = tabulacoes[cleanText_(row['N° Proposta'])] || {};
    detalhe.push({
      dataProposta: dataHoraProposta ? Utilities.formatDate(dataHoraProposta, 'America/Sao_Paulo', 'dd/MM/yyyy') : '',
      motivo: motivo, proposta: row['N° Proposta'], cliente: row['Cliente'], celular: row['Celular'],
      analista, correspondente: row['Correspondente'], valor: isNaN(valor) ? null : valor,
      supervisor: supervisores[analista] || '',
      retornouCliente: retorno.retornouCliente,
      tentativasRetorno: retorno.tentativasRetorno,
      tentativasBemSucedidas: retorno.tentativasBemSucedidas,
      tabulacaoTentativa1: tab.t1 || '',
      tabulacaoTentativa2: tab.t2 || ''
    });
  });

  const distribuicaoPorApi = {};
  APP_GROUPS.forEach(([rotulo, keyword]) => {
    const grupo = canceladas.filter(r => cleanText_(r['Analista']).toUpperCase().indexOf(keyword) !== -1);
    if (grupo.length === 0) return;
    const porMotivo = {};
    grupo.forEach(row => {
      let motivo = cleanText_(row['Motivo_Cancelamento']);
      if (!motivo || motivo.toUpperCase() === '#N/D') return;
      motivo = normalizarMotivo_(motivo);
      const valor = convertMonetaryToFloat_(row['Val. Solic.']);
      if (!porMotivo[motivo]) porMotivo[motivo] = { qtd: 0, valor: 0 };
      porMotivo[motivo].qtd += 1;
      porMotivo[motivo].valor += isNaN(valor) ? 0 : valor;
    });
    distribuicaoPorApi[rotulo] = porMotivo;
  });

  return { pivot, detalhe, distribuicaoPorApi };
}


function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Assina AI : Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(tokenSessao, dataInicio, dataFim, forcar) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) throw new Error('Sessão inválida ou expirada. Faça login novamente.');

  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard_data_v12_' + (dataInicio || 'x') + '_' + (dataFim || 'x');
  if (!forcar) {
    const cached = cacheGetJson_(cache, cacheKey);
    if (cached) return cached;
  }

  const { rows } = getRawRows_(dataInicio, dataFim);
  const clientesMaisRecentes = agruparMaisRecentePorCliente_(rows);
  const supervisores = getSupervisoresMap_();
  const callLogMap = getCallLogPorTelefone_();
  const tabulacoesMap = getMapaTabulacoes_();
  const serieConversao = computeSerieConversao_();
  const cfg = getConfig_();

  const payload = {
    atualizadoEm: new Date().toISOString(),
    ultimaSincronizacao: getUltimaSincronizacao_(),
    funil: computeFunil_(clientesMaisRecentes),
    projecao: {
      Callcenter: computeProjecao_(serieConversao.Callcenter, cfg, dataInicio, dataFim),
      API: computeProjecao_(serieConversao.API, cfg, dataInicio, dataFim),
      Todos: computeProjecao_(serieConversao.Todos, cfg, dataInicio, dataFim)
    },
    serie: montarSerieTemporal_(serieConversao, cfg),
    naoAssinados: computeNaoAssinados_(clientesMaisRecentes, supervisores, callLogMap, tabulacoesMap),
    cancelados: computeCancelados_(rows, supervisores, callLogMap, tabulacoesMap)
  };
  cachePutJson_(cache, cacheKey, payload, 600);
  return payload;
}

function deslocarDataIso_(dataIso, dias) {
  const p = String(dataIso).split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12, 0, 0, 0);
  d.setDate(d.getDate() + dias);
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function inicioSemanaIso_(dataIso) {
  const p = String(dataIso).split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12, 0, 0, 0);
  const dia = d.getDay();
  d.setDate(d.getDate() - (dia === 0 ? 6 : dia - 1));
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function getDashboardDataInicial(tokenSessao, dataInicio, dataFim) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) throw new Error('Sessão inválida ou expirada. Faça login novamente.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dataInicio || '') || !/^\d{4}-\d{2}-\d{2}$/.test(dataFim || '')) {
    throw new Error('Período inicial inválido.');
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard_inicial_v4_' + dataInicio.replace(/-/g, '') + '_' + dataFim.replace(/-/g, '');
  const cached = cacheGetJson_(cache, cacheKey);
  if (cached) return cached;

  const cfg = getConfig_();
  const inicioSemana = inicioSemanaIso_(dataFim);
  const inicioLeitura = dataInicio < inicioSemana ? dataInicio : inicioSemana;
  const { rows: rowsSemana } = getRawRows_(inicioLeitura, dataFim);
  const limiteInicial = new Date(dataInicio + 'T00:00:00');
  const limiteFinal = new Date(dataFim + 'T23:59:59');
  let periodoEfetivoDe = dataInicio;
  let periodoEfetivoAte = dataFim;
  let rows = rowsSemana.filter(row => {
    const d = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    return d && d >= limiteInicial && d <= limiteFinal;
  });

  if (!rows.length && rowsSemana.length) {
    let ultimoDiaDisponivel = '';
    rowsSemana.forEach(row => {
      const d = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
      if (!d) return;
      const iso = Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
      if (iso <= dataFim && iso > ultimoDiaDisponivel) ultimoDiaDisponivel = iso;
    });
    if (ultimoDiaDisponivel) {
      periodoEfetivoDe = ultimoDiaDisponivel;
      periodoEfetivoAte = ultimoDiaDisponivel;
      rows = rowsSemana.filter(row => {
        const d = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
        return d && Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd') === ultimoDiaDisponivel;
      });
    }
  }

  const serieSemana = agruparSerieConversao_(
    agruparMaisRecentePorCliente_(rowsSemana),
    mapaPropostasPagas_(inicioLeitura, dataFim)
  );
  const serieMontada = montarSerieTemporal_(serieSemana, cfg);
  const serieVisivel = {};
  Object.keys(serieMontada).forEach(setor => {
    serieVisivel[setor] = serieMontada[setor].filter(item => item.data >= periodoEfetivoDe && item.data <= periodoEfetivoAte);
  });
  const clientesMaisRecentes = agruparMaisRecentePorCliente_(rows);
  const supervisores = getSupervisoresMap_();
  const telefonesNecessarios = new Set(clientesMaisRecentes.map(row => telefonePadrao10_(row['Celular'])).filter(Boolean));
  const callLogMap = getCallLogPorTelefone_(telefonesNecessarios);
  const tabulacoesMap = getMapaTabulacoes_();
  const payload = {
    modoInicial: true,
    janelaCarregada: { de: inicioLeitura, ate: dataFim },
    periodoEfetivo: {
      de: periodoEfetivoDe,
      ate: periodoEfetivoAte,
      fallback: periodoEfetivoDe !== dataInicio || periodoEfetivoAte !== dataFim
    },
    atualizadoEm: new Date().toISOString(),
    ultimaSincronizacao: getUltimaSincronizacao_(),
    funil: computeFunil_(clientesMaisRecentes),
    projecao: {
      Callcenter: computeProjecao_(serieSemana.Callcenter, cfg, periodoEfetivoDe, periodoEfetivoAte),
      API: computeProjecao_(serieSemana.API, cfg, periodoEfetivoDe, periodoEfetivoAte),
      Todos: computeProjecao_(serieSemana.Todos, cfg, periodoEfetivoDe, periodoEfetivoAte)
    },
    serie: serieVisivel,
    naoAssinados: computeNaoAssinados_(clientesMaisRecentes, supervisores, callLogMap, tabulacoesMap),
    cancelados: computeCancelados_(rows, supervisores, callLogMap, tabulacoesMap)
  };
  cachePutJson_(cache, cacheKey, payload, 600);
  return payload;
}

function getDashboardComplemento(tokenSessao, dataInicio, dataFim) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) throw new Error('Sessão inválida ou expirada. Faça login novamente.');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'dashboard_complemento_v2_' + (dataInicio || 'x') + '_' + (dataFim || 'x');
  const cached = cacheGetJson_(cache, cacheKey);
  if (cached) return cached;
  const cfg = getConfig_();
  let serieConversao = cacheGetJson_(cache, 'serie_conversao_v4');
  let opcoesExportacao = cacheGetJson_(cache, 'opcoes_exportacao_v1');
  let rowsHistoricos = null;
  if (!serieConversao || !opcoesExportacao) rowsHistoricos = getRawRows_().rows;
  if (!serieConversao) serieConversao = computeSerieConversao_(rowsHistoricos);
  if (!opcoesExportacao) {
    opcoesExportacao = montarOpcoesFiltroDeRows_(rowsHistoricos);
    cachePutJson_(cache, 'opcoes_exportacao_v1', opcoesExportacao, 600);
  }
  const payload = {
    atualizadoEm: new Date().toISOString(),
    serie: montarSerieTemporal_(serieConversao, cfg),
    opcoesExportacao: opcoesExportacao,
    projecao: {
      Callcenter: computeProjecao_(serieConversao.Callcenter, cfg, dataInicio, dataFim),
      API: computeProjecao_(serieConversao.API, cfg, dataInicio, dataFim),
      Todos: computeProjecao_(serieConversao.Todos, cfg, dataInicio, dataFim)
    }
  };
  cachePutJson_(cache, cacheKey, payload, 600);
  return payload;
}

const CANAIS_API_KEYWORDS = [
  ['Channel A', 'CHANNEL A'],
  ['Channel B', 'CHANNEL B'],
  ['App', 'APP'],
  ['Landing Page', 'LANDING']
];

function somaValSolic_(lista) {
  return lista.reduce((soma, r) => {
    const v = convertMonetaryToFloat_(r['Val. Solic.']);
    return soma + (isNaN(v) ? 0 : v);
  }, 0);
}

function computeGerencialCallcenter_(rowsTotal, rowsEfetivados) {
  const cc = rowsTotal.filter(r => classificarSetor_(r['Analista']) === 'Callcenter');
  const ccPagos = rowsEfetivados.filter(r => classificarSetor_(r['Analista']) === 'Callcenter');
  const comSeguro = cc.filter(r => {
    const v = convertMonetaryToFloat_(r['Seguro']);
    return !isNaN(v) && v > 0;
  }).length;

  return {
    canal: 'Call Center', disponivel: true,
    totalPropostas: cc.length,
    valorTotalPropostas: somaValSolic_(cc),
    qtdeEfetivados: ccPagos.length,
    valorLiberado: somaValSolic_(ccPagos),
    comSeguro: comSeguro,
    semSeguro: cc.length - comSeguro
  };
}

function computeGerencialCanalApi_(nome, keyword, rowsRaw, rowsEfetivadosApi) {
  const bate = (r) => classificarSetor_(r['Analista']) === 'API' && cleanText_(r['Analista']).toUpperCase().indexOf(keyword) !== -1;
  const total = rowsRaw.filter(bate);
  const pagos = rowsEfetivadosApi.filter(bate);
  const comSeguro = total.filter(r => {
    const v = convertMonetaryToFloat_(r['Seguro']);
    return !isNaN(v) && v > 0;
  }).length;

  return {
    canal: nome, disponivel: true,
    totalPropostas: total.length,
    valorTotalPropostas: somaValSolic_(total),
    qtdeEfetivados: pagos.length,
    valorLiberado: somaValSolic_(pagos),
    comSeguro: comSeguro,
    semSeguro: total.length - comSeguro
  };
}

function getGerencialData(tokenSessao, dataInicio, dataFim) {
  const sessao = exigirAdmin_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito a administradores.');

  const { rows: rowsTotal } = getGerTotalRows_(dataInicio, dataFim);
  const { rows: rowsEfetivados } = getGerEfetivadosRows_(dataInicio, dataFim);
  const { rows: rowsRaw } = getRawRows_(dataInicio, dataFim);
  const { rows: rowsEfetivadosApi } = getGerEfetivadosApiRows_(dataInicio, dataFim);

  return {
    atualizadoEm: new Date().toISOString(),
    canais: [
      ...CANAIS_API_KEYWORDS.map(([nome, keyword]) => computeGerencialCanalApi_(nome, keyword, rowsRaw, rowsEfetivadosApi)),
      computeGerencialCallcenter_(rowsTotal, rowsEfetivados)
    ]
  };
}

function filtroCanalGerencial_(nome) {
  if (nome === 'Call Center') return r => classificarSetor_(r['Analista']) === 'Callcenter';
  const par = CANAIS_API_KEYWORDS.find(([n]) => n === nome);
  const keyword = par ? par[1] : null;
  return r => classificarSetor_(r['Analista']) === 'API' && cleanText_(r['Analista']).toUpperCase().indexOf(keyword) !== -1;
}

function fontesPorCanalGerencial_(dataInicio, dataFim) {
  const { rows: rowsTotalCc } = getGerTotalRows_(dataInicio, dataFim);
  const { rows: rowsEfetivadosCc } = getGerEfetivadosRows_(dataInicio, dataFim);
  const { rows: rowsTotalApi } = getRawRows_(dataInicio, dataFim);
  const { rows: rowsEfetivadosApi } = getGerEfetivadosApiRows_(dataInicio, dataFim);
  const fontes = {};
  CANAIS_API_KEYWORDS.forEach(([nome]) => { fontes[nome] = { rowsTotal: rowsTotalApi, rowsEfetivados: rowsEfetivadosApi }; });
  fontes['Call Center'] = { rowsTotal: rowsTotalCc, rowsEfetivados: rowsEfetivadosCc };
  return fontes;
}

const NOMES_CANAIS_GERENCIAL = [...CANAIS_API_KEYWORDS.map(([nome]) => nome), 'Call Center'];

function computeCanalPorHora_(nome, filtro, rowsTotalDia, rowsEfetivadosDia) {
  const vazio = () => ({ totalPropostas: 0, valorTotalPropostas: 0, comSeguro: 0, semSeguro: 0, qtdeEfetivados: 0, valorLiberado: 0 });
  const horas = {};
  for (let h = 0; h < 24; h++) horas[h] = vazio();

  rowsTotalDia.forEach(row => {
    if (!filtro(row)) return;
    const dt = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (!dt) return;
    const bucket = horas[dt.getHours()];
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    bucket.totalPropostas++;
    bucket.valorTotalPropostas += isNaN(valor) ? 0 : valor;
    const seguro = convertMonetaryToFloat_(row['Seguro']);
    if (!isNaN(seguro) && seguro > 0) bucket.comSeguro++; else bucket.semSeguro++;
  });
  rowsEfetivadosDia.forEach(row => {
    if (!filtro(row)) return;
    const dt = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (!dt) return;
    const bucket = horas[dt.getHours()];
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    bucket.qtdeEfetivados++;
    bucket.valorLiberado += isNaN(valor) ? 0 : valor;
  });

  const resultado = {};
  Object.keys(horas).forEach(h => { resultado[h] = Object.assign({ canal: nome, disponivel: true }, horas[h]); });
  return resultado;
}

function getGerencialDataPorHora(tokenSessao, dataIso) {
  const sessao = exigirAdmin_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito a administradores.');
  if (!dataIso) throw new Error('Informe o dia.');

  const fontes = fontesPorCanalGerencial_(dataIso, dataIso);
  const porHora = {};
  NOMES_CANAIS_GERENCIAL.forEach(nome => {
    porHora[nome] = computeCanalPorHora_(nome, filtroCanalGerencial_(nome), fontes[nome].rowsTotal, fontes[nome].rowsEfetivados);
  });

  return {
    atualizadoEm: new Date().toISOString(),
    dia: dataIso,
    porHora: porHora
  };
}

function computeCanalSerie_(nome, filtro, rowsTotal, rowsEfetivados, chaveFn, rotuloFn) {
  const buckets = {};
  const pegar = (chave) => {
    if (!buckets[chave]) buckets[chave] = { totalPropostas: 0, valorTotalPropostas: 0, comSeguro: 0, semSeguro: 0, qtdeEfetivados: 0, valorLiberado: 0 };
    return buckets[chave];
  };

  rowsTotal.forEach(row => {
    if (!filtro(row)) return;
    const dt = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (!dt) return;
    const chave = chaveFn(dt);
    const b = pegar(chave);
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    b.totalPropostas++;
    b.valorTotalPropostas += isNaN(valor) ? 0 : valor;
    const seguro = convertMonetaryToFloat_(row['Seguro']);
    if (!isNaN(seguro) && seguro > 0) b.comSeguro++; else b.semSeguro++;
  });
  rowsEfetivados.forEach(row => {
    if (!filtro(row)) return;
    const dt = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (!dt) return;
    const chave = chaveFn(dt);
    const b = pegar(chave);
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    b.qtdeEfetivados++;
    b.valorLiberado += isNaN(valor) ? 0 : valor;
  });

  return Object.keys(buckets).sort().map(chave =>
    Object.assign({ chave: chave, rotulo: rotuloFn(chave), canal: nome, disponivel: true }, buckets[chave])
  );
}

function chaveDia_(dt) {
  return dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + String(dt.getDate()).padStart(2, '0');
}
function chaveMes_(dt) {
  return dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
}

function serieTodosCanais_(dataInicio, dataFim, chaveFn, rotuloFn) {
  const fontes = fontesPorCanalGerencial_(dataInicio, dataFim);
  const serie = {};
  NOMES_CANAIS_GERENCIAL.forEach(nome => {
    serie[nome] = computeCanalSerie_(nome, filtroCanalGerencial_(nome), fontes[nome].rowsTotal, fontes[nome].rowsEfetivados, chaveFn, rotuloFn);
  });
  return serie;
}

function getGerencialSerieDiaria(tokenSessao, dataInicio, dataFim) {
  const sessao = exigirAdmin_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito a administradores.');

  return {
    atualizadoEm: new Date().toISOString(),
    serie: serieTodosCanais_(dataInicio, dataFim, chaveDia_, chave => chave.slice(8, 10) + '/' + chave.slice(5, 7))
  };
}

function getGerencialSerieMensal(tokenSessao, dataInicio, dataFim) {
  const sessao = exigirAdmin_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito a administradores.');

  const MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  return {
    atualizadoEm: new Date().toISOString(),
    serie: serieTodosCanais_(dataInicio, dataFim, chaveMes_, chave => MESES_ABREV[Number(chave.slice(5, 7)) - 1] + '/' + chave.slice(2, 4))
  };
}

function getGerencialBootstrap(tokenSessao, periodos) {
  const sessao = exigirAdmin_(tokenSessao);
  if (!sessao) throw new Error('Acesso restrito a administradores.');
  const datasDoBootstrap = periodos && [
    periodos.hoje,
    periodos.mensal && periodos.mensal.de, periodos.mensal && periodos.mensal.ate,
    periodos.mensalAnterior && periodos.mensalAnterior.de, periodos.mensalAnterior && periodos.mensalAnterior.ate,
    periodos.diaria && periodos.diaria.de, periodos.diaria && periodos.diaria.ate,
    periodos.diariaAnterior && periodos.diariaAnterior.de, periodos.diariaAnterior && periodos.diariaAnterior.ate,
    periodos.serieDiaria && periodos.serieDiaria.de, periodos.serieDiaria && periodos.serieDiaria.ate,
    periodos.serieMensal && periodos.serieMensal.de, periodos.serieMensal && periodos.serieMensal.ate
  ];
  if (!datasDoBootstrap || datasDoBootstrap.some(data => !/^\d{4}-\d{2}-\d{2}$/.test(data || ''))) {
    throw new Error('Períodos da visão gerencial inválidos.');
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'gerencial_bootstrap_v2_' + datasDoBootstrap.join('_').replace(/-/g, '');
  const cached = cacheGetJson_(cache, cacheKey);
  if (cached) return cached;

  const inicio = [periodos.mensalAnterior.de, periodos.diariaAnterior.de,
    periodos.serieDiaria.de, periodos.serieMensal.de].sort()[0];
  const fim = [periodos.mensal.ate, periodos.diaria.ate,
    periodos.serieDiaria.ate, periodos.serieMensal.ate].sort().slice(-1)[0];
  const fontes = {
    total: getGerTotalRows_(inicio, fim).rows,
    efetivados: getGerEfetivadosRows_(inicio, fim).rows,
    raw: getRawRows_(inicio, fim).rows,
    efetivadosApi: getGerEfetivadosApiRows_(inicio, fim).rows
  };
  const dataDaLinha = new WeakMap();
  function diaDaLinha(row) {
    if (!dataDaLinha.has(row)) {
      const data = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
      dataDaLinha.set(row, data ? chaveDia_(data) : '');
    }
    return dataDaLinha.get(row);
  }
  function noPeriodo(lista, de, ate) {
    return lista.filter(row => {
      const dia = diaDaLinha(row);
      return dia && dia >= de && dia <= ate;
    });
  }
  function recorte(de, ate) {
    return {
      total: noPeriodo(fontes.total, de, ate),
      efetivados: noPeriodo(fontes.efetivados, de, ate),
      raw: noPeriodo(fontes.raw, de, ate),
      efetivadosApi: noPeriodo(fontes.efetivadosApi, de, ate)
    };
  }
  function canais(dados) {
    return [
      ...CANAIS_API_KEYWORDS.map(([nome, keyword]) =>
        computeGerencialCanalApi_(nome, keyword, dados.raw, dados.efetivadosApi)),
      computeGerencialCallcenter_(dados.total, dados.efetivados)
    ];
  }
  function fontesDe(dados) {
    const porCanal = {};
    CANAIS_API_KEYWORDS.forEach(([nome]) => {
      porCanal[nome] = { rowsTotal: dados.raw, rowsEfetivados: dados.efetivadosApi };
    });
    porCanal['Call Center'] = { rowsTotal: dados.total, rowsEfetivados: dados.efetivados };
    return porCanal;
  }
  function serie(dados, chaveFn, rotuloFn) {
    const porCanal = fontesDe(dados);
    const resultado = {};
    NOMES_CANAIS_GERENCIAL.forEach(nome => {
      const fonte = porCanal[nome];
      resultado[nome] = computeCanalSerie_(nome, filtroCanalGerencial_(nome),
        fonte.rowsTotal, fonte.rowsEfetivados, chaveFn, rotuloFn);
    });
    return resultado;
  }

  const mensal = recorte(periodos.mensal.de, periodos.mensal.ate);
  const mensalAnterior = recorte(periodos.mensalAnterior.de, periodos.mensalAnterior.ate);
  const diaria = recorte(periodos.diaria.de, periodos.diaria.ate);
  const diariaAnterior = recorte(periodos.diariaAnterior.de, periodos.diariaAnterior.ate);
  const horario = fontesDe(diaria);
  const porHora = {};
  NOMES_CANAIS_GERENCIAL.forEach(nome => {
    const fonte = horario[nome];
    porHora[nome] = computeCanalPorHora_(nome, filtroCanalGerencial_(nome),
      fonte.rowsTotal, fonte.rowsEfetivados);
  });

  const MESES_ABREV = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const resultado = {
    atualizadoEm: new Date().toISOString(),
    mensal: { canais: canais(mensal) },
    mensalAnterior: { canais: canais(mensalAnterior) },
    diaria: { canais: canais(diaria) },
    diariaAnterior: { canais: canais(diariaAnterior) },
    hora: { dia: periodos.hoje, porHora: porHora },
    serieDiaria: serie(recorte(periodos.serieDiaria.de, periodos.serieDiaria.ate),
      chaveDia_, chave => chave.slice(8, 10) + '/' + chave.slice(5, 7)),
    serieMensal: serie(recorte(periodos.serieMensal.de, periodos.serieMensal.ate),
      chaveMes_, chave => MESES_ABREV[Number(chave.slice(5, 7)) - 1] + '/' + chave.slice(2, 4))
  };
  cachePutJson_(cache, cacheKey, resultado, 600);
  return resultado;
}

const CABECALHO_EXPORT = ['Data', 'Proposta', 'Cliente', 'Celular', 'Analista', 'Setor', 'Correspondente', 'Status', 'Motivo Cancelamento', 'Valor', 'Supervisor'];
const COR_CABECALHO_XLSX = '#1B2A4A';
const COR_FONTE_CABECALHO_XLSX = '#FFFFFF';

function chaveCacheExportacao_(filtros) {
  const normalizado = {
    setor: filtros.setor || 'Todos',
    estagios: (Array.isArray(filtros.estagios) ? filtros.estagios.slice() : []).sort(),
    motivos: (Array.isArray(filtros.motivos) ? filtros.motivos.slice() : []).sort(),
    dataInicio: filtros.dataInicio || '', dataFim: filtros.dataFim || '',
    cliente: String(filtros.cliente || '').trim().toLowerCase(),
    correspondente: String(filtros.correspondente || '').trim().toLowerCase(),
    formato: filtros.formato === 'csv' ? 'csv' : 'xlsx'
  };
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(normalizado));
  return 'export_v1_' + bytes.slice(0, 16).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function montarOpcoesFiltroDeRows_(rows) {
  const correspondentes = new Set();
  const clientes = new Set();
  rows.forEach(row => {
    const corr = cleanText_(row['Correspondente']);
    if (corr) correspondentes.add(corr);
    const cli = cleanText_(row['Cliente']);
    if (cli) clientes.add(cli);
  });
  return {
    correspondentes: Array.from(correspondentes).sort(),
    clientes: Array.from(clientes).sort()
  };
}

function obterOpcoesFiltro(tokenSessao) {
  if (!exigirSessaoValida_(tokenSessao)) throw new Error('Sessão inválida ou expirada. Faça login novamente.');
  const cache = CacheService.getScriptCache();
  const cacheado = cacheGetJson_(cache, 'opcoes_exportacao_v1');
  if (cacheado) return cacheado;
  const { rows } = getRawRows_();
  const resultado = montarOpcoesFiltroDeRows_(rows);
  cachePutJson_(cache, 'opcoes_exportacao_v1', resultado, 600);
  return resultado;
}

function linhasFiltradas_(filtros) {
  const { rows } = getRawRows_(filtros.dataInicio || null, filtros.dataFim || null);
  const supervisores = getSupervisoresMap_();
  const setor = filtros.setor || 'Todos';
  const estagiosSel = Array.isArray(filtros.estagios) && filtros.estagios.length ? filtros.estagios : null;
  const motivosSel = Array.isArray(filtros.motivos) && filtros.motivos.length ? filtros.motivos : null;
  const cliente = String(filtros.cliente || '').trim().toLowerCase();
  const correspondente = String(filtros.correspondente || '').trim().toLowerCase();
  const dataInicio = filtros.dataInicio ? new Date(filtros.dataInicio + 'T00:00:00') : null;
  const dataFim = filtros.dataFim ? new Date(filtros.dataFim + 'T23:59:59') : null;

  const linhas = [];
  rows.forEach(row => {
    const setorRow = classificarSetor_(row['Analista']);
    if (setor !== 'Todos' && setorRow !== setor) return;

    const estagio = classificarEstagioAtual_(row['Status Admin.']);
    if (estagiosSel && estagiosSel.indexOf(estagio) === -1) return;

    let motivoFiltro = '', motivoExport = '';
    if (estagio === ESTAGIO_CANCELADO) {
      const bruto = cleanText_(row['Motivo_Cancelamento']);
      const temMotivo = bruto && bruto.toUpperCase() !== '#N/D';
      motivoExport = temMotivo ? normalizarMotivo_(bruto) : '';
      motivoFiltro = temMotivo ? motivoExport : 'CANCELADO';
      if (motivosSel && motivosSel.indexOf(motivoFiltro) === -1) return;
    }

    if (cliente && cleanText_(row['Cliente']).toLowerCase().indexOf(cliente) === -1) return;
    if (correspondente && cleanText_(row['Correspondente']).toLowerCase().indexOf(correspondente) === -1) return;

    const dataProposta = parseDataHoraProposta_(row['Solic. Proposta'], row['Hora Solic. Proposta']);
    if (dataInicio && (!dataProposta || dataProposta < dataInicio)) return;
    if (dataFim && (!dataProposta || dataProposta > dataFim)) return;

    const analista = cleanText_(row['Analista']);
    const valor = convertMonetaryToFloat_(row['Val. Solic.']);
    linhas.push([
      dataProposta ? Utilities.formatDate(dataProposta, 'America/Sao_Paulo', 'dd/MM/yyyy') : '',
      row['N° Proposta'], row['Cliente'], row['Celular'], analista, setorRow, row['Correspondente'],
      estagio, motivoExport, isNaN(valor) ? '' : valor, supervisores[analista] || ''
    ]);
  });
  return linhas;
}

function paraCsv_(headers, linhas) {
  const escape = (v) => {
    const s = v === null || v === undefined ? '' : String(v);
    return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const linhasCsv = [headers.map(escape).join(';')];
  linhas.forEach(l => linhasCsv.push(l.map(escape).join(';')));
  return linhasCsv.join('\n');
}

function gerarArquivoExport_(headers, linhas, formato, nomeBase) {
  if (formato === 'csv') {
    const csv = paraCsv_(headers, linhas);
    const blob = Utilities.newBlob('﻿' + csv, 'text/csv', nomeBase + '.csv');
    return { base64: Utilities.base64Encode(blob.getBytes()), mime: 'text/csv', nome: nomeBase + '.csv' };
  }

  const ssTemp = SpreadsheetApp.create(nomeBase);
  let bytes;
  try {
    const sh = ssTemp.getActiveSheet();
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground(COR_CABECALHO_XLSX).setFontColor(COR_FONTE_CABECALHO_XLSX).setFontWeight('bold');
    if (linhas.length > 0) {
      if (headers[0] === 'Data') sh.getRange(2, 1, linhas.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, linhas.length, headers.length).setValues(linhas);
    }
    SpreadsheetApp.flush();
    const url = ssTemp.getUrl().replace(/\/edit.*$/, '') + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    bytes = resp.getBlob().getBytes();
  } finally {
    try { DriveApp.getFileById(ssTemp.getId()).setTrashed(true); } catch (e) {}
  }

  return {
    base64: Utilities.base64Encode(bytes),
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    nome: nomeBase + '.xlsx'
  };
}

function rotuloExportacao_(estagiosSel) {
  if (estagiosSel && estagiosSel.length) {
    if (estagiosSel.length === 1 && estagiosSel[0] === ESTAGIO_CANCELADO) return 'Cancelados';
    if (estagiosSel.every(e => STATUS_NAO_ASSINADOS_ORDEM.indexOf(e) !== -1)) return 'Casos Parados';
  }
  return 'Exportação';
}

function exportarComFiltros(tokenSessao, filtros) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) throw new Error('Sessão inválida ou expirada. Faça login novamente.');

  filtros = filtros || {};
  const cache = CacheService.getScriptCache();
  const cacheKey = chaveCacheExportacao_(filtros);
  const arquivoCacheado = cacheGetJson_(cache, cacheKey);
  if (arquivoCacheado) {
    if (!filtros._preparacao) registrarAcessoAsync(sessao.login, sessao.papel, true, 'export:cache:' + (filtros.formato === 'csv' ? 'csv' : 'xlsx'));
    return arquivoCacheado;
  }
  const linhas = linhasFiltradas_(filtros);
  const formato = filtros.formato === 'csv' ? 'csv' : 'xlsx';
  const rotulo = rotuloExportacao_(Array.isArray(filtros.estagios) ? filtros.estagios : null);
  const nomeBase = 'AssinaAI - ' + rotulo + ' - ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd-MM-yyyy');
  const arquivo = gerarArquivoExport_(CABECALHO_EXPORT, linhas, formato, nomeBase);
  cachePutJson_(cache, cacheKey, arquivo, 300);
  if (!filtros._preparacao) registrarAcessoAsync(sessao.login, sessao.papel, true, 'export:' + formato + ':' + linhas.length + 'linhas');
  return arquivo;
}

function prepararExportacaoRapida(tokenSessao, deslocamentoDias) {
  if (!exigirSessaoValida_(tokenSessao)) throw new Error('Sessão inválida ou expirada. Faça login novamente.');
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  const deslocamento = Number(deslocamentoDias) === -1 ? -1 : 0;
  const data = deslocarDataIso_(hoje, deslocamento);
  exportarComFiltros(tokenSessao, {
    setor: 'Todos', estagios: [], motivos: [], dataInicio: data, dataFim: data,
    cliente: '', correspondente: '', formato: 'xlsx', _preparacao: true
  });
  return { ok: true, data: data, preparadoEm: new Date().toISOString() };
}

function prepararExportacoesRapidas(tokenSessao) {
  const hoje = prepararExportacaoRapida(tokenSessao, 0);
  const ontem = prepararExportacaoRapida(tokenSessao, -1);
  return { ok: true, datas: [hoje.data, ontem.data], preparadoEm: new Date().toISOString() };
}

function exportarBloco(tokenSessao, bloco, formato, dataInicio, dataFim) {
  const sessao = exigirSessaoValida_(tokenSessao);
  if (!sessao) throw new Error('Sessão inválida ou expirada. Faça login novamente.');

  const { rows } = getRawRows_(dataInicio, dataFim);
  const supervisores = getSupervisoresMap_();
  const callLogMap = getCallLogPorTelefone_();
  const tabulacoesMap = getMapaTabulacoes_();

  let headers, linhas, rotulo;
  if (bloco === 'cancelados') {
    rotulo = 'Cancelados';
    headers = ['Data', 'Motivo', 'Proposta', 'Cliente', 'Celular', 'Analista', 'Correspondente', 'Valor', 'Supervisor', 'Retornou ao Cliente', 'Tentativas de Retorno', 'Tentativas Bem Sucedidas', 'Retorno Tentativa 1º', 'Retorno Tentativa 2º'];
    linhas = computeCancelados_(rows, supervisores, callLogMap, tabulacoesMap).detalhe.map(d => [d.dataProposta, d.motivo, d.proposta, d.cliente, d.celular, d.analista, d.correspondente, d.valor, d.supervisor, d.retornouCliente ? 'Sim' : 'Não', d.tentativasRetorno, d.tentativasBemSucedidas, d.tabulacaoTentativa1, d.tabulacaoTentativa2]);
  } else {
    rotulo = 'Casos Parados';
    const clientesMaisRecentes = agruparMaisRecentePorCliente_(rows);
    headers = ['Data', 'Validação', 'Proposta', 'Cliente', 'Celular', 'Analista', 'Correspondente', 'Valor', 'Supervisor', 'Retornou ao Cliente', 'Tentativas de Retorno', 'Tentativas Bem Sucedidas', 'Retorno Tentativa 1º', 'Retorno Tentativa 2º'];
    linhas = computeNaoAssinados_(clientesMaisRecentes, supervisores, callLogMap, tabulacoesMap).detalhe.map(d => [d.dataProposta, d.validacao, d.proposta, d.cliente, d.celular, d.analista, d.correspondente, d.valor, d.supervisor, d.retornouCliente ? 'Sim' : 'Não', d.tentativasRetorno, d.tentativasBemSucedidas, d.tabulacaoTentativa1, d.tabulacaoTentativa2]);
  }
  const nomeBase = 'AssinaAI - ' + rotulo + ' - ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd-MM-yyyy');
  registrarAcessoAsync(sessao.login, sessao.papel, true, 'export-bloco:' + bloco + ':' + formato);
  return gerarArquivoExport_(headers, linhas, formato === 'csv' ? 'csv' : 'xlsx', nomeBase);
}

function inicializarPlanilha() {
  const ss = planilha_();
  const definicoes = {
    [SHEET_RAW]: [],
    [SHEET_CONFIG]: ['Chave', 'Valor'],
    [SHEET_SNAPSHOT]: ['Data', 'Setor', 'Estagio', 'Qtd', 'Valor'],
    [SHEET_SUPERVISORES]: ['Analista', 'Supervisor'],
    [SHEET_ATRIBUICOES]: ['Proposta', 'Analista', 'DataAtribuicao', 'DataConversao'],
    [SHEET_USUARIOS]: ['Login', 'SenhaHash', 'Salt', 'Papel', 'Ativo', 'CriadoEm'],
    [SHEET_TENTATIVAS]: ['Chave', 'Tentativas', 'UltimaTentativa', 'BloqueadoAte'],
    [SHEET_LOG_ACESSOS]: ['DataHora', 'Login', 'Papel', 'Sucesso', 'Detalhe'],
    [SHEET_CALL_LOG]: []
  };
  Object.keys(definicoes).forEach(nome => {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    if (sh.getLastRow() === 0 && definicoes[nome].length > 0) sh.getRange(1, 1, 1, definicoes[nome].length).setValues([definicoes[nome]]);
  });

  const cfgSheet = ss.getSheetByName(SHEET_CONFIG);
  if (cfgSheet.getLastRow() < 2) {
    cfgSheet.getRange(2, 1, 5, 2).setValues([
      ['Meta_Conversao_Diaria', 20],
      ['Limiar_Alerta_Vermelho', 0.15],
      ['Dias_Media_Movel', 7],
      ['Horario_Corte_Dia', '23:59'],
      ['Ultima_Sincronizacao', '']
    ]);
  }

  definirUsuariosPadrao();
}

function definirUsuariosPadrao() {
  const aba = planilha_().getSheetByName(SHEET_USUARIOS);
  const seeds = [
    { login: 'admin.demo', papel: 'Admin' },
    { login: 'Master', papel: 'Admin' },
    { login: 'supervisor.demo', papel: 'Supervisor' },
    { login: 'operator.demo', papel: 'Operador' }
  ];
  const props = PropertiesService.getScriptProperties();
  const dados = aba.getDataRange().getValues();
  const prontos = [];

  seeds.forEach(seed => {
    const senha = props.getProperty('ASSINA_AI_SEED_' + seed.login.toUpperCase());
    if (!senha) { Logger.log('Sem senha inicial para ' + seed.login + ' : conta pulada.'); return; }
    const salt = Utilities.getUuid();
    const hash = hashSenha_(senha, salt);
    let linha = -1;
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).toLowerCase() === seed.login.toLowerCase()) { linha = i + 1; break; }
    }
    if (linha === -1) {
      aba.appendRow([seed.login, hash, salt, seed.papel, true, new Date()]);
    } else {
      aba.getRange(linha, 2, 1, 3).setValues([[hash, salt, seed.papel]]);
    }
    prontos.push(seed.login + ' (' + seed.papel + ')');
  });

  try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {}
  Logger.log('Usuários prontos: ' + (prontos.join(', ') || 'nenhum'));
  return prontos;
}

function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncSnapshotDiario_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncSnapshotDiario_').timeBased().everyMinutes(15).create();
}
