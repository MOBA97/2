/*** === КОНФИГ (всё внутри одной таблицы) === ***/
const USERS_SHEET   = 'Users';       // логины
const SECTIONS_SHEET= 'Sections';    // alias | title | description | enabled | salt | passhash
const AUDIT_SHEET   = 'Audit';       // журнал

// Алиасы (ключ -> имя листа в ЭТОЙ таблице)
const ALIASES_SHEET = {
  'клиенты':'Клиенты',
  'заказы':'Заказы',
  'склад':'Склад',
  'брак':'Брак',
  'возвраты':'Возвраты',
  'сервисы':'Сервисы',
  'оплаты':'Оплаты',
};

// Подписи по умолчанию
const ALIAS_LABEL = {
  'клиенты':'Клиенты',
  'заказы':'Заказы',
  'склад':'Склад',
  'брак':'Брак',
  'возвраты':'Возвраты',
  'сервисы':'Сервисы',
  'оплаты':'Оплаты',
};

// Домашний лист (Моя панель)
const ROLE_HOME_SHEET = { admin:'Главная', editor:'Клиенты', viewer:'Склад' };

// Синонимы (поддержка старых английских алиасов)
const ALIAS_SYNONYMS = { clients:'клиенты', orders:'заказы', stock:'склад' };
function normAlias(a){ a=String(a||'').trim().toLowerCase(); return ALIAS_SYNONYMS[a] || a; }

/*** === УТИЛИТЫ === ***/
function doGet(){
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Вход в CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sha256Hex_(str){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  return raw.map(b => (b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function sheetUrlByName_(name){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Нет листа "${name}"`);
  return `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${sh.getSheetId()}`;
}

function countRows_(sheetName){
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return 0;
  const last = sh.getLastRow(); if (last<2) return 0;
  const a = sh.getRange(2,1,last-1,1).getValues();
  return a.filter(r => String(r[0]).trim()!=='').length;
}

function getOrCreateSheet_(name){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function logEvent_(kind, login, detail){
  const sh = getOrCreateSheet_(AUDIT_SHEET);
  if (sh.getLastRow()===0) sh.appendRow(['ts','kind','login','detail']);
  sh.appendRow([new Date(), kind||'', login||'', detail||'']);
}

/*** === USERS === ***/
function loadUsers_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
  if (!sh) throw new Error(`Не найден лист "${USERS_SHEET}"`);
  const rng = sh.getDataRange().getValues();
  const head = rng.shift().map(h=>String(h).trim().toLowerCase());
  const idx = n => head.indexOf(n);
  ['login','salt','passhash','role','aliases','enabled'].forEach(n=>{
    if (idx(n)<0) throw new Error(`В листе Users нет колонки: ${n}`);
  });
  const map = new Map();
  rng.forEach(r=>{
    const login = String(r[idx('login')]||'').trim();
    if (!login) return;
    const role  = String(r[idx('role')]||'').trim().toLowerCase();
    const rec = {
      salt:     String(r[idx('salt')]||'').trim(),
      passhash: String(r[idx('passhash')]||'').trim().toLowerCase(),
      role,
      aliases:  String(r[idx('aliases')]||'').trim(),
      enabled:  String(r[idx('enabled')]).toUpperCase()==='TRUE',
      display:  head.includes('display') ? String(r[idx('display')]).trim() : login,
      home:     head.includes('home')    ? String(r[idx('home')]).trim()    : sheetUrlByName_(ROLE_HOME_SHEET[role] || ROLE_HOME_SHEET.viewer),
      can_edit_all: head.includes('can_edit_all')
        ? (String(r[idx('can_edit_all')]).toUpperCase()==='TRUE')
        : (role==='admin'),
    };
    map.set(login.toLowerCase(), rec);
  });
  return map;
}

/*** === РАЗДЕЛЫ === ***/
function readSections_(aliasesArr){
  const meta = {};
  const sh = SpreadsheetApp.getActive().getSheetByName(SECTIONS_SHEET);
  if (sh){
    const v = sh.getDataRange().getValues();
    const head = v.shift().map(h=>String(h).trim().toLowerCase());
    const ix = n => head.indexOf(n);
    v.forEach(r=>{
      const alias = normAlias(r[ix('alias')]);
      if (!alias) return;
      meta[alias] = {
        title: String(r[ix('title')]||'').trim(),
        description: String(r[ix('description')]||'').trim(),
        enabled: String(r[ix('enabled')]).toUpperCase()==='TRUE',
        salt: String(r[ix('salt')]||'').trim(),
        passhash: String(r[ix('passhash')]||'').trim().toLowerCase(),
      };
    });
  }

  const out = [];
  (aliasesArr||[]).forEach(_a=>{
    const a = normAlias(_a);
    if (!(a in ALIASES_SHEET)) return;
    const m = meta[a] || {};
    if ('enabled' in m && !m.enabled) return;

    const sheetName = ALIASES_SHEET[a];
    const url = sheetUrlByName_(sheetName);
    const count = countRows_(sheetName);
    const title = m.title || ALIAS_LABEL[a] || a;
    const description = m.description || '';
    const protected_ = Boolean(m.salt && m.passhash);

    out.push({ alias:a, title, description, url, count, protected: protected_ });
  });
  return out;
}

function verifySectionPassword(alias, password){
  alias = normAlias(alias);
  const sh = SpreadsheetApp.getActive().getSheetByName(SECTIONS_SHEET);
  if (!sh) return { ok:false, msg:'Раздел не настроен' };
  const v = sh.getDataRange().getValues();
  const head = v.shift().map(h=>String(h).trim().toLowerCase());
  const ia=head.indexOf('alias'), is=head.indexOf('salt'), ih=head.indexOf('passhash'), ie=head.indexOf('enabled');
  for (let i=0;i<v.length;i++){
    const r=v[i];
    if (normAlias(r[ia])===alias){
      if (String(r[ie]).toUpperCase()!=='TRUE') return { ok:false, msg:'Раздел отключён' };
      const salt=String(r[is]||'').trim(), ph=String(r[ih]||'').trim().toLowerCase();
      if (!salt || !ph) return { ok:true }; // пароль не требуется
      const calc = sha256Hex_(salt + password);
      return { ok: calc===ph, msg: calc===ph ? '' : 'Неверный пароль раздела' };
    }
  }
  return { ok:false, msg:'Раздел не найден' };
}

/*** === АВТОРИЗАЦИЯ === ***/
function login(login, password){
  try{
    login = String(login||'').trim().toLowerCase();
    if (!login || !password) return { ok:false, msg:'Пустой логин или пароль' };

    const users = loadUsers_();
    const u = users.get(login);
    if (!u) return { ok:false, msg:'Пользователь не найден' };
    if (!u.enabled) return { ok:false, msg:'Учётная запись отключена' };

    const calc = sha256Hex_(u.salt + password);
    if (calc !== u.passhash) return { ok:false, msg:'Неверный пароль' };

    const aliases = u.aliases ? u.aliases.split(/[,\s;]+/).map(normAlias).filter(Boolean) : [];
    const sections = readSections_(aliases);

    const stamp = new Date().toISOString();
    const token = sha256Hex_([login, u.role, stamp, Session.getTemporaryActiveUserKey()||''].join('|'));

    logEvent_('login_ok', login, '');
    return {
      ok:true,
      user:{ login, display:u.display, role:u.role, can_edit_all:u.can_edit_all, home:u.home, aliases },
      sections,
      token
    };
  }catch(e){
    logEvent_('login_fail', login, e.message);
    return { ok:false, msg:'Ошибка сервера: '+e.message };
  }
}

function changePassword(login, oldPassword, newPassword){
  login = String(login||'').trim().toLowerCase();
  if (!login || !oldPassword || !newPassword) return {ok:false,msg:'Пустые поля'};

  const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
  const v = sh.getDataRange().getValues();
  const head = v[0].map(h=>String(h).trim().toLowerCase());
  const iL=head.indexOf('login'), iS=head.indexOf('salt'), iH=head.indexOf('passhash'), iE=head.indexOf('enabled');

  for (let r=1;r<v.length;r++){
    const row=v[r];
    if (String(row[iL]).trim().toLowerCase()!==login) continue;
    if (String(row[iE]).toUpperCase()!=='TRUE') return {ok:false,msg:'Учётка отключена'};
    const calc = sha256Hex_(String(row[iS]) + oldPassword);
    if (String(row[iH]).toLowerCase()!==calc) return {ok:false,msg:'Старый пароль неверен'};
    const salt = Utilities.getUuid().replace(/-/g,'').slice(0,16);
    const passhash = sha256Hex_(salt + newPassword);
    sh.getRange(r+1,iS+1).setValue(salt);
    sh.getRange(r+1,iH+1).setValue(passhash);
    logEvent_('password_changed', login, '');
    return {ok:true,msg:'Пароль обновлён'};
  }
  return {ok:false,msg:'Пользователь не найден'};
}

/*** === АДМИН === ***/
function assertAdmin_(requestorLogin){
  const u = loadUsers_().get(String(requestorLogin||'').toLowerCase());
  if (!u || !u.enabled) throw new Error('Нет доступа');
  if (!(u.role==='admin' || u.can_edit_all===true)) throw new Error('Требуется роль admin');
  return u;
}

function getDashboard(requestorLogin){
  try{
    assertAdmin_(requestorLogin);
    const stats = {};
    Object.entries(ALIASES_SHEET).forEach(([key, sheetName])=>{
      stats[key] = countRows_(sheetName);
    });

    const users = [];
    const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    const v = sh.getDataRange().getValues();
    const head = v.shift().map(h=>String(h).trim().toLowerCase());
    const idx = n => head.indexOf(n);
    v.forEach(r=>{
      const rec = {
        login: String(r[idx('login')]).trim(),
        role:  String(r[idx('role')]).trim(),
        aliases: String(r[idx('aliases')]).trim(),
        enabled: String(r[idx('enabled')]).toUpperCase()==='TRUE',
        display: head.includes('display') ? String(r[idx('display')]).trim() : '',
        home:    head.includes('home')    ? String(r[idx('home')]).trim()    : '',
        can_edit_all: head.includes('can_edit_all')
          ? (String(r[idx('can_edit_all')]).toUpperCase()==='TRUE')
          : (String(r[idx('role')]).trim().toLowerCase()==='admin')
      };
      if (rec.login) users.push(rec);
    });

    return {ok:true, stats, users};
  }catch(e){ return {ok:false, msg:e.message}; }
}

function adminUpsertUser(requestorLogin, payload){
  try{
    assertAdmin_(requestorLogin);
    const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    const v = sh.getDataRange().getValues();
    const head = v[0].map(h=>String(h).trim());
    const M = new Map(head.map((h,i)=>[h.toLowerCase(),i]));
    const need = ['login','salt','passhash','role','aliases','enabled'];
    need.forEach(n=>{ if(!M.has(n)) throw new Error('Нет колонки '+n); });
    const col = n => M.get(n);

    let row = -1;
    for (let r=1;r<v.length;r++){
      if (String(v[r][col('login')]).trim().toLowerCase()===String(payload.login).toLowerCase()){ row=r+1; break; }
    }
    if (row<0) row = sh.getLastRow()+1;

    sh.getRange(row,col('login')+1).setValue(payload.login);
    sh.getRange(row,col('role')+1).setValue(payload.role||'viewer');
    sh.getRange(row,col('aliases')+1).setValue(payload.aliases||'');
    sh.getRange(row,col('enabled')+1).setValue(Boolean(payload.enabled));
    if (M.has('display')) sh.getRange(row,col('display')+1).setValue(payload.display||'');
    if (M.has('home'))    sh.getRange(row,col('home')+1).setValue(payload.home||'');
    if (M.has('can_edit_all')) sh.getRange(row,col('can_edit_all')+1).setValue(Boolean(payload.can_edit_all));

    if (payload.newPassword){
      const salt = Utilities.getUuid().replace(/-/g,'').slice(0,16);
      const passhash = sha256Hex_(salt + payload.newPassword);
      sh.getRange(row,col('salt')+1).setValue(salt);
      sh.getRange(row,col('passhash')+1).setValue(passhash);
    } else if (v.length===1) {
      sh.getRange(row,col('enabled')+1).setValue(false);
    }

    logEvent_('user_upsert', requestorLogin, payload.login);
    return {ok:true};
  }catch(e){ return {ok:false, msg:e.message}; }
}

function adminToggleUser(requestorLogin, login, enabled){
  try{
    assertAdmin_(requestorLogin);
    const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    const v = sh.getDataRange().getValues();
    const head = v[0].map(h=>String(h).trim().toLowerCase());
    const iL=head.indexOf('login'), iE=head.indexOf('enabled');
    for (let r=1;r<v.length;r++){
      if (String(v[r][iL]).trim().toLowerCase()===String(login).toLowerCase()){
        sh.getRange(r+1, iE+1).setValue(Boolean(enabled));
        logEvent_('user_toggle', requestorLogin, login+' -> '+enabled);
        return {ok:true};
      }
    }
    return {ok:false, msg:'Пользователь не найден'};
  }catch(e){ return {ok:false, msg:e.message}; }
}

function adminResetPassword(requestorLogin, login, newPassword){
  try{
    assertAdmin_(requestorLogin);
    const sh = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    const v = sh.getDataRange().getValues();
    const head = v[0].map(h=>String(h).trim().toLowerCase());
    const iL=head.indexOf('login'), iS=head.indexOf('salt'), iH=head.indexOf('passhash');
    for (let r=1;r<v.length;r++){
      if (String(v[r][iL]).trim().toLowerCase()===String(login).toLowerCase()){
        const salt = Utilities.getUuid().replace(/-/g,'').slice(0,16);
        const passhash = sha256Hex_(salt + newPassword);
        sh.getRange(r+1,iS+1).setValue(salt);
        sh.getRange(r+1,iH+1).setValue(passhash);
        logEvent_('user_reset_password', requestorLogin, login);
        return {ok:true};
      }
    }
    return {ok:false, msg:'Пользователь не найден'};
  }catch(e){ return {ok:false, msg:e.message}; }
}

/*** === МЕНЮ: сгенерировать salt+hash === ***/
function makeSaltHash(password){
  if (!password) throw new Error('Пароль пуст');
  const salt = Utilities.getUuid().replace(/-/g,'').slice(0,16);
  const passhash = sha256Hex_(salt + password);
  return { salt, passhash };
}

function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('CRM-Auth')
    .addItem('Сгенерировать salt+hash…', 'menuMakeSaltHash_')
    .addToUi();
}

function menuMakeSaltHash_(){
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Введите новый пароль', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton()!==ui.Button.OK) return;
  const {salt, passhash} = makeSaltHash(res.getResponseText());
  ui.alert('Готово', `salt: ${salt}\npasshash: ${passhash}\n\nВставьте в лист Users или Sections.`, ui.ButtonSet.OK);
}
