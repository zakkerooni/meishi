// ================================================================
//  名刺帳 — GAS バックエンド v3 (Security, Performance & OCR Updated)
// ================================================================

const SH_CARDS   = '名刺データ';
const SH_MEMBERS = 'メンバー';
const SH_CONFIG  = 'アプリ設定';

const CARD_HEADERS = [
  'id','company','dept','name','role','email',
  'tel','mobile','addr','url','memo',
  'org','owner','imageUrl','createdAt',
  'kana','exchangedAt','imageUrl2','memoPhotos'
];
const MEMBER_HEADERS = [
  'id','name','org','email','tel','role','gEmail','avatar','createdAt','profile','password','disabled'
];

// GETリクエスト（URL直叩き）はセキュリティのためすべて遮断します
function doGet(e) {
  return ContentService.createTextOutput("Access Denied: URLからの直接アクセスは許可されていません。");
}

// ================================================================
//  POST ハンドラ (全ての通信をここに集約)
// ================================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    // --- セキュリティ: アプリケーショントークンの検証 ---
    const config = getConfig();
    let appToken = config.appToken;
    if (!appToken) {
      appToken = Utilities.getUuid();
      saveConfig({ ...config, appToken: appToken });
    }

    // 修正：ログインアクション以外、かつトークンが一致しない場合のみ弾く
    if (action !== 'verifyAdmin') {
      if (body.appToken !== appToken) {
        return jsonRes({ ok: false, error: 'Unauthorized: 無効なアクセスです。' });
      }
    }

    // verifyAdmin（ログイン時）以外の通信は、正しいトークンがないと弾く
    if (action !== 'verifyAdmin') {
      if (body.appToken !== appToken) {
        return jsonRes({ ok: false, error: 'Unauthorized: 無効なアクセスです。' });
      }
    }

    // ── 認証チェック（ログイン時） ──
    if (action === 'verifyAdmin') {
      const members = getMembers();
      const email    = (body.email||'').toLowerCase().trim();
      const password = (body.password||'').trim();
      const matched = members.find(m =>
        (m.gEmail||'').toLowerCase().trim() === email ||
        (m.email||'').toLowerCase().trim() === email
      );

      if (matched && (matched.disabled||'').toString().toLowerCase() === 'true') {
        return jsonRes({ ok:false, disabledError: true });
      }
      if (matched && matched.password && matched.password.trim() !== '') {
        if (password !== matched.password.trim()) {
          return jsonRes({ ok:false, passwordError: true });
        }
      }

      const adminEmails = members.filter(m => m.role === 'admin' && m.gEmail).map(m => m.gEmail.toLowerCase().trim());
      const isAdmin = adminEmails.includes(email);

      return jsonRes({ ok:true, isAdmin, isMember: !!matched, member: matched || null, appToken: (isAdmin || matched) ? appToken : null });
    }

    // ── データ取得系 ──
    if (action === 'list')        return jsonRes({ ok:true, cards: getCards() });
    if (action === 'getMembers')  return jsonRes({ ok:true, members: getMembers() });
    if (action === 'getConfig')   return jsonRes({ ok:true, config: getConfig() });
    if (action === 'syncAll')     return jsonRes({ ok:true, cards: getCards(), members: getMembers(), config: getConfig() });

    // ── 名刺 CRUD ──
    if (action === 'addCard' || action === 'add') {
      const card = body.card;
      const imageUrl = card.imageData ? uploadImage(card.id + '_front', card.imageData) : '';
      const imageUrl2 = card.imageData2 ? uploadImage(card.id + '_back', card.imageData2) : '';
      const memoUrls = uploadMemoPhotos(card.id, card.memoPhotos || []);
      addCard({ ...card, imageUrl, imageUrl2, memoPhotos: JSON.stringify(memoUrls), imageData: undefined, imageData2: undefined });
      return jsonRes({ ok:true, id: card.id, imageUrl, imageUrl2, memoPhotos: memoUrls });
    }
    
    if (action === 'updateCard' || action === 'update') {
      const card = body.card;
      const existing = getCardById(card.id);
      const imageUrl = card.imageData ? uploadImage(card.id + '_front', card.imageData) : (existing?.imageUrl || card.imageUrl || '');
      const imageUrl2 = card.imageData2 ? uploadImage(card.id + '_back', card.imageData2) : (existing?.imageUrl2 || card.imageUrl2 || '');
      let memoPhotos = card.memoPhotos || [];
      if (Array.isArray(memoPhotos)) memoPhotos = uploadMemoPhotos(card.id, memoPhotos);
      updateCard({ ...card, imageUrl, imageUrl2, memoPhotos: JSON.stringify(memoPhotos), imageData: undefined, imageData2: undefined });
      return jsonRes({ ok:true, imageUrl, imageUrl2, memoPhotos });
    }
    
    if (action === 'deleteCard' || action === 'delete') {
      deleteCard(body.id);
      return jsonRes({ ok:true });
    }

    // ── その他設定・メンバー ──
    if (action === 'saveMember') { saveMember(body.member); return jsonRes({ ok:true }); }
    if (action === 'deleteMember') { deleteMember(body.id); return jsonRes({ ok:true }); }
    if (action === 'saveConfig') { saveConfig(body.config); return jsonRes({ ok:true }); }

    return jsonRes({ ok:false, error:'unknown action: '+action });
  } catch(err) {
    return jsonRes({ ok:false, error: err.message });
  }
}

// ================================================================
//  パフォーマンス最適化: スプレッドシートアクセスキャッシュ
// ================================================================
let _ssCache = null;
function getSS() {
  if (!_ssCache) {
    _ssCache = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  }
  return _ssCache;
}

function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#f0efe9').setFontWeight('bold');
    }
    if (name === SH_CONFIG) seedConfig(sheet);
    if (name === SH_MEMBERS) seedMembers(sheet);
  }
  return sheet;
}

function seedConfig(sheet) {
  const defaults = {
    org1Name: 'MiKS', org1Color: '#1a56db', org2Name: 'Linnas', org2Color: '#be185d',
    appTitle: '名刺帳', geminiKey: '', anthropicKey: '', allowPublicView: 'false',
    appToken: Utilities.getUuid(), updatedAt: new Date().toISOString()
  };
  Object.entries(defaults).forEach(([k,v]) => sheet.appendRow([k, v]));
}

function seedMembers(sheet) {
  sheet.appendRow(['m_init','管理者','miks','','','admin','','',new Date().toISOString()]);
}

function getCards() {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const hdrs = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h, i) => {
      const val = String(row[i] ?? '');
      if (h === 'memoPhotos' && val) {
        try { obj[h] = JSON.parse(val); } catch(e) { obj[h] = []; }
      } else { obj[h] = val; }
    });
    return obj;
  }).filter(r => r.id);
}

function getCardById(id) { return getCards().find(c => c.id === id); }
function addCard(card) { const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS); sheet.appendRow(CARD_HEADERS.map(h => card[h] ?? '')); }
function updateCard(card) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(card.id)) {
      sheet.getRange(i+1, 1, 1, CARD_HEADERS.length).setValues([CARD_HEADERS.map(h => card[h] ?? '')]);
      return;
    }
  }
  addCard(card);
}
function deleteCard(id) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return; }
  }
}

function getMembers() {
  const sheet = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const hdrs = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h, i) => { obj[h] = String(row[i] ?? ''); });
    return obj;
  }).filter(r => r.id);
}

function saveMember(member) {
  const sheet = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const gEmailIdx = headers.indexOf('gEmail');
  const emailIdx = headers.indexOf('email');
  
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][idIdx] || '');
    const rowGEmail = String(data[i][gEmailIdx] || '').toLowerCase();
    const rowEmail = String(data[i][emailIdx] || '').toLowerCase();
    const memberEmail = (member.gEmail || member.email || '').toLowerCase();

    if ((member.id && rowId === String(member.id)) || (memberEmail && rowGEmail === memberEmail) || (memberEmail && rowEmail === memberEmail)) {
      member.id = member.id || data[i][idIdx];
      member.createdAt = data[i][headers.indexOf('createdAt')] || now;
      sheet.getRange(i+1, 1, 1, MEMBER_HEADERS.length).setValues([MEMBER_HEADERS.map(h => member[h] ?? '')]);
      return;
    }
  }
  if (!member.id) member.id = 'm_' + Date.now();
  if (!member.createdAt) member.createdAt = now;
  sheet.appendRow(MEMBER_HEADERS.map(h => member[h] ?? ''));
}

function deleteMember(id) {
  const sheet = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return; }
  }
}

function getConfig() {
  const sheet = getOrCreateSheet(SH_CONFIG, ['key', 'value']);
  const data  = sheet.getDataRange().getValues();
  const config = {};
  data.slice(1).forEach(row => { if (row[0]) config[String(row[0])] = String(row[1] ?? ''); });
  return config;
}

function saveConfig(config) {
  const sheet = getOrCreateSheet(SH_CONFIG, ['key', 'value']);
  const data  = sheet.getDataRange().getValues();
  const keyRow = {};
  data.forEach((row, i) => { if (i > 0 && row[0]) keyRow[String(row[0])] = i + 1; });
  Object.entries(config).forEach(([k, v]) => {
    if (keyRow[k]) sheet.getRange(keyRow[k], 2).setValue(v);
    else sheet.appendRow([k, v]);
  });
  const updRow = keyRow['updatedAt'];
  const now = new Date().toISOString();
  if (updRow) sheet.getRange(updRow, 2).setValue(now);
  else sheet.appendRow(['updatedAt', now]);
}

function uploadMemoPhotos(cardId, photos) {
  if (!photos || !photos.length) return [];
  if (typeof photos === 'string') { try { photos = JSON.parse(photos); } catch(e) { return []; } }
  return photos.map((p, i) => {
    if (!p) return '';
    if (p.startsWith('http')) return p;
    return uploadImage('memo_' + cardId + '_' + i, p);
  }).filter(Boolean);
}

function uploadImage(id, dataUrl) {
  try {
    const config = getConfig();
    const folderId = config.folderId || PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    if (!folderId) return '';
    const folder = DriveApp.getFolderById(folderId);
    const existing = folder.getFilesByName('meishi_'+id+'.jpg');
    while (existing.hasNext()) existing.next().setTrashed(true);
    const match = dataUrl.match(/^data:(.+);base64,(.+)$/);
    if (!match) return '';
    const [, mimeType, base64] = match;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, 'meishi_'+id+'.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch(err) { return ''; }
}

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function initializeSheets() {
  getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  getOrCreateSheet(SH_CONFIG, ['key','value']);
  SpreadsheetApp.getUi().alert('✅ シートの初期化が完了しました');
}
