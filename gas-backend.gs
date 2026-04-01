// ================================================================
//  名刺帳 — GAS バックエンド v2
//  Google Apps Script に貼り付けてWebアプリとしてデプロイ
//  デプロイ設定: 種類=ウェブアプリ / アクセス=全員 / 実行=自分
//
//  管理シート構成:
//    「名刺データ」  — 名刺マスター
//    「メンバー」    — ユーザー・権限管理
//    「アプリ設定」  — 組織名・APIキー等の設定値
// ================================================================

// ── シート名定数 ──
const SH_CARDS   = '名刺データ';
const SH_MEMBERS = 'メンバー';
const SH_CONFIG  = 'アプリ設定';

const CARD_HEADERS = [
  'id','company','dept','name','role','email',
  'tel','mobile','addr','url','memo',
  'org','owner','imageUrl','createdAt'
];

const MEMBER_HEADERS = [
  'id','name','org','email','tel','role','gEmail','avatar','createdAt','profile','password','disabled'
];

const CONFIG_KEYS = [
  'org1Name','org1Color','org2Name','org2Color',
  'appTitle','geminiKey','anthropicKey','folderId',
  'allowPublicView','updatedAt'
];

// ================================================================
//  GET ハンドラ
// ================================================================
function doGet(e) {
  try {
    const action = e.parameter.action || 'list';
    const q      = e.parameter.q || '';

    if (action === 'list')        return jsonRes({ ok:true, cards: getCards(q) });
    if (action === 'getMembers')  return jsonRes({ ok:true, members: getMembers() });
    if (action === 'getConfig')   return jsonRes({ ok:true, config: getConfig() });
    if (action === 'syncAll')     return jsonRes({ ok:true,
      cards:   getCards(''),
      members: getMembers(),
      config:  getConfig()
    });

    return jsonRes({ ok:false, error:'unknown action: '+action });
  } catch(err) {
    return jsonRes({ ok:false, error: err.message });
  }
}

// ================================================================
//  POST ハンドラ
// ================================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    // ── 名刺 CRUD ──
    if (action === 'addCard') {
      const imageUrl = body.card.imageData
        ? uploadImage(body.card.id, body.card.imageData) : '';
      addCard({ ...body.card, imageUrl, imageData: undefined });
      return jsonRes({ ok:true, id: body.card.id, imageUrl });
    }
    if (action === 'updateCard') {
      const existing = getCardById(body.card.id);
      const imageUrl = body.card.imageData
        ? uploadImage(body.card.id, body.card.imageData)
        : (existing?.imageUrl || '');
      updateCard({ ...body.card, imageUrl, imageData: undefined });
      return jsonRes({ ok:true, imageUrl });
    }
    if (action === 'deleteCard') {
      deleteCard(body.id);
      return jsonRes({ ok:true });
    }

    // ── メンバー管理 ──
    if (action === 'saveMember') {
      saveMember(body.member);
      return jsonRes({ ok:true });
    }
    if (action === 'deleteMember') {
      deleteMember(body.id);
      return jsonRes({ ok:true });
    }

    // ── 設定同期 ──
    if (action === 'saveConfig') {
      saveConfig(body.config);
      return jsonRes({ ok:true });
    }

    // ── 認証チェック（管理者 + 一般メンバー両対応）──
    if (action === 'verifyAdmin') {
      const members = getMembers();
      const email    = (body.email||'').toLowerCase().trim();
      const password = (body.password||'').trim();

      // メンバー照合（gEmail or email列）
      const matched = members.find(m =>
        (m.gEmail||'').toLowerCase().trim() === email ||
        (m.email||'').toLowerCase().trim() === email
      );

      // アカウント無効チェック
      if (matched && (matched.disabled||'').toString().toLowerCase() === 'true') {
        return jsonRes({ ok:false, disabledError: true });
      }

      // パスワードチェック（passwordカラムが設定されている場合のみ）
      if (matched && matched.password && matched.password.trim() !== '') {
        if (password !== matched.password.trim()) {
          return jsonRes({ ok:false, passwordError: true });
        }
      }

      const adminEmails = members
        .filter(m => m.role === 'admin' && m.gEmail)
        .map(m => m.gEmail.toLowerCase().trim());
      const isAdmin = adminEmails.includes(email);

      return jsonRes({ ok:true, isAdmin, isMember: !!matched, member: matched || null });
    }

    // 後方互換 (旧クライアント)
    if (action === 'add')    return doPost({ postData:{ contents: JSON.stringify({...body, action:'addCard'}) }});
    if (action === 'update') return doPost({ postData:{ contents: JSON.stringify({...body, action:'updateCard'}) }});
    if (action === 'delete') return doPost({ postData:{ contents: JSON.stringify({...body, action:'deleteCard'}) }});

    return jsonRes({ ok:false, error:'unknown action: '+action });
  } catch(err) {
    return jsonRes({ ok:false, error: err.message });
  }
}

// ================================================================
//  スプレッドシート取得 & 初期化
// ================================================================
function getSS() {
  return SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SHEET_ID')
  );
}

function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      // ヘッダー行の書式
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#f0efe9')
        .setFontWeight('bold');
    }
    // デフォルトデータ投入
    if (name === SH_CONFIG) seedConfig(sheet);
    if (name === SH_MEMBERS) seedMembers(sheet);
  }
  return sheet;
}

// ── デフォルト設定を投入 ──
function seedConfig(sheet) {
  const defaults = {
    org1Name: 'MiKS',
    org1Color: '#1a56db',
    org2Name: 'Linnas',
    org2Color: '#be185d',
    appTitle: '名刺帳',
    geminiKey: '',
    anthropicKey: '',
    folderId: PropertiesService.getScriptProperties().getProperty('FOLDER_ID') || '',
    allowPublicView: 'true',
    updatedAt: new Date().toISOString()
  };
  Object.entries(defaults).forEach(([k,v]) => sheet.appendRow([k, v]));
}

// ── デフォルトメンバーを投入（初回のみ） ──
function seedMembers(sheet) {
  sheet.appendRow([
    'm_init',             // id
    '管理者',             // name
    'miks',               // org
    '',                   // email
    '',                   // tel
    'admin',              // role
    '',                   // gEmail (← 管理者が後で設定)
    '',                   // avatar
    new Date().toISOString() // createdAt
  ]);
}

// ================================================================
//  名刺データ操作
// ================================================================
function getCards(q) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const hdrs = data[0];
  const rows = data.slice(1)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => obj[h] = String(row[i] ?? ''));
      return obj;
    })
    .filter(r => r.id);

  if (!q) return rows;
  const ql = q.toLowerCase();
  return rows.filter(r =>
    ['company','dept','name','role','email','tel','addr','owner','memo']
      .some(k => (r[k]||'').toLowerCase().includes(ql))
  );
}

function getCardById(id) {
  return getCards('').find(c => c.id === id);
}

function addCard(card) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  sheet.appendRow(CARD_HEADERS.map(h => card[h] ?? ''));
}

function updateCard(card) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(card.id)) {
      sheet.getRange(i+1, 1, 1, CARD_HEADERS.length)
        .setValues([CARD_HEADERS.map(h => card[h] ?? '')]);
      return;
    }
  }
  // IDが見つからなければ追加
  addCard(card);
}

function deleteCard(id) {
  const sheet = getOrCreateSheet(SH_CARDS, CARD_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return;
    }
  }
}

// ================================================================
//  メンバー操作
// ================================================================
function getMembers() {
  const sheet = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const hdrs = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      hdrs.forEach((h, i) => obj[h] = String(row[i] ?? ''));
      return obj;
    })
    .filter(r => r.id);
}

function saveMember(member) {
  const sheet   = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data    = sheet.getDataRange().getValues();
  const now     = new Date().toISOString();
  const headers = data[0];

  // idカラムとgEmailカラムのインデックスを取得
  const idIdx     = headers.indexOf('id');
  const gEmailIdx = headers.indexOf('gEmail');
  const emailIdx  = headers.indexOf('email');

  // 既存行を検索（id → gEmail → email の順で照合）
  for (let i = 1; i < data.length; i++) {
    const rowId     = String(data[i][idIdx]     || '');
    const rowGEmail = String(data[i][gEmailIdx] || '').toLowerCase();
    const rowEmail  = String(data[i][emailIdx]  || '').toLowerCase();
    const memberEmail = (member.gEmail || member.email || '').toLowerCase();

    if (
      (member.id && rowId === String(member.id)) ||
      (memberEmail && rowGEmail === memberEmail) ||
      (memberEmail && rowEmail  === memberEmail)
    ) {
      // 既存行を更新（idは変えない、createdAtも保持）
      const existingId        = data[i][idIdx];
      const existingCreatedAt = data[i][headers.indexOf('createdAt')] || now;
      member.id        = member.id || existingId;
      member.createdAt = existingCreatedAt;
      sheet.getRange(i+1, 1, 1, MEMBER_HEADERS.length)
        .setValues([MEMBER_HEADERS.map(h => member[h] ?? '')]);
      return;
    }
  }
  // 新規追加
  if (!member.id)        member.id        = 'm_' + Date.now();
  if (!member.createdAt) member.createdAt = now;
  sheet.appendRow(MEMBER_HEADERS.map(h => member[h] ?? ''));
}

function deleteMember(id) {
  const sheet = getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return;
    }
  }
}

// ================================================================
//  アプリ設定操作（キーバリュー形式）
// ================================================================
function getConfig() {
  const sheet = getOrCreateSheet(SH_CONFIG, ['key', 'value']);
  const data  = sheet.getDataRange().getValues();
  const config = {};
  data.slice(1).forEach(row => {
    if (row[0]) config[String(row[0])] = String(row[1] ?? '');
  });
  return config;
}

function saveConfig(config) {
  const sheet = getOrCreateSheet(SH_CONFIG, ['key', 'value']);
  const data  = sheet.getDataRange().getValues();
  const keyRow = {};

  // 既存キーの行番号を収集
  data.forEach((row, i) => {
    if (i > 0 && row[0]) keyRow[String(row[0])] = i + 1;
  });

  Object.entries(config).forEach(([k, v]) => {
    if (keyRow[k]) {
      sheet.getRange(keyRow[k], 2).setValue(v);
    } else {
      sheet.appendRow([k, v]);
    }
  });

  // updatedAt を更新
  const updRow = keyRow['updatedAt'];
  const now = new Date().toISOString();
  if (updRow) sheet.getRange(updRow, 2).setValue(now);
  else sheet.appendRow(['updatedAt', now]);
}

// ================================================================
//  Google Drive 画像アップロード
// ================================================================
function uploadImage(id, dataUrl) {
  try {
    // フォルダIDはシート設定 → スクリプトプロパティの順で取得
    const config   = getConfig();
    const folderId = config.folderId
      || PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    if (!folderId) return '';

    const folder = DriveApp.getFolderById(folderId);

    // 同名ファイルを削除
    const existing = folder.getFilesByName('meishi_'+id+'.jpg');
    while (existing.hasNext()) existing.next().setTrashed(true);

    const match = dataUrl.match(/^data:(.+);base64,(.+)$/);
    if (!match) return '';
    const [, mimeType, base64] = match;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64), mimeType, 'meishi_'+id+'.jpg'
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch(err) {
    console.error('uploadImage failed:', err);
    return '';
  }
}

// ================================================================
//  ユーティリティ
// ================================================================
function jsonRes(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 手動初期化: スクリプトエディタから実行してシートを作成 ──
function initializeSheets() {
  getOrCreateSheet(SH_CARDS,   CARD_HEADERS);
  getOrCreateSheet(SH_MEMBERS, MEMBER_HEADERS);
  getOrCreateSheet(SH_CONFIG,  ['key','value']);
  SpreadsheetApp.getUi().alert('✅ シートの初期化が完了しました');
}
