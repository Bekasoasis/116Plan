/*******************************************************************
 * 116學年度系所特色發展計畫 — 線上填報系統（Google Apps Script 後端）
 * 目標試算表：見 SPREADSHEET_ID
 * 部署方式：Apps Script → 部署 → 新增部署 → 網頁應用程式
 *           執行身分「我」、存取權「南華大學網域內所有人」（或任何人）
 *******************************************************************/

// 學校帳號執行、資料寫入學校試算表。
/**
 * @OnlyCurrentDoc
 */
const SPREADSHEET_ID = '13UsTozJDEr3f0MmFi5ShacvvKbIGaWGf-CAKHwFvW9c';

// 填報頁（GitHub Pages）。「設定_系統」的「前端網址」有填時以該欄為準。
const FRONTEND_URL = 'https://bekasoasis.github.io/116Plan/';

// 寄送通知信（學校帳號寄出）。需 appsscript.json 含 script.send_mail 權限。
const SEND_MAIL = true;
const SH = {
  main : '填報主檔',
  equip: '設備明細',
  units: '設定_學院系所',
  plan : '設定_校發計畫',
  feat : '設定_辦學特色',
  sys  : '設定_系統'
};
const MAIN_HEADERS = [
  '填報編號','編輯碼','狀態','學院','系所','計畫名稱',
  '聯絡人','分機','Email',
  '對應校發主軸','對應重點項目','對應辦學特色','與學校特色關聯說明',
  '現況分析與需求','計畫目標','執行做法','預期達成目標','量化指標',
  '設備項數','設備總經費','估價單連結',
  '建立時間','最後更新','送出時間'
];
const EQUIP_HEADERS = [
  '填報編號','學院','系所','序號','設備名稱','基本規格',
  '單價','數量','小計','設置地點','用途／對應做法'
];

/* ---------- 網頁進入點 ----------
 * 兩種用法並存：
 *  A. 直接開 Apps Script 網址 → 內建 Index.html
 *  B. 前端放 GitHub Pages → 以 fetch 呼叫本網址 ?action=config（GET）
 *     或 POST JSON {action:'save'|'load', ...}（Content-Type 用 text/plain 免 CORS 預檢）
 */
function doGet(e) {
  const a = (e && e.parameter.action) || '';
  if (a === 'config') return json_(getConfig());
  if (a === 'load')   return json_(safe_(() => loadRecord(e.parameter.id, e.parameter.code)));
  try {
    const t = HtmlService.createTemplateFromFile('Index');
    t.prefillId   = (e && e.parameter.id)   || '';
    t.prefillCode = (e && e.parameter.code) || '';
    return t.evaluate()
      .setTitle('116學年度系所特色發展計畫填報')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // 專案內沒有 Index.html：轉到 GitHub 填報頁，續填參數一併帶過去
    const q = (e && e.parameter && e.parameter.id && e.parameter.code)
      ? '?id=' + encodeURIComponent(e.parameter.id) + '&code=' + encodeURIComponent(e.parameter.code) : '';
    const to = frontendUrl_() + q;
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font-family:sans-serif;padding:40px;text-align:center">' +
      '<p>填報系統已移至新網址，正在轉跳…</p>' +
      '<p><a href="' + to + '" target="_top" style="font-size:18px">點此前往填報頁</a></p></div>' +
      '<script>try{window.top.location.replace(' + JSON.stringify(to) + ')}catch(e){}</script>'
    ).setTitle('116學年度系所特色發展計畫填報');
  }
}
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'JSON 格式錯誤' }); }
  if (p.action === 'save') return json_(safe_(() => saveRecord(p)));
  if (p.action === 'load') return json_(safe_(() => loadRecord(p.id, p.code)));
  if (p.action === 'config') return json_(getConfig());
  return json_({ error: '未知的 action' });
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function frontendUrl_() {
  try { return readSys_(SpreadsheetApp.openById(SPREADSHEET_ID))['前端網址'] || FRONTEND_URL; }
  catch (e) { return FRONTEND_URL; }
}

/* 在編輯器直接執行：寄一封測試信給自己，確認寄信權限與信箱 */
function testMail() {
  const me = Session.getActiveUser().getEmail();
  MailApp.sendEmail({ to: me, name: '南華大學研發處校務發展組', subject: '[116填報系統] 寄信測試', body: '收到這封信代表通知信功能正常。' });
  Logger.log('已寄測試信至 %s，今日剩餘配額 %s 封', me, MailApp.getRemainingDailyQuota());
}

function safe_(fn) { try { return fn(); } catch (err) { return { error: err.message }; } }

/* ---------- 第一次使用：建立工作表、種子資料、上傳資料夾 ---------- */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  ensureSheet_(ss, SH.main,  MAIN_HEADERS);
  ensureSheet_(ss, SH.equip, EQUIP_HEADERS);

  const units = ensureSheet_(ss, SH.units, ['學院','系所']);
  if (units.getLastRow() < 2) {
    // 依校網教學單位頁整理（2026-09）；系所異動時直接改試算表即可
    units.getRange(2,1,SEED_UNITS.length,2).setValues(SEED_UNITS);
  }
  const plan = ensureSheet_(ss, SH.plan, ['主軸代碼','主軸名稱','重點項目代碼','重點項目名稱']);
  if (plan.getLastRow() < 2) {
    // 112-116 校發計畫主軸與重點項目
    plan.getRange(2,1,SEED_PLAN.length,4).setValues(SEED_PLAN);
  }
  const feat = ensureSheet_(ss, SH.feat, ['特色代碼','辦學特色','說明']);
  if (feat.getLastRow() < 2) {
    // ★ 依現行辦學特色預載，如 117-121 修正後請同步更新
    feat.getRange(2,1,SEED_FEAT.length,3).setValues(SEED_FEAT);
  }
  const sys = ensureSheet_(ss, SH.sys, ['項目','值']);
  if (sys.getLastRow() < 2) {
    const folderId = Drive.Files.create({ name: '116系所特色發展計畫_估價單', mimeType: 'application/vnd.google-apps.folder' }).id;
    sys.getRange(2,1,5,2).setValues([
      ['上傳資料夾ID', folderId],
      ['填報截止', ''],
      ['公告文字', '請於截止日前完成送出；送出後仍可用編輯碼修改。'],
      ['通知副本Email', ''],
      ['前端網址', '']   // 若前端放 GitHub Pages，填該網址，通知信的續填連結會指向它
    ]);
  }
  return '完成：工作表與上傳資料夾已建立';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#5C4033').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- 前端讀取設定 ---------- */
function getConfig() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rows = name => { const sh = ss.getSheetByName(name); return sh && sh.getLastRow() > 1 ? sh.getDataRange().getValues().slice(1) : []; };

  const colleges = {};
  rows(SH.units).filter(r => r[0] && r[1])
    .forEach(r => { (colleges[String(r[0]).trim()] = colleges[String(r[0]).trim()] || []).push(String(r[1]).trim()); });

  const plan = rows(SH.plan).filter(r => r[0] && r[2])
    .map(r => ({ g: String(r[0]), gName: String(r[1]), s: String(r[2]), sName: String(r[3]) }));

  const features = rows(SH.feat).filter(r => r[0] && r[1])
    .map(r => ({ code: String(r[0]), name: String(r[1]), desc: String(r[2] || '') }));

  const missing = [SH.units, SH.plan, SH.feat, SH.sys].filter(n => !ss.getSheetByName(n));
  return { colleges, plan, features, sys: readSys_(ss), missing };
}

/* 在編輯器直接執行，看 Logger 輸出確認設定讀得到 */
function testConfig() {
  const c = getConfig();
  Logger.log('學院數：%s　系所數：%s　重點項目數：%s　特色數：%s　缺少工作表：%s',
    Object.keys(c.colleges).length,
    Object.values(c.colleges).reduce((a, b) => a + b.length, 0),
    c.plan.length, c.features.length, c.missing.join('、') || '無');
}

function readSys_(ss) {
  const o = {};
  const sh = ss.getSheetByName(SH.sys);
  if (!sh) return o;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const v = r[1];
    o[String(r[0]).trim()] = (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : String(v || '');
  });
  return o;
}

/* ---------- 儲存（草稿／送出） ---------- */
function saveRecord(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const main = ss.getSheetByName(SH.main);
    const eq   = ss.getSheetByName(SH.equip);
    const sys  = readSys_(ss);
    const f    = p.form || {};
    const equipment = (p.equipment || []).filter(r => r.name);
    const now  = new Date();

    // 1. 編號與編輯碼：新案配發，舊案驗證
    let id = p.id, code = p.code, rowIdx = -1, created = now, oldFiles = '';
    if (id) {
      rowIdx = findRow_(main, id);
      if (rowIdx < 0) throw new Error('找不到此填報編號');
      const row = main.getRange(rowIdx, 1, 1, MAIN_HEADERS.length).getValues()[0];
      if (String(row[1]) !== String(code)) throw new Error('編輯碼錯誤');
      created  = row[21] || now;
      oldFiles = row[20] || '';
    } else {
      id   = nextId_(main);
      code = genCode_();
    }

    // 2. 上傳估價單（用進階 Drive 服務，只需 drive.file 範圍）
    const links = String(oldFiles).split('\n').filter(Boolean);
    if (p.files && p.files.length) {
      const parentId = getOrCreateSub_(sys['上傳資料夾ID'], id + '_' + f.college + '_' + f.dept);
      p.files.forEach(fl => {
        const blob = Utilities.newBlob(Utilities.base64Decode(fl.data), fl.mime, fl.name);
        const meta = { name: fl.name, parents: [parentId] };
        const created = Drive.Files.create(meta, blob, { fields: 'id,webViewLink' });
        const link = created.webViewLink || ('https://drive.google.com/file/d/' + created.id + '/view');
        links.push(fl.name + ' | ' + link);
      });
    }

    // 3. 設備彙總
    const total = equipment.reduce((s, r) => s + (Number(r.price)||0) * (Number(r.qty)||0), 0);

    // 4. 主檔列
    const status    = p.status === 'submitted' ? '已送出' : '草稿';
    const submitted = p.status === 'submitted' ? now : (rowIdx > 0 ? main.getRange(rowIdx, 24).getValue() : '');
    const row = [
      id, code, status, f.college, f.dept, f.title,
      f.contact, f.ext, f.email,
      (f.goals || []).join('、'), (f.strategies || []).join('、'),
      (f.features || []).join('、'), f.featureNote,
      f.context, f.objective, f.method, f.outcome, f.kpi,
      equipment.length, total, links.join('\n'),
      created, now, submitted
    ];
    if (rowIdx > 0) main.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else main.appendRow(row);

    // 5. 設備明細：先刪舊列再寫入
    deleteRowsById_(eq, id);
    if (equipment.length) {
      eq.getRange(eq.getLastRow()+1, 1, equipment.length, EQUIP_HEADERS.length).setValues(
        equipment.map((r, i) => [
          id, f.college, f.dept, i+1, r.name, r.spec,
          Number(r.price)||0, Number(r.qty)||0, (Number(r.price)||0)*(Number(r.qty)||0),
          r.location, r.purpose
        ])
      );
    }

    // 6. 續填連結（不寄信，避免 Gmail 授權被 Google 封鎖；連結顯示在畫面上並可存回試算表）
    const base = sys['前端網址'] || FRONTEND_URL;
    const url  = base + (base.indexOf('?') > -1 ? '&' : '?') + 'id=' + id + '&code=' + code;

    // 7. 通知信（失敗不影響存檔，結果回傳給前端顯示）
    let mailed = false, mailError = '';
    if (SEND_MAIL && f.email) {
      try {
        const cc = String(sys['通知副本Email'] || '').trim();
        const opt = {
          to: f.email,
          name: '南華大學研發處校務發展組',
          subject: '[116系所特色發展計畫] ' + f.dept + '：' + status + '（' + id + '）',
          htmlBody:
            '<p>' + f.college + ' ' + f.dept + ' 的填報已' + (status === '已送出' ? '送出' : '儲存為草稿') + '。</p>' +
            '<p>計畫名稱：' + (f.title || '') + '</p>' +
            '<p>填報編號：<b>' + id + '</b>　編輯碼：<b>' + code + '</b></p>' +
            '<p>要繼續填寫或修改，請點：<a href="' + url + '">' + url + '</a></p>' +
            '<p>設備 ' + equipment.length + ' 項，合計 NT$ ' + total.toLocaleString() + '。</p>' +
            '<p style="color:#888">本信由系統自動寄出，請勿直接回覆。</p>'
        };
        if (cc) opt.cc = cc;
        MailApp.sendEmail(opt);
        mailed = true;
      } catch (err) { mailError = err.message; }
    }

    return { id, code, url, status, total, files: links, mailed, mailError };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 重新連接：讀回既有資料 ---------- */
function loadRecord(id, code) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const main = ss.getSheetByName(SH.main);
  const r    = findRow_(main, id);
  if (r < 0) throw new Error('找不到此填報編號');
  const v = main.getRange(r, 1, 1, MAIN_HEADERS.length).getValues()[0];
  if (String(v[1]) !== String(code)) throw new Error('編輯碼錯誤');

  const eq = ss.getSheetByName(SH.equip).getDataRange().getValues().slice(1)
    .filter(x => String(x[0]) === String(id))
    .map(x => ({ name:x[4], spec:x[5], price:x[6], qty:x[7], location:x[9], purpose:x[10] }));

  return {
    id, code, status: v[2],
    form: {
      college:v[3], dept:v[4], title:v[5], contact:v[6], ext:v[7], email:v[8],
      goals: split_(v[9]), strategies: split_(v[10]),
      features: split_(v[11]), featureNote: v[12],
      context:v[13], objective:v[14], method:v[15], outcome:v[16], kpi:v[17]
    },
    equipment: eq,
    files: String(v[20] || '').split('\n').filter(Boolean)
  };
}

/* ---------- 工具 ---------- */
function findRow_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow()-1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function nextId_(sh) {
  let max = 0;
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
    const m = /(\d+)$/.exec(String(r[0])); if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'NHU116-' + ('000' + (max + 1)).slice(-3);
}
function genCode_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = ''; for (let i = 0; i < 6; i++) s += chars[Math.floor(Math.random()*chars.length)];
  return s;
}
function deleteRowsById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last-1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
}
function getOrCreateSub_(parentId, name) {
  const q = "name = '" + name.replace(/'/g, "\\'") + "' and '" + parentId +
            "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
  const found = Drive.Files.list({ q: q, fields: 'files(id)', pageSize: 1 });
  if (found.files && found.files.length) return found.files[0].id;
  return Drive.Files.create({ name: name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] }).id;
}
function split_(s) { return String(s || '').split('、').filter(Boolean); }

/* ---------- 種子資料（★請核對修正） ---------- */
const SEED_UNITS = [
  // 依校網「教學單位」頁（2026-09）整理；文創系 115 學年度已併入企管系，故不列
  ['管理學院','企業管理學系'],['管理學院','財務金融學系'],['管理學院','旅遊管理學系'],
  ['管理學院','管理學院國際企業學士學位學程'],
  ['人文學院','文學系'],['人文學院','生死學系'],['人文學院','宗教學研究所'],
  ['人文學院','幼兒教育學系'],['人文學院','外國語文學系'],
  ['人文學院','運動與健康促進學士學位學程'],['人文學院','全齡運動科學與健康促進碩士學位學程'],
  ['人文學院','人間佛教數位學習碩士在職專班'],
  ['人文學院','語文教學中心'],
  ['社會科學院','國際事務與企業學系'],['社會科學院','社會工作學系'],['社會科學院','傳播學系'],
  ['藝術與設計學院','視覺藝術與設計學系'],['藝術與設計學院','建築學系'],
  ['藝術與設計學院','產品與室內設計學系'],['藝術與設計學院','民族音樂學系'],
  ['藝術與設計學院','音樂跨域設計與藝術管理學士學位學程'],
  ['藝術與設計學院','藝術與設計跨域碩士學位學程'],
  ['科技學院','資訊管理學系'],['科技學院','資訊工程學系'],['科技學院','資訊科技進修學士班'],
  ['科技學院','自然生物科技學系'],['科技學院','永續綠色科技碩士學位學程'],
  ['科技學院','半導體應用學士學位學程'],
  ['通識教育中心','通識教育中心']
];
const SEED_FEAT = [
  ['F1','生命教育','以生命教育為核心，涵養學生生命素養與品德'],
  ['F2','環境永續','推動永續發展、綠色校園與 SDGs 實踐'],
  ['F3','智慧創新','發展智慧科技應用與跨域創新教學'],
  ['F4','三好校園','落實做好事、說好話、存好心之校園文化']
];
const SEED_PLAN = [
  // 112-116 學年度校務發展計畫：主軸 → 重點項目（2026-09 由校務發展組提供）
  ['主軸A','精進創新教學','A1','強化教師的教學能力，激勵教師教學創新'],
  ['主軸A','精進創新教學','A2','深化新科技支持系統，涵養資訊科技能力'],
  ['主軸A','精進創新教學','A3','精進自主及跨域課程，提升學生跨域素養'],
  ['主軸A','精進創新教學','A4','以校務分析證據為本，強化教學品保機制'],
  ['主軸A','精進創新教學','A5','精進校務研究及管理機制，落實自我課責'],
  ['主軸B','深化學輔工作','B1','精進三好友善校園，提升三好生命力與競爭力'],
  ['主軸B','深化學輔工作','B2','精進南華志工制度，實踐服務利他與民主素養'],
  ['主軸B','深化學輔工作','B3','優化多元輔導機制，強化學生競爭力與就業力'],
  ['主軸C','強化職涯輔導','C1','深化自主學習課程與設計思考教學，提升競爭力'],
  ['主軸C','強化職涯輔導','C2','精進實習、職場體驗及證照，強化學用證賽合一'],
  ['主軸C','強化職涯輔導','C3','強化四年一貫職涯輔導機制，深化職涯關懷輔導'],
  ['主軸D','厚實產學研發','D1','強化區域產學聯盟，促進區域產業發展'],
  ['主軸D','厚實產學研發','D2','創新創業生態基地，強化創客實作能力'],
  ['主軸D','厚實產學研發','D3','強化教師產學能量，提升教師實務能力'],
  ['主軸D','厚實產學研發','D4','延攬師資獎勵研發，擴增產研發展能量'],
  ['主軸D','厚實產學研發','D5','厚植研究資源共享，建構優質產研環境'],
  ['主軸D','厚實產學研發','D6','優化多元評鑑機制，激發教師研究動能'],
  ['主軸E','追求永續發展','E1','深化生命教育至各校，提升南華校園生命教育影響力'],
  ['主軸E','追求永續發展','E2','拓展生命教育到社會，朝向國際級生命教育典範大學'],
  ['主軸E','追求永續發展','E3','推展環境永續到國際，擴展南華環境永續發展影響力'],
  ['主軸E','追求永續發展','E4','創設實踐專責的基地，以落實大學社會責任與實踐力'],
  ['主軸E','追求永續發展','E5','推動綠色的行動辦公，邁向南華碳中和的綠色生活'],
  ['主軸E','追求永續發展','E6','深化青銀共學的場域，促進青銀、校園及社會的共榮'],
  ['主軸E','追求永續發展','E7','強化各類的風險管理，建構韌性校園與全面品質管理'],
  ['主軸F','增進全球交流','F1','結合佛光山系統資源，提升國際競爭力'],
  ['主軸F','增進全球交流','F2','強化EMI教學與學習，塑造國際化校園'],
  ['主軸F','增進全球交流','F3','促進海外學習與交流，讓南華走向世界'],
  ['主軸F','增進全球交流','F4','構建國際化友善校園，讓世界走向南華'],
  ['主軸G','促進身心靈健康','G1','建調身心靈健康、體育組織與課程，深耕南華學生身心靈更健康'],
  ['主軸G','促進身心靈健康','G2','創設師生健康運動護照與樂活日，促使南華教職員身心靈更健康'],
  ['主軸G','促進身心靈健康','G3','應用物聯網與精準化培訓，引領雲嘉青年與高齡者身心靈更健康']
];
