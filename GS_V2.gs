/**
 * 崇華十六期整合管理系統 - GAS 後端主程式 (全功能最終整合版)
 * 已修復：getTeamsData 缺失問題、重複函式清理、全成員點名邏輯整合、操作日誌系統
 */

const SPREADSHEET_ID = '1sOH7UWJ6w55FHyHCvW_zW0rbCfrFWKtgzgEapUHj7ZI';

/**
 * 處理 GET 請求：資料讀取
 */
function legacyDoGet(e) {
  const action = e.parameter.action;
  const name = e.parameter.name;
  const group = e.parameter.group;
  const role = e.parameter.role;
  const topic = e.parameter.topic;
  
  try {
    switch (action) {
      case 'getHistory': return JSONResponse(getHistory(group, name));
      case 'getConfig': return JSONResponse(handleGetConfig());
      case 'checkAttendance': return JSONResponse(checkAttendanceStatus(group, name));
      case 'getAttendanceReport': return JSONResponse(getAttendanceReport(role, group, e.parameter.bonusSheetName));
      case 'getAnnouncements': return JSONResponse(getAnnouncements());
      case 'getLocations': return JSONResponse(getLocationsData());
      case 'getHomeworkConfig': 
        if (e.parameter.force === '1') {
          removeFromCache("hw_config");
        }
        return JSONResponse(getHomeworkConfig());
      case 'getHomeworkReport': return JSONResponse(getHomeworkReport(topic));
      case 'getUserHomeworkHistory': return JSONResponse(getUserHomeworkHistory(group, name));
      case 'getBonusStats': return JSONResponse(getBonusStats(e.parameter.sheetName));
      case 'getTeamsData':
        if (e.parameter.force === '1') {
          removeFromCache("teams_data");
        }
        return JSONResponse(getTeamsData());
      case 'getRegionReport': return JSONResponse(getRegionReport());
      default: throw new Error("未定義的讀取動作: " + action);
    }
  } catch (err) {
    return JSONResponse({ success: false, error: err.toString() });
  }
}

/**
 * 處理 POST 請求：資料寫入
 */
function legacyDoPost(e) {
  let params;
  try {
    params = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : e.parameter;
  } catch (err) { params = e.parameter; }

  const action = params.action;
  try {
    switch (action) {
      case 'register': return JSONResponse(handleRegister(params));
      case 'login': return JSONResponse(handleLogin(params));
      case 'signIn': return JSONResponse(handleSignIn(params));
      case 'submitQuiz': return JSONResponse(handleSubmitQuiz(params));
      case 'updateConfig': return JSONResponse(handleUpdateConfig(params.config));
      case 'updateAnnouncement': return JSONResponse(updateAnnouncement(params));
      case 'uploadFile': return JSONResponse(handleFileUpload(params));
      case 'saveTransferRecord': return JSONResponse(handleSaveTransferRecord(params));
      case 'addBonusPoints': return JSONResponse(handleAddBonusPoints(params));
      case 'submitHomework': return JSONResponse(handleSubmitHomework(params));
      case 'updateHomeworkSettings': return JSONResponse(handleUpdateHomeworkSettings(params.config));
      case 'createBonusSheet': return JSONResponse(handleCreateBonusSheet());
      default: throw new Error("未定義的寫入動作: " + action);
    }
  } catch (err) {
    return JSONResponse({ success: false, error: err.message || err.toString() });
  }
}

/**
 * 格式化 JSON 回訊
 */
function JSONResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 0. 快取輔助模組 (CacheService)
// ==========================================

function getFromCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    const data = cache.get(key);
    return data ? JSON.parse(data) : null;
  } catch (e) {
    return null;
  }
}

function putInCache(key, data, seconds) {
  try {
    const cache = CacheService.getScriptCache();
    const str = JSON.stringify(data);
    if (str.length < 100000) {
      cache.put(key, str, seconds || 21600); // 預設快取 6 小時
    }
  } catch (e) {}
}

function removeFromCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
  } catch (e) {}
}

/**
 * 試算表手動編輯自動觸發器 (Simple Trigger)
 * 當使用者在 Google Sheets / EXCEL 上手動修改工作表時，自動清除對應快取
 */
function onEdit(e) {
  try {
    const sheet = e && e.range ? e.range.getSheet() : null;
    if (sheet) {
      const sheetName = sheet.getName();
      if (sheetName === 'Users') {
        removeFromCache("teams_data");
      } else if (sheetName === 'HomeworkSettings') {
        removeFromCache("hw_config");
      } else if (sheetName === 'Locations') {
        removeFromCache("locations_data");
      } else if (sheetName === 'Bulletin') {
        removeFromCache("announcements_data");
      } else if (sheetName === 'Config') {
        removeFromCache("sys_config");
      }
    }
  } catch (err) {}
}

// ==========================================
// 1. 系統輔助功能 (日誌)
// ==========================================

/**
 * 記錄操作日誌
 * @param {string} user 目標使用者
 * @param {string} group 目標組別
 * @param {string} action 操作項目
 * @param {string} detail 詳細內容
 * @param {string} operator 執行操作的人 (Operator)
 */
function logAction(user, group, action, detail, operator) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Log') || ss.insertSheet('Log');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間', '操作人員 (Operator)', '目標使用者', '組別', '操作項目', '詳細內容']);
      sheet.getRange(1, 1, 1, 6).setBackground('#f1f5f9').setFontWeight('bold');
    }
    sheet.appendRow([new Date(), operator || user, user, group, action, detail || '']);
  } catch (e) {
    console.error("Log Error:", e);
  }
}

// ==========================================
// 2. 使用者與組別管理
// ==========================================


// 兼容既有 Users 與新名稱 USERS；不自動新增或搬移工作表，避免覆蓋既有帳號。
// 帳號正式儲存於 Apps Script Script Properties，僅供後端存取。
// 第一次遷移必須由專案擁有者手動執行 migrateUsersToScriptProperties()。
// 遷移未啟用前仍讀取原 Users/USERS 工作表；成功後完全改由 GS 讀取。
const V3_USER_PREFIX = 'V3_USER_';
const V3_USER_ACTIVE = 'V3_USERS_ACTIVE';
function v2UsersSheet(ss) {
  const sheet = ss.getSheetByName('USERS') || ss.getSheetByName('Users');
  if (!sheet) throw Error('找不到原有 USERS / Users 工作表；若已完成遷移，請確認 V3_USERS_ACTIVE');
  return sheet;
}
function v3UsersEnabled() {
  return PropertiesService.getScriptProperties().getProperty(V3_USER_ACTIVE) === '1';
}
function v3UserKey(group,name) {
  // UTF-8 JSON identity encoded losslessly: distinct (group,name) pairs have distinct keys.
  // Avoid relying on ambiguous digest/charset handling during the one-time migration.
  const identity=JSON.stringify([String(group||'').trim(),String(name||'').trim()]);
  const key=V3_USER_PREFIX + Utilities.base64EncodeWebSafe(Utilities.newBlob(identity,'text/plain','identity').getBytes()).replace(/=+$/,'');
  if(key.length>240) throw Error('組別或姓名過長，無法作為 Script Properties 帳號鍵值');
  return key;
}
function v3AllUsers() {
  if (!v3UsersEnabled()) throw Error('帳號尚未完成遷移');
  const props=PropertiesService.getScriptProperties().getProperties();
  return Object.keys(props).filter(k=>k.startsWith(V3_USER_PREFIX)).map(k=>JSON.parse(props[k]));
}
function v3GetUser(group,name) {
  const raw=PropertiesService.getScriptProperties().getProperty(v3UserKey(group,name));
  if(!raw) return null;
  const record=JSON.parse(raw);
  if(String(record[0]).trim()!==String(group).trim() || String(record[1]).trim()!==String(name).trim()) throw Error('帳號資料驗證失敗');
  return record;
}
function v3SaveUser(record) {
  const key=v3UserKey(record[0],record[1]);
  const value=JSON.stringify(record);
  if(Utilities.newBlob(value).getBytes().length>9000) throw Error('帳號紀錄超過 Script Properties 單筆容量限制');
  PropertiesService.getScriptProperties().setProperty(key,value);
}
// 此函式只供 Apps Script 編輯器手動執行，不提供任何 HTTP 入口。
// 不刪除原始工作表；先驗證全部資料成功複製，再啟用新儲存模式。
// 遷移前檢查：只輸出重複的組別、姓名與來源列號，不輸出密碼。
// 依遷移使用的實際儲存鍵檢查，僅輸出姓名、組別、列號，不輸出密碼。
function inspectDuplicateUsersForMigration() {
  if(v3UsersEnabled()) return {success:false,message:'已啟用 GS 帳號儲存，無須檢查原表'};
  const sheet=v2UsersSheet(SpreadsheetApp.openById(SPREADSHEET_ID));
  const rows=sheet.getDataRange().getValues();
  const byKey=new Map(),byIdentity=new Map();
  let validRows=0;
  for(let i=1;i<rows.length;i++) {
    const group=String(rows[i][0]||'').trim(),name=String(rows[i][1]||'').trim();
    if(!group||!name||group==='組別') continue;
    validRows++;
    const identity=JSON.stringify([group,name]),key=v3UserKey(group,name);
    for(const [map,id] of [[byKey,key],[byIdentity,identity]]) {
      if(!map.has(id)) map.set(id,[]);
      map.get(id).push({row:i+1,group,name});
    }
  }
  const duplicates=[];
  for(const entries of byKey.values()) {
    if(entries.length>1) duplicates.push({group:entries[0].group,name:entries[0].name,rows:entries.map(e=>e.row),distinctIdentities:new Set(entries.map(e=>JSON.stringify([e.group,e.name]))).size});
  }
  const result={success:true,sheet:sheet.getName(),validRows,uniqueKeys:byKey.size,uniqueIdentities:byIdentity.size,duplicateGroups:duplicates.length,duplicates};
  Logger.log(JSON.stringify(result));
  return result;
}
function migrateUsersToScriptProperties() {
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if(v3UsersEnabled()) throw Error('已完成遷移，為避免覆蓋線上帳號，禁止重複執行');
    const sheet=v2UsersSheet(SpreadsheetApp.openById(SPREADSHEET_ID));
    const rows=sheet.getDataRange().getValues().slice(1);
    const records=rows.filter(r=>String(r[0]||'').trim() && String(r[1]||'').trim() && String(r[0]).trim()!=='組別')
      .map(r=>[String(r[0]).trim(),String(r[1]).trim(),String(r[2]||''),String(r[3]||'學員'),r[4] instanceof Date?r[4].toISOString():String(r[4]||'')]);
    if(!records.length) throw Error('來源工作表沒有帳號，已取消遷移');
    const keys=records.map(r=>v3UserKey(r[0],r[1]));
    if(new Set(keys).size!==keys.length) {
      const diagnostic=inspectDuplicateUsersForMigration();
      throw Error('遷移鍵值重複，來源工作表：'+sheet.getName()+'；有效帳號：'+records.length+'；唯一鍵值：'+new Set(keys).size+'；請查看執行記錄中的診斷結果：'+JSON.stringify(diagnostic));
    }
    const props=PropertiesService.getScriptProperties();
    const existing=props.getProperties();
    if(Object.keys(existing).some(k=>k.startsWith(V3_USER_PREFIX))) throw Error('偵測到先前未完成的遷移資料，請先人工核對，避免覆蓋');
    // Script Properties 具有總容量上限；預估超量時直接停止，不啟用新模式。
    const bytes=records.reduce((n,r)=>n+Utilities.newBlob(JSON.stringify(r)).getBytes().length,0);
    if(bytes>350000) throw Error('帳號資料量過大，不適合 Script Properties，請改用獨立受限存取的帳號資料庫');
    records.forEach(v3SaveUser);
    const copied=records.every(r=>JSON.stringify(v3GetUser(r[0],r[1]))===JSON.stringify(r));
    if(!copied || v3CountStoredUsers()!==records.length) throw Error('帳號資料複製驗證失敗，未切換新模式');
    props.setProperty('V3_USERS_MIGRATED_AT',new Date().toISOString());
    props.setProperty('V3_USERS_COUNT',String(records.length));
    props.setProperty(V3_USER_ACTIVE,'1');
    removeFromCache('teams_data');
    return {success:true,migrated:records.length,active:true,message:'遷移完成。請先測試登入、註冊與報表，再自行備份並移除原工作表。'};
  } finally {lock.releaseLock();}
}
function v3CountStoredUsers() {
  return Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k=>k.startsWith(V3_USER_PREFIX)).length;
}
function verifyUsersMigration() {
  if(!v3UsersEnabled()) return {success:false,active:false,message:'尚未遷移'};
  const count=v3CountStoredUsers();
  return {success:count===Number(PropertiesService.getScriptProperties().getProperty('V3_USERS_COUNT')),active:true,count:count};
}
function getTeamsData() {
  const cached=getFromCache('teams_data');
  if(cached) return {success:true,data:cached};
  try {
    const rows=v3UsersEnabled()?v3AllUsers():v2UsersSheet(SpreadsheetApp.openById(SPREADSHEET_ID)).getDataRange().getValues().slice(1);
    const teams={};
    rows.forEach(r=>{
      const group=String(r[0]||'').trim(),name=String(r[1]||'').trim();
      if(!group||!name||group==='組別') return;
      if(!teams[group]) teams[group]=[];
      if(!teams[group].includes(name)) teams[group].push(name);
    });
    putInCache('teams_data',teams,21600);
    return {success:true,data:teams};
  } catch(err) {return {success:false,error:String(err)};}
}

function legacyHandleRegister(data) { return v2Register(data); }

function legacyHandleLogin(data) { return v2Login(data); }

// ==========================================
// 3. 簽到系統
// ==========================================

function legacyHandleSignIn(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
  const sheetName = '簽到' + dateStr;
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['時間', '組別', '姓名', '狀態', '備註']);
    sheet.getRange(1, 1, 1, 5).setBackground('#e2f3e7').setFontWeight('bold');
  }
  sheet.appendRow([new Date(), data.group, data.name, data.leaveType || '正常出席', data.reason || '']);
  logAction(data.name, data.group, "簽到", `狀態: ${data.leaveType || '正常出席'}, 備註: ${data.reason || '無'}`, data.admin || data.name);
  return { success: true };
}

function checkAttendanceStatus(group, name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
  const sheetName = '簽到' + dateStr;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { signedIn: false };
  const today = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  const row = sheet.getDataRange().getValues().slice().reverse().find(r => 
    r[0] instanceof Date && Utilities.formatDate(r[0], "GMT+8", "yyyy-MM-dd") === today && r[1] == group && r[2] == name
  );
  return { signedIn: !!row, time: row ? Utilities.formatDate(new Date(row[0]), "GMT+8", "HH:mm") : "" };
}

function getAttendanceReport(role, userGroup, bonusSheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
    const attendanceSheet = ss.getSheetByName('簽到' + todayStr);
    const userRows = v3UsersEnabled()?v3AllUsers():v2UsersSheet(ss).getDataRange().getValues().slice(1);
    
    // 1. 獲取所有使用者 (過濾名單，徹底去除前後空格與大小寫容錯)
    const safeRole = (role || '').toString().toUpperCase();
    const safeUserGroup = (userGroup || '').toString().trim();
    const isFullAccess = safeRole.includes('ADMIN') || safeRole.includes('TEACHER') || safeRole.includes('管理員') || safeRole.includes('老師');

    const allUsers = userRows
      .map(u => ({
        group: u[0] ? u[0].toString().trim() : "",
        name: u[1] ? u[1].toString().trim() : ""
      }))
      .filter(u => {
        if (!u.group || !u.name || u.group === "組別") return false;
        if (isFullAccess) return true;
        return u.group === safeUserGroup;
      });

    // 2. 獲取今日簽到紀錄
    const attendanceMap = {};
    if (attendanceSheet) {
      const rows = attendanceSheet.getDataRange().getValues();
      rows.slice(1).forEach(r => {
        let extractedTime = "";
        try {
          if (r[0]) extractedTime = Utilities.formatDate(new Date(r[0]), "GMT+8", "HH:mm");
        } catch (e) {}
        const aName = r[2] ? r[2].toString().trim() : "";
        if (aName) {
          attendanceMap[aName] = { time: extractedTime || "", note: r[3] ? r[3].toString().trim() : "" };
        }
      });
    }

    // 3. 獲取累計加分 (單表與全期總計)
    const bonusMap = {};
    const totalBonusMap = {};
    const allBonusSheets = sheets.filter(s => s.getName().startsWith("加分_"));
    
    let targetBonusSheet = allBonusSheets.sort((a,b) => b.getName().localeCompare(a.getName()))[0];
    if (bonusSheetName) {
      const found = allBonusSheets.find(s => s.getName() === bonusSheetName);
      if (found) targetBonusSheet = found;
    }

    allBonusSheets.forEach(s => {
      const bRows = s.getDataRange().getValues();
      for(let i=1; i<bRows.length; i++) {
        const bName = bRows[i][2] ? bRows[i][2].toString().trim() : "";
        const bPts = parseFloat(bRows[i][3] || 0);
        if (bName) {
           totalBonusMap[bName] = (totalBonusMap[bName] || 0) + bPts;
           if (s.getName() === targetBonusSheet?.getName()) {
              bonusMap[bName] = (bonusMap[bName] || 0) + bPts;
           }
        }
      }
    });

    // 4. 獲取累計作業繳交次數 (掃描所有作業表)
    const hwCountMap = {};
    sheets.filter(s => s.getName().startsWith("作業_")).forEach(s => {
      const hRows = s.getDataRange().getValues();
      for(let i=1; i<hRows.length; i++) {
        const hName = hRows[i][2] ? hRows[i][2].toString().trim() : "";
        if (hName) hwCountMap[hName] = (hwCountMap[hName] || 0) + 1;
      }
    });

    // 5. 整合資料：以所有使用者為基礎
    const result = allUsers.map(user => {
      const sign = attendanceMap[user.name];
      return {
        group: user.group,
        name: user.name,
        time: sign ? sign.time : "",
        note: sign ? sign.note : "",
        pts: bonusMap[user.name] || 0,
        totalPts: totalBonusMap[user.name] || 0,
        hwCount: hwCountMap[user.name] || 0
      };
    });

    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ==========================================
// 4. 加分功能模組
// ==========================================

function handleCreateBonusSheet(operator) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");
  let sheetName = '加分_' + dateStr;
  
  let finalName = sheetName;
  let counter = 1;
  while (ss.getSheetByName(finalName)) {
    finalName = sheetName + "_" + counter;
    counter++;
  }
  
  const sheet = ss.insertSheet(finalName);
  sheet.appendRow(['時間', '組別', '姓名', '加分分值', '原因', '操作人']);
  sheet.getRange(1, 1, 1, 6).setBackground('#fef3c7').setFontWeight('bold');
  logAction("管理員", "ADMIN", "建立加分頁", `分頁名稱: ${finalName}`, operator || "管理員");
  return { success: true, sheetName: finalName };
}

function handleAddBonusPoints(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const bonusSheets = sheets
    .filter(s => s.getName().indexOf('加分_') === 0)
    .sort((a, b) => b.getName().localeCompare(a.getName()));
    
  let sheet;
  if (bonusSheets.length > 0) {
    sheet = bonusSheets[0];
  } else {
    const res = handleCreateBonusSheet(data.admin);
    sheet = ss.getSheetByName(res.sheetName);
  }
  
  sheet.appendRow([new Date(), data.group, data.name, data.points, data.reason || '無備註', data.admin || '管理員']);
  logAction(data.name, data.group, "加分", `分數: ${data.points}, 原因: ${data.reason}`, data.admin || '管理員');
  return { success: true };
}

function getBonusStats(targetSheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const bonusSheets = sheets
    .filter(s => s.getName().indexOf('加分_') === 0)
    .sort((a, b) => b.getName().localeCompare(a.getName()));
    
  const groups = ["第一組", "第二組", "第三組", "第四組", "第五組", "第六組", "第七組", "馬來西亞"];
  const sheetStats = {};
  const totalStats = {};
  
  // 初始化數據
  groups.forEach(g => {
    sheetStats[g] = 0;
    totalStats[g] = 0;
  });

  const availableSheets = bonusSheets.map(s => s.getName());

  if (bonusSheets.length === 0) return { success: true, latest: sheetStats, total: totalStats, sheetName: "", availableSheets: [] };

  // 1. 單表數據 (預設最新，若傳入指定名稱則使用該表)
  let targetSheet = bonusSheets[0];
  if (targetSheetName) {
    const found = bonusSheets.find(s => s.getName() === targetSheetName);
    if (found) targetSheet = found;
  }
  
  const sheetData = targetSheet.getDataRange().getValues();
  if (sheetData.length > 1) {
    sheetData.slice(1).forEach(row => {
      const group = row[1] ? row[1].toString().trim() : "";
      const pts = parseFloat(row[3]) || 0;
      if (groups.indexOf(group) !== -1) sheetStats[group] += pts;
    });
  }

  // 2. 總累積數據 (掃描所有加分表)
  bonusSheets.forEach(s => {
    const d = s.getDataRange().getValues();
    if (d.length > 1) {
      d.slice(1).forEach(row => {
        const group = row[1] ? row[1].toString().trim() : "";
        const pts = parseFloat(row[3]) || 0;
        if (groups.indexOf(group) !== -1) totalStats[group] += pts;
      });
    }
  });

  return { 
    success: true, 
    latest: sheetStats, // 維持 latest 鍵名避免改動太多前端
    total: totalStats, 
    sheetName: targetSheet.getName(),
    availableSheets: availableSheets,
    labels: groups
  };
}

// ==========================================
// 5. 作業管理模組
// ==========================================

function getHomeworkConfig() {
  const cached = getFromCache("hw_config");
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('HomeworkSettings') || ss.insertSheet('HomeworkSettings');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['作業主題', '最低字數', '題目資料JSON', '是否顯示心得', '允許上傳檔案數']);
    sheet.getRange(1,1,1,5).setBackground('#e0e7ff').setFontWeight('bold');
    const defaultQs = JSON.stringify([{ q: "今日學習狀態", opts: ["法喜充滿", "收穫良多", "還在消化"], img: "" }]);
    sheet.appendRow(['預設心得作業', 200, defaultQs, 'Y', 1]);
  }
  const data = sheet.getDataRange().getValues();
  
  // 自動檢查並更新標題列以容納新欄位
  if (data.length > 0) {
    const headers = data[0];
    if (headers.length < 5 || headers.indexOf('允許上傳檔案數') === -1) {
      sheet.getRange(1, 5).setValue('允許上傳檔案數');
      sheet.getRange(1, 1, 1, 5).setBackground('#e0e7ff').setFontWeight('bold');
    }
  }
  
  const config = data.slice(1).map(row => {
    let questions = [];
    try { questions = row[2] ? JSON.parse(row[2].toString()) : []; } catch (e) {}
    return { 
      topic: row[0] != null ? row[0].toString().trim() : "", 
      minChars: parseInt(row[1]) || 0, 
      questions: questions,
      showReflection: row[3] !== 'N', // 預設為 Y
      maxFiles: (row[4] !== undefined && row[4] !== "") ? parseInt(row[4]) : 1
    };
  }).filter(c => c.topic !== "");
  
  putInCache("hw_config", config, 21600); // 快取 6 小時
  return { success: true, data: config };
}

function handleUpdateHomeworkSettings(configStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('HomeworkSettings') || ss.insertSheet('HomeworkSettings');
  const configs = JSON.parse(configStr);
  sheet.clear();
  sheet.appendRow(['作業主題', '最低字數', '題目資料JSON', '是否顯示心得', '允許上傳檔案數']);
  sheet.getRange(1, 1, 1, 5).setBackground('#e0e7ff').setFontWeight('bold');
  configs.forEach(cfg => {
    const maxFiles = cfg.maxFiles !== undefined ? parseInt(cfg.maxFiles) : 1;
    sheet.appendRow([cfg.topic, cfg.minChars, JSON.stringify(cfg.questions || []), cfg.showReflection ? 'Y' : 'N', maxFiles]);
  });
  removeFromCache("hw_config"); // 清除快取
  return { success: true };
}

function handleSubmitHomework(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = '作業_' + (data.topic || '').toString().trim();
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  let headers = ['時間戳記', '組別', '姓名', '作業主題', '心得內容', '參考網址'];
  let answersArr = [];
  if (data.answersJson) { try { answersArr = JSON.parse(data.answersJson); } catch (e) {} }
  answersArr.forEach(ans => { if (headers.indexOf(ans.question) === -1) headers.push(ans.question); });
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#e0e7ff').setFontWeight('bold');
  } else {
    let existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let headerUpdated = false;
    headers.forEach(h => { if (existingHeaders.indexOf(h) === -1) { existingHeaders.push(h); headerUpdated = true; } });
    if (headerUpdated) {
      sheet.getRange(1, 1, 1, existingHeaders.length).setValues([existingHeaders]);
      sheet.getRange(1, 1, 1, existingHeaders.length).setBackground('#e0e7ff').setFontWeight('bold');
    }
  }
  
  const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = new Array(finalHeaders.length).fill('');
  rowData[0] = new Date();
  rowData[1] = (data.group || '').toString().trim();
  rowData[2] = (data.name || '').toString().trim();
  rowData[3] = (data.topic || '').toString().trim();
  rowData[4] = data.reflection || '';
  rowData[5] = data.url || '';
  answersArr.forEach(ans => {
     const colIndex = finalHeaders.indexOf(ans.question);
     if (colIndex !== -1) rowData[colIndex] = ans.answer;
  });
  sheet.appendRow(rowData);
  logAction(data.name, data.group, "繳交作業", `主題: ${data.topic}`, data.name);
  return { success: true, time: Utilities.formatDate(rowData[0], "GMT+8", "yyyy/MM/dd HH:mm:ss") };
}

function getHomeworkReport(topic) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const safeTopic = (topic || '').toString().trim();
    const sheet = ss.getSheetByName('作業_' + safeTopic);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const list = data.slice(1)
      .filter(row => row && (row[1] || row[2]))
      .map(row => ({
        group: row[1] != null ? row[1].toString().trim() : "",
        name: row[2] != null ? row[2].toString().trim() : "",
        reflection: row[4] != null ? row[4].toString() : "",
        url: row[5] != null ? row[5].toString() : "",
        time: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+8", "MM/dd HH:mm") : (row[0] ? row[0].toString() : "")
      }));
    return { success: true, data: list };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getUserHomeworkHistory(group, name) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const history = [];
    const safeGroup = (group || '').toString().trim();
    const safeName = (name || '').toString().trim();

    ss.getSheets().forEach(s => {
      const sName = s.getName();
      if (sName.indexOf('作業_') === 0) {
        const topic = sName.replace('作業_', '');
        const data = s.getDataRange().getValues();
        if (data.length >= 2) {
          const headers = data[0];
          data.slice(1).forEach(row => {
            const rowGroup = row[1] != null ? row[1].toString().trim() : "";
            const rowName = row[2] != null ? row[2].toString().trim() : "";
            if (rowGroup === safeGroup && rowName === safeName) {
              const answers = [];
              if (headers.length > 6) {
                for (let i = 6; i < headers.length; i++) {
                  if (headers[i]) {
                    answers.push({ question: headers[i], answer: row[i] != null ? row[i].toString() : "" });
                  }
                }
              }
              history.push({
                topic: topic,
                time: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+8", "yyyy/MM/dd HH:mm") : (row[0] ? row[0].toString() : ""),
                reflection: row[4] != null ? row[4].toString() : "",
                url: row[5] != null ? row[5].toString() : "",
                answers: answers
              });
            }
          });
        }
      }
    });
    return { success: true, data: history.sort((a,b) => new Date(b.time) - new Date(a.time)) };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ==========================================
// 6. 測驗與成績
// ==========================================

function handleSubmitQuiz(p) {
  const monthStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = '結果' + monthStr;
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  let headers = ['時間戳記', '組別', '姓名', '科目', '分數', '詳情'];
  const ua = JSON.parse(p.userAnswersJson), qd = JSON.parse(p.quizDataJson);
  let score = 0;
  
  const details = ua.map((u, i) => {
    const q = qd[i];
    const isCorrect = u.isBonus ? (u.userAnswer !== '未作答') : [].concat(q.answer).some(a => a.toString().trim().toLowerCase() === u.userAnswer.toString().trim().toLowerCase());
    if (isCorrect) score += (parseFloat(q.score) || 0);
    return { question: q.question, userAnswer: u.userAnswer, correctAnswer: [].concat(q.answer).join(','), isCorrect, photo: q.photo };
  });

  details.forEach((d, index) => {
    const qPrefix = `[Q${index + 1}]`;
    const h1 = `${qPrefix} 正確答案`, h2 = `${qPrefix} 學員作答`, h3 = `${qPrefix} 是否答對`;
    if (headers.indexOf(h1) === -1) headers.push(h1);
    if (headers.indexOf(h2) === -1) headers.push(h2);
    if (headers.indexOf(h3) === -1) headers.push(h3);
  });

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#e0e7ff').setFontWeight('bold');
  } else {
    let existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let headerUpdated = false;
    headers.forEach(h => { if (existingHeaders.indexOf(h) === -1) { existingHeaders.push(h); headerUpdated = true; } });
    if (headerUpdated) {
      sheet.getRange(1, 1, 1, existingHeaders.length).setValues([existingHeaders]);
      sheet.getRange(1, 1, 1, existingHeaders.length).setBackground('#e0e7ff').setFontWeight('bold');
    }
  }

  const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = new Array(finalHeaders.length).fill('');
  rowData[0] = new Date();
  rowData[1] = p.group; rowData[2] = p.name; rowData[3] = p.quizType; rowData[4] = score; rowData[5] = JSON.stringify(details);

  details.forEach((d, index) => {
     const qPrefix = `[Q${index + 1}]`;
     const idx1 = finalHeaders.indexOf(`${qPrefix} 正確答案`), idx2 = finalHeaders.indexOf(`${qPrefix} 學員作答`), idx3 = finalHeaders.indexOf(`${qPrefix} 是否答對`);
     if (idx1 !== -1) rowData[idx1] = d.correctAnswer;
     if (idx2 !== -1) rowData[idx2] = d.userAnswer;
     if (idx3 !== -1) rowData[idx3] = d.isCorrect ? '⭕ 是' : '❌ 否';
  });
  sheet.appendRow(rowData);
  logAction(p.name, p.group, "提交測驗", `科目: ${p.quizType}, 分數: ${score}`, p.name);
  return { status: 'success', score };
}

function getHistory(group, name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const history = [];
  ss.getSheets().forEach(s => {
    if (s.getName().indexOf('結果') === 0) {
      const data = s.getDataRange().getValues();
      data.slice(1).forEach(r => {
        if (r[0] && r[1] == group && r[2] == name) {
          history.push({ timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "yyyy-MM-dd HH:mm") : r[0].toString(), quizType: r[3], score: r[4], details: r[5] });
        }
      });
    }
  });
  return { success: true, data: history.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp)) };
}

// ==========================================
// 7. 系統配置、公告與地點
// ==========================================

function handleGetConfig() {
  const cached = getFromCache("sys_config");
  if (cached) return cached;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.slice(1).forEach(r => { if(r[0]) config[r[0].toString().trim()] = { start: r[1] instanceof Date ? Utilities.formatDate(r[1], "GMT+8", "yyyy-MM-dd HH:mm") : r[1].toString(), end: r[2] ? (r[2] instanceof Date ? Utilities.formatDate(r[2], "GMT+8", "yyyy-MM-dd HH:mm") : r[2].toString()) : "" }; });
  putInCache("sys_config", config, 600); // 快取 10 分鐘
  return config;
}

function handleUpdateConfig(configStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
  const updates = JSON.parse(configStr);
  const data = sheet.getDataRange().getValues();
  for (let key in updates) {
    let found = false;
    for(let i=1; i<data.length; i++) { if(data[i][0] == key) { sheet.getRange(i+1, 2, 1, 2).setValues([[updates[key].start, updates[key].end || ""]]); found = true; break; } }
    if(!found) sheet.appendRow([key, updates[key].start, updates[key].end || ""]);
  }
  removeFromCache("sys_config"); // 清除快取
  return { success: true };
}

function getAnnouncements() {
  const cached = getFromCache("announcements_data");
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Bulletin') || ss.insertSheet('Bulletin');
  const data = sheet.getDataRange().getValues();
  const list = data.slice(1).map((r, i) => ({ id: i+2, time: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+8", "yyyy/MM/dd HH:mm") : r[0].toString(), title: r[1], content: r[2], link: r[3], visible: r[4] || 'Y' })).reverse();
  putInCache("announcements_data", list, 600); // 快取 10 分鐘
  return { success: true, data: list };
}

function updateAnnouncement(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Bulletin');
  const vals = [[new Date(), p.title, p.content, p.link, p.visible]];
  if (p.id) sheet.getRange(parseInt(p.id), 1, 1, 5).setValues(vals); else sheet.appendRow(vals[0]);
  removeFromCache("announcements_data"); // 清除快取
  return { success: true };
}

function getLocationsData() {
  const cached = getFromCache("locations_data");
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Locations') || ss.insertSheet('Locations');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [{ name: "中心地點", lat: 24.13388, lng: 120.71792, radius: 500 }] };
  const locs = data.slice(1).map(row => ({ name: row[0].toString(), lat: parseFloat(row[1]), lng: parseFloat(row[2]), radius: parseFloat(row[3]) || 500 })).filter(loc => !isNaN(loc.lat));
  putInCache("locations_data", locs, 21600); // 快取 6 小時
  return { success: true, data: locs };
}

// ==========================================
// 8. 檔案系統
// ==========================================


/**
 * 改進：自動從 Config 工作表讀取 DRIVE_FOLDER_ID
 */
function handleFileUpload(p) {
  try {
    const config = handleGetConfig();
    const targetFolderId = (config['DRIVE_FOLDER_ID'] && config['DRIVE_FOLDER_ID'].start) ? config['DRIVE_FOLDER_ID'].start : PropertiesService.getScriptProperties().getProperty('V2_UPLOAD_FOLDER_ID');
    if (!targetFolderId) throw Error('尚未設定 DRIVE_FOLDER_ID 或 V2_UPLOAD_FOLDER_ID');
    
    const folder = DriveApp.getFolderById(targetFolderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName);
    const file = folder.createFile(blob);
    //file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: "上傳失敗。請確認資料夾 ID 是否正確，並已共用給後端帳號。錯誤資訊: " + err.toString() };
  }
}

function handleSaveTransferRecord(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('圖傳') || ss.insertSheet('圖傳');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['時間戳記', '組別', '上傳者', '原始檔名', '圖片網址']);
    sheet.getRange(1, 1, 1, 5).setBackground('#f1f5f9').setFontWeight('bold');
  }
  sheet.appendRow([new Date(), p.group, p.name, p.fileName, p.url]);
  return { success: true };
}

// ==========================================
// 9. 額外報表與區域統計
// ==========================================

function getRegionReport() {
  try {
    const targetSSId = '157JOJsLvrUlNVwZ93D4D0awO4KKIyM_G8sbXbKufuK4';
    const ss = SpreadsheetApp.openById(targetSSId);
    const sheet = ss.getSheetByName('區域');
    if (!sheet) return { success: false, error: '找不到工作表: 區域' };
    
    const data = sheet.getRange('F1:O84').getDisplayValues();
    return { success: true, data: data };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}


// ================================================================
// V2 安全與相容層：新增於原有後端，不刪除或清空任何既有工作表
// ================================================================
const V2_TZ = 'Asia/Taipei';
const V2_SESSION_SECONDS = 8 * 60 * 60;
const V2_ADMIN = ['ADMIN','管理員'];
const V2_TEACHER = ['TEACHER','老師','教師'];
const V2_LEADER = ['LEADER','組長'];
function v2Role(role) {
  const s = String(role || '').trim().toUpperCase();
  if (V2_ADMIN.some(x => s === x.toUpperCase())) return 'admin';
  if (V2_TEACHER.some(x => s === x.toUpperCase())) return 'teacher';
  if (V2_LEADER.some(x => s === x.toUpperCase())) return 'leader';
  if (s === '發布員') return 'publisher';
  return 'student';
}
function v2B64(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,''); }
function v2Bytes(s) { return Utilities.newBlob(String(s)).getBytes(); }
function v2Secret() {
  const p = PropertiesService.getScriptProperties();
  let secret = p.getProperty('V2_SESSION_SECRET');
  if (!secret) {
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try { secret = p.getProperty('V2_SESSION_SECRET'); if (!secret) { secret = Utilities.getUuid()+Utilities.getUuid(); p.setProperty('V2_SESSION_SECRET',secret); } }
    finally { lock.releaseLock(); }
  }
  return secret;
}
function v2Digest(password,salt) {
  // SHA-256 with per-account salt; upgrade to dedicated identity provider for high-risk deployments.
  let v = String(salt)+':'+String(password);
  for (let i=0;i<1500;i++) v = v2B64(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,v));
  return 'v2$'+salt+'$'+v;
}
function v2Compare(a,b) {
  a=String(a); b=String(b);
  let d=a.length ^ b.length;
  for (let i=0;i<Math.max(a.length,b.length);i++) d |= (a.charCodeAt(i)||0) ^ (b.charCodeAt(i)||0);
  return d===0;
}
function v2Find(group,name) {
  if(v3UsersEnabled()) {
    const data=v3GetUser(group,name);
    return data?{data:data,source:'script-properties'}:null;
  }
  const sh=v2UsersSheet(SpreadsheetApp.openById(SPREADSHEET_ID)),values=sh.getDataRange().getValues();
  const i=values.findIndex((r,j)=>j>0&&String(r[0]).trim()===String(group).trim()&&String(r[1]).trim()===String(name).trim());
  return i<0?null:{sheet:sh,row:i+1,data:values[i],source:'sheet'};
}

function v2Issue(user) {
  const payload=v2B64(v2Bytes(JSON.stringify({g:String(user.group),n:String(user.name),exp:Math.floor(Date.now()/1000)+V2_SESSION_SECONDS})));
  const sig=v2B64(Utilities.computeHmacSha256Signature(payload,v2Secret()));
  return payload+'.'+sig;
}
function v2Session(token) {
  const parts=String(token||'').split('.');
  if(parts.length!==2) throw Error('請重新登入');
  const expected=v2B64(Utilities.computeHmacSha256Signature(parts[0],v2Secret()));
  if(!v2Compare(expected,parts[1])) throw Error('登入憑證無效');
  const payload=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if(!payload.exp || payload.exp<Math.floor(Date.now()/1000)) throw Error('登入已逾時，請重新登入');
  const found=v2Find(payload.g,payload.n);
  if(!found || !found.data[2]) throw Error('使用者不存在或未註冊');
  return {group:String(found.data[0]),name:String(found.data[1]),role:String(found.data[3]||'學員'),level:v2Role(found.data[3])};
}
function v2Require(user,roles) { if(!roles.includes(user.level)) throw Error('權限不足，無法執行此操作'); }
function v2Register(data) {
  const code=PropertiesService.getScriptProperties().getProperty('V2_REGISTRATION_CODE');
  if(!code) return {success:false,error:'首次註冊尚未啟用，請管理員先設定 V2_REGISTRATION_CODE'};
  if(!v2Compare(String(data.registrationCode||''),code)) return {success:false,error:'註冊驗證碼錯誤'};
  const pw=String(data.password||'');
  if(pw.length<8 || pw.length>128) return {success:false,error:'密碼請設定 8 至 128 個字元'};
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const u=v2Find(data.group,data.name);
    if(!u) return {success:false,error:'名單查無此人'};
    if(String(u.data[2]||'').trim()) return {success:false,error:'此帳號已註冊'};
    if(v3UsersEnabled()) {
      const updated=u.data.slice();
      updated[2]=pw; updated[4]=new Date().toISOString();
      v3SaveUser(updated);
    } else {
      u.sheet.getRange(u.row,3).setValue(pw);
      u.sheet.getRange(u.row,5).setValue(new Date());
    }
    logAction(data.name,data.group,'註冊','V2 安全註冊',data.name);
    return {success:true};
  } finally { lock.releaseLock(); }
}
function v2Login(data) {
  const u=v2Find(data.group,data.name);
  if(!u) return {success:false,error:'帳號或密碼錯誤'};
  const pw=String(data.password||''), stored=String(u.data[2]||'');
  // 保留原有 Users C 欄密碼格式，不自動轉換、不改寫。
  // 兼容先前測試版可能已產生的 v2$ 密碼紀錄。
  let valid=false;
  if(stored.startsWith('v2$')) {
    const parts=stored.split('$');
    valid=parts.length===3 && v2Compare(v2Digest(pw,parts[1]),stored);
  } else {
    valid=!!stored && v2Compare(stored,pw);
  }
  if(!valid) return {success:false,error:'帳號或密碼錯誤'};
  const user={name:String(u.data[1]),group:String(u.data[0]),role:String(u.data[3]||'學員')};
  logAction(user.name,user.group,'登入','V2 登入成功',user.name);
  return {success:true,...user,token:v2Issue(user)};
}
function v2WithLock(fn) {
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {return fn();} finally {lock.releaseLock();}
}
function v2SignIn(p,user) {
  return v2WithLock(()=>{
    const group=String(p.group||user.group),name=String(p.name||user.name);
    const proxy=(group!==user.group || name!==user.name || !!p.leaveType);
    if(proxy) v2Require(user,['admin','teacher','leader']);
    if(user.level==='leader' && group!==user.group) throw Error('組長只能處理自己的組別');
    if(!v2Find(group,name)) throw Error('簽到名單不存在');
    const status=checkAttendanceStatus(group,name);
    if(status.signedIn) return {success:false,error:'今日已有簽到紀錄，請勿重複提交'};
    return legacyHandleSignIn({...p,group,name,admin:user.name});
  });
}
// V3.2 使用者管理：僅管理員可透過已驗證的工作階段操作。
// 只回傳公開管理欄位，永不回傳密碼或雜湊。
function v3UserPublic(r) {
  return {group:String(r[0]||''),name:String(r[1]||''),role:String(r[3]||'學員'),registeredAt:String(r[4]||''),registered:!!String(r[2]||'')};
}
function v3AdminUsersList(user) {
  v2Require(user,['admin']);
  if(!v3UsersEnabled()) throw Error('請先完成 USERS 遷移，才能使用帳號管理');
  return {success:true,data:v3AllUsers().map(v3UserPublic).sort((a,b)=>a.group.localeCompare(b.group,'zh-Hant')||a.name.localeCompare(b.name,'zh-Hant'))};
}
// V3.3：USERS/Users A、B 欄增量同步。C/D/E 永不讀取或覆寫。
// 同步只新增 GS 中不存在的 (組別,姓名)；既有記錄完全不變。
function v3GetImportSheet() {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh=ss.getSheetByName('USERS')||ss.getSheetByName('Users');
  if(!sh) {
    sh=ss.insertSheet('USERS');
    sh.getRange(1,1,1,2).setValues([['組別','姓名']]);
  }
  return sh;
}
function v3ImportABPreview() {
  if(!v3UsersEnabled()) throw Error('請先完成 USERS 帳號遷移');
  const sh=v3GetImportSheet(), last=sh.getLastRow();
  const rows=last>1?sh.getRange(2,1,last-1,2).getDisplayValues():[];
  const props=PropertiesService.getScriptProperties().getProperties();
  const seen=new Set(),pending=[],duplicates=[],invalid=[];
  let existing=0,blank=0;
  rows.forEach((r,i)=>{
    const group=String(r[0]||'').trim(),name=String(r[1]||'').trim(),row=i+2;
    if(!group&&!name){blank++;return;}
    if(!group||!name){invalid.push(row);return;}
    if(group==='組別'&&name==='姓名') {invalid.push(row);return;}
    if(group.length>80||name.length>80){invalid.push(row);return;}
    let key;
    try {key=v3UserKey(group,name);} catch(e) {invalid.push(row);return;}
    if(seen.has(key)){duplicates.push(row);return;}
    seen.add(key);
    if(Object.prototype.hasOwnProperty.call(props,key)) {existing++;return;}
    pending.push({row,group,name});
  });
  return {success:true,sheet:sh.getName(),scanned:rows.length,existing,blank,duplicates,invalid,
    newCount:pending.length,pending:pending.slice(0,30),_pending:pending};
}
function v3SyncImportAB() {
  return v2WithLock(()=>{
    const plan=v3ImportABPreview();
    if(plan.invalid.length||plan.duplicates.length) throw Error('A/B 名單有不完整或重複資料，請先修正。缺漏列：'+plan.invalid.join(',')+'；重複列：'+plan.duplicates.join(','));
    // 容量及筆數保守限制；超過時不寫入任何新資料。
    if(plan._pending.length>500) throw Error('單次新增超過 500 筆，請分批匯入');
    const props=PropertiesService.getScriptProperties();
    const original=props.getProperties();
    const toWrite={};
    plan._pending.forEach(p=>{
      const key=v3UserKey(p.group,p.name);
      if(Object.prototype.hasOwnProperty.call(original,key)) return;
      toWrite[key]=JSON.stringify([p.group,p.name,'','學員','']);
    });
    const currentBytes=Object.keys(original).reduce((n,k)=>n+Utilities.newBlob(k+original[k]).getBytes().length,0);
    const addedBytes=Object.keys(toWrite).reduce((n,k)=>n+Utilities.newBlob(k+toWrite[k]).getBytes().length,0);
    if(currentBytes+addedBytes>420000) throw Error('Script Properties 儲存容量接近上限，請停止匯入並評估資料庫');
    if(Object.keys(toWrite).length) props.setProperties(toWrite,false);
    removeFromCache('teams_data');
    return {success:true,sheet:plan.sheet,added:Object.keys(toWrite).length,skippedExisting:plan.existing,blank:plan.blank,message:'已新增未註冊帳號，既有密碼、權限與註冊日期完全未修改'};
  });
}
function v3AdminImportPreview(user) {
  v2Require(user,['admin']);
  const p=v3ImportABPreview();
  delete p._pending;
  return p;
}
// Apps Script 編輯器也可手動執行，須由有專案編輯權的人操作。
function syncUsersABToScriptProperties() {
  const result=v3SyncImportAB();
  Logger.log(JSON.stringify(result));
  return result;
}
function v3AdminUserMutation(p,user) {
  v2Require(user,['admin']);
  if(!v3UsersEnabled()) throw Error('請先完成 USERS 遷移');
  const action=String(p.action||'');
  const group=String(p.group||'').trim(),name=String(p.name||'').trim();
  if(!group||!name||group.length>80||name.length>80) throw Error('組別及姓名必填，且不得超過 80 字');
  const role=String(p.role||'學員').trim();
  if(!['ADMIN','TEACHER','LEADER','發布員','學員','管理員','老師','教師','組長'].includes(role)) throw Error('不支援的使用者權限');
  return v2WithLock(()=>{
    const existing=v3GetUser(group,name);
    const props=PropertiesService.getScriptProperties();
    if(action==='adminCreateUser') {
      if(existing) throw Error('此組別與姓名已存在');
      const password=String(p.password||'');
      if(password && (password.length<8||password.length>128)) throw Error('初始密碼須為 8 至 128 字');
      const stored=password?v2Digest(password,Utilities.getUuid()):'';
      const record=[group,name,stored,role,password?new Date().toISOString():''];
      v3SaveUser(record);
    } else if(action==='adminUpdateUser') {
      const oldGroup=String(p.oldGroup||'').trim(),oldName=String(p.oldName||'').trim();
      const old=v3GetUser(oldGroup,oldName);
      if(!old) throw Error('找不到原使用者');
      if(oldGroup===user.group&&oldName===user.name && (group!==oldGroup||name!==oldName||v2Role(role)!=='admin')) throw Error('不能變更自己的帳號識別或取消自己的管理員權限');
      if((oldGroup!==group||oldName!==name)&&existing) throw Error('新組別與姓名已有帳號');
      const updated=old.slice();updated[0]=group;updated[1]=name;updated[3]=role;
      v3SaveUser(updated);
      if(oldGroup!==group||oldName!==name) props.deleteProperty(v3UserKey(oldGroup,oldName));
    } else if(action==='adminResetUserPassword') {
      if(!existing) throw Error('找不到使用者');
      const password=String(p.password||'');
      if(password.length<8||password.length>128) throw Error('新密碼須為 8 至 128 字');
      const updated=existing.slice();
      updated[2]=v2Digest(password,Utilities.getUuid());
      updated[4]=updated[4]||new Date().toISOString();
      v3SaveUser(updated);
    } else throw Error('不支援的管理操作');
    removeFromCache('teams_data');
    return {success:true,message:'使用者資料已更新'};
  });
}
function v2Get(e) {
  try {
    const p=e&&e.parameter||{}, action=String(p.action||'');
    const publicActions=['getTeamsData'];
    const session=publicActions.includes(action)?null:v2Session(p.token);
    if(action==='adminListUsers') return JSONResponse(v3AdminUsersList(session));
    if(action==='adminPreviewABImport') return JSONResponse(v3AdminImportPreview(session));
    if(action==='validateSession') return JSONResponse({success:true,name:session.name,group:session.group,role:session.role});
    if(action==='getQuizQuestions') {
      v2Require(session,['admin','teacher','leader','student','publisher']);
      return JSONResponse({success:true,data:v2QuizQuestions(String(p.quizType||''))});
    }
    if(action==='getAttendanceReport') {v2Require(session,['admin','teacher','leader']); p.role=session.role;p.group=session.group;}
    if(action==='getHomeworkReport') {
      v2Require(session,['admin','teacher','leader']);
      // 先以登入憑證判定權限，再於伺服器端過濾，不能信任前端傳來的組別。
      if(session.level==='leader') {
        const report=getHomeworkReport(p.topic);
        if(report.success) report.data=(report.data||[]).filter(row=>String(row.group||'').trim()===String(session.group||'').trim());
        return JSONResponse(report);
      }
    }
    if(action==='getBonusStats') v2Require(session,['admin','teacher','leader']);
    // 測驗開放時間由所有已登入使用者共用；不能只允許管理員讀取 Config。
    if(action==='getRegionReport') v2Require(session,['admin','teacher']);
    if(action==='getConfig') v2Require(session,['admin','teacher','leader','student','publisher']);
    if(['getHistory','getUserHomeworkHistory','checkAttendance'].includes(action)) {
      const targetGroup=String(p.group||''),targetName=String(p.name||'');
      if(targetGroup!==session.group || targetName!==session.name) {
        v2Require(session,['admin','teacher','leader']);
        if(session.level==='leader' && targetGroup!==session.group) throw Error('不可查看其他組別');
      }
    }
    if(action==='getLocations') v2Require(session,['admin','teacher','leader','student','publisher']);
    return legacyDoGet({parameter:p});
  } catch(err) {return JSONResponse({success:false,error:err.message||String(err)});}
}

function v2QuizQuestions(type) {
  const ids={'隨堂':0,'論語':1250135497,'性理':1342680959,'道德經':1246480043,'六祖':1302889740};
  if(!Object.prototype.hasOwnProperty.call(ids,type)) throw Error('不支援的測驗科目');
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  function read(gid) {
    const sh=ss.getSheets().find(x=>x.getSheetId()===gid);
    if(!sh) throw Error('找不到題庫工作表 gid='+gid);
    return sh.getDataRange().getValues().map(r=>{
      const cells=r.map(x=>String(x===null?'':x));
      const c0=(cells[0]||'').trim(),c1=(cells[1]||'').trim();
      const offset=(c0.length>0 && c0.length<=5 && /^[Qq\d\.\s]+$/.test(c0) && c1.length>c0.length)?1:0;
      const question=(cells[offset]||'').trim();
      if(!question || question==='題目') return null;
      const split=x=>String(x||'').split(/[,，\n;]+/).map(v=>v.trim()).filter(Boolean);
      return {question,type:(cells[offset+1]||'single').toLowerCase().trim(),options:split(cells[offset+2]),answer:split(cells[offset+3]),score:parseFloat(cells[offset+4])||0,
        photo:(cells[offset+5]||'').trim(),isBonus:['是','yes','true','y','1'].includes((cells[offset+6]||'').toLowerCase().trim())};
    }).filter(Boolean);
  }
  return [...read(ids['隨堂']),...read(ids[type])];
}
function v2VerifyQuiz(p) {
  const sent=JSON.parse(p.quizDataJson||'[]'),answers=JSON.parse(p.userAnswersJson||'[]');
  const actual=v2QuizQuestions(String(p.quizType||''));
  if(!Array.isArray(sent)||!Array.isArray(answers)||sent.length!==actual.length||answers.length!==actual.length) throw Error('題庫已更新，請重新載入考卷');
  actual.forEach((q,i)=>{
    const a=sent[i];
    if(String(a.question)!==q.question || Number(a.score)!==q.score ||
       JSON.stringify([].concat(a.answer).map(String))!==JSON.stringify(q.answer) ||
       Boolean(a.isBonus)!==q.isBonus) throw Error('題目資料驗證失敗，請重新載入考卷');
  });
  // 不再信任前端提供的正解、配分、加分旗標與圖片網址。
  p.quizDataJson=JSON.stringify(actual);
  p.userAnswersJson=JSON.stringify(answers.map((a,i)=>({...a,isBonus:actual[i].isBonus})));
}
function v2Post(e) {
  try {
    const p=JSON.parse(e.postData.contents||'{}'),action=String(p.action||'');
    if(action==='login') return JSONResponse(v2Login(p));
    if(action==='register') return JSONResponse(v2Register(p));
    const user=v2Session(p.token);
    if(action==='adminSyncABImport') {v2Require(user,['admin']);return JSONResponse(v3SyncImportAB());}
    if(['adminCreateUser','adminUpdateUser','adminResetUserPassword'].includes(action)) return JSONResponse(v3AdminUserMutation(p,user));
    const admin=['admin'],staff=['admin','teacher'],lead=['admin','teacher','leader'];
    const permissions={updateConfig:admin,updateAnnouncement:[...staff,'publisher'],updateHomeworkSettings:staff,createBonusSheet:admin,addBonusPoints:lead,uploadFile:['admin','teacher','leader','publisher','student'],saveTransferRecord:['admin','teacher','leader','publisher','student'],signIn:lead};
    if(permissions[action] && action!=='signIn') v2Require(user,permissions[action]);
    if(action==='signIn') return JSONResponse(v2SignIn(p,user));
    if(action==='addBonusPoints' && user.level==='leader' && String(p.group)!==user.group) throw Error('組長只能為自己組別加分');
    if(['submitHomework','submitQuiz'].includes(action)) {
      if(String(p.group)!==user.group || String(p.name)!==user.name) throw Error('只能提交自己的資料');
    }
    if(['addBonusPoints','saveTransferRecord'].includes(action)) p.admin=user.name;
    if(action==='submitQuiz') v2VerifyQuiz(p);
    if(action==='createBonusSheet') return JSONResponse(v2WithLock(()=>handleCreateBonusSheet(user.name)));
    if(['addBonusPoints','submitHomework','submitQuiz','updateConfig','updateAnnouncement','updateHomeworkSettings'].includes(action))
      return v2WithLock(()=>legacyDoPost({postData:{contents:JSON.stringify(p)}}));
    return legacyDoPost({postData:{contents:JSON.stringify(p)}});
  } catch(err) {return JSONResponse({success:false,error:err.message||String(err)});}
}
function doGet(e) {return v2Get(e);}
function doPost(e) {return v2Post(e);}