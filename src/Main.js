/**
 * Main.gs - 伊美系統 v10.3.7 入口與控制器
 * 
 * 職責：
 * 1. 處理 Web App 請求 (doGet)
 * 2. 路由分發
 * 3. 初始與自動化介面 (onOpen, showControlCenter)
 */

/**
 * 當試算表開啟時自動執行：建立自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 伊美系統')
    .addItem('📱 開啟控制台 (連結/QR)', 'showControlCenter')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ 系統初始化作業')
      .addItem('🔒 步驟 1: 設定管理密碼', 'setupAdminPasswordUI')
      .addItem('🧠 步驟 2: 配置 AI 大腦', 'setupGeminiApiKeyUI'))
    .addSeparator()
    .addItem('⚙️ 進入管理後台', 'directToAdmin')
    .addToUi();
}

/**
 * [UI] 初始化管理密碼 (免改程式碼)
 */
function setupAdminPasswordUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔒 系統安全性設定', '請輸入新的管理員密碼（至少 8 字元）：\n(此設定會加密儲存，系統不會保留明文)', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const pwd = response.getResponseText();
    try {
      Service_Security.setAdminPassword(pwd);
      ui.alert('✅ 設定成功', '管理員密碼已安全加密並儲存完畢！', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('❌ 設定失敗', e.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * [UI] 配置 Gemini API Key (免改專案設定)
 */
function setupGeminiApiKeyUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🧠 AI 大腦連線設定', '請貼入您的 Gemini API Key：', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const key = response.getResponseText();
    if (key && key.trim().length > 10) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key.trim());
      ui.alert('✅ 連線成功', 'AI 大腦已經接通，您可以開始使用自動回饋功能了！', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ 設定無效', '請輸入正確的 API Key。', ui.ButtonSet.OK);
    }
  }
}

/**
 * 彈出指揮中心對話框 (v10.3.7 升級版)
 * 優先順序重新調整：① 試算表 M2/M3 (尊重使用者設定) → ② 自動抓取並轉換 /dev 為 /exec
 */
function showControlCenter() {
  var portalUrl = "";
  var adminUrl  = "";
  var status    = "offline";

  // 第一層：優先從試算表讀取 (唯一真理)
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = ss.getSheetByName("系統設定");
    if (cfg) {
      var m2 = cfg.getRange("M2").getValue();
      var m3 = cfg.getRange("M3").getValue();
      if (m2 && String(m2).indexOf('http') > -1) {
        portalUrl = String(m2).trim();
        // 如果手動亂填填到 dev，幫他轉成 exec
        if (portalUrl.indexOf('/dev') > -1) {
          portalUrl = portalUrl.replace('/dev', '/exec');
          cfg.getRange("M2").setValue(portalUrl); // 順手修復
        }
        adminUrl  = m3 ? String(m3).trim() : (portalUrl + (portalUrl.indexOf('?') > -1 ? '&page=admin' : '?page=admin'));
        status    = "online";
      }
    }
  } catch(e) {
    console.warn("讀取 M2/M3 失敗: " + e.message);
  }

  // 第二層：如果試算表沒設定或是設定錯誤
  if (!portalUrl) {
    try {
      var serviceUrl = ScriptApp.getService().getUrl();
      if (serviceUrl && serviceUrl.indexOf('http') > -1) {
        // [核心修復] 將 /dev 網址轉換為 /exec 正式網址
        portalUrl = serviceUrl.replace('/dev', '/exec');
        adminUrl  = portalUrl + (portalUrl.indexOf('?') > -1 ? '&page=admin' : '?page=admin');
        status    = "online";
        
        // 成功抓取且轉換後，寫入試算表快取，達成全自動化
        try {
          var ss2 = SpreadsheetApp.getActiveSpreadsheet();
          var cfg2 = ss2.getSheetByName("系統設定");
          if (cfg2) {
            cfg2.getRange("M2").setValue(portalUrl);
            cfg2.getRange("M3").setValue(adminUrl);
          }
        } catch(e2) { /* 忽略 */ }
      }
    } catch(e) {
      console.warn("getUrl 失敗: " + e.message);
    }
  }

  // 第三層：若都失敗，顯示友善提示
  if (!portalUrl) {
    portalUrl = "⚠️ 尚未設定正式網址 — 請至「部署 -> 管理部署」複製網址，貼上到系統設定分頁 M2。";
    adminUrl  = "⚠️ 尚未設定正式網址";
    status    = "offline";
  }

  var tpl = HtmlService.createTemplateFromFile('UI_Console');
  tpl.portalUrl = portalUrl;
  tpl.adminUrl  = adminUrl;
  tpl.status    = status;

  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(420).setHeight(620).setTitle('伊美系統：智慧管理控制台'),
    ' '
  );
}

/**
 * 快速跳转至後台 (輔助函式)
 */
function directToAdmin() {
  var url = "";
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = ss.getSheetByName("系統設定");
    if (cfg) {
      var m3 = cfg.getRange("M3").getValue();
      if (m3 && String(m3).indexOf('http') > -1) {
        url = String(m3).trim();
        // 對於 /dev 網址的防呆修復
        if (url.indexOf('/dev') > -1) {
          url = url.replace('/dev', '/exec');
          cfg.getRange("M3").setValue(url);
        }
      }
    }
  } catch(e) {}

  if (!url) {
    try {
      const serviceUrl = ScriptApp.getService().getUrl();
      if (serviceUrl) {
        url = serviceUrl.replace('/dev', '/exec');
        url = url + (url.indexOf('?') > -1 ? '&page=admin' : '?page=admin');
      }
    } catch(e) {}
    
    if (!url) {
      SpreadsheetApp.getUi().alert("系統尚未設定網址，請先至「部署 -> 管理部署」複製正式網址，貼到「系統設定」M2 儲存格");
      return;
    }
  }

  const html = `<script>window.open("${url}", "_blank"); google.script.host.close();</script>正在轉向管理中心...`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(100), ' ');
}

/**
 * 處理 Web App HTTP GET 請求
 * @param {Object} e - 事件物件
 * @returns {HtmlOutput}
 */
/**
 * 處理 Web App HTTP GET 請求
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const route = params.route || params.page || 'portal';

  // --- 關鍵優化：自動偵測初始化狀態 ---
  const masterId = PropertiesService.getScriptProperties().getProperty('MASTER_ID');
  const pwdHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
  
  if (!masterId || !pwdHash) {
    if (route === 'setup') {
      return _renderSetup(e);
    }
    // 強制進入安裝精靈，除非已經在安裝頁
    return _renderSetup(e);
  }
  
  if (route === 'portal') {
    return _renderPortal(e);
  } else if (route === 'admin') {
    return _renderAdmin(e);
  } else if (route === 'setup') {
    return _renderSetup(e);
  } else {
    return HtmlService.createHtmlOutput("未知路由");
  }
}

/**
 * 渲染安裝精靈 (v10.3.7)
 */
function _renderSetup(e) {
  return HtmlService.createTemplateFromFile('UI_Setup')
      .evaluate()
      .setTitle('系統安裝精靈 - 伊美：簡報同步互動學習系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 渲染管理後台
 */
function _renderAdmin(e) {
  return HtmlService.createTemplateFromFile('UI_Admin')
      .evaluate()
      .setTitle('後台管理 - 伊美：簡報同步互動學習系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 渲染學員入口頁面
 */
function _renderPortal(e) {
  // 註：安全性與權限檢查已整合於 Service_Security 與 Service_Engine 中。
  const template = HtmlService.createTemplateFromFile('UI_Portal');
  template.title = '伊美：簡報同步互動學習系統 v10.3.7 (Multi-Activity)';
  
  // 固定顯示正確版號，防止試算表名稱過舊導致誤導
  template.activityName = 'emedu-Slides-Sync-Interactive-System - v10.3.7';
  
  return template.evaluate()
      .setTitle('伊美：簡報同步互動學習系統 v10.3.7')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * [API] 終極一鍵安裝 (由 UI_Setup 呼叫)
 */
function apiFinalizeSetup(config) {
  console.log("[Setup] 開始全系統自動安裝...");
  try {
    // 1. 設定管理密碼
    if (config.password) {
      Service_Security.setAdminPassword(config.password);
    } else {
      throw new Error("請提供管理員密碼");
    }

    // 2. 配置 Gemini API Key
    if (config.apiKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', config.apiKey.trim());
    }

    // 3. 建立資料庫與主控台
    const url = Service_DB.installControlSheet();
    
    return {
      status: 'success',
      message: '系統安裝成功！',
      spreadsheetUrl: url
    };
  } catch (e) {
    console.error("[Setup] 失敗: " + e.toString());
    return { status: 'error', message: "安裝失敗：" + e.message };
  }
}

/**
 * 🆘 [緊急救援] 強制跳過 UI 完成初始化
 * 如果網頁安裝精靈持續空白或失敗，請在 GAS 編輯器執行此函式。
 */
function setupSystem_EMERGENCY_FORCE() {
  console.log("🚀 啟動緊急強制初始化...");
  try {
    // 預設密碼：admin123 (請務必在進入後台後修改)
    Service_Security.setAdminPassword("admin123456"); 
    console.log("1. 密碼已強制重置為: admin123456");

    const url = Service_DB.installControlSheet();
    console.log("2. 試算表主控台已建立: " + url);
    console.log("✅ 緊急初始化完成！現在請重新整理 Web App 網頁，您應該能看到登入畫面。");
  } catch (e) {
    console.error("❌ 緊急初始化失敗: " + e.message);
  }
}

/**
 * 🟢 [手動執行] 系統初始化 (保留作為備用地底工具)
 */
function setupSystem() {
  console.log("正在建立主控台與資料表...");
  try {
    const url = Service_DB.installControlSheet();
    console.log("✅ 系統初始化成功！");
    console.log("試算表網址: " + url);
  } catch (e) {
    console.error("❌ 初始化失敗: " + e.toString());
  }
}

/**
 * 🌐 [手動執行] 偵測並同步系統網址 (v10.3.7 自動轉換修復版)
 * 
 * 邏輯：
 * 1. 優先讀取試算表 M2/M3 (如果使用者已經手動貼上，絕對不覆蓋)
 * 2. 如果試算表是空的，嘗試自動抓取，若抓到 /dev 網址，自動轉換為 /exec 正式網址
 */
function getSystemUrls() {
  console.log("══════════════════════════════════════");
  console.log("🔍 伊美系統 v10.3.7 — 網址偵測與同步");
  console.log("══════════════════════════════════════\n");

  var webAppUrl = "";
  var adminUrl  = "";
  var masterId  = Service_DB.getMasterId();
  var isFromSpreadsheet = false;

  // --- 第 1 步：優先讀取試算表 (尊重使用者手動貼上的網址) ---
  if (masterId) {
    try {
      var ss = SpreadsheetApp.openById(masterId);
      var cfg = ss.getSheetByName("系統設定");
      if (cfg) {
        var m2 = cfg.getRange("M2").getValue();
        var m3 = cfg.getRange("M3").getValue();
        if (m2 && String(m2).indexOf("http") > -1) {
          webAppUrl = String(m2).trim();
          if (webAppUrl.indexOf('/dev') > -1) {
             console.log("ℹ️ 偵測到 /dev 開發者網址，系統將自動修復為 /exec 正式網址");
             webAppUrl = webAppUrl.replace('/dev', '/exec');
          }
          adminUrl  = m3 ? String(m3).trim() : (webAppUrl + "?page=admin");
          if (adminUrl.indexOf('/dev') > -1) adminUrl = adminUrl.replace('/dev', '/exec');
          
          isFromSpreadsheet = true;
          console.log("ℹ️ 已讀取試算表內設定的自訂網址");
        }
      }
    } catch(e) { }
  }

  // --- 第 2 步：如果試算表是空的，才自動抓取 ---
  if (!webAppUrl) {
    try {
      var serviceUrl = ScriptApp.getService().getUrl();
      if (serviceUrl && serviceUrl.indexOf("http") > -1) {
        webAppUrl = serviceUrl.replace('/dev', '/exec');
        adminUrl  = webAppUrl + "?page=admin";
        console.log("ℹ️ 已從系統部署自動偵測網址，並已轉換為正式版");
      }
    } catch(e) { }
  }

  // --- 第 3 步：印出結果 ---
  if (webAppUrl && webAppUrl.indexOf("http") > -1) {
    console.log("📱 學員入口 (前台)：");
    console.log("   " + webAppUrl);
    console.log("");
    console.log("⚙️ 管理中心 (後台)：");
    console.log("   " + adminUrl);
  } else {
    console.log("⚠️ 尚未偵測到正式 Web App 網址。");
    console.log("   請先完成以下步驟：");
    console.log("   1. 點擊編輯器右上角 [部署] → [管理部署]");
    console.log("   2. 複製結尾為 /exec 的網址");
    console.log("   3. 將得到的網址貼入試算表 M2 儲存格！");
  }

  console.log("");

  // --- 第 4 步：印出試算表網址 ---
  if (masterId) {
    console.log("📊 控制主控台 (試算表)：");
    console.log("   https://docs.google.com/spreadsheets/d/" + masterId);
  }

  // --- 第 5 步：自動回填 (如果原本有錯或為空，覆寫/回填) ---
  if (webAppUrl && masterId) {
    try {
      var ss2 = SpreadsheetApp.openById(masterId);
      var cfg2 = ss2.getSheetByName("系統設定");
      if (cfg2) {
        cfg2.getRange("M2").setValue(webAppUrl);
        cfg2.getRange("M3").setValue(adminUrl);
        if (!isFromSpreadsheet) {
           console.log("\n📝 已將偵測到的正式網址自動填入試算表 M2/M3 儲存格！");
        }
      }
    } catch(e) {
      console.log("\n⚠️ 自動填入試算表失敗: " + e.message);
    }
  }

  console.log("\n══════════════════════════════════════");
  console.log("✅ 執行完畢。未來如有變動，修改試算表 M2/M3 即可生效。");
  console.log("══════════════════════════════════════");
}

/**
 * 🔧 [修復工具] 強制將 MASTER_ID 指向當前綁定試算表
 * 
 * 使用時機：後台「開啟主控台」連結到錯誤試算表時，
 * 請在 GAS 編輯器選取此函式並執行。
 * 
 * 執行後請查看「執行記錄」確認新舊 ID。
 */
function fixMasterId() {
  console.log("🔧 [fixMasterId] 開始修復 MASTER_ID...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error("❌ 找不到綁定的試算表，請確認此腳本是從試算表的「擴充功能 > Apps Script」開啟的。");
      return;
    }

    const correctId  = ss.getId();
    const correctUrl = ss.getUrl();
    const oldId      = PropertiesService.getScriptProperties().getProperty('MASTER_ID') || "(空)";

    console.log("📋 舊 MASTER_ID: " + oldId);
    console.log("✅ 新 MASTER_ID: " + correctId);
    console.log("🔗 試算表網址 :  " + correctUrl);

    PropertiesService.getScriptProperties().setProperty('MASTER_ID', correctId);

    // 同步清除快取，確保後續 API 讀到最新設定
    try {
      Service_DB.clearActivityConfigCache(correctId);
    } catch(e) { /* 忽略 */ }

    console.log("✅ MASTER_ID 已成功修正！請重新整理後台頁面，「開啟主控台」連結即可正常。");
  } catch(e) {
    console.error("❌ fixMasterId 失敗: " + e.toString());
  }
}

// --- 公開 API 供前端 google.script.run 呼叫 ---

/**
 * API: 學生登入並取得題目
 */
function apiLogin(studentId) {
  console.log(`[API] 學員登入開始: ${studentId}`);
  if (!studentId || studentId.trim() === "") {
    return { status: 'error', message: '請輸入學號才能開始學習唷！' };
  }
  
  try {
    // v10.4.0: 使用當前活動的試算表 ID，而非固定主控台
    const ssid = Service_DB.getActiveSSId() || Service_DB.getMasterId();
    console.log(`[API] Active SS ID: ${ssid}`);
    if (!ssid) throw new Error("系統主控台尚未安裝，請聯絡系統管理員。");
    
    // 直接呼叫引擎取得下一題
    const nextTaskResult = Service_Engine.getStudentNextTask(ssid, studentId);
    console.log(`[API] 下一步任務結果: ${JSON.stringify(nextTaskResult)}`);
    
    if (nextTaskResult.status === 'completed') {
       return {
         status: 'success',
         completed: true,
         message: "恭喜！您已完成所有課程任務。"
       };
    }
    
    // 獲取所有題目以計算進度 (stats)
    const allQs = Service_DB.getActivityConfig(ssid);
    const q = nextTaskResult.question;

    if (!q) {
      throw new Error("尚未設定題目");
    }
    
    const result = {
      status: 'success',
      datestamp: new Date().toISOString(),
      task: {
        stage: q.label,
        question: q.question,
        type: q.type,
        desc: q.helpText || "(無提示)"
      },
      stats: {
        totalQuestions: allQs.length,
        currentOrder: allQs.findIndex(item => item.question === q.question) + 1
      }
    };
    
    console.log(`[API] 登入結果回傳成功: ${studentId}`);
    return result;
  } catch (e) {
    console.error(`[API] 登入失敗意外錯誤: ${e.toString()}`);
    return { status: 'error', message: "系統發生錯誤: " + e.toString() }; 
  }
}


/**
 * API: 管理員登入與資料獲取
 */
function apiAdminLogin(password) {
  try {
    if (!Service_Security.verifyAdmin(password)) {
      return { status: 'error', message: '密碼錯誤' };
    }
    
    // 登入成功，獲取儀表板數據
    const masterId = Service_DB.getMasterId();
    // v10.4.0: 統計資料從當前活動試算表讀取
    const activeSSId = Service_DB.getActiveSSId() || masterId;
    const mockStat = {
      totalStudents: 0,
      totalSubmissions: 0,
      masterUrl: masterId ? `https://docs.google.com/spreadsheets/d/${masterId}` : "#"
    };
    
    // 如果有試算表，嘗試取得真實數據（從當前活動 SS）
    if (activeSSId) {
      try {
        const ss = SpreadsheetApp.openById(activeSSId);
        const dataSheet = ss.getSheetByName(Service_DB.getActiveDataSheetName());
        if (dataSheet) {
           mockStat.totalStudents = Math.max(0, dataSheet.getLastRow() - 1);
           mockStat.dataSheetId = dataSheet.getSheetId();
        }
      } catch(ignore) {}
    }

    return { 
      status: 'success', 
      data: mockStat 
    };
    
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 新增題目 (Admin Only)
 */
function apiAddQuestion(password, questionData) {
  try {
    if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
    
    const ssid = Service_DB.getMasterId();
    if (!ssid) return { status: 'error', message: '找不到主控台 ID' };
    
    return Service_DB.addQuestionConfig(ssid, questionData);
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 提交答案
 */
function apiSubmit(studentId, stage, question, answer) {
  try {
    // v10.4.0: 答案寫入當前活動的試算表
    const ssid = Service_DB.getActiveSSId() || Service_DB.getMasterId();
    
    // 建構 answers 物件
    const ansObj = {};
    ansObj[question] = answer;
    
    // 呼叫引擎
    const result = Service_Engine.processSubmission(ssid, studentId, stage, ansObj);
    return result;
    
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 取得所有階段清單 (Admin Only)
 * 從試算表動態讀取，確保 UI 與設定同步
 */
function apiGetStages(password) {
  try {
    if (!Service_Security.verifyAdmin(password)) {
      return { status: 'error', message: '權限不足' };
    }
    const masterId = Service_DB.getMasterId();
    if (!masterId) return { status: 'error', message: '系統尚未初始化' };
    // v10.4.0: 從當前活動試算表讀取階段清單
    const ssid = Service_DB.getActiveSSId() || masterId;

    const allQs = Service_DB.getActivityConfig(ssid);
    // 依照出現順序去重，保留唯一階段標籤
    const seen = new Set();
    const stages = [];
    allQs.forEach(q => {
      if (q.label && !seen.has(q.label)) {
        seen.add(q.label);
        stages.push(q.label);
      }
    });

    return { status: 'success', stages };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// =============================================
// 多活動管理 API (v10.3.7)
// =============================================

/**
 * API: 取得所有活動清單 (Admin Only)
 */
function apiGetActivityList(password) {
  try {
    if (!Service_Security.verifyAdmin(password)) {
      return { status: 'error', message: '權限不足' };
    }
    const ssid = Service_DB.getMasterId();
    if (!ssid) return { status: 'error', message: '系統尚未初始化' };

    const activities = Service_DB.getActivityList(ssid);
    const activeId   = Service_DB.getActiveActivityId();
    return {
      status: 'success',
      activities,
      activeId: activeId || 'default'
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 建立新活動 (Admin Only)
 * @param {string} password
 * @param {string} activityName - 活動名稱
 */
function apiCreateActivity(password, activityName) {
  try {
    if (!Service_Security.verifyAdmin(password)) {
      return { status: 'error', message: '權限不足' };
    }
    const ssid = Service_DB.getMasterId();
    if (!ssid) return { status: 'error', message: '系統尚未初始化' };

    return Service_DB.createActivity(ssid, activityName);
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 切換當前活動 (Admin Only)
 * @param {string} password
 * @param {string} activityId - 活動 ID（傳入 'default' 可回到預設活動）
 */
function apiSwitchActivity(password, activityId) {
  try {
    if (!Service_Security.verifyAdmin(password)) {
      return { status: 'error', message: '權限不足' };
    }
    const target = (activityId === 'default') ? '' : activityId;
    Service_DB.setActiveActivityId(target);
    return {
      status: 'success',
      message: target ? '已切換至活動「' + activityId + '」' : '已切換回預設活動'
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * 保留 doPost 以備未來擴充 (如自訂表單提交)
 */
function doPost(e) {
  return ContentService.createTextOutput("POST request received");
}

// --- Admin CMS 2.0 擴充 API ---

/**
 * 取得系統環境診斷報告
 */
function apiGetSystemStatus(password) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
  
  const masterId = Service_DB.getMasterId();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  return {
    status: 'success',
    data: {
      spreadsheetId: masterId,
      hasApiKey: !!apiKey,
      version: 'v10.3.7',
      timezone: Session.getScriptTimeZone(),
      user: Session.getActiveUser().getEmail()
    }
  };
}

/**
 * 清除系統全域快取
 */
function apiClearCache(password) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
  const masterId = Service_DB.getMasterId();
  Service_DB.clearActivityConfigCache(masterId);
  return { status: 'success', message: '已成功清除活動設定快取' };
}

/**
 * 取得說明文件內容
 */
function apiGetManual(password, manualId) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };

  const sections = [
    {
      icon: '🎯',
      title: '多活動管理（v10.3.7 新功能）',
      color: '#1a2a6c',
      items: [
        { label: '什麼是「活動」', text: '一個活動就是一門課程。系統可同時管理多個課程，各課程有獨立的題目與學員資料。' },
        { label: '新增活動', text: '「系統工具」分頁 → 「🎯 活動管理」→ 輸入活動名稱 → 點「➕ 新增活動」。系統會自動建立三個試算表分頁。' },
        { label: '切換活動', text: '在活動清單對應項目點「切換為當前」，前台學員下次登入即讀取新課程題目。' },
        { label: '回到預設活動', text: '點「預設活動」的切換，就會回到「系統設定」與「學員資料總表」的原始狀態。' }
      ]
    },
    {
      icon: '📂',
      title: '試算表分頁結構說明',
      color: '#217346',
      items: [
        { label: '系統設定', text: '預設活動的題目設定表。每列一題，A欄=階段標籤、C欄=題目內容、D欄=答案欄位名稱（必須唯一）、E欄=題型。' },
        { label: '學員資料總表', text: '預設活動的學員作答記錄。學員登入後系統自動建立，不需手動新增。' },
        { label: '[活動名]-設定', text: '新活動的題目表。建立完活動後，在此表填入題目即可。格式與「系統設定」相同。' },
        { label: '[活動名]-學員', text: '新活動的學員作答區域（自動建立，不需手動操作）。' },
        { label: '活動清單', text: '所有活動的索引表。請勿手動編輯此表，由系統自動維護。' }
      ]
    },
    {
      icon: '🛠️',
      title: '題目設定 SOP',
      color: '#b21f1f',
      items: [
        { label: '答案欄位名稱 (D欄)', text: '必須全專案唯一。設定後不建議修改，否則舊有資料將無法對應欄位。' },
        { label: '選項 (F欄)', text: '單選/多選題請用「英文半形逗號 ,」分隔。簡答/段落題可留空。' },
        { label: '標準答案 (G欄)', text: '單選填選項字母（如 A），多選用逗號分隔（如 A,C），簡答填關鍵詞。' },
        { label: '新增題目後', text: '請到「系統工具」→「🧹 清除全域快取」，或等待 60 秒快取自動失效，再請學員重新登入。' }
      ]
    },
    {
      icon: '⚠️',
      title: '常見問題與注意事項',
      color: '#e67e22',
      items: [
        { label: '學員登入後看到舊題', text: '請到「系統工具」清除快取，學員再次登入即可讀取最新題目。' },
        { label: '切換活動後學員看到舊課程', text: '活動切換屬即時生效，學員下次登入就會讀取新活動題目。如果學員已登入中，需請對方重新整理。' },
        { label: '各活動學員資料會分隔', text: '學員在活動 A 的作答不會出現在活動 B 的學員表，規劃上完全隔離。' },
        { label: '不要手動刪除分頁', text: '請透過系統工具管理活動，手動刪除 [x]-設定 或 [x]-學員 將導致系統讀取失敗。' }
      ]
    }
  ];

  return { status: 'success', sections: sections };
}

/**
 * API: 自動偵測應用程式網址 (極簡穩定版 v10.3.7)
 * 解決原因：ScriptApp.getService().getUrl() 可能在未完全授權環境下導致執行掛起
 */
function apiGetAppUrls() {
  var result = { portal: "", admin: "", status: "checking" };
  try {
    // 優先使用當前活動試算表，這絕對不會掛起
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("系統設定");
    if (sheet) {
      const m2 = sheet.getRange("M2").getValue() || "";
      if (m2) {
         result.portal = String(m2).replace('/dev', '/exec');
         result.admin = sheet.getRange("M3").getValue() ? String(sheet.getRange("M3").getValue()).replace('/dev', '/exec') : (result.portal + "?page=admin");
      }
    }
  } catch(e) {
    result.error = "ReadSheetError: " + e.toString();
  }

  // 二次檢驗：若試算表內沒存，最後才嘗試 getUrl，且放在最後一刻
  if (!result.portal || String(result.portal).indexOf('http') === -1) {
    try {
      var serviceUrl = ScriptApp.getService().getUrl();
      if (serviceUrl) { 
        result.portal = serviceUrl.replace('/dev', '/exec'); // 自動修復
        result.admin = result.portal + (result.portal.indexOf('?') > -1 ? '&page=admin' : '?page=admin');
      }
    } catch(e) {
       result.apiError = e.toString();
    }
  }

  // 整理最終文字
  if (!result.portal || String(result.portal).indexOf('http') === -1) {
    result.portal = "尚未設定正式網址 (請參照使用手冊部署新版本)";
    result.admin = "尚未設定正式網址";
    result.status = "offline";
  } else {
    result.status = "online";
  }

  return result;
}
