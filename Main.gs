/**
 * Main.gs - 伊美系統 v10.0 入口與控制器
 * 
 * 職責：
 * 1. 處理 Web App 請求 (doGet)
 * 2. 路由分發
 * 3. 初始化服務
 */

/**
 * 處理 Web App HTTP GET 請求
 * @param {Object} e - 事件物件
 * @returns {HtmlOutput}
 */
function doGet(e) {
  // 防止在編輯器直接執行時報錯
  const params = e && e.parameter ? e.parameter : {};
  const route = params.route || 'portal';
  
  if (route === 'portal') {
    return _renderPortal(e);
  } else if (route === 'install') {
    return _handleInstall(e);
  } else if (route === 'admin') {
    return _renderAdmin(e);
  } else {
    return HtmlService.createHtmlOutput("未知路由");
  }
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
  // TODO: 檢查活動狀態與學員身分 (呼叫 Service_Security)
  // 如果需要登入，回傳登入頁
  // 如果已登入，呼叫 Service_Engine 取得當前狀態
  
  const template = HtmlService.createTemplateFromFile('UI_Portal');
  template.title = '伊美：簡報同步互動學習系統 v10.2.0 (Admin CMS 2.0)';
  
  // 固定顯示正確版號，防止試算表名稱過舊導致誤導
  template.activityName = 'emedu-Slides-Sync-Interactive-System - v10.2.0';
  
  return template.evaluate()
      .setTitle('伊美：簡報同步互動學習系統 v10.2.0')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 處理安裝請求 (僅供管理員)
 */
function _handleInstall(e) {
  try {
    const url = Service_DB.installControlSheet();
    return HtmlService.createHtmlOutput(`核心主控台安裝完成。<br>網址：<a href="${url}" target="_blank">${url}</a>`);
  } catch (err) {
    return HtmlService.createHtmlOutput("安裝失敗：" + err.toString());
  }
}

/**
 * 🟢 [手動執行] 系統初始化
 * 請在上方下拉選單選擇此函式，並點擊「執行」以建立試算表。
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
    const ssid = Service_DB.getMasterId();
    console.log(`[API] Master ID: ${ssid}`);
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
    
    // 若無任何題目 (空設定)
    if (!nextTaskResult.question) {
      throw new Error("尚未設定題目");
    }
    
    const q = nextTaskResult.question;
    const result = {
      status: 'success',
      datestamp: new Date().toISOString(), // 轉為 ISO 字串，避免序列化問題
      task: {
        stage: q.label,
        question: q.question,
        type: q.type,
        desc: q.helpText || "(無提示)"
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
    const ssid = Service_DB.getMasterId();
    const mockStat = {
      totalStudents: 0,
      totalSubmissions: 0,
      masterUrl: ssid ? `https://docs.google.com/spreadsheets/d/${ssid}` : "#"
    };
    
    // 如果有試算表，嘗試取得真實數據
    if (ssid) {
      const ss = SpreadsheetApp.openById(ssid);
      try {
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
    const ssid = Service_DB.getMasterId();
    
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
    const ssid = Service_DB.getMasterId();
    if (!ssid) return { status: 'error', message: '系統尚未初始化' };

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
      version: 'v10.2.0',
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
  
  // 這裡可根據 manualId 回傳不同的文件內容
  // 為求精確，我們直接回傳「試算表全功能維護手冊」的內容
  // (在實際 GAS 環境，如果 .md 檔沒一起 push 上去，這裡改用硬編碼或字串)
  const content = `
### 📂 系統三大核心分頁概覽
1. **[系統設定]**：系統的「大腦」，決定題目、分數與題型。
2. **[學員資料總表]**：系統的「保險箱」，存放原始答案。
3. **[進度追蹤表]**：系統的「儀表板」，監看完成進度。

### 🛠️ 核心設定 SOP
- **選項 (F欄)**：請用 **英文半形逗號 ,** 分隔。
- **答案欄位名稱 (D欄)**：必須唯一，不可重複且設定後不建議修改。
- **自動計分**：系統會將學員答案與 **標準答案 (G欄)** 進行比對。
  `;
  
  return { status: 'success', content: content };
}

// ============================================================
// 多活動管理 API
// ============================================================

/**
 * API: 取得活動清單 (Admin Only)
 */
function apiGetActivityList(password) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
  const ssid = Service_DB.getMasterId();
  if (!ssid) return { status: 'error', message: '系統尚未初始化' };
  try {
    const list = Service_DB.getActivityList(ssid);
    const currentId = Service_DB.getActiveActivityId() || 'default';
    return { status: 'success', activities: list, currentId: currentId };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * API: 建立新活動 (Admin Only)
 * @param {string} password - 管理密碼
 * @param {string} activityName - 活動名稱（如 "2025秋季班"）
 */
function apiCreateActivity(password, activityName) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
  const ssid = Service_DB.getMasterId();
  if (!ssid) return { status: 'error', message: '系統尚未初始化' };
  return Service_DB.createActivity(ssid, activityName);
}

/**
 * API: 切換當前活動 (Admin Only)
 * @param {string} password - 管理密碼
 * @param {string} activityId - 目標活動 ID（傳入 "default" 或空字串可回退至預設）
 */
function apiSwitchActivity(password, activityId) {
  if (!Service_Security.verifyAdmin(password)) return { status: 'error', message: '權限不足' };
  try {
    Service_DB.setActiveActivityId(activityId === 'default' ? '' : activityId);
    return { status: 'success', message: '已切換至活動：' + (activityId === 'default' ? '預設活動' : activityId) };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

