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
  template.title = '伊美：簡報同步互動學習系統 v10.1';
  
  // 嘗試取得活動名稱
  let actName = "互動學習門戶";
  try {
    const ssid = Service_DB.getMasterId();
    if (ssid) {
      actName = SpreadsheetApp.openById(ssid).getName();
    }
  } catch(e) { console.warn("無法取得活動名稱: " + e.message); }
  
  template.activityName = actName;
  
  return template.evaluate()
      .setTitle('伊美：簡報同步互動學習系統 v10.1')
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
    
    return {
      status: 'success',
      datestamp: new Date(),
      task: {
        stage: q.label,
        question: q.question,
        type: q.type,
        desc: q.helpText || "(無提示)"
      }
    };
  } catch (e) {
    return { status: 'error', message: e.toString() }; 
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
        const dataSheet = ss.getSheetByName(Service_DB.CONFIG.DATA_SHEET_NAME);
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
