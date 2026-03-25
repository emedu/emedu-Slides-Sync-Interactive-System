/**
 * Service_DB.gs - 資料庫抽象層 v10.3.7 (多活動版)
 * 
 * 職責：
 * 1. 封裝 SpreadsheetApp 操作
 * 2. 管理主控台與活動資料簿
 * 3. 提供設定讀取與資料寫入介面
 * 
 * 遷移自 v9.4.2: Utils, install_ControlSheet
 */

const Service_DB = (function() {

  // --- 設定常數 (保留 v9.4.2 設定) ---
  const CONFIG = {
    MASTER_SPREADSHEET_NAME: "emedu-Slides-Sync-Interactive-System - 主控台 v10.3.7",
    MASTER_SETTINGS_SHEET: "系統設定",
    ACT_SETTINGS_SNAPSHOT: "系統設定(快照)",
    DATA_SHEET_NAME: "學員資料總表",
    TRACKING_SHEET_NAME: "進度追蹤表",
    ACTIVITIES_OVERVIEW_SHEET: "活動清單"
  };

  const ALLOWED_TYPES = new Set(["單選題", "多選題", "簡答題", "段落"]);

  // --- 私有工具函式 (Internal Utils) ---
  
  function _getScriptProp(key, def = null) {
     const val = PropertiesService.getScriptProperties().getProperty(key);
     return val === null ? def : val;
  }

  function _setScriptProp(key, val) {
    PropertiesService.getScriptProperties().setProperty(key, String(val));
  }

  // --- 公開介面 ---
  return {
    
    /**
     * 安裝主控台 (v10.3.7 優化：支援綁定型腳本)
     */
    installControlSheet: function() {
      let ss;
      const existingId = _getScriptProp('MASTER_ID');
      
      // 優先偵測是否為試算表綁定型腳本環境
      try {
        const activeSS = SpreadsheetApp.getActiveSpreadsheet();
        if (activeSS) {
          ss = activeSS;
          console.log("🔗 偵測到綁定型腳本，使用現有試算表: " + ss.getUrl());
        }
      } catch(e) {
        // 非綁定型腳本環境，繼續以下邏輯
      }

      if (!ss && existingId) {
        try {
          ss = SpreadsheetApp.openById(existingId);
          console.log("♻️ 偵測到現有主控台，執行熱更新 (Hot Update)...");
          ss.rename(CONFIG.MASTER_SPREADSHEET_NAME);
        } catch(e) {
          console.log("ℹ️ 舊有 MASTER_ID 無法存取，將建立新主控台。");
        }
      }

      if (!ss) {
        ss = SpreadsheetApp.create(CONFIG.MASTER_SPREADSHEET_NAME);
        console.log("✅ 主控台試算表已建立：" + ss.getUrl());
      }

      const masterId = ss.getId();
      console.log("✅ 主控台試算表已建立：" + ss.getUrl());

      // 1. 系統設定表 (檢查後更新)
      let settings = ss.getSheetByName(CONFIG.MASTER_SETTINGS_SHEET);
      if (!settings) settings = ss.insertSheet(CONFIG.MASTER_SETTINGS_SHEET);
      const headers = [
        "階段標籤 (Label)", "表單標題 (Form Title)", "表單問題 (Question)", 
        "答案欄位名稱 (Answer Column)", "題目類型 (單選題/多選題/簡答題/段落)", 
        "選項 (若為選擇題, 用 , 分隔)", "標準答案 (簡答/選擇；多選用 , 分隔)", 
        "分數 (整題分值)", "提示文字 (選填)", "截止日 (YYYY-MM-DD/選填)", "必修題數 (選填)"
      ];
      settings.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      
      // 範例資料
      const examples = [
        ["階段一","學習系統：階段一","問題 1-1（單選）","答案 1-1","單選題","A,B,C","A",10,"提示 1-1","",""],
        ["階段一","學習系統：階段一","問題 1-2（簡答）","答案 1-2","簡答題","","正確詞",10,"提示 1-2","",""],
        ["階段二","學習系統：階段二","問題 2-1（多選）","答案 2-1","多選題","A,B,C,D","A,C",20,"提示 2-1","",1]
      ];
      settings.getRange(2, 1, examples.length, headers.length).setValues(examples);
      settings.setFrozenRows(1);

      // --- [NEW] 智慧連結面板 (建立於 L, M 欄位以避開資料區) ---
      const dashTitle = [["🚀 系統管理快捷面板 (自動偵測)"]];
      settings.getRange("L1").setValues(dashTitle).setFontWeight("bold").setFontColor("#ffffff").setBackground("#1a73e8");
      
      let appUrl = "";
      try {
        appUrl = ScriptApp.getService().getUrl();
      } catch(e) {
        console.warn("無法取得 Web App 網址 (可能是尚未部署或權限限制)");
      }
      appUrl = appUrl || "尚未部署 (請完成部署後於試算表選單更新)";

      const links = [
        ["📱 使用者入口 (前台)", appUrl],
        ["⚙️ 管理中心 (後台)", appUrl + (appUrl.indexOf('?') > -1 ? '&page=admin' : '?page=admin')]
      ];
      settings.getRange("L2:M3").setValues(links);
      settings.getRange("M2:M3").setFontColor("#1a73e8").setFontLine("underline");
      
      // 自動調整欄寬
      settings.autoResizeColumn(12);
      settings.autoResizeColumn(13);
      
      // 2. 活動清單
      let list = ss.getSheetByName(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      if (!list) list = ss.insertSheet(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      list.getRange(1,1,1,6).setValues([["活動ID","活動名稱","狀態","建立時間","設定分頁","學員分頁"]]).setFontWeight("bold");
      list.setFrozenRows(1);

      // 3. 學員資料總表 (Student Data Sheet)
      let dataSheet;
      try {
        dataSheet = ss.insertSheet(CONFIG.DATA_SHEET_NAME);
      } catch (e) {
        dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
      }
      const dataHeaders = ["學號", "姓名", "Email", "備註"]; // Basic headers
      dataSheet.getRange(1, 1, 1, dataHeaders.length).setValues([dataHeaders]).setFontWeight("bold");
      dataSheet.setFrozenRows(1);


      // 清除預設工作表
      try { ss.deleteSheet(ss.getSheetByName('工作表1')); } catch(e){}

      // 儲存 Master ID
      _setScriptProp('MASTER_ID', masterId);
      
      // 自動部署安裝式觸發器 (重要：解決選單不顯示問題)
      this.setupInstallableTriggers(masterId);
      
      return ss.getUrl();
    },

    /**
     * 自動部署安裝式觸發器
     * 解決獨立腳本無法自動執行 onOpen 的問題
     */
    setupInstallableTriggers: function(ssId) {
      console.log("🛠️ 正在部屬自動化觸發器...");
      try {
        const triggers = ScriptApp.getProjectTriggers();
        const triggerFunction = 'onOpen'; 
        
        triggers.forEach(t => {
          if (t.getHandlerFunction() === triggerFunction) ScriptApp.deleteTrigger(t);
        });

        ScriptApp.newTrigger(triggerFunction)
          .forSpreadsheet(ssId)
          .onOpen()
          .create();
        console.log("✅ 觸發器部屬成功。");
      } catch(e) {
        console.warn("⚠️ 觸發器部屬失敗: " + e.message);
      }
    },

    /**
     * 讀取活動設定 (遷移自 Utils.readConfig)
     * @param {Sheet} sheet - 設定工作表
     */
    readConfig: function(sheet) {
      if (!sheet || sheet.getLastRow() < 2) return [];
      const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues();
      const out = [];

      data.forEach((r, i) => {
        const rowIndex = i + 2;
        const [label, formTitle, question, targetColumn, type, opts, stdAns, score, help, deadlineStr, reqCount] = r;
        
        if (!label || !targetColumn || !type) return;

        let deadline = null;
        if (deadlineStr) {
           const d = new Date(deadlineStr);
           if (!isNaN(d.getTime())) deadline = d;
        }

        out.push({
          label: String(label).trim(),
          formTitle: String(formTitle).trim(),
          question: String(question).trim(),
          targetColumn: String(targetColumn).trim(),
          type: String(type).trim(),
          options: opts ? String(opts).split(",").map(s=>s.trim()) : [],
          standardAnswer: stdAns ? String(stdAns).trim() : "",
          score: Number(score) || 0,
          helpText: help ? String(help).trim() : "",
          deadline: deadline,
          requiredCount: Number(reqCount) || 0,
          timestampColumn: `${String(targetColumn).trim()} (提交時間)`,
          scoreColumn: `${String(targetColumn).trim()} (得分)`,
          rowIndex: rowIndex
        });
      });
      return out;
    },

    /**
     * 準備資料工作表 (遷移自 Utils.prepareDataSheet)
     */
    prepareDataSheet: function(sheet, config) {
      sheet.clear();
      const headers = ["學號", "Email"];
      const added = new Set();
      
      config.forEach(q => {
        [q.targetColumn, q.scoreColumn, q.timestampColumn].forEach(h => {
          if (!added.has(h)) {
            headers.push(h);
            added.add(h);
          }
        });
      });
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sheet.setFrozenRows(1);
    },

    /**
     * 新增題目設定 (Admin CMS)
     * v10.4.0: ssId 傳入的是「當前活動的獨立試算表 ID」(由 apiAddQuestion 負責路由)
     * 因此直接讀取固定的「系統設定」分頁，不再使用 [活動名]-設定 動態分頁名稱
     */
    addQuestionConfig: function(ssId, questionData) {
       const ss = SpreadsheetApp.openById(ssId);
       
       // v10.4.0: 固定使用「系統設定」分頁（每個活動的獨立試算表均採固定名稱）
       let sheet = ss.getSheetByName(CONFIG.ACT_SETTINGS_SNAPSHOT);
       if (!sheet) sheet = ss.getSheetByName(CONFIG.MASTER_SETTINGS_SHEET);
       if (!sheet) return { status: 'error', message: '找不到設定表（系統設定 / 系統設定(快照)）' };
       
       const newRow = [
         questionData.label || "未分類",
         `學習系統：${questionData.label || "未命名"}`,
         questionData.question || "未命名題目",
         questionData.key || `Q${Date.now()}`,
         questionData.type || "單選題",
         questionData.options || "",
         questionData.answer || "",
         Number(questionData.score) || 0,
         questionData.desc || "",
         "",
         ""
       ];
       
       sheet.appendRow(newRow);
       // 題目新增後清除快取，確保下次 getActivityConfig 讀到最新資料
       this.clearActivityConfigCache(ssId);
       return { status: 'success' };
    },

    /**
     * 取得目前 Master ID
     */
    getMasterId: function() {
      return _getScriptProp('MASTER_ID');
    },

    // =============================================
    // 多活動管理 API (v10.4.0 — 一活動一試算表版)
    // =============================================

    /**
     * [新增] 根據活動 ID 取得其對應的獨立試算表 ID
     * 若為預設活動 (default/null) 則回傳 masterId
     * @param {string} activityId
     * @returns {string} Spreadsheet ID
     */
    getActivitySSId: function(activityId) {
      if (!activityId || activityId === 'default') return _getScriptProp('MASTER_ID');
      const masterId = _getScriptProp('MASTER_ID');
      if (!masterId) return null;

      const ss = SpreadsheetApp.openById(masterId);
      const listSheet = ss.getSheetByName(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      if (!listSheet || listSheet.getLastRow() < 2) return masterId;

      const rows = listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 5).getValues();
      for (const row of rows) {
        if (String(row[0]) === activityId) {
          const ssId = String(row[4]).trim(); // 欄位 E = 試算表ID
          return ssId || masterId;
        }
      }
      return masterId; // fallback
    },

    /**
     * 取得目前進行中的活動 ID
     * 若未設定則回傳 null（代表使用預設的 系統設定 分頁）
     */
    getActiveActivityId: function() {
      return _getScriptProp('ACTIVE_ACTIVITY_ID', null);
    },

    /**
     * 設定目前進行中的活動 ID（切換活動）
     * @param {string|null} activityId - 活動名稱，傳入空字串代表回到預設活動
     */
    setActiveActivityId: function(activityId) {
      if (!activityId || activityId.trim() === '') {
        PropertiesService.getScriptProperties().deleteProperty('ACTIVE_ACTIVITY_ID');
        console.log('[Activity] 已切換回預設活動 (系統設定)');
      } else {
        _setScriptProp('ACTIVE_ACTIVITY_ID', activityId.trim());
        console.log('[Activity] 已切換至活動: ' + activityId);
      }
      // 清除快取，確保下一次讀取最新設定
      const masterId = _getScriptProp('MASTER_ID');
      if (masterId) {
        const currentId = activityId ? activityId.trim() : 'default';
        CacheService.getScriptCache().removeAll([
          'activity_config_' + masterId + '_default',
          'activity_config_' + masterId + '_' + currentId
        ]);
      }
    },

    /**
     * 建立新活動 (v10.4.0 — 每個活動建立獨立的 Google Spreadsheet)
     * @param {string} ssId - 主控台試算表 ID
     * @param {string} activityName - 活動名稱（中英文均可，勿含特殊符號）
     */
    createActivity: function(ssId, activityName) {
      if (!activityName || activityName.trim() === '') {
        return { status: 'error', message: '活動名稱不可為空！' };
      }
      const name = activityName.trim();
      const masterSS = SpreadsheetApp.openById(ssId);

      // 先確認活動清單是否已有同名活動
      let listSheet = masterSS.getSheetByName(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      if (listSheet && listSheet.getLastRow() >= 2) {
        const rows = listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 1).getValues();
        if (rows.some(r => String(r[0]) === name)) {
          return { status: 'error', message: '活動「' + name + '」已存在！' };
        }
      }

      // --- 核心改動：建立獨立試算表 ---
      const newSS = SpreadsheetApp.create('emedu-活動-' + name);
      const newSSId  = newSS.getId();
      const newSSUrl = newSS.getUrl();
      console.log('[Activity v10.4.0] 已建立獨立試算表: ' + newSSUrl);

      // 在新試算表中建立固定名稱的分頁（與預設活動結構一致）
      const headers = [
        '階段標籤 (Label)', '表單標題 (Form Title)', '表單問題 (Question)',
        '答案欄位名稱 (Answer Column)', '題目類型 (單選題/多選題/簡答題/段落)',
        '選項 (若為選擇題, 用 , 分隔)', '標準答案 (簡答/選擇；多選用 , 分隔)',
        '分數 (整題分值)', '提示文字 (選填)', '截止日 (YYYY-MM-DD/選填)', '必修題數 (選填)'
      ];

      // 1. 系統設定分頁
      const settingsSheet = newSS.getActiveSheet();
      settingsSheet.setName(CONFIG.MASTER_SETTINGS_SHEET);
      settingsSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      settingsSheet.setFrozenRows(1);

      // 2. 學員資料總表
      const dataSheet = newSS.insertSheet(CONFIG.DATA_SHEET_NAME);
      dataSheet.getRange(1, 1, 1, 2).setValues([['學號', 'Email']]).setFontWeight('bold');
      dataSheet.setFrozenRows(1);

      // 3. 進度追蹤表
      const trackingSheet = newSS.insertSheet(CONFIG.TRACKING_SHEET_NAME);
      trackingSheet.getRange(1, 1, 1, 3).setValues([['學號', '完成數', '總分']]).setFontWeight('bold');
      trackingSheet.setFrozenRows(1);

      // --- 在主控台「活動清單」中登記 (新欄位結構: 活動ID | 活動名稱 | 狀態 | 建立時間 | 試算表ID | 試算表連結) ---
      if (!listSheet) {
        listSheet = masterSS.insertSheet(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
        listSheet.getRange(1, 1, 1, 6).setValues([[
          '活動ID', '活動名稱', '狀態', '建立時間', '試算表ID', '試算表連結'
        ]]).setFontWeight('bold');
      }
      // 確保標題列是新格式（向下相容舊版活動清單）
      const firstHeader = listSheet.getRange(1, 5).getValue();
      if (firstHeader !== '試算表ID') {
        listSheet.getRange(1, 5, 1, 2).setValues([['試算表ID', '試算表連結']]).setFontWeight('bold');
      }
      // 若尚無預設活動的登記，補上
      if (listSheet.getLastRow() < 2) {
        listSheet.appendRow(['default', '預設活動 (主控台)', '進行中', new Date(), ssId, masterSS.getUrl()]);
      }
      listSheet.appendRow([name, name, '進行中', new Date(), newSSId, newSSUrl]);

      console.log('[Activity v10.4.0] 已在活動清單登記: ' + name);
      return {
        status: 'success',
        message: '活動「' + name + '」已建立！',
        activityId: name,
        spreadsheetUrl: newSSUrl
      };
    },

    /**
     * 取得所有活動清單 (v10.4.0 — 新增 spreadsheetUrl 欄位)
     * @param {string} ssId - 主控台試算表 ID
     * @returns {Array} 活動資料陣列
     */
    getActivityList: function(ssId) {
      const ss = SpreadsheetApp.openById(ssId);
      const listSheet = ss.getSheetByName(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      const activeId = this.getActiveActivityId();

      // 若活動清單不存在，回傳預設活動
      if (!listSheet || listSheet.getLastRow() < 2) {
        return [{
          id: 'default', name: '預設活動 (主控台)',
          status: '進行中', isActive: !activeId,
          spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + ssId
        }];
      }

      const rows = listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 6).getValues();
      return rows.map(r => ({
        id:             String(r[0]),
        name:           String(r[1]),
        status:         String(r[2]),
        created:        r[3] ? new Date(r[3]).toLocaleDateString() : '',
        isActive:       (String(r[0]) === activeId) || (!activeId && String(r[0]) === 'default'),
        spreadsheetUrl: String(r[5]) || '' // 新欄位：試算表連結
      }));
    },


    /**
     * 取得目前活動的資料工作表名稱 (v10.4.0 — 所有試算表統一使用固定名稱)
     */
    getActiveDataSheetName: function() {
      return CONFIG.DATA_SHEET_NAME; // 固定: 學員資料總表
    },

    /**
     * 取得目前活動的進度追蹤表名稱 (v10.4.0 — 固定名稱)
     */
    getActiveTrackingSheetName: function() {
      return CONFIG.TRACKING_SHEET_NAME; // 固定: 進度追蹤表
    },

    /**
     * [v10.4.0] 取得目前活動對應的試算表 ID
     * 這是所有讀寫操作的核心入口 — 呼叫此函式取代直接使用 MASTER_ID
     */
    getActiveSSId: function() {
      const activeId = this.getActiveActivityId();
      return this.getActivitySSId(activeId);
    },
    
    /**
     * 寫入單一儲存格 (核心寫入邏輯)
     * 遷移自 TriggerHandler._updateSingleDataCell
     */
    /**
     * 讀取活動設定表 (快照或設定)
     * 效能優化：使用 CacheService 快取 60 秒，避免每次 API 呼叫重複讀取 Spreadsheet
     */
    getActivityConfig: function(ssId) {
       // v10.4.0: ssId 參數保留供相容性，但實際讀取的 SS 由 getActiveSSId() 決定
       const activeId = this.getActiveActivityId();
       const activeSSId = this.getActiveSSId(); // 取得當前活動對應的試算表 ID
       const cacheKey = 'activity_config_' + (activeSSId || ssId) + '_' + (activeId || 'default');
       const cache = CacheService.getScriptCache();
       const cached = cache.get(cacheKey);
       if (cached) {
         try { return JSON.parse(cached); } catch(e) { /* 快取損毀，重新讀取 */ }
       }

       // 開啟當前活動對應的試算表
       const targetSSId = activeSSId || ssId;
       const ss = SpreadsheetApp.openById(targetSSId);

       // v10.4.0: 所有試算表統一使用固定分頁名稱
       let sheet = ss.getSheetByName(CONFIG.ACT_SETTINGS_SNAPSHOT);
       if (!sheet) sheet = ss.getSheetByName(CONFIG.MASTER_SETTINGS_SHEET);

       if (!sheet) {
         console.warn('[Activity v10.4.0] 找不到設定分頁，targetSSId=' + targetSSId);
         return [];
       }

       const config = this.readConfig(sheet);

       // 寫入快取 (CacheService 最大 100KB，序列化前先確認)
       try {
         const serialized = JSON.stringify(config);
         if (serialized.length < 90000) {
           cache.put(cacheKey, serialized, 60); // 60 秒快取
         }
       } catch(e) { /* 序列化失敗不影響主流程 */ }

       return config;
    },

    /**
     * 清除活動設定快取 (題目異動後應呼叫)
     */
    clearActivityConfigCache: function(ssId) {
       // v10.4.0: 清除主控台與當前活動試算表的快取
       const activeId = this.getActiveActivityId();
       const activeSSId = this.getActivitySSId(activeId);
       const keysToRemove = [
         'activity_config_' + ssId + '_default',
         'activity_config_' + (activeSSId || ssId) + '_' + (activeId || 'default')
       ];
       CacheService.getScriptCache().removeAll(keysToRemove);
    },

    /**
     * 更新資料工作表
     */
    /**
     * 更新學員資料 (單筆或批次)
     */
    updateStudentData: function(ssid, sheetName, studentId, email, updates) {
       const ss = SpreadsheetApp.openById(ssid);
       const sheet = ss.getSheetByName(sheetName);
       if (!sheet) return;
       
       const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       const headerMap = {};
       headers.forEach((h, i) => headerMap[h] = i + 1);
       
       const sidCol = headerMap["學號"];
       const emCol = headerMap["Email"];
       if (!sidCol) return;
       
       // 優化搜尋：一次讀取所有學號
       let row = -1;
       const lastRow = sheet.getLastRow();
       if (lastRow > 1) {
         const sidValues = sheet.getRange(2, sidCol, lastRow - 1, 1).getValues();
         const searchId = String(studentId).toLowerCase();
         for(let i=0; i<sidValues.length; i++) {
           if(String(sidValues[i][0]).toLowerCase() === searchId) {
             row = i + 2;
             break;
           }
         }
       }
       
       if (row === -1) {
         const newRow = new Array(sheet.getLastColumn()).fill("");
         newRow[sidCol - 1] = studentId;
         if (email && emCol) newRow[emCol - 1] = email;
         sheet.appendRow(newRow);
         row = sheet.getLastRow();
       } else if (email && emCol) {
         sheet.getRange(row, emCol).setValue(email);
       }
       
       // 批次寫入：一次讀取整列，批次更新，最後一次 setValues 寫回，避免 N 次 API 呼叫
       const totalCols = sheet.getLastColumn();
       const rowRange = sheet.getRange(row, 1, 1, totalCols);
       const rowValues = rowRange.getValues()[0]; // 取得現有列資料

       let hasUpdate = false;
       for (const key in updates) {
         const col = headerMap[key];
         if (col) {
           rowValues[col - 1] = updates[key]; // 更新陣列 (0-indexed)
           hasUpdate = true;
         }
       }

       if (hasUpdate) {
         rowRange.setValues([rowValues]); // 一次寫回整列
       }
    },
    
    /**
     * 取得學員完整資料列 (回傳 Map: Header -> Value)
     */
     getStudentRowData: function(ssid, studentId) {
        const ss = SpreadsheetApp.openById(ssid);
        const sheet = ss.getSheetByName(this.getActiveDataSheetName());
        if (!sheet) return {};
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return {};
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const sidColIndex = headers.indexOf("學號");
        if (sidColIndex === -1) return {};
        
        const fullData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        const searchId = String(studentId).toLowerCase().trim();
        for (let i = 0; i < fullData.length; i++) {
          if (String(fullData[i][sidColIndex]).toLowerCase().trim() === searchId) {
             const rowMap = {};
             headers.forEach((h, idx) => {
               rowMap[h] = fullData[i][idx];
             });
             return rowMap;
          }
        }
        return {};
     },
    
    /**
     * 更新追蹤表
     */
    updateTrackingData: function(ssid, sheetName, studentId, stageMarks, completedCount, totalScore) {
       const ss = SpreadsheetApp.openById(ssid);
       let sheet = ss.getSheetByName(sheetName);
       if (!sheet) return;
       
       // 確保標題列 (學號, 階段1✔, ..., completed, total)
       // 這裡簡化：假設標題已由 install/snapshot 建立好，直接找學號寫入
       // 若欄位不夠彈性，可在此加強動態擴充邏輯
       
       const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       const sidCol = 1; // 假設學號在第一欄
       
       let row = -1;
       const data = sheet.getRange(2, sidCol, Math.max(sheet.getLastRow()-1, 1), 1).getValues();
        for(let i=0; i<data.length; i++) {
         if(String(data[i][0]).toLowerCase() === String(studentId).toLowerCase()) {
           row = i + 2;
           break;
         }
       }
       
       // 準備整列資料 (部分更新較難，這裡假設覆寫該列 tracking info)
       // 實際上應根據 header 填入
       if (row === -1) {
          // Append
          const newRow = [studentId, ...stageMarks, completedCount, totalScore];
          sheet.appendRow(newRow);
       } else {
          // Update
          // 假設 tracking sheet 結構: [學號, Stage1, Stage2, ..., Completed, Total]
          // 需對齊
          const range = sheet.getRange(row, 2, 1, stageMarks.length + 2);
          range.setValues([[...stageMarks, completedCount, totalScore]]);
       }
    },

    
    // --- 輔助 ---
    getHeaderIndexMap: function(sheet) {
        const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
        const map={}; 
        headers.forEach((h,i)=>map[h]=i+1); 
        return map;
    },
    
    /**
     * 生成學員學習歷程 (Portfolio) - 用於備份與存檔
     */
    generateLearningPortfolio: function(ssid, studentId) {
       const studentData = this.getStudentRowData(ssid, studentId);
       const config = this.getActivityConfig(ssid);
       
       let md = `# emedu 學習歷程備份 - 學號: ${studentId}\n\n`;
       md += `> 系統名稱: ${CONFIG.MASTER_SPREADSHEET_NAME}\n`;
       md += `> 備份時間: ${new Date().toLocaleString()}\n\n`;
       md += `## 📚 課程互動紀錄\n\n`;
       
       config.forEach(q => {
         const ans = studentData[q.targetColumn];
         const score = studentData[q.scoreColumn];
         const time = studentData[q.timestampColumn];
         
         if (ans) {
           md += `### 🏷️ [${q.label}] ${q.question}\n`;
           md += `- **作答內容**: ${ans}\n`;
           md += `- **得分**: ${score} / ${q.score}\n`;
           md += `- **提交時間**: ${time ? new Date(time).toLocaleString() : "N/A"}\n`;
           md += `\n`;
         }
       });
       
       return md;
    },

    CONFIG // 導出設定供其他模組參考
  };

})();
