/**
 * Service_DB.gs - 資料庫抽象層
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
    MASTER_SPREADSHEET_NAME: "伊美：簡報同步互動學習系統 - 主控台 v10.0",
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
     * 安裝主控台 (遷移自 v9.4.2 install_ControlSheet)
     */
    installControlSheet: function() {
      const ss = SpreadsheetApp.create(CONFIG.MASTER_SPREADSHEET_NAME);
      const masterId = ss.getId();
      console.log("✅ 主控台試算表已建立：" + ss.getUrl());

      // 1. 系統設定表
      const settings = ss.insertSheet(CONFIG.MASTER_SETTINGS_SHEET);
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
      
      // 2. 活動清單
      const list = ss.insertSheet(CONFIG.ACTIVITIES_OVERVIEW_SHEET);
      list.getRange(1,1,1,6).setValues([["活動ID","活動名稱","狀態","活動資料簿連結","表單數","建立時間"]]).setFontWeight("bold");
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
      
      return ss.getUrl();
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
     */
    addQuestionConfig: function(ssId, questionData) {
       const ss = SpreadsheetApp.openById(ssId);
       let sheet = ss.getSheetByName(CONFIG.MASTER_SETTINGS_SHEET);
       if (!sheet) return { status: 'error', message: '找不到設定表' };
       
       // Headers參考: Label, Form Title, Question, Ans Col, Type, Options, Std Ans, Score, Help, ...
       // 取得最後一列
       const lastRow = sheet.getLastRow();
       
       const newRow = [
         questionData.label || "未分類", // 階段 default
         `學習系統：${questionData.label || "未命名"}`, // 預設標題
         questionData.question || "未命名題目",       // 問題
         questionData.key || `Q${Date.now()}`, // 答案欄位 (自動產生)
         questionData.type || "單選題",           // 類型
         questionData.options || "",  // 選項
         questionData.answer || "",   // 標準答案
         Number(questionData.score) || 0,     // 分數 (Ensure number)
         questionData.desc || "",     // 提示
         "", // 截止日
         ""  // 必修數
       ];
       
       sheet.appendRow(newRow);
       return { status: 'success' };
    },

    /**
     * 取得目前 Master ID
     */
    getMasterId: function() {
      return _getScriptProp('MASTER_ID');
    },
    
    /**
     * 寫入單一儲存格 (核心寫入邏輯)
     * 遷移自 TriggerHandler._updateSingleDataCell
     */
    /**
     * 讀取活動設定表 (快照或設定)
     */
    getActivityConfig: function(ssId) {
       const ss = SpreadsheetApp.openById(ssId);
       // 優先找快照，找不到找一般設定
       let sheet = ss.getSheetByName(CONFIG.ACT_SETTINGS_SNAPSHOT);
       if (!sheet) sheet = ss.getSheetByName(CONFIG.MASTER_SETTINGS_SHEET);
       return this.readConfig(sheet);
    },

    /**
     * 更新資料工作表
     */
    updateStudentData: function(ssid, sheetName, studentId, email, keyColumn, value, scoreColumn, score, timestampColumn) {
       // ... existing code ...
       const ss = SpreadsheetApp.openById(ssid);
       const sheet = ss.getSheetByName(sheetName);
       if (!sheet) return;
       
       const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       const headerMap = {};
       headers.forEach((h, i) => headerMap[h] = i + 1);
       
       const sidCol = headerMap["學號"];
       const emCol = headerMap["Email"];
       if (!sidCol) return;
       
       // 尋找或建立列
       let row = -1;
       const data = sheet.getRange(2, sidCol, Math.max(sheet.getLastRow()-1, 1), 1).getValues();
       for(let i=0; i<data.length; i++) {
         if(String(data[i][0]).toLowerCase() === String(studentId).toLowerCase()) {
           row = i + 2;
           break;
         }
       }
       
       if (row === -1) {
         const newRow = new Array(sheet.getLastColumn()).fill("");
         newRow[sidCol - 1] = studentId;
         if (email && emCol) newRow[emCol - 1] = email;
         sheet.appendRow(newRow);
         row = sheet.getLastRow();
       } else {
          // 更新 Email?
          if (email && emCol) sheet.getRange(row, emCol).setValue(email);
       }
       
       // 寫入資料
       const now = new Date();
       if (headerMap[keyColumn]) sheet.getRange(row, headerMap[keyColumn]).setValue(value);
       if (scoreColumn && headerMap[scoreColumn]) sheet.getRange(row, headerMap[scoreColumn]).setValue(score);
       if (timestampColumn && headerMap[timestampColumn]) sheet.getRange(row, headerMap[timestampColumn]).setValue(now);
    },
    
    /**
     * 取得學員完整資料列 (回傳 Map: Header -> Value)
     */
    getStudentRowData: function(ssid, studentId) {
       const ss = SpreadsheetApp.openById(ssid);
       const sheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
       if (!sheet || sheet.getLastRow() < 2) return {};
       
       const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       const sidColIndex = headers.indexOf("學號");
       if (sidColIndex === -1) return {};
       
       const data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
       for (let i = 0; i < data.length; i++) {
         if (String(data[i][sidColIndex]).toLowerCase() === String(studentId).toLowerCase()) {
            const rowMap = {};
            headers.forEach((h, idx) => {
              rowMap[h] = data[i][idx];
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
    
    CONFIG // 導出設定供其他模組參考
  };

})();
