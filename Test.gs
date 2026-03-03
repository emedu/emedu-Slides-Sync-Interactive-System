/**
 * Test.gs - 自動化測試腳本
 * 
 * 用途：
 * 此腳本包含針對 v10.0 核心模組的單元測試。
 * 請在 Google Apps Script 編輯器中執行 `runAllTests()` 進行驗證。
 */

function runAllTests() {
  console.log("🚀 開始執行伊美系統 v10.0 測試...");
  
  let passed = 0;
  let failed = 0;
  
  const tests = [
    test_Scoring_SingleChoice,
    test_Scoring_MultiChoice,
    test_Scoring_ShortAnswer,
    test_Scoring_ShortAnswer,
    test_Security_RateLimit,
    test_Integration_MockSubmission,
    test_Admin_AddQuestion,
    test_Admin_AddQuestion,
    test_DB_Install,
    test_Engine_NextTask
  ];
  
  tests.forEach(testFn => {
    try {
      console.log(`\n🔹 執行測試: ${testFn.name}`);
      testFn();
      console.log(`✅ ${testFn.name} 通過`);
      passed++;
    } catch (e) {
      console.error(`❌ ${testFn.name} 失敗: ${e.message}`);
      failed++;
    }
  });
  
  console.log(`\n🏁 測試結果總結: 通過 ${passed}, 失敗 ${failed}`);
}

/**
 * 驗證：單選題計分
 */
function test_Scoring_SingleChoice() {
  const q = { type: "單選題", score: 10, standardAnswer: "A" };
  
  assertEq(Service_Engine.calculateScore(q, "A"), 10, "正確答案應得滿分");
  assertEq(Service_Engine.calculateScore(q, "B"), 0, "錯誤答案應得 0 分");
  assertEq(Service_Engine.calculateScore(q, "a"), 10, "大小寫應忽略");
}

/**
 * 驗證：多選題計分 (Partial with Penalty)
 */
function test_Scoring_MultiChoice() {
  const q = { type: "多選題", score: 20, standardAnswer: "A,B" };
  
  // 完全正確
  assertEq(Service_Engine.calculateScore(q, "A,B"), 20, "A,B 應得 20 分");
  assertEq(Service_Engine.calculateScore(q, ["A","B"]), 20, "Array 輸入應得 20 分");
  
  // 部分正確 (1對1錯) -> (1/2) - (1/2) = 0
  assertEq(Service_Engine.calculateScore(q, "A,C"), 0, "A,C 應得 0 分 (Penalty)");
  
  // 部分正確 (1對) -> 1/2 = 10
  assertEq(Service_Engine.calculateScore(q, "A"), 10, "只選 A 應得 10 分");
  
  // 全錯
  assertEq(Service_Engine.calculateScore(q, "C,D"), 0, "C,D 應得 0 分");
}

/**
 * 驗證：簡答題
 */
function test_Scoring_ShortAnswer() {
  const q = { type: "簡答題", score: 10, standardAnswer: "Hello World" };
  
  assertEq(Service_Engine.calculateScore(q, "hello world"), 10, "忽略大小寫");
  assertEq(Service_Engine.calculateScore(q, "  Hello World  "), 10, "忽略前後空白");
  assertEq(Service_Engine.calculateScore(q, "Hello"), 0, "部分符合不給分");
}

/**
 * 驗證：防秒刷 (需 Mock CacheService，若在 GAS 跑則用真實 Cache)
 */
function test_Security_RateLimit() {
  const sid = "TEST_USER_001";
  const stage = "STAGE_TEST";
  
  // 清除舊 Lock (若有)
  try { CacheService.getScriptCache().remove(`v10submit:${sid}:${stage}`); } catch(e){}
  
  // 第一次：應通過
  const allow1 = Service_Security.checkSubmitRateLimit(sid, stage);
  assertEq(allow1, true, "第一次提交應允許");
  
  // 第二次 (立即)：應阻擋
  const allow2 = Service_Security.checkSubmitRateLimit(sid, stage);
  assertEq(allow2, false, "立即重複提交應被阻擋");

  // 清理
  try { CacheService.getScriptCache().remove(`v10submit:${sid}:${stage}`); } catch(e){}
}

/**
 * 驗證：整合流程 (Mock Service_DB)
 */
function test_Integration_MockSubmission() {
  const runId = "TEST_RUN";
  const sid = "S001";
  const stage = "S1";
  
  // Mock Service_DB.getActivityConfig
  const originalGetConfig = Service_DB.getActivityConfig;
  const originalUpdateStudent = Service_DB.updateStudentData;
  const originalUpdateTracking = Service_DB.updateTrackingData;
  const originalGetStudentRow = Service_DB.getStudentRowData;
  
  let updateDataCalled = false;
  let updateTrackingCalled = false;
  
  try {
    // 注入 Mock
    Service_DB.getActivityConfig = function(id) {
      return [
        { label: "S1", question: "Q1", type: "单选題", standardAnswer: "A", score: 10, targetColumn: "A1", scoreColumn: "S1", timestampColumn: "T1" }
      ];
    };
    
    Service_DB.updateStudentData = function(ssid, sheet, sid, email, updates) {
      updateDataCalled = true;
      console.log("Mock updateStudentData called with updates:", JSON.stringify(updates));
      if (updates["A1"] !== "A") throw new Error("Update value mismatch");
    };
    
    Service_DB.updateTrackingData = function() {
      updateTrackingCalled = true;
      console.log("Mock updateTrackingData called");
    };
    
    Service_DB.getStudentRowData = function() {
      return {};
    };
    
    // 執行
    const result = Service_Engine.processSubmission(runId, sid, stage, { "Q1": "A" });
    
    // 驗證
    assertEq(result.status, "success", "提交應成功");
    assertEq(updateDataCalled, true, "應呼叫 updateStudentData");
    assertEq(updateTrackingCalled, true, "應呼叫 updateTrackingData");
    
  } finally {
    // 還原
    Service_DB.getActivityConfig = originalGetConfig;
    Service_DB.updateStudentData = originalUpdateStudent;
    Service_DB.updateTrackingData = originalUpdateTracking;
    Service_DB.getStudentRowData = originalGetStudentRow;
  }
}

/**
 * 驗證：安裝程序 (Mock Install)
 */
function test_DB_Install() {
  // Mock SpreadsheetApp.create & insertSheet
  const createdSheets = [];
  
  // Backup globals
  const originalCreate = SpreadsheetApp.create;
  
  try {
    // Mock
    SpreadsheetApp.create = function(name) {
      return {
        getId: () => "NEW_SS_ID",
        getUrl: () => "mock://new_ss",
        insertSheet: (n) => { 
            createdSheets.push(n); 
            const mockRange = {
                setValues: function() { return this; },
                setFontWeight: function() { return this; }
            };
            return { 
                getRange:() => mockRange, 
                setFrozenRows:()=>{} 
            }; 
        },
        getSheetByName: (n) => null, // 模擬新表無舊sheet
        deleteSheet: () => {}
      };
    };
    
    // Run Install
    Service_DB.installControlSheet();
    
    // Verify
    assertEq(createdSheets.includes(Service_DB.CONFIG.DATA_SHEET_NAME), true, "應建立學員資料總表");
    assertEq(createdSheets.includes(Service_DB.CONFIG.MASTER_SETTINGS_SHEET), true, "應建立系統設定表");
    assertEq(createdSheets.includes(Service_DB.CONFIG.ACTIVITIES_OVERVIEW_SHEET), true, "應建立活動清單表");
    
  } finally {
    // Restore
    SpreadsheetApp.create = originalCreate;
  }
}

// --- 輔助 ---
function assertEq(actual, expected, msg) {
  if (actual !== expected) {
    throw new Error(`${msg} (預期: ${expected}, 實際: ${actual})`);
  }
}

/**
 * 驗證：後台 CMS 新增題目 (Mock)
 */
function test_Admin_AddQuestion() {
  const mockPwd = "correctPropertiesPwd";
  const inputData = { label: "S2", question: "NewQ", type: "單選題" };
  
  // Mock Security
  const origVerify = Service_Security.verifyAdmin;
  Service_Security.verifyAdmin = (p) => p === mockPwd;
  
  // Mock DB
  const origAdd = Service_DB.addQuestionConfig;
  const origGetId = Service_DB.getMasterId;
  
  let addCalled = false;
  
  try {
     Service_DB.getMasterId = () => "MOCK_SS_ID";
     Service_DB.addQuestionConfig = (ssid, data) => {
       addCalled = true;
       if(ssid !== "MOCK_SS_ID") throw new Error("SSID mismatch");
       if(data.question !== "NewQ") throw new Error("Data mismatch");
       return { status: "success" };
     };
     
     // 1. 密碼錯誤
     const resFail = apiAddQuestion("wrong", inputData);
     assertEq(resFail.status, "error", "密碼錯誤應拒絕");
     
     // 2. 密碼正確
     const resOk = apiAddQuestion(mockPwd, inputData);
     assertEq(resOk.status, "success", "密碼正確應成功");
     assertEq(addCalled, true, "應呼叫底層 DB 寫入");
     
  } finally {
    Service_Security.verifyAdmin = origVerify;
    Service_DB.addQuestionConfig = origAdd;
    Service_DB.getMasterId = origGetId;
  }
}

/**
 * 驗證：動態取得下一題 (Mock Engine)
 */
function test_Engine_NextTask() {
  const sid = "S999";
  
  // Mock DB
  const origGetConfig = Service_DB.getActivityConfig;
  const origGetData = Service_DB.getStudentRowData;
  
  try {
     const mockQs = [
       { label:"S1", question:"Q1", targetColumn:"C1", timestampColumn:"T1" },
       { label:"S2", question:"Q2", targetColumn:"C2", timestampColumn:"T2" }
     ];
     Service_DB.getActivityConfig = () => mockQs;
     
     // Case 1: 什麼都沒做 -> Q1
     Service_DB.getStudentRowData = () => ({});
     let res = Service_Engine.getStudentNextTask("ssid", sid);
     assertEq(res.status, 'pending', "應回傳 pending");
     assertEq(res.question.question, 'Q1', "第一題應為 Q1");
     
     // Case 2: Q1 做完了 -> Q2
     Service_DB.getStudentRowData = () => ({ "C1": "Ans", "T1": "Time" });
     res = Service_Engine.getStudentNextTask("ssid", sid);
     assertEq(res.status, 'pending', "應回傳 pending");
     assertEq(res.question.question, 'Q2', "下一題應為 Q2");
     
     // Case 3: 全做完了 -> completed
     Service_DB.getStudentRowData = () => ({ "C1":"A", "T1":"t", "C2":"A", "T2":"t" });
     res = Service_Engine.getStudentNextTask("ssid", sid);
     assertEq(res.status, 'completed', "應回傳 completed");
     
  } finally {
     Service_DB.getActivityConfig = origGetConfig;
     Service_DB.getStudentRowData = origGetData;
  }
}
