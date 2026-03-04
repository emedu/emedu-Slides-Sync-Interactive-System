/**
 * Service_Engine.gs - 教學引擎與評分核心
 * 
 * 職責：
 * 1. 處理評分邏輯 (Scoring)
 * 2. 判斷學習進度
 * 3. 處理作答提交的核心業務流程
 * 
 * 遷移自 v9.4.2: Scoring behavior
 */

const Service_Engine = (function() {

  // --- 內部評分邏輯 (遷移自 v9.4.2 Scoring) ---
  const Scoring = {
    toHalfWidth: s => String(s||"").replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0)).replace(/\u3000/g," "),
    normalize: s => Scoring.toHalfWidth(s).trim().toLowerCase(),
    splitOptions: s => String(s||"").split(",").map(x=>Scoring.normalize(x)).filter(Boolean),
    normalizeAnswerForStore: resp => Array.isArray(resp) ? resp.join(", ") : String(resp||""), // Fix: Added missing function
    
    scoreAnswer: function(q, rawResp) {
      let score = 0;
      if (q.type === "單選題" || q.type === "簡答題") {
        const ans = Scoring.normalize(rawResp);
        const std = Scoring.normalize(q.standardAnswer);
        if (std && ans === std) score = q.score;
      } else if (q.type === "多選題") {
        const respArr = Array.isArray(rawResp) ? rawResp.map(Scoring.normalize) : Scoring.splitOptions(rawResp);
        const stdArr = Scoring.splitOptions(q.standardAnswer);
        const respSet = new Set(respArr), stdSet = new Set(stdArr);
        
        // 預設採用 partialWithPenalty 模式 (參照 v9.4.2 CONFIG)
        // 若需支援其他模式，可從 Service_DB 讀取 CONFIG
        const mode = "partialWithPenalty"; 
        
        if (mode === "exact") {
           const sameSize = respSet.size === stdSet.size; 
           let allIn = true; respSet.forEach(x=>{ if(!stdSet.has(x)) allIn=false; });
           if (sameSize && allIn) score = q.score;
        } else {
           // partialWithPenalty
           const totalStd = stdSet.size;
           if (totalStd > 0) {
             let correct=0, wrong=0;
             respSet.forEach(x=>{ if(stdSet.has(x)) correct++; else wrong++; });
             const raw = (correct/totalStd) - (wrong/Math.max(respSet.size,1));
             score = Math.max(0, Math.round(raw * q.score));
           }
        }
      }
      return score;
    }
  };

  return {
    /**
     * 處理學員提交的答案
     * @param {string} runId - 活動ID
     * @param {string} studentId - 學號
     * @param {string} stageLabel - 階段標籤
     * @param {Object} answers -作答內容 { questionText: answerValue }
     */
    /**
     * 處理學員提交的答案
     * @param {string} runId - 活動ID (對應主控台或是單一活動資料簿)
     * @param {string} studentId - 學號
     * @param {string} stageLabel - 階段標籤
     * @param {Object} answers -作答內容 { questionText: answerValue }
     */
    processSubmission: function(runId, studentId, stageLabel, answers) {
      if (!runId || !studentId || !stageLabel) {
        return { status: 'error', message: '缺少必要參數' };
      }

      // 1. 取得活動設定 (需從 Service_DB 讀取設定表)
      // 若是單機版(活動資料簿本身)，直接讀取快照或設定
      const ssId = runId; // 假設 runId 即為 Spreadsheet ID (或需 lookup)
      const allQs = Service_DB.getActivityConfig(ssId); 
      
      // 篩選出該階段的題目
      const stageQs = allQs.filter(q => q.label === stageLabel);
      if (stageQs.length === 0) {
        return { status: 'error', message: '找不到該階段的題目設定' };
      }
      
      let totalStageScore = 0;
      
      // 2. 遍歷題目計算分數並寫入
      const updates = {};
      const now = new Date();
      
      // 2. 遍歷題目計算分數並寫入
      stageQs.forEach(q => {
        const rawAns = answers[q.question];
        if (rawAns !== undefined) {
          // 評分
          const score = Scoring.scoreAnswer(q, rawAns);
          totalStageScore += score;
          
          // 收集更新
          updates[q.targetColumn] = Scoring.normalizeAnswerForStore(rawAns);
          if (q.scoreColumn) updates[q.scoreColumn] = score;
          if (q.timestampColumn) updates[q.timestampColumn] = now;
        }
      });

      // 批次寫入資料表 (Data Sheet) - 使用動態路由，寫到當前活動對應分頁
      if (Object.keys(updates).length > 0) {
        Service_DB.updateStudentData(
          ssId, 
          Service_DB.getActiveDataSheetName(),
          studentId, 
          answers['Email'] || null,
          updates
        );
      }

      // 3. 更新進度追蹤表 (Tracking Sheet) - 動態路由
      const completeInfo = this._calculateTrackingStatus(allQs, answers, studentId, ssId);
      Service_DB.updateTrackingData(
        ssId,
        Service_DB.getActiveTrackingSheetName(),
        studentId,
        completeInfo.stageMarks,
        completeInfo.completedCount,
        completeInfo.totalAccumulatedScore
      );
      
      // 4. (Optimization) 呼叫 AI 分析回饋
      let aiFeedback = null;
      try {
        // 僅针对簡答題或特定情境觸發 AI
        // 這裡做簡單示範：所有題目都嘗試取得回饋
        const qConfig = stageQs.find(q => q.question in answers);
        if (qConfig) {
             const ansText = answers[qConfig.question];
             const aiRes = Service_AI.analyzeResponse(qConfig.question, ansText);
             aiFeedback = aiRes.feedback;
        }
      } catch(e) {
        console.log("AI Feedback Error: " + e.toString());
      }
      
      return {
        status: 'success',
        message: '已儲存作答',
        progress: completeInfo.completedCount,
        feedback: aiFeedback // 回傳 AI 回饋
      };
    },
    
    /**
     * 計算追蹤狀態 (內部用)
     * 遷移自 TriggerHandler._updateTrackingSheet 邏輯
     */
    _calculateTrackingStatus: function(allQs, currentAnswers, studentId, ssId) {
      // 需要取得該生目前在資料表的所有作答紀錄，以判斷歷史階段是否完成
      // 這裡簡化：假設 Service_DB 可以取得該生完整 row data map
      const studentData = Service_DB.getStudentRowData(ssId, studentId);
      // 注意：currentAnswers 的 key 是 questionText，但 studentData 的 key 可能是 targetColumn
      
      // 合併當前作答 (currentAnswers) 到 studentData (模擬更新後狀態)
      const mergedData = { ...studentData, ...currentAnswers };
      // 注意：currentAnswers 的 key 是 questionText，但 studentData 的 key 可能是 targetColumn
      // 這裡需做 mapping，為簡化邏輯，假設 Service_DB.getStudentRowData 回傳的是 { columnHeader: value }
      
      const stages = this._buildStageInfo(allQs);
      let completedCount = 0;
      let totalAccumulatedScore = 0;
      
      const stageMarks = stages.map(st => {
        const qs = allQs.filter(q => q.label === st.label);
        const req = Math.max(...qs.map(q => q.requiredCount || 0), 0);
        
        // 檢查每一題是否已作答 (檢查 Timestamp 欄位是否有值)
        // 若是當次提交，檢查 answers 是否有值
        const doneFlags = qs.map(q => {
             const timeVal = mergedData[q.timestampColumn];
             // 或是當次提交有答案
             const hasCurrent = !!currentAnswers[q.question]; 
             return (!!timeVal) || hasCurrent;
        });
        
        const doneCount = doneFlags.filter(Boolean).length;
        const isComplete = req > 0 ? (doneCount >= req) : doneFlags.every(Boolean);
        
        if (isComplete) completedCount++;
        
        // 計算總分
        qs.forEach(q => {
           const sVal = parseFloat(mergedData[q.scoreColumn] || 0);
           if (!isNaN(sVal)) totalAccumulatedScore += sVal;
        });

        return isComplete ? "✔" : "";
      });
      
      return { stageMarks, completedCount, totalAccumulatedScore };
    },

    /**
     * 取得學員下一題 (NEW)
     */
    getStudentNextTask: function(runId, studentId) {
       const allQs = Service_DB.getActivityConfig(runId);
       // 讀取學員資料：Map<HeaderName, Value>
        // 1. 驗證學員是否存在
        const studentData = Service_DB.getStudentRowData(runId, studentId);
        if (!studentData || Object.keys(studentData).length === 0) {
           throw new Error("查無此學號，請檢查輸入是否正確或聯繫老師。");
        }

       // 依序尋找未作答題目
       // 假設 allQs 已經是正確順序 (readConfig 讀取順序)
       for (const q of allQs) {
           // 判定是否已作答：檢查 Timestamp 或 Target Column 是否有值
           const tsVal = studentData[q.timestampColumn];
           const ansVal = studentData[q.targetColumn];
           const isDone = (tsVal && String(tsVal).trim() !== "") || (ansVal && String(ansVal).trim() !== "");
           
           if (!isDone) {
               return { status: 'pending', question: q };
           }
       }
       
       return { status: 'completed' };
    },

    _buildStageInfo: function(allQs) {
        const grouped = {};
        allQs.forEach(q => { if (!grouped[q.label]) grouped[q.label]=[]; grouped[q.label].push(q); });
        // 依照 rowIndex 排序
        return Object.keys(grouped).map(label=>{
             const qs = grouped[label];
             const first = Math.min(...qs.map(q=>q.rowIndex));
             return { label, firstRow: first };
        }).sort((a,b)=>a.firstRow-b.firstRow);
    },

    
    // 公開工具給前端或測試用
    calculateScore: function(questionObj, userAnswer) {
      return Scoring.scoreAnswer(questionObj, userAnswer);
    }
  };
})();
