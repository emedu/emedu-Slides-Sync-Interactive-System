/**
 * Service_Security.gs - 安全防護與驗證
 * 
 * 職責：
 * 1. 防秒刷 (Spam Protection)
 * 2. 身份驗證 (Identity Verification)
 * 
 * 遷移自 v9.4.2: withSubmitGuard_
 */

const Service_Security = (function() {
  
  const SUBMIT_GUARD_SECONDS = 60; // 預設 60 秒

  return {
    
    /**
     * 檢查並鎖定提交 (防秒刷)
     * @param {string} studentId - 學號
     * @param {string} stageLabel - 階段
     * @returns {boolean} - 是否允許提交 (true=允許, false=阻擋)
     */
    checkSubmitRateLimit: function(studentId, stageLabel) {
      const key = `v10submit:${studentId}:${stageLabel}`;
      const cache = CacheService.getScriptCache();
      
      if (cache.get(key)) {
        return false; // 鎖定中，阻擋
      }
      
      // 設定鎖定
      cache.put(key, "1", SUBMIT_GUARD_SECONDS);
      return true;
    },
    
    /**
     * 驗證學號格式 (簡單範例)
     */
    validateStudentId: function(sid) {
      if (!sid || sid.trim().length === 0) return false;
      return true;
    }
  };
})();
