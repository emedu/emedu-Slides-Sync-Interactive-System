/**
 * Service_Security.gs - 安全防護與驗證
 * 
 * 職責：
 * 1. 防秒刷 (Spam Protection)
 * 2. 身份驗證 (Identity Verification)
 * 
 * 遷移自 v9.4.2: withSubmitGuard_
 * v10.3.0 安全更新：全面導入網頁安裝精靈，移除代碼層級之密碼初始化
 */

const Service_Security = (function() {
  
  const SUBMIT_GUARD_SECONDS = 60; // 預設 60 秒
  const ADMIN_PASSWORD_PROP = 'ADMIN_PASSWORD_HASH'; // 儲存雜湊值的 Property key

  /**
   * 計算字串的 SHA-256 雜湊值 (16 進位字串)
   * @param {string} input
   * @returns {string}
   */
  function _sha256(input) {
    const rawBytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(input),
      Utilities.Charset.UTF_8
    );
    return rawBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  }

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
      
      cache.put(key, "1", SUBMIT_GUARD_SECONDS);
      return true;
    },
    
    /**
     * 驗證學號格式
     * @param {string} sid
     * @returns {boolean}
     */
    validateStudentId: function(sid) {
      if (!sid || sid.trim().length === 0) return false;
      return true;
    },

    /**
     * 驗證管理員密碼 (以 SHA-256 雜湊比對)
     * 注意：必須先呼叫 setAdminPassword() 完成初始化，否則會拒絕所有登入。
     * @param {string} password - 使用者輸入的明文密碼
     * @returns {boolean}
     */
    verifyAdmin: function(password) {
      const storedHash = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROP);
      
      // 安全政策：若尚未設定密碼，一律拒絕，避免以預設值登入
      if (!storedHash) {
        console.warn('[Security] ADMIN_PASSWORD_HASH 尚未設定，請先執行 Service_Security.setAdminPassword() 初始化密碼。');
        return false;
      }
      
      return _sha256(password) === storedHash;
    },

    /**
     * 🟢 [手動執行] 設定管理員密碼
     * 請在 Apps Script 編輯器中直接呼叫此函式（選擇函式後點「執行」）。
     * 密碼會以 SHA-256 雜湊形式儲存至 Script Properties，明文不會被保留。
     * 
     * @param {string} newPassword - 新密碼（至少 8 字元，建議含英數混合）
     */
    setAdminPassword: function(newPassword) {
      if (!newPassword || String(newPassword).trim().length < 8) {
        throw new Error('密碼長度不足，請至少設定 8 個字元的密碼。');
      }
      const hash = _sha256(String(newPassword).trim());
      PropertiesService.getScriptProperties().setProperty(ADMIN_PASSWORD_PROP, hash);
      console.log('✅ 管理員密碼已更新（已雜湊儲存）。請妥善保管您的密碼，系統無法還原明文。');
    }
  };
})();

// --- 註：管理密碼初始化現已遷移至網頁安裝精靈 (UI_Setup.html) 與試算表選單 ---
