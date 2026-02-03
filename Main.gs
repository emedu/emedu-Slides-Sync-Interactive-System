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
  const route = e.parameter.route || 'portal';
  
  if (route === 'portal') {
    return _renderPortal(e);
  } else if (route === 'install') {
    return _handleInstall(e);
  } else {
    return HtmlService.createHtmlOutput("未知路由");
  }
}

/**
 * 渲染學員入口頁面
 */
function _renderPortal(e) {
  // TODO: 檢查活動狀態與學員身分 (呼叫 Service_Security)
  // 如果需要登入，回傳登入頁
  // 如果已登入，呼叫 Service_Engine 取得當前狀態
  
  const template = HtmlService.createTemplateFromFile('UI_Portal');
  template.data = {
    // 假資料，待 Service_Engine 實作後替換
    title: '伊美互動學習',
    activityName: '載入中...'
  };
  
  return template.evaluate()
      .setTitle('伊美：簡報同步互動學習系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 處理安裝請求 (僅供管理員)
 */
function _handleInstall(e) {
  // 簡單的安裝觸發
  try {
    Service_DB.installControlSheet();
    return HtmlService.createHtmlOutput("核心主控台安裝完成。");
  } catch (err) {
    return HtmlService.createHtmlOutput("安裝失敗：" + err.toString());
  }
}

/**
 * 保留 doPost 以備未來擴充 (如自訂表單提交)
 */
function doPost(e) {
  return ContentService.createTextOutput("POST request received");
}
