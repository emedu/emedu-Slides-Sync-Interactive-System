/**
 * Service_AI.gs - Gemini AI 整合模組
 * 
 * 職責：
 * 1. 介接 Gemini API
 * 2. 生成學習筆記與建議
 * 
 * (目前為 Placeholder)
 */

const Service_AI = (function() {
  
  const API_KEY_PROP = 'GEMINI_API_KEY';

  return {
    
    /**
     * 分析作答並提供建議
     */
    analyzeResponse: function(question, answer, context) {
      // TODO: 實作 Gemini API 呼叫
      return {
        feedback: "AI 分析功能尚未啟用",
        suggestions: []
      };
    },
    
    /**
     * 生成個人化筆記
     */
    generateSummary: function(studentData) {
      // TODO: 彙整學員該次活動的作答，生成重點摘要
      return "自動生成的學習筆記...";
    }
  };
})();
