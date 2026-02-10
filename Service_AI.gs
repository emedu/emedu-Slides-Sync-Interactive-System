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
  const MODEL_NAME = 'gemini-1.5-flash'; // 使用較快速且經濟的模型

  return {
    
    /**
     * 分析作答並提供建議
     */
    analyzeResponse: function(question, answer, context) {
      const apiKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
      
      if (!apiKey) {
        return {
          feedback: "（AI 功能尚未啟用：請於專案設定中設定 GEMINI_API_KEY）",
          suggestions: []
        };
      }

      // 建構提示詞 (Prompt)
      const prompt = `
        你是一位專業的簡報與溝通教練。請針對以下學員的作答提供簡短、建設性的回饋。
        
        題目：${question}
        學員回答：${answer}
        
        請提供：
        1. 評語 (50字以內)
        2. 一個具體的改進建議
        
        回傳格式請直接給予純文字建議即可。
      `;

      try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
        const payload = {
          contents: [{
            parts: [{ text: prompt }]
          }]
        };

        const options = {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        if (json.error) {
          console.error("Gemini API Error:", json.error);
          return { feedback: "無法取得 AI 建議 (API Error)", suggestions: [] };
        }

        const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (text) {
          return {
            feedback: text.trim(),
            suggestions: [] // 簡單起見，將建議包含在 feedback 文字中
          };
        } else {
           return { feedback: "AI 無法產生回應", suggestions: [] };
        }

      } catch (e) {
        console.error("AI Service Exception:", e);
        return { feedback: "連線發生錯誤，請稍後再試", suggestions: [] };
      }
    },
    
    /**
     * 生成個人化筆記
     */
    generateSummary: function(studentData) {
      // TODO: 彙整學員該次活動的作答，生成重點摘要
      return "自動生成的學習筆記功能開發中...";
    }
  };
})();
