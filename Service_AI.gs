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

      // 構建系統指令與提示詞 (NLP 優化: 教練人格與結構化輸出)
      const systemInstruction = `你是一位專業的「emedu 簡報同步互動教練」。您的目標是協助學員提升簡報與溝通技巧。
請針對學員的作答提供專業、正向且具體的回饋。
必須以 JSON 格式回傳，格式如下：
{
  "rating": "優秀|良好|待加強",
  "comment": "50字內的短評",
  "suggestions": ["建議1", "建議2"]
}`;

      const userPrompt = `
課程情境：${context || "通用簡報技巧課程"}
題目：${question}
學員回答：${answer}

請分析上述作答並提供回饋。`;

      try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
        const payload = {
          system_instruction: {
            parts: [{ text: systemInstruction }]
          },
          contents: [{
            parts: [{ text: userPrompt }]
          }],
          generationConfig: {
            response_mime_type: "application/json"
          }
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

        const responseText = json.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (responseText) {
          const aiResult = JSON.parse(responseText);
          return {
            rating: aiResult.rating,
            feedback: aiResult.comment,
            suggestions: aiResult.suggestions || []
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
