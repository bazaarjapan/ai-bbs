// GASの[プロジェクトの設定]から[スクリプトのプロパティ]に以下のキーと値を設定する
// apikey: APIキー
// apikeyの取得先 https://aistudio.google.com/app/apikey
const apikey = PropertiesService.getScriptProperties().getProperty('apikey');
const model = 'gemini-2.5-flash';
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apikey}`;

// 再試行のための定数
const MAX_RETRIES = 3;
const RETRY_DELAY = 1000; // ミリ秒単位

// ロックサービスの初期化
const lockService = LockService.getScriptLock();
const LOCK_TIMEOUT = 30000; // 30秒のタイムアウト

// プロンプト自動生成関数
function generateText(value) {
  const promptText = `role: あなたは連絡用掲示板への文書作成のプロです。
input: |
  ${value}
task: |
  入力文の内容をもとに、連絡用掲示板に適した分かりやすく簡潔な投稿文を作成してください。
conditions:
  - 敬語を使い、丁寧で分かりやすい表現にする。
  - 前置きや後書き、あいさつ文は無で、必要な情報のみ漏れなく伝える。
  - 箇条書きや改行を適切に使い、読みやすくする。
output_format: |
  （掲示板への投稿文をテキスト形式（マークダウン等の装飾なし）で記載してください。）
`; 
  const payload = {
    contents: [{
      parts: [{
        text: promptText
      }]
    }],
    generationConfig: {
      temperature: 0.1,
      topK: 40,
      topP: 0.90,
      maxOutputTokens: 8192,
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  for (let i = 0; i < MAX_RETRIES; i++) {
    try {
      // ロックを取得してAPIを呼び出す
      if (lockService.tryLock(LOCK_TIMEOUT)) {
        try {
          const response = UrlFetchApp.fetch(GEMINI_URL, options);
          const responseJson = JSON.parse(response.getContentText());
          
          if (responseJson.candidates && responseJson.candidates.length > 0) {
            const text = responseJson.candidates[0].content.parts[0].text;
            // console.log(text);
            return text;
          }
          return 'No response from Gemini API';
        } catch (e) {
          console.error('API呼び出しエラー:', e);
          throw e; // エラーを再スローしてリトライ処理に任せる
        } finally {
          lockService.releaseLock();
        }
      } else {
        throw new Error('ロックを取得できませんでした');
      }
    } catch (e) {
      if (i === MAX_RETRIES - 1) {
        console.error('Error in generateText:', e);
        return 'Error retrieving response: ' + e.toString();
      }
      console.log(`再試行 ${i+1}/${MAX_RETRIES}`);
      Utilities.sleep(RETRY_DELAY * (i + 1));
    }
  }
  return 'すべての再試行が失敗しました';
}
