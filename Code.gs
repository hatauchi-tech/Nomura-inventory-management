/**
 * 在庫管理アプリ - メインエントリーポイント（修正版）
 * 修正内容：エラーハンドリングとログ出力を強化
 * 修正内容（ロギング）：getScriptUrlにロギングラッパーを追加
 */

function doGet(e) {
  try {
    const user = getCurrentUser();
    
    // ユーザーが認証されていない、または無効な場合はログイン画面を表示
    if (!user) {
      Logger.log('doGet: ユーザー情報が取得できませんでした');
      return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('在庫管理システム - ログイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    if (!user.valid) {
      Logger.log('doGet: ユーザーが無効です - ' + user.email);
      return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('在庫管理システム - ログイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // ページパラメータを取得
    const page = e.parameter.page || 'home';
    Logger.log('doGet: ページ表示 - ユーザー: ' + user.name + ', ページ: ' + page);
    
    // ページに応じたHTMLファイルを選択
    let htmlFile;
    let pageTitle = '在庫管理システム';
    
    switch(page) {
      case 'home':
        htmlFile = user.role === '管理者' ? 'HomeAdmin' : 'HomeUser';
        pageTitle += ' - ホーム';
        break;
      case 'stock-movement':
        htmlFile = 'StockMovement';
        pageTitle += ' - 入出庫登録';
        break;
      case 'stock-movement-history':
        htmlFile = 'StockMovementHistory';
        pageTitle += ' - 入出庫履歴確認';
        break;
      case 'inventory-count':
        htmlFile = 'InventoryCount';
        pageTitle += ' - 棚卸カウント入力';
        break;
      case 'inventory-verify':
        htmlFile = 'InventoryVerify';
        pageTitle += ' - 棚卸照合・確定';
        break;
      case 'stock-list':
        htmlFile = 'StockList';
        pageTitle += ' - 在庫一覧・検索';
        break;
      case 'master-management':
        htmlFile = 'MasterManagement';
        pageTitle += ' - マスタ管理';
        break;
      case 'past-stock-adjustment':
        htmlFile = 'PastStockAdjustment';
        pageTitle += ' - 過去データ調整';
        break;
      default:
        Logger.log('doGet: 未知のページ指定 - ' + page);
        htmlFile = user.role === '管理者' ? 'HomeAdmin' : 'HomeUser';
        pageTitle += ' - ホーム';
    }
    
    // テンプレートにユーザー情報を渡す
    const template = HtmlService.createTemplateFromFile(htmlFile);
    template.userName = user.name;
    template.userRole = user.role;
    
    return template.evaluate()
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    Logger.log('doGet Error: ' + error.toString());
    Logger.log('doGet Error Stack: ' + error.stack);
    
    // エラー画面を表示
    const errorHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>エラー - 在庫管理システム</title>
          <style>
            body {
              font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              display: flex;
              justify-content: center;
              align-items: center;
              min-height: 100vh;
              margin: 0;
              padding: 20px;
            }
            .error-container {
              background: white;
              border-radius: 10px;
              box-shadow: 0 10px 40px rgba(0,0,0,0.2);
              padding: 40px;
              max-width: 600px;
              width: 100%;
            }
            h1 {
              color: #e74c3c;
              margin-top: 0;
              font-size: 28px;
            }
            .error-icon {
              font-size: 60px;
              text-align: center;
              margin-bottom: 20px;
            }
            .error-message {
              background: #fff5f5;
              border-left: 4px solid #e74c3c;
              padding: 15px;
              margin: 20px 0;
              border-radius: 4px;
            }
            .error-details {
              color: #666;
              font-size: 14px;
              line-height: 1.6;
            }
            .action-buttons {
              margin-top: 30px;
              display: flex;
              gap: 10px;
              flex-wrap: wrap;
            }
            button {
              background: #667eea;
              color: white;
              border: none;
              padding: 12px 24px;
              border-radius: 5px;
              cursor: pointer;
              font-size: 16px;
              transition: background 0.3s;
            }
            button:hover {
              background: #5568d3;
            }
            .secondary-button {
              background: #95a5a6;
            }
            .secondary-button:hover {
              background: #7f8c8d;
            }
            .help-text {
              margin-top: 20px;
              padding: 15px;
              background: #f8f9fa;
              border-radius: 5px;
              font-size: 14px;
              color: #666;
            }
          </style>
        </head>
        <body>
          <div class="error-container">
            <div class="error-icon">⚠️</div>
            <h1>エラーが発生しました</h1>
            <div class="error-message">
              <strong>エラー内容:</strong>
              <div class="error-details">${error.toString().replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>
            </div>
            <div class="help-text">
              <strong>考えられる原因:</strong>
              <ul>
                <li>スプレッドシートの必要なシートが存在しない</li>
                <li>シート名が正しくない（例: 「T_棚卸履歴」など）</li>
                <li>必要な列（ヘッダー）が存在しない</li>
                <li>アクセス権限の問題</li>
              </ul>
              <strong>対処方法:</strong>
              <ul>
                <li>スプレッドシートの構成を確認してください</li>
                <li>管理者に連絡してください</li>
                <li>Google Apps Scriptのログを確認してください</li>
              </ul>
            </div>
            <div class="action-buttons">
              <button onclick="location.reload()">再読み込み</button>
              <button class="secondary-button" onclick="window.history.back()">戻る</button>
            </div>
          </div>
        </body>
      </html>
    `;
    
    return HtmlService.createHtmlOutput(errorHtml)
      .setTitle('エラー - 在庫管理システム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * HTMLファイルの内容を取得して埋め込む
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('include Error: ' + error.toString() + ' - ファイル名: ' + filename);
    return '<p style="color: red;">ファイル「' + filename + '」の読み込みに失敗しました</p>';
  }
}

/**
 * WebアプリのURLを取得
 */
function getScriptUrl() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  // ServerSide.gsで定義されているloggable関数を使用します
  return loggable('getScriptUrl', arguments, function() {
    try {
      return ScriptApp.getService().getUrl();
    } catch (error) {
      Logger.log('getScriptUrl Error: ' + error.toString());
      return '';
    }
  });
}