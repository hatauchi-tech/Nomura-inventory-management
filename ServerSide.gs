/**
 * 在庫管理アプリ - サーバーサイドロジック統合（修正版）
 * 修正内容：シート取得時のnullチェックを追加し、適切なエラーメッセージを表示
 * 修正内容（客先対応）：registerStockMovementに客先フィールドを追加
 * 修正内容（検索対応）：searchStocksにカテゴリ1, 2の絞り込み機能を追加
 * 修正内容（検索項目変更）：searchStocksからproductIdの絞り込みを削除
 * 修正内容（ロギング）：google.script.runで呼び出される全関数の実行ログを強化
 * 修正内容（バグ修正）：重複していた「マスタ管理」の関数定義を削除
 */

// ========================================
// ★★★ ロギングラッパー (追加) ★★★
// ========================================

/**
 * 関数実行の開始、終了、エラーをログに記録するラッパー関数
 * @param {string} fnName - 実行する関数名（ログ用）
 * @param {IArguments} args - 実行する関数の引数 (argumentsオブジェクト)
 * @param {Function} fn - 実行する実際の関数
 * @returns {*} - 実行した関数の戻り値
 */
function loggable(fnName, args, fn) {
  // 引数をJSON文字列に変換（長すぎる場合は省略）
  let argsString = '[failed to stringify]';
  try {
    argsString = JSON.stringify(Array.from(args));
    if (argsString.length > 500) {
      argsString = argsString.substring(0, 500) + '... (truncated)';
    }
  } catch (e) {
    // 循環参照などでJSON.stringifyが失敗した場合のフォールバック
    argsString = `[${args.length} arguments]`;
  }
  
  Logger.log(`[START] ${fnName} | Args: ${argsString}`);
  
  try {
    // 実際の関数を実行
    const result = fn();
    
    // 戻り値をJSON文字列に変換（長すぎる場合は省略）
    let resultString = '[failed to stringify]';
    try {
      resultString = JSON.stringify(result);
      if (resultString.length > 500) {
        resultString = resultString.substring(0, 500) + '... (truncated)';
      }
    } catch (e) {
      resultString = '[complex object]';
    }

    Logger.log(`[ END ] ${fnName} | Result: ${resultString}`);
    return result;
    
  } catch (error) {
    // 実行時エラーをログに記録
    Logger.log(`[ERROR] ${fnName} | Error: ${error.toString()} | Stack: ${error.stack}`);
    // エラーを再度スローし、クライアントサイドの.withFailureHandler()でキャッチできるようにする
    throw error;
  }
}


// ========================================
// 設定とユーティリティ
// ========================================

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function generateUniqueId(prefix) {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return prefix + timestamp + random;
}

function getCurrentDateTime() {
  return new Date();
}

function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function getSheetData(sheet) {
  // 【修正】nullチェック追加
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function appendRowToSheet(sheet, rowData) {
  // 【修正】nullチェック追加
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => rowData[header] !== undefined ? rowData[header] : '');
  sheet.appendRow(row);
}

function updateSheetRow(sheet, rowIndex, rowData) {
  // 【修正】nullチェック追加
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => rowData[header] !== undefined ? rowData[header] : '');
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function findRowByColumn(sheet, columnName, value) {
  // 【修正】nullチェック追加
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const columnIndex = headers.indexOf(columnName);
  
  if (columnIndex === -1) return null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnIndex] === value) {
      const row = { _rowIndex: i + 1 };
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      return row;
    }
  }
  return null;
}

function findAllRowsByColumn(sheet, columnName, value) {
  // 【修正】nullチェック追加
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const columnIndex = headers.indexOf(columnName);
  const results = [];
  
  if (columnIndex === -1) return results;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnIndex] === value) {
      const row = { _rowIndex: i + 1 };
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      results.push(row);
    }
  }
  return results;
}

function validateRequiredFields(data, requiredFields) {
  const errors = [];
  requiredFields.forEach(field => {
    if (!data[field] || data[field] === '') {
      errors.push(field + 'は必須項目です');
    }
  });
  return { valid: errors.length === 0, errors: errors };
}

function validateNumber(value, min, max) {
  const num = Number(value);
  if (isNaN(num)) return false;
  if (min !== undefined && num < min) return false;
  if (max !== undefined && num > max) return false;
  return true;
}

// ========================================
// 認証・認可
// ========================================

function getCurrentUser() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getCurrentUser', arguments, function() {
    try {
      const email = Session.getActiveUser().getEmail();
      if (!email) return null;
      
      const ss = getSpreadsheet();
      const userSheet = ss.getSheetByName('M_ユーザー');
      
      // 【修正】nullチェック追加
      if (!userSheet) {
        Logger.log('M_ユーザーシートが見つかりません');
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const data = userSheet.getDataRange().getValues();
      const headers = data[0];
      
      const emailIndex = headers.indexOf('メールアドレス');
      const userIdIndex = headers.indexOf('ユーザーID');
      const nameIndex = headers.indexOf('ユーザー名');
      const roleIndex = headers.indexOf('権限');
      const validIndex = headers.indexOf('有効');
      const deptIndex = headers.indexOf('部門');
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][emailIndex] === email) {
          const isValid = data[i][validIndex] === true || data[i][validIndex] === 'TRUE';
          return {
            userId: data[i][userIdIndex],
            name: data[i][nameIndex],
            email: data[i][emailIndex],
            role: data[i][roleIndex],
            department: data[i][deptIndex],
            valid: isValid
          };
        }
      }
      return null;
    } catch (error) {
      Logger.log('getCurrentUser Error: ' + error.toString());
      return null;
    }
  });
}

function checkUserPermission(requiredRole) {
  const user = getCurrentUser();
  if (!user || !user.valid) return false;
  if (user.role === '管理者') return true;
  return user.role === requiredRole;
}

function requireAdminPermission() {
  if (!checkUserPermission('管理者')) {
    throw new Error('この操作には管理者権限が必要です');
  }
}

// ========================================
// 通知機能
// ========================================

/**
 * 在庫が適正在庫数を下回った際にメール通知を送信
 * @param {string} productId - 製品ID
 * @param {string} locationId - 保管場所ID
 * @param {number} currentStock - 現在在庫数
 * @param {number} optimalStock - 適正在庫数
 */
function sendLowStockAlert(productId, locationId, currentStock, optimalStock) {
  return loggable('sendLowStockAlert', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const productSheet = ss.getSheetByName('M_製品');
      const locationSheet = ss.getSheetByName('M_保管場所');
      const userSheet = ss.getSheetByName('M_ユーザー');

      if (!productSheet || !locationSheet || !userSheet) {
        Logger.log('sendLowStockAlert: 必要なシートが見つかりません');
        return;
      }

      // 製品情報を取得
      const product = findRowByColumn(productSheet, '製品ID', productId);
      if (!product) {
        Logger.log('sendLowStockAlert: 製品が見つかりません - 製品ID: ' + productId);
        return;
      }

      const department = product['担当部署'];
      if (!department || department === '') {
        Logger.log('sendLowStockAlert: 担当部署が設定されていません - 製品ID: ' + productId);
        return;
      }

      // 保管場所情報を取得
      const location = findRowByColumn(locationSheet, '保管場所ID', locationId);
      const locationName = location ? location['場所名'] : locationId;

      // 担当部署の管理者を検索
      const users = getSheetData(userSheet);
      Logger.log('sendLowStockAlert: 全ユーザー数: ' + users.length);

      // デバッグ: 担当部署と一致するユーザーを確認
      const departmentUsers = users.filter(u => u['部門'] === department);
      Logger.log('sendLowStockAlert: 担当部署「' + department + '」のユーザー数: ' + departmentUsers.length);

      const managers = users.filter(u =>
        u['部門'] === department &&
        u['権限'] === '管理者' &&
        u['有効'] === true &&
        u['メールアドレス'] && u['メールアドレス'] !== ''
      );

      Logger.log('sendLowStockAlert: 検索条件 - 部署: ' + department + ', 権限: 管理者, 有効: true, メールアドレス: 必須');
      Logger.log('sendLowStockAlert: 該当する管理者数: ' + managers.length);

      if (managers.length === 0) {
        Logger.log('sendLowStockAlert: 担当部署の管理者が見つかりません - 部署: ' + department);
        // デバッグ: 部署が一致するユーザーの詳細を表示
        if (departmentUsers.length > 0) {
          departmentUsers.forEach(u => {
            Logger.log('  - ユーザー: ' + u['ユーザー名'] + ', 権限: ' + u['権限'] + ', 有効: ' + u['有効'] + ', メール: ' + u['メールアドレス']);
          });
        }
        return;
      }

      // メール本文を作成
      const subject = '【在庫アラート】適正在庫数を下回りました';
      const body = `
在庫管理システムからの通知

以下の製品の在庫が適正在庫数を下回りました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
製品名: ${product['製品名']}
製品ID: ${productId}
カテゴリ: ${product['カテゴリ1']} / ${product['カテゴリ2']}
保管場所: ${locationName}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

現在在庫数: ${currentStock}
適正在庫数: ${optimalStock}
不足数: ${optimalStock - currentStock}

担当部署: ${department}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

在庫の補充をご検討ください。

※このメールは自動送信されています。
`;

      // 各管理者にメール送信
      Logger.log('sendLowStockAlert: メール送信を開始します - 送信先: ' + managers.length + '名');
      managers.forEach((manager, index) => {
        try {
          Logger.log('sendLowStockAlert: [' + (index + 1) + '/' + managers.length + '] メール送信中 - 宛先: ' + manager['メールアドレス'] + ', 名前: ' + manager['ユーザー名']);
          MailApp.sendEmail({
            to: manager['メールアドレス'],
            subject: subject,
            body: body
          });
          Logger.log('sendLowStockAlert: [' + (index + 1) + '/' + managers.length + '] メール送信完了 - 宛先: ' + manager['メールアドレス']);
        } catch (error) {
          Logger.log('sendLowStockAlert: [' + (index + 1) + '/' + managers.length + '] メール送信失敗 - 宛先: ' + manager['メールアドレス'] + ', エラー: ' + error.toString());
        }
      });

      Logger.log('sendLowStockAlert: メール送信処理完了');

    } catch (error) {
      Logger.log('sendLowStockAlert Error: ' + error.toString());
    }
  });
}

// ========================================
// 入出庫管理
// ========================================

function getStorageLocations() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getStorageLocations', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      return getSheetData(sheet);
    } catch (error) {
      Logger.log('getStorageLocations Error: ' + error.toString());
      throw new Error('保管場所の取得に失敗しました: ' + error.message);
    }
  });
}

function getCategory1List() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getCategory1List', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const data = getSheetData(sheet);
      
      const category1Set = new Set();
      data.forEach(row => {
        if (row['カテゴリ1'] && row['有効'] === true) {
          category1Set.add(row['カテゴリ1']);
        }
      });
      return Array.from(category1Set).sort();
    } catch (error) {
      Logger.log('getCategory1List Error: ' + error.toString());
      throw new Error('カテゴリ1リストの取得に失敗しました: ' + error.message);
    }
  });
}

function getCategory2List(category1) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getCategory2List', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const data = getSheetData(sheet);
      
      const category2Set = new Set();
      data.forEach(row => {
        if (row['カテゴリ1'] === category1 && row['カテゴリ2'] && row['有効'] === true) {
          category2Set.add(row['カテゴリ2']);
        }
      });
      return Array.from(category2Set).sort();
    } catch (error) {
      Logger.log('getCategory2List Error: ' + error.toString());
      throw new Error('カテゴリ2リストの取得に失敗しました: ' + error.message);
    }
  });
}

function getProductsByCategory(category1, category2) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getProductsByCategory', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const data = getSheetData(sheet);
      
      return data.filter(row => {
        const matchCategory1 = !category1 || row['カテゴリ1'] === category1;
        const matchCategory2 = !category2 || row['カテゴリ2'] === category2;
        return matchCategory1 && matchCategory2 && row['有効'] === true;
      });
    } catch (error) {
      Logger.log('getProductsByCategory Error: ' + error.toString());
      throw new Error('製品の取得に失敗しました: ' + error.message);
    }
  });
}

function getCurrentStock(productId, locationId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getCurrentStock', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('T_在庫');

      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      const data = getSheetData(sheet);

      // 【修正】型を統一して比較（String型に変換）
      const stock = data.find(row =>
        String(row['製品ID']) === String(productId) && String(row['保管場所ID']) === String(locationId)
      );
      return stock ? Number(stock['現在在庫数']) : 0;
    } catch (error) {
      Logger.log('getCurrentStock Error: ' + error.toString());
      throw new Error('在庫数の取得に失敗しました: ' + error.message);
    }
  });
}

function registerStockMovement(data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('registerStockMovement', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      // 【修正】入庫時は管理者のみ、出庫は全員OK
      if (data.type === '入庫' && user.role !== '管理者') {
        throw new Error('入庫登録は管理者のみ実行できます');
      }

      const validation = validateRequiredFields(data, ['type', 'locationId', 'productId', 'quantity']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      if (!validateNumber(data.quantity, 1)) {
        throw new Error('数量は1以上の数値を入力してください');
      }

      const quantity = Number(data.quantity);
      const ss = getSpreadsheet();
      
      if (data.type === '出庫') {
        const currentStock = getCurrentStock(data.productId, data.locationId);
        if (currentStock < quantity) {
          throw new Error(`在庫不足: 現在在庫数${currentStock}に対して${quantity}の出庫はできません`);
        }
      }
      
      const historySheet = ss.getSheetByName('T_入出庫履歴');
      
      // 【修正】nullチェック追加
      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const historyId = generateUniqueId('H');
      const historyData = {
        '履歴ID': historyId,
        '製品ID': data.productId,
        '数量': quantity,
        '入出庫タイプ': data.type,
        '現場名': data.siteName || '',
        '客先': data.customerName || '', // (客先対応)
        '発生日時': getCurrentDateTime(),
        '操作ユーザーID': user.userId,
        '保管場所ID': data.locationId
      };
      appendRowToSheet(historySheet, historyData);

      const newStock = updateStock(data.productId, data.locationId, quantity, data.type);

      // 【追加】出庫時に適正在庫数をチェックし、下回った場合にメール通知
      if (data.type === '出庫') {
        Logger.log('registerStockMovement: 出庫処理 - 製品ID: ' + data.productId + ', 新在庫数: ' + newStock);
        const productSheet = ss.getSheetByName('M_製品');
        if (productSheet) {
          const product = findRowByColumn(productSheet, '製品ID', data.productId);
          if (product) {
            const optimalStock = Number(product['適正在庫数'] || 0);
            const department = product['担当部署'] || '';
            Logger.log('registerStockMovement: 製品情報取得 - 製品名: ' + product['製品名'] + ', 適正在庫数: ' + optimalStock + ', 担当部署: ' + department);

            if (optimalStock > 0 && newStock < optimalStock) {
              // 適正在庫数を下回った場合、メール通知を送信
              Logger.log('registerStockMovement: 在庫アラート条件満たす - メール通知送信開始');
              sendLowStockAlert(data.productId, data.locationId, newStock, optimalStock);
            } else if (optimalStock > 0) {
              Logger.log('registerStockMovement: 適正在庫数は設定されているが、まだ下回っていない - 新在庫数: ' + newStock + ', 適正在庫数: ' + optimalStock);
            } else {
              Logger.log('registerStockMovement: 適正在庫数が設定されていない（0または未設定）');
            }
          } else {
            Logger.log('registerStockMovement: 製品情報が取得できませんでした - 製品ID: ' + data.productId);
          }
        } else {
          Logger.log('registerStockMovement: M_製品シートが取得できませんでした');
        }
      }

      return {
        success: true,
        message: data.type + 'を登録しました',
        historyId: historyId,
        newStock: newStock
      };
    } catch (error) {
      Logger.log('registerStockMovement Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function updateStock(productId, locationId, quantity, type) {
  const ss = getSpreadsheet();
  const stockSheet = ss.getSheetByName('T_在庫');

  // 【修正】nullチェック追加
  if (!stockSheet) {
    throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
  }

  const data = stockSheet.getDataRange().getValues();
  const headers = data[0];

  const productIdIndex = headers.indexOf('製品ID');
  const locationIdIndex = headers.indexOf('保管場所ID');
  const stockIndex = headers.indexOf('現在在庫数');
  const dateIndex = headers.indexOf('最終更新日時');

  let found = false;
  let newStockQuantity;

  for (let i = 1; i < data.length; i++) {
    // 【修正】型を統一して比較（String型に変換）
    if (String(data[i][productIdIndex]) === String(productId) && String(data[i][locationIdIndex]) === String(locationId)) {
      const currentQuantity = Number(data[i][stockIndex]);
      newStockQuantity = type === '入庫' ? currentQuantity + quantity : currentQuantity - quantity;
      stockSheet.getRange(i + 1, stockIndex + 1).setValue(newStockQuantity);
      stockSheet.getRange(i + 1, dateIndex + 1).setValue(getCurrentDateTime());
      found = true;
      break;
    }
  }

  if (!found) {
    newStockQuantity = type === '入庫' ? quantity : 0;
    const stockId = generateUniqueId('S');
    const newStockData = {
      '在庫ID': stockId,
      '製品ID': productId,
      '保管場所ID': locationId,
      '現在在庫数': newStockQuantity,
      '最終更新日時': getCurrentDateTime()
    };
    appendRowToSheet(stockSheet, newStockData);
  }

  return newStockQuantity;
}

/**
 * 過去の入出庫履歴を登録する（在庫連動なし）
 * 棚卸後の理論在庫調整用
 * @param {Object} data - 登録データ
 * @returns {Object} 結果オブジェクト
 */
function registerPastStockMovement(data) {
  return loggable('registerPastStockMovement', arguments, function() {
    try {
      requireAdminPermission(); // 管理者のみ実行可能

      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      // 必須項目チェック
      const validation = validateRequiredFields(data, ['type', 'locationId', 'productId', 'quantity', 'occurrenceDate']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      // 数量チェック
      if (!validateNumber(data.quantity, 1)) {
        throw new Error('数量は1以上の数値を入力してください');
      }

      // 日付チェック（過去3ヶ月以内）
      const occurrenceDate = new Date(data.occurrenceDate);
      const today = new Date();
      const threeMonthsAgo = new Date();
      threeMonthsAgo.setMonth(today.getMonth() - 3);

      if (occurrenceDate > today) {
        throw new Error('未来の日付は指定できません');
      }
      if (occurrenceDate < threeMonthsAgo) {
        throw new Error('過去3ヶ月以内の日付のみ指定できます');
      }

      const quantity = Number(data.quantity);
      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_入出庫履歴');

      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      // 履歴レコードを作成（在庫連動なし）
      const historyId = generateUniqueId('H');
      const historyData = {
        '履歴ID': historyId,
        '製品ID': data.productId,
        '数量': quantity,
        '入出庫タイプ': data.type + '(過去調整)', // タイプに「(過去調整)」を付加
        '現場名': data.siteName || '過去データ調整',
        '客先': data.customerName || '',
        '発生日時': occurrenceDate,
        '操作ユーザーID': user.userId,
        '保管場所ID': data.locationId
      };
      appendRowToSheet(historySheet, historyData);

      return {
        success: true,
        message: '過去の' + data.type + 'を登録しました（在庫連動なし）',
        historyId: historyId,
        warning: 'この操作は現在の在庫数には影響しません。理論在庫の記録のみ追加されました。'
      };
    } catch (error) {
      Logger.log('registerPastStockMovement Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

// ========================================
// 入出庫履歴管理
// ========================================

function getMyStockMovements(filters) {
  // ★★★ ロギングラッパーで囲む (新規) ★★★
  return loggable('getMyStockMovements', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_入出庫履歴');
      const productSheet = ss.getSheetByName('M_製品');
      const locationSheet = ss.getSheetByName('M_保管場所');
      const userSheet = ss.getSheetByName('M_ユーザー');

      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。');
      }
      if (!productSheet) {
        throw new Error('M_製品シートが見つかりません。');
      }
      if (!locationSheet) {
        throw new Error('M_保管場所シートが見つかりません。');
      }
      if (!userSheet) {
        throw new Error('M_ユーザーシートが見つかりません。');
      }

      const histories = getSheetData(historySheet);
      const products = getSheetData(productSheet);
      const locations = getSheetData(locationSheet);
      const users = getSheetData(userSheet);

      // 【修正】管理者は全履歴、現場担当者は自分の履歴のみ
      let filteredHistories = user.role === '管理者'
        ? histories
        : histories.filter(row => row['操作ユーザーID'] === user.userId);

      // 絞り込み条件を適用
      if (filters) {
        if (filters.type && filters.type !== '') {
          filteredHistories = filteredHistories.filter(row => row['入出庫タイプ'] === filters.type);
        }
        if (filters.locationId && filters.locationId !== '') {
          filteredHistories = filteredHistories.filter(row => String(row['保管場所ID']) === String(filters.locationId));
        }
        if (filters.category1 && filters.category1 !== '') {
          filteredHistories = filteredHistories.filter(row => {
            const product = products.find(p => String(p['製品ID']) === String(row['製品ID']));
            return product && product['カテゴリ1'] === filters.category1;
          });
        }
        if (filters.category2 && filters.category2 !== '') {
          filteredHistories = filteredHistories.filter(row => {
            const product = products.find(p => String(p['製品ID']) === String(row['製品ID']));
            return product && product['カテゴリ2'] === filters.category2;
          });
        }
        // 【追加】発生日時の絞り込み（開始日）
        if (filters.dateFrom && filters.dateFrom !== '') {
          const fromDate = new Date(filters.dateFrom);
          fromDate.setHours(0, 0, 0, 0);
          filteredHistories = filteredHistories.filter(row => {
            if (!row['発生日時']) return false;
            const rowDate = new Date(row['発生日時']);
            return rowDate >= fromDate;
          });
        }
        // 【追加】発生日時の絞り込み（終了日）
        if (filters.dateTo && filters.dateTo !== '') {
          const toDate = new Date(filters.dateTo);
          toDate.setHours(23, 59, 59, 999);
          filteredHistories = filteredHistories.filter(row => {
            if (!row['発生日時']) return false;
            const rowDate = new Date(row['発生日時']);
            return rowDate <= toDate;
          });
        }
        // 【追加】現場名の絞り込み（部分一致）
        if (filters.siteName && filters.siteName !== '') {
          filteredHistories = filteredHistories.filter(row => {
            const siteName = String(row['現場名'] || '');
            return siteName.indexOf(filters.siteName) !== -1;
          });
        }
        // 【追加】お客様名の絞り込み（部分一致）
        if (filters.customerName && filters.customerName !== '') {
          filteredHistories = filteredHistories.filter(row => {
            const customerName = String(row['客先'] || '');
            return customerName.indexOf(filters.customerName) !== -1;
          });
        }
      }

      // 製品名、場所名、ユーザー名を追加
      const result = filteredHistories.map(history => {
        const product = products.find(p => String(p['製品ID']) === String(history['製品ID']));
        const location = locations.find(l => String(l['保管場所ID']) === String(history['保管場所ID']));
        const operationUser = users.find(u => String(u['ユーザーID']) === String(history['操作ユーザーID']));

        return {
          履歴ID: String(history['履歴ID'] || ''),
          製品ID: String(history['製品ID'] || ''),
          製品名: product ? String(product['製品名'] || '') : '',
          カテゴリ1: product ? String(product['カテゴリ1'] || '') : '',
          カテゴリ2: product ? String(product['カテゴリ2'] || '') : '',
          数量: Number(history['数量'] || 0),
          入出庫タイプ: String(history['入出庫タイプ'] || ''),
          現場名: String(history['現場名'] || ''),
          客先: String(history['客先'] || ''),
          保管場所ID: String(history['保管場所ID'] || ''),
          場所名: location ? String(location['場所名'] || '') : '',
          発生日時: history['発生日時'] ? formatDateTime(history['発生日時']) : '',
          操作ユーザーID: String(history['操作ユーザーID'] || ''),
          操作ユーザー名: operationUser ? String(operationUser['ユーザー名'] || '') : ''
        };
      });

      // 発生日時で降順ソート（新しい順）
      result.sort((a, b) => {
        if (a.発生日時 > b.発生日時) return -1;
        if (a.発生日時 < b.発生日時) return 1;
        return 0;
      });

      return result;

    } catch (error) {
      Logger.log('getMyStockMovements Error: ' + error.toString());
      throw new Error('入出庫履歴の取得に失敗しました: ' + error.message);
    }
  });
}

function updateStockMovement(historyId, newData) {
  // ★★★ ロギングラッパーで囲む (新規) ★★★
  return loggable('updateStockMovement', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_入出庫履歴');

      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。');
      }

      // 既存の履歴を取得
      const existingHistory = findRowByColumn(historySheet, '履歴ID', historyId);
      if (!existingHistory) {
        throw new Error('指定された履歴が見つかりません');
      }

      // 自分の履歴かチェック
      if (existingHistory['操作ユーザーID'] !== user.userId) {
        throw new Error('他のユーザーの履歴は編集できません');
      }

      // バリデーション
      const validation = validateRequiredFields(newData, ['quantity']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      if (!validateNumber(newData.quantity, 1)) {
        throw new Error('数量は1以上の数値を入力してください');
      }

      const newQuantity = Number(newData.quantity);
      const oldQuantity = Number(existingHistory['数量']);
      const productId = existingHistory['製品ID'];
      const locationId = existingHistory['保管場所ID'];
      const movementType = existingHistory['入出庫タイプ'];

      // 出庫の場合、在庫不足チェック
      if (movementType === '出庫') {
        const currentStock = getCurrentStock(productId, locationId);
        const stockAfterRevert = currentStock + oldQuantity; // 旧数量を戻す
        if (stockAfterRevert < newQuantity) {
          throw new Error(`在庫不足: 現在在庫数${currentStock}に対して${newQuantity}の出庫はできません`);
        }
      }

      // 在庫を調整（旧数量を戻して新数量を適用）
      // 1. 旧数量を逆操作で戻す
      const revertType = movementType === '入庫' ? '出庫' : '入庫';
      updateStock(productId, locationId, oldQuantity, revertType);

      // 2. 新数量を適用
      updateStock(productId, locationId, newQuantity, movementType);

      // 履歴を更新
      const updatedHistory = {
        '履歴ID': historyId,
        '製品ID': productId,
        '数量': newQuantity,
        '入出庫タイプ': movementType,
        '現場名': newData.siteName !== undefined ? newData.siteName : existingHistory['現場名'],
        '客先': newData.customerName !== undefined ? newData.customerName : existingHistory['客先'],
        '発生日時': existingHistory['発生日時'],
        '操作ユーザーID': user.userId,
        '保管場所ID': locationId
      };
      updateSheetRow(historySheet, existingHistory._rowIndex, updatedHistory);

      return { success: true, message: '入出庫履歴を更新しました' };

    } catch (error) {
      Logger.log('updateStockMovement Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function deleteStockMovement(historyId) {
  // ★★★ ロギングラッパーで囲む (新規) ★★★
  return loggable('deleteStockMovement', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_入出庫履歴');

      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。');
      }

      // 既存の履歴を取得
      const existingHistory = findRowByColumn(historySheet, '履歴ID', historyId);
      if (!existingHistory) {
        throw new Error('指定された履歴が見つかりません');
      }

      // 自分の履歴かチェック
      if (existingHistory['操作ユーザーID'] !== user.userId) {
        throw new Error('他のユーザーの履歴は削除できません');
      }

      const quantity = Number(existingHistory['数量']);
      const productId = existingHistory['製品ID'];
      const locationId = existingHistory['保管場所ID'];
      const movementType = existingHistory['入出庫タイプ'];

      // 在庫を逆操作で戻す
      const revertType = movementType === '入庫' ? '出庫' : '入庫';
      updateStock(productId, locationId, quantity, revertType);

      // 履歴を削除
      historySheet.deleteRow(existingHistory._rowIndex);

      return { success: true, message: '入出庫履歴を削除しました（在庫を元に戻しました）' };

    } catch (error) {
      Logger.log('deleteStockMovement Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

/**
 * 入出庫履歴の製品を変更する（在庫も連動して再計算）
 * @param {string} historyId - 履歴ID
 * @param {string} newProductId - 新しい製品ID
 * @returns {Object} 結果オブジェクト
 */
function changeProductInHistory(historyId, newProductId) {
  return loggable('changeProductInHistory', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      if (!historyId || !newProductId) {
        throw new Error('履歴IDと新しい製品IDは必須です');
      }

      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_入出庫履歴');
      const productSheet = ss.getSheetByName('M_製品');

      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。');
      }
      if (!productSheet) {
        throw new Error('M_製品シートが見つかりません。');
      }

      // 既存の履歴を取得
      const existingHistory = findRowByColumn(historySheet, '履歴ID', historyId);
      if (!existingHistory) {
        throw new Error('指定された履歴が見つかりません');
      }

      // 自分の履歴かチェック（管理者は全員分編集可能）
      if (user.role !== '管理者' && existingHistory['操作ユーザーID'] !== user.userId) {
        throw new Error('他のユーザーの履歴は編集できません');
      }

      // 新しい製品が存在するか確認
      const newProduct = findRowByColumn(productSheet, '製品ID', newProductId);
      if (!newProduct) {
        throw new Error('指定された製品が見つかりません');
      }
      if (newProduct['有効'] !== true) {
        throw new Error('指定された製品は無効です');
      }

      const oldProductId = existingHistory['製品ID'];
      const locationId = existingHistory['保管場所ID'];
      const quantity = Number(existingHistory['数量']);
      const movementType = existingHistory['入出庫タイプ'];

      // 同じ製品の場合は何もしない
      if (String(oldProductId) === String(newProductId)) {
        return { success: true, message: '製品は変更されませんでした（同じ製品です）' };
      }

      // 在庫の再計算
      // 1. 旧製品の在庫を戻す（入庫なら減算、出庫なら加算）
      const revertType = movementType === '入庫' ? '出庫' : '入庫';
      updateStock(oldProductId, locationId, quantity, revertType);

      // 2. 新製品の在庫を更新（元の入出庫タイプで適用）
      // 出庫の場合は在庫チェック
      if (movementType === '出庫') {
        const newProductStock = getCurrentStock(newProductId, locationId);
        if (newProductStock < quantity) {
          // 在庫不足の場合は旧製品の在庫を元に戻す
          updateStock(oldProductId, locationId, quantity, movementType);
          throw new Error(`新製品の在庫不足: 現在在庫数${newProductStock}に対して${quantity}の出庫はできません`);
        }
      }
      updateStock(newProductId, locationId, quantity, movementType);

      // 履歴の製品IDを更新
      const updatedHistory = {
        '履歴ID': historyId,
        '製品ID': newProductId,
        '数量': quantity,
        '入出庫タイプ': movementType,
        '現場名': existingHistory['現場名'],
        '客先': existingHistory['客先'],
        '発生日時': existingHistory['発生日時'],
        '操作ユーザーID': existingHistory['操作ユーザーID'],
        '保管場所ID': locationId
      };
      updateSheetRow(historySheet, existingHistory._rowIndex, updatedHistory);

      return {
        success: true,
        message: '製品を変更しました（在庫も再計算されました）',
        oldProductId: oldProductId,
        newProductId: newProductId
      };

    } catch (error) {
      Logger.log('changeProductInHistory Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

// ========================================
// 棚卸管理
// ========================================

function createInventoryEvent(eventName) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('createInventoryEvent', arguments, function() {
    try {
      requireAdminPermission();
      if (!eventName || eventName.trim() === '') throw new Error('イベント名を入力してください');
      
      const user = getCurrentUser();
      const ss = getSpreadsheet();
      const historySheet = ss.getSheetByName('T_棚卸履歴');
      
      // 【修正】nullチェック追加
      if (!historySheet) {
        throw new Error('T_棚卸履歴シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const inventoryId = generateUniqueId('I');
      const historyData = {
        '棚卸ID': inventoryId,
        'イベント名': eventName,
        '棚卸実施日': getCurrentDateTime(),
        'ステータス': '実施中',
        '担当ユーザーID': user.userId
      };
      appendRowToSheet(historySheet, historyData);
      
      createInventorySnapshot(inventoryId);
      
      return { success: true, message: '棚卸イベントを作成しました', inventoryId: inventoryId };
    } catch (error) {
      Logger.log('createInventoryEvent Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function createInventorySnapshot(inventoryId) {
  const ss = getSpreadsheet();
  const stockSheet = ss.getSheetByName('T_在庫');
  const detailSheet = ss.getSheetByName('T_棚卸明細');
  
  // 【修正】nullチェック追加
  if (!stockSheet) {
    throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
  }
  if (!detailSheet) {
    throw new Error('T_棚卸明細シートが見つかりません。スプレッドシートの設定を確認してください。');
  }
  
  const stocks = getSheetData(stockSheet);
  
  stocks.forEach(stock => {
    const detailId = generateUniqueId('ID');
    const detailData = {
      '棚卸明細ID': detailId,
      '棚卸ID': inventoryId,
      '製品ID': stock['製品ID'],
      '保管場所ID': stock['保管場所ID'],
      '理論在庫数': stock['現在在庫数'],
      '確定実在庫数': '',
      '差異': '',
      '差異理由': ''
    };
    appendRowToSheet(detailSheet, detailData);
  });
}

function getActiveInventoryEvents() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getActiveInventoryEvents', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('T_棚卸履歴');
      
      // 【修正】nullチェック追加 - これが今回のエラーの主な原因
      if (!sheet) {
        throw new Error('T_棚卸履歴シートが見つかりません。スプレッドシートに「T_棚卸履歴」という名前のシートを作成してください。');
      }
      
      const data = getSheetData(sheet);
      return data
        .filter(row => row['ステータス'] === '実施中' || row['ステータス'] === '照合中')
        .map(row => ({
          棚卸ID: String(row['棚卸ID'] || ''),
          イベント名: String(row['イベント名'] || ''),
          棚卸実施日: row['棚卸実施日'] ? formatDateTime(row['棚卸実施日']) : '',
          ステータス: String(row['ステータス'] || ''),
          担当ユーザーID: String(row['担当ユーザーID'] || '')
        }));
    } catch (error) {
      Logger.log('getActiveInventoryEvents Error: ' + error.toString());
      throw new Error('棚卸イベントの取得に失敗しました: ' + error.message);
    }
  });
}

function registerInventoryCount(data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('registerInventoryCount', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');
      
      const validation = validateRequiredFields(data, ['inventoryId', 'locationId', 'productId', 'count']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));
      
      if (!validateNumber(data.count, 0)) {
        throw new Error('カウント数は0以上の数値を入力してください');
      }
      
      const ss = getSpreadsheet();
      const inputSheet = ss.getSheetByName('T_棚卸担当者別入力');
      
      // 【修正】nullチェック追加
      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const inputId = generateUniqueId('IC');
      const inputData = {
        '入力ID': inputId,
        '棚卸ID': data.inventoryId,
        '担当ユーザーID': user.userId,
        '製品ID': data.productId,
        '保管場所ID': data.locationId,
        'カウント数': Number(data.count),
        '入力日時': getCurrentDateTime()
      };
      appendRowToSheet(inputSheet, inputData);
      
      return { success: true, message: 'カウントを登録しました', inputId: inputId };
    } catch (error) {
      Logger.log('registerInventoryCount Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function registerInventoryCountBatch(dataList) {
  // ★★★ ロギングラッパーで囲む (新規) ★★★
  return loggable('registerInventoryCountBatch', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');

      if (!dataList || !Array.isArray(dataList) || dataList.length === 0) {
        throw new Error('登録データがありません');
      }

      const ss = getSpreadsheet();
      const inputSheet = ss.getSheetByName('T_棚卸担当者別入力');

      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      let successCount = 0;
      let errorCount = 0;
      const errors = [];

      dataList.forEach((data, index) => {
        try {
          // 必須項目チェック（countは0を許容するため除外）
          const validation = validateRequiredFields(data, ['inventoryId', 'locationId', 'productId']);
          if (!validation.valid) {
            throw new Error(`行${index + 1}: ${validation.errors.join(', ')}`);
          }

          // カウント数の存在チェック（0も許容）
          if (data.count === undefined || data.count === null || data.count === '') {
            throw new Error(`行${index + 1}: カウント数は必須項目です`);
          }

          // 数値チェック（0以上）
          if (!validateNumber(data.count, 0)) {
            throw new Error(`行${index + 1}: カウント数は0以上の数値を入力してください`);
          }

          // データ登録
          const inputId = generateUniqueId('IC');
          const inputData = {
            '入力ID': inputId,
            '棚卸ID': data.inventoryId,
            '担当ユーザーID': user.userId,
            '製品ID': data.productId,
            '保管場所ID': data.locationId,
            'カウント数': Number(data.count),
            '入力日時': getCurrentDateTime()
          };
          appendRowToSheet(inputSheet, inputData);
          successCount++;

        } catch (error) {
          errorCount++;
          errors.push(error.toString());
          Logger.log('registerInventoryCountBatch - Item Error: ' + error.toString());
        }
      });

      if (errorCount > 0) {
        return {
          success: false,
          error: `${successCount}件登録成功、${errorCount}件失敗\n${errors.slice(0, 5).join('\n')}`,
          successCount: successCount,
          errorCount: errorCount
        };
      }

      return {
        success: true,
        message: `${successCount}件のカウントを一括登録しました`,
        successCount: successCount
      };

    } catch (error) {
      Logger.log('registerInventoryCountBatch Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function getInventoryCountsByEvent(inventoryId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getInventoryCountsByEvent', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const inputSheet = ss.getSheetByName('T_棚卸担当者別入力');
      
      // 【修正】nullチェック追加
      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      return findAllRowsByColumn(inputSheet, '棚卸ID', inventoryId);
    } catch (error) {
      Logger.log('getInventoryCountsByEvent Error: ' + error.toString());
      throw new Error('カウントデータの取得に失敗しました: ' + error.message);
    }
  });
}

function getMyInventoryCounts(inventoryId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getMyInventoryCounts', arguments, function() {
    try {
      const user = getCurrentUser();
      if (!user || !user.valid) throw new Error('ログインが必要です');
      
      const ss = getSpreadsheet();
      const inputSheet = ss.getSheetByName('T_棚卸担当者別入力');
      
      // 【修正】nullチェック追加
      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const allCounts = findAllRowsByColumn(inputSheet, '棚卸ID', inventoryId);
      return allCounts.filter(count => count['担当ユーザーID'] === user.userId);
    } catch (error) {
      Logger.log('getMyInventoryCounts Error: ' + error.toString());
      throw new Error('カウントデータの取得に失敗しました: ' + error.message);
    }
  });
}

function verifyInventoryCounts(inventoryId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('verifyInventoryCounts', arguments, function() {
    try {
      requireAdminPermission();

      const ss = getSpreadsheet();
      const inputSheet = ss.getSheetByName('T_棚卸担当者別入力');
      const detailSheet = ss.getSheetByName('T_棚卸明細');
      const productSheet = ss.getSheetByName('M_製品');
      const locationSheet = ss.getSheetByName('M_保管場所');
      const userSheet = ss.getSheetByName('M_ユーザー');

      // 【修正】nullチェック追加
      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!detailSheet) {
        throw new Error('T_棚卸明細シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!productSheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!locationSheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!userSheet) {
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      const counts = findAllRowsByColumn(inputSheet, '棚卸ID', inventoryId);
      const details = findAllRowsByColumn(detailSheet, '棚卸ID', inventoryId);
      const products = getSheetData(productSheet);
      const locations = getSheetData(locationSheet);
      const users = getSheetData(userSheet);

      const grouped = {};
      counts.forEach(count => {
        const key = count['製品ID'] + '_' + count['保管場所ID'];
        if (!grouped[key]) {
          grouped[key] = {
            productId: count['製品ID'],
            locationId: count['保管場所ID'],
            counts: []
          };
        }
        const user = users.find(u => u['ユーザーID'] === count['担当ユーザーID']);
        grouped[key].counts.push({
          userId: count['担当ユーザーID'],
          userName: user ? user['ユーザー名'] : count['担当ユーザーID'],
          count: count['カウント数']
        });
      });

      const discrepancies = [];
      Object.keys(grouped).forEach(key => {
        const item = grouped[key];
        if (item.counts.length > 1) {
          const firstCount = item.counts[0].count;
          // 【修正】型変換を明示的に行い、数値として比較
          const hasDiscrepancy = item.counts.some(c => Number(c.count) !== Number(firstCount));

          if (hasDiscrepancy) {
            const detail = details.find(d =>
              d['製品ID'] === item.productId && d['保管場所ID'] === item.locationId
            );
            const product = products.find(p => String(p['製品ID']) === String(item.productId));
            const location = locations.find(l => String(l['保管場所ID']) === String(item.locationId));

            discrepancies.push({
              productId: item.productId,
              productName: product ? product['製品名'] : item.productId,
              locationId: item.locationId,
              locationName: location ? location['場所名'] : item.locationId,
              theoreticalStock: detail ? detail['理論在庫数'] : 0,
              counts: item.counts
            });
          }
        }
      });

      // 一致している場合も照合結果として返す
      const allResults = [];
      Object.keys(grouped).forEach(key => {
        const item = grouped[key];
        const detail = details.find(d =>
          d['製品ID'] === item.productId && d['保管場所ID'] === item.locationId
        );
        const product = products.find(p => String(p['製品ID']) === String(item.productId));
        const location = locations.find(l => String(l['保管場所ID']) === String(item.locationId));

        // 平均カウント数を計算
        const avgCount = item.counts.reduce((sum, c) => sum + Number(c.count), 0) / item.counts.length;

        allResults.push({
          productId: item.productId,
          productName: product ? product['製品名'] : item.productId,
          locationId: item.locationId,
          locationName: location ? location['場所名'] : item.locationId,
          theoreticalStock: detail ? detail['理論在庫数'] : 0,
          confirmedCount: Math.round(avgCount), // 平均値を四捨五入
          counts: item.counts
        });
      });

      return {
        success: true,
        totalItems: Object.keys(grouped).length,
        discrepancies: discrepancies,
        hasDiscrepancies: discrepancies.length > 0,
        allResults: allResults // 全ての照合結果
      };
    } catch (error) {
      Logger.log('verifyInventoryCounts Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function getInventoryDetails(inventoryId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getInventoryDetails', arguments, function() {
    try {
      requireAdminPermission();
      
      const ss = getSpreadsheet();
      const detailSheet = ss.getSheetByName('T_棚卸明細');
      const productSheet = ss.getSheetByName('M_製品');
      const locationSheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!detailSheet) {
        throw new Error('T_棚卸明細シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!productSheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!locationSheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const details = findAllRowsByColumn(detailSheet, '棚卸ID', inventoryId);
      const products = getSheetData(productSheet);
      const locations = getSheetData(locationSheet);
      
      return details.map(detail => {
        const product = products.find(p => p['製品ID'] === detail['製品ID']);
        const location = locations.find(l => l['保管場所ID'] === detail['保管場所ID']);
        return {
          ...detail,
          製品名: product ? product['製品名'] : '',
          場所名: location ? location['場所名'] : ''
        };
      });
    } catch (error) {
      Logger.log('getInventoryDetails Error: ' + error.toString());
      throw new Error('棚卸明細の取得に失敗しました: ' + error.message);
    }
  });
}

function finalizeInventory(inventoryId, details) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('finalizeInventory', arguments, function() {
    try {
      requireAdminPermission();
      if (!details || details.length === 0) throw new Error('確定する明細データがありません');
      
      const ss = getSpreadsheet();
      const detailSheet = ss.getSheetByName('T_棚卸明細');
      const stockSheet = ss.getSheetByName('T_在庫');
      const historySheet = ss.getSheetByName('T_入出庫履歴');
      const inventoryHistorySheet = ss.getSheetByName('T_棚卸履歴');
      
      // 【修正】nullチェック追加
      if (!detailSheet) {
        throw new Error('T_棚卸明細シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!stockSheet) {
        throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!historySheet) {
        throw new Error('T_入出庫履歴シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!inventoryHistorySheet) {
        throw new Error('T_棚卸履歴シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const user = getCurrentUser();
      
      let updatedStocks = 0;
      let adjustmentHistories = 0;
      
      const detailData = detailSheet.getDataRange().getValues();
      const detailHeaders = detailData[0];
      const stockData = stockSheet.getDataRange().getValues();
      const stockHeaders = stockData[0];
      
      details.forEach(detail => {
        const confirmedCount = Number(detail.confirmedCount);
        const theoreticalStock = Number(detail.theoreticalStock || 0);
        const discrepancy = confirmedCount - theoreticalStock;
        
        // 棚卸明細を更新
        for (let i = 1; i < detailData.length; i++) {
          if (detailData[i][detailHeaders.indexOf('棚卸明細ID')] === detail.detailId) {
            detailSheet.getRange(i + 1, detailHeaders.indexOf('確定実在庫数') + 1).setValue(confirmedCount);
            detailSheet.getRange(i + 1, detailHeaders.indexOf('差異') + 1).setValue(discrepancy);
            detailSheet.getRange(i + 1, detailHeaders.indexOf('差異理由') + 1).setValue(detail.discrepancyReason || '');
            break;
          }
        }
        
        // T_在庫を更新
        for (let i = 1; i < stockData.length; i++) {
          if (stockData[i][stockHeaders.indexOf('製品ID')] === detail.productId && 
              stockData[i][stockHeaders.indexOf('保管場所ID')] === detail.locationId) {
            stockSheet.getRange(i + 1, stockHeaders.indexOf('現在在庫数') + 1).setValue(confirmedCount);
            stockSheet.getRange(i + 1, stockHeaders.indexOf('最終更新日時') + 1).setValue(getCurrentDateTime());
            updatedStocks++;
            break;
          }
        }
        
        // 差異がある場合、調整ログを追加
        if (discrepancy !== 0) {
          const historyId = generateUniqueId('H');
          const adjustmentType = discrepancy > 0 ? '棚卸調整(入庫)' : '棚卸調整(出庫)';
          const historyData = {
            '履歴ID': historyId,
            '製品ID': detail.productId,
            '数量': Math.abs(discrepancy),
            '入出庫タイプ': adjustmentType,
            '現場名': '棚卸調整: ' + inventoryId,
            '発生日時': getCurrentDateTime(),
            '操作ユーザーID': user.userId,
            '保管場所ID': detail.locationId
          };
          appendRowToSheet(historySheet, historyData);
          adjustmentHistories++;
        }
      });
      
      // ステータスを「確定済」に更新
      const inventoryData = inventoryHistorySheet.getDataRange().getValues();
      const inventoryHeaders = inventoryData[0];
      for (let i = 1; i < inventoryData.length; i++) {
        if (inventoryData[i][inventoryHeaders.indexOf('棚卸ID')] === inventoryId) {
          inventoryHistorySheet.getRange(i + 1, inventoryHeaders.indexOf('ステータス') + 1).setValue('確定済');
          break;
        }
      }
      
      return {
        success: true,
        message: '棚卸を確定しました',
        updatedStocks: updatedStocks,
        adjustmentHistories: adjustmentHistories
      };
    } catch (error) {
      Logger.log('finalizeInventory Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

// ========================================
// 在庫照会
// ========================================

function getAllStocks() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getAllStocks', arguments, function() {
    try {
      const ss = getSpreadsheet();
      const stockSheet = ss.getSheetByName('T_在庫');
      const productSheet = ss.getSheetByName('M_製品');
      const locationSheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!stockSheet) {
        throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!productSheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!locationSheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const stocks = getSheetData(stockSheet);
      const products = getSheetData(productSheet);
      const locations = getSheetData(locationSheet);

      // デバッグログ追加
      Logger.log('getAllStocks: stocks count = ' + stocks.length);
      Logger.log('getAllStocks: products count = ' + products.length);
      Logger.log('getAllStocks: locations count = ' + locations.length);

      const result = stocks.map(stock => {
        // 型を文字列に統一して比較
        const stockProductId = String(stock['製品ID']);
        const stockLocationId = String(stock['保管場所ID']);

        const product = products.find(p => String(p['製品ID']) === stockProductId);
        const location = locations.find(l => String(l['保管場所ID']) === stockLocationId);

        // デバッグ: マッチングできなかった場合のログ
        if (!product) {
          Logger.log('getAllStocks: 製品が見つかりません - 製品ID: ' + stockProductId);
        }
        if (!location) {
          Logger.log('getAllStocks: 保管場所が見つかりません - 保管場所ID: ' + stockLocationId);
        }

        return {
          在庫ID: String(stock['在庫ID'] || ''),
          製品ID: String(stock['製品ID'] || ''),
          製品名: product ? String(product['製品名'] || '') : '',
          カテゴリ1: product ? String(product['カテゴリ1'] || '') : '',
          カテゴリ2: product ? String(product['カテゴリ2'] || '') : '',
          保管場所ID: String(stock['保管場所ID'] || ''),
          場所名: location ? String(location['場所名'] || '') : '',
          現在在庫数: Number(stock['現在在庫数'] || 0),
          適正在庫数: product ? Number(product['適正在庫数'] || 0) : 0,
          担当部署: product ? String(product['担当部署'] || '') : '',
          最終更新日時: stock['最終更新日時'] ? formatDateTime(stock['最終更新日時']) : ''
        };
      });

      Logger.log('getAllStocks: result count = ' + result.length);
      return result;
    } catch (error) {
      Logger.log('getAllStocks Error: ' + error.toString());
      throw new Error('在庫データの取得に失敗しました: ' + error.message);
    }
  });
}

// ★★★ 検索処理 (修正) ★★★
function searchStocks(query) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('searchStocks', arguments, function() {
    try {
      const allStocks = getAllStocks(); // 既存の関数で全在庫と製品情報を取得

      // queryがnull、またはすべての検索条件が空の場合は全件返す
      // (修正) productIdのチェックを削除
      if (!query || (!query.productName && !query.category1 && !query.category2)) {
        return allStocks;
      }
      
      return allStocks.filter(stock => {
        let match = true;

        // (削除) productIdのフィルター処理を削除
        // if (query.productId) {
        //   match = match && stock['製品ID'].indexOf(query.productId) !== -1;
        // }

        if (query.productName) {
          // 製品名（部分一致）
          const productName = stock['製品名'] || '';
          match = match && productName.indexOf(query.productName) !== -1;
        }
        if (query.category1) {
          // カテゴリ1（完全一致）
          match = match && (stock['カテゴリ1'] || '') === query.category1;
        }
        if (query.category2) {
          // カテゴリ2（完全一致）
          match = match && (stock['カテゴリ2'] || '') === query.category2;
        }

        return match;
      });
    } catch (error) {
      Logger.log('searchStocks Error: ' + error.toString());
      throw new Error('在庫検索に失敗しました: ' + error.message);
    }
  });
}

function getStocksByProduct() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getStocksByProduct', arguments, function() {
    try {
      const allStocks = getAllStocks();
      const grouped = {};
      
      allStocks.forEach(stock => {
        const key = stock['製品ID'];
        if (!grouped[key]) {
          grouped[key] = {
            製品ID: stock['製品ID'],
            製品名: stock['製品名'],
            カテゴリ1: stock['カテゴリ1'],
            カテゴリ2: stock['カテゴリ2'],
            総在庫数: 0,
            保管場所: []
          };
        }
        grouped[key].総在庫数 += Number(stock['現在在庫数']);
        grouped[key].保管場所.push({
          場所名: stock['場所名'],
          在庫数: stock['現在在庫数']
        });
      });
      return Object.values(grouped);
    } catch (error) {
      Logger.log('getStocksByProduct Error: ' + error.toString());
      throw new Error('製品別集計に失敗しました: ' + error.message);
    }
  });
}

function getStocksByLocation() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getStocksByLocation', arguments, function() {
    try {
      const allStocks = getAllStocks();
      const grouped = {};
      
      allStocks.forEach(stock => {
        const key = stock['保管場所ID'];
        if (!grouped[key]) {
          grouped[key] = {
            保管場所ID: stock['保管場所ID'],
            場所名: stock['場所名'],
            製品数: 0,
            製品リスト: []
          };
        }
        grouped[key].製品数++;
        grouped[key].製品リスト.push({
          製品ID: stock['製品ID'],
          製品名: stock['製品名'],
          在庫数: stock['現在在庫数']
        });
      });
      return Object.values(grouped);
    } catch (error) {
      Logger.log('getStocksByLocation Error: ' + error.toString());
      throw new Error('保管場所別集計に失敗しました: ' + error.message);
    }
  });
}

// ========================================
// マスタ管理
// ========================================

function getAllProducts() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getAllProducts', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      return getSheetData(sheet);
    } catch (error) {
      Logger.log('getAllProducts Error: ' + error.toString());
      throw new Error('製品データの取得に失敗しました: ' + error.message);
    }
  });
}

function createProduct(data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('createProduct', arguments, function() {
    try {
      requireAdminPermission();
      const validation = validateRequiredFields(data, ['製品名', 'カテゴリ1', 'カテゴリ2']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');

      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      // 【追加】製品IDを自動発行（既存の最大ID+1）
      const allData = getSheetData(sheet);
      let maxId = 0;
      allData.forEach(row => {
        const id = parseInt(String(row['製品ID'] || '0'));
        if (!isNaN(id) && id > maxId) {
          maxId = id;
        }
      });
      const newProductId = String(maxId + 1);

      const productData = {
        '製品ID': newProductId,
        '製品名': data['製品名'],
        'カテゴリ1': data['カテゴリ1'],
        'カテゴリ2': data['カテゴリ2'],
        '有効': data['有効'] !== undefined ? data['有効'] : true,
        '単価': data['単価'] || 0,
        '適正在庫数': data['適正在庫数'] || 0,
        '担当部署': data['担当部署'] || ''
      };
      appendRowToSheet(sheet, productData);
      return { success: true, message: '製品を登録しました（ID: ' + newProductId + '）' };
    } catch (error) {
      Logger.log('createProduct Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function updateProduct(productId, data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('updateProduct', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, '製品ID', productId);
      if (!existing) throw new Error('製品ID「' + productId + '」が見つかりません');
      
      const updateData = {
        '製品ID': productId,
        '製品名': data['製品名'] || existing['製品名'],
        'カテゴリ1': data['カテゴリ1'] || existing['カテゴリ1'],
        'カテゴリ2': data['カテゴリ2'] || existing['カテゴリ2'],
        '有効': data['有効'] !== undefined ? data['有効'] : existing['有効'],
        '単価': data['単価'] !== undefined ? data['単価'] : existing['単価'],
        '適正在庫数': data['適正在庫数'] !== undefined ? data['適正在庫数'] : (existing['適正在庫数'] || 0),
        '担当部署': data['担当部署'] !== undefined ? data['担当部署'] : (existing['担当部署'] || '')
      };
      updateSheetRow(sheet, existing._rowIndex, updateData);
      return { success: true, message: '製品を更新しました' };
    } catch (error) {
      Logger.log('updateProduct Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function deleteProduct(productId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('deleteProduct', arguments, function() {
    try {
      requireAdminPermission();
      return updateProduct(productId, { '有効': false });
    } catch (error) {
      Logger.log('deleteProduct Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function getAllUsers() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getAllUsers', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_ユーザー');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      return getSheetData(sheet);
    } catch (error) {
      Logger.log('getAllUsers Error: ' + error.toString());
      throw new Error('ユーザーデータの取得に失敗しました: ' + error.message);
    }
  });
}

function createUser(data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('createUser', arguments, function() {
    try {
      requireAdminPermission();
      const validation = validateRequiredFields(data, ['ユーザー名', 'メールアドレス', '権限']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_ユーザー');

      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      const existingEmail = findRowByColumn(sheet, 'メールアドレス', data['メールアドレス']);
      if (existingEmail) throw new Error('メールアドレス「' + data['メールアドレス'] + '」は既に登録されています');

      // 【追加】ユーザーIDを自動発行（U001形式）
      const allData = getSheetData(sheet);
      let maxNum = 0;
      allData.forEach(row => {
        const id = String(row['ユーザーID'] || '');
        const match = id.match(/^U(\d+)$/);
        if (match) {
          const num = parseInt(match[1]);
          if (!isNaN(num) && num > maxNum) {
            maxNum = num;
          }
        }
      });
      const newUserId = 'U' + String(maxNum + 1).padStart(3, '0');

      const userData = {
        'ユーザーID': newUserId,
        '部門': data['部門'] || '',
        'ユーザー名': data['ユーザー名'],
        'メールアドレス': data['メールアドレス'],
        '権限': data['権限'],
        '有効': data['有効'] !== undefined ? data['有効'] : true
      };
      appendRowToSheet(sheet, userData);
      return { success: true, message: 'ユーザーを登録しました（ID: ' + newUserId + '）' };
    } catch (error) {
      Logger.log('createUser Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function updateUser(userId, data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('updateUser', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_ユーザー');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, 'ユーザーID', userId);
      if (!existing) throw new Error('ユーザーID「' + userId + '」が見つかりません');
      
      const updateData = {
        'ユーザーID': userId,
        '部門': data['部門'] !== undefined ? data['部門'] : existing['部門'],
        'ユーザー名': data['ユーザー名'] || existing['ユーザー名'],
        'メールアドレス': data['メールアドレス'] || existing['メールアドレス'],
        '権限': data['権限'] || existing['権限'],
        '有効': data['有効'] !== undefined ? data['有効'] : existing['有効']
      };
      updateSheetRow(sheet, existing._rowIndex, updateData);
      return { success: true, message: 'ユーザーを更新しました' };
    } catch (error) {
      Logger.log('updateUser Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function deleteUser(userId) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('deleteUser', arguments, function() {
    try {
      requireAdminPermission();
      return updateUser(userId, { '有効': false });
    } catch (error) {
      Logger.log('deleteUser Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function getAllLocations() {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('getAllLocations', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      return getSheetData(sheet);
    } catch (error) {
      Logger.log('getAllLocations Error: ' + error.toString());
      throw new Error('保管場所データの取得に失敗しました: ' + error.message);
    }
  });
}

function createLocation(data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('createLocation', arguments, function() {
    try {
      requireAdminPermission();
      const validation = validateRequiredFields(data, ['場所名']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));

      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');

      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }

      // 【追加】保管場所IDを自動発行（P001形式）
      const allData = getSheetData(sheet);
      let maxNum = 0;
      allData.forEach(row => {
        const id = String(row['保管場所ID'] || '');
        const match = id.match(/^P(\d+)$/);
        if (match) {
          const num = parseInt(match[1]);
          if (!isNaN(num) && num > maxNum) {
            maxNum = num;
          }
        }
      });
      const newLocationId = 'P' + String(maxNum + 1).padStart(3, '0');

      const locationData = {
        '保管場所ID': newLocationId,
        '場所名': data['場所名']
      };
      appendRowToSheet(sheet, locationData);
      return { success: true, message: '保管場所を登録しました（ID: ' + newLocationId + '）' };
    } catch (error) {
      Logger.log('createLocation Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function updateLocation(locationId, data) {
  // ★★★ ロギングラッパーで囲む (修正) ★★★
  return loggable('updateLocation', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, '保管場所ID', locationId);
      if (!existing) throw new Error('保管場所ID「' + locationId + '」が見つかりません');
      
      const updateData = {
        '保管場所ID': locationId,
        '場所名': data['場所名'] || existing['場所名']
      };
      updateSheetRow(sheet, existing._rowIndex, updateData);
      return { success: true, message: '保管場所を更新しました' };
    } catch (error) {
      Logger.log('updateLocation Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

function deleteLocation(locationId) {
  // ★★★ ロギ級ラッパーで囲む (修正) ★★★
  return loggable('deleteLocation', arguments, function() {
    try {
      requireAdminPermission();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, '保管場所ID', locationId);
      if (!existing) throw new Error('保管場所ID「' + locationId + '」が見つかりません');
      
      const stockSheet = ss.getSheetByName('T_在庫');
      if (!stockSheet) {
        throw new Error('T_在庫シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const stockData = findAllRowsByColumn(stockSheet, '保管場所ID', locationId);
      if (stockData.length > 0) {
        throw new Error('この保管場所は在庫データで使用中のため削除できません');
      }
      sheet.deleteRow(existing._rowIndex);
      return { success: true, message: '保管場所を削除しました' };
    } catch (error) {
      Logger.log('deleteLocation Error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
}

// ★★★ ここから下の重複していたブロックを削除 ★★★