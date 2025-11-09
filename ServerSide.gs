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
      
      const stock = data.find(row => 
        row['製品ID'] === productId && row['保管場所ID'] === locationId
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
    if (data[i][productIdIndex] === productId && data[i][locationIdIndex] === locationId) {
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
      return data.filter(row => row['ステータス'] === '実施中' || row['ステータス'] === '照合中');
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
      
      // 【修正】nullチェック追加
      if (!inputSheet) {
        throw new Error('T_棚卸担当者別入力シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      if (!detailSheet) {
        throw new Error('T_棚卸明細シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const counts = findAllRowsByColumn(inputSheet, '棚卸ID', inventoryId);
      const details = findAllRowsByColumn(detailSheet, '棚卸ID', inventoryId);
      
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
        grouped[key].counts.push({
          userId: count['担当ユーザーID'],
          count: count['カウント数']
        });
      });
      
      const discrepancies = [];
      Object.keys(grouped).forEach(key => {
        const item = grouped[key];
        if (item.counts.length > 1) {
          const firstCount = item.counts[0].count;
          const hasDiscrepancy = item.counts.some(c => c.count !== firstCount);
          
          if (hasDiscrepancy) {
            const detail = details.find(d => 
              d['製品ID'] === item.productId && d['保管場所ID'] === item.locationId
            );
            discrepancies.push({
              productId: item.productId,
              locationId: item.locationId,
              theoreticalStock: detail ? detail['理論在庫数'] : 0,
              counts: item.counts
            });
          }
        }
      });
      
      return {
        success: true,
        totalItems: Object.keys(grouped).length,
        discrepancies: discrepancies,
        hasDiscrepancies: discrepancies.length > 0
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
      
      return stocks.map(stock => {
        const product = products.find(p => p['製品ID'] === stock['製品ID']);
        const location = locations.find(l => l['保管場所ID'] === stock['保管場所ID']);
        return {
          在庫ID: stock['在庫ID'],
          製品ID: stock['製品ID'],
          製品名: product ? product['製品名'] : '',
          カテゴリ1: product ? product['カテゴリ1'] : '',
          カテゴリ2: product ? product['カテゴリ2'] : '',
          保管場所ID: stock['保管場所ID'],
          場所名: location ? location['場所名'] : '',
          現在在庫数: stock['現在在庫数'],
          最終更新日時: stock['最終更新日時']
        };
      });
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
          match = match && stock['製品名'].indexOf(query.productName) !== -1;
        }
        if (query.category1) {
          // カテゴリ1（完全一致）
          match = match && stock['カテゴリ1'] === query.category1;
        }
        if (query.category2) {
          // カテゴリ2（完全一致）
          match = match && stock['カテゴリ2'] === query.category2;
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
      const validation = validateRequiredFields(data, ['製品ID', '製品名', 'カテゴリ1', 'カテゴリ2']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));
      
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_製品');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_製品シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, '製品ID', data['製品ID']);
      if (existing) throw new Error('製品ID「' + data['製品ID'] + '」は既に存在します');
      
      const productData = {
        '製品ID': data['製品ID'],
        '製品名': data['製品名'],
        'カテゴリ1': data['カテゴリ1'],
        'カテゴリ2': data['カテゴリ2'],
        '有効': data['有効'] !== undefined ? data['有効'] : true,
        '単価': data['単価'] || 0
      };
      appendRowToSheet(sheet, productData);
      return { success: true, message: '製品を登録しました' };
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
        '単価': data['単価'] !== undefined ? data['単価'] : existing['単価']
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
      const validation = validateRequiredFields(data, ['ユーザーID', 'ユーザー名', 'メールアドレス', '権限']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));
      
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_ユーザー');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_ユーザーシートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, 'ユーザーID', data['ユーザーID']);
      if (existing) throw new Error('ユーザーID「' + data['ユーザーID'] + '」は既に存在します');
      
      const existingEmail = findRowByColumn(sheet, 'メールアドレス', data['メールアドレス']);
      if (existingEmail) throw new Error('メールアドレス「' + data['メールアドレス'] + '」は既に登録されています');
      
      const userData = {
        'ユーザーID': data['ユーザーID'],
        '部門': data['部門'] || '',
        'ユーザー名': data['ユーザー名'],
        'メールアドレス': data['メールアドレス'],
        '権限': data['権限'],
        '有効': data['有効'] !== undefined ? data['有効'] : true
      };
      appendRowToSheet(sheet, userData);
      return { success: true, message: 'ユーザーを登録しました' };
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
      const validation = validateRequiredFields(data, ['保管場所ID', '場所名']);
      if (!validation.valid) throw new Error(validation.errors.join(', '));
      
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName('M_保管場所');
      
      // 【修正】nullチェック追加
      if (!sheet) {
        throw new Error('M_保管場所シートが見つかりません。スプレッドシートの設定を確認してください。');
      }
      
      const existing = findRowByColumn(sheet, '保管場所ID', data['保管場所ID']);
      if (existing) throw new Error('保管場所ID「' + data['保管場所ID'] + '」は既に存在します');
      
      const locationData = {
        '保管場所ID': data['保管場所ID'],
        '場所名': data['場所名']
      };
      appendRowToSheet(sheet, locationData);
      return { success: true, message: '保管場所を登録しました' };
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