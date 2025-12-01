[不動産価格監視システム]
 * 
 * 【課題】
 * - 不動産価格の動向に関心があったが、手作業で集計しなくてならず時間がかかる
 * - 収集の時間が取れない
 * 
 * 【解決策】
 * - このスクリプトで自動化
 * - AI（claude）を活用して開発
 * 
 * 【成果】
 * - 作業時間：指示を出すと自動的にデータを収集する
 * - 大幅な時間の節約
 * 
 * 【使用技術】
 * - Google Apps Script
 * - Googleスプレッドシート
 * - Gmail API

// ==============================================
// 沖縄不動産取引データ監視システム
// 不動産情報ライブラリAPI対応版
// ==============================================

const CONFIG = {
  emailRecipient: 'samon.miyagi@gmail.com',
  reinfolibApiKey: '3e0be87849df4e8ea44e802b6a6662ab',
  checkIntervalHours: 24,
  prefectureCode: '47',
  targetYears: [2024],
  maxRecordsPerQuarter: 50
};

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('沖縄不動産取引データ監視システム');
  
  let conditionsSheet = ss.getSheetByName('監視条件');
  if (!conditionsSheet) {
    conditionsSheet = ss.insertSheet('監視条件', 0);
    
    const headers = [
      '有効/無効', '条件名', '取引種別', '最低価格(万円)', '最高価格(万円)',
      '最低土地面積(㎡)', '最高土地面積(㎡)', '最低建物面積(㎡)', '最高建物面積(㎡)',
      '市区町村', '建築年(以降)', '駅距離(分以内)', '用途', '備考'
    ];
    conditionsSheet.appendRow(headers);
    
    const headerRange = conditionsSheet.getRange('A1:N1');
    headerRange.setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff').setHorizontalAlignment('center');
    
    const samples = [
      ['ON', '那覇市の土地', '宅地(土地)', 1000, 3000, 100, '', '', '', '那覇市', '', '', '住宅', ''],
      ['ON', '浦添市マンション', 'マンション', 2000, 5000, '', '', 50, '', '浦添市', 2000, 15, '', '駅近希望'],
      ['OFF', '県内土地建物', '宅地(土地と建物)', 1000, 5000, '', '', '', '', '', 2010, '', '住宅', '全域'],
      ['ON', '沖縄市の土地', '宅地(土地)', 500, 2000, 200, '', '', '', '沖縄市', '', '', '', '広め'],
      ['OFF', '農地', '農地', '', 1000, 500, '', '', '', '', '', '', '', '将来用']
    ];
    
    samples.forEach(function(sample) { conditionsSheet.appendRow(sample); });
    
    for (let i = 1; i <= 14; i++) {
      conditionsSheet.setColumnWidth(i, i === 2 ? 150 : (i === 3 ? 150 : (i === 14 ? 150 : 110)));
    }
    
    const validationRange = conditionsSheet.getRange('A2:A100');
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['ON', 'OFF'], true).build();
    validationRange.setDataValidation(rule);
    
    const range = conditionsSheet.getRange('A2:N100');
    const formatRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="OFF"')
      .setBackground('#f0f0f0')
      .setFontColor('#999999')
      .setRanges([range])
      .build();
    conditionsSheet.setConditionalFormatRules([formatRule]);
  }
  
  let trackedSheet = ss.getSheetByName('取引データ履歴');
  if (!trackedSheet) {
    trackedSheet = ss.insertSheet('取引データ履歴');
    trackedSheet.appendRow([
      'ID', '取引時期', '市区町村', '地区名', '取引種別', '取引価格(万円)', 
      '土地面積(㎡)', '建物面積(㎡)', '建築年', '用途', '最寄駅', '駅距離(分)',
      '合致条件', '登録日時'
    ]);
    const headerRange = trackedSheet.getRange('A1:N1');
    headerRange.setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    trackedSheet.setColumnWidth(1, 250);
    trackedSheet.setColumnWidth(13, 150);
  }
  
  let notificationSheet = ss.getSheetByName('通知履歴');
  if (!notificationSheet) {
    notificationSheet = ss.insertSheet('通知履歴');
    notificationSheet.appendRow([
      '通知日時', '市区町村', '地区名', '取引価格(万円)', '取引種別', '合致条件', 'ステータス'
    ]);
    const headerRange = notificationSheet.getRange('A1:G1');
    headerRange.setFontWeight('bold').setBackground('#fbbc04').setFontColor('#000000');
  }
  
  let errorSheet = ss.getSheetByName('エラーログ');
  if (!errorSheet) {
    errorSheet = ss.insertSheet('エラーログ');
    errorSheet.appendRow(['日時', 'ソース', 'エラー内容']);
    const headerRange = errorSheet.getRange('A1:C1');
    headerRange.setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
    errorSheet.setColumnWidth(3, 400);
  }
  
  let guideSheet = ss.getSheetByName('使い方ガイド');
  if (!guideSheet) {
    guideSheet = ss.insertSheet('使い方ガイド');
    const guide = [
      ['沖縄不動産取引データ監視システム - 使い方ガイド'],
      [''],
      ['このシステムについて'],
      ['国土交通省の不動産情報ライブラリAPIを使用して、沖縄県内の不動産取引データを監視します。'],
      ['新規公開された取引データのうち、設定した条件に合致するものをメールで通知します。'],
      [''],
      ['取得できるデータ'],
      ['・過去の不動産取引価格情報（実際に成約した取引データ）'],
      ['・データは四半期ごとに更新されます'],
      ['※現在募集中の物件情報ではありません'],
      [''],
      ['基本的な使い方'],
      ['1. 監視条件シートで監視したい条件を設定'],
      ['2. 有効/無効列をONにすると監視開始、OFFで監視停止'],
      ['3. 条件に合う取引データが新規公開されるとメールで通知'],
      [''],
      ['初期設定'],
      ['実行 > setupTriggers を実行して自動監視を開始'],
      [''],
      ['テスト実行'],
      ['実行 > testMonitoring でテスト可能']
    ];
    
    guide.forEach(function(row, index) {
      guideSheet.appendRow(row);
      if (index === 0) {
        guideSheet.getRange(index + 1, 1).setFontSize(14).setFontWeight('bold')
          .setBackground('#4285f4').setFontColor('#ffffff');
      }
    });
    
    guideSheet.setColumnWidth(1, 800);
  }
  
  Logger.log('シートの初期化完了');
  return {
    conditionsSheet: conditionsSheet,
    trackedSheet: trackedSheet,
    notificationSheet: notificationSheet,
    errorSheet: errorSheet
  };
}

function loadActiveConditions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('監視条件');
  const data = sheet.getDataRange().getValues();
  
  const conditions = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (row[0] === 'ON') {
      const condition = {
        name: row[1] || '条件' + i,
        tradeType: row[2] ? row[2].trim() : '',
        minPrice: parseFloat(row[3]) || 0,
        maxPrice: parseFloat(row[4]) || 0,
        minLandArea: parseFloat(row[5]) || 0,
        maxLandArea: parseFloat(row[6]) || 0,
        minBuildingArea: parseFloat(row[7]) || 0,
        maxBuildingArea: parseFloat(row[8]) || 0,
        city: row[9] ? row[9].trim() : '',
        buildYear: parseFloat(row[10]) || 0,
        stationDistance: parseFloat(row[11]) || 0,
        usage: row[12] ? row[12].trim() : '',
        note: row[13] || ''
      };
      
      conditions.push(condition);
    }
  }
  
  Logger.log(conditions.length + '件の有効な監視条件を読み込みました');
  return conditions;
}

function monitorProperties() {
  try {
    Logger.log('=== 不動産取引データ監視を開始します ===');
    const sheets = initializeSheets();
    
    const conditions = loadActiveConditions();
    
    if (conditions.length === 0) {
      Logger.log('有効な監視条件がありません。');
      return;
    }
    
    Logger.log('監視条件: ' + conditions.map(function(c) { return c.name; }).join(', '));
    
    const allTransactions = fetchTransactionData();
    
    Logger.log(allTransactions.length + '件の取引データを取得しました');
    
    const matchedTransactions = [];
    
    allTransactions.forEach(function(transaction) {
      if (isAlreadyTracked(transaction.id)) {
        return;
      }
      
      conditions.forEach(function(condition) {
        if (matchesCondition(transaction, condition)) {
          const existing = matchedTransactions.find(function(t) { return t.id === transaction.id; });
          if (existing) {
            existing.matchedConditions.push(condition.name);
          } else {
            transaction.matchedConditions = [condition.name];
            matchedTransactions.push(transaction);
          }
        }
      });
    });
    
    Logger.log(matchedTransactions.length + '件の新規取引データが条件に合致しました');
    
    if (matchedTransactions.length > 0) {
      sendNotificationEmail(matchedTransactions);
      saveTransactions(matchedTransactions, sheets.trackedSheet, sheets.notificationSheet);
      Logger.log(matchedTransactions.length + '件の取引データを通知しました');
    } else {
      Logger.log('条件に合致する新規取引データはありませんでした');
    }
    
    PropertiesService.getScriptProperties().setProperty('lastRun', new Date().toISOString());
    Logger.log('=== 監視完了 ===');
    
  } catch (error) {
    Logger.log('エラー発生: ' + error.toString());
    logError('メイン処理', error.toString());
    sendErrorEmail(error);
  }
}

function fetchTransactionData() {
  const transactions = [];
  
  try {
    Logger.log('不動産情報ライブラリAPIから沖縄県のデータを取得中...');
    
    CONFIG.targetYears.forEach(function(year) {
      for (let quarter = 1; quarter <= 4; quarter++) {
        try {
          const url = 'https://www.reinfolib.mlit.go.jp/ex-api/external/XIT001';
          
          const params = {
            year: year.toString(),
            area: CONFIG.prefectureCode,
            from: year + 'Q' + quarter,
            to: year + 'Q' + quarter
          };
          
          const queryString = Object.keys(params).map(function(key) {
            return key + '=' + encodeURIComponent(params[key]);
          }).join('&');
          
          const fullUrl = url + '?' + queryString;
          
          Logger.log('リクエスト: ' + year + '年第' + quarter + '四半期');
          
          const options = {
            method: 'get',
            headers: {
              'Ocp-Apim-Subscription-Key': CONFIG.reinfolibApiKey
            },
            muteHttpExceptions: true
          };
          
          const response = UrlFetchApp.fetch(fullUrl, options);
          const responseCode = response.getResponseCode();
          
          Logger.log('HTTPステータス: ' + responseCode);
          
          if (responseCode === 200) {
            const responseText = response.getContentText();
            Logger.log('レスポンス受信（最初の200文字）: ' + responseText.substring(0, 200));
            
            const data = JSON.parse(responseText);
            
            if (data && data.data && Array.isArray(data.data)) {
              const limited = data.data.slice(0, CONFIG.maxRecordsPerQuarter);
              Logger.log(year + '年第' + quarter + '四半期: ' + data.data.length + '件中' + limited.length + '件を処理');
              
              limited.forEach(function(item) {
                const transaction = parseTransactionData(item, year, quarter);
                if (transaction) {
                  transactions.push(transaction);
                }
              });
            } else if (data && Array.isArray(data)) {
              const limited = data.slice(0, CONFIG.maxRecordsPerQuarter);
              Logger.log(year + '年第' + quarter + '四半期: ' + data.length + '件中' + limited.length + '件を処理');
              
              limited.forEach(function(item) {
                const transaction = parseTransactionData(item, year, quarter);
                if (transaction) {
                  transactions.push(transaction);
                }
              });
            } else {
              Logger.log('データ形式が想定外: ' + typeof data);
            }
          } else {
            Logger.log('HTTPエラー ' + responseCode + ': ' + response.getContentText().substring(0, 200));
          }
          
          Utilities.sleep(1000);
          
        } catch (e) {
          Logger.log('四半期データ取得エラー: ' + e.toString());
          logError('API取得', year + 'Q' + quarter + ': ' + e.toString());
        }
      }
    });
    
  } catch (error) {
    Logger.log('データ取得エラー: ' + error.toString());
    logError('データ取得', error.toString());
  }
  
  return transactions;
}

function parseTransactionData(item, year, quarter) {
  try {
    const period = year + 'Q' + quarter;
    const municipality = item.cityName || item.Municipality || item.city || '';
    const districtName = item.districtName || item.DistrictName || item.district || '';
    const tradeType = item.tradeType || item.Type || item.type || '';
    const priceValue = item.tradePrice || item.TradePrice || item.price || 0;
    
    const id = [
      period,
      municipality,
      districtName,
      tradeType,
      priceValue
    ].join('_');
    
    const price = typeof priceValue === 'number' ? priceValue / 10000 : 
                  (typeof priceValue === 'string' ? parseFloat(priceValue) / 10000 : 0);
    
    return {
      id: id,
      period: period,
      prefecture: '沖縄県',
      municipality: municipality,
      districtName: districtName,
      type: tradeType,
      price: price,
      landArea: parseFloat(item.landArea || item.Area || 0),
      buildingArea: parseFloat(item.buildingArea || item.BuildingArea || 0),
      buildYear: item.buildingYear ? parseInt(item.buildingYear) : 0,
      usage: item.use || item.Use || '',
      nearestStation: item.nearestStation || item.NearestStation || '',
      stationDistance: parseFloat(item.stationDistance || item.TimeToNearestStation || 0),
      rawData: item
    };
  } catch (e) {
    Logger.log('データ解析エラー: ' + e.toString());
    return null;
  }
}

function matchesCondition(transaction, condition) {
  if (condition.tradeType && transaction.type.indexOf(condition.tradeType) === -1) {
    return false;
  }
  
  if (condition.minPrice > 0 && transaction.price < condition.minPrice) {
    return false;
  }
  if (condition.maxPrice > 0 && transaction.price > condition.maxPrice) {
    return false;
  }
  
  if (condition.minLandArea > 0 && transaction.landArea < condition.minLandArea) {
    return false;
  }
  if (condition.maxLandArea > 0 && transaction.landArea > condition.maxLandArea) {
    return false;
  }
  
  if (condition.minBuildingArea > 0 && transaction.buildingArea < condition.minBuildingArea) {
    return false;
  }
  if (condition.maxBuildingArea > 0 && transaction.buildingArea > condition.maxBuildingArea) {
    return false;
  }
  
  if (condition.city && transaction.municipality.indexOf(condition.city) === -1) {
    return false;
  }
  
  if (condition.buildYear > 0 && transaction.buildYear > 0 && transaction.buildYear < condition.buildYear) {
    return false;
  }
  
  if (condition.stationDistance > 0 && transaction.stationDistance > 0 && 
      transaction.stationDistance > condition.stationDistance) {
    return false;
  }
  
  if (condition.usage && transaction.usage.indexOf(condition.usage) === -1) {
    return false;
  }
  
  return true;
}

function isAlreadyTracked(transactionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('取引データ履歴');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === transactionId) {
      return true;
    }
  }
  return false;
}

function saveTransactions(transactions, trackedSheet, notificationSheet) {
  const now = new Date();
  
  transactions.forEach(function(t) {
    const matchedStr = t.matchedConditions.join(', ');
    
    trackedSheet.appendRow([
      t.id,
      t.period,
      t.municipality,
      t.districtName,
      t.type,
      t.price,
      t.landArea,
      t.buildingArea,
      t.buildYear || '',
      t.usage,
      t.nearestStation,
      t.stationDistance || '',
      matchedStr,
      now
    ]);
    
    notificationSheet.appendRow([
      now,
      t.municipality,
      t.districtName,
      t.price,
      t.type,
      matchedStr,
      '通知済み'
    ]);
  });
}

function sendNotificationEmail(transactions) {
  const subject = '【沖縄不動産】' + transactions.length + '件の新規取引データが見つかりました';
  
  let body = '条件に合致する新規取引データが' + transactions.length + '件見つかりました。\n\n';
  body += '============================================================\n\n';
  
  transactions.forEach(function(t, index) {
    body += '【取引' + (index + 1) + '】\n';
    body += '▼ 合致した条件: ' + t.matchedConditions.join(', ') + '\n\n';
    body += '取引時期: ' + t.period + '\n';
    body += '所在地: ' + t.municipality + ' ' + t.districtName + '\n';
    body += '取引種別: ' + t.type + '\n';
    body += '取引価格: ' + t.price.toLocaleString() + '万円\n';
    if (t.landArea > 0) {
      body += '土地面積: ' + t.landArea + '㎡\n';
    }
    if (t.buildingArea > 0) {
      body += '建物面積: ' + t.buildingArea + '㎡\n';
    }
    if (t.buildYear > 0) {
      body += '建築年: ' + t.buildYear + '年\n';
    }
    if (t.usage) {
      body += '用途: ' + t.usage + '\n';
    }
    if (t.nearestStation) {
      body += '最寄駅: ' + t.nearestStation + '（徒歩' + t.stationDistance + '分）\n';
    }
    body += '\n------------------------------------------------------------\n\n';
  });
  
  body += '\n【システム情報】\n';
  body += '送信日時: ' + new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'}) + '\n';
  body += 'データソース: 国土交通省 不動産情報ライブラリ\n';
  body += '※このメールは自動送信されています\n';
  
  try {
    MailApp.sendEmail({
      to: CONFIG.emailRecipient,
      subject: subject,
      body: body
    });
    Logger.log('通知メール送信完了');
  } catch (error) {
    Logger.log('メール送信エラー: ' + error.toString());
    logError('メール送信', error.toString());
  }
}

function sendErrorEmail(error) {
  const subject = '【エラー】沖縄不動産監視システム';
  const body = 'エラーが発生しました:\n\n' + error.toString() + '\n\n' +
               '発生日時: ' + new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'});
  
  try {
    MailApp.sendEmail(CONFIG.emailRecipient, subject, body);
  } catch (e) {
    Logger.log('エラーメール送信失敗: ' + e.toString());
  }
}

function logError(source, errorMessage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('エラーログ');
    sheet.appendRow([new Date(), source, errorMessage]);
  } catch (e) {
    Logger.log('エラーログ記録失敗: ' + e.toString());
  }
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  
  ScriptApp.newTrigger('monitorProperties').timeBased().everyHours(CONFIG.checkIntervalHours).create();
  
  Logger.log('トリガー設定完了');
  
  const conditions = loadActiveConditions();
  
  MailApp.sendEmail({
    to: CONFIG.emailRecipient,
    subject: '沖縄不動産監視システム - 設定完了',
    body: '監視システムの設定が完了しました。\n\n' +
          '監視間隔: ' + CONFIG.checkIntervalHours + '時間ごと\n' +
          'データソース: 国土交通省 不動産情報ライブラリ\n' +
          '有効な条件: ' + conditions.length + '件'
  });
}

function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log('トリガー削除完了');
  
  MailApp.sendEmail({
    to: CONFIG.emailRecipient,
    subject: '沖縄不動産監視システム - 監視停止',
    body: '監視を一時停止しました。'
  });
}

function testMonitoring() {
  Logger.log('=== テスト実行開始 ===');
  initializeSheets();
  
  const conditions = loadActiveConditions();
  Logger.log('有効な条件数: ' + conditions.length);
  
  monitorProperties();
  Logger.log('=== テスト実行完了 ===');
}
