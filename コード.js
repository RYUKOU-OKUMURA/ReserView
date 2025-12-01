/**
 * ReserView（リザビュー）- 予約集計ビューア
 * Google Apps Script バックエンド
 */

// ★ ここにスプレッドシートIDを設定
var SPREADSHEET_ID = '1z-OuS5riqLp8PYKECOnPzjBWPjgvUa6KKg5c4Ne-g08';
var SHEET_NAME = 'Reservations';

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ReserView - 予約集計ビューア')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * シートのヘッダーインデックスを取得
 */
function getColumnIndexes_(headers) {
  return {
    date: headers.indexOf('日付'),
    start: headers.indexOf('開始'),
    end: headers.indexOf('終了'),
    patient: headers.indexOf('患者名'),
    menu: headers.indexOf('メニュー'),
    amount: headers.indexOf('金額'),
    staff: headers.indexOf('担当'),
    payment: headers.indexOf('決済方法'),
    memo: headers.indexOf('メモ'),
    status: headers.indexOf('ステータス'),
    id: headers.indexOf('ID'),
    lane: headers.indexOf('レーン')
  };
}

/**
 * 金額を数値に変換
 */
function parseAmount_(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    return parseInt(value.replace(/[¥,]/g, '')) || 0;
  }
  return 0;
}

/**
 * 日付をフォーマット
 */
function formatDate_(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(dateValue);
}

/**
 * 時間をフォーマット
 */
function formatTime_(timeValue) {
  if (!timeValue) return '';
  if (timeValue instanceof Date) {
    return Utilities.formatDate(timeValue, 'Asia/Tokyo', 'HH:mm');
  }
  return String(timeValue);
}

/**
 * 日付から日（1-31）を抽出
 */
function extractDay_(dateStr) {
  if (!dateStr) return 0;
  var parts = dateStr.split('/');
  if (parts.length === 3) {
    return parseInt(parts[2]) || 0;
  }
  return 0;
}

/**
 * フィルタ用の選択肢を取得（初期読み込み用）
 */
function getFilterOptions() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    var staffSet = {};
    var menuSet = {};
    var paymentSet = {};
    var statusSet = {};
    var yearMonthSet = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[col.staff]) staffSet[row[col.staff]] = true;
      if (row[col.menu]) menuSet[row[col.menu]] = true;
      if (row[col.payment]) paymentSet[row[col.payment]] = true;
      if (row[col.status]) statusSet[row[col.status]] = true;
      
      var dateStr = formatDate_(row[col.date]);
      if (dateStr) {
        var ym = dateStr.substring(0, 7).replace('/', '-');
        yearMonthSet[ym] = true;
      }
    }
    
    return {
      staff: Object.keys(staffSet).sort(),
      menu: Object.keys(menuSet).sort(),
      payment: Object.keys(paymentSet).sort(),
      status: Object.keys(statusSet).sort(),
      yearMonths: Object.keys(yearMonthSet).sort().reverse()
    };
    
  } catch (error) {
    console.error('getFilterOptions error:', error);
    throw new Error('オプション取得エラー: ' + error.message);
  }
}

/**
 * 予約データを取得
 */
function getReservations(filters) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { reservations: [], summary: createEmptySummary_() };
    }
    
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    var reservations = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[col.id] && !row[col.patient]) continue;
      
      var formattedDate = formatDate_(row[col.date]);
      var formattedStart = formatTime_(row[col.start]);
      var formattedEnd = formatTime_(row[col.end]);
      var amount = parseAmount_(row[col.amount]);
      var yearMonth = formattedDate ? formattedDate.substring(0, 7).replace('/', '-') : '';
      var day = extractDay_(formattedDate);
      
      var record = {
        date: formattedDate,
        yearMonth: yearMonth,
        day: day,
        status: row[col.status] || '',
        staff: row[col.staff] || '',
        payment: row[col.payment] || '',
        menu: row[col.menu] || '',
        patient: row[col.patient] || ''
      };
      
      // フィルタ適用
      if (!matchesFilters_(filters, record)) {
        continue;
      }
      
      var reservation = {
        rowIndex: i + 1,
        date: formattedDate,
        yearMonth: yearMonth,
        day: day,
        start: formattedStart,
        end: formattedEnd,
        patient: row[col.patient] || '',
        menu: row[col.menu] || '',
        amount: amount,
        staff: row[col.staff] || '',
        payment: row[col.payment] || '',
        memo: row[col.memo] || '',
        status: row[col.status] || '',
        id: row[col.id] || '',
        lane: row[col.lane] || ''
      };
      
      reservations.push(reservation);
    }
    
    // ソート
    var sortKey = filters.sortKey || 'date';
    var sortOrder = filters.sortOrder || 'desc';
    reservations = sortReservations_(reservations, sortKey, sortOrder);
    
    // 集計
    var summary = calculateSummary_(reservations);
    
    return { reservations: reservations, summary: summary };
    
  } catch (error) {
    console.error('getReservations error:', error);
    throw new Error('データ取得エラー: ' + error.message);
  }
}

/**
 * フィルタ条件にマッチするかチェック
 */
function matchesFilters_(filters, record) {
  // 月別フィルタ
  if (filters.yearMonth && filters.yearMonth !== '' && record.yearMonth !== filters.yearMonth) {
    return false;
  }
  
  // 日付範囲フィルタ（開始日）
  if (filters.startDay && filters.startDay > 0) {
    if (record.day < filters.startDay) {
      return false;
    }
  }
  
  // 日付範囲フィルタ（終了日）
  if (filters.endDay && filters.endDay > 0) {
    if (record.day > filters.endDay) {
      return false;
    }
  }
  
  // ステータスフィルタ
  if (filters.status && filters.status !== 'all' && record.status !== filters.status) {
    return false;
  }
  
  // 担当フィルタ
  if (filters.staff && filters.staff !== 'all' && record.staff !== filters.staff) {
    return false;
  }
  
  // 決済方法フィルタ
  if (filters.payment && filters.payment !== 'all' && record.payment !== filters.payment) {
    return false;
  }
  
  // メニューフィルタ
  if (filters.menu && filters.menu !== 'all' && record.menu !== filters.menu) {
    return false;
  }
  
  // 患者名検索
  if (filters.search && filters.search !== '') {
    var searchLower = filters.search.toLowerCase();
    var patientLower = record.patient.toLowerCase();
    if (patientLower.indexOf(searchLower) === -1) {
      return false;
    }
  }
  
  return true;
}

/**
 * ソート処理
 */
function sortReservations_(reservations, sortKey, sortOrder) {
  var multiplier = (sortOrder === 'asc') ? 1 : -1;
  
  reservations.sort(function(a, b) {
    var valA, valB;
    
    switch (sortKey) {
      case 'date':
        // 日付 → 時間
        if (a.date !== b.date) {
          return (a.date < b.date ? -1 : 1) * multiplier;
        }
        return (a.start < b.start ? -1 : 1) * multiplier;
        
      case 'amount':
        return (a.amount - b.amount) * multiplier;
        
      case 'patient':
        return a.patient.localeCompare(b.patient, 'ja') * multiplier;
        
      case 'menu':
        return a.menu.localeCompare(b.menu, 'ja') * multiplier;
        
      default:
        return 0;
    }
  });
  
  return reservations;
}

/**
 * 空の集計オブジェクトを作成
 */
function createEmptySummary_() {
  return {
    totalCount: 0,
    totalAmount: 0,
    byPayment: {
      '現金': 0,
      'クレジット': 0,
      '回数券': 0,
      'PayPay': 0
    },
    byMenu: {},
    byPatient: {}
  };
}

/**
 * 集計を計算
 */
function calculateSummary_(reservations) {
  var summary = createEmptySummary_();
  summary.totalCount = reservations.length;
  
  for (var i = 0; i < reservations.length; i++) {
    var r = reservations[i];
    summary.totalAmount += r.amount;
    
    // 決済方法別
    if (summary.byPayment.hasOwnProperty(r.payment)) {
      summary.byPayment[r.payment] += r.amount;
    }
    
    // メニュー別
    if (r.menu) {
      if (!summary.byMenu[r.menu]) {
        summary.byMenu[r.menu] = { count: 0, amount: 0 };
      }
      summary.byMenu[r.menu].count++;
      summary.byMenu[r.menu].amount += r.amount;
    }
    
    // 患者別
    if (r.patient) {
      if (!summary.byPatient[r.patient]) {
        summary.byPatient[r.patient] = { count: 0, amount: 0 };
      }
      summary.byPatient[r.patient].count++;
      summary.byPatient[r.patient].amount += r.amount;
    }
  }
  
  return summary;
}

/**
 * 予約データを更新
 */
function updateReservation(rowIndex, updates) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var col = getColumnIndexes_(headers);
    
    var fieldMap = {
      amount: col.amount,
      payment: col.payment,
      status: col.status,
      memo: col.memo
    };
    
    for (var field in updates) {
      if (updates.hasOwnProperty(field) && fieldMap.hasOwnProperty(field) && fieldMap[field] >= 0) {
        var colNum = fieldMap[field] + 1;
        sheet.getRange(rowIndex, colNum).setValue(updates[field]);
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true };
    
  } catch (error) {
    console.error('updateReservation error:', error);
    throw new Error('更新エラー: ' + error.message);
  }
}