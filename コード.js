/**
 * 経営みえる化くん - 予約管理・経営分析システム
 * Google Apps Script バックエンド
 * 
 * Phase 1: モード切替UI + 経理モード
 * Phase 2: 分析モード（TODO）
 * Phase 3: CFモード（TODO）
 */

// ========================================
// 設定
// ========================================
var SPREADSHEET_ID = '1z-OuS5riqLp8PYKECOnPzjBWPjgvUa6KKg5c4Ne-g08';
var SHEET_NAME = 'Reservations';
var PATIENTS_SHEET = 'Patients';
var EXPENSES_SHEET = 'Expenses';  // Phase 3で使用

// ========================================
// Webアプリ エントリーポイント
// ========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('経営みえる化くん')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ========================================
// ユーティリティ関数
// ========================================

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

// ========================================
// 経理モード用関数（既存）
// ========================================

/**
 * フィルタ用の選択肢を取得
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

// ========================================
// 分析モード用関数（Phase 2）
// ========================================

/**
 * 分析ダッシュボード用データを取得
 * @param {string} yearMonth - 対象年月（例: "2025-12"）
 * @return {Object} ダッシュボードデータ
 */
function getAnalysisSummary(yearMonth) {
  // TODO: Phase 2で実装
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var reservationsSheet = ss.getSheetByName(SHEET_NAME);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    // 仮のデータを返す
    return {
      currentMonth: {
        sales: 850000,
        count: 95,
        uniquePatients: 42
      },
      previousMonth: {
        sales: 780000,
        count: 88,
        uniquePatients: 38
      },
      comparison: {
        salesDiff: 70000,
        salesRate: 8.97,
        countDiff: 7,
        countRate: 7.95
      },
      repeatRate: 78.5,
      churnWarning: 6,
      churnConfirmed: 7
    };
    
  } catch (error) {
    console.error('getAnalysisSummary error:', error);
    throw new Error('分析データ取得エラー: ' + error.message);
  }
}

/**
 * 顧客分析データを取得
 * @param {string} yearMonth - 対象年月
 * @return {Object} 顧客分析データ
 */
function getCustomerAnalysis(yearMonth) {
  // TODO: Phase 2で実装
  return {
    repeatRatePeriod: 78.5,
    retentionRate: 65.2,
    visitDistribution: {
      '1回': 8,
      '2-5回': 19,
      '6-10回': 16,
      '11回以上': 32
    }
  };
}

/**
 * 離反リストを取得
 * @param {string} baseDate - 基準日
 * @return {Array} 離反リスト
 */
function getChurnList(baseDate) {
  // TODO: Phase 2で実装
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    if (!patientsSheet) {
      return [];
    }
    
    var data = patientsSheet.getDataRange().getValues();
    var headers = data[0];
    
    // 列インデックスを取得
    var col = {
      id: headers.indexOf('患者ID'),
      name: headers.indexOf('患者名'),
      lastVisit: headers.indexOf('最終来院日'),
      visitCount: headers.indexOf('来院回数'),
      totalAmount: headers.indexOf('総支払額')
    };
    
    var today = baseDate ? new Date(baseDate) : new Date();
    var churnList = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var lastVisit = row[col.lastVisit];
      
      if (!lastVisit || !(lastVisit instanceof Date)) continue;
      
      var daysSince = Math.floor((today - lastVisit) / (1000 * 60 * 60 * 24));
      
      // 90日以上未来院の患者を抽出
      if (daysSince >= 90) {
        churnList.push({
          patientId: row[col.id] || '',
          patientName: row[col.name] || '',
          lastVisit: formatDate_(lastVisit),
          daysSinceVisit: daysSince,
          status: daysSince >= 180 ? 'churned' : 'warning',
          totalVisits: row[col.visitCount] || 0,
          totalAmount: row[col.totalAmount] || 0
        });
      }
    }
    
    // 未来院日数でソート
    churnList.sort(function(a, b) {
      return b.daysSinceVisit - a.daysSinceVisit;
    });
    
    return churnList;
    
  } catch (error) {
    console.error('getChurnList error:', error);
    throw new Error('離反リスト取得エラー: ' + error.message);
  }
}

/**
 * 売上トレンドデータを取得
 * @param {number} months - 取得する月数
 * @return {Array} 月別売上データ
 */
function getSalesTrend(months) {
  // TODO: Phase 2で実装
  return [];
}

/**
 * メニュー分析データを取得
 * @param {string} yearMonth - 対象年月
 * @return {Object} メニュー分析データ
 */
function getMenuAnalysis(yearMonth) {
  // TODO: Phase 2で実装
  return {
    byMenu: [],
    totalSales: 0
  };
}

// ========================================
// CFモード用関数（Phase 3）
// ========================================

/**
 * 経費データを取得
 * @param {string} yearMonth - 対象年月
 * @return {Object} 経費データ
 */
function getExpenses(yearMonth) {
  // TODO: Phase 3で実装
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EXPENSES_SHEET);
    
    if (!sheet) {
      return null;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === yearMonth) {
        return {
          yearMonth: yearMonth,
          variable: data[i][1] || 0,
          labor: data[i][2] || 0,
          otherFixed: data[i][3] || 0,
          depreciation: data[i][4] || 0,
          loanPayment: data[i][5] || 0,
          capex: data[i][6] || 0
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('getExpenses error:', error);
    throw new Error('経費データ取得エラー: ' + error.message);
  }
}

/**
 * 経費データを保存
 * @param {string} yearMonth - 対象年月
 * @param {Object} data - 経費データ
 * @return {Object} 結果
 */
function saveExpenses(yearMonth, data) {
  // TODO: Phase 3で実装
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EXPENSES_SHEET);
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet(EXPENSES_SHEET);
      sheet.appendRow(['年月', '変動費', '人件費', 'その他固定費', '減価償却費', '借入返済', '設備投資', '更新日時', '更新者']);
    }
    
    var existingData = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // 既存データを検索
    for (var i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === yearMonth) {
        rowIndex = i + 1;
        break;
      }
    }
    
    var rowData = [
      yearMonth,
      data.variable || 0,
      data.labor || 0,
      data.otherFixed || 0,
      data.depreciation || 0,
      data.loanPayment || 0,
      data.capex || 0,
      new Date(),
      Session.getActiveUser().getEmail()
    ];
    
    if (rowIndex > 0) {
      // 更新
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 新規追加
      sheet.appendRow(rowData);
    }
    
    SpreadsheetApp.flush();
    return { success: true };
    
  } catch (error) {
    console.error('saveExpenses error:', error);
    throw new Error('経費データ保存エラー: ' + error.message);
  }
}

/**
 * キャッシュフロー履歴を取得
 * @param {number} months - 取得する月数
 * @return {Array} CF履歴データ
 */
function getCashFlowHistory(months) {
  // TODO: Phase 3で実装
  return [];
}