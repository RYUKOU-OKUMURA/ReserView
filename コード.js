/**
 * 経営みえる化くん - 予約管理・経営分析システム
 * Google Apps Script バックエンド
 * 
 * Phase 1: モード切替UI + 経理モード ✅
 * Phase 2: 分析モード ✅
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
 * Patientsシートのヘッダーインデックスを取得
 */
function getPatientColumnIndexes_(headers) {
  return {
    id: headers.indexOf('患者ID'),
    name: headers.indexOf('患者名'),
    furigana: headers.indexOf('フリガナ'),
    gender: headers.indexOf('性別'),
    phone: headers.indexOf('電話番号'),
    memo: headers.indexOf('メモ'),
    firstVisit: headers.indexOf('初回来院日'),
    lastVisit: headers.indexOf('最終来院日'),
    visitCount: headers.indexOf('来院回数'),
    totalAmount: headers.indexOf('総支払額')
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
 * 日付をフォーマット（yyyy/MM/dd）
 */
function formatDate_(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(dateValue);
}

/**
 * 日付をyyyy-MM形式に変換
 */
function formatYearMonth_(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy-MM');
  }
  // 文字列の場合
  var str = String(dateValue);
  if (str.indexOf('/') > 0) {
    var parts = str.split('/');
    if (parts.length >= 2) {
      return parts[0] + '-' + ('0' + parts[1]).slice(-2);
    }
  }
  return str.substring(0, 7);
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
 * 2つの日付間の日数を計算
 */
function daysBetween_(date1, date2) {
  var oneDay = 24 * 60 * 60 * 1000;
  return Math.floor((date1 - date2) / oneDay);
}

/**
 * 年月文字列からDateオブジェクトを作成（月初）
 */
function yearMonthToDate_(yearMonth) {
  var parts = yearMonth.split('-');
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
}

/**
 * 前月の年月を取得
 */
function getPreviousMonth_(yearMonth) {
  var date = yearMonthToDate_(yearMonth);
  date.setMonth(date.getMonth() - 1);
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
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
 * 分析データ一括取得（サマリー・トレンド・離反リスト）
 * @param {string} yearMonth - 対象年月（例: "2025-12"）
 * @param {number} months - トレンド取得月数（デフォルト12）
 * @return {Object} { summary, trend, churnList }
 */
function getAnalysisBundle(yearMonth, months) {
  try {
    months = months || 12;
    if (!yearMonth) {
      yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    }
    
    var cache = CacheService.getScriptCache();
    var cacheKey = ['analysisBundle', yearMonth, months].join(':');
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // 破損キャッシュは無視して計算し直す
      }
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var reservationsSheet = ss.getSheetByName(SHEET_NAME);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    if (!reservationsSheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }
    if (!patientsSheet) {
      throw new Error('シートが見つかりません: ' + PATIENTS_SHEET);
    }
    
    var resData = reservationsSheet.getDataRange().getValues();
    var resHeaders = resData[0];
    var resCol = getColumnIndexes_(resHeaders);
    
    var patData = patientsSheet.getDataRange().getValues();
    var patHeaders = patData[0];
    var patCol = getPatientColumnIndexes_(patHeaders);
    
    var prevMonth = getPreviousMonth_(yearMonth);
    var currentMonthData = { sales: 0, count: 0, patients: {} };
    var prevMonthData = { sales: 0, count: 0, patients: {} };
    var allReservationsByPatient = {};
    var monthlyData = {};
    
    // 予約データ集計
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      var dateValue = row[resCol.date];
      if (!dateValue) continue;
      
      var ym = formatYearMonth_(dateValue);
      var amount = parseAmount_(row[resCol.amount]);
      var patient = row[resCol.patient] || '';
      
      // 患者別訪問月を記録
      if (patient) {
        if (!allReservationsByPatient[patient]) {
          allReservationsByPatient[patient] = [];
        }
        if (allReservationsByPatient[patient].indexOf(ym) === -1) {
          allReservationsByPatient[patient].push(ym);
        }
      }
      
      // 月次集計
      if (!monthlyData[ym]) {
        monthlyData[ym] = { sales: 0, count: 0, patients: {} };
      }
      monthlyData[ym].sales += amount;
      monthlyData[ym].count++;
      if (patient) {
        monthlyData[ym].patients[patient] = true;
      }
      
      // 今月/前月集計
      if (ym === yearMonth) {
        currentMonthData.sales += amount;
        currentMonthData.count++;
        if (patient) currentMonthData.patients[patient] = true;
      }
      if (ym === prevMonth) {
        prevMonthData.sales += amount;
        prevMonthData.count++;
        if (patient) prevMonthData.patients[patient] = true;
      }
    }
    
    // リピート率計算
    var currentUniquePatients = Object.keys(currentMonthData.patients).length;
    var repeatCount = 0;
    var currentPatients = Object.keys(currentMonthData.patients);
    for (var j = 0; j < currentPatients.length; j++) {
      var patientName = currentPatients[j];
      var visitMonths = allReservationsByPatient[patientName] || [];
      var hasPastVisit = false;
      for (var k = 0; k < visitMonths.length; k++) {
        if (visitMonths[k] < yearMonth) {
          hasPastVisit = true;
          break;
        }
      }
      if (hasPastVisit) {
        repeatCount++;
      }
    }
    var repeatRate = currentUniquePatients > 0 ? (repeatCount / currentUniquePatients * 100) : 0;
    repeatRate = Math.round(repeatRate * 10) / 10;
    
    // 前月比
    var currentUniqueCount = currentUniquePatients;
    var prevUniquePatients = Object.keys(prevMonthData.patients).length;
    var salesDiff = currentMonthData.sales - prevMonthData.sales;
    var salesRate = prevMonthData.sales > 0 ? (salesDiff / prevMonthData.sales * 100) : 0;
    var countDiff = currentMonthData.count - prevMonthData.count;
    var countRate = prevMonthData.count > 0 ? (countDiff / prevMonthData.count * 100) : 0;
    
    // 離反リスト
    var today = new Date();
    var churnWarning = 0;
    var churnConfirmed = 0;
    var churnList = [];
    
    for (var p = 1; p < patData.length; p++) {
      var patRow = patData[p];
      var lastVisit = patRow[patCol.lastVisit];
      
      if (!lastVisit || !(lastVisit instanceof Date)) continue;
      
      var daysSince = daysBetween_(today, lastVisit);
      
      if (daysSince >= 180) {
        churnConfirmed++;
      } else if (daysSince >= 90) {
        churnWarning++;
      }
      
      if (daysSince >= 90) {
        churnList.push({
          patientId: patRow[patCol.id] || '',
          patientName: patRow[patCol.name] || '',
          lastVisit: formatDate_(lastVisit),
          daysSinceVisit: daysSince,
          status: daysSince >= 180 ? 'churned' : 'warning',
          totalVisits: parseInt(patRow[patCol.visitCount]) || 0,
          totalAmount: parseAmount_(patRow[patCol.totalAmount])
        });
      }
    }
    
    churnList.sort(function(a, b) {
      return b.daysSinceVisit - a.daysSinceVisit;
    });
    
    // トレンドデータ
    var sortedMonths = Object.keys(monthlyData).sort().reverse();
    var currentYM = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    var filteredMonths = sortedMonths.filter(function(ym) {
      return ym <= currentYM;
    });
    var targetMonths = filteredMonths.slice(0, months);
    targetMonths.reverse();
    
    var trend = targetMonths.map(function(ym) {
      var d = monthlyData[ym];
      return {
        yearMonth: ym,
        sales: d.sales,
        count: d.count,
        uniquePatients: Object.keys(d.patients).length
      };
    });
    
    var summary = {
      currentMonth: {
        sales: currentMonthData.sales,
        count: currentMonthData.count,
        uniquePatients: currentUniqueCount
      },
      previousMonth: {
        sales: prevMonthData.sales,
        count: prevMonthData.count,
        uniquePatients: prevUniquePatients
      },
      comparison: {
        salesDiff: salesDiff,
        salesRate: Math.round(salesRate * 100) / 100,
        countDiff: countDiff,
        countRate: Math.round(countRate * 100) / 100
      },
      repeatRate: repeatRate,
      churnWarning: churnWarning,
      churnConfirmed: churnConfirmed
    };
    
    var result = {
      summary: summary,
      trend: trend,
      churnList: churnList
    };
    
    cache.put(cacheKey, JSON.stringify(result), 60);
    return result;
    
  } catch (error) {
    console.error('getAnalysisBundle error:', error);
    throw new Error('分析バンドル取得エラー: ' + error.message);
  }
}

/**
 * 分析ダッシュボード用データを取得
 * @param {string} yearMonth - 対象年月（例: "2025-12"）
 * @return {Object} ダッシュボードデータ
 */
function getAnalysisSummary(yearMonth) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var reservationsSheet = ss.getSheetByName(SHEET_NAME);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    // 予約データを取得
    var resData = reservationsSheet.getDataRange().getValues();
    var resHeaders = resData[0];
    var resCol = getColumnIndexes_(resHeaders);
    
    // 患者データを取得
    var patData = patientsSheet.getDataRange().getValues();
    var patHeaders = patData[0];
    var patCol = getPatientColumnIndexes_(patHeaders);
    
    // 対象年月がなければ今月を使用
    if (!yearMonth) {
      yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    }
    var prevMonth = getPreviousMonth_(yearMonth);
    
    // 月別集計用
    var currentMonthData = { sales: 0, count: 0, patients: {} };
    var prevMonthData = { sales: 0, count: 0, patients: {} };
    var allReservationsByPatient = {};  // リピート率計算用
    
    // 予約データをループして集計
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      var dateValue = row[resCol.date];
      if (!dateValue) continue;
      
      var ym = formatYearMonth_(dateValue);
      var amount = parseAmount_(row[resCol.amount]);
      var patient = row[resCol.patient] || '';
      
      if (!patient) continue;
      
      // 患者ごとの来院月を記録（リピート率計算用）
      if (!allReservationsByPatient[patient]) {
        allReservationsByPatient[patient] = [];
      }
      if (allReservationsByPatient[patient].indexOf(ym) === -1) {
        allReservationsByPatient[patient].push(ym);
      }
      
      // 今月の集計
      if (ym === yearMonth) {
        currentMonthData.sales += amount;
        currentMonthData.count++;
        currentMonthData.patients[patient] = true;
      }
      
      // 前月の集計
      if (ym === prevMonth) {
        prevMonthData.sales += amount;
        prevMonthData.count++;
        prevMonthData.patients[patient] = true;
      }
    }
    
    // ユニーク患者数
    var currentUniquePatients = Object.keys(currentMonthData.patients).length;
    var prevUniquePatients = Object.keys(prevMonthData.patients).length;
    
    // リピート率計算（今月来院者のうち、過去にも来院履歴がある人の割合）
    var repeatCount = 0;
    var currentPatients = Object.keys(currentMonthData.patients);
    for (var j = 0; j < currentPatients.length; j++) {
      var patientName = currentPatients[j];
      var visitMonths = allReservationsByPatient[patientName] || [];
      // 今月以外にも来院履歴があればリピーター
      var hasPastVisit = false;
      for (var k = 0; k < visitMonths.length; k++) {
        if (visitMonths[k] < yearMonth) {
          hasPastVisit = true;
          break;
        }
      }
      if (hasPastVisit) {
        repeatCount++;
      }
    }
    var repeatRate = currentUniquePatients > 0 ? (repeatCount / currentUniquePatients * 100) : 0;
    
    // 離反数を計算
    var today = new Date();
    var churnWarning = 0;
    var churnConfirmed = 0;
    
    for (var p = 1; p < patData.length; p++) {
      var patRow = patData[p];
      var lastVisit = patRow[patCol.lastVisit];
      
      if (!lastVisit || !(lastVisit instanceof Date)) continue;
      
      var daysSince = daysBetween_(today, lastVisit);
      
      if (daysSince >= 180) {
        churnConfirmed++;
      } else if (daysSince >= 90) {
        churnWarning++;
      }
    }
    
    // 前月比計算
    var salesDiff = currentMonthData.sales - prevMonthData.sales;
    var salesRate = prevMonthData.sales > 0 ? (salesDiff / prevMonthData.sales * 100) : 0;
    var countDiff = currentMonthData.count - prevMonthData.count;
    var countRate = prevMonthData.count > 0 ? (countDiff / prevMonthData.count * 100) : 0;
    
    return {
      currentMonth: {
        sales: currentMonthData.sales,
        count: currentMonthData.count,
        uniquePatients: currentUniquePatients
      },
      previousMonth: {
        sales: prevMonthData.sales,
        count: prevMonthData.count,
        uniquePatients: prevUniquePatients
      },
      comparison: {
        salesDiff: salesDiff,
        salesRate: Math.round(salesRate * 100) / 100,
        countDiff: countDiff,
        countRate: Math.round(countRate * 100) / 100
      },
      repeatRate: Math.round(repeatRate * 10) / 10,
      churnWarning: churnWarning,
      churnConfirmed: churnConfirmed
    };
    
  } catch (error) {
    console.error('getAnalysisSummary error:', error);
    throw new Error('分析データ取得エラー: ' + error.message);
  }
}

/**
 * 顧客分析データを取得（来院回数分布）
 * @param {string} yearMonth - 対象年月（未使用、将来の拡張用）
 * @return {Object} 顧客分析データ
 */
function getCustomerAnalysis(yearMonth) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    var data = patientsSheet.getDataRange().getValues();
    var headers = data[0];
    var col = getPatientColumnIndexes_(headers);
    
    // 来院回数分布
    var distribution = {
      '1回': 0,
      '2-5回': 0,
      '6-10回': 0,
      '11-20回': 0,
      '21回以上': 0
    };
    
    var totalPatients = 0;
    var totalVisits = 0;
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var visitCount = parseInt(row[col.visitCount]) || 0;
      
      if (visitCount === 0) continue;
      
      totalPatients++;
      totalVisits += visitCount;
      
      if (visitCount === 1) {
        distribution['1回']++;
      } else if (visitCount <= 5) {
        distribution['2-5回']++;
      } else if (visitCount <= 10) {
        distribution['6-10回']++;
      } else if (visitCount <= 20) {
        distribution['11-20回']++;
      } else {
        distribution['21回以上']++;
      }
    }
    
    var avgVisits = totalPatients > 0 ? Math.round(totalVisits / totalPatients * 10) / 10 : 0;
    
    return {
      visitDistribution: distribution,
      totalPatients: totalPatients,
      averageVisits: avgVisits
    };
    
  } catch (error) {
    console.error('getCustomerAnalysis error:', error);
    throw new Error('顧客分析データ取得エラー: ' + error.message);
  }
}

/**
 * 離反リストを取得
 * @param {string} baseDate - 基準日（省略時は今日）
 * @return {Array} 離反リスト
 */
function getChurnList(baseDate) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);
    
    if (!patientsSheet) {
      return [];
    }
    
    var data = patientsSheet.getDataRange().getValues();
    var headers = data[0];
    var col = getPatientColumnIndexes_(headers);
    
    var today = baseDate ? new Date(baseDate) : new Date();
    var churnList = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var lastVisit = row[col.lastVisit];
      
      if (!lastVisit || !(lastVisit instanceof Date)) continue;
      
      var daysSince = daysBetween_(today, lastVisit);
      
      // 90日以上未来院の患者を抽出
      if (daysSince >= 90) {
        churnList.push({
          patientId: row[col.id] || '',
          patientName: row[col.name] || '',
          lastVisit: formatDate_(lastVisit),
          daysSinceVisit: daysSince,
          status: daysSince >= 180 ? 'churned' : 'warning',
          totalVisits: parseInt(row[col.visitCount]) || 0,
          totalAmount: parseAmount_(row[col.totalAmount])
        });
      }
    }
    
    // 未来院日数でソート（多い順）
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
 * @param {number} months - 取得する月数（デフォルト12）
 * @return {Array} 月別売上データ
 */
function getSalesTrend(months) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    months = months || 12;
    
    // 月別集計
    var monthlyData = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateValue = row[col.date];
      if (!dateValue) continue;
      
      var ym = formatYearMonth_(dateValue);
      var amount = parseAmount_(row[col.amount]);
      var patient = row[col.patient] || '';
      
      if (!monthlyData[ym]) {
        monthlyData[ym] = { sales: 0, count: 0, patients: {} };
      }
      
      monthlyData[ym].sales += amount;
      monthlyData[ym].count++;
      if (patient) {
        monthlyData[ym].patients[patient] = true;
      }
    }
    
    // 直近N ヶ月を抽出
    var sortedMonths = Object.keys(monthlyData).sort().reverse();
    
    // 未来の月（2026年以降）を除外し、過去のデータのみ取得
    var today = new Date();
    var currentYM = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM');
    
    var filteredMonths = sortedMonths.filter(function(ym) {
      return ym <= currentYM;
    });
    
    var targetMonths = filteredMonths.slice(0, months);
    
    // 結果を古い順に並べ替え
    targetMonths.reverse();
    
    var result = targetMonths.map(function(ym) {
      var d = monthlyData[ym];
      return {
        yearMonth: ym,
        sales: d.sales,
        count: d.count,
        uniquePatients: Object.keys(d.patients).length
      };
    });
    
    return result;
    
  } catch (error) {
    console.error('getSalesTrend error:', error);
    throw new Error('売上トレンド取得エラー: ' + error.message);
  }
}

/**
 * メニュー分析データを取得
 * @param {string} yearMonth - 対象年月（省略時は今月）
 * @return {Object} メニュー分析データ
 */
function getMenuAnalysis(yearMonth) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    // 対象年月がなければ今月を使用
    if (!yearMonth) {
      yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    }
    
    // メニュー別集計
    var menuData = {};
    var totalSales = 0;
    var totalCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateValue = row[col.date];
      if (!dateValue) continue;
      
      var ym = formatYearMonth_(dateValue);
      if (ym !== yearMonth) continue;
      
      var menu = row[col.menu] || '未設定';
      var amount = parseAmount_(row[col.amount]);
      
      if (!menuData[menu]) {
        menuData[menu] = { count: 0, amount: 0 };
      }
      
      menuData[menu].count++;
      menuData[menu].amount += amount;
      totalSales += amount;
      totalCount++;
    }
    
    // 配列に変換して売上順にソート
    var byMenu = [];
    for (var menuName in menuData) {
      if (menuData.hasOwnProperty(menuName)) {
        var d = menuData[menuName];
        byMenu.push({
          menu: menuName,
          count: d.count,
          amount: d.amount,
          percentage: totalSales > 0 ? Math.round(d.amount / totalSales * 1000) / 10 : 0
        });
      }
    }
    
    // 売上額でソート（降順）
    byMenu.sort(function(a, b) {
      return b.amount - a.amount;
    });
    
    return {
      byMenu: byMenu,
      totalSales: totalSales,
      totalCount: totalCount
    };
    
  } catch (error) {
    console.error('getMenuAnalysis error:', error);
    throw new Error('メニュー分析取得エラー: ' + error.message);
  }
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

/**
 * 指定月の売上合計を取得（CFモード用）
 * @param {string} yearMonth - 対象年月
 * @return {number} 売上合計
 */
function getMonthlySales(yearMonth) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    var totalSales = 0;
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateValue = row[col.date];
      if (!dateValue) continue;
      
      var ym = formatYearMonth_(dateValue);
      if (ym === yearMonth) {
        totalSales += parseAmount_(row[col.amount]);
      }
    }
    
    return totalSales;
    
  } catch (error) {
    console.error('getMonthlySales error:', error);
    throw new Error('月別売上取得エラー: ' + error.message);
  }
}
