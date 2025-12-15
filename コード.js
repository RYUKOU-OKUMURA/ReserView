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
var ANNUAL_PLAN_SHEET = 'AnnualPlans'; // 年次の予算/計画（CF用）

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
 * シートから実データ範囲のみ取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array} 2次元配列
 */
function getSheetValues_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return [];
  }
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
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

/**
 * CacheService を用いた汎用キャッシュヘルパー（JSONシリアライズ）
 * @param {string} key - キャッシュキー
 * @param {number} ttlSeconds - 有効期限（秒）
 * @param {function(): *} computeFn - キャッシュミス時の計算処理
 * @return {*} 計算結果
 */
function withCache_(key, ttlSeconds, computeFn) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached !== null && cached !== '') {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // 壊れたキャッシュは無視して計算する
    }
  }

  var result = computeFn();

  try {
    cache.put(key, JSON.stringify(result), ttlSeconds);
  } catch (e) {
    // キャッシュ失敗は致命的でないので無視
  }

  return result;
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
    var data = getSheetValues_(sheet);
    if (data.length === 0) {
      return { staff: [], menu: [], payment: [], status: [], yearMonths: [] };
    }
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
 * 初期データ一括取得（フィルタ選択肢 + 対象月の予約 + サマリー）
 * @param {string} yearMonth - 対象年月（省略時は今月 or 最新月）
 * @return {Object} { filterOptions, reservations, summary, selectedYearMonth }
 */
function getInitialData(yearMonth) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }
    var data = getSheetValues_(sheet);
    if (data.length < 2) {
      return {
        filterOptions: { staff: [], menu: [], payment: [], status: [], yearMonths: [] },
        reservations: [],
        summary: createEmptySummary_(),
        selectedYearMonth: yearMonth || ''
      };
    }

    var headers = data[0];
    var col = getColumnIndexes_(headers);
    var staffSet = {};
    var menuSet = {};
    var paymentSet = {};
    var statusSet = {};
    var yearMonthSet = {};
    var allReservations = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[col.id] && !row[col.patient]) continue;

      var formattedDate = formatDate_(row[col.date]);
      var formattedStart = formatTime_(row[col.start]);
      var formattedEnd = formatTime_(row[col.end]);
      var amount = parseAmount_(row[col.amount]);
      var ym = formattedDate ? formattedDate.substring(0, 7).replace('/', '-') : '';
      var day = extractDay_(formattedDate);

      if (row[col.staff]) staffSet[row[col.staff]] = true;
      if (row[col.menu]) menuSet[row[col.menu]] = true;
      if (row[col.payment]) paymentSet[row[col.payment]] = true;
      if (row[col.status]) statusSet[row[col.status]] = true;
      if (ym) yearMonthSet[ym] = true;

      allReservations.push({
        rowIndex: i + 1,
        date: formattedDate,
        yearMonth: ym,
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
      });
    }

    var yearMonths = Object.keys(yearMonthSet).sort().reverse();
    var selectedYM = yearMonth;
    var currentYM = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    if (!selectedYM) {
      if (yearMonths.indexOf(currentYM) !== -1) {
        selectedYM = currentYM;
      } else {
        selectedYM = yearMonths[0] || '';
      }
    }

    var filteredReservations = allReservations.filter(function(r) {
      if (!selectedYM) return true;
      return r.yearMonth === selectedYM;
    });

    var summary = calculateSummary_(filteredReservations);

    return {
      filterOptions: {
        staff: Object.keys(staffSet).sort(),
        menu: Object.keys(menuSet).sort(),
        payment: Object.keys(paymentSet).sort(),
        status: Object.keys(statusSet).sort(),
        yearMonths: yearMonths
      },
      reservations: filteredReservations,
      summary: summary,
      selectedYearMonth: selectedYM
    };
  } catch (error) {
    console.error('getInitialData error:', error);
    throw new Error('初期データ取得エラー: ' + error.message);
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
    
    var data = getSheetValues_(sheet);
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
  months = months || 12;
  if (!yearMonth) {
    yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  }

  var cacheKey = ['analysisBundle', yearMonth, months].join(':');
  return withCache_(cacheKey, 600, function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var reservationsSheet = ss.getSheetByName(SHEET_NAME);
      var patientsSheet = ss.getSheetByName(PATIENTS_SHEET);

      if (!reservationsSheet) {
        throw new Error('シートが見つかりません: ' + SHEET_NAME);
      }
      if (!patientsSheet) {
        throw new Error('シートが見つかりません: ' + PATIENTS_SHEET);
      }

      var resData = getSheetValues_(reservationsSheet);
      var patData = getSheetValues_(patientsSheet);

      if (resData.length < 1) {
        return {
          summary: {
            currentMonth: { sales: 0, count: 0, uniquePatients: 0 },
            previousMonth: { sales: 0, count: 0, uniquePatients: 0 },
            comparison: { salesDiff: 0, salesRate: 0, countDiff: 0, countRate: 0 },
            repeatRate: 0,
            churnWarning: 0,
            churnConfirmed: 0
          },
          trend: [],
          churnList: []
        };
      }

      var resHeaders = resData[0];
      var resCol = getColumnIndexes_(resHeaders);

      var patHeaders = patData[0] || [];
      var patCol = patHeaders.length ? getPatientColumnIndexes_(patHeaders) : {};

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

      return {
        summary: summary,
        trend: trend,
        churnList: churnList
      };
    } catch (error) {
      console.error('getAnalysisBundle error:', error);
      throw new Error('分析バンドル取得エラー: ' + error.message);
    }
  });
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
    var resData = getSheetValues_(reservationsSheet);
    if (resData.length < 1) {
      return {
        currentMonth: { sales: 0, count: 0, uniquePatients: 0 },
        previousMonth: { sales: 0, count: 0, uniquePatients: 0 },
        comparison: { salesDiff: 0, salesRate: 0, countDiff: 0, countRate: 0 },
        repeatRate: 0,
        churnWarning: 0,
        churnConfirmed: 0
      };
    }
    var resHeaders = resData[0];
    var resCol = getColumnIndexes_(resHeaders);
    
    // 患者データを取得
    var patData = getSheetValues_(patientsSheet);
    var patHeaders = patData[0] || [];
    var patCol = patHeaders.length ? getPatientColumnIndexes_(patHeaders) : {};
    
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
    var reservationsSheet = ss.getSheetByName(SHEET_NAME);
    
    var patData = getSheetValues_(patientsSheet);
    var resData = getSheetValues_(reservationsSheet);
    if (patData.length < 1 || resData.length < 1) {
      return {
        visitDistribution: {
          '1回': 0,
          '2-5回': 0,
          '6-10回': 0,
          '11-20回': 0,
          '21回以上': 0
        },
        totalPatients: 0,
        averageVisits: 0
      };
    }
    var patHeaders = patData[0];
    var patCol = getPatientColumnIndexes_(patHeaders);
    var resHeaders = resData[0];
    var resCol = getColumnIndexes_(resHeaders);
    
    // 対象年月に来院した患者のみを抽出
    var targetYM = yearMonth || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    var targetPatients = {};
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      var dateValue = row[resCol.date];
      if (!dateValue) continue;
      var ym = formatYearMonth_(dateValue);
      if (ym !== targetYM) continue;
      var patientName = row[resCol.patient] || '';
      if (patientName) {
        targetPatients[patientName] = true;
      }
    }
    
    // 来院していない場合は0を返す
    var targetNames = Object.keys(targetPatients);
    if (targetNames.length === 0) {
      return {
        visitDistribution: {
          '1回': 0,
          '2-5回': 0,
          '6-10回': 0,
          '11-20回': 0,
          '21回以上': 0
        },
        totalPatients: 0,
        averageVisits: 0
      };
    }
    
    // 患者シートをマップ化（名前ベース）
    var patientMap = {};
    for (var p = 1; p < patData.length; p++) {
      var prow = patData[p];
      var name = prow[patCol.name] || '';
      if (name) {
        patientMap[name] = prow;
      }
    }
    
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
    
    for (var t = 0; t < targetNames.length; t++) {
      var name = targetNames[t];
      var prowData = patientMap[name];
      var visitCount = 0;
      if (prowData) {
        visitCount = parseInt(prowData[patCol.visitCount]) || 0;
      }
      // visitCountが0でも、対象月に来院があれば1とみなす
      if (visitCount === 0) {
        visitCount = 1;
      }
      
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
    
    var data = getSheetValues_(patientsSheet);
    if (data.length < 1) {
      return [];
    }
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
    
    var data = getSheetValues_(sheet);
    if (data.length < 1) {
      return [];
    }
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
    
    var data = getSheetValues_(sheet);
    if (data.length < 1) {
      return { byMenu: [], totalSales: 0, totalCount: 0 };
    }
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
    
    var data = getSheetValues_(sheet);
    if (data.length < 2) {
      return null;
    }
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
    
    var existingData = getSheetValues_(sheet);
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
  var cacheKey = ['monthlySales', yearMonth].join(':');
  return withCache_(cacheKey, 600, function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      
      var data = getSheetValues_(sheet);
      if (data.length < 1) {
        return 0;
      }
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
  });
}

/**
 * 指定年の売上合計を取得（ヘッダー表示用）
 * @param {number} year - 西暦（例: 2025）
 * @return {number} 年間売上合計
 */
function getYearlySales(year) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var data = getSheetValues_(sheet);
    if (data.length < 1) {
      return 0;
    }
    var headers = data[0];
    var col = getColumnIndexes_(headers);
    
    var totalSales = 0;
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateValue = row[col.date];
      if (!dateValue) continue;
      
      var dateObj = (dateValue instanceof Date) ? dateValue : null;
      var yearStr = '';
      if (dateObj) {
        yearStr = String(dateObj.getFullYear());
      } else {
        yearStr = String(dateValue).split('/')[0] || '';
      }
      if (String(year) !== yearStr) continue;
      
      totalSales += parseAmount_(row[col.amount]);
    }
    
    return totalSales;
    
  } catch (error) {
    console.error('getYearlySales error:', error);
    throw new Error('年間売上取得エラー: ' + error.message);
  }
}

// ========================================
// 年次/期間集計 + ダウンロード（Sales Summary）
// ========================================

/**
 * 年月文字列(yyyy-MM)の妥当性チェック
 * @param {string} yearMonth
 * @return {boolean}
 */
function isValidYearMonth_(yearMonth) {
  return /^\d{4}-\d{2}$/.test(String(yearMonth || ''));
}

/**
 * 指定した年月の範囲（両端含む）を yyyy-MM の配列で返す
 * @param {string} startYearMonth
 * @param {string} endYearMonth
 * @return {Array<string>}
 */
function listYearMonthsBetween_(startYearMonth, endYearMonth) {
  if (!isValidYearMonth_(startYearMonth) || !isValidYearMonth_(endYearMonth)) {
    throw new Error('年月は yyyy-MM 形式で指定してください');
  }

  var start = yearMonthToDate_(startYearMonth);
  var end = yearMonthToDate_(endYearMonth);
  if (start > end) {
    var tmp = start;
    start = end;
    end = tmp;
  }

  var yms = [];
  var cursor = new Date(start.getTime());
  while (cursor <= end) {
    yms.push(Utilities.formatDate(cursor, 'Asia/Tokyo', 'yyyy-MM'));
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return yms;
}

/**
 * 期間の売上サマリーをGoogleスプレッドシートに出力してURLを返す
 * - 画面側でExcelとしてダウンロード可能（ファイル→ダウンロード）
 * @param {string} startYearMonth - yyyy-MM
 * @param {string} endYearMonth - yyyy-MM
 * @return {Object} { spreadsheetId, url, name }
 */
function exportSalesSummary(startYearMonth, endYearMonth) {
  try {
    var yms = listYearMonthsBetween_(startYearMonth, endYearMonth);
    var rangeLabel = yms.length === 1 ? yms[0] : (yms[0] + '〜' + yms[yms.length - 1]);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var reservationsSheet = ss.getSheetByName(SHEET_NAME);
    if (!reservationsSheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }

    var data = getSheetValues_(reservationsSheet);
    if (data.length < 2) {
      data = data.slice(0, 1);
    }

    var headers = data[0] || [];
    var col = headers.length ? getColumnIndexes_(headers) : {};

    var ymSet = {};
    for (var i = 0; i < yms.length; i++) ymSet[yms[i]] = true;

    var monthly = {};
    yms.forEach(function(ym) {
      monthly[ym] = { sales: 0, count: 0, patients: {} };
    });
    var byPayment = {};
    var byMenu = {};
    var totalSales = 0;
    var totalCount = 0;
    var totalPatients = {};

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var dateValue = row[col.date];
      if (!dateValue) continue;

      var ym = formatYearMonth_(dateValue);
      if (!ymSet[ym]) continue;

      var amount = parseAmount_(row[col.amount]);
      var payment = (row[col.payment] || '').toString();
      var menu = (row[col.menu] || '未設定').toString();
      var patient = (row[col.patient] || '').toString();

      monthly[ym].sales += amount;
      monthly[ym].count++;
      if (patient) {
        monthly[ym].patients[patient] = true;
        totalPatients[patient] = true;
      }

      totalSales += amount;
      totalCount++;

      if (payment) {
        if (!byPayment[payment]) byPayment[payment] = 0;
        byPayment[payment] += amount;
      }
      if (menu) {
        if (!byMenu[menu]) byMenu[menu] = { count: 0, amount: 0 };
        byMenu[menu].count++;
        byMenu[menu].amount += amount;
      }
    }

    // 出力用スプレッドシート作成
    var name = '売上サマリー_' + rangeLabel.replace('〜', '-') + '_' +
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    var out = SpreadsheetApp.create(name);

    // フォルダへ移動（任意）
    try {
      var folderName = 'ReserView_Exports';
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var file = DriveApp.getFileById(out.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) {
      // フォルダ移動失敗は致命的ではない
    }

    // シート1: 月別サマリー
    var sheet1 = out.getSheets()[0];
    sheet1.setName('月別サマリー');
    sheet1.getRange(1, 1, 1, 6).setValues([['期間', '年月', '売上', '件数', 'ユニーク患者数', '備考']]);

    var rows1 = [];
    for (var j = 0; j < yms.length; j++) {
      var k = yms[j];
      var d = monthly[k];
      rows1.push([
        rangeLabel,
        k,
        d.sales,
        d.count,
        Object.keys(d.patients).length,
        ''
      ]);
    }
    if (rows1.length > 0) {
      sheet1.getRange(2, 1, rows1.length, 6).setValues(rows1);
    }
    sheet1.getRange(2, 3, Math.max(rows1.length, 1), 1).setNumberFormat('¥#,##0');
    sheet1.getRange(2, 3, Math.max(rows1.length, 1), 1).setHorizontalAlignment('right');
    sheet1.autoResizeColumns(1, 6);

    // 総計行
    sheet1.getRange(rows1.length + 3, 1, 1, 6).setValues([[
      '合計',
      '',
      totalSales,
      totalCount,
      Object.keys(totalPatients).length,
      ''
    ]]);
    sheet1.getRange(rows1.length + 3, 3, 1, 1).setNumberFormat('¥#,##0');

    // シート2: 決済方法別
    var sheet2 = out.insertSheet('決済方法別');
    sheet2.getRange(1, 1, 1, 4).setValues([['期間', '決済方法', '売上', '構成比(%)']]);
    var payments = Object.keys(byPayment).sort();
    var rows2 = payments.map(function(p) {
      var amt = byPayment[p] || 0;
      var pct = totalSales > 0 ? Math.round(amt / totalSales * 1000) / 10 : 0;
      return [rangeLabel, p, amt, pct];
    });
    if (rows2.length > 0) {
      sheet2.getRange(2, 1, rows2.length, 4).setValues(rows2);
      sheet2.getRange(2, 3, rows2.length, 1).setNumberFormat('¥#,##0');
      sheet2.getRange(2, 4, rows2.length, 1).setNumberFormat('0.0');
    }
    sheet2.autoResizeColumns(1, 4);

    // シート3: メニュー別
    var sheet3 = out.insertSheet('メニュー別');
    sheet3.getRange(1, 1, 1, 5).setValues([['期間', 'メニュー', '件数', '売上', '構成比(%)']]);
    var menus = [];
    for (var mn in byMenu) {
      if (byMenu.hasOwnProperty(mn)) {
        menus.push({
          menu: mn,
          count: byMenu[mn].count,
          amount: byMenu[mn].amount
        });
      }
    }
    menus.sort(function(a, b) { return b.amount - a.amount; });
    var rows3 = menus.map(function(m) {
      var pct = totalSales > 0 ? Math.round(m.amount / totalSales * 1000) / 10 : 0;
      return [rangeLabel, m.menu, m.count, m.amount, pct];
    });
    if (rows3.length > 0) {
      sheet3.getRange(2, 1, rows3.length, 5).setValues(rows3);
      sheet3.getRange(2, 4, rows3.length, 1).setNumberFormat('¥#,##0');
      sheet3.getRange(2, 5, rows3.length, 1).setNumberFormat('0.0');
    }
    sheet3.autoResizeColumns(1, 5);

    SpreadsheetApp.flush();

    return {
      spreadsheetId: out.getId(),
      url: out.getUrl(),
      name: name
    };
  } catch (error) {
    console.error('exportSalesSummary error:', error);
    throw new Error('売上サマリー出力エラー: ' + error.message);
  }
}

// ========================================
// CFモード：年次（実績/計画）
// ========================================

function normalizeAnnualPlan_(year, planRow) {
  var y = parseInt(year, 10);
  return {
    year: y,
    sales: planRow && planRow.sales ? planRow.sales : 0,
    variable: planRow && planRow.variable ? planRow.variable : 0,
    labor: planRow && planRow.labor ? planRow.labor : 0,
    otherFixed: planRow && planRow.otherFixed ? planRow.otherFixed : 0,
    depreciation: planRow && planRow.depreciation ? planRow.depreciation : 0,
    loanPayment: planRow && planRow.loanPayment ? planRow.loanPayment : 0,
    capex: planRow && planRow.capex ? planRow.capex : 0
  };
}

/**
 * 年次CF用：実績（Reservationsの年合計 + Expensesの年合計）
 * @param {number} year
 * @return {Object}
 */
function getYearlyActualCashFlow_(year) {
  var y = parseInt(year, 10);
  if (!y) throw new Error('年が不正です');

  var sales = getYearlySales(y);
  var totals = {
    year: y,
    sales: sales,
    variable: 0,
    labor: 0,
    otherFixed: 0,
    depreciation: 0,
    loanPayment: 0,
    capex: 0
  };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(EXPENSES_SHEET);
  if (!sheet) return totals;

  var values = getSheetValues_(sheet);
  if (values.length < 2) return totals;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var ym = String(row[0] || '');
    if (ym.substring(0, 4) !== String(y)) continue;
    totals.variable += parseAmount_(row[1]);
    totals.labor += parseAmount_(row[2]);
    totals.otherFixed += parseAmount_(row[3]);
    totals.depreciation += parseAmount_(row[4]);
    totals.loanPayment += parseAmount_(row[5]);
    totals.capex += parseAmount_(row[6]);
  }

  return totals;
}

/**
 * 年次計画を取得（存在しなければ0で返す）
 * @param {number} year
 * @return {Object}
 */
function getAnnualPlan(year) {
  try {
    var y = parseInt(year, 10);
    if (!y) throw new Error('年が不正です');

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ANNUAL_PLAN_SHEET);
    if (!sheet) return normalizeAnnualPlan_(y, null);

    var values = getSheetValues_(sheet);
    if (values.length < 2) return normalizeAnnualPlan_(y, null);

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (parseInt(row[0], 10) !== y) continue;
      return normalizeAnnualPlan_(y, {
        sales: parseAmount_(row[1]),
        variable: parseAmount_(row[2]),
        labor: parseAmount_(row[3]),
        otherFixed: parseAmount_(row[4]),
        depreciation: parseAmount_(row[5]),
        loanPayment: parseAmount_(row[6]),
        capex: parseAmount_(row[7])
      });
    }

    return normalizeAnnualPlan_(y, null);
  } catch (error) {
    console.error('getAnnualPlan error:', error);
    throw new Error('年次計画取得エラー: ' + error.message);
  }
}

/**
 * 年次計画を保存
 * @param {number} year
 * @param {Object} data
 * @return {Object}
 */
function saveAnnualPlan(year, data) {
  try {
    var y = parseInt(year, 10);
    if (!y) throw new Error('年が不正です');

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ANNUAL_PLAN_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(ANNUAL_PLAN_SHEET);
      sheet.appendRow(['年', '売上(計画)', '変動費', '人件費', 'その他固定費', '減価償却費', '借入返済', '設備投資', '更新日時', '更新者']);
    }

    var values = getSheetValues_(sheet);
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (parseInt(values[i][0], 10) === y) {
        rowIndex = i + 1;
        break;
      }
    }

    var rowData = [
      y,
      data.sales || 0,
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
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (error) {
    console.error('saveAnnualPlan error:', error);
    throw new Error('年次計画保存エラー: ' + error.message);
  }
}

/**
 * CF年次表示用データ（実績/計画）
 * @param {number} year
 * @param {string} mode - "actual" | "plan"
 * @return {Object}
 */
function getYearlyCashFlow(year, mode) {
  try {
    mode = mode || 'actual';
    if (mode === 'plan') return getAnnualPlan(year);
    return getYearlyActualCashFlow_(year);
  } catch (error) {
    console.error('getYearlyCashFlow error:', error);
    throw new Error('年次CF取得エラー: ' + error.message);
  }
}
