// ============================================================
// Session 3：SpreadsheetApp 基本操作
// 日期：115/05/09（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Apps Script 介面深入
//   2. SpreadsheetApp 基本操作
//   3. 建立／開啟試算表
//   4. 觸發器應用（自動開啟／時間觸發）
//   5. 實作：自動建立新工作表
// ============================================================

// ============================================================
// 第一部分：SpreadsheetApp 基本操作
// ============================================================

/**
 * SpreadsheetApp 常用方法總覽
 * 說明：示範如何操作試算表、工作表、儲存格
 */
function SpreadsheetApp基本操作() {
  // --- 取得試算表 ---

  // 方法 1：取得「目前開啟」的試算表（最常用）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("試算表名稱：" + ss.getName());
  Logger.log("試算表 ID：" + ss.getId());
  Logger.log("試算表 URL：" + ss.getUrl());

  // --- 取得工作表 (Sheet) ---

  // 方法 1：依名稱取得
  var sheet = ss.getSheetByName("員工資料");
  if (sheet) {
    Logger.log("找到工作表：" + sheet.getName());
  } else {
    Logger.log("找不到「員工資料」工作表");
  }

  // 方法 2：取得所有工作表
  var 所有工作表 = ss.getSheets();
  Logger.log("工作表數量：" + 所有工作表.length);
  for (var i = 0; i < 所有工作表.length; i++) {
    Logger.log("  - " + 所有工作表[i].getName());
  }

  // 方法 3：取得目前作用中的工作表
  var 目前工作表 = ss.getActiveSheet();
  Logger.log("目前工作表：" + 目前工作表.getName());
}

/**
 * 工作表基本資訊讀取
 */
function 讀取工作表資訊() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("員工資料");

  if (!sheet) {
    Logger.log("❌ 找不到「員工資料」，請先執行初始化");
    return;
  }

  // 取得工作表的基本資訊
  Logger.log("===== 工作表資訊 =====");
  Logger.log("最後一列（有資料）：" + sheet.getLastRow());
  Logger.log("最後一欄（有資料）：" + sheet.getLastColumn());
  Logger.log("列數上限：" + sheet.getMaxRows());
  Logger.log("欄數上限：" + sheet.getMaxColumns());

  // 取得所有有資料的範圍
  var 資料範圍 = sheet.getDataRange();
  Logger.log("資料範圍：" + 資料範圍.getA1Notation()); // e.g. "A1:F11"
}

// ============================================================
// 第二部分：建立與管理工作表
// ============================================================

/**
 * 示範建立新工作表的各種方式
 */
function 建立工作表示範() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 方式 1：建立空白工作表（自動命名）
    var 新表1 = ss.insertSheet();
    Logger.log("建立新表：" + 新表1.getName());

    // 方式 2：指定名稱建立
    var 表名 = "測試工作表_" + new Date().getTime();
    var 新表2 = ss.insertSheet(表名);
    Logger.log("建立新表：" + 新表2.getName());

    // 方式 3：在指定位置插入（第 1 個位置）
    var 新表3 = ss.insertSheet("置頂工作表", 0);
    Logger.log("建立新表（置頂）：" + 新表3.getName());

    // 清理：刪除剛建立的測試工作表
    ss.deleteSheet(新表1);
    ss.deleteSheet(新表2);
    ss.deleteSheet(新表3);
    Logger.log("✅ 測試工作表已清理");
  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 自動建立每月報表工作表
 * 說明：根據目前月份，自動建立當月的工作表（含標題與格式）
 */
function 自動建立月報表() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 員工表 = ss.getSheetByName("員工資料");
    
    if (!員工表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「員工資料」工作表，請先點選「初始化員工資料」。");
      return;
    }

    // 1. 取得員工基本資料
    var 最後一列 = 員工表.getLastRow();
    if (最後一列 < 2) {
      SpreadsheetApp.getUi().alert("⚠️ 員工資料中沒有數據。");
      return;
    }
    // 取得資料 (姓名、部門、月薪分別在第 1, 2, 5 欄)
    var 員工原始資料 = 員工表.getRange(2, 1, 最後一列 - 1, 5).getValues();

    // 2. 準備新工作表
    var 今天 = new Date();
    var 年 = 今天.getFullYear();
    var 月 = 今天.getMonth() + 1;
    var 表名 = 年 + "年" + 月 + "月薪資報表";

    // 檢查工作表是否已存在
    var 既有表 = ss.getSheetByName(表名);
    if (既有表) {
      SpreadsheetApp.getUi().alert("⚠️ 「" + 表名 + "」已存在，請先刪除舊表或更名。");
      return;
    }

    var 新表 = ss.insertSheet(表名);

    // 3. 設定標題列與格式
    var 標題 = [["姓名", "部門", "基本月薪", "加班津貼", "應領金額"]];
    新表.getRange("A1:E1").setValues(標題);
    
    var 標題範圍 = 新表.getRange("A1:E1");
    標題範圍.setBackground("#34a853").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    // 4. 準備並寫入報表內容
    var 報表內容 = [];
    for (var i = 0; i < 員工原始資料.length; i++) {
      var 姓名 = 員工原始資料[i][0];
      var 部門 = 員工原始資料[i][1];
      var 月薪 = 員工原始資料[i][4];
      
      // 示範：隨機產生一些加班津貼
      var 加班費 = Math.floor(Math.random() * 3000) + 1000;
      var 總計 = 月薪 + 加班費;
      
      報表內容.push([姓名, 部門, 月薪, 加班費, 總計]);
    }

    // 寫入所有資料
    新表.getRange(2, 1, 報表內容.length, 5).setValues(報表內容);

    // 5. 最後修飾
    新表.getRange(2, 3, 報表內容.length, 3).setNumberFormat("#,##0"); // 格式化金額
    新表.setFrozenRows(1); // 凍結首列
    
    for (var j = 1; j <= 5; j++) {
      新表.autoResizeColumn(j);
      var 目前寬度 = 新表.getColumnWidth(j);
      新表.setColumnWidth(j, 目前寬度 + 30); // 增加 30 像素緩衝
    }

    Logger.log("✅ 「" + 表名 + "」建立完成，共計 " + 報表內容.length + " 位員工。");
    SpreadsheetApp.getUi().alert("✅ 「" + 表名 + "」建立完成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 自動建立週報表
 * 說明：根據目前週數，自動建立該週的工作報表（含標題與隨機工時）
 */
function 自動建立週報表() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 員工表 = ss.getSheetByName("員工資料");
    
    if (!員工表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「員工資料」工作表，請先執行初始化。");
      return;
    }

    // 1. 取得員工基本資料 (姓名、部門)
    var 最後一列 = 員工表.getLastRow();
    if (最後一列 < 2) {
      SpreadsheetApp.getUi().alert("⚠️ 員工資料中沒有數據。");
      return;
    }
    var 員工原始資料 = 員工表.getRange(2, 1, 最後一列 - 1, 2).getValues();

    // 2. 計算目前週數、日期區間與準備表名
    var 今天 = new Date();
    var 一月一日 = new Date(今天.getFullYear(), 0, 1);
    var 經過天數 = Math.floor((今天 - 一月一日) / (24 * 60 * 60 * 1000));
    var 週數 = Math.ceil((經過天數 + 一月一日.getDay() + 1) / 7);
    var 表名 = 今天.getFullYear() + "年第" + 週數 + "週工作週報";

    // 計算日期區間 (以本週日到週六為準)
    var 距離週日 = -今天.getDay();
    var 週日 = new Date(今天);
    週日.setDate(今天.getDate() + 距離週日);
    var 週六 = new Date(週日);
    週六.setDate(週日.getDate() + 6);
    
    var 格式化日期 = function(d) { 
      var 星期幾 = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
      return Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd") + " (" + 星期幾[d.getDay()] + ")"; 
    };
    var 日期區間文字 = "📅 日期區間：" + 格式化日期(週日) + " ~ " + 格式化日期(週六);

    // 檢查工作表是否已存在
    var 既有表 = ss.getSheetByName(表名);
    if (既有表) {
      SpreadsheetApp.getUi().alert("⚠️ 「" + 表名 + "」已存在，請先刪除舊表或更名。");
      return;
    }

    var 新表 = ss.insertSheet(表名);

    // 3. 設定日期區間、標題列與格式
    var 標題 = [["姓名", "部門", "週日", "週一", "週二", "週三", "週四", "週五", "週六", "本週總時數"]];
    
    // 寫入日期區間 (第 1 列)
    新表.getRange(1, 1).setValue(日期區間文字);
    新表.getRange(1, 1, 1, 10).merge()
      .setBackground("#f3f3f3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // 寫入欄位標題 (第 2 列)
    新表.getRange(2, 1, 1, 10).setValues(標題);
    新表.getRange(2, 1, 1, 10)
      .setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    // 4. 準備並寫入週報表內容
    var 報表內容 = [];
    for (var i = 0; i < 員工原始資料.length; i++) {
      var 姓名 = 員工原始資料[i][0];
      var 部門 = 員工原始資料[i][1];
      
      // 隨機產生週一至週日的工時 (0~8 小時)
      var 每日工時 = [];
      var 總時數 = 0;
      for (var d = 0; d < 7; d++) {
        var 時數 = Math.floor(Math.random() * 9); 
        每日工時.push(時數);
        總時數 += 時數;
      }
      
      var 每一列 = [姓名, 部門].concat(每日工時).concat([總時數]);
      報表內容.push(每一列);
    }

    // 寫入所有資料 (從第 3 列開始)
    新表.getRange(3, 1, 報表內容.length, 10).setValues(報表內容);

    // 5. 最後修飾
    新表.setFrozenRows(2); // 凍結前兩列 (日期區間與標題)
    for (var j = 1; j <= 10; j++) {
      新表.autoResizeColumn(j);
      var 目前寬度 = 新表.getColumnWidth(j);
      新表.setColumnWidth(j, 目前寬度 + 20);
    }

    Logger.log("✅ 「" + 表名 + "」建立完成。");
    SpreadsheetApp.getUi().alert("✅ 「" + 表名 + "」建立完成！");

    // --- 新增：同步更新月度彙整 ---
    更新月度加班彙整(報表內容, "W" + 週數);

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 更新月度加班彙整
 * 說明：將週報表的工時資料自動彙整到當月的總表中
 * @param {Array} 報表內容 週報表的資料陣列
 * @param {string} 週數文字 週數代碼，例如 "W19"
 */
function 更新月度加班彙整(報表內容, 週數文字) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 今天 = new Date();
    var 月份表名 = 今天.getFullYear() + "年" + (今天.getMonth() + 1) + "月加班工時彙整";
    
    var 月總表 = ss.getSheetByName(月份表名);
    
    // 1. 如果月總表不存在，初始化它
    if (!月總表) {
      月總表 = ss.insertSheet(月份表名);
      var 員工表 = ss.getSheetByName("員工資料");
      if (!員工表) return;
      
      var 最後一列 = 員工表.getLastRow();
      if (最後一列 < 2) return;
      
      var 員工名單 = 員工表.getRange(2, 1, 最後一列 - 1, 2).getValues();
      
      // 設定初始標題與名單
      月總表.getRange(1, 1, 1, 2).setValues([["姓名", "部門"]]);
      月總表.getRange(2, 1, 員工名單.length, 2).setValues(員工名單);
      
      // 格式化
      月總表.getRange("A1:B1").setBackground("#fbbc04").setFontWeight("bold").setHorizontalAlignment("center");
      月總表.setFrozenRows(1);
      月總表.setFrozenColumns(2);
    }
    
    // 2. 尋找或新增週數標題 (例如 W19)
    var 最後一欄 = 月總表.getLastColumn();
    var 標題行 = 月總表.getRange(1, 1, 1, 最後一欄).getValues()[0];
    var 週數欄位 = 標題行.indexOf(週數文字) + 1;
    
    if (週數欄位 === 0) {
      週數欄位 = 最後一欄 + 1;
      月總表.getRange(1, 週數欄位).setValue(週數文字);
      月總表.getRange(1, 週數欄位).setBackground("#fbbc04").setFontWeight("bold").setHorizontalAlignment("center");
    }
    
    // 3. 建立姓名到列號的映射
    var 姓名資料 = 月總表.getRange(1, 1, 月總表.getLastRow(), 1).getValues();
    var 姓名映射 = {};
    for (var r = 1; r < 姓名資料.length; r++) {
      if (姓名資料[r][0]) {
        姓名映射[姓名資料[r][0]] = r + 1;
      }
    }
    
    // 4. 將週報資料寫入對應的週數欄位
    for (var i = 0; i < 報表內容.length; i++) {
      var 姓名 = 報表內容[i][0];
      var 總工時 = 報表內容[i][9]; // 週報表中「本週總時數」在第 10 欄 (index 9)
      
      var 目標列 = 姓名映射[姓名];
      if (目標列) {
        月總表.getRange(目標列, 週數欄位).setValue(總工時);
      }
    }
    
    月總表.autoResizeColumn(週數欄位);
    Logger.log("✅ 月度加班彙整更新成功：" + 週數文字);

  } catch (錯誤) {
    Logger.log("❌ 更新月度彙整失敗：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：讀寫儲存格
// ============================================================

/**
 * 示範讀取與寫入儲存格
 */
function 讀寫儲存格示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("員工資料");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到「員工資料」");
    return;
  }

  // --- 讀取單一儲存格 ---
  var A1值 = sheet.getRange("A1").getValue();
  Logger.log("A1 的值：" + A1值);

  // --- 讀取範圍 ---
  var 範圍值 = sheet.getRange("A1:C3").getValues();
  Logger.log("A1:C3 的值：" + JSON.stringify(範圍值));

  // --- 寫入 11 筆資料（一次寫入多欄多列） ---
  var 現在時間 = "2026/5/9"; // Hardcode 固定時間示範
  
  var 示範資料 = [
    ["更新時間", "加班時數"],   // 第 1 筆：標題
    [現在時間, 8],              // 第 2 筆
    [現在時間, 4],              // 第 3 筆
    [現在時間, 12],             // 第 4 筆
    [現在時間, 0],              // 第 5 筆
    [現在時間, 6],              // 第 6 筆
    [現在時間, 10],             // 第 7 筆
    [現在時間, 2],              // 第 8 筆
    [現在時間, 5],              // 第 9 筆
    [現在時間, 9],              // 第 10 筆
    [現在時間, 3]               // 第 11 筆
  ];

  // 使用 setValues 一次寫入 11 列、2 欄 (範圍從 H1 開始，即第 1 列、第 8 欄)
  sheet.getRange(1, 8, 11, 2).setValues(示範資料);
  
  // 自動調整欄寬以符合時間字串長度
  // 自動調整欄寬並增加緩衝空間
  sheet.autoResizeColumn(8);
  var 欄寬8 = sheet.getColumnWidth(8);
  sheet.setColumnWidth(8, 欄寬8 + 30);

  Logger.log("✅ 讀寫示範完成，已寫入 11 筆資料（包含當前格式化時間：" + 現在時間 + "）");
}

// ============================================================
// 第四部分：觸發器進階 — 時間觸發
// ============================================================

/**
 * 手動設定時間觸發器
 * 說明：每天早上 9 點自動執行「每日報告」函數
 */
function 設定每日觸發器() {
  // 先刪除既有的相同觸發器（避免重複）
  var 觸發器列表 = ScriptApp.getProjectTriggers();
  for (var i = 0; i < 觸發器列表.length; i++) {
    if (觸發器列表[i].getHandlerFunction() === "每日自動報告") {
      ScriptApp.deleteTrigger(觸發器列表[i]);
      Logger.log("已刪除舊的觸發器");
    }
  }

  // 建立新的時間觸發器：每天 9~10 點之間執行
  ScriptApp.newTrigger("每日自動報告")
    .timeBased()
    .everyDays(1) // 每天
    .atHour(9) // 早上 9 點
    .create();

  Logger.log("✅ 每日觸發器已設定（9:00~10:00）");
  SpreadsheetApp.getUi().alert(
    "✅ 每日觸發器已設定！\n每天 9:00~10:00 會自動執行。",
  );
}

/**
 * 被觸發器呼叫的函數
 * 說明：自動記錄每日報告資訊到「每日紀錄」工作表
 */
function 每日自動報告() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("每日紀錄");

    // 如果「每日紀錄」不存在，自動建立
    if (!sheet) {
      sheet = ss.insertSheet("每日紀錄");
      sheet.getRange("A1:D1").setValues([["日期", "時間", "事件", "狀態"]]);
      sheet.getRange("A1:D1").setBackground("#fbbc04").setFontWeight("bold");
    }

    // 寫入紀錄
    var 新列 = sheet.getLastRow() + 1;
    var 現在 = new Date();

    sheet
      .getRange(新列, 1)
      .setValue(Utilities.formatDate(現在, "Asia/Taipei", "yyyy/MM/dd"));
    sheet
      .getRange(新列, 2)
      .setValue(Utilities.formatDate(現在, "Asia/Taipei", "HH:mm:ss"));
    sheet.getRange(新列, 3).setValue("每日自動報告已執行");
    sheet.getRange(新列, 4).setValue("✅ 正常");

    Logger.log("✅ 每日報告已記錄：" + 現在);
  } catch (錯誤) {
    Logger.log("❌ 每日報告錯誤：" + 錯誤.message);
  }
}

/**
 * 刪除所有觸發器
 * 說明：方便清理測試時建立的觸發器
 */
function 刪除所有觸發器() {
  var 觸發器列表 = ScriptApp.getProjectTriggers();
  for (var i = 0; i < 觸發器列表.length; i++) {
    ScriptApp.deleteTrigger(觸發器列表[i]);
  }
  Logger.log("✅ 已刪除 " + 觸發器列表.length + " 個觸發器");
  SpreadsheetApp.getUi().alert("✅ 已刪除 " + 觸發器列表.length + " 個觸發器");
}

// ============================================================
// 初始化範例資料
// ============================================================

/**
 * 建立「員工資料」工作表與範例資料
 */
function 初始化員工資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("員工資料");

  if (!sheet) {
    sheet = ss.insertSheet("員工資料");
  } else {
    sheet.clear();
  }

  var 標題 = [["姓名", "部門", "職稱", "到職日", "月薪", "電話", "Email"]];
  var 資料 = [
    [
      "王小明",
      "業務部",
      "業務專員",
      "2023/03/15",
      38000,
      "0912-345-678",
      "wang@example.com",
    ],
    [
      "李小華",
      "行銷部",
      "行銷主管",
      "2021/08/01",
      52000,
      "0923-456-789",
      "lee@example.com",
    ],
    [
      "張美玲",
      "人資部",
      "人資專員",
      "2022/11/20",
      40000,
      "0934-567-890",
      "chang@example.com",
    ],
    [
      "陳大文",
      "研發部",
      "工程師",
      "2024/01/10",
      55000,
      "0945-678-901",
      "chen@example.com",
    ],
    [
      "林小芬",
      "財務部",
      "會計",
      "2020/06/15",
      42000,
      "0956-789-012",
      "lin@example.com",
    ],
    [
      "黃志偉",
      "業務部",
      "業務主管",
      "2019/04/01",
      58000,
      "0967-890-123",
      "huang@example.com",
    ],
    [
      "劉家豪",
      "研發部",
      "資深工程師",
      "2018/09/10",
      68000,
      "0978-901-234",
      "liu@example.com",
    ],
    [
      "吳雅琪",
      "行銷部",
      "行銷專員",
      "2023/07/20",
      36000,
      "0989-012-345",
      "wu@example.com",
    ],
    [
      "周建國",
      "業務部",
      "業務專員",
      "2024/03/01",
      35000,
      "0911-123-456",
      "chou@example.com",
    ],
    [
      "許文馨",
      "人資部",
      "人資主管",
      "2020/01/15",
      56000,
      "0922-234-567",
      "hsu@example.com",
    ],
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);

  // 格式化
  var 標題範圍 = sheet.getRange("A1:G1");
  標題範圍.setBackground("#4285f4");
  標題範圍.setFontColor("#ffffff");
  標題範圍.setFontWeight("bold");
  標題範圍.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  // 薪資欄格式
  sheet.getRange("E2:E11").setNumberFormat("#,##0");

  for (var i = 1; i <= 7; i++) {
    sheet.autoResizeColumn(i);
    var 目前寬度 = sheet.getColumnWidth(i);
    sheet.setColumnWidth(i, 目前寬度 + 30); // 增加 30 像素緩衝
  }

  Logger.log("✅ 員工資料已建立！");
  SpreadsheetApp.getUi().alert("✅ 員工資料已建立！共 " + 資料.length + " 筆");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 3 工具")
    .addItem("👥 初始化員工資料", "初始化員工資料")
    .addItem("📋 SpreadsheetApp 基本操作", "SpreadsheetApp基本操作")
    .addItem("📊 讀取工作表資訊", "讀取工作表資訊")
    .addItem("📝 讀寫儲存格示範", "讀寫儲存格示範")
    .addSeparator()
    .addItem("📅 建立當月報表", "自動建立月報表")
    .addItem("📅 建立週報表", "自動建立週報表")
    .addItem("⏰ 設定每日觸發器", "設定每日觸發器")
    .addItem("🗑️ 刪除所有觸發器", "刪除所有觸發器")
    .addToUi();
}
