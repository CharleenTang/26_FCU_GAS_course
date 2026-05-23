// ============================================================
// 報價單系統 - 完整版
// 功能：初始化資料、一鍵生成報價單、轉 PDF 建立 Gmail 草稿
// ============================================================


// ────────────────────────────────────────────────────────────
// 選單
// ────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 報價單系統")
    .addItem("📦 初始化報價資料", "初始化報價資料")
    .addItem("📋 生成報價單", "生成報價單")
    .addItem("📤 寄出報價單", "寄出報價單")
    .addToUi();
}


// ────────────────────────────────────────────────────────────
// 初始化報價資料
// ────────────────────────────────────────────────────────────

function 初始化報價資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 產生資料編號（PropertiesService 計數器遞增，跨初始化不重置）
  var props = PropertiesService.getScriptProperties();
  var 目前計數 = parseInt(props.getProperty("資料編號計數器") || "0");
  var 下一個計數 = 目前計數 + 1;
  var 資料編號 = "RD-" + ("000" + 下一個計數).slice(-4);
  props.setProperty("資料編號計數器", 下一個計數.toString());

  // Sheet 名稱直接帶入 RD 編號
  var 新表名稱 = "報價資料 " + 資料編號;
  var sheet = ss.getSheetByName(新表名稱);
  if (!sheet) sheet = ss.insertSheet(新表名稱); else sheet.clear();

  // 品項欄位標題與預設資料
  var 標題 = [["品名/規格", "單位", "說明", "數量", "單價"]];
  var 資料 = [
    ["AI 智慧會議系統（基本版）", "套", "含語音辨識、自動會議紀要", 1, 180000],
    ["智慧文件管理模組", "授權", "OCR + AI 分類，10 人授權", 10, 12000],
    ["自動化報表工具", "授權", "每日/週/月報表自動產生", 5, 8500],
    ["教育訓練", "小時", "現場教育訓練（含教材）", 16, 3000],
    ["系統維護（年約）", "年", "含系統更新與技術支援", 1, 50000],
    ["客製化開發", "人天", "依需求客製功能開發", 10, 8000]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  sheet.getRange("A1:E1").setBackground("#283593").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  // G1/H1 資料編號（唯讀顯示）
  sheet.getRange("G1").setValue("資料編號：").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("H1").setValue(資料編號)
    .setFontWeight("bold").setFontColor("#1a237e")
    .setBackground("#e8eaf6").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "#1a237e", SpreadsheetApp.BorderStyle.SOLID);

  // G2/H2 業務人員（可手動修改）
  sheet.getRange("G2").setValue("業務人員：").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("H2").setValue("林冠廷")
    .setFontWeight("bold").setBackground("#fff3e0")
    .setBorder(true, true, true, true, false, false, "#ff9800", SpreadsheetApp.BorderStyle.SOLID);

  // G3/H3 客戶 Email（可手動填入）
  sheet.getRange("G3").setValue("客戶Email：").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("H3").setValue("")
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, false, false, "#4caf50", SpreadsheetApp.BorderStyle.SOLID);

  sheet.autoResizeColumn(7);
  sheet.setColumnWidth(8, 180);

  // 歷史追蹤表：只有不存在才新建，不清空 → 流水號不中斷
  var 歷史表 = ss.getSheetByName("報價歷史追蹤");
  if (!歷史表) {
    歷史表 = ss.insertSheet("報價歷史追蹤");
    歷史表.appendRow(["報價編號", "報價日期", "客戶名稱", "總金額", "經辦業務", "資料單號", "快速連結"]);
    歷史表.getRange("A1:G1")
      .setBackground("#1a237e").setFontColor("#fff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    歷史表.setFrozenRows(1);
    歷史表.setColumnWidth(1, 160);
    歷史表.setColumnWidth(2, 110);
    歷史表.setColumnWidth(3, 180);
    歷史表.setColumnWidth(4, 120);
    歷史表.setColumnWidth(5, 100);
    歷史表.setColumnWidth(6, 200);
    歷史表.setColumnWidth(7, 120);
  }

  SpreadsheetApp.getUi().alert(
    "✅ 報價資料已初始化！\n\n" +
    "📋 資料編號：" + 資料編號 + "\n" +
    "📄 工作表名稱：" + 新表名稱 + "\n" +
    "👤 H2 可修改業務姓名\n" +
    "📧 H3 請填入客戶 Email"
  );
}


// ────────────────────────────────────────────────────────────
// 生成報價單
// ────────────────────────────────────────────────────────────

function 生成報價單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 動態找名稱以「報價資料」開頭的 sheet
    var 客戶表 = 找報價資料表(ss);
    if (!客戶表) { SpreadsheetApp.getUi().alert("❌ 請先初始化報價資料"); return; }

    var 業務人員 = 客戶表.getRange("H2").getValue().toString().trim() || "林冠廷";
    var 資料編號 = 客戶表.getRange("H1").getValue().toString().trim() || "";

    var ui = SpreadsheetApp.getUi();
    var 客戶回應 = ui.prompt("📝 報價單", "請輸入客戶名稱：", ui.ButtonSet.OK_CANCEL);
    if (客戶回應.getSelectedButton() !== ui.Button.OK) return;
    var 客戶 = 客戶回應.getResponseText().trim() || "範例客戶";

    var 今日序號 = 取得今日下一個流水號(ss);
    var 編號 = "QT-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd") + "-" + 今日序號;

    var sheet = ss.insertSheet(編號, 0);

    // ── 公司 Logo 區 ──
    sheet.getRange("A1:F1").merge();
    sheet.getRange("A1").setValue("ABC 科技股份有限公司")
      .setFontSize(20).setFontWeight("bold").setFontColor("#1a237e");
    sheet.setRowHeight(1, 50);

    sheet.getRange("A2:F2").merge();
    sheet.getRange("A2").setValue("台北市信義區信義路五段 7 號 ｜ Tel: 02-2345-6789 ｜ www.abc-tech.com")
      .setFontSize(9).setFontColor("#666");

    // ── 報價單標題 ──
    sheet.getRange("A4:F4").merge();
    sheet.getRange("A4").setValue("📋 報 價 單")
      .setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground("#1a237e").setFontColor("#fff");
    sheet.setRowHeight(4, 45);

    // ── 客戶資訊 ──
    var 今天 = new Date();
    var 有效日 = new Date(今天);
    有效日.setDate(今天.getDate() + 30);

    sheet.getRange(6, 1, 4, 6).setValues([
      ["報價編號", 編號,                                                         "", "客戶名稱", 客戶, ""],
      ["報價日期", Utilities.formatDate(今天,   "Asia/Taipei", "yyyy/MM/dd"),    "", "聯絡人",   "",   ""],
      ["有效期限", Utilities.formatDate(有效日, "Asia/Taipei", "yyyy/MM/dd"),    "", "聯絡電話", "",   ""],
      ["業務人員", 業務人員,                                                     "", "傳真",     "",   ""]
    ]);

    for (var r = 6; r <= 9; r++) {
      sheet.getRange(r, 1).setFontWeight("bold").setBackground("#e8eaf6");
      sheet.getRange(r, 4).setFontWeight("bold").setBackground("#e8eaf6");
    }
    sheet.getRange("B6:B9").setHorizontalAlignment("left");

    // ── 品項表格標題 ──
    sheet.getRange("A11:F11").setValues([["項次", "品名/規格", "單位", "數量", "單價", "金額"]]);
    sheet.getRange("A11:F11")
      .setBackground("#283593").setFontColor("#fff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setRowHeight(11, 35);

    // ── 品項資料 ──
    var 品項資料 = 客戶表.getDataRange().getValues();
    var 小計 = 0;

    for (var i = 1; i < 品項資料.length; i++) {
      var 列 = 11 + i;
      var 金額 = 品項資料[i][3] * 品項資料[i][4];
      小計 += 金額;

      sheet.getRange(列, 1, 1, 6).setValues([[
        i, 品項資料[i][0], 品項資料[i][1], 品項資料[i][3], 品項資料[i][4], 金額
      ]]);
      sheet.getRange(列, 1).setHorizontalAlignment("center");
      sheet.getRange(列, 4, 1, 3).setNumberFormat("#,##0").setHorizontalAlignment("right");

      // 斑馬紋
      if (i % 2 === 0) sheet.getRange(列, 1, 1, 6).setBackground("#f5f5f5");
    }

    // ── 合計區 ──
    var 稅金 = Math.round(小計 * 0.05);
    var 總計 = 小計 + 稅金;
    var 合計起始 = 12 + 品項資料.length - 1;

    sheet.getRange(合計起始, 1, 3, 6).setValues([
      ["", "", "", "", "小　計", 小計],
      ["", "", "", "", "稅金 (5%)", 稅金],
      ["", "", "", "", "總　計", 總計]
    ]);
    sheet.getRange(合計起始, 5, 3, 1).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(合計起始, 6, 3, 1).setNumberFormat("NT$#,##0").setFontWeight("bold");
    sheet.getRange(合計起始 + 2, 5, 1, 2)
      .setBackground("#e8eaf6").setFontSize(14)
      .setBorder(true, true, true, true, false, false, "#1a237e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // ── 備註 ──
    var 備註列 = 合計起始 + 5;
    sheet.getRange(備註列, 1, 1, 6).merge();
    sheet.getRange(備註列, 1).setValue("📌 備註：").setFontWeight("bold");
    sheet.getRange(備註列 + 1, 1, 1, 6).merge();
    sheet.getRange(備註列 + 1, 1).setValue(
      "1. 報價有效期限 30 天\n2. 付款條件：月結 30 天\n3. 交貨期：訂購後 7~14 個工作天\n4. 以上報價含安裝及教育訓練"
    ).setWrap(true).setFontColor("#555");

    // ── 欄寬 ──
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 60);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 120);

    // ── 整體框線 ──
    sheet.getRange(11, 1, 品項資料.length + 3, 6)
      .setBorder(true, true, true, true, true, true, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setFrozenRows(0);

    記錄至歷史追蹤(編號, 今天, 客戶, 總計, 業務人員, 資料編號);

    SpreadsheetApp.getUi().alert(
      "✅ 報價單 " + 編號 + " 已生成！\n總金額：NT$ " + 總計.toLocaleString()
    );

  } catch (錯誤) {
    Logger.log("❌ " + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ " + 錯誤.message);
  }
}


// ────────────────────────────────────────────────────────────
// 寄出報價單（轉 PDF → 建立 Gmail 草稿）
// ────────────────────────────────────────────────────────────

function 寄出報價單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    // ── Step 1：列出所有 QT- 報價單讓使用者選擇 ──
    var sheets = ss.getSheets();
    var 報價單清單 = sheets
      .map(function(s) { return s.getName(); })
      .filter(function(n) { return n.indexOf("QT-") === 0; });

    if (報價單清單.length === 0) {
      ui.alert("❌ 目前沒有任何報價單，請先生成報價單。");
      return;
    }

    var 回應 = ui.prompt(
      "📤 寄出報價單",
      "請輸入要寄出的報價單編號：\n\n" + 報價單清單.join("\n"),
      ui.ButtonSet.OK_CANCEL
    );
    if (回應.getSelectedButton() !== ui.Button.OK) return;

    var 選擇編號 = 回應.getResponseText().trim();
    var 目標Sheet = ss.getSheetByName(選擇編號);
    if (!目標Sheet) {
      ui.alert("❌ 找不到報價單「" + 選擇編號 + "」，請確認編號是否正確。");
      return;
    }

    // ── Step 2：從報價資料表讀取 Email 資訊 ──
    var 資料表 = 找報價資料表(ss);
    if (!資料表) {
      ui.alert("❌ 找不到報價資料表，請先初始化。");
      return;
    }

    var 客戶Email = 資料表.getRange("H3").getValue().toString().trim();
    var 業務人員 = 資料表.getRange("H2").getValue().toString().trim() || "業務人員";

    if (!客戶Email) {
      ui.alert("❌ 報價資料表 H3 尚未填入客戶 Email，請先填寫後再寄出。");
      return;
    }

    // 業務人員 Email 自動抓當前登入 Google 帳號
    var 業務Email = Session.getActiveUser().getEmail();

    // ── Step 3：匯出 PDF ──
    var pdf = 產生報價單PDF(ss, 目標Sheet);

    // ── Step 4：建立 Gmail 草稿 ──
    var 客戶名稱 = 目標Sheet.getRange("E6").getValue().toString().trim() || "客戶";
    var 主旨 = "【報價單】" + 選擇編號 + "　" + 客戶名稱;
    var 附件名稱 = 選擇編號 + ".pdf";

    GmailApp.createDraft(
      客戶Email,
      主旨,
      "",
      {
        htmlBody: 產生Email內文(選擇編號, 客戶名稱, 業務人員),
        cc: 業務Email,
        attachments: [pdf.setName(附件名稱)],
        name: "ABC 科技股份有限公司"
      }
    );

    ui.alert(
      "✅ Gmail 草稿已建立！\n\n" +
      "📧 收件人：" + 客戶Email + "\n" +
      "📋 副本：" + 業務Email + "\n" +
      "📎 附件：" + 附件名稱 + "\n\n" +
      "請至 Gmail 草稿匣確認內容後再送出。"
    );

  } catch (錯誤) {
    Logger.log("❌ " + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ " + 錯誤.message);
  }
}

/**
 * 將指定 Sheet 匯出為 PDF Blob
 */
function 產生報價單PDF(ss, sheet) {
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() +
    "/export?" +
    "exportFormat=pdf" +
    "&format=pdf" +
    "&size=A4" +
    "&portrait=true" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&gid=" + sheet.getSheetId();

  var response = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PDF 匯出失敗，HTTP " + response.getResponseCode());
  }

  return response.getBlob().setContentType("application/pdf");
}

/**
 * 產生 Email HTML 內文
 */
function 產生Email內文(編號, 客戶名稱, 業務人員) {
  return [
    "<p>親愛的 " + 客戶名稱 + " 您好，</p>",
    "<p>感謝您對本公司產品及服務的支持與信賴。</p>",
    "<p>隨信附上報價單 <strong>" + 編號 + "</strong>，請惠予參閱。",
    "如有任何疑問或需要進一步說明，歡迎隨時與我們聯繫。</p>",
    "<p>期待與您合作，謝謝！</p>",
    "<br>",
    "<p>敬上<br>",
    "<strong>" + 業務人員 + "</strong><br>",
    "ABC 科技股份有限公司<br>",
    "Tel: 02-2345-6789<br>",
    "www.abc-tech.com</p>"
  ].join("\n");
}


// ────────────────────────────────────────────────────────────
// 共用工具函式
// ────────────────────────────────────────────────────────────

/**
 * 找名稱以「報價資料」開頭的工作表（相容帶 RD 編號的命名格式）
 */
function 找報價資料表(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf("報價資料") === 0) return sheets[i];
  }
  return null;
}

/**
 * 從「報價歷史追蹤」計算今日已生成的報價單數量，回傳下一個四位數流水號
 */
function 取得今日下一個流水號(ss) {
  var 歷史表 = ss.getSheetByName("報價歷史追蹤");
  if (!歷史表) return "0001";

  var 今天字串 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
  var 資料 = 歷史表.getDataRange().getValues();
  var 今天計數 = 0;

  for (var i = 1; i < 資料.length; i++) {
    var 記錄日期 = 資料[i][1];
    var 記錄日期字串 = (記錄日期 instanceof Date)
      ? Utilities.formatDate(記錄日期, "Asia/Taipei", "yyyy/MM/dd")
      : String(記錄日期).trim();
    if (記錄日期字串 === 今天字串) 今天計數++;
  }

  return ("000" + (今天計數 + 1)).slice(-4);
}

/**
 * 將生成的報價單資訊記錄至「報價歷史追蹤」工作表
 */
function 記錄至歷史追蹤(編號, 日期, 客戶, 總計, 業務人員, 資料單號) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 歷史表 = ss.getSheetByName("報價歷史追蹤");

  if (!歷史表) {
    歷史表 = ss.insertSheet("報價歷史追蹤");
    歷史表.appendRow(["報價編號", "報價日期", "客戶名稱", "總金額", "經辦業務", "資料單號", "快速連結"]);
    歷史表.getRange("A1:G1")
      .setBackground("#1a237e").setFontColor("#fff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    歷史表.setFrozenRows(1);
    歷史表.setColumnWidth(1, 160);
    歷史表.setColumnWidth(2, 110);
    歷史表.setColumnWidth(3, 180);
    歷史表.setColumnWidth(4, 120);
    歷史表.setColumnWidth(5, 100);
    歷史表.setColumnWidth(6, 200);
    歷史表.setColumnWidth(7, 120);
  }

  var targetSheet = ss.getSheetByName(編號);
  var 連結公式 = targetSheet
    ? '=HYPERLINK("#gid=' + targetSheet.getSheetId() + '", "🔍 開啟報價單")'
    : "（工作表不存在）";

  歷史表.appendRow([
    編號,
    Utilities.formatDate(日期, "Asia/Taipei", "yyyy/MM/dd"),
    客戶,
    總計,
    業務人員 || "未填寫",
    資料單號 || "",
    連結公式
  ]);

  var 最後列 = 歷史表.getLastRow();
  歷史表.getRange(最後列, 4).setNumberFormat("NT$#,##0").setHorizontalAlignment("right");
  歷史表.getRange(最後列, 6).setHorizontalAlignment("center");
  歷史表.getRange(最後列, 7).setHorizontalAlignment("center");
}