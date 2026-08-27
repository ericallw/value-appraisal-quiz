/**
 * 值錢鑑定卡 — 後端（已部署 v5+）
 *
 * PDF 排版全部由 buildPdf() 程式生成（顏色跟返網頁「檸檬與亞麻」品牌色），
 * TEMPLATE_DOC_ID 指嘅 Google Doc 只係一個乾淨容器，內文會每次被清空重寫——
 * 唔使再喺個 Doc 度手動貼字/調格式，想改設計就直接改 pdfBox/pdfStatRow 呢幾個函數。
 * SHEET_ID／TEMPLATE_DOC_ID 已經填好，ADMIN_EMAIL 已經填咗 lilokwa122@gmail.com。
 * 前端個 fetch 用 text/plain 避開 CORS preflight，所以呢度用 e.postData.contents 解析。
 */

var SHEET_ID = "173Wr44n3hPjhie8ZeXSBmRyW9TV0pFZWthHJ6yK0yrk";
var TEMPLATE_DOC_ID = "1_7U5jubQMFWQL6QHhRmOvU2PuZLcF7REbcdb60-0TXc";
var ADMIN_EMAIL = "lilokwa122@gmail.com";

var HEADER_ROW = [
  "時間", "Email", "鎖定卡", "分類", "分類名", "類型", "熱度分級",
  "第一輪張數", "第二輪張數", "第三輪張數", "撐得住指數",
  "最卡原因", "最快變現形態", "數碼產品形態"
];

var TYPE_HEAD = {
  hidden: "你不是沒有可以賣的東西，是從來沒有人知道你有。",
  signal: "你手上不只一件值錢的事，你卡在還沒開價。",
  material: "現在還沒有，這是正常的，也是唯一誠實的答案。"
};
var TYPE_ACTION = {
  hidden: "把這件事做一次，公開發布出來，附上「有需要可以找我」。",
  signal: "公開標價，收下第一個陌生人的錢。",
  material: "從第一輪圈選過的卡片裡挑一張，公開記錄 30 天。"
};
var TYPE_QUOTE = {
  hidden: "「沒有需求」和「需求還沒找到你」，是兩件完全不同的事。",
  signal: "專業不是準備好才發生，是收第一次錢之後才開始。",
  material: "方向不是想出來的結論，是做出來的副產品。"
};

// 現在就能做的一步——PDF專屬內容，免費、簡單、今天就做得到。
// 更深嘅內容生產／漏斗設計留返做未來付費內容，唔喺呢度畀晒。
var IMMEDIATE_ACTION = {
  signal: "傳一則訊息給最近一個問過你這件事的人：告訴他你現在可以幫他，附上一個價。",
  hidden: "在限時動態用一句話說你會做這件事——不用多，一句就好。",
  material: "把這件事做一次，拍下或寫下那個過程——不用想怎麼發，先做這一次就好。"
};

/**
 * 熱度分級 —— 用「鑑定類型」+「撐得住指數」判斷呢個人幾接近會畀錢：
 *   hot    = 已經有市場訊號，而且撐得住 → 最值得你今日就親自 DM
 *   warm   = 有嘢但仲未被人發現，而且撐得住 → 中期培育
 *   nurture = 素材仲未夠，或者撐得住指數低 → 長期培育，未急
 */
function computeTier(type, persistScore) {
  if (type === "signal" && persistScore >= 5) return "hot";
  if (type === "hidden" && persistScore >= 5) return "warm";
  return "nurture";
}

function doGet(e) {
  return ContentService.createTextOutput("OK — 值錢鑑定卡後端運行緊");
}

// 喺 editor 揀呢個函數撳「執行」——逐個直接掂一掂 Sheet/Drive/Docs/Gmail，
// 逼 Apps Script 一次過問晒所有授權（唔會真係改動任何嘢）。
function grantAccess() {
  SpreadsheetApp.openById(SHEET_ID).getName();
  DriveApp.getFileById(TEMPLATE_DOC_ID).getName();
  DocumentApp.openById(TEMPLATE_DOC_ID).getName();
  GmailApp.getAliases();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) {
      return jsonOut({ ok: false, error: "invalid email" });
    }

    var tier = computeTier(data.type, data.persistScore);
    appendRow(data, tier);

    var pdfBlob = buildPdf(data);
    sendResultEmail(data.email, pdfBlob, data);

    if (tier === "hot") {
      notifyHotLead(data);
    }

    return jsonOut({ ok: true, tier: tier });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// 同一個 email 只保留一行——搵到就更新嗰行，搵唔到先 append 新行。
// 呢個防埋因為網絡重試/手快撳兩下而出現嘅重複行。
function appendRow(data, tier) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  }
  var row = [
    data.submittedAt || new Date().toISOString(),
    data.email,
    data.lockedCardText,
    data.category,
    data.categoryName,
    data.type,
    tier,
    data.r1Count,
    data.r2Count,
    data.r3Count,
    data.persistScore,
    data.finalChoiceText,
    data.fastForm,
    data.digiForm
  ];

  var existingRowNum = findRowByEmail(sheet, data.email);
  if (existingRowNum) {
    sheet.getRange(existingRowNum, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function findRowByEmail(sheet, email) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var emails = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < emails.length; i++) {
    if (emails[i][0] === email) return i + 2;
  }
  return null;
}

// 顏色跟返網頁「檸檬與亞麻」品牌色
var PDF_INK = "#3A4420";
var PDF_INK_SOFT = "#5C6A3D";
var PDF_BRASS = "#8F6416";
var PDF_LINE = "#D9C9A0";
var PDF_LINEN = "#F3ECDA";

function buildPdf(data) {
  // 範本 Doc 淨係用嚟做一個乾淨嘅容器，實際排版全部由程式生成，
  // 唔使再靠人手喺 Doc 度貼字/調格式——每次都自動保持同一套設計。
  var templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  var copy = templateFile.makeCopy("值錢鑑定證書 - " + data.email);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  // 清空範本原有內容（Body.clear() 會自動保留一個空段落，符合 Docs 嘅結構要求）
  body.clear();

  body.setMarginTop(38).setMarginBottom(38).setMarginLeft(56).setMarginRight(56);

  pdfPara(body.getChild(0).asParagraph(), "夢 想 人 生 研 究 所 · 免 費 鑑 定", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 9, bold: true,
    color: PDF_BRASS, family: "Courier New", spaceAfter: 5
  });

  pdfAppend(body, "你身上哪件事值錢", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 22, bold: true,
    color: PDF_INK, family: "Georgia", spaceAfter: 2
  });
  pdfAppend(body, "鑑 定 證 書", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 11, italic: true,
    color: PDF_BRASS, family: "Georgia", spaceAfter: 12
  });

  body.appendHorizontalRule();

  pdfAppend(body, pdfTypeBadge(data.type), {
    size: 9, bold: true, color: PDF_BRASS, family: "Courier New",
    spaceBefore: 12, spaceAfter: 5
  });
  pdfAppend(body, TYPE_HEAD[data.type] || "", {
    size: 15.5, bold: true, color: PDF_INK, family: "Georgia", spaceAfter: 10
  });

  if (data.type !== "material" && data.lockedCardText) {
    pdfBox(body, "你的第一個產品", data.lockedCardText);
  }
  pdfBox(body, "7 天內做一件事", TYPE_ACTION[data.type] || "");

  pdfAppend(body, "「" + (TYPE_QUOTE[data.type] || "") + "」", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 12, italic: true,
    color: PDF_INK, family: "Georgia", spaceBefore: 6, spaceAfter: 10
  });

  body.appendHorizontalRule();

  pdfStatRow(body, [
    ["第二輪．不費力", String(data.r2Count)],
    ["第三輪．有人道謝過", String(data.r3Count)],
    ["撐得住指數", data.persistScore + " / 8"]
  ]);

  pdfAppend(body, "這張卡，可以怎麼賣", {
    size: 13, bold: true, color: PDF_INK, family: "Georgia",
    spaceBefore: 12, spaceAfter: 2
  });
  pdfAppend(body, "分類：" + (data.categoryName || ""), {
    size: 10.5, color: PDF_INK_SOFT, spaceAfter: 6
  });
  pdfStatRow(body, [
    ["最快變現形態", data.fastForm || ""],
    ["最容易做成數碼產品", data.digiForm || ""]
  ]);

  pdfAppend(body, "你的答案", {
    size: 10, bold: true, color: PDF_BRASS, family: "Courier New",
    spaceBefore: 12, spaceAfter: 4
  });
  pdfAppend(body, data.finalChoiceText || "", {
    size: 12, color: PDF_INK, spaceAfter: 14
  });

  // ---- 以下一節係PDF專屬加碼內容，網頁結果頁完全冇；唔強行分頁，跟返實際內容自然接落去 ----
  body.appendHorizontalRule();

  pdfAppend(body, "PDF 專屬加碼", {
    size: 9, bold: true, color: PDF_BRASS, family: "Courier New",
    spaceBefore: 12, spaceAfter: 10
  });

  pdfAppend(body, "現在就能做的一步", {
    size: 16, bold: true, color: PDF_INK, family: "Georgia", spaceAfter: 4
  });
  pdfAppend(body, "不用等到「準備好」，也不用先想清楚整套計畫——今天就做這一件事。", {
    size: 10.5, color: PDF_INK_SOFT, spaceAfter: 10
  });
  pdfBox(body, "今天", IMMEDIATE_ACTION[data.type] || "");

  pdfAppend(body, "這一步之後，是把它變成持續的內容、完整的銷售漏斗——這部分需要更系統化的協助，屬於夢想人生研究所之後會推出的內容。", {
    size: 10, color: PDF_INK_SOFT, spaceBefore: 8, spaceAfter: 10
  });

  body.appendHorizontalRule();
  pdfAppend(body, "這張卡只是第一步。夢想人生研究所會陸續推出更多工具，幫你一步步研究出自己的方向。", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 9.5, italic: true,
    color: PDF_INK_SOFT, family: "Georgia", spaceBefore: 10, spaceAfter: 10
  });
  pdfAppend(body, "夢想人生研究所 · 你身上哪件事值錢", {
    align: DocumentApp.HorizontalAlignment.CENTER, size: 8.5,
    color: PDF_INK_SOFT, family: "Courier New"
  });

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(copy.getId()).getAs("application/pdf");
  DriveApp.getFileById(copy.getId()).setTrashed(true); // 清走臨時副本，PDF已經攞埋做blob
  return pdfBlob;
}

function pdfTypeBadge(type) {
  if (type === "signal") return "鑑定類型．已經有訊號";
  if (type === "hidden") return "鑑定類型．有東西沒人知道";
  return "鑑定類型．素材還不夠";
}

// 幫段落套用字體/顏色/對齊/前後留白
function pdfPara(p, text, opt) {
  if (text != null) p.setText(text);
  p.setFontFamily(opt.family || "Georgia");
  p.setFontSize(opt.size || 11);
  p.setBold(!!opt.bold);
  p.setItalic(!!opt.italic);
  p.setForegroundColor(opt.color || PDF_INK);
  if (opt.align) p.setAlignment(opt.align);
  p.setSpacingBefore(opt.spaceBefore || 0);
  p.setSpacingAfter(opt.spaceAfter || 0);
  return p;
}

function pdfAppend(body, text, opt) {
  return pdfPara(body.appendParagraph(text), null, opt);
}

// 淺亞麻底色嘅重點方塊，畀「你的第一個產品」/「7 天內做一件事」用
function pdfBox(body, label, value) {
  var table = body.appendTable([[" "]]);
  table.setBorderWidth(0);
  var cell = table.getCell(0, 0);
  cell.setBackgroundColor(PDF_LINEN);
  cell.setPaddingTop(9).setPaddingBottom(9).setPaddingLeft(16).setPaddingRight(16);
  pdfPara(cell.getChild(0).asParagraph(), label, {
    size: 9, bold: true, color: PDF_BRASS, family: "Courier New", spaceAfter: 3
  });
  var valuePara = cell.appendParagraph(value);
  pdfPara(valuePara, null, { size: 13, bold: true, color: PDF_INK });

  var spacer = body.appendParagraph(" ");
  spacer.setSpacingAfter(8);
  return table;
}

// 多欄統計列，用嚟顯示落差數字/變現形態
function pdfStatRow(body, pairs) {
  var cells = pairs.map(function () { return " "; });
  var table = body.appendTable([cells]);
  table.setBorderWidth(0.5);
  table.setBorderColor(PDF_LINE);
  for (var i = 0; i < pairs.length; i++) {
    var cell = table.getCell(0, i);
    cell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
    pdfPara(cell.getChild(0).asParagraph(), pairs[i][0], {
      size: 8.5, bold: true, color: PDF_BRASS, family: "Courier New", spaceAfter: 4
    });
    var v = cell.appendParagraph(pairs[i][1]);
    pdfPara(v, null, { size: 12, bold: true, color: PDF_INK });
  }
  var p = body.appendParagraph(" ");
  p.setSpacingAfter(6);
  return table;
}

function sendResultEmail(email, pdfBlob, data) {
  GmailApp.sendEmail(email, "你的鑑定證書：" + (data.lockedCardText || ""),
    "附件是你的完整鑑定分析 PDF。\n\n如果有問題，直接回覆這封 email。",
    {
      from: ADMIN_EMAIL,
      attachments: [pdfBlob.setName("值錢鑑定證書.pdf")]
    });
}

function notifyHotLead(data) {
  GmailApp.sendEmail(ADMIN_EMAIL, "🔥 熱門名單：" + data.email,
    "有一個「已經有訊號」而且撐得住指數高的人剛剛完成鑑定，值得今天就親自跟進。\n\n" +
    "Email：" + data.email + "\n" +
    "鎖定卡：" + data.lockedCardText + "\n" +
    "分類：" + data.categoryName + "\n" +
    "第二輪／第三輪：" + data.r2Count + " / " + data.r3Count + "\n" +
    "撐得住指數：" + data.persistScore + " / 8\n" +
    "最快變現形態：" + data.fastForm,
    { from: ADMIN_EMAIL });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
