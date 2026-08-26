/**
 * 值錢鑑定卡 — 後端草稿（未部署，未接通任何真實 Sheet/Doc）
 *
 * 落手用之前要做：
 * 1. 開一個新 Google Sheet，複製佢個 ID 填入 SHEET_ID，第一分頁加返 HEADER_ROW 嗰行做標題。
 * 2. 開一個新 Google Doc 做 PDF 範本，內文用 {{headline}} {{product}} {{action}} {{quote}}
 *    {{categoryName}} {{fastForm}} {{digiForm}} {{r2Count}} {{r3Count}} {{persistScore}} 呢啲
 *    佔位字（同下面 TEMPLATE_TAGS 對照），複製個 Doc ID 填入 TEMPLATE_DOC_ID。
 * 3. clasp push 之後，Deploy → New deployment → Web app，Execute as "Me"，
 *    Who has access "Anyone" —— 攞到嗰條 /exec URL 填返去前端 index.html 嘅 APP_SCRIPT_URL。
 * 4. 前端個 fetch 用 text/plain 避開 CORS preflight，所以呢度用 e.postData.contents 解析。
 * 5. ADMIN_EMAIL 已經填咗你個email，「熱門」名單一出現就會寄一封通知畀你自己。
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

function appendRow(data, tier) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  }
  sheet.appendRow([
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
  ]);
}

function buildPdf(data) {
  var templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  var copy = templateFile.makeCopy("值錢鑑定證書 - " + data.email);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  var tags = {
    "{{headline}}": TYPE_HEAD[data.type] || "",
    "{{product}}": data.lockedCardText || "",
    "{{action}}": TYPE_ACTION[data.type] || "",
    "{{quote}}": TYPE_QUOTE[data.type] || "",
    "{{categoryName}}": data.categoryName || "",
    "{{fastForm}}": data.fastForm || "",
    "{{digiForm}}": data.digiForm || "",
    "{{r2Count}}": String(data.r2Count),
    "{{r3Count}}": String(data.r3Count),
    "{{persistScore}}": String(data.persistScore),
    "{{finalChoiceText}}": data.finalChoiceText || ""
  };
  Object.keys(tags).forEach(function (key) {
    body.replaceText(key.replace(/[{}]/g, "\\$&"), tags[key]);
  });

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(copy.getId()).getAs("application/pdf");
  DriveApp.getFileById(copy.getId()).setTrashed(true); // 清走臨時副本，PDF已經攞埋做blob
  return pdfBlob;
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
