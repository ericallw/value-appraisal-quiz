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
 */

var SHEET_ID = "TODO_填你個Sheet_ID";
var TEMPLATE_DOC_ID = "TODO_填你個PDF範本Doc_ID";

var HEADER_ROW = [
  "時間", "Email", "鎖定卡", "分類", "分類名", "類型",
  "第一輪張數", "第二輪張數", "第三輪張數", "撐得住指數",
  "最卡原因", "最快變現形態", "數碼產品形態"
];

var TYPE_HEAD = {
  hidden: "你唔係冇嘢好賣，係從來冇人知你有。",
  signal: "你手上唔只一件值錢嘅事，你卡喺仲未開價。",
  material: "而家仲未有，呢個係正常，亦係唯一誠實嘅答案。"
};
var TYPE_ACTION = {
  hidden: "將呢件事做一次，公開發出嚟，附上「有需要可以搵我」。",
  signal: "公開標價，收第一個陌生人嘅錢。",
  material: "由第一輪圈過嘅卡度揀一張，公開記錄 30 日。"
};
var TYPE_QUOTE = {
  hidden: "「冇需求」同「需求仲未搵到你」，係兩件完全唔同嘅事。",
  signal: "專業唔係準備好先發生，係收第一次錢之後先開始。",
  material: "方向唔係諗出嚟嘅結論，係做出嚟嘅副產品。"
};

function doGet(e) {
  return ContentService.createTextOutput("OK — 值錢鑑定卡後端運行緊");
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) {
      return jsonOut({ ok: false, error: "invalid email" });
    }

    appendRow(data);

    var pdfBlob = buildPdf(data);
    sendResultEmail(data.email, pdfBlob, data);

    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function appendRow(data) {
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
    "{{persistScore}}": String(data.persistScore) + " / 8",
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
  MailApp.sendEmail({
    to: email,
    subject: "你嘅鑑定證書：" + (data.lockedCardText || ""),
    body: "附件係你嘅完整鑑定分析 PDF。\n\n如果有問題，直接reply呢封email。",
    attachments: [pdfBlob.setName("值錢鑑定證書.pdf")]
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
