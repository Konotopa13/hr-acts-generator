// ============================================================
// ВАРІАНТ А: ОСОБИСТИЙ ДИСК (Мій диск)
// ============================================================
// Не потребує Drive API. Працює з DriveApp.
// Підходить, коли шаблони та вихідні файли на вашому особистому диску.
// ============================================================

const SHEET_DATA = "Дані";
const SHEET_SETTINGS = "Налаштування";

const MONTH_NAMES_UK = {
  1: "січень", 2: "лютий", 3: "березень", 4: "квітень",
  5: "травень", 6: "червень", 7: "липень", 8: "серпень",
  9: "вересень", 10: "жовтень", 11: "листопад", 12: "грудень"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Акти")
    .addItem("Згенерувати акти за місяць", "generateAllActs")
    .addItem("Створити структуру папок", "createFolderStructure")
    .addSeparator()
    .addItem("Інструкція", "showHelp")
    .addToUi();
}

function generateAllActs() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) { ui.alert("Не знайдено аркуш Налаштування!"); return; }

  var month = Number(settingsSheet.getRange("B2").getValue());
  var year = Number(settingsSheet.getRange("B3").getValue());
  var templatesFolderId = String(settingsSheet.getRange("B4").getValue()).trim();
  var outputFolderId = String(settingsSheet.getRange("B5").getValue()).trim();

  if (!month || !year || !templatesFolderId || !outputFolderId) {
    ui.alert("Заповніть всі поля на аркуші Налаштування");
    return;
  }

  var monthName = MONTH_NAMES_UK[month];
  var confirm = ui.alert("Підтвердження", "Згенерувати акти за " + monthName + " " + year + "?", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var dataSheet = ss.getSheetByName(SHEET_DATA);
  if (!dataSheet) { ui.alert("Не знайдено аркуш Дані!"); return; }

  var lastRow = dataSheet.getLastRow();
  if (lastRow < 2) { ui.alert("Немає даних спеціалістів!"); return; }

  var data = dataSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // Створюємо підпапку місяця
  var outputFolder = DriveApp.getFolderById(outputFolderId);
  var monthFolderName = year + "_" + String(month).padStart(2, "0");
  var monthFolder;
  var existing = outputFolder.getFoldersByName(monthFolderName);
  if (existing.hasNext()) {
    monthFolder = existing.next();
  } else {
    monthFolder = outputFolder.createFolder(monthFolderName);
  }

  var lastDay = new Date(year, month, 0).getDate();
  var successCount = 0;
  var errors = [];

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    var shortName = data[i][1];
    var taxId = data[i][2];
    var templateId = String(data[i][3]).trim();
    var lastActNumber = Number(data[i][4]);
    var services = data[i][5];
    var price = data[i][6];
    var total = data[i][7];

    if (!name || !templateId) continue;

    var actNumber = lastActNumber + 1;
    var rowIndex = i + 2;

    try {
      var templateFile = DriveApp.getFileById(templateId);
      var copyName = taxId + "_" + shortName + "_Акт_за_" + String(month).padStart(2, "0") + "_" + year;
      var copy = templateFile.makeCopy(copyName, monthFolder);
      var actSS = SpreadsheetApp.openById(copy.getId());
      var actSheet = actSS.getSheets()[0];

      // D2: Номер акта
      var oldTitle = String(actSheet.getRange("D2").getValue() || "");
      var idx = oldTitle.lastIndexOf("\u2116");
      var baseTitle = (idx >= 0) ? oldTitle.substring(0, idx + 1) + " " : oldTitle;
      actSheet.getRange("D2").setValue(baseTitle + actNumber);

      // K5: Дата
      actSheet.getRange("K5").setValue(new Date(year, month - 1, lastDay));

      // C15: Послуги
      if (services) actSheet.getRange("C15").setValue(services);

      // I15: Ціна
      if (price) actSheet.getRange("I15").setValue(price);

      // K15: Сума
      if (total) actSheet.getRange("K15").setValue(total);

      // G22: Місяць
      actSheet.getRange("G22").setValue(monthName);

      // B23: Сума прописом (у дужках)
      if (total) actSheet.getRange("B23").setValue(numberToWordsUAH(total));

      SpreadsheetApp.flush();
      dataSheet.getRange(rowIndex, 5).setValue(actNumber);
      successCount++;
    } catch (e) {
      errors.push(name + ": " + e.message);
    }
  }

  var message = "Згенеровано: " + successCount + " актів\nПапка: " + monthFolderName;
  if (errors.length > 0) {
    message += "\n\nПомилки (" + errors.length + "):\n" + errors.join("\n");
  }
  ui.alert("Результат", message, ui.ButtonSet.OK);
}

// ─── СТВОРЕННЯ ПАПОК (особистий диск) ───────────────────────

function createFolderStructure() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Створення структури", "Введіть назву головної папки (наприклад: Акти):", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var folderName = result.getResponseText().trim() || "Акти";
  var root = DriveApp.createFolder(folderName);
  var templates = root.createFolder("Шаблони");
  var output = root.createFolder("Вихідні файли");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (settingsSheet) {
    settingsSheet.getRange("B4").setValue(templates.getId());
    settingsSheet.getRange("B5").setValue(output.getId());
  }

  ui.alert("Папки створено!", folderName + "/\n  Шаблони\n  Вихідні файли\n\nID записано в налаштування.", ui.ButtonSet.OK);
}

// ─── СУМА ПРОПИСОМ ──────────────────────────────────────────

function numberToWordsUAH(amount) {
  var intPart = Math.floor(amount);
  var kopecks = Math.round((amount - intPart) * 100);
  var hryvniWords = integerToWordsUk(intPart);
  var hryvniSuffix = getCurrencySuffix(intPart, "гривня", "гривні", "гривень");
  var kopeckWords = integerToWordsUk(kopecks);
  var kopeckSuffix = getCurrencySuffix(kopecks, "копійка", "копійки", "копійок");
  return "(" + hryvniWords + " " + hryvniSuffix + ", " + kopeckWords + " " + kopeckSuffix + ")";
}

function getCurrencySuffix(n, one, twoFour, fivePlus) {
  var abs = Math.abs(n) % 100;
  if (abs >= 11 && abs <= 19) return fivePlus;
  var lastDigit = abs % 10;
  if (lastDigit === 1) return one;
  if (lastDigit >= 2 && lastDigit <= 4) return twoFour;
  return fivePlus;
}

function integerToWordsUk(n) {
  if (n === 0) return "нуль";
  var ones = ["", "одна", "дві", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"];
  var onesM = ["", "один", "два", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять"];
  var teens = ["десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п'ятнадцять", "шістнадцять", "сімнадцять", "вісімнадцять", "дев'ятнадцять"];
  var tens = ["", "десять", "двадцять", "тридцять", "сорок", "п'ятдесят", "шістдесят", "сімдесят", "вісімдесят", "дев'яносто"];
  var hundreds = ["", "сто", "двісті", "триста", "чотириста", "п'ятсот", "шістсот", "сімсот", "вісімсот", "дев'ятсот"];

  function threeDigits(num, feminine) {
    if (num === 0) return "";
    var parts = [];
    var h = Math.floor(num / 100);
    var remainder = num % 100;
    var t = Math.floor(remainder / 10);
    var o = remainder % 10;
    if (h > 0) parts.push(hundreds[h]);
    if (remainder >= 10 && remainder <= 19) {
      parts.push(teens[remainder - 10]);
    } else {
      if (t > 0) parts.push(tens[t]);
      if (o > 0) parts.push(feminine ? ones[o] : onesM[o]);
    }
    return parts.join(" ");
  }

  var result = [];
  var millions = Math.floor(n / 1000000);
  if (millions > 0) result.push(threeDigits(millions, false) + " " + getCurrencySuffix(millions, "мільйон", "мільйони", "мільйонів"));
  var thousands = Math.floor((n % 1000000) / 1000);
  if (thousands > 0) result.push(threeDigits(thousands, true) + " " + getCurrencySuffix(thousands, "тисяча", "тисячі", "тисяч"));
  var units = n % 1000;
  if (units > 0) result.push(threeDigits(units, true));
  return result.join(" ").replace(/\s+/g, " ").trim();
}

function showHelp() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:16px;line-height:1.6}.s{margin:8px 0;padding:8px;background:#f8f9fa;border-radius:8px}</style>' +
    '<h3>Варіант А: Особистий диск</h3>' +
    '<div class="s"><b>1.</b> Акти > Створити структуру папок</div>' +
    '<div class="s"><b>2.</b> Завантажте шаблони у папку Шаблони</div>' +
    '<div class="s"><b>3.</b> Заповніть аркуш Дані</div>' +
    '<div class="s"><b>Щомісяця:</b> Оновіть послуги + суму > Акти > Згенерувати</div>' +
    '<p>Drive API НЕ потрібен. Працює тільки з особистим диском.</p>'
  ).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, "Інструкція");
}
