// ============================================================
// ВАРІАНТ В: ГНУЧКИЙ (налаштовувані клітинки)
// ============================================================
// Для кожного спеціаліста можна задати свої клітинки у шаблоні.
// Працює зі спільним диском. Потребує Drive API v2.
//
// Аркуш "Маппінг" визначає, яку клітинку чим заповнювати:
//   Колонка A = ID шаблону (або "default")
//   Колонка B = Клітинка (наприклад D2, K5, C15)
//   Колонка C = Тип: act_number / date / services / price / total / month / words / custom
//   Колонка D = Значення для custom (статичний текст або посилання на колонку аркуша Дані)
// ============================================================

const SHEET_DATA = "Дані";
const SHEET_SETTINGS = "Налаштування";
const SHEET_MAPPING = "Маппінг";

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

// ─── ЗЧИТУВАННЯ МАППІНГУ ────────────────────────────────────

function loadMappings(ss) {
  var mappingSheet = ss.getSheetByName(SHEET_MAPPING);
  if (!mappingSheet) return null;

  var lastRow = mappingSheet.getLastRow();
  if (lastRow < 2) return {};

  var raw = mappingSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var mappings = {};

  for (var i = 0; i < raw.length; i++) {
    var key = String(raw[i][0]).trim(); // template ID або "default"
    var cell = String(raw[i][1]).trim(); // клітинка
    var type = String(raw[i][2]).trim().toLowerCase(); // тип
    var customValue = raw[i][3]; // значення для custom

    if (!key || !cell || !type) continue;

    if (!mappings[key]) mappings[key] = [];
    mappings[key].push({ cell: cell, type: type, customValue: customValue });
  }

  return mappings;
}

function getMappingForTemplate(mappings, templateId) {
  if (!mappings) {
    // Дефолтний маппінг (як у варіантах А/Б)
    return [
      { cell: "D2", type: "act_number", customValue: "" },
      { cell: "K5", type: "date", customValue: "" },
      { cell: "C15", type: "services", customValue: "" },
      { cell: "I15", type: "price", customValue: "" },
      { cell: "K15", type: "total", customValue: "" },
      { cell: "G22", type: "month", customValue: "" },
      { cell: "B23", type: "words", customValue: "" }
    ];
  }

  // Спочатку шукаємо спеціфічний маппінг для цього шаблону
  if (mappings[templateId]) return mappings[templateId];

  // Інакше — дефолтний
  if (mappings["default"]) return mappings["default"];

  // Якщо нічого немає — стандартний
  return [
    { cell: "D2", type: "act_number", customValue: "" },
    { cell: "K5", type: "date", customValue: "" },
    { cell: "C15", type: "services", customValue: "" },
    { cell: "I15", type: "price", customValue: "" },
    { cell: "K15", type: "total", customValue: "" },
    { cell: "G22", type: "month", customValue: "" },
    { cell: "B23", type: "words", customValue: "" }
  ];
}

// ─── ГОЛОВНА ФУНКЦІЯ ────────────────────────────────────────

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

  // Зчитуємо дані (колонки A-H + додаткові колонки I+ для custom)
  var numCols = dataSheet.getLastColumn();
  var data = dataSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var headers = dataSheet.getRange(1, 1, 1, numCols).getValues()[0];

  // Завантажуємо маппінг
  var mappings = loadMappings(ss);

  // Створюємо підпапку місяця
  var monthFolderName = year + "_" + String(month).padStart(2, "0");
  var monthFolderId;

  var searchQuery = "title='" + monthFolderName + "' and '" + outputFolderId + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false";
  var searchResult = Drive.Files.list({
    q: searchQuery,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    corpora: "allDrives"
  });

  if (searchResult.items && searchResult.items.length > 0) {
    monthFolderId = searchResult.items[0].id;
  } else {
    var created = Drive.Files.insert({
      title: monthFolderName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [{ id: outputFolderId }]
    }, null, { supportsAllDrives: true });
    monthFolderId = created.id;
  }

  var lastDay = new Date(year, month, 0).getDate();
  var successCount = 0;
  var errors = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = row[0];
    var shortName = row[1];
    var taxId = row[2];
    var templateId = String(row[3]).trim();
    var lastActNumber = Number(row[4]);
    var services = row[5];
    var price = row[6];
    var total = row[7];

    if (!name || !templateId) continue;

    var actNumber = lastActNumber + 1;
    var rowIndex = i + 2;

    try {
      var templateFile = DriveApp.getFileById(templateId);
      var mimeType = templateFile.getMimeType();
      var copyName = taxId + "_" + shortName + "_Акт_за_" + String(month).padStart(2, "0") + "_" + year;

      var actSS;

      if (mimeType === "application/vnd.google-apps.spreadsheet") {
        var copyResource = { title: copyName, parents: [{ id: monthFolderId }] };
        var copied = Drive.Files.copy(copyResource, templateId, { supportsAllDrives: true });
        actSS = SpreadsheetApp.openById(copied.id);
      } else {
        var blob = templateFile.getBlob();
        var resource = {
          title: copyName,
          mimeType: "application/vnd.google-apps.spreadsheet",
          parents: [{ id: monthFolderId }]
        };
        var converted = Drive.Files.insert(resource, blob, { convert: true, supportsAllDrives: true });
        actSS = SpreadsheetApp.openById(converted.id);
      }

      var actSheet = actSS.getSheets()[0];

      // Отримуємо маппінг для цього шаблону
      var cellMappings = getMappingForTemplate(mappings, templateId);

      // Заповнюємо клітинки за маппінгом
      for (var j = 0; j < cellMappings.length; j++) {
        var m = cellMappings[j];
        var cellRef = m.cell;
        var type = m.type;

        switch (type) {
          case "act_number":
            var oldTitle = String(actSheet.getRange(cellRef).getValue() || "");
            var idx = oldTitle.lastIndexOf("\u2116");
            var base = (idx >= 0) ? oldTitle.substring(0, idx + 1) + " " : oldTitle;
            actSheet.getRange(cellRef).setValue(base + actNumber);
            break;

          case "date":
            actSheet.getRange(cellRef).setValue(new Date(year, month - 1, lastDay));
            break;

          case "services":
            if (services) actSheet.getRange(cellRef).setValue(services);
            break;

          case "price":
            if (price) actSheet.getRange(cellRef).setValue(price);
            break;

          case "total":
            if (total) actSheet.getRange(cellRef).setValue(total);
            break;

          case "month":
            actSheet.getRange(cellRef).setValue(monthName);
            break;

          case "words":
            if (total) actSheet.getRange(cellRef).setValue(numberToWordsUAH(total));
            break;

          case "custom":
            // customValue може бути: рядок тексту АБО назва колонки з аркуша Дані (наприклад "I" або назва заголовка)
            var val = m.customValue;
            // Перевіряємо, чи це посилання на колонку (починається з "col:")
            if (typeof val === "string" && val.indexOf("col:") === 0) {
              var colName = val.substring(4).trim();
              var colIndex = headers.indexOf(colName);
              if (colIndex >= 0) {
                actSheet.getRange(cellRef).setValue(row[colIndex]);
              }
            } else {
              actSheet.getRange(cellRef).setValue(val);
            }
            break;
        }
      }

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

// ─── СТВОРЕННЯ ПАПОК ────────────────────────────────────────

function createFolderStructure() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Створення структури",
    "Введіть ID батьківської папки на спільному диску\n(або залиште порожнім для особистого диска):",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var parentId = result.getResponseText().trim();
  var templatesFolderId, outputFolderId;

  if (parentId) {
    var templatesRes = Drive.Files.insert({
      title: "Шаблони", mimeType: "application/vnd.google-apps.folder",
      parents: [{ id: parentId }]
    }, null, { supportsAllDrives: true });
    var outputRes = Drive.Files.insert({
      title: "Вихідні файли", mimeType: "application/vnd.google-apps.folder",
      parents: [{ id: parentId }]
    }, null, { supportsAllDrives: true });
    templatesFolderId = templatesRes.id;
    outputFolderId = outputRes.id;
  } else {
    var root = DriveApp.createFolder("Акти");
    var templates = root.createFolder("Шаблони");
    var output = root.createFolder("Вихідні файли");
    templatesFolderId = templates.getId();
    outputFolderId = output.getId();
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (settingsSheet) {
    settingsSheet.getRange("B4").setValue(templatesFolderId);
    settingsSheet.getRange("B5").setValue(outputFolderId);
  }

  ui.alert("Папки створено!", "Шаблони: " + templatesFolderId + "\nВихідні: " + outputFolderId, ui.ButtonSet.OK);
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
    if (remainder >= 10 && remainder <= 19) { parts.push(teens[remainder - 10]); }
    else { if (t > 0) parts.push(tens[t]); if (o > 0) parts.push(feminine ? ones[o] : onesM[o]); }
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
    '<style>body{font-family:Arial;padding:16px;line-height:1.6}.s{margin:8px 0;padding:8px;background:#f8f9fa;border-radius:8px}code{background:#eee;padding:2px 4px}</style>' +
    '<h3>Варіант В: Гнучкий</h3>' +
    '<p>Якщо шаблони різних спеціалістів мають різну структуру клітинок, створіть аркуш <b>Маппінг</b>.</p>' +
    '<div class="s"><b>Аркуш Маппінг:</b><br>' +
    'A = ID шаблону (або <code>default</code>)<br>' +
    'B = Клітинка (напр. <code>D2</code>, <code>K5</code>)<br>' +
    'C = Тип: <code>act_number</code> / <code>date</code> / <code>services</code> / <code>price</code> / <code>total</code> / <code>month</code> / <code>words</code> / <code>custom</code><br>' +
    'D = Для custom: текст або <code>col:Назва колонки</code></div>' +
    '<div class="s">Якщо аркуш Маппінг відсутній — скрипт використовує стандартні клітинки (D2, K5, C15 тощо).</div>' +
    '<p>Потребує Drive API v2.</p>'
  ).setWidth(450).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, "Інструкція");
}
