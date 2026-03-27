// ============================================================
// 📋 SETUP — запустіть один раз для створення структури аркушів
// ============================================================
// Файл: Setup.gs
// Додайте цей код поруч з Code.gs у Apps Script Editor
// Після запуску можна видалити цей файл
// ============================================================

function setupPanelSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ─── Аркуш "Налаштування" ──────────────────────────────

  let settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(SHEET_SETTINGS);
  }
  settings.clear();

  // Заголовки
  const settingsData = [
    ["Параметр", "Значення", "Опис"],
    ["Місяць", "", "Число від 1 до 12"],
    ["Рік", 2026, ""],
    ["ID папки шаблонів", "", "Натисніть 'Створити структуру папок' або вставте вручну"],
    ["ID папки вихідних файлів", "", "Натисніть 'Створити структуру папок' або вставте вручну"],
  ];
  settings.getRange(1, 1, settingsData.length, 3).setValues(settingsData);

  // Форматування
  settings.getRange("A1:C1").setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  settings.setColumnWidth(1, 220);
  settings.setColumnWidth(2, 400);
  settings.setColumnWidth(3, 350);
  settings.getRange("A1:C5").setBorder(true, true, true, true, true, true);

  // Dropdown для місяця
  const monthRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 12)
    .setHelpText("Введіть число від 1 до 12")
    .build();
  settings.getRange("B1").setDataValidation(null); // clear first row header
  settings.getRange("B2").setDataValidation(monthRule);

  // ─── Аркуш "Дані" ─────────────────────────────────────

  let data = ss.getSheetByName(SHEET_DATA);
  if (!data) {
    data = ss.insertSheet(SHEET_DATA);
  }
  data.clear();

  // Заголовки
  const headers = [
    "ПІБ спеціаліста",           // A
    "Коротке ім'я",               // B
    "РНОКПП",                     // C
    "ID шаблону (Google Sheet)",   // D
    "Останній № акта",            // E
    "Опис послуг (з Gemini)",     // F
    "Ціна (грн)",                 // G
    "Сума (грн)",                 // H
  ];
  data.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Форматування заголовків
  data.getRange("A1:H1")
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white")
    .setWrap(true)
    .setVerticalAlignment("middle");

  // Ширини колонок
  data.setColumnWidth(1, 280);  // ПІБ
  data.setColumnWidth(2, 140);  // Коротке ім'я
  data.setColumnWidth(3, 120);  // РНОКПП
  data.setColumnWidth(4, 300);  // ID шаблону
  data.setColumnWidth(5, 120);  // Останній номер
  data.setColumnWidth(6, 500);  // Послуги
  data.setColumnWidth(7, 100);  // Ціна
  data.setColumnWidth(8, 100);  // Сума

  // Приклад першого рядка
  const exampleRow = [
    "Прізвище Ім'я По батькові",
    "Прізвище І.",
    "1234567890",
    "",  // ID шаблону — заповнити після завантаження
    20,
    "Послуги, що включають в себе: \nПроведено 1:1 зустрічі;\nПроведено адаптаційні зустрічі;",
    45000,
    45000,
  ];
  data.getRange(2, 1, 1, exampleRow.length).setValues([exampleRow]);

  // Додаємо ще 14 порожніх рядків
  for (let i = 3; i <= 16; i++) {
    data.getRange(i, 5).setValue(20); // default last act number
  }

  // Рамки для 16 рядків
  data.getRange("A1:H16").setBorder(true, true, true, true, true, true);

  // Формат числових колонок
  data.getRange("G2:G16").setNumberFormat("#,##0.00");
  data.getRange("H2:H16").setNumberFormat("#,##0.00");

  // Замороження заголовка
  data.setFrozenRows(1);

  // Видаляємо дефолтний Sheet1 якщо є
  const sheet1 = ss.getSheetByName("Sheet1") || ss.getSheetByName("Аркуш1");
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  SpreadsheetApp.getUi().alert(
    "✅ Структуру створено!",
    "Аркуші «Дані» та «Налаштування» готові.\n\n" +
    "Наступний крок:\n" +
    "1. Меню Акти → Створити структуру папок\n" +
    "2. Завантажте шаблони у папку «Шаблони»\n" +
    "3. Заповніть дані спеціалістів",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
