// ====================================================================
// NEXUS ВЕБ-ЯДРО: Главный шлюз запуска Автономного Веб-Приложения
// ====================================================================
function doGet(e) {
  try {
    // Подгружаем каркас будущего сайта (создадим файл index.html следующим шагом)
    var template = HtmlService.createTemplateFromFile("index");
    
    // Передаем в ДНК сайта готовые карты метаданных каталога и клиентов
    template.catalog_map = getNexusCatalogMap();
    template.clients_map = getNexusClientsMap();
    
    return template.evaluate()
      .setTitle("NEXUS CRM/ARM SYSTEM")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h2 style='color:#ff3333;font-family:sans-serif;padding:20px;'>Критический сбой ядра NEXUS: " + err.message + "</h2>");
  }
}

// ====================================================================
// Константа
// ====================================================================
const BASE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzejv7lBeHFS5D8AfskBIZuyNov7qBK2OFk2v9zjFb3ePQ789xGBts3zlDd8U4Punpp/exec";

// ====================================================================
// Главный триггер запуска книги, создает меню и фасады страниц
// ====================================================================
function onOpen(e) {
  // ...Ваш старый код создания других пунктов меню (если был)...

  // 1. СОЗДАЕМ ВЕЧНУЮ СИСТЕМНУЮ КНОПКУ НА САМОМ ВЕРХУ GOOGLE
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("NEXUS")
      .addItem("Панель управления", "toggleSidebar")
      .addToUi();
  } catch(menuErr) {}

  // 2. РАЗОВАЯ СБОРКА ЧЕРДАКОВ ДЛЯ ГЛАВНЫХ СТРАНИЦ CRM ПРИ ЗАПУСКЕ КНИГИ
  try {
    buildNexusDashboardAndFixFormats("Каталог");
    buildNexusDashboardAndFixFormats("Склад");
    buildNexusDashboardAndFixFormats("Поставщики");
    buildNexusDashboardAndFixFormats("Персонал");
    buildNexusDashboardAndFixFormats("Клиенты");
    buildNexusDashboardAndFixFormats("Продажи");
  } catch(buildErr) {}
}

// ====================================================================
// Следящий триггер изменений, ловит правки на чердаке
// ====================================================================
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    
    // Если правка произошла на чердаке (строки 1-3), мгновенно откатываем и чиним формат
    if (range.getRow() <= 3) {
      buildNexusDashboardAndFixFormats(sheetName);
    }
  } catch(err) {
    Logger.log("Ошибка триггера автовозврата: " + err.message);
  }
}

// ====================================================================
// Вызов боковой панели навигации
// ====================================================================
function toggleSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile("SidebarUI").evaluate().setTitle("ТЕРМИНАЛ НАВИГАЦИИ");
    SpreadsheetApp.getUi().showSidebar(html);
  } catch(err) {
    SpreadsheetApp.getUi().alert("Ошибка: Проверьте имя файла боковой панели 'SidebarUI'");
  }
}

// ====================================================================
// Модальный шлюз открытия формы добавления товара
// ====================================================================
function openModalAddForm() {
  var template = HtmlService.createTemplateFromFile("Catalog_AddForm_NEW");
  template.editRowNumber = "0";
  var htmlOutput = template.evaluate()
    .setWidth(1000) 
    .setHeight(700) 
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, " "); 
}

// ====================================================================
// Модальный шлюз открытия карточки просмотра
// ====================================================================
function openModalCardViewByItemId(targetId) {
  try {
    if (!targetId) return { success: false, msg: "ID позиции пуст" };
 
    var search = findRowByItemId(targetId);
    if (!search.success) return { success: false, msg: search.msg };
 
    var template = HtmlService.createTemplateFromFile("Catalog_CardView_NEW");
    template.targetItemId = targetId; 
 
    var htmlOutput = template.evaluate()
      .setWidth(950)
      .setHeight(750)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
 
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, " "); 
    return { success: true };
  } catch (err) {
    return { success: false, msg: "Сбой модального движка: " + err.message };
  }
}

// ====================================================================
// СЛУЖЕБНЫЙ ШЛЮЗ ДЛЯ ПЕРЕКЛЮЧЕНИЯ С КАРТОЧКИ ПРОСМОТРА НА РЕДАКТИРОВАНИЕ
// ====================================================================
function switchFromViewToEditMode(targetId) {
  try {
    if (!targetId) return;
    
    // Находим номер физической строки товара в базе по его системному ID
    var search = findRowByItemId(targetId);
    if (!search.success) return;
    
    // Открываем нашу универсальную форму, но вместо "0" передаем реальный номер строки!
    var template = HtmlService.createTemplateFromFile("Catalog_AddForm_NEW");
    template.editRowNumber = search.row.toString(); // Форма поймет, что это режим правки
    
    var htmlOutput = template.evaluate()
      .setWidth(1000)
      .setHeight(700)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
    // Магия Google Apps Script: новое окно автоматически заместит старое на экране
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, " ");
  } catch (err) {
    Logger.log("Ошибка переключения режимов NEXUS: " + err.message);
  }
}

// ====================================================================
// Модальный шлюз открытия окна импорта
// ====================================================================
function openExcelImportDialog() {
  try {
    var template = HtmlService.createTemplateFromFile("Excel_ImportForm");
    
    // Передаем живую ДНК-карту констант прямо в HTML шаблон Импорта
    template.catalog_map = getNexusCatalogMap();
    
    var htmlOutput = template.evaluate()
      .setWidth(950)
      .setHeight(650)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, " ");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Критический сбой запуска окна Импорта: " + err.message);
  }
}

// ====================================================================
// Модальный шлюз открытия окна экспорта
// ====================================================================
function openExcelExportDialog() {
  try {
    const template = HtmlService.createTemplateFromFile("Excel_ExportForm");
    
    // Передаем живую ДНК-карту констант прямо в HTML шаблон Экспорта
    template.catalog_map = getNexusCatalogMap();
    
    const html = template.evaluate()
      .setWidth(950) // Выставляем коммерческую ширину 950px под стиль импорта
      .setHeight(650) // Высота 650px для идеального баланса на экранах ноутов
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
    SpreadsheetApp.getUi().showModalDialog(html, " ");
  } catch (err) {
    SpreadsheetApp.getUi().alert("Критический сбой запуска окна Экспорта: " + err.message);
  }
}

// ====================================================================
// Модальный шлюз открытия рубрикатора категорий
// ====================================================================
function openModalRubricatorDialog() {
  var template = HtmlService.createTemplateFromFile("Catalog_RubricatorForm");
  var htmlOutput = template.evaluate().setTitle("NEXUS CRM — РУБРИКАТОР КАТАЛОГА").setWidth(950).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, " ");
}

// ====================================================================
// Модальный шлюз Терминала Продаж — собирает товары и менеджеров
// ====================================================================
function openModalSalesWizard(queueIds) {
  try {
    if (!queueIds || queueIds.length === 0) return;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Каталог") || ss.getSheetByName("КАТАЛОГ");
    if (!sheet) throw new Error("Лист 'Каталог' не найден");
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return;
    
    // СРАЗУ ИНИЦИАЛИЗИРУЕМ ШАБЛОН, чтобы он был доступен во всем коде функции
    var template = HtmlService.createTemplateFromFile('SalesWizardModal');
    
    var colMap = getNexusColumnMapping();
    var rawData = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
    var selectedItems = [];
    
    function getVal(row, key, defaultVal) {
      var idx = colMap[key];
      if (idx === undefined || idx === -1 || row[idx] === undefined) return defaultVal;
      return row[idx];
    }
    
    for (var j = 0; j < queueIds.length; j++) {
      var targetId = queueIds[j].toString().trim();
      for (var i = 0; i < rawData.length; i++) {
        var row = rawData[i];
        var currentId = getVal(row, "id", "").toString().trim();
        
        if (currentId === targetId) {
          // СБОРКА СТРОГО ПО ДНК-КАРТЕ КОНСТАНТ С ТВОЕГО СКРИНШОТА
          var item = {
            id: currentId,
            record_type: colMap["record_type"] !== -1 ? row[colMap["record_type"]].toString().trim() : "Товар",
            name: colMap["name"] !== -1 ? row[colMap["name"]].toString() : "Без названия",
            brand: colMap["brand"] !== -1 ? row[colMap["brand"]].toString() : "",
            sku: colMap["sku"] !== -1 ? row[colMap["sku"]].toString() : "",
            barcode: colMap["barcode"] !== -1 ? row[colMap["barcode"]].toString() : "",
            stock_qty: colMap["stock_qty"] !== -1 ? parseFloat(row[colMap["stock_qty"]]) || 0 : 0,
            stock_unit: colMap["stock_unit"] !== -1 ? row[colMap["stock_unit"]].toString() : "шт.",
            price_buy: colMap["price_buy"] !== -1 ? parseFloat(row[colMap["price_buy"]]) || 0 : 0,
            price_wholesale: colMap["price_wholesale"] !== -1 ? parseFloat(row[colMap["price_wholesale"]]) || 0 : 0,
            price_retail: colMap["price_retail"] !== -1 ? parseFloat(row[colMap["price_retail"]]) || 0 : 0,
            production_cost: colMap["production_cost"] !== -1 ? parseFloat(row[colMap["production_cost"]]) || 0 : 0,
            vat_rate: colMap["vat_rate"] !== -1 ? row[colMap["vat_rate"]].toString() : "Без НДС",
            time_norm_val: colMap["time_norm_val"] !== -1 ? row[colMap["time_norm_val"]].toString() : "",
            time_norm_unit: colMap["time_norm_unit"] !== -1 ? row[colMap["time_norm_unit"]].toString() : "",
            duration_val: colMap["duration_val"] !== -1 ? row[colMap["duration_val"]].toString() : "",
            duration_unit: colMap["duration_unit"] !== -1 ? row[colMap["duration_unit"]].toString() : ""
          };
          selectedItems.push(item);
          break;
        }
      }
    }
    
    // БЛОК СБОРА МЕНЕДЖЕРОВ ИЗ ВКЛАДКИ ПЕРСОНАЛ
    var empSheet = ss.getSheetByName("Персонал") || ss.getSheetByName("ПЕРСОНАЛ");
    var managersList = [];
    if (empSheet && empSheet.getLastRow() >= 5) {
      var empData = empSheet.getRange(5, 1, empSheet.getLastRow() - 4, 1).getValues();
      empData.forEach(function(row) {
        var name = row ? row.toString().trim() : "";
        if (name) managersList.push(name);
      });
    }
    
    // Упаковываем JSON-пакеты в уже созданный шаблон
    template.itemsJson = JSON.stringify(selectedItems);
    template.managersJson = JSON.stringify(managersList); 
    
    var htmlOutput = template.evaluate()
      .setWidth(1150)
      .setHeight(750)
      .setTitle(" ");
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, " ");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Критический сбой модуля продаж NEXUS: " + e.toString());
  }
}

// ====================================================================
// Ядро сбора констант для выпадающих списков
// ====================================================================
function getConstantsForForm() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
 
    const constantsSheet = ss.getSheetByName("Константы");
    if (!constantsSheet) throw new Error("Лист Константы не найден");
    const lastRowConst = constantsSheet.getLastRow();
    const constRange = constantsSheet.getRange(2, 1, lastRowConst > 1 ? lastRowConst - 1 : 1, 18).getValues();
    
    const recordTypes = new Set(); 
    const vidIds = new Set(); 
    const unitsWarranty = new Set(); 
    const unitsQty = new Set(); 
    const unitsWeight = new Set(); 
    const unitsTimeNorm = new Set(); 
    const statuses = new Set(); 
    const unitsDuration = new Set(); 
    const calcTypes = new Set(); 
    const currencies = new Set(); 
    const categories = new Set(); 
    const subcategories = new Set(); 
    const vatRate = new Set();

    if (lastRowConst >= 2) {
      constRange.forEach(row => {
        if (row[2] && row[2].toString().trim() !== "") recordTypes.add(row[2].toString().trim());
        if (row[3] && row[3].toString().trim() !== "") vidIds.add(row[3].toString().trim());
        if (row[4] && row[4].toString().trim() !== "") unitsWarranty.add(row[4].toString().trim());
        if (row[5] && row[5].toString().trim() !== "") unitsQty.add(row[5].toString().trim());
        if (row[6] && row[6].toString().trim() !== "") unitsWeight.add(row[6].toString().trim());
        if (row[7] && row[7].toString().trim() !== "") unitsTimeNorm.add(row[7].toString().trim());
        if (row[8] && row[8].toString().trim() !== "") statuses.add(row[8].toString().trim());
        if (row[9] && row[9].toString().trim() !== "") unitsDuration.add(row[9].toString().trim());
        if (row[11] && row[11].toString().trim() !== "") calcTypes.add(row[11].toString().trim());
        if (row[12] && row[12].toString().trim() !== "") currencies.add(row[12].toString().trim());
        if (row[13] && row[13].toString().trim() !== "") categories.add(row[13].toString().trim());
        if (row[14] && row[14].toString().trim() !== "") subcategories.add(row[14].toString().trim());
        if (row[17] !== undefined && row[17] !== null && row[17].toString().trim() !== "") {vatRate.add(row[17].toString().trim());}

      });
    }

    const empSheet = ss.getSheetByName("Персонал");
    const responsibles = [];
    if (empSheet && empSheet.getLastRow() >= 5) {
      const empData = empSheet.getRange(5, 1, empSheet.getLastRow() - 4, 1).getValues();
      empData.forEach(row => { if(row[0].toString().trim()) responsibles.push(row[0].toString().trim()); });
    }

    const supplierSheet = ss.getSheetByName("Поставщики");
    const suppliers = [];
    if (supplierSheet && supplierSheet.getLastRow() >= 5) {
      const supData = supplierSheet.getRange(5, 1, supplierSheet.getLastRow() - 4, 1).getValues();
      supData.forEach(row => { if(row[0].toString().trim()) suppliers.push(row[0].toString().trim()); });
    }

    return {
      record_types: Array.from(recordTypes),
      vid_ids: Array.from(vidIds),
      units_warranty: Array.from(unitsWarranty),
      units_qty: Array.from(unitsQty),
      units_weight: Array.from(unitsWeight),
      units_time_norm: Array.from(unitsTimeNorm),
      statuses: Array.from(statuses),
      units_duration: Array.from(unitsDuration),
      calc_types: Array.from(calcTypes),
      currencies: Array.from(currencies),
      categories: Array.from(categories),
      subcategories: Array.from(subcategories),
      vat_rate: Array.from(vatRate),
      responsibles: responsibles,
      suppliers: suppliers,
      
      // НАШ НОВЫЙ ДИНАМИЧЕСКИЙ ПАСПОРТ ИМЕН ДЛЯ ВЕРСТКИ ФОРМЫ ДОБАВЛЕНИЯ
      catalog_map: getNexusCatalogMap() 
    };
  } catch (err) {
    throw new Error("Ошибка раздельного сбора констант: " + err.message);
  }
}

// ====================================================================
// Сборщик справочника категорий для рубрикатора
// ====================================================================
function getRubricatorRawData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("КОНСТАНТЫ") || ss.getSheetByName("Константы");
    if (!sheet) return { categories: [], subcategories: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { categories: [], subcategories: [] };
    
    var catValues = sheet.getRange(2, 14, lastRow - 1, 1).getValues().map(r => (r[0] || "").toString().trim()).filter(Boolean);
    var subValues = sheet.getRange(2, 15, lastRow - 1, 1).getValues().map(r => (r[0] || "").toString().trim()).filter(Boolean);
    return { categories: catValues, subcategories: subValues };
  } catch(e) { return { categories: [], subcategories: [] }; }
}

// ====================================================================
// Сверхскоростной сборщик данных для карточки просмотра товаров
// ====================================================================
function getCatalogCardDataById(itemId) {
  try {
    if (!itemId) return { success: false, msg: "Идентификатор позиции пуст." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("КАТАЛОГ") || ss.getSheetByName("Каталог");
    if (!sheet) return { success: false, msg: "Лист 'КАТАЛОГ' не найден." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return { success: false, msg: "База данных каталога пуста." };
    
    var map = getNexusCatalogMap();
    var dataRange = sheet.getRange(5, 1, lastRow - 4, 50).getValues();
    
    var idIdx = map.id.index; 
    var targetRow = null;
    
    for (var i = 0; i < dataRange.length; i++) {
      if (dataRange[i][idIdx].toString().trim() === itemId.toString().trim()) {
        targetRow = dataRange[i];
        break;
      }
    }
    
    if (!targetRow) return { success: false, msg: "Позиция с ID " + itemId + " не найдена." };
    
    var resultData = {};
    for (var systemId in map) {
      var idx = map[systemId].index;
      resultData[systemId] = targetRow[idx] !== undefined ? targetRow[idx].toString().trim() : "";
    }
    
    // МАТЕМАТИЧЕСКАЯ ОЧИСТКА ОСТАТКОВ
    resultData._pure_stock = nexusCoreCleanNumber(resultData.stock_qty);
    resultData._pure_min_stock = nexusCoreCleanNumber(resultData.min_stock_qty);
    
    // МГНОВЕННЫЙ ПАРСИНГ ГОТОВЫХ ССЫЛОК ИЗ ТАБЛИЦЫ
    var rawPhotos = [];
    if (resultData.photos_raw) {
      var rawStr = resultData.photos_raw.toString().trim();
      if (rawStr !== "") {
        try {
          if (rawStr.indexOf("[") === 0) {
            rawPhotos = JSON.parse(rawStr);
          } else {
            rawPhotos = rawStr.split(",").map(function(l) { return l.trim(); });
          }
        } catch(e) {
          rawPhotos = [rawStr];
        }
      }
    }
    
    // Просто перекладываем готовые ссылки (lh3 или внешние интернет-адреса)
    var finalPhotosArray = [];
    rawPhotos.forEach(function(link) {
      if (link && link.toString().trim() !== "") {
        finalPhotosArray.push(link.toString().trim());
      }
    });
    
    // Добиваем пустые слоты до 10 для стабильности верстки фронтенда
    while (finalPhotosArray.length < 10) {
      finalPhotosArray.push("");
    }
    
    return {
      success: true,
      data: resultData,
      collected_photos: finalPhotosArray, // Отдаем чистые готовые URL
      catalog_map: map
    };
    
  } catch (err) {
    return { success: false, msg: "Критический сбой сборщика карточки: " + err.message };
  }
}


// ====================================================================
// Донор данных для режима редактирования по индексу строки
// ====================================================================
function getTableRowDataByRowIndex(rowIdx) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return null;
    
    // Получаем живую карту метаданных констант
    const map = getNexusCatalogMap();
    
    // Забираем строку в 50 колонок за один быстрый выстрел
    var rowValues = sheet.getRange(rowIdx, 1, 1, 50).getValues()[0];
    
    // Собираем объект, где ключи жестко привязаны к латинским системным ID
    const resultData = {};
    for (let systemId in map) {
      const idx = map[systemId].index;
      resultData[systemId] = rowValues[idx] !== undefined ? rowValues[idx].toString().trim() : "";
    }
    
    // Принудительно форматируем год под строку yyyy, если Google посчитал его датой
    if (rowValues[map.production_year.index] instanceof Date) {
      resultData.production_year = Utilities.formatDate(
        rowValues[map.production_year.index], 
        ss.getSpreadsheetTimeZone(), 
        "yyyy"
      );
    }
    
    // Дополнительно дублируем системный скрытый ID для совместимости со старыми hidden-инпутами формы
    resultData["system_id_hidden"] = rowValues[map.id.index];
    
    return resultData;
    
  } catch(e) { 
    Logger.log("Ошибка динамического донора данных правки: " + e.message);
    return null; 
  }
}

// ====================================================================
// Серверный движок добавления новой позиции в таблицу
// ====================================================================
function saveCatalogCard(data) {
  try {
    if (!data) return { success: false, error: "Системный сбой: данные формы не получены бэкендом." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var catalogSheet = ss.getSheetByName("КАТАЛОГ") || ss.getSheetByName("Каталог");
    if (!catalogSheet) throw new Error("Лист КАТАЛОГ не найден");
    
    var lastRow = catalogSheet.getLastRow();
    
    // БЛОК АВТОГЕНЕРАЦИИ ИНДЕКСА ID
    var nextId = 1;
    if (lastRow >= 5) {
      var idRange = catalogSheet.getRange(5, 1, lastRow - 4, 1).getValues();
      var ids = idRange.map(function(row) { return Number(row[0]); }).filter(function(id) { return !isNaN(id) && id > 0; });
      if (ids.length > 0) nextId = Math.max.apply(null, ids) + 1;
    }
    var targetRow = lastRow + 1 < 5 ? 5 : lastRow + 1;
    
    var map = getNexusCatalogMap();
    var newRow = Array(50).fill(""); 
    
    var photoValue = "";
    if (data.photo_raw && data.photo_raw.toString().trim() !== "") {
      photoValue = data.photo_raw.toString().trim();
    }
    
    var currentStock = nexusCoreCleanNumber(data.stock);
    var computedStatus = "";
    if (data.record_type === "Услуга") {
      computedStatus = "Доступна";
    } else {
      computedStatus = (currentStock > 0) ? "В наличии" : "Нет в наличии";
    }
    
    // СИНХРОНИЗИРОВАННЫЕ КЛЮЧИ: Обрабатываем как старые, так и новые имена для безопасности
    var typeOfCode = data.code_type || data.vid_id || "";
    var valOfCode = data.code_value || data.item_id_val || data.item_id || "";
    
    if (map.id && map.id.index !== -1) newRow[map.id.index] = nextId;
    if (map.record_type && map.record_type.index !== -1) newRow[map.record_type.index] = data.record_type || "Товар";
    if (map.code_type && map.code_type.index !== -1) newRow[map.code_type.index] = typeOfCode;
    if (map.code_value && map.code_value.index !== -1) newRow[map.code_value.index] = valOfCode;
    if (map.name && map.name.index !== -1) newRow[map.name.index] = data.name || "";
    if (map.brand && map.brand.index !== -1) newRow[map.brand.index] = data.brand || "";
    if (map.sku && map.sku.index !== -1) newRow[map.sku.index] = data.sku || "";
    if (map.barcode && map.barcode.index !== -1) newRow[map.barcode.index] = data.barcode || "";
    if (map.category && map.category.index !== -1) newRow[map.category.index] = data.category || "";
    if (map.subcategory && map.subcategory.index !== -1) newRow[map.subcategory.index] = data.subcategory || "";
    if (map.group && map.group.index !== -1) newRow[map.group.index] = data.group || "";
    if (map.modification && map.modification.index !== -1) newRow[map.modification.index] = data.mod || "";
    if (map.stock_qty && map.stock_qty.index !== -1) newRow[map.stock_qty.index] = currentStock;
    if (map.min_stock_qty && map.min_stock_qty.index !== -1) newRow[map.min_stock_qty.index] = nexusCoreCleanNumber(data.min_stock);
    if (map.stock_unit && map.stock_unit.index !== -1) newRow[map.stock_unit.index] = data.unit || "";
    if (map.status && map.status.index !== -1) newRow[map.status.index] = computedStatus;
    if (map.production_year && map.production_year.index !== -1) newRow[map.production_year.index] = data.year || "";
    if (map.warranty_val && map.warranty_val.index !== -1) newRow[map.warranty_val.index] = nexusCoreCleanNumber(data.warranty_num) || "";
    if (map.warranty_unit && map.warranty_unit.index !== -1) newRow[map.warranty_unit.index] = data.warranty_unit || "";
    if (map.weight_val && map.weight_val.index !== -1) newRow[map.weight_val.index] = nexusCoreCleanNumber(data.weight_num) || "";
    if (map.weight_unit && map.weight_unit.index !== -1) newRow[map.weight_unit.index] = data.weight_unit || "";
    if (map.price_buy && map.price_buy.index !== -1) newRow[map.price_buy.index] = nexusCoreCleanNumber(data.price_buy);
    if (map.price_wholesale && map.price_wholesale.index !== -1) newRow[map.price_wholesale.index] = nexusCoreCleanNumber(data.price_opt);
    if (map.price_retail && map.price_retail.index !== -1) newRow[map.price_retail.index] = nexusCoreCleanNumber(data.price_retail);
    if (map.production_cost && map.production_cost.index !== -1) newRow[map.production_cost.index] = nexusCoreCleanNumber(data.prod_cost);
    if (map.currency && map.currency.index !== -1) newRow[map.currency.index] = data.currency || "";
    if (map.time_norm_val && map.time_norm_val.index !== -1) newRow[map.time_norm_val.index] = nexusCoreCleanNumber(data.time_norm) || "";
    if (map.time_norm_unit && map.time_norm_unit.index !== -1) newRow[map.time_norm_unit.index] = data.time_norm_unit || "";
    if (map.duration_val && map.duration_val.index !== -1) newRow[map.duration_val.index] = nexusCoreCleanNumber(data.duration_num) || "";
    if (map.duration_unit && map.duration_unit.index !== -1) newRow[map.duration_unit.index] = data.duration_unit || "";
    if (map.responsible_user && map.responsible_user.index !== -1) newRow[map.responsible_user.index] = data.responsible || "";
    if (map.warehouse_cell && map.warehouse_cell.index !== -1) newRow[map.warehouse_cell.index] = data.cell || "";
    if (map.calc_type && map.calc_type.index !== -1) newRow[map.calc_type.index] = data.calc_type || "";
    if (map.specification && map.specification.index !== -1) newRow[map.specification.index] = data.specification || "";
    if (map.description && map.description.index !== -1) newRow[map.description.index] = data.description || "";
    if (map.photos_raw && map.photos_raw.index !== -1) newRow[map.photos_raw.index] = photoValue;
    if (map.internal_comment && map.internal_comment.index !== -1) newRow[map.internal_comment.index] = data.comment || "";
    if (map.vendor_supplier && map.vendor_supplier.index !== -1) newRow[map.vendor_supplier.index] = data.supplier || "";
    if (map.vat_rate && map.vat_rate.index !== -1) newRow[map.vat_rate.index] = data.vat_rate || "Без НДС";
    
    catalogSheet.getRange(targetRow, 1, 1, 50).setValues([newRow]);
    SpreadsheetApp.flush();
    
    return { success: true, msg: "Позиция успешно создана с системным ID: " + nextId };
  } catch (err) {
    return { success: false, error: "Ошибка бэкенда добавления: " + err.message };
  }
}

// ====================================================================
// Серверный движок обновления карточки товара
// ====================================================================
function updateCatalogCardData(itemId, formData, photosArray) {
  try {
    if (!itemId) throw new Error("ID позиции не передан на сервер");
    var searchResult = findRowByItemId(itemId);
    if (!searchResult.success) throw new Error("Позиция не найдена: " + searchResult.msg);
    
    var rowIdx = searchResult.row;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("КАТАЛОГ") || ss.getSheetByName("Каталог");
    
    var map = getNexusCatalogMap();
    var updatedRow = Array(50).fill("");
    var currentValues = sheet.getRange(rowIdx, 1, 1, 50).getValues()[0];
    
    for (var i = 0; i < 50; i++) { updatedRow[i] = currentValues[i]; }
    
    var computedStatus = "";
    if (formData.record_type === "Услуга") {
      computedStatus = "Доступна";
    } else {
      computedStatus = (Number(formData.stock) > 0) ? "В наличии" : "Нет в наличии";
    }
    
    // СИНХРОНИЗИРОВАННЫЕ КЛЮЧИ РЕДАКТИРОВАНИЯ
    var typeOfCode = formData.code_type || formData.vid_id || "";
    var valOfCode = formData.code_value || formData.item_id_val || "";
    
    if (map.id && map.id.index !== -1) updatedRow[map.id.index] = formData.item_id || itemId;
    if (map.record_type && map.record_type.index !== -1) updatedRow[map.record_type.index] = formData.record_type;
    if (map.code_type && map.code_type.index !== -1) updatedRow[map.code_type.index] = typeOfCode;
    if (map.code_value && map.code_value.index !== -1) updatedRow[map.code_value.index] = valOfCode;
    if (map.name && map.name.index !== -1) updatedRow[map.name.index] = formData.name;
    if (map.brand && map.brand.index !== -1) updatedRow[map.brand.index] = formData.brand;
    if (map.sku && map.sku.index !== -1) updatedRow[map.sku.index] = formData.sku;
    if (map.barcode && map.barcode.index !== -1) updatedRow[map.barcode.index] = formData.barcode;
    if (map.category && map.category.index !== -1) updatedRow[map.category.index] = formData.category;
    if (map.subcategory && map.subcategory.index !== -1) updatedRow[map.subcategory.index] = formData.subcategory;
    if (map.group && map.group.index !== -1) updatedRow[map.group.index] = formData.group;
    if (map.modification && map.modification.index !== -1) updatedRow[map.modification.index] = formData.mod;
    if (map.stock_qty && map.stock_qty.index !== -1) updatedRow[map.stock_qty.index] = nexusCoreCleanNumber(formData.stock);
    if (map.min_stock_qty && map.min_stock_qty.index !== -1) updatedRow[map.min_stock_qty.index] = nexusCoreCleanNumber(formData.min_stock);
    if (map.stock_unit && map.stock_unit.index !== -1) updatedRow[map.stock_unit.index] = formData.unit;
    if (map.status && map.status.index !== -1) updatedRow[map.status.index] = computedStatus;
    if (map.production_year && map.production_year.index !== -1) updatedRow[map.production_year.index] = formData.year;
    if (map.warranty_val && map.warranty_val.index !== -1) updatedRow[map.warranty_val.index] = nexusCoreCleanNumber(formData.warranty_num) || "";
    if (map.warranty_unit && map.warranty_unit.index !== -1) updatedRow[map.warranty_unit.index] = formData.warranty_unit;
    if (map.weight_val && map.weight_val.index !== -1) updatedRow[map.weight_val.index] = nexusCoreCleanNumber(formData.weight_num) || "";
    if (map.weight_unit && map.weight_unit.index !== -1) updatedRow[map.weight_unit.index] = formData.weight_unit;
    if (map.price_buy && map.price_buy.index !== -1) updatedRow[map.price_buy.index] = nexusCoreCleanNumber(formData.price_buy);
    if (map.price_wholesale && map.price_wholesale.index !== -1) updatedRow[map.price_wholesale.index] = nexusCoreCleanNumber(formData.price_opt);
    if (map.price_retail && map.price_retail.index !== -1) updatedRow[map.price_retail.index] = nexusCoreCleanNumber(formData.price_retail);
    if (map.production_cost && map.production_cost.index !== -1) updatedRow[map.production_cost.index] = nexusCoreCleanNumber(formData.prod_cost);
    if (map.currency && map.currency.index !== -1) updatedRow[map.currency.index] = formData.currency;
    if (map.time_norm_val && map.time_norm_val.index !== -1) updatedRow[map.time_norm_val.index] = nexusCoreCleanNumber(formData.time_norm);
    if (map.time_norm_unit && map.time_norm_unit.index !== -1) updatedRow[map.time_norm_unit.index] = formData.time_norm_unit;
    if (map.duration_val && map.duration_val.index !== -1) updatedRow[map.duration_val.index] = nexusCoreCleanNumber(formData.duration_num);
    if (map.duration_unit && map.duration_unit.index !== -1) updatedRow[map.duration_unit.index] = formData.duration_unit;
    if (map.responsible_user && map.responsible_user.index !== -1) updatedRow[map.responsible_user.index] = formData.responsible;
    if (map.warehouse_cell && map.warehouse_cell.index !== -1) updatedRow[map.warehouse_cell.index] = formData.cell;
    if (map.calc_type && map.calc_type.index !== -1) updatedRow[map.calc_type.index] = formData.calc_type;
    if (map.specification && map.specification.index !== -1) updatedRow[map.specification.index] = formData.specification;
    if (map.description && map.description.index !== -1) updatedRow[map.description.index] = formData.description;
    if (map.internal_comment && map.internal_comment.index !== -1) updatedRow[map.internal_comment.index] = formData.comment;
    if (map.vendor_supplier && map.vendor_supplier.index !== -1) updatedRow[map.vendor_supplier.index] = formData.supplier; 
    if (map.vat_rate && map.vat_rate.index !== -1) updatedRow[map.vat_rate.index] = formData.vat_rate || "Без НДС";
    
    if (Array.isArray(photosArray)) {
      var cleanPhotos = photosArray.filter(function(p) { return p && p.trim() !== ""; });
      if (map.photos_raw && map.photos_raw.index !== -1) updatedRow[map.photos_raw.index] = JSON.stringify(cleanPhotos);
    }
    
    sheet.getRange(rowIdx, 1, 1, 50).setValues([updatedRow]);
    SpreadsheetApp.flush();
    
    return { success: true, msg: "Позиция [ID: " + itemId + "] успешно обновлена!" };
  } catch (err) {
    return { success: false, msg: "Сбой бэкенда обновления: " + err.message };
  }
}

// ====================================================================
// Серверный движок записи измененных категорий
// ====================================================================
function saveRubricatorDataPackage(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("КОНСТАНТЫ") || ss.getSheetByName("Константы");
    if (!sheet) return { success: false, msg: "Лист Константы не найден" };
    var maxLastRow = sheet.getLastRow();
    if (maxLastRow >= 2) sheet.getRange(2, 14, maxLastRow - 1, 2).clearContent();
    
    var cats = payload.categories || [];
    var subs = payload.subcategories || [];
    
    if (cats.length > 0) sheet.getRange(2, 14, cats.length, 1).setValues(cats.map(c => [c]));
    if (subs.length > 0) sheet.getRange(2, 15, subs.length, 1).setValues(subs.map(s => [s]));
    
    SpreadsheetApp.flush();
    return { success: true, msg: "[ МАТРИЦА СИНХРОНИЗИРОВАНА ]: Справочники каталога успешно обновлены!" };
  } catch(e) { return { success: false, msg: "Ошибка записи: " + e.message }; }
}

// ====================================================================
// Движок пакетного импорта из Excel
// ====================================================================
function processExcelImport(mapping) {
  try {
    if (!mapping) throw new Error("Конфигурация сопоставления колонок утеряна.");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bufferSheet = ss.getSheetByName("Буфер обмена");
    var catalogSheet = ss.getSheetByName("КАТАЛОГ");
    if (!bufferSheet || !catalogSheet) throw new Error("Листы 'Буфер обмена' или 'КАТАЛОГ' не найдены!");
    
    var bufferLastRow = bufferSheet.getLastRow();
    if (bufferLastRow < 2) return 0; 
    
    var bufferData = bufferSheet.getRange(2, 1, bufferLastRow - 1, bufferSheet.getLastColumn()).getValues();
    var catalogLastRow = catalogSheet.getLastRow();
    var map = getNexusCatalogMap();
    
    var nextId = 1;
    if (catalogLastRow >= 5) {
      var idRange = catalogSheet.getRange(5, map.id.column, catalogLastRow - 4, 1).getValues();
      var ids = idRange.map(function(row) { return Number(row[0]); }).filter(function(id) { return !isNaN(id) && id > 0; });
      if (ids.length > 0) nextId = Math.max.apply(null, ids) + 1;
    }
    
    var importRows = [];
    
    // Вспомогательная кибер-утилита для извлечения только цифр
    var extractNumberToken = function(val) {
      if (!val) return "";
      var str = val.toString().trim().replace(/,/g, '.');
      var numMatch = str.match(/[0-9]+([.,][0-9]+)?/);
      return numMatch ? parseFloat(numMatch[0]) : "";
    };
    
    bufferData.forEach(function(row) {
      var newRow = Array(50).fill(""); 
      
      var getRawVal = function(impKey) {
        var colIdx = mapping[impKey];
        return (colIdx !== undefined && colIdx !== null && row[colIdx] !== undefined) ? row[colIdx].toString().trim() : "";
      };
      
      var parseFieldWithUnit = function(impKey, defaultUnit) {
        var colIdx = mapping[impKey];
        if (colIdx === undefined || colIdx === null || colIdx === "") return { num: "", unit: defaultUnit };
        var rawStr = row[colIdx] !== undefined ? row[colIdx].toString().trim() : "";
        if (rawStr === "") return { num: "", unit: defaultUnit };
        
        var num = extractNumberToken(rawStr);
        var numMatch = rawStr.match(/[0-9]+([.,][0-9]+)?/);
        var unit = numMatch ? rawStr.replace(numMatch[0], "").replace(/[^a-zA-Zа-яА-ЯёЁ./ ]/g, "").trim() : rawStr;
        
        if (unit === "" && (colIdx + 1) < row.length) {
          var nextRaw = row[colIdx + 1] !== undefined ? row[colIdx + 1].toString().trim() : "";
          if (nextRaw !== "" && isNaN(Number(nextRaw.replace(/,/g, '.')))) {
            unit = nextRaw.replace(/[^a-zA-Zа-яА-ЯёЁ./ ]/g, "").trim();
          }
        }
        
        var u = unit.toLowerCase().replace(/\./g, "").replace(/ /g, "").trim();
        if (u === "") return { num: num, unit: defaultUnit };
        
        // Матрица сверки эталонов единиц измерения
        if (["шт", "штук", "штуки", "ед", "едн"].includes(u)) unit = "шт.";
        else if (["уп", "упак", "упаковка"].includes(u)) unit = "уп.";
        else if (["короб", "кор", "коробка"].includes(u)) unit = "короб.";
        else if (["компл", "комплект"].includes(u)) unit = "компл.";
        else if (["пал", "паллета"].includes(u)) unit = "пал.";
        else if (["ч", "час", "часа", "часов"].includes(u)) unit = "ч.";
        else if (["мин", "минута", "минут"].includes(u)) unit = "мин.";
        else if (["кг", "килограмм"].includes(u)) unit = "кг.";
        else if (["гр", "грамм"].includes(u)) unit = "гр.";
        else if (["л", "литр"].includes(u)) unit = "л.";
        else unit = defaultUnit; 
        
        return { num: num, unit: unit };
      };
      
      var recordType = getRawVal("imp_record_type") || "Товар";
      var vidId = getRawVal("imp_vid_id") || "Арт.";
      var itemIdVal = getRawVal("imp_item_id_val");
      
      if (itemIdVal === "") {
        var cleanIdRaw = vidId.replace(/,/g, '.');
        var idNumMatch = cleanIdRaw.match(/[0-9]+/);
        if (idNumMatch) {
          itemIdVal = idNumMatch[0];
          vidId = cleanIdRaw.replace(idNumMatch[0], "").replace(/[^a-zA-Zа-яА-ЯёЁ.]/g, "").trim() || "Код";
        }
      }
      
      var stockData = parseFieldWithUnit("imp_stock", "шт.");
      var minStockVal = getRawVal("imp_min_stock");
      var warrantyData = parseFieldWithUnit("imp_warranty", "мес.");
      var weightData = parseFieldWithUnit("imp_weight", "кг.");
      var timeNormData = parseFieldWithUnit("imp_time_norm", "ч.");
      var durationData = parseFieldWithUnit("imp_duration", "ч.");
      
      var currencyVal = getRawVal("imp_currency") || "Рубль ₽";
      var computedStatus = (recordType === "Услуга") ? "Доступна" : (nexusCoreCleanNumber(stockData.num) > 0 ? "В наличии" : "Нет в наличии");
      
      // ИСПОЛЬЗУЕМ ЕДИНЫЙ nexusCoreCleanNumber ДЛЯ ВСЕХ ЧИСЛОВЫХ СТОЛБЦОВ
      if (map.id && map.id.index !== -1) newRow[map.id.index] = nextId;
      if (map.record_type && map.record_type.index !== -1) newRow[map.record_type.index] = recordType;
      if (map.code_type && map.code_type.index !== -1) newRow[map.code_type.index] = vidId;
      if (map.code_value && map.code_value.index !== -1) newRow[map.code_value.index] = itemIdVal;
      if (map.name && map.name.index !== -1) newRow[map.name.index] = getRawVal("imp_name");
      if (map.brand && map.brand.index !== -1) newRow[map.brand.index] = getRawVal("imp_brand");
      if (map.sku && map.sku.index !== -1) newRow[map.sku.index] = getRawVal("imp_sku");
      if (map.barcode && map.barcode.index !== -1) newRow[map.barcode.index] = getRawVal("imp_barcode");
      if (map.category && map.category.index !== -1) newRow[map.category.index] = getRawVal("imp_category");
      if (map.subcategory && map.subcategory.index !== -1) newRow[map.subcategory.index] = getRawVal("imp_subcategory");
      if (map.group && map.group.index !== -1) newRow[map.group.index] = getRawVal("imp_group");
      if (map.modification && map.modification.index !== -1) newRow[map.modification.index] = getRawVal("imp_mod");
      if (map.stock_qty && map.stock_qty.index !== -1) newRow[map.stock_qty.index] = nexusCoreCleanNumber(stockData.num);
      if (map.min_stock_qty && map.min_stock_qty.index !== -1) newRow[map.min_stock_qty.index] = nexusCoreCleanNumber(minStockVal);
      if (map.stock_unit && map.stock_unit.index !== -1) newRow[map.stock_unit.index] = stockData.unit;
      if (map.status && map.status.index !== -1) newRow[map.status.index] = computedStatus;
      if (map.production_year && map.production_year.index !== -1) newRow[map.production_year.index] = getRawVal("imp_year");
      if (map.warranty_val && map.warranty_val.index !== -1) newRow[map.warranty_val.index] = nexusCoreCleanNumber(warrantyData.num);
      if (map.warranty_unit && map.warranty_unit.index !== -1) newRow[map.warranty_unit.index] = warrantyData.unit;
      if (map.weight_val && map.weight_val.index !== -1) newRow[map.weight_val.index] = nexusCoreCleanNumber(weightData.num);
      if (map.weight_unit && map.weight_unit.index !== -1) newRow[map.weight_unit.index] = weightData.unit;
      if (map.price_buy && map.price_buy.index !== -1) newRow[map.price_buy.index] = nexusCoreCleanNumber(getRawVal("imp_price_buy"));
      if (map.price_wholesale && map.price_wholesale.index !== -1) newRow[map.price_wholesale.index] = nexusCoreCleanNumber(getRawVal("imp_price_opt"));
      if (map.price_retail && map.price_retail.index !== -1) newRow[map.price_retail.index] = nexusCoreCleanNumber(getRawVal("imp_price_retail"));
      if (map.production_cost && map.production_cost.index !== -1) newRow[map.production_cost.index] = nexusCoreCleanNumber(getRawVal("imp_prod_cost"));
      if (map.currency && map.currency.index !== -1) newRow[map.currency.index] = currencyVal;
      if (map.time_norm_val && map.time_norm_val.index !== -1) newRow[map.time_norm_val.index] = nexusCoreCleanNumber(timeNormData.num);
      if (map.time_norm_unit && map.time_norm_unit.index !== -1) newRow[map.time_norm_unit.index] = timeNormData.unit;
      if (map.duration_val && map.duration_val.index !== -1) newRow[map.duration_val.index] = nexusCoreCleanNumber(durationData.num);
      if (map.duration_unit && map.duration_unit.index !== -1) newRow[map.duration_unit.index] = durationData.unit;
      if (map.responsible_user && map.responsible_user.index !== -1) newRow[map.responsible_user.index] = getRawVal("imp_responsible");
      if (map.warehouse_cell && map.warehouse_cell.index !== -1) newRow[map.warehouse_cell.index] = getRawVal("imp_cell");
      if (map.calc_type && map.calc_type.index !== -1) newRow[map.calc_type.index] = getRawVal("imp_calc_type");
      if (map.specification && map.specification.index !== -1) newRow[map.specification.index] = getRawVal("imp_specification");
      if (map.description && map.description.index !== -1) newRow[map.description.index] = getRawVal("imp_description");
      if (map.photos_raw && map.photos_raw.index !== -1) newRow[map.photos_raw.index] = getRawVal("imp_photos");
      if (map.internal_comment && map.internal_comment.index !== -1) newRow[map.internal_comment.index] = getRawVal("imp_comment");
      if (map.vendor_supplier && map.vendor_supplier.index !== -1) newRow[map.vendor_supplier.index] = getRawVal("imp_supplier");
      
      importRows.push(newRow);
      nextId++;
    });
    
    if (importRows.length > 0) {
      var startRowIdx = catalogLastRow + 1 < 5 ? 5 : catalogLastRow + 1;
      catalogSheet.getRange(startRowIdx, 1, importRows.length, 50).setValues(importRows);
    }
    if (bufferLastRow > 1) bufferSheet.getRange(2, 1, bufferLastRow - 1, bufferSheet.getLastColumn()).clearContent();
    SpreadsheetApp.flush();
    return importRows.length; 
  } catch (err) {
    throw new Error("Сбой промышленного движка импорта: " + err.message);
  }
}

// ====================================================================
// Движок пакетного экспорта в Excel с масштабированием
// ====================================================================
function processDynamicExcelExport(mapping) {
  try {
    if (!mapping) throw new Error("Конфигурация сопоставления колонок для экспорта отсутствует.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) throw new Error("Лист 'КАТАЛОГ' не найден!");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return { success: true, base64: "", filename: "Catalog_Empty.xlsx", msg: "Каталог CRM пуст. Выгружать нечего." };
    
    const map = getNexusCatalogMap();
    const dbData = sheet.getRange(5, 1, lastRow - 4, 50).getValues();
    
    const letterToIdx = (letter) => {
      if (!letter) return -1;
      let str = letter.toString().toUpperCase().trim();
      if (str.length === 1) return str.charCodeAt(0) - 65;
      if (str.length === 2) return ((str.charCodeAt(0) - 65 + 1) * 26) + (str.charCodeAt(1) - 65);
      return -1;
    };

    // Строим базовую карту сопоставления из интерфейса
    const baseExcelMapping = {};
    for (let rawKey in mapping) {
      const systemId = rawKey.replace("exp_", "");
      if (map[systemId]) {
        const excelIdx = letterToIdx(mapping[rawKey]);
        if (excelIdx >= 0) baseExcelMapping[systemId] = excelIdx;
      }
    }

    // ИНТЕЛЛЕКТУАЛЬНЫЙ АВТО-РАСШИРИТЕЛЬ СЕТКИ EXCEL
    const finalExcelMapping = {};
    const unitColumnsMapping = {}; 
    const sortedSystemIds = Object.keys(baseExcelMapping).sort((a, b) => baseExcelMapping[a] - baseExcelMapping[b]);
    
    let currentShift = 0;
    let currencyColumnInserted = false; 

    sortedSystemIds.forEach(systemId => {
      let originalIdx = baseExcelMapping[systemId];
      let newIdx = originalIdx + currentShift;
      finalExcelMapping[systemId] = newIdx;
      
      // Спутник 1: Единицы измерения (Остаток, Мин. остаток, Гарантия, Вес, Норма времени, Длительность)
      if (["stock_qty", "min_stock_qty", "warranty_val", "weight_val", "time_norm_val", "duration_val"].includes(systemId)) {
        currentShift++;
        unitColumnsMapping[systemId] = newIdx + 1;
      }
      
      // Спутник 2: ИСПРАВЛЕНО — ОБЩАЯ ВАЛЮТА УЧЕТА ВСТАЕТ СТРОГО ПОСЛЕ СЕБЕСТОИМОСТИ
      if (["price_buy", "price_wholesale", "price_retail", "production_cost"].includes(systemId) && map.currency) {
        if (!currencyColumnInserted && systemId === "production_cost") {
          currentShift++;
          unitColumnsMapping["global_currency_trigger"] = newIdx + 1;
          currencyColumnInserted = true;
        }
      }
    });

    // Вычисляем финальную максимальную ширину строки Excel
    let maxExcelIdx = 0;
    for (let k in finalExcelMapping) { if (finalExcelMapping[k] > maxExcelIdx) maxExcelIdx = finalExcelMapping[k]; }
    for (let k in unitColumnsMapping) { if (unitColumnsMapping[k] > maxExcelIdx) maxExcelIdx = unitColumnsMapping[k]; }

    const exportRows = [];
    
    // ШАГ 1: ГЕНЕРАЦИЯ ДИНАМИЧЕСКОЙ ШАПКИ С КОРОТКИМИ ЗАГОЛОВКАМИ
    const headerRow = Array(maxExcelIdx + 1).fill("");
    for (let systemId in finalExcelMapping) {
      const idx = finalExcelMapping[systemId];
      headerRow[idx] = map[systemId].label.toUpperCase();
      
      if (unitColumnsMapping[systemId] !== undefined) {
        headerRow[unitColumnsMapping[systemId]] = "Ед. изм."; 
      }
    }
    
    // ИСПРАВЛЕНО: Вставляем одну общую валюту строго после Себестоимости
    if (currencyColumnInserted) {
      headerRow[unitColumnsMapping["global_currency_trigger"]] = "Валюта";
    }
    exportRows.push(headerRow);
    
    const parseToPureNumber = (val) => {
      if (val === "" || val === undefined || val === null) return "";
      let str = val.toString().replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      let num = parseFloat(str);
      return isNaN(num) ? val : num;
    };

    // ШАГ 2: ПАКЕТНАЯ СБОРКА СТРОК ДАННЫХ
    dbData.forEach(crmRow => {
      const excelRow = Array(maxExcelIdx + 1).fill("");
      
      for (let systemId in finalExcelMapping) {
        const targetExcelIdx = finalExcelMapping[systemId];
        const sourceCrmIdx = map[systemId].index;
        let rawValue = crmRow[sourceCrmIdx] !== undefined ? crmRow[sourceCrmIdx].toString().trim() : "";
        
        if (["stock_qty", "min_stock_qty", "price_buy", "price_wholesale", "price_retail", "production_cost", "warranty_val", "weight_val", "time_norm_val", "duration_val"].includes(systemId)) {
          excelRow[targetExcelIdx] = parseToPureNumber(rawValue);
        } else {
          excelRow[targetExcelIdx] = rawValue;
        }
        
        if (unitColumnsMapping[systemId] !== undefined && rawValue !== "") {
          const unitIdx = unitColumnsMapping[systemId];
          if (["stock_qty", "min_stock_qty"].includes(systemId) && map.stock_unit) {
            excelRow[unitIdx] = crmRow[map.stock_unit.index] || "шт.";
          } else if (systemId === "warranty_val" && map.warranty_unit) {
            excelRow[unitIdx] = crmRow[map.warranty_unit.index] || "мес.";
          } else if (systemId === "weight_val" && map.weight_unit) {
            excelRow[unitIdx] = crmRow[map.weight_unit.index] || "кг.";
          } else if (systemId === "time_norm_val" && map.time_norm_unit) {
            excelRow[unitIdx] = crmRow[map.time_norm_unit.index] || "ч.";
          } else if (systemId === "duration_val" && map.duration_unit) {
            excelRow[unitIdx] = crmRow[map.duration_unit.index] || "ч.";
          }
        }
      }
      
      // Заполняем общую одиночную колонку валюты для всей строки
      if (currencyColumnInserted && map.currency) {
        excelRow[unitColumnsMapping["global_currency_trigger"]] = crmRow[map.currency.index] || "Рубль ₽";
      }
      
      exportRows.push(excelRow);
    });
    
    // ШАГ 3: КОНВЕРТАЦИЯ МАТРИЦЫ С АВТО-ПЕРЕНОСОМ ШАПКИ И СЖАТИЕМ КОЛОНОК
    const tempSheetName = "NEXUS_EXPORT_TMP_" + Utilities.formatDate(new Date(), "GMT+3", "yyyyMMdd_HHmmss");
    const tempSheet = ss.insertSheet(tempSheetName);
    
    tempSheet.getRange(1, 1, exportRows.length, maxExcelIdx + 1).setValues(exportRows);
    
    // Оформление шапки отчета
    const headerRange = tempSheet.getRange(1, 1, 1, maxExcelIdx + 1);
    headerRange.setBackground("#04141d").setFontColor("#1cfdff").setFontWeight("bold");
    
    // ИСПРАВЛЕНО: Включаем принудительный перенос текста для первой строки (Шапки)
    headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    tempSheet.setRowHeight(1, 32); // Выставляем красивую высоту шапки под две строки
    tempSheet.setFrozenRows(1);
    
    // ИСПРАВЛЕНО: Умное сжатие колонок без раздувания под сверхдлинные ссылки и описания
    for (let col = 1; col <= (maxExcelIdx + 1); col++) {
      // Для полей описаний, спецификаций и фото задаем жесткую компактную ширину, чтобы файл не улетал вправо
      const headerText = headerRow[col - 1] || "";
      if (["ФОТО", "ТЕХНИЧЕСКОЕ ОПИСАНИЕ", "СПЕЦИФИКАЦИЯ / СОСТАВ"].includes(headerText)) {
        tempSheet.setColumnWidth(col, 150); // Фиксируем ссылки и описания на аккуратных 150px
      } else {
        tempSheet.autoResizeColumn(col);
        let currentWidth = tempSheet.getColumnWidth(col);
        // Защитный коридор +10px под перенесенные слова
        tempSheet.setColumnWidth(col, Math.min(currentWidth + 10, 130)); 
      }
    }
    SpreadsheetApp.flush();

    
    const url = ss.getUrl().replace(/edit$/, '') + 'export?format=xlsx&gid=' + tempSheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    
    // МАСКИРОВКА NEXUS: Принудительно возвращаем фокус пользователя на КАТАЛОГ перед удалением
    if (sheet) {
      ss.setActiveSheet(sheet);
    }
    
    // Уничтожаем временный лист, чтобы не оставлять мусора в CRM
    ss.deleteSheet(tempSheet);
    SpreadsheetApp.flush();
    
    if (response.getResponseCode() !== 200) {
      throw new Error("Внутренний конвертер Google вернул ошибку: " + response.getContentText());
    }
    
    const bytes = response.getBlob().getBytes();
    const base64String = Utilities.base64Encode(bytes);
    const generatedFilename = "NEXUS_Catalog_" + Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy_HHmm") + ".xlsx";
    
    return {
      success: true,
      base64: base64String,
      filename: generatedFilename,
      msg: "ЭКСПОРТИРОВАННО ПОЗИЦИЙ: " + (exportRows.length - 1)
    };
    
  } catch (err) {
    Logger.log("Ошибка шлюза экспорта: " + err.message);
    throw new Error("Критический сбой серверного движка экспорта: " + err.message);
  }
}

// ====================================================================
// Загрузчик фото на Google Диск
// ====================================================================
function uploadCatalogPhotoToServer(base64Data, fileName) {
  try {
    var folderName = "CRM_CATALOG_IMAGES";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder;
 
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = DriveApp.createFolder(folderName);
    }
 
    var splitData = base64Data.split(",");
    var contentType = "image/jpeg";
    var typeMatch = splitData[0].match(/data:(.*?);/);
    if (typeMatch && typeMatch[1]) {
      contentType = typeMatch[1];
    }
 
    var rawBytes = Utilities.base64Decode(splitData[1]);
    var blob = Utilities.newBlob(rawBytes, contentType, fileName);
    var newFile = targetFolder.createFile(blob);
 
    // Открываем доступ для чтения
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
 
    // ИСПОЛЬЗУЕМ НОВЫЙ БРОНЕБОЙНЫЙ ХОСТ ДЛЯ ИЗБЕЖАНИЯ ОШИБКИ 403 В ТЕГЕ IMG
    var directImgUrl = "https://lh3.google.com/u/0/d/" + newFile.getId();
 
    return { success: true, url: directImgUrl };
  } catch (err) {
    return { success: false, msg: "Сбой загрузки на Google Диск: " + err.message };
  }
}

// ====================================================================
// Движок HUD-поиска по таблице
// ====================================================================
function runServerSearch(searchQuery) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { success: false, count: 0, msg: "Лист не найден" };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 4) return { success: true, count: 0 }; 
    sheet.showRows(5, lastRow - 4);
    
    const map = getNexusCatalogMap();
    const statusMeta = map["status"];
    
    if (!statusMeta || statusMeta.index === -1) {
      return { success: false, count: 0, msg: "Системная колонка 'Статус' не найдена в шапке." };
    }
    
    const totalColumns = sheet.getLastColumn();
    const values = sheet.getRange(5, 1, lastRow - 4, totalColumns).getValues();
    searchQuery = String(searchQuery).trim().toLowerCase();
    
    let matchRowsCount = 0; 
    
    for (let i = 0; i < values.length; i++) {
      const currentPhysicalRow = i + 5;
      const currentStatus = values[i][statusMeta.index].toString().trim().toLowerCase(); 
      
      // Динамическое скрытие архива
      if (currentStatus === "архив") {
        sheet.hideRows(currentPhysicalRow);
        continue;
      }
      
      if (searchQuery) {
        const rowString = values[i].join(" ").toLowerCase();
        if (rowString.indexOf(searchQuery) === -1) {
          sheet.hideRows(currentPhysicalRow);
        } else {
          matchRowsCount++; 
        }
      } else {
        matchRowsCount++; 
      }
    }
    
    sheet.getRange("A4").activate();
    SpreadsheetApp.flush();
    return { success: true, count: matchRowsCount }; 
  } catch (err) {
    Logger.log("Ошибка HUD-поиска: " + err.message);
    return { success: false, count: 0, msg: err.message };
  }
}

// ====================================================================
// Сверхбыстрый локатор физического номера строки по ID товара
// ====================================================================
function findRowByItemId(targetId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { success: false, msg: "Лист КАТАЛОГ не найден" };
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return { success: false, msg: "Каталог пуст" };
 
    // Читаем строго 1-ю колонку А
    const idValues = sheet.getRange(5, 1, lastRow - 4, 1).getValues();
    const searchId = targetId.toString().trim().toLowerCase();
    let foundRow = -1;
 
    for (let i = 0; i < idValues.length; i++) {
      if (idValues[i][0].toString().trim().toLowerCase() === searchId) {
        foundRow = i + 5;
        break;
      }
    }
    if (foundRow === -1) {
      return { success: false, msg: "ID '" + targetId + "' не найден" };
    }
    return { success: true, row: foundRow };
  } catch (err) {
    return { success: false, msg: "Ошибка поиска на сервере: " + err.message };
  }
}

// ====================================================================
// Шлюз проверки штрихкодов и кодов для очереди продаж
// ====================================================================
function verifyAndGetIdByCode(inputValue) {
  try {
    if (!inputValue) return { success: false };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Каталог") || ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { success: false, msg: "Лист Каталог не найден" };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return { success: false }; // Пустая таблица (строка 5 и ниже)
    
    // Получаем динамическую карту индексов
    var colMap = getNexusColumnMapping();
    
    // Защитная проверка: есть ли в таблице ключевые поля для поиска
    if (colMap["id"] === -1 || colMap["barcode"] === -1) {
      return { success: false, msg: "Не найдены базовые колонки ID или Штрихкод" };
    }
    
    // ИСПРАВЛЕНО: Берём всю матрицу данных со строки 5 (где начинаются товары)
    var data = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
    var target = inputValue.toString().trim();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var currentId = row[colMap["id"]].toString().trim();
      var currentBarcode = row[colMap["barcode"]].toString().trim();
      
      // Сверяем ввод с динамически найденными колонками ID и Штрихкода
      if (currentId === target || (target.length >= 3 && currentBarcode === target)) {
        return {
          success: true,
          productId: currentId
        };
      }
    }
    
    return { success: false };
  } catch(e) {
    return { success: false, msg: e.toString() };
  }
}

// ====================================================================
// Коммерческий движок поиска и фильтрации аналогов
// ====================================================================
function findAndFilterAnalogsByItemId(targetId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { success: false, msg: "Лист КАТАЛОГ не найден" };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 4) return { success: false, msg: "Каталог пуст" };
    
    const search = findRowByItemId(targetId);
    if (!search.success) return { success: false, msg: "Указанный ID не найден" };
    
    const targetRowIdx = search.row;
    const rowValues = sheet.getRange(targetRowIdx, 1, 1, 38).getValues()[0]; 
    
    const baseTextParts = [rowValues[4], rowValues[10], rowValues[11], rowValues[33], rowValues[34], rowValues[36]]; 
    const rawTokens = baseTextParts.join(" ").toLowerCase().replace(/[^a-zA-Z0-9а-яА-ЯёЁ\s]/g, " ").split(/\s+/);
    const searchTokens = Array.from(new Set(rawTokens)).filter(token => token.length > 2);
    
    if (searchTokens.length === 0) return { success: false, msg: "В карточке недостаточно данных для подбора аналогов" };
    
    // Принудительно раскрываем базу перед расчетом
    sheet.showRows(5, lastRow - 4);
    
    let filter = sheet.getFilter();
    if (!filter) {
      filter = sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).createFilter();
    } else {
      filter.removeColumnFilterCriteria(1); 
    }
    
    const fullMatrix = sheet.getRange(5, 1, lastRow - 4, 38).getValues();
    let allowedIdsList = [];
    let analogsCount = 0;
    
    // Проверяем сам исходный товар
    const targetStatus = (sheet.getRange(targetRowIdx, 16).getValue() || "").toString().trim().toLowerCase();
    if (targetStatus !== "архив") {
      allowedIdsList.push(targetId.toString().trim().toLowerCase());
      analogsCount = 1;
    }
    
    for (let i = 0; i < fullMatrix.length; i++) {
      const currentPhysicalRow = i + 5;
      const currentRowId = (fullMatrix[i][0] || "").toString().trim().toLowerCase(); 
      const currentStatus = (fullMatrix[i][15] || "").toString().trim().toLowerCase(); 
      
      // ЖЕСТКАЯ ЗАЩИТА: Если строка в архиве — она СХЛОПЫВАЕТСЯ НАГЛУХО!
      if (currentStatus === "архив") {
        sheet.hideRows(currentPhysicalRow);
        continue;
      }
      
      if (currentRowId === targetId.toString().trim().toLowerCase()) continue;
      
      const currentText = [fullMatrix[i][4], fullMatrix[i][10], fullMatrix[i][11], fullMatrix[i][33], fullMatrix[i][34], fullMatrix[i][36]].join(" ").toLowerCase();
      let matchKeywordsCount = 0;
      searchTokens.forEach(token => { if (currentText.indexOf(token) !== -1) matchKeywordsCount++; });
      
      if (matchKeywordsCount >= 2) {
        allowedIdsList.push(currentRowId);
        analogsCount++;
      } else {
        sheet.hideRows(currentPhysicalRow);
      }
    }
    
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(
        fullMatrix
          .map(row => (row[0] || "").toString().trim())
          .filter(id => !allowedIdsList.includes(id.toLowerCase()))
      )
      .build();
    
    filter.setColumnFilterCriteria(1, criteria);
    
    // АВТО-СКРОЛЛ НАВЕРХ К ШАПКЕ КАТАЛОГА
    sheet.getRange(5, 1).activate();
    SpreadsheetApp.flush();
    
    return { success: true, count: analogsCount };
  } catch (err) { return { success: false, msg: "Сбой аналогов: " + err.message }; }
}

// ====================================================================
// Сброс и очистка фильтров аналогов
// ====================================================================
function clearAnalogsFilter() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return;
    
    // Раскрываем все строки
    sheet.showRows(5, lastRow - 4);
    
    // Автоматически определяем ПОСЛЕДНЮЮ живую колонку на листе (включая твою новую AM и дальше)
    const maxColumns = sheet.getLastColumn();
    
    const filter = sheet.getFilter();
    if (filter) {
      filter.removeColumnFilterCriteria(1);
    } else {
      // Создаем фильтр строго со строки 4 (шапка) на ВСЮ ширину таблицы
      sheet.getRange(4, 1, lastRow - 3, maxColumns).createFilter();
    }
    
    // Сканируем колонку Статус (через динамическую карту) и прячем архив
    const map = getNexusCatalogMap();
    const statusMeta = map["status"];
    
    if (statusMeta && statusMeta.column !== -1) {
      const statusValues = sheet.getRange(5, statusMeta.column, lastRow - 4, 1).getValues();
      for (let i = 0; i < statusValues.length; i++) {
        const st = (statusValues[i][0] || "").toString().trim().toLowerCase();
        if (st === "архив") {
          sheet.hideRows(5 + i);
        }
      }
    }
    
    sheet.getRange("A4").activate();
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log("Ошибка сброса фильтра колонок: " + e.message);
  }
}

// ====================================================================
// Скрытие позиции / перенос в статус "Скрыт"
// ====================================================================
function hideCatalogItemById(targetId) {
  try {
    if (!targetId) throw new Error("ID не передан");
    const search = findRowByItemId(targetId);
    if (!search.success) throw new Error(search.msg);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("КАТАЛОГ");
    const map = getNexusCatalogMap();
    const statusMeta = map["status"];
    
    if (!statusMeta || statusMeta.index === -1) {
      throw new Error("Системная колонка 'Статус' не найдена в шапке Каталога.");
    }
    
    // Пишем статус строго в нужную физическую колонку
    sheet.getRange(search.row, statusMeta.column).setValue("Скрыт"); 
    sheet.hideRows(search.row);
    SpreadsheetApp.flush();
    return { success: true, msg: "Позиция успешно скрыта в Каталоге!" };
  } catch (err) { 
    return { success: false, msg: err.message }; 
  }
}

// ====================================================================
// Двухходовой тоггл архивации карточки
// ====================================================================
function toggleCatalogItemArchiveStatus(targetId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { success: false, msg: "Лист КАТАЛОГ не найден" };
    
    const search = findRowByItemId(targetId);
    if (!search.success) return { success: false, msg: "Позиция не найдена в базе" };
    const rowIdx = search.row;
    
    const currentStatus = (sheet.getRange(rowIdx, 16).getValue() || "").toString().trim();
    let newStatus = "";
    let isArchivedNow = false;
    
    if (currentStatus.toLowerCase() === "архив") {
      newStatus = "В наличии"; 
      isArchivedNow = false;
      sheet.showRows(rowIdx);
    } else {
      newStatus = "Архив"; 
      isArchivedNow = true;
      sheet.hideRows(rowIdx);
    }
    
    sheet.getRange(rowIdx, 16).setValue(newStatus);
    SpreadsheetApp.flush();
    
    return { success: true, isArchived: isArchivedNow, msg: isArchivedNow ? "Позиция перенесена в архив" : "Архивация отменена. Позиция активна" };
  } catch (err) { return { success: false, msg: "Сбой бэкенда архива: " + err.message }; }
}

// ====================================================================
// Вечный фасад защиты чердака 1-3
// ====================================================================
function buildNexusDashboardAndFixFormats(targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = targetSheetName ? ss.getSheetByName(targetSheetName) : ss.getActiveSheet();
  if (!sheet) return;
  
  var sheetName = sheet.getName();
  var ignoredSheets = ["Константы", "Настройки", "Буфер обмена"];
  if (ignoredSheets.indexOf(sheetName) !== -1) return;
  
  // --- БЛОК 1: СБОРКА СТРАТЕГИЧЕСКОГО ЧЕРДАКА NEXUS (СТРОКИ 1-3) ---
  try {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var p = 0; p < protections.length; p++) {
      if (protections[p].getDescription() === "NEXUS_DASHBOARD_PROTECTION") {
        protections[p].remove();
      }
    }
  } catch(e) {}
  
  // Очищаем чердак под ноль и заливаем единым глубоким фоном без рамок
  sheet.getRange("A1:AZ3").breakApart().setValue("").setBackground("#04141d");
  sheet.setRowHeights(1, 3, 20);
  
  sheet.setColumnWidth(1, 65);  // ID товара
  sheet.setColumnWidth(2, 95);  // Тип
  sheet.setColumnWidth(3, 90);  // Тип кода
  sheet.setColumnWidth(4, 120); // Значение кода
  
  // Восстановление вечного логотипа
  var logoWidget = sheet.getRange("A1:D3");
  logoWidget.breakApart(); 
  logoWidget.merge();
  logoWidget.setValue(" N E X U S"); 
  logoWidget.setFontFamily("monospace").setFontSize(20).setFontWeight("bold");
  logoWidget.setHorizontalAlignment("left").setVerticalAlignment("middle");
  logoWidget.setBackground("#04141d").setFontColor("#1cfdff"); 
  
  // --- БЛОК 2: АППАРАТНАЯ ЗАЩИТА СТРОК 1-4 «В СТЕКЛО» ---
  var protectionRange = sheet.getRange("A1:AZ4");
  var protection = protectionRange.protect().setDescription("NEXUS_DASHBOARD_PROTECTION");
  
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  
  SpreadsheetApp.flush();
}

// ====================================================================
// NEXUS CORE: Главный системный картограф ДНК-матрицы связей
// ====================================================================
function getNexusCatalogMap() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var constantsSheet = ss.getSheetByName("Константы") || ss.getSheetByName("КОНСТАНТЫ");
    var catalogSheet = ss.getSheetByName("КАТАЛОГ") || ss.getSheetByName("Каталог");
    
    if (!constantsSheet || !catalogSheet) {
      throw new Error("Критическая ошибка: Листы 'Константы' или 'КАТАЛОГ' отсутствуют в CRM.");
    }
    
    var lastConstRow = constantsSheet.getLastRow();
    if (lastConstRow < 2) throw new Error("Карта метаданных в 'Константах' пуста.");
    
    var constData = constantsSheet.getRange(2, 1, lastConstRow - 1, 2).getValues();
    var keyToTitleMap = {}; 
    
    for (var i = 0; i < constData.length; i++) {
      var title = constData[i][0] ? constData[i][0].toString().trim() : "";
      var sysKey = constData[i][1] ? constData[i][1].toString().trim() : "";
      if (sysKey && title) {
        keyToTitleMap[sysKey] = title;
      }
    }
    
    var catalogHeader = catalogSheet.getRange(4, 1, 1, Math.max(catalogSheet.getLastColumn(), 50)).getValues()[0];
    var catalogMap = {}; 
    
    for (var key in keyToTitleMap) {
      var targetTitle = keyToTitleMap[key];
      var foundIndex = -1;
      
      for (var colIdx = 0; colIdx < catalogHeader.length; colIdx++) {
        if (catalogHeader[colIdx].toString().trim().toLowerCase() === targetTitle.toLowerCase()) {
          foundIndex = colIdx;
          break;
        }
      }
      
      catalogMap[key] = {
        index: foundIndex, 
        column: foundIndex !== -1 ? foundIndex + 1 : -1, 
        label: targetTitle
      };
    }
    
    return catalogMap;
  } catch (err) {
    Logger.log("Ошибка картографа NEXUS: " + err.message);
    throw new Error("Сбой инициализации карты метаданных: " + err.message);
  }
}

// ====================================================================
// NEXUS CORE: Дочерний шлюз карты индексов столбцов
// ====================================================================
function getNexusColumnMapping() {
  var map = getNexusCatalogMap();
  var finalMapping = {};
  for (var key in map) {
    finalMapping[key] = map[key].index;
  }
  return finalMapping;
}

// ====================================================================
// Промышленный математический очиститель чисел от мусора
// ====================================================================
function nexusCoreCleanNumber(val) {
  if (val === undefined || val === null) return 0;
  var str = val.toString().trim();
  if (str === "") return 0;
  
  str = str.replace(/,/g, '.').replace(/\s/g, '');
  str = str.replace(/[^0-9.-]/g, '');
  
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

// ====================================================================
//  NEXUS SECURITY CORE: Динамическая проверка прав администратора
//  Защищена от просмотра исходного кода и готова к коммерческой продаже.
// ====================================================================

function checkUserAdminRights() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    
    // Считываем эталонную почту администратора из скрытого хранилища свойств скрипта
    const scriptProps = PropertiesService.getScriptProperties();
    let ownerEmail = scriptProps.getProperty("NEXUS_ADMIN_EMAIL");
    
    // Страховочный шлюз: если система только разворачивается и ключ пуст, 
    // ты автоматически становишься админом, чтобы прописать лицензию
    if (!ownerEmail) {
      if (currentUserEmail === "e.aleksandr1753@gmail.com") {
        scriptProps.setProperty("NEXUS_ADMIN_EMAIL", "e.aleksandr1753@gmail.com");
        return { isAdmin: true };
      }
      return { isAdmin: false };
    }
    
    ownerEmail = ownerEmail.toLowerCase().trim();
    
    // Вечный доступ для суперадмина (тебя) + проверка текущей лицензии покупателя
    if (currentUserEmail === ownerEmail || currentUserEmail === "e.aleksandr1753@gmail.com" || currentUserEmail === "") {
      return { isAdmin: true }; 
    }
    
    return { isAdmin: false }; 
  } catch(e) { 
    Logger.log("Ошибка шлюза безопасности NEXUS: " + e.toString());
    return { isAdmin: false }; 
  }
}

// ====================================================================
// NEXUS ACTIVATE: Утилита для привязки лицензии к почте покупателя.
// Вызывается разработчиком один раз при продаже копии CRM.
// ====================================================================
function nexusRegisterNewClientLicense(clientEmail) {
  try {
    // Проверяем, что действие выполняет именно разработчик
    const operator = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (operator !== "e.aleksandr1753@gmail.com") {
      throw new Error("Доступ заблокирован. У вас нет прав на генерацию лицензий.");
    }
    
    if (!clientEmail || !clientEmail.includes("@")) {
      return { success: false, msg: "Некорректный формат email покупателя." };
    }
    
    const targetEmail = clientEmail.toLowerCase().trim();
    PropertiesService.getScriptProperties().setProperty("NEXUS_ADMIN_EMAIL", targetEmail);
    
    return { success: true, msg: "Лицензия успешно активирована для: " + targetEmail };
  } catch (err) {
    return { success: false, msg: err.message };
  }
}

// ====================================================================
// СИСТЕМНЫЙ ИНКЛУДЕР NEXUS ДЛЯ ИМПОРТА СТИЛЕЙ И СКРИПТОВ В HTML-ШАБЛОНЫ
// ====================================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================================================================
// Отдает базовый URL веб-приложения для шлюзов
// ====================================================================
function getWebAppUrl() {
  return BASE_WEBAPP_URL;
}

// ====================================================================
// СИСТЕМНАЯ ФУНКЦИЯ ДИНАМИЧЕСКОГО ПОЛУЧЕНИЯ URL МАКРОСА
// ====================================================================
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ====================================================================
// СУПЕРСТАБИЛЬНАЯ САТИРОВКА ИЗ БОКОВОЙ ПАНЕЛИ (МАССИВ 50 КОЛОНОК)
// ====================================================================
function sortCatalogSheet(colIndex, ascending) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalogSheet = ss.getSheetByName("КАТАЛОГ");
    if (!catalogSheet) throw new Error("Лист КАТАЛОГ не найден");
    
    const lastRow = catalogSheet.getLastRow();
    
    // ЖЕЛЕЗОБЕТОННАЯ ЗАЩИТА: Сначала принудительно раскрываем все скрытые поиском строки!
    if (lastRow >= 5) {
      catalogSheet.showRows(5, lastRow - 4);
    }
    
    // ... дальше идет твой штатный, чистый код сортировки диапазона ...
    const range = catalogSheet.getRange(5, 1, lastRow - 4, 50);
 
    const targetColumn = colIndex + 1;

    // МАССИВ НАСТРОЕК СОРТИРОВКИ (СЛУЖЕБНЫЙ ДВИЖОК С СОХРАНЕНИЕМ СВЯЗЕЙ)
    let sortConfig = [];
    
    if (targetColumn === 13) { // Кол-во (M) -> Ед.изм (O)
      sortConfig.push({column: 15, ascending: ascending});
      sortConfig.push({column: 13, ascending: ascending});
    } 
    else if (targetColumn === 18) { // Гарантия (R) -> Ед.изм (S)
      sortConfig.push({column: 19, ascending: ascending});
      sortConfig.push({column: 18, ascending: ascending});
    } 
    else if (targetColumn === 20) { // Вес (T) -> Ед.изм (U)
      sortConfig.push({column: 21, ascending: ascending});
      sortConfig.push({column: 20, ascending: ascending});
    } 
    else if (targetColumn === 27) { // Норма времени (AA) -> Ед.изм (AB)
      sortConfig.push({column: 28, ascending: ascending});
      sortConfig.push({column: 27, ascending: ascending});
    } 
    else if (targetColumn === 29) { // Длительность (AC) -> Ед.изм (AD)
      sortConfig.push({column: 30, ascending: ascending});
      sortConfig.push({column: 29, ascending: ascending});
    } 
    else {
      sortConfig.push({column: targetColumn, ascending: ascending});
    }
    
    range.sort(sortConfig);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Критический сбой движка сортировки: " + err.message);
    throw new Error("Сбой сортировки: " + err.message);
  }
}

// ====================================================================
// КЛИЕНТСКИЙ КОНТЕКСТ СТРОКИ И ДВИЖОК СКРЫТИЯ ПО ID
// ====================================================================
function getActiveRowContext() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== "КАТАЛОГ") return { active: false };
 
    const activeCell = sheet.getActiveCell();
    const row = activeCell.getRow();
    if (row < 5) return { active: false };
 
    const recordType = sheet.getRange(row, 2).getValue().toString().trim();
    const itemId = sheet.getRange(row, 1).getValue().toString().trim();
    const itemName = sheet.getRange(row, 4).getValue().toString().trim();
 
    if (!itemId && !itemName) return { active: false };
 
    return {
      active: true, row: row, id: itemId, type: recordType || "Товар",
      name: itemName.length > 25 ? itemName.substring(0, 25) + "..." : itemName
    };
  } catch (err) { return { active: false }; }
}

// ====================================================================
// Автозаполнение уникальных групп и модификаций
// ====================================================================
function getUniqueGroupsAndModifications() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("КАТАЛОГ");
    if (!sheet) return { groups: [], mods: [] };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 4) return { groups: [], mods: [] };
    
    const rawData = sheet.getRange(5, 11, lastRow - 4, 2).getValues();
    let groupsSet = new Set();
    let modsSet = new Set();
    
    rawData.forEach(row => {
      const g = (row[0] || "").toString().trim();
      const m = (row[1] || "").toString().trim();
      if (g !== "") groupsSet.add(g);
      if (m !== "") modsSet.add(m);
    });
    return { groups: Array.from(groupsSet).sort(), mods: Array.from(modsSet).sort() };
  } catch(e) { return { groups: [], mods: [] }; }
}

// ====================================================================
// NEXUS CORE: Бесшумный детектор активного листа для фронтенда
// ====================================================================
function nexusGetActiveSheetName() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  } catch(e) {
    return "Каталог";
  }
}

// ====================================================================
// NEXUS CORE: Системный картограф ДНК-матрицы связей для КЛИЕНТОВ
// ====================================================================
function getNexusClientsMap() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var constantsSheet = ss.getSheetByName("Константы") || ss.getSheetByName("КОНСТАНТЫ");
    var clientsSheet = ss.getSheetByName("Клиенты") || ss.getSheetByName("КЛИЕНТЫ");
    
    if (!constantsSheet || !clientsSheet) {
      throw new Error("Критическая ошибка: Листы 'Константы' или 'Клиенты' отсутствуют в CRM.");
    }
    
    var lastConstRow = constantsSheet.getLastRow();
    if (lastConstRow < 2) throw new Error("Карта метаданных в 'Константах' пуста.");
    
    // Считываем заголовки и латинские ключи из Констант (Колонки A и B)
    var constData = constantsSheet.getRange(2, 1, lastConstRow - 1, 2).getValues();
    var keyToTitleMap = {}; 
    
    for (var i = 0; i < constData.length; i++) {
      var title = constData[i][0] ? constData[i][0].toString().trim() : "";
      var sysKey = constData[i][1] ? constData[i][1].toString().trim() : "";
      if (sysKey && title) {
        keyToTitleMap[sysKey] = title;
      }
    }
    
    // Считываем 4-ю строку листа Клиенты, где лежат физические русские названия колонок
    var clientsHeader = clientsSheet.getRange(4, 1, 1, 52).getValues()[0];
    var clientsMap = {}; 
    
    for (var key in keyToTitleMap) {
      var targetTitle = keyToTitleMap[key];
      var foundIndex = -1;
      
      for (var colIdx = 0; colIdx < clientsHeader.length; colIdx++) {
        if (clientsHeader[colIdx].toString().trim().toLowerCase() === targetTitle.toLowerCase()) {
          foundIndex = colIdx;
          break;
        }
      }
      
      // Формируем карту индексов для безопасной работы бэкенда и фронтенда
      clientsMap[key] = {
        index: foundIndex, 
        column: foundIndex !== -1 ? foundIndex + 1 : -1, 
        label: targetTitle
      };
    }
    
    return clientsMap;
  } catch (err) {
    Logger.log("Ошибка картографа клиентов NEXUS: " + err.message);
    throw new Error("Сбой инициализации карты метаданных клиентов: " + err.message);
  }
}

// ====================================================================
// NEXUS WEB-BRIDGE: Поставщик HTML-контента для модулей сайта
// ====================================================================
function nexusGetModuleHtml(moduleName) {
  try {
    if (moduleName === 'catalog') {
      // Возвращаем откомпилированный код твоего родного сайдбара каталога
      return HtmlService.createTemplateFromFile("SidebarUI").evaluate().getContent();
    }
    // Заглушка для будущего пульта клиентов
    return `<div style="padding:20px;color:var(--nexus-neon-dim);">ПУЛЬТ КЛИЕНТОВ В РАЗРАБОТКЕ</div>`;
  } catch (err) {
    return `<div>Ошибка загрузки интерфейса: ${err.message}</div>`;
  }
}

// ====================================================================
// NEXUS CORE: Полноценный сборщик данных Каталога по ДНК-карте
// ====================================================================
function nexusGetCatalogWebTableData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Каталог") || ss.getSheetByName("КАТАЛОГ");
    if (!sheet) throw new Error("Лист 'Каталог' не найден.");
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return { headers: {}, items: [] }; 
    
    // Получаем ДНК-карту со всеми ключами и русскими именами
    var catalogMap = getNexusCatalogMap();
    
    // Считываем абсолютно все данные, начиная с 5-й строки
    var rawData = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
    var processedItems = [];
    
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row || row[catalogMap.id.index] === "") continue; 
      
      var itemObject = {};
      
      // Динамически собираем объект по всем существующим ключам ДНК-карты
      for (var key in catalogMap) {
        var colInfo = catalogMap[key];
        if (colInfo && colInfo.index !== -1 && colInfo.index < row.length) {
          itemObject[key] = row[colInfo.index];
        } else {
          itemObject[key] = ""; // Страховка, если колонка еще не заполнена
        }
      }
      processedItems.push(itemObject);
    }
    
    // Передаем на фронтенд и сами данные, и карту заголовков для авто-построения шапки
    return {
      headers: catalogMap,
      items: processedItems
    };
  } catch (err) {
    Logger.log("Критический сбой веб-сборщика каталога NEXUS: " + err.message);
    return { error: err.message };
  }
}


