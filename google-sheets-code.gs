const TASK_HEADERS = ['المعرّف', 'عنوان المهمة', 'الأولوية', 'الفئة', 'تاريخ الاستحقاق', 'مكتملة؟'];

// يعيد المهام بصيغة JSON أو JSONP حتى يستطيع الموقع قراءتها.
function doGet(event) {
  const callback = event && event.parameter ? event.parameter.callback : '';

  try {
    const tasks = readTasks_();
    return respond_({ ok: true, tasks: tasks }, callback);
  } catch (error) {
    return respond_({ ok: false, error: String(error) }, callback);
  }
}

// يحفظ القائمة القادمة من الموقع داخل ورقة المهام.
function doPost(event) {
  try {
    const payload = JSON.parse(event.parameter.payload || '{}');
    if (payload.action !== 'save' || !Array.isArray(payload.tasks)) {
      throw new Error('بيانات المهام غير صالحة.');
    }
    writeTasks_(payload.tasks);
    return respond_({ ok: true });
  } catch (error) {
    return respond_({ ok: false, error: String(error) });
  }
}

function respond_(data, callback) {
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// شغّل هذه الدالة مرة واحدة من محرر Apps Script المرتبط بالشيت.
function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('افتح السكربت من داخل ملف Google Sheets.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
  getTaskSheet_();
}

function getTaskSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('اربط هذا السكربت بملف Google Sheets أولًا.');

  const sheet = spreadsheet.getSheets()[0] || spreadsheet.insertSheet('المهام');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, TASK_HEADERS.length).setFontWeight('bold');
  } else {
    const firstCell = sheet.getRange(1, 1).getDisplayValue();
    if (firstCell !== TASK_HEADERS[0]) {
      throw new Error('استخدم قالب المهام المرفق دون تغيير عناوين الأعمدة.');
    }
  }
  return sheet;
}

function readTasks_() {
  const sheet = getTaskSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, TASK_HEADERS.length).getValues();
  const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  return rows.map(function (row, index) {
    const title = String(row[1] || '').trim();
    if (!title) return null;

    let id = String(row[0] || '').trim();
    if (!id) {
      id = Utilities.getUuid();
      sheet.getRange(index + 2, 1).setValue(id);
    }

    let dueDate = '';
    if (row[4] instanceof Date) {
      dueDate = Utilities.formatDate(row[4], timezone, 'yyyy-MM-dd');
    } else if (row[4]) {
      dueDate = String(row[4]).trim();
    }

    return {
      id: id,
      title: title,
      priority: ({ 'عالية': 'high', 'متوسطة': 'medium', 'منخفضة': 'low' })[String(row[2]).trim()] || 'medium',
      category: ({ 'شخصي': 'personal', 'عمل': 'work', 'إنجازات': 'achievement' })[String(row[3]).trim()] || 'personal',
      dueDate: dueDate,
      completed: row[5] === true || ['نعم', 'مكتملة', 'مكتمل', 'true', '1'].indexOf(String(row[5]).trim().toLowerCase()) !== -1
    };
  }).filter(function (task) {
    return task !== null;
  });
}

function writeTasks_(tasks) {
  const sheet = getTaskSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, TASK_HEADERS.length).clearContent();
  }
  if (tasks.length === 0) return;

  const rows = tasks.map(function (task) {
    return [
      task.id || Utilities.getUuid(),
      String(task.title || '').trim(),
      ({ high: 'عالية', medium: 'متوسطة', low: 'منخفضة' })[task.priority] || 'متوسطة',
      ({ personal: 'شخصي', work: 'عمل', achievement: 'إنجازات' })[task.category] || 'شخصي',
      task.dueDate || '',
      task.completed ? 'نعم' : 'لا'
    ];
  });
  sheet.getRange(2, 1, rows.length, TASK_HEADERS.length).setValues(rows);
}