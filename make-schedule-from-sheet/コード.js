function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カレンダー連携')
    .addItem('ツールを開く', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('カレンダー一括登録')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSelectedPreview() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) throw new Error("範囲が選択されていません。");

    const values = range.getValues();
    if (values.length < 2 || values[0].length < 2) {
      throw new Error("日付と役割を含めて選択してください。");
    }

    let previewList = [];
    const headerRow = values[0]; 

    // 【修正】列（日付）を外側のループに、行（役割）を内側のループにすることで日付順にする
    for (let j = 1; j < values[0].length; j++) { // 列（日付）
      const dateVal = headerRow[j];
      const dateObj = new Date(dateVal);
      if (isNaN(dateObj.getTime())) continue;

      for (let i = 1; i < values.length; i++) { // 行（役割）
        const roleName = values[i][0];
        const personName = values[i][j];
        if (!personName || !roleName) continue;

        const dateString = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd");
        previewList.push({
          dateStr: dateString,
          dateValue: dateObj.getTime(), // ソート用に数値を持たせる
          dateObj: dateObj.toISOString(),
          title: `${roleName} ${personName}`
        });
      }
    }

    // 念のため日付順（昇順）にソートを確定させる
    previewList.sort((a, b) => a.dateValue - b.dateValue);

    return previewList;
  } catch (e) {
    throw new Error(e.message);
  }
}

function registerEvents(events) {
  const calendar = CalendarApp.getCalendarById('8f1bce7581c514e6e241989446b32497b561129a04bb9f806a8e4e396de7840d@group.calendar.google.com');
  let count = 0;
  events.forEach(ev => {
    const date = new Date(ev.dateObj);
    const existing = calendar.getEventsForDay(date);
    const isExist = existing.some(e => e.getTitle() === ev.title);
    if (!isExist) {
      calendar.createAllDayEvent(ev.title, date);
      count++;
    }
  });
  return count;
}