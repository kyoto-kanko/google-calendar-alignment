function addEventsToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // カレンダーを取得
  var calendar = CalendarApp.getDefaultCalendar();

  // ヘッダー行をスキップするために2行目からループを開始
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var eventName = row[0];
    var startDate = new Date(row[1]);

    // 開始時間と終了時間を取得
    var startTime = getTimeParts(row[2]);
    var endTime = getTimeParts(row[3]);

    // 開始日と開始時間を組み合わせる
    startDate.setHours(startTime[0], startTime[1], 0);
    var startDateTime = startDate;

    // 終了時間を設定
    var endDate = new Date(row[1]); // 開始日と同じ日を基にする
    endDate.setHours(endTime[0], endTime[1], 0);
    var endDateTime = endDate;

    var location = row[4];
    var description = row[5];

    // イベントをカレンダーに追加
    calendar.createEvent(eventName, startDateTime, endDateTime, {
      location: location,
      description: description
    });
  }
}

// 時間の文字列またはDateオブジェクトから時間と分を取得する関数
function getTimeParts(time) {
  if (typeof time === 'string') {
    return time.split(":").map(Number);
  } else if (time instanceof Date) {
    return [time.getHours(), time.getMinutes()];
  } else {
    throw new Error("Invalid time format");
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // カスタムメニューを追加
  ui.createMenu('カレンダー登録')
      .addItem('イベントを追加', 'addEventsToCalendar')
      .addToUi();
}
