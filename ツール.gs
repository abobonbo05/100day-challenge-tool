/**
 * メイン関数：シートのセットアップを実行します。
 * この関数を実行してください。
 */
function setup100DayChallenge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. 設定シートの準備 ---
  let settingsSheet = ss.getSheetByName('設定');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('設定', 0);
  }
  setupPlaceholderSettings(settingsSheet);


  // --- 2. 必要なシートのクリアと作成 ---
  const sheetNames = ['一覧', '1', '2', '3', '4', '5', '6', '7'];
  const allSheets = ss.getSheets();
  const allSheetNames = allSheets.map(sheet => sheet.getName());
  
  sheetNames.forEach(name => {
    if (allSheetNames.includes(name)) {
      const sheetToDelete = ss.getSheetByName(name);
      ss.deleteSheet(sheetToDelete);
    }
  });

  const createdSheets = {};
  sheetNames.forEach(name => {
    createdSheets[name] = ss.insertSheet(name, ss.getSheets().length); 
  });

  const listSheet = createdSheets['一覧'];

  // --- 3. 「一覧」シートの1-3行目を設定 ---
  Logger.log('「一覧」シートの1-3行目を設定中...');
  listSheet.getRange('A1').setValue('100日チャレンジ').setFontSize(27);
  
  const headerData = settingsSheet.getRange('B2:I2').getValues()[0];
  
  listSheet.getRange('A2').setValue(headerData[0]);
  listSheet.getRange('C2').setValue(headerData[1]);
  listSheet.getRange('E2').setValue(headerData[2]);
  listSheet.getRange('G2').setValue(headerData[3]);
  listSheet.getRange('A3').setValue(headerData[4]);
  listSheet.getRange('C3').setValue(headerData[5]);
  listSheet.getRange('E3').setValue(headerData[6]);
  listSheet.getRange('G3').setValue(headerData[7]);

  listSheet.getRange('A2:G3').setFontSize(27);


  // --- 4. 「一覧」シートの4行目以降（100日分）を設定 ---
  Logger.log('「一覧」シートの4行目以降（100日分）を設定中...');
  
  const itemHeaders = settingsSheet.getRange('B3:I3').getValues()[0];
  const dataMatrix = [
    [itemHeaders[0], itemHeaders[1]],
    [itemHeaders[2], itemHeaders[3]],
    [itemHeaders[4], itemHeaders[5]],
    [itemHeaders[6], itemHeaders[7]]
  ];

  let startDate;
  const settingsStartDateValue = settingsSheet.getRange('A2').getValue();

  if (settingsStartDateValue && settingsStartDateValue instanceof Date && !isNaN(settingsStartDateValue.getTime())) {
    startDate = settingsStartDateValue;
    Logger.log('設定シートA2の日付を開始日として使用します: ' + startDate);
  } else {
    startDate = new Date(); 
    Logger.log('設定シートA2が空または無効なため、本日の日付を開始日として使用します: ' + startDate);
  }

  const weekdays = ['(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'];
  const dataRangesToFormat = []; 

  for (let day = 1; day <= 100; day++) {
    const weekIndex = Math.floor((day - 1) / 7);
    const dayInWeekIndex = (day - 1) % 7;
    
    const headerRow = 4 + weekIndex * 5;
    const dataStartRow = headerRow + 1;
    const col = 1 + dayInWeekIndex * 2;
    
    listSheet.getRange(headerRow, col).setValue(`${day}日目`);
    
    const targetDate = new Date(startDate); 
    targetDate.setDate(startDate.getDate() + (day - 1));
    
    const dateString = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'M/d');
    const dayOfWeek = weekdays[targetDate.getDay()];
    const fullDateHeader = `${dateString}${dayOfWeek}`;
    
    listSheet.getRange(headerRow, col + 1).setValue(fullDateHeader);
    dataRangesToFormat.push(listSheet.getRange(headerRow, col, 1, 2)); 

    const dataRange = listSheet.getRange(dataStartRow, col, 4, 2);
    dataRange.setValues(dataMatrix);
    dataRangesToFormat.push(dataRange);
  }

  dataRangesToFormat.forEach(range => range.setFontSize(14));


  // --- 5. シート「1」〜「7」に転記 ---
  Logger.log('シート「1」〜「7」に転記中...');
  
  const listHeaderRange = listSheet.getRange('A1:G3');
  const maxCol = 14; 
  const targetRowHeight = 21 * 3; // 63px (標準21pxの3倍)
  
  for (let i = 1; i <= 7; i++) {
    const sheet = createdSheets[String(i)]; 
    
    listHeaderRange.copyTo(sheet.getRange('A1'));
    
    const sourceStartRow = (i - 1) * 10 + 4;

    let numRows;
    if (i === 7) {
      numRows = 15; // Days 85-100
    } else {
      numRows = 10; // 2週間分
    }
    
    if (numRows > 0) {
      const sourceRange = listSheet.getRange(sourceStartRow, 1, numRows, maxCol);
      sourceRange.copyTo(sheet.getRange(4, 1));
      
      // 4行目以降の行の高さを3倍に
      sheet.setRowHeights(4, numRows, targetRowHeight);
    }
  }
  
  // --- 6. 列幅の調整 (すべて150pxに統一) ---
  Logger.log('列幅を調整中...');
  const targetColumnWidth = 150; 
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { 
      sheet.setColumnWidths(1, maxCol, targetColumnWidth); 
    }
  });

  Logger.log('すべての設定が完了しました。');
}


/**
 * ヘルパー関数：
 * '設定'シートが空の場合に、見本のヘッダー情報を入力します。
 */
function setupPlaceholderSettings(settingsSheet) {
  try {
    if (settingsSheet.getRange('B2').getValue() === '') {
      
      const spec3Data = [
        ['今日の目標', '目標1', '目標2', '目標3', '今日の振り返り', '振り返り1', '振り返り2', '振り返り3']
      ];
      settingsSheet.getRange('B2:I2').setValues(spec3Data);
      
      const spec2Data = [
        ['タスク1', 'ステータス1', 'タスク2', 'ステータス2', 'タスク3', 'ステータス3', 'タスク4', 'ステータス4']
      ];
      settingsSheet.getRange('B3:I3').setValues(spec2Data);
      
      settingsSheet.getRange('A1').setValue('【要確認】B2:I2 と B3:I3 に見本データを入力しました。チャレンジ内容に合わせて書き換えてください。');
      
      if (settingsSheet.getRange('A2').getValue() === '') {
        settingsSheet.getRange('A2').setValue('（ここに開始日 (例: 11/1) を入力。空欄の場合は実行日）');
        settingsSheet.getRange('A2').setFontColor('#888888').setFontStyle('italic');
      }
    }
  } catch (e) {
    Logger.log('プレースホルダー設定エラー: ' + e.message);
  }
}