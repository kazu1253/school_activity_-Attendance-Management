function doGet(e) {
  if (e.parameter.mode === 'adviser') {
    return HtmlService.createHtmlOutputFromFile('AdviserForm')
        .setTitle('【顧問用】管理画面')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (e.parameter.mode === 'Viewer') {
    return HtmlService.createHtmlOutputFromFile('ViewerPage')
        .setTitle('【顧問用】出席状況確認')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('StudentApp')
      .setTitle('部活出席連絡')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/**
 * 先生が練習予定を保存する（日付列に自動で判別し自動追加）
 */
function saveSchedule(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('出席表');
  
  if (!sheet) {
    throw new Error('出席表シートが見つかりません');
  }
  
  // 単発登録の場合
  if (data.mode === 'single') {
    const targetDate = new Date(data.date);
    const targetCol = findOrCreateDateColumn(sheet, targetDate);
    
    // 2行目：練習可否（○/×）
    sheet.getRange(2, targetCol).setValue(data.practice || "○");
    
    // 3行目：時間
    sheet.getRange(3, targetCol).setValue(data.startTime + "～");
    
    // 4行目：顧問名
    sheet.getRange(4, targetCol).setValue(data.adviser || "");
    
    // 5行目：学校行事
    if (data.schoolEvent) {
      sheet.getRange(5, targetCol).setValue(data.schoolEvent);
    }
    
    // 6行目：大会
    if (data.tournament) {
      sheet.getRange(6, targetCol).setValue(data.tournament);
    }
    
    return "練習日を登録しました。";
  }
  
  // 一括登録の場合
  if (data.mode === 'bulk') {
    const start = new Date(data.startDate);
    const end = new Date(data.endDate);
    let count = 0;
    
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const dayOfWeek = d.getDay();
      const dayMap = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
      
      if (data.days.includes(dayMap[dayOfWeek])) {
        const targetCol = findOrCreateDateColumn(sheet, new Date(d));
        
        sheet.getRange(2, targetCol).setValue(data.practice || "○");
        sheet.getRange(4, targetCol).setValue(data.startTime + "～");
        sheet.getRange(5, targetCol).setValue(data.adviser || "");
        count++;
      }
    }
    return count + "日分の練習日を一括登録しました！";
  }
}

/**
 * 日付の列を探す、なければ作成する
 */
function findOrCreateDateColumn(sheet, targetDate) {
  const dateRow = 1; // 1行目が日付
  const lastCol = sheet.getLastColumn();
  const dateRange = sheet.getRange(dateRow, 1, 1, lastCol).getValues()[0];
  
  const targetDateStr = Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd");
  
  // 既存の日付列を探す
  for (let i = 0; i < dateRange.length; i++) {
    const cellDate = dateRange[i];
    if (cellDate instanceof Date) {
      const cellDateStr = Utilities.formatDate(cellDate, "JST", "yyyy-MM-dd");
      if (cellDateStr === targetDateStr) {
        return i + 1; // 見つかった列番号を返す
      }
    }
  }
  
  // 見つからなかった場合、新しい列を追加
  // E列(5列目)以降に日付を追加
  let insertCol = 5; // E列から開始
  
  // E列以降の最後の日付列を探す
  for (let col = 5; col <= lastCol; col++) {
    const cellValue = sheet.getRange(dateRow, col).getValue();
    if (!cellValue || cellValue === "") {
      insertCol = col;
      break;
    }
    if (cellValue instanceof Date) {
      insertCol = col + 1;
    }
  }
  
  // 日付が昇順になるように挿入位置を決定
  let finalCol = insertCol;
  for (let col = 5; col < insertCol; col++) {
    const cellDate = sheet.getRange(dateRow, col).getValue();
    if (cellDate instanceof Date) {
      const cellDateStr = Utilities.formatDate(cellDate, "JST", "yyyy-MM-dd");
      if (targetDateStr < cellDateStr) {
        // この位置に挿入すべき
        sheet.insertColumnBefore(col);
        finalCol = col;
        break;
      }
    }
  }
  
  // 最後の列に追加する場合
  if (finalCol === insertCol && finalCol > lastCol) {
    finalCol = lastCol + 1;
  }
  
  // 1行目に日付を入力
  const dateCell = sheet.getRange(dateRow, finalCol);
  dateCell.setValue(targetDate);
  dateCell.setNumberFormat("m/d");
  
  // 2〜7行目にラベルを設定（D列にある場合のみ）
  if (finalCol > 5) {
    // 既に他の日付があるので、ラベルはコピー不要
  } else if (finalCol === 5) {
    // 最初の日付列なので、D列にラベルを設定
    sheet.getRange(2, 4).setValue("練習");
    sheet.getRange(3, 4).setValue("時間");
    sheet.getRange(4, 4).setValue("顧問");
    sheet.getRange(5, 4).setValue("学校行事");
    sheet.getRange(6, 4).setValue("大会");
  }
  
  return finalCol;
}

/**
 * 学校行事・大会を保存する（日付列を自動追加）
 */
function saveEvent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('出席表');
  
  if (!sheet) {
    throw new Error('出席表シートが見つかりません');
  }
  
  const targetDate = new Date(data.date);
  const targetCol = findOrCreateDateColumn(sheet, targetDate);
  
  // 5行目：学校行事
  if (data.schoolEvent) {
    sheet.getRange(5, targetCol).setValue(data.schoolEvent);
  }
  
  // 6行目：大会
  if (data.tournament) {
    sheet.getRange(6, targetCol).setValue(data.tournament);
  }
  
  // 2行目：練習可否を×にする（行事や大会がある場合）
  if (data.schoolEvent || data.tournament) {
    sheet.getRange(2, targetCol).setValue("×");
  }
  
  return "イベントを登録しました。";
}


/**
 * 学校行事・大会を保存する
 */
function saveEvent(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('出席表');
  
  if (!sheet) {
    throw new Error('出席表シートが見つかりません');
  }
  
  const targetDate = new Date(data.date);
  const lastCol = sheet.getLastColumn();
  const dateRow = 1;
  const dateRange = sheet.getRange(dateRow, 1, 1, lastCol).getValues()[0];
  
  let targetCol = -1;
  for (let i = 0; i < dateRange.length; i++) {
    const cellDate = dateRange[i];
    if (cellDate instanceof Date) {
      const cellDateStr = Utilities.formatDate(cellDate, "JST", "yyyy-MM-dd");
      const targetDateStr = Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd");
      if (cellDateStr === targetDateStr) {
        targetCol = i + 1;
        break;
      }
    }
  }
  
  if (targetCol === -1) {
    return "指定された日付の列が見つかりません。";
  }
  
  // 6行目：学校行事
  if (data.schoolEvent) {
    sheet.getRange(6, targetCol).setValue(data.schoolEvent);
  }
  
  // 7行目：大会
  if (data.tournament) {
    sheet.getRange(7, targetCol).setValue(data.tournament);
  }
  
  // 2行目：練習可否を×にする（行事や大会がある場合）
  if (data.schoolEvent || data.tournament) {
    sheet.getRange(2, targetCol).setValue("×");
  }
  
  return "イベントを登録しました。";
}

/**
 * 週間予定と参加者リストを取得する
 */
/**
 * 週間予定と参加者リスト（時刻付き）を取得する
 */
function getWeeklySchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('出席表');
  
  if (!sheet) {
    console.log("出席表シートが存在しません");
    return [];
  }
  
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const nextWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7, 23, 59, 59);
  
  const dateRow = 1;      // 日付
  const practiceRow = 2;  // 練習可否
  const timeRow = 4;      // 時間
  const adviserRow = 5;   // 顧問
  const eventRow = 6;     // 学校行事
  const tournamentRow = 7;// 大会
  
  const lastCol = sheet.getLastColumn();
  
  const dates = sheet.getRange(dateRow, 1, 1, lastCol).getValues()[0];
  const practices = sheet.getRange(practiceRow, 1, 1, lastCol).getValues()[0];
  const times = sheet.getRange(timeRow, 1, 1, lastCol).getValues()[0];
  const advisers = sheet.getRange(adviserRow, 1, 1, lastCol).getValues()[0];
  const events = sheet.getRange(eventRow, 1, 1, lastCol).getValues()[0];
  const tournaments = sheet.getRange(tournamentRow, 1, 1, lastCol).getValues()[0];
  
  let result = [];
  
  for (let col = 1; col <= lastCol; col++) {
    const dateCell = dates[col - 1];
    const practiceCell = practices[col - 1];
    const timeCell = times[col - 1];
    const adviserCell = advisers[col - 1];
    const eventCell = events[col - 1];
    const tournamentCell = tournaments[col - 1];
    
    if (!dateCell) continue;
    
    let scheduleDate = null;
    
    if (dateCell instanceof Date) {
      scheduleDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
    } else {
      const match = dateCell.toString().match(/(\d+)\/(\d+)/);
      if (match) {
        const month = parseInt(match[1]) - 1;
        const day = parseInt(match[2]);
        scheduleDate = new Date(now.getFullYear(), month, day);
      }
    }
    
    if (!scheduleDate) continue;
    
    if (scheduleDate >= today && scheduleDate <= nextWeek) {
      const dateStr = Utilities.formatDate(scheduleDate, "JST", "yyyy-MM-dd");
      const dateLabel = Utilities.formatDate(scheduleDate, "JST", "M/d (E)");
      
      // 練習可否をチェック
      const isPractice = practiceCell && practiceCell.toString() !== "×";
      
      // 開始時刻を取得
      let startTimeStr = '';
      if (timeCell instanceof Date) {
        startTimeStr = Utilities.formatDate(timeCell, "JST", "HH:mm");
      } else if (timeCell && timeCell.toString().includes('～')) {
        startTimeStr = timeCell.toString().split('～')[0];
      } else if (timeCell) {
        startTimeStr = timeCell.toString();
      }
      
      // 顧問名を取得
      let adviserName = adviserCell ? adviserCell.toString() : "";
      if (!adviserName || adviserName.trim() === "") {
        adviserName = "自主練";
      }
      
      // イベント情報を取得
      const schoolEvent = eventCell ? eventCell.toString() : "";
      const tournament = tournamentCell ? tournamentCell.toString() : "";
      
      // 参加者と時刻を取得（8行目以降、D列に名前）
      const nameCol = 4; // D列
      const attendees = [];
      const attendeeTimes = [];
      
      for (let row = 8; row <= sheet.getLastRow(); row++) {
        const name = sheet.getRange(row, nameCol).getValue();
        const time = sheet.getRange(row, col).getValue();
        
        if (name && time) {
          attendees.push(name.toString());
          
          // 時刻のフォーマット
          let timeStr = '';
          if (time instanceof Date) {
            timeStr = Utilities.formatDate(time, "JST", "HH:mm");
          } else {
            timeStr = time.toString();
          }
          attendeeTimes.push(timeStr);
        }
      }
      
      result.push({
        date: dateStr,
        dateLabel: dateLabel,
        startTime: startTimeStr || '未定',
        adviser: adviserName,
        isPractice: isPractice,
        schoolEvent: schoolEvent,
        tournament: tournament,
        attendees: attendees,
        attendeeTimes: attendeeTimes
      });
    }
  }
  
  result.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  console.log("取得した予定:", result);
  return result;
}

/**
 * 生徒の出席データを記録する
 */
function recordAttendance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('出席表');
  
  if (!sheet) {
    throw new Error('出席表シートが見つかりません');
  }
  
  // 1. 日付の列を探す
  const targetDate = new Date(data.date);
  const lastCol = sheet.getLastColumn();
  const dateRow = 1;
  const dateRange = sheet.getRange(dateRow, 1, 1, lastCol).getValues()[0];
  
  let targetCol = -1;
  for (let i = 0; i < dateRange.length; i++) {
    const cellDate = dateRange[i];
    if (cellDate instanceof Date) {
      const cellDateStr = Utilities.formatDate(cellDate, "JST", "yyyy-MM-dd");
      const targetDateStr = Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd");
      if (cellDateStr === targetDateStr) {
        targetCol = i + 1;
        break;
      }
    }
  }
  
  if (targetCol === -1) {
    throw new Error('指定された日付の列が見つかりません');
  }
  
  // 2. 生徒名の行を探す（D列、8行目以降）
  const nameCol = 4; // D列
  const lastRow = sheet.getLastRow();
  const nameRange = sheet.getRange(8, nameCol, lastRow - 7, 1).getValues();
  
  let targetRow = -1;
  for (let i = 0; i < nameRange.length; i++) {
    const cellName = nameRange[i][0];
    if (cellName && cellName.toString().includes(data.name)) {
      targetRow = i + 8; // 8行目から開始
      break;
    }
  }
  
  if (targetRow === -1) {
    throw new Error('名前が見つかりません: ' + data.name);
  }
  
  // 3. 該当セルに時刻を書き込む
  sheet.getRange(targetRow, targetCol).setValue(data.time);
  
  return true;
}