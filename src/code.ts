// GASのスクリプトプロパティに適宜設定値を登録する
// カレンダーID
const calendarId = PropertiesService.getScriptProperties().getProperty('calendarId');
// scriptのデプロイID
const deployId = PropertiesService.getScriptProperties().getProperty('deployId');
// 通知メールの送信先アドレス
const notificationAddress = PropertiesService.getScriptProperties().getProperty('notificationAddress');
// 稽古日程メールのCC送信先アドレス
const notificationCcAddress = PropertiesService.getScriptProperties().getProperty('notificationCcAddress');

// getリクエストを受け取ってページを表示する関数
const doGet = (e) => {
    switch (e.parameter.action) {
        case "register":
            return HtmlService.createTemplateFromFile("register").evaluate();
        case "list":
        default:
            const sheetName = e.parameter.sheetName || `${getNow().year}${getPaddingNumber(getNow().month)}`;
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
            if (sheet === null) {
                throw new Error("該当月の稽古日程は未登録です。");
            }
            const values = sheet?.getDataRange().getValues()
            const template = HtmlService.createTemplateFromFile("list");
            template.links = values;
            template.deployId = deployId;
            template.sheetName = sheetName;
            // シート名の一覧を取得する
            const sheetNames = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(sheet => sheet.getName()).filter(sheet => sheet !== "template");
            template.sheetNames = sheetNames;
            return template.evaluate();
    }
}

// メンバーの追加: 暫定でGoogleフォームへのリンクを追加
// リストの見た目を整える・アイコンで表示
// TODO: 備品状況を端的に送信できる機能

// cssファイルを読み込むための関数
const include = (filename: string) => {
    const script = HtmlService.createHtmlOutputFromFile(filename)
      .getContent()
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">");
    return script;
  }

// html側のvalueと定義を揃える必要あり
const statusDefinition = {
    key: "鍵",
    join: "出",
    absent: "欠",
    late: "遅",
    early: "早"
};

// 現在年月を取得する関数
const getNow = () => {
    let now = new Date();
    let year = now.getFullYear();
    let month = now.getMonth() + 1;
    return { year, month };
}

const getPaddingNumber = (num: number) => {
    return String(num).padStart(2, '0');
}

// リクエストを受け取って出欠表の各自の行を更新する関数
const updateIndivisualRow = (sheetName: string, name: string, values: string[]) => {
    // 該当のシートを取得する
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet === null) {
        throw new Error("該当月の稽古日程は未登録です。");
    }
    // 指定された名前の行数を取得する
    let row = getRowByName(sheet, name);
    // 取得した行数が-1の場合はエラーにする
    if (row === -1) {
        throw new Error("指定された名前が存在しません。");
    }
    // 取得した行数の2列目以降を更新する
    return sheet.getRange(row + 1, 2, 1, values.length).setValues([values]);
}

// 指定した名前の行数を取得する関数
const getRowByName = (sheet: GoogleAppsScript.Spreadsheet.Sheet, name: string) => {
    let values = sheet.getDataRange().getValues();
    let row = values.findIndex((row) => row[0] === name);
    return row;
}


// カレンダーから指定された月のイベントを取得する関数
const getEvents = (year:number, month: number) => {
    if (calendarId === null) {
        return [];
    }
    let calendar = CalendarApp.getCalendarById(calendarId);
    let start = new Date(year, month - 1, 1);
    let end = new Date(year, month, 1);
    let events = calendar.getEvents(start, end);
    return events;
}

// カレンダーイベントプロパティから必要な項目を抽出して整形する関数
const getEventProps = (event: GoogleAppsScript.Calendar.CalendarEvent) => {
    const startDate = event.getStartTime();
    const endDate = event.getEndTime();
    const [name, place] = event.getTitle().split('@');
    return {
        date: getDateString(startDate),
        time: `${startDate.getHours()}:${getPaddingNumber(startDate.getMinutes())}-${endDate.getHours()}:${getPaddingNumber(endDate.getMinutes())}`,
        name: name.replace("創玄会", "") || "正規練",
        place: place ?? event.getLocation() ?? "場所情報なし"
    }
}

const getDateString = (date: Date | GoogleAppsScript.Base.Date) => {
    const weeks = ["日", "月", "火", "水", "木", "金", "土"];
    return `${date.getMonth() + 1}/${date.getDate()}(${weeks[date.getDay()]})`
}


// カレンダーのイベントを出欠表に反映する関数
// 稽古が追加された場合は列が挿入される・削除された場合は列が削除される
// 日時が変わらず稽古場所が変わった場合は列は保持され、場所のみ更新される
const updateSheet = (sheet, events: GoogleAppsScript.Calendar.CalendarEvent[]) => {
    const eventsProps = events.map(getEventProps);
    const data: string[][] = [
        eventsProps.map(prop => prop.date),
        eventsProps.map(prop => prop.time),
        eventsProps.map(prop => prop.name),
        eventsProps.map(prop => prop.place)
    ];
    // シートに記載された日程データを確認し、カレンダー側に存在しない場合（後から稽古が削除された場合）は列を削除する
    const lastColumn = sheet.getLastColumn();
    for (let i = 3; i <= lastColumn; i++) {
        const date = sheet.getRange(1, i).getValue();
        const time = sheet.getRange(2, i).getValue();
        const index = data[0].findIndex((d, i) => d === date && data[1][i] === time);
        if (index === -1) {
            sheet.deleteColumn(i);
        }
    }
    // シートに記載された日程データを確認し、シート側に存在しない場合（後から稽古が追加された場合）は列を挿入する
    data[0].forEach((date, index) => {
        const time = data[1][index];
        if (sheet.getRange(1, index + 3).getValue() !== date || sheet.getRange(2, index + 3).getValue() !== time) {
            sheet.insertColumnBefore(index + 3);
        }
    });

    sheet.getRange(1, 3, data.length, data[0].length).setValues(data);
}

// 指定した名称のシートが存在するか確認し、存在しなければ追加する関数
const addSheetIfNotExists = (sheetName) => {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet === null) {
        // templateシートをコピーして新しいシートを作成する
        let templateSheet = spreadsheet.getSheetByName("template");
        // templateSheetが存在しなければエラーにする
        if (templateSheet === null) {
            throw new Error("templateシートが存在しません。管理者に連絡してください。");
        }
        sheet = templateSheet.copyTo(spreadsheet).setName(sheetName);
    }
    return sheet;
}

// 指定した年月のシートにカレンダーのイベントを反映する関数
const updateSheetByCalendar = (year: number, month: number) => {
    let sheetName = `${year}${getPaddingNumber(month)}`;
    let events = getEvents(year, month);
    if (events.length == 0) {
        Logger.log('該当月のカレンダーにはイベントが存在しませんでした');
        return '該当月のカレンダーにはイベントが存在しませんでした';
    }
    let sheet = addSheetIfNotExists(sheetName);
    updateSheet(sheet, events);
    return events.map(getEventProps).map(prop => `${prop.date} ${prop.time}<br/> ${prop.name} @${prop.place}<br/>`).join("<br/>");
}

// 文言を引数に取り、通知先アドレスに稽古日程メールを送る関数
const sendAnnouncementEmail = (month: number, text: string, optionalText?: string) => {
    if (notificationAddress === null) {
        Logger.log('通知先アドレスが正しく設定されていません');
        return '通知先アドレスが正しく設定されていません';
    } else if (notificationCcAddress === null) {
        Logger.log('通知先CCアドレスが正しく設定されていません');
        return '通知先CCアドレスが正しく設定されていません'
    }
    const messageBody = `
    みなさま<br/>
    <br/>
    お世話になっております。${month}月の稽古日程をお知らせします。<br/>
    <br/>
    ${text}<br/>
    <br/>
    ${optionalText ? "[担当者からの補足]<br/>" + optionalText : ""}<br/>
    ------<br/>
    以下のリンクから出欠登録をお願いします。<br/>
    https://script.google.com/macros/s/${deployId}/exec?sheetName=${getNow().year}${getPaddingNumber(month)} <br/>
    <br/>`;

    MailApp.sendEmail({
        name: "創玄会 出欠管理システム",
        to: notificationAddress,
        cc: notificationCcAddress,
        subject: `創玄会 ${month}月 稽古日程のお知らせ`,
        htmlBody: messageBody
    });
}

// 日次で起動し、翌日に稽古がある場合はメールを送信する関数
const sendAttendanceEmailIfNecessary = () => {

    // 翌日の日付を取得する
    let tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);

    // 翌日の日付が存在し得るシートを取得する
    const sheetName = `${tomorrow.getFullYear()}${getPaddingNumber(tomorrow.getMonth() + 1)}`;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    const columnIndex = getColumnIndexFromSheet(sheet, tomorrow);
    // 翌日に稽古がない場合は何もしない
    if (sheet === null || columnIndex.length === 0) {
        return;
    }

    const textList = columnIndex.map((index) => {
        // 稽古の概要情報を取得する
        const date = sheet.getRange(1, index + 1).getValue();
        const time = sheet.getRange(2, index + 1).getValue();
        const name = sheet.getRange(3, index + 1).getValue();
        const place = sheet.getRange(4, index + 1).getValue();

        const data = sheet.getDataRange().getValues();
        const getName = (member: string[]) => member[0];
        // 鍵開け担当者を取得する
        const keyHolder = data.filter((member) => member[index] === statusDefinition.key).map(getName);
        // 参加者を取得する
        const joiners = data.filter((member) => member[index] === statusDefinition.join).map(getName);
        // 遅刻者を取得する
        const lateComers = data.filter((member) => member[index] === statusDefinition.late).map(getName);
        // 早退者を取得する
        const earlyLeavers = data.filter((member) => member[index] === statusDefinition.early).map(getName);
        // TODO: 備品の保持者を取得して連絡・更新日時も保持する

        // 文言を統合して作成する
        const text = `${date} ${time}<br/>${name} @${place}<br/>
        [鍵開け担当者]: ${keyHolder.length > 0 ? keyHolder.join(", ") : "※登録なし・要確認"}<br/>
        [参加者]: ${joiners.join(", ")}<br/>
        ${lateComers.length > 0 ? "[遅刻]:" + lateComers.join(", ") + "<br/>" : ""}
        ${earlyLeavers.length > 0 ? "[早退]:" + earlyLeavers.join(", ") + "<br/>" : ""}`;
        return text;
    })
    
    // メールを送信する
    if (notificationAddress === null) {
        Logger.log('通知先アドレスが正しく設定されていません');
        return '通知先アドレスが正しく設定されていません';
    }
    const messageBody = `
    みなさま<br/>
    <br/>
    お疲れさまです。出欠管理より明日の稽古場所および参加者をお知らせします。<br/>
    <br/>
    ${textList.join("<br/>")}<br/>
    <br/>
    ※このメールは自動送信です。<br/>
    `;

    MailApp.sendEmail({
        name: "創玄会 出欠管理システム",
        to: notificationAddress,
        subject: `創玄会 ${(tomorrow.getMonth() + 1)}/${(tomorrow.getDate())}の稽古参加者`,
        htmlBody: messageBody
    });
}

// 指定されたシートから、指定された日付の列番号を返す関数
const getColumnIndexFromSheet = (sheet: GoogleAppsScript.Spreadsheet.Sheet | null, date: Date) => {
    if (sheet === null) {
        return [];
    }
    const dateRow = sheet.getDataRange().getValues()[0];
    const dateText = getDateString(date);
    const columnIndex = dateRow.flatMap((date, i) => date === dateText ? i : []);
    return columnIndex;
}
