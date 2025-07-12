/**
 * @fileoverview Gmailで受信したホットペッパービューティーの予約確定メールを解析し、
 * Googleカレンダーに予定を自動で登録するGoogle Apps Scriptです。
 */

// --- 設定項目 ---

/**
 * 予定のリマインダー（通知）を何分前に設定するかを配列で定義します。
 * 例: [120, 30] と設定すると、予定の2時間前と30分前に通知が設定されます。
 * @type {number[]}
 */
const REMINDER_MINUTES_BEFORE = [120, 30];


// --- メイン処理 ---

/**
 * スクリプトのメイン関数。
 * 未読の予約確定メールを検索し、カレンダーへの登録処理を呼び出します。
 * この関数をトリガーで定期実行してください。
 */
function createEventFromHotPepperMail() {
  // 検索条件: 指定の差出人、件名、かつ未読のメール
  const query = 'from:reserve@beauty.hotpepper.jp subject:ご予約が確定いたしました is:unread';

  try {
    const threads = GmailApp.search(query);
    Logger.log(`${threads.length}件の対象スレッドが見つかりました。`);

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        // 念のため、再度未読であることを確認
        if (message.isUnread()) {
          processSingleMail(message);
        }
      }
    }
  } catch (e) {
    Logger.log(`スクリプト実行中にエラーが発生しました: ${e.toString()}`);
  }
}

/**
 * 1通のメールを処理し、カレンダー登録と既読化を行います。
 * @param {GoogleAppsScript.Gmail.GmailMessage} message 処理対象のメールオブジェクト
 */
function processSingleMail(message) {
  const mailBody = message.getPlainBody(); // HTMLではなくプレーンテキストの本文を取得
  Logger.log('未読メールを発見。本文の解析を開始します。');

  try {
    // メール本文から予約情報を抽出
    const visitDateTime = extractVisitDateTime(mailBody);
    const salonName = extractSalonName(mailBody);

    // 日時かサロン名のどちらかが抽出できなければ処理を中断
    if (!visitDateTime || !salonName) {
      Logger.log('来店日時またはサロン名が本文から抽出できませんでした。このメールの処理をスキップします。');
      // 予期せぬフォーマットの可能性があるため、無限ループを避けるために既読にする
      message.markRead();
      return;
    }
    
    Logger.log(`抽出成功: サロン名「${salonName}」, 来店日時「${visitDateTime}」`);

    // Googleカレンダーに予定を作成
    createCalendarEvent(salonName, visitDateTime);

    // 重複処理を防ぐため、処理が完了したメールを既読にする
    message.markRead();
    Logger.log('カレンダーへの登録が完了したため、メールを既読にしました。');

  } catch (e) {
    Logger.log(`メール処理中にエラーが発生しました: ${e.toString()}\nエラーが発生したメールの件名: ${message.getSubject()}`);
    // エラー発生時はメールを未読のままにし、次回の実行で再試行できるようにします。
  }
}


// --- ヘルパー関数 ---

/**
 * メール本文から「来店日時」を正規表現で抽出します。
 * @param {string} body メールの本文
 * @return {Date|null} 抽出・変換済みのDateオブジェクト、または見つからない場合はnull
 */
function extractVisitDateTime(body) {
  // "■来店日時"と日時の間にある空白(全角/半角)や改行(\s*)に柔軟に対応する正規表現
  const regex = /■来店日時\s*(\d{4}年\d{1,2}月\d{1,2}日（.）\d{1,2}:\d{2})/;
  const match = body.match(regex);

  if (!match || !match[1]) {
    return null;
  }

  // "2025年07月11日（金）14:00" のような形式の文字列を取得
  let dateTimeString = match[1];

  // new Date()で正しく解釈できるよう、日本語と曜日を整形します
  // 1. "（金）" のような曜日の部分を削除 -> "2025年07月11日 14:00"
  dateTimeString = dateTimeString.replace(/（.）/, '');
  // 2. "年", "月"を"/"に、"日"を空白に置換 -> "2025/07/11 14:00"
  dateTimeString = dateTimeString.replace('年', '/').replace('月', '/').replace('日', ' ');
  
  return new Date(dateTimeString);
}

/**
 * メール本文から「サロン名」を正規表現で抽出します。
 * @param {string} body メールの本文
 * @return {string|null} 抽出したサロン名、または見つからない場合はnull
 */
function extractSalonName(body) {
  // "■サロン名"と実際の名前の間にある空白(全角/半角)や改行(\s*)に対応する正規表現
  const regex = /■サロン名\s*(.+)/;
  const match = body.match(regex);
  
  // マッチし、かつキャプチャした文字列が存在する場合
  if (match && match[1]) {
    return match[1].trim(); // 前後の不要な空白を削除して返す
  }
  
  return null;
}

/**
 * 抽出した情報を用いてGoogleカレンダーに予定を作成します。
 * @param {string} title 予定のタイトル（サロン名）
 * @param {Date} startTime 予定の開始日時
 */
function createCalendarEvent(title, startTime) {
  const calendar = CalendarApp.getDefaultCalendar();
  
  // 終了日時を開始日時の1時間後に設定
  const endTime = new Date(startTime.getTime() + 60 * 60 * 1000);

  const options = {
    location: title, // 場所にもサロン名を設定
    description: 'ホットペッパービューティーの予約確定メールから自動登録されました。'
  };

  const event = calendar.createEvent(title, startTime, endTime, options);
  
  // グローバル変数で設定されたリマインダーをループで追加
  for (const minutes of REMINDER_MINUTES_BEFORE) {
    event.addPopupReminder(minutes);
  }

  Logger.log(`カレンダーに予定「${title}」を登録しました。（${startTime} - ${endTime}）`);
}
