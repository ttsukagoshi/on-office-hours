/* exported resetTriggers, createOfficeHourCheckTrigger, officeHourCheck */

// 実行するGoogleアカウントのカレンダーに登録されている「日本の祝日」カレンダーのIDを参照する↓
const CAL_ID_HOLIDAY_JA = 'en.japanese#holiday@group.v.calendar.google.com';
// オフィスアワーの開始・終了時刻をHHmmss形式のstringで設定↓
const OFFICE_HOURS = { start: '090000', end: '170000' }; 
// エラーとなったときの再実行までの時間（ミリ秒）↓
const RETRY_MILLISEC = 5 * 60 * 1000;
// officeHourCheckがエラーとなったときの再実行までの時間（ミリ秒）↓
const RETRY_MILLISEC_OFFICE_HOUR_CHECK = 5 * 1000;
const TIMEZONE = Session.getScriptTimeZone();
const UP_KEY_DATE = 'upDate';
const UP_KEY_OFFICE_OPEN_STATUS = 'officeOpenStatus';
const UP_KEY_SCRIPT_STATUS = 'scriptStatus';

/**
 * 一連の時刻判定トリガーをリセットする。
 * 1日1回、午前0～1時で実行するようにトリガー設定。
 */
function resetTriggers() {
  var scriptStatus = 'ERROR';
  var today = new Date();
  var up = PropertiesService.getUserProperties();
  try {
    let todayString = Utilities.formatDate(today, TIMEZONE, 'yyyyMMdd');
    // 日付が変わっている、または前回実行時にエラーが発生していた場合に行う、トリガーの初期化処理
    if (todayString !== up.getProperty(UP_KEY_DATE) || up.getProperty(UP_KEY_SCRIPT_STATUS) !== 'RUNNING') {
      console.log(`User Property: ${up.getProperty(UP_KEY_DATE)} -> todayString: ${todayString}; scriptStatus: ${scriptStatus}`);
      up.setProperty(UP_KEY_DATE, todayString);
      // この関数（resetTriggers）のトリガー以外の全てのトリガーを削除
      ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getHandlerFunction() !== 'resetTriggers') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      // 平日判定。
      // 土日祝日の場合は、次の日まで本トリガー以外のトリガーを発動させない。
      if (!isWeekendOrHolidayJa(today)) {
        console.log(`Today is a weekday.`); //log
        ScriptApp.newTrigger('createOfficeHourCheckTrigger')
          .timeBased()
          .atHour(parseOfficeHourTrigger(OFFICE_HOURS.start))
          .nearMinute(40) // https://developers.google.com/apps-script/reference/script/clock-trigger-builder#nearminuteminute
          .everyDays(1)
          .create();
        ScriptApp.newTrigger('createOfficeHourCheckTrigger')
          .timeBased()
          .atHour(parseOfficeHourTrigger(OFFICE_HOURS.end))
          .nearMinute(40)
          .everyDays(1)
          .create();
        console.log('Set trigger officeHourCheck'); //log
      }
    }
    scriptStatus = 'RUNNING';
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] triggerReset', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('resetTriggers')
      .timeBased()
      .after(RETRY_MILLISEC)
      .create();
    console.log(error.message); //log
  } finally {
    up.setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.log(`scriptStatus: ${scriptStatus}`); //log
  }
}

/**
 * オフィスアワー開始・終了時刻の前の1時間内（例：オフィスアワーの開始時刻が8:30であれば、7時台）に実行され、
 * 1分おきのトリガーを設定する。
 */
function createOfficeHourCheckTrigger() {
  var scriptStatus = 'ERROR';
  try {
    ScriptApp.newTrigger('officeHourCheck')
      .timeBased()
      .everyMinutes(1)
      .create();
    scriptStatus = 'RUNNING';
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] createOfficeHourCheckTrigger', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('createOfficeHourCheckTrigger')
      .timeBased()
      .after(RETRY_MILLISEC)
      .create();
    console.log(error.message); //log
  } finally {
    PropertiesService.getUserProperties().setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.log(`scriptStatus: ${scriptStatus}`); //log
  }
}

/**
 * オフィスアワーを判定する。必要に応じてオフィスアワーの開始・終了時にそれぞれ実行させたい処理を挿入する。
 */
function officeHourCheck() {
  var scriptStatus = 'ERROR';
  var up = PropertiesService.getUserProperties();
  var officeOpenStatus = up.getProperty(UP_KEY_OFFICE_OPEN_STATUS);
  var now = Utilities.formatDate(new Date(), TIMEZONE, 'HHmmss');
  try {
    if (!officeOpenStatus || officeOpenStatus === 'CLOSED') {
      // 現在オフィスが閉まっている場合
      console.log('Office is closed.'); //log
      if (now >= OFFICE_HOURS.start) {
        officeOpenStatus = 'OPEN';
        console.log('Office is now open.'); // log
        ///// ↓↓↓ここにオフィスアワー開始時に実行させたい処理を入れる↓↓↓ /////
        functionToExecuteWhenOpen();
        ///// ↑↑↑ここにオフィスアワー開始時に実行させたい処理を入れる↑↑↑ /////
        ScriptApp.getProjectTriggers().forEach(trigger => {
          if (trigger.getHandlerFunction() === 'officeHourCheck') {
            ScriptApp.deleteTrigger(trigger);
            console.log('Trigger for officeHourCheck is deleted.') // log
          }
        });
      }
    } else if (officeOpenStatus === 'OPEN') {
      // 現在オフィスが開いている場合
      console.log('Office is open.'); //log
      if (now > OFFICE_HOURS.end) {
        officeOpenStatus = 'CLOSED';
        console.log('Office is now closed.'); // log
        ///// ↓↓↓ここにオフィスアワー終了時に実行させたい処理を入れる↓↓↓ /////
        functionToExecuteWhenClosed();
        ///// ↑↑↑ここにオフィスアワー終了時に実行させたい処理を入れる↑↑↑ /////
        ScriptApp.getProjectTriggers().forEach(trigger => {
          if (trigger.getHandlerFunction() === 'officeHourCheck') {
            ScriptApp.deleteTrigger(trigger);
            console.log('Trigger for officeHourCheck is deleted.') // log
          }
        });
      }
    } else {
      throw new Error('Unknown officeOpenStatus value.');
    }
    up.setProperty(UP_KEY_OFFICE_OPEN_STATUS, officeOpenStatus);
    scriptStatus = 'RUNNING';
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] officeHourCheck', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('officeHourCheck')
      .timeBased()
      .after(RETRY_MILLISEC_OFFICE_HOUR_CHECK)
      .create();
    console.log(error.message); // log
  } finally {
    up.setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.log(`scriptStatus: ${scriptStatus}`); // log
  }
}

/**
 * 入力したDateオブジェクトが休日（土日または日本の祝日）かどうかを判定する。
 * スクリプトのタイムゾーンがJST (Asia/Tokyo)となっている前提。
 * @param {Date} dateObj
 * @return {boolean}
 */
function isWeekendOrHolidayJa(dateObj) {
  var weekday = dateObj.getDay(); // Assuming that the script's time zone is set to JST (Asia/Tokyo)
  var holidayEventsJa = CalendarApp.getCalendarById(CAL_ID_HOLIDAY_JA).getEventsForDay(dateObj);
  return (weekday === 0 || weekday === 6 || holidayEventsJa.length > 0);
}

/**
 * HHmmss形式となっている時刻について、その前の時間を0～23のnumberで返す。
 * 例）17時30分（173000）に対しては、16（時台）を返す。
 * @param {string} officeHourString 
 */
function parseOfficeHourTrigger(officeHourString) {
  var officeHourNum = parseInt(officeHourString);
  if (officeHourNum < 0 || officeHourNum > 235959) {
    throw new Error('OFFICE_HOURS.start and OFFICE_HOURS.end must be between 000000 and 235959');
  }
  var officeHour = Math.trunc(officeHourNum / 10000);
  return (officeHour === 0 ? 23 : officeHour - 1);
}

/**
 * オフィスアワー時に実行させたい処理
 */
function functionToExecuteWhenOpen() {
  console.log('[functionToExecuteWhenOpen] Office is now open.');
}

/**
 * オフィスアワー外で実行させたい処理
 */
function functionToExecuteWhenClosed() {
  console.log('[functionToExecuteWhenClosed] Office is now closed.');
}
