/* exported initialTrigger, resetTriggers, createOfficeHourCheckTrigger, officeHourCheck */

// 実行するGoogleアカウントのカレンダーに登録されている「日本の祝日」カレンダーのIDを参照する↓
const CAL_ID_HOLIDAY_JA = 'en.japanese#holiday@group.v.calendar.google.com';
// オフィスアワーの開始・終了時刻をHHmmss形式のstringで設定↓
const OFFICE_HOURS = { start: '100000', end: '130000' };
// エラーとなったときの再実行までの時間（ミリ秒）↓
const RETRY_MILLISEC = 5 * 60 * 1000;
// officeHourCheckがエラーとなったときの再実行までの時間（ミリ秒）↓
const RETRY_MILLISEC_OFFICE_HOUR_CHECK = 5 * 1000;
const TIMEZONE = Session.getScriptTimeZone();
const UP_KEY_DATE = 'upDate';
const UP_KEY_OFFICE_OPEN_STATUS = 'officeOpenStatus';
const UP_KEY_SCRIPT_STATUS = 'scriptStatus';

/**
 * 最初に手動で実行する。
 */
function initialTrigger() {
  console.info('[initialTrigger] Initiating...'); // log
  var now = new Date();
  var upDate = Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd');
  var nowTime = Utilities.formatDate(now, TIMEZONE, 'HHmmss');
  var officeOpenStatus = 'CLOSED';
  var outOfOfficeHours = isWeekendOrHolidayJa(now);
  console.info(`[initialTrigger] outOfOfficeHours: ${outOfOfficeHours}\nnowTime: ${nowTime}\nOFFICE_HOURS.start: ${OFFICE_HOURS.start}\nOFFICE_HOURS.end: ${OFFICE_HOURS.end}`); // log
  if (!outOfOfficeHours && nowTime >= OFFICE_HOURS.start && nowTime <= OFFICE_HOURS.end) {
    console.info(`[initialTrigger] Switching officeOpenStatus to "OPEN"\n\outOfOfficeHours: ${outOfOfficeHours}\nnowTime: ${nowTime}`); // log
    officeOpenStatus = 'OPEN';
  }
  var up = PropertiesService.getUserProperties()
    .setProperty(UP_KEY_DATE, upDate)
    .setProperty(UP_KEY_OFFICE_OPEN_STATUS, officeOpenStatus)
    .setProperty(UP_KEY_SCRIPT_STATUS, 'ERROR');
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'resetTriggers') {
      ScriptApp.deleteTrigger(trigger);
      console.info(`[initialTrigger] Deleted trigger for resetTriggers`); // log
    }
  });
  resetTriggers();
  ScriptApp.newTrigger('resetTriggers')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
  console.info(`[initialTrigger] Complete. Current user properties:\n\n${JSON.stringify(up.getProperties())}`); // log
}

/**
 * 日付が変わったり、scriptStatusがERRORとなっている場合に、
 * 一連の時刻判定トリガーを初期化して、再設定する。
 * initialTriggerによって、1日1回、午前0～1時で実行するようにトリガー設定される。
 */
function resetTriggers() {
  console.info('[resetTriggers] Initiating...'); // log
  var scriptStatus = 'ERROR';
  var today = new Date();
  var up = PropertiesService.getUserProperties();
  try {
    let todayString = Utilities.formatDate(today, TIMEZONE, 'yyyyMMdd');
    // 日付が変わっている、または前回実行時にエラーが発生していた場合に行う、トリガーの初期化処理
    if (todayString !== up.getProperty(UP_KEY_DATE) || up.getProperty(UP_KEY_SCRIPT_STATUS) !== 'RUNNING') {
      console.info(`[resetTriggers] User Property: ${up.getProperty(UP_KEY_DATE)} -> todayString: ${todayString}; scriptStatus: ${scriptStatus}`); // log
      up.setProperty(UP_KEY_DATE, todayString);
      // この関数（resetTriggers）のトリガー以外の全てのトリガーを削除
      ScriptApp.getProjectTriggers().forEach(trigger => {
        let handler = trigger.getHandlerFunction();
        if (handler !== 'resetTriggers') {
          ScriptApp.deleteTrigger(trigger);
          console.info(`[resetTriggers] Deleted trigger for ${handler}`); // log
        }
      });
      // 平日判定。
      // 土日祝日の場合は、次の日まで本トリガー以外のトリガーを発動させない。
      if (!isWeekendOrHolidayJa(today)) {
        console.info(`Today is a weekday.`); //log
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
        console.info('[resetTriggers] Set trigger officeHourCheck'); //log
      }
    }
    scriptStatus = 'RUNNING';
    console.info(`[resetTriggers] ${UP_KEY_DATE}: ${todayString}; scriptStatus: ${scriptStatus}`); // log
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] triggerReset', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('resetTriggers')
      .timeBased()
      .after(RETRY_MILLISEC)
      .create();
    console.info(`[resetTriggers] ${error.message}`); //log
  } finally {
    up.setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.info(`[resetTriggers] scriptStatus: ${scriptStatus}`); //log
  }
}

/**
 * オフィスアワー開始・終了時刻の前の1時間内（例：オフィスアワーの開始時刻が8:30であれば、7時台）に実行され、
 * 1分おきのトリガーを設定する。
 */
function createOfficeHourCheckTrigger() {
  console.info('[createOfficeHourCheckTrigger] Initiating...'); // log
  var scriptStatus = 'ERROR';
  try {
    ScriptApp.newTrigger('officeHourCheck')
      .timeBased()
      .everyMinutes(1)
      .create();
    scriptStatus = 'RUNNING';
    console.info('[createOfficeHourCheckTrigger] Set trigger officeHourCheck'); //log
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] createOfficeHourCheckTrigger', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('createOfficeHourCheckTrigger')
      .timeBased()
      .after(RETRY_MILLISEC)
      .create();
    console.info(`[createOfficeHourCheckTrigger] ${error.message}`); //log
  } finally {
    PropertiesService.getUserProperties().setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.info(`[createOfficeHourCheckTrigger] scriptStatus: ${scriptStatus}`); //log
  }
}

/**
 * オフィスアワーを判定する。必要に応じてオフィスアワーの開始・終了時にそれぞれ実行させたい処理を挿入する。
 */
function officeHourCheck() {
  console.info('[officeHourCheck] Initiating...'); // log
  var scriptStatus = 'ERROR';
  var up = PropertiesService.getUserProperties();
  var officeOpenStatus = up.getProperty(UP_KEY_OFFICE_OPEN_STATUS);
  var now = Utilities.formatDate(new Date(), TIMEZONE, 'HHmmss');
  try {
    if (!officeOpenStatus || officeOpenStatus === 'CLOSED') {
      // 現在オフィスが閉まっている場合
      console.info(`[officeHourCheck] Office is closed.\n\nofficeOpenStatus: ${officeOpenStatus}`); //log
      if (now >= OFFICE_HOURS.start) {
        officeOpenStatus = 'OPEN';
        console.info(`[officeHourCheck] Office is now open.\n\nnow: ${now}\nOFFICE_HOURS start: ${OFFICE_HOURS.start}, end: ${OFFICE_HOURS.end}\nofficeOpenStatus: ${officeOpenStatus}`); // log
        ///// ↓↓↓ここにオフィスアワー開始時に実行させたい処理を入れる↓↓↓ /////
        functionToExecuteWhenOpen();
        ///// ↑↑↑ここにオフィスアワー開始時に実行させたい処理を入れる↑↑↑ /////
        ScriptApp.getProjectTriggers().forEach(trigger => {
          if (trigger.getHandlerFunction() === 'officeHourCheck') {
            ScriptApp.deleteTrigger(trigger);
            console.log('[officeHourCheck] Trigger for officeHourCheck is deleted.'); // log
          }
        });
      }
    } else if (officeOpenStatus === 'OPEN') {
      // 現在オフィスが開いている場合
        console.info(`[officeHourCheck] Office is open.\n\nofficeOpenStatus: ${officeOpenStatus}`); // log
      if (now > OFFICE_HOURS.end) {
        officeOpenStatus = 'CLOSED';
        console.info(`[officeHourCheck] Office is now closed.\n\nnow: ${now}\nOFFICE_HOURS start: ${OFFICE_HOURS.start}, end: ${OFFICE_HOURS.end}\nofficeOpenStatus: ${officeOpenStatus}`); // log
        ///// ↓↓↓ここにオフィスアワー終了時に実行させたい処理を入れる↓↓↓ /////
        functionToExecuteWhenClosed();
        ///// ↑↑↑ここにオフィスアワー終了時に実行させたい処理を入れる↑↑↑ /////
        ScriptApp.getProjectTriggers().forEach(trigger => {
          if (trigger.getHandlerFunction() === 'officeHourCheck') {
            ScriptApp.deleteTrigger(trigger);
            console.info('[officeHourCheck] Trigger for officeHourCheck is deleted.'); // log
          }
        });
      }
    } else {
      throw new Error('Unknown officeOpenStatus value.');
    }
    up.setProperty(UP_KEY_OFFICE_OPEN_STATUS, officeOpenStatus);
    scriptStatus = 'RUNNING';
    console.info(`[officeHourCheck] Final officeOpenStatus: ${officeOpenStatus}`); // log
  } catch (error) {
    // メールでエラー通知
    let myEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(myEmail, '[Error] officeHourCheck', error.message); // 使用者に応じてわかりやすいメッセージにしておく。
    // 再度実行するためのトリガー設定
    ScriptApp.newTrigger('officeHourCheck')
      .timeBased()
      .after(RETRY_MILLISEC_OFFICE_HOUR_CHECK)
      .create();
    console.log(`[officeHourCheck] ${error.message}`); // log
  } finally {
    up.setProperty(UP_KEY_SCRIPT_STATUS, scriptStatus);
    console.log(`[officeHourCheck] scriptStatus: ${scriptStatus}`); // log
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
  console.info(`[parseOfficeHourTrigger] Initiating with officeHourString: ${officeHourString}...`); // log
  var officeHourNum = parseInt(officeHourString);
  if (officeHourNum < 0 || officeHourNum > 235959) {
    throw new Error('OFFICE_HOURS.start and OFFICE_HOURS.end must be between 000000 and 235959');
  }
  var officeHour = Math.trunc(officeHourNum / 10000);
  var triggerTime = (officeHour === 0 ? 23 : officeHour - 1);
  console.info(`[parseOfficeHourTrigger] Returning triggerTime: ${triggerTime}`); // log
  return triggerTime;
}

/**
 * オフィスアワー時に実行させたい処理
 */
function functionToExecuteWhenOpen() {
  console.log('[functionToExecuteWhenOpen] Office is now open.'); // log
}

/**
 * オフィスアワー外で実行させたい処理
 */
function functionToExecuteWhenClosed() {
  console.log('[functionToExecuteWhenClosed] Office is now closed.'); // log
}
