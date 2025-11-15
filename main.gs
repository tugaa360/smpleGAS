/***********************************************************
 * ポイントサイト横断｜クレカ案件トラッキング自動化 GAS
 * main.gs
 *
 * 含まれる機能：
 * - ランキング更新（API利用、エラー通知、排他制御）
 * - 契約管理シートのテンプレ生成
 * - 契約記録登録（安全な日付計算）
 * - 解約推奨日のカレンダー登録（重複防止）
 * - メール通知（重複防止、冗長化）
 * - トリガー自動設定（ランキング更新とメール通知を毎朝実行）
 * - エラー通知の実装
(C)tugaa All rights reserved. 
 ***********************************************************/

const API_URL = "https://point-site-research.com/api/creditcard_ranking.json";
const ADMIN_EMAIL = Session.getActiveUser().getEmail(); // 管理者メール

/***********************************************
 * 1. ランキング更新：外部APIから30件取得
 ***********************************************/
function mainUpdate() {
  const lock = LockService.getScriptLock();
  // 30秒待機し、ロックが取得できなければ終了
  if (!lock.tryLock(30000)) {
    console.warn("別のプロセスが実行中です");
    return;
  }

  try {
    const response = UrlFetchApp.fetch(API_URL, { muteHttpExceptions: true });
    
    // HTTPステータスコードの検証
    if (response.getResponseCode() !== 200) {
      throw new Error(`API Error: ${response.getResponseCode()}`);
    }

    const json = response.getContentText();
    const data = JSON.parse(json);

    // データ検証
    if (!Array.isArray(data) || data.length === 0) {
      throw new Error("APIレスポンスが空または不正な形式です");
    }

    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName("ランキング");
    if (!sheet) sheet = ss.insertSheet("ランキング");

    sheet.clear();
    sheet.appendRow([
      "順位", "カード名", "最高報酬", "次点",
      "終了予定", "条件", "指標", "URL"
    ]);

    data.slice(0, 30).forEach((c, i) => {
      // データ存在チェック
      if (!c || typeof c !== 'object') return;

      let mark = "";
      if (c.is_new) mark += "【NEW】";
      if (c.is_hot) mark += "【急げ！】";

      sheet.appendRow([
        i + 1,
        c.card_name || "不明",
        c.best_site && c.best_point 
          ? `${c.best_site} ${c.best_point.toLocaleString()}円`
          : "情報なし",
        c.second_site && c.second_point
          ? `${c.second_site} ${c.second_point.toLocaleString()}円`
          : "なし",
        c.end_date || "不明",
        c.conditions || "情報なし",
        mark || "-",
        c.best_url || ""
      ]);
    });

    sheet.setFrozenRows(1);
    console.log(`ランキング更新成功: ${data.length}件`);

  } catch (e) {
    console.error("mainUpdate Error:", e);
    // エラー通知
    sendErrorNotification("ランキング更新エラー", e.toString());
  } finally {
    lock.releaseLock();
  }
}


/***********************************************
 * 2. 契約管理シートのテンプレート自動生成
 ***********************************************/
function createContractManagementSheet() {
  const ss = SpreadsheetApp.getActive();
  const name = "契約管理";
  let sheet = ss.getSheetByName(name);

  if (sheet) {
    console.log("契約管理シートは既に存在します");
    return;
  }

  sheet = ss.insertSheet(name);

  const header = [
    "カード名",
    "契約日",
    "解約推奨日",
    "獲得ポイント",
    "最高報酬サイト",
    "URL",
    "備考",
    "ポイント付与条件",
    "解約推奨日の算出根拠",
    "通知済み" // メール送信管理用フラグ
  ];

  sheet.appendRow(header);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, header.length, 150);

  const comments = {
    "カード名": "案件名を入力",
    "契約日": "申し込み完了日",
    "解約推奨日": "自動計算 or 手動入力",
    "獲得ポイント": "獲得後更新",
    "最高報酬サイト": "ハピタス、モッピー等",
    "URL": "申し込みURL",
    "備考": "キャンペーン情報など",
    "ポイント付与条件": "例：●円利用など",
    "解約推奨日の算出根拠": "例：契約から6ヶ月後の月末",
    "通知済み": "システム用（触らないでください）"
  };

  header.forEach((title, i) => {
    if (comments[title]) {
      sheet.getRange(1, i + 1).setNote(comments[title]);
    }
  });

  sheet.getRange(1, 1, 1, header.length)
    .setBackground("#efefef")
    .setFontWeight("bold");
    
  console.log("契約管理シート作成完了");
}


/***********************************************
 * 3. 契約記録の登録（手動 or UI で利用）
 ***********************************************/
function recordContract({ cardName, points, site, url, memo, conditions, basis }) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("契約管理");
  if (!sheet) {
    throw new Error("契約管理シートがありません。initialSetup() を実行してください。");
  }

  if (!cardName) {
    throw new Error("カード名は必須です");
  }

  const today = new Date();

  // 解約推奨日の計算（6ヶ月後、月末考慮）
  const cancelDefault = addMonthsSafely(today, 6);

  sheet.appendRow([
    cardName,
    today,
    cancelDefault,
    points || "",
    site || "",
    url || "",
    memo || "",
    conditions || "",
    basis || "契約から6ヶ月後（デフォルト）",
    "" // 通知済みフラグ（空白）
  ]);
  
  console.log(`契約記録追加: ${cardName}`);
}


/***********************************************
 * 4. 解約推奨日を Google カレンダーへ登録
 ***********************************************/
function registerCancelDateToCalendar() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("契約管理");
  if (!sheet) {
    console.warn("契約管理シートが見つかりません");
    return;
  }

  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const idxCard = header.indexOf("カード名");
  const idxCancel = header.indexOf("解約推奨日");

  if (idxCard === -1 || idxCancel === -1) {
    throw new Error("必須カラムが見つかりません（カード名, 解約推奨日）");
  }

  const cal = CalendarApp.getDefaultCalendar();
  let addedCount = 0;

  values.forEach(r => {
    const card = r[idxCard];
    const cancel = r[idxCancel];

    if (!card || !(cancel instanceof Date)) return;

    // 重複チェック：同名・同日のイベントが既存か確認
    const existingEvents = cal.getEventsForDay(cancel);
    const eventTitle = `【解約推奨】${card}`;
    
    const isDuplicate = existingEvents.some(e => e.getTitle() === eventTitle);

    if (isDuplicate) {
      console.log(`既存イベントをスキップ: ${card}`);
      return;
    }

    cal.createAllDayEvent(
      eventTitle,
      cancel,
      {
        description: `${card} の解約推奨日です。\n\nこのリマインダーは自動生成されました。`,
        reminders: [
          { method: "popup", minutes: 7 * 24 * 60 }, // 1週間前
          { method: "popup", minutes: 1 * 24 * 60 }  // 前日
        ]
      }
    );
    
    addedCount++;
  });
  
  console.log(`カレンダー登録完了: ${addedCount}件`);
}


/***********************************************
 * 5. メール通知（重複防止版）
 ***********************************************/
function sendCancelReminderEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("契約管理");
  if (!sheet) {
    console.warn("契約管理シートが見つかりません");
    return;
  }

  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const idxCard = header.indexOf("カード名");
  const idxCancel = header.indexOf("解約推奨日");
  const idxUrl = header.indexOf("URL");
  const idxMemo = header.indexOf("備考");
  const idxNotified = header.indexOf("通知済み");

  if (idxCard === -1 || idxCancel === -1 || idxNotified === -1) {
    console.error("必須カラムが見つかりません（カード名, 解約推奨日, 通知済み）");
    return;
  }

  const today = new Date();
  const week = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // 7日後

  const todayStr = formatDateKey(today);
  const weekStr = formatDateKey(week);

  const user = ADMIN_EMAIL;
  let sentCount = 0;

  values.forEach((row, index) => {
    const card = row[idxCard];
    const cancel = row[idxCancel];
    const url = row[idxUrl] || "なし";
    const memo = row[idxMemo] || "なし";
    const notified = row[idxNotified] || "";

    if (!card || !(cancel instanceof Date)) return;

    const cancelStr = formatDateKey(cancel);

    // 今日または1週間後の場合、かつ当該通知日で未通知の場合
    if ((cancelStr === todayStr || cancelStr === weekStr) && !notified.includes(cancelStr)) {
      try {
        GmailApp.sendEmail(
          user,
          `【解約推奨】${card} の解約日が迫っています`,
          `以下カードの解約推奨日が近づいています。\n\n` +
          `カード名: ${card}\n` +
          `解約推奨日: ${Utilities.formatDate(cancel, "JST", "yyyy/MM/dd")}\n` +
          `URL: ${url}\n` +
          `備考: ${memo}\n\n` +
          `※このメールは自動送信されています`
        );

        // 通知済みフラグを更新
        const rowNum = index + 2; // ヘッダー分+1
        const newNotified = notified ? `${notified},${cancelStr}` : cancelStr;
        sheet.getRange(rowNum, idxNotified + 1).setValue(newNotified);
        
        sentCount++;
        console.log(`メール送信完了: ${card}`);

      } catch (e) {
        console.error(`メール送信エラー (${card}):`, e);
      }
    }
  });

  if (sentCount > 0) {
    console.log(`メール送信合計: ${sentCount}件`);
  }
}


/***********************************************
 * 6. 毎朝のトリガー自動設定
 ***********************************************/
function setupDailyTriggers() {
  const functionsToTrigger = ["sendCancelReminderEmails", "mainUpdate"];

  // 既存トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (functionsToTrigger.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 新規トリガー作成 (毎朝8時に実行)
  functionsToTrigger.forEach(funcName => {
    ScriptApp.newTrigger(funcName)
      .timeBased()
      .atHour(8)
      .everyDays(1)
      .create();
    console.log(`トリガー設定完了: ${funcName} を毎朝8時に実行`);
  });
}


/***********************************************
 * 7. 初期セットアップ（初回実行用）
 ***********************************************/
function initialSetup() {
  console.log("=== 初期セットアップ開始 ===");
  
  try {
    // 1. 管理シートの作成
    createContractManagementSheet();
    
    // 2. 毎朝実行トリガーの設定 (ランキング更新 & メール通知)
    setupDailyTriggers();
    
    // 3. ランキングの初回取得
    mainUpdate();
    
    console.log("=== セットアップ完了 ===");
    console.log("次のステップ:");
    console.log("1. カードを契約したら、契約管理シートに手動または recordContract() で記録");
    console.log("2. registerCancelDateToCalendar() を実行してカレンダーに登録");
    
  } catch (e) {
    console.error("初期セットアップ中に致命的なエラー:", e);
    sendErrorNotification("初期セットアップ失敗", e.toString());
  }
}


/***********************************************
 * ユーティリティ関数
 ***********************************************/

/**
 * 月末を考慮した月加算 (日付ズレ防止)
 * 例：1/31 + 1ヶ月 = 2/28（うるう年は2/29）
 */
function addMonthsSafely(date, months) {
  const result = new Date(date);
  const targetDay = result.getDate();
  const targetMonth = result.getMonth() + months;
  
  result.setMonth(targetMonth);
  
  // 元の日付と加算後の日付が違う場合 (例: 1/31を2/31に設定しようとして3/3になる)
  if (result.getDate() !== targetDay) {
    result.setDate(0); // その月の前月末日に設定
  }
  
  return result;
}

/**
 * 日付を比較用キー（YYYY-MM-DD）に変換
 */
function formatDateKey(date) {
  return Utilities.formatDate(date, "JST", "yyyy-MM-dd");
}

/**
 * エラー通知メール送信
 */
function sendErrorNotification(subject, errorMessage) {
  try {
    GmailApp.sendEmail(
      ADMIN_EMAIL,
      `[GASエラー] ${subject}`,
      `以下のエラーが発生しました:\n\n${errorMessage}\n\n実行時刻: ${new Date()}`
    );
  } catch (e) {
    console.error("エラー通知の送信に失敗:", e);
  }
}
