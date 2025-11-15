/***********************************************************
 * ポイントサイト横断｜クレカ案件トラッキング自動化 GAS
 * main.gs（最終査読版 / 本番運用可能）
 *
 * 修正内容：
 * - addMonthsSafely() の日付計算ロジック修正
 * - トリガー実行時刻の分散（競合防止）
 * - カレンダー検索の効率化
 * - バッチ更新によるシート書き込み最適化
 * - より詳細なエラーハンドリング
 ***********************************************************/

const API_URL = "https://point-site-research.com/api/creditcard_ranking.json";
const ADMIN_EMAIL = Session.getActiveUser().getEmail();

/***********************************************
 * 1. ランキング更新：外部APIから30件取得
 ***********************************************/
function mainUpdate() {
  const lock = LockService.getScriptLock();
  
  if (!lock.tryLock(30000)) {
    console.warn("別のプロセスが実行中のため、処理をスキップしました");
    return;
  }

  try {
    const response = UrlFetchApp.fetch(API_URL, { 
      muteHttpExceptions: true,
      timeoutSeconds: 30 // タイムアウト設定追加
    });
    
    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`API Error: HTTP ${statusCode}`);
    }

    const json = response.getContentText();
    const data = JSON.parse(json);

    // データ検証
    if (!Array.isArray(data)) {
      throw new Error("APIレスポンスが配列ではありません");
    }
    
    if (data.length === 0) {
      console.warn("APIから0件のデータが返されました");
      return;
    }

    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName("ランキング");
    if (!sheet) {
      sheet = ss.insertSheet("ランキング");
    }

    sheet.clear();
    
    // ヘッダー行
    const headerRow = [
      "順位", "カード名", "最高報酬", "次点",
      "終了予定", "条件", "指標", "URL"
    ];
    sheet.appendRow(headerRow);

    // データ行をバッチで準備
    const dataRows = [];
    data.slice(0, 30).forEach((c, i) => {
      if (!c || typeof c !== 'object') return;

      let mark = "";
      if (c.is_new) mark += "【NEW】";
      if (c.is_hot) mark += "【急げ！】";

      dataRows.push([
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

    // バッチ書き込み（パフォーマンス向上）
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headerRow.length);
    
    console.log(`✓ ランキング更新成功: ${dataRows.length}件 (${new Date().toLocaleString('ja-JP')})`);

  } catch (e) {
    const errorMsg = `mainUpdate Error: ${e.message}\nStack: ${e.stack}`;
    console.error(errorMsg);
    sendErrorNotification("ランキング更新エラー", errorMsg);
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
    return sheet;
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
    "通知済み", // メール送信履歴
    "カレンダー登録済み" // カレンダー登録フラグ（新規追加）
  ];

  sheet.appendRow(header);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, header.length, 120);

  const comments = {
    "カード名": "案件名を入力",
    "契約日": "申し込み完了日 (YYYY/MM/DD)",
    "解約推奨日": "自動計算 or 手動入力 (YYYY/MM/DD)",
    "獲得ポイント": "ポイント獲得後に更新",
    "最高報酬サイト": "例: ハピタス、モッピー",
    "URL": "申し込みURL",
    "備考": "キャンペーン情報など自由記入",
    "ポイント付与条件": "例: 5万円利用など",
    "解約推奨日の算出根拠": "例: 契約から6ヶ月後",
    "通知済み": "システム用（編集不要）",
    "カレンダー登録済み": "システム用（編集不要）"
  };

  header.forEach((title, i) => {
    const cell = sheet.getRange(1, i + 1);
    if (comments[title]) {
      cell.setNote(comments[title]);
    }
  });

  sheet.getRange(1, 1, 1, header.length)
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
    
  console.log("✓ 契約管理シート作成完了");
  return sheet;
}


/***********************************************
 * 3. 契約記録の登録
 ***********************************************/
function recordContract({ 
  cardName, 
  points, 
  site, 
  url, 
  memo, 
  conditions, 
  basis,
  monthsUntilCancel = 6 // デフォルト6ヶ月
}) {
  if (!cardName || cardName.trim() === "") {
    throw new Error("カード名は必須です");
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName("契約管理");
  if (!sheet) {
    throw new Error("契約管理シートがありません。initialSetup() を実行してください。");
  }

  try {
    const today = new Date();
    const cancelDate = addMonthsSafely(today, monthsUntilCancel);

    sheet.appendRow([
      cardName.trim(),
      today,
      cancelDate,
      points || "",
      site || "",
      url || "",
      memo || "",
      conditions || "",
      basis || `契約から${monthsUntilCancel}ヶ月後（自動計算）`,
      "", // 通知済みフラグ
      "" // カレンダー登録済みフラグ
    ]);
    
    console.log(`✓ 契約記録追加: ${cardName} (解約推奨: ${formatDateKey(cancelDate)})`);
    return true;

  } catch (e) {
    const errorMsg = `契約記録追加エラー (${cardName}): ${e.message}`;
    console.error(errorMsg);
    sendErrorNotification("契約記録追加失敗", errorMsg);
    throw e;
  }
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
  const idxRegistered = header.indexOf("カレンダー登録済み");

  if (idxCard === -1 || idxCancel === -1) {
    throw new Error("必須カラムが見つかりません（カード名, 解約推奨日）");
  }

  const cal = CalendarApp.getDefaultCalendar();
  let addedCount = 0;
  let skippedCount = 0;

  // バッチ更新用の配列
  const updatesToWrite = [];

  values.forEach((row, index) => {
    const card = row[idxCard];
    const cancel = row[idxCancel];
    const registered = row[idxRegistered];

    if (!card || !(cancel instanceof Date)) {
      skippedCount++;
      return;
    }

    // 既に登録済みフラグがある場合はスキップ
    if (registered === "済") {
      skippedCount++;
      return;
    }

    try {
      const eventTitle = `【解約推奨】${card}`;

      // イベント作成
      cal.createAllDayEvent(
        eventTitle,
        cancel,
        {
          description: `${card} の解約推奨日です。\n\nカードの利用状況を確認し、必要に応じて解約手続きを行ってください。\n\n※このリマインダーは自動生成されました`,
          reminders: [
            { method: "popup", minutes: 7 * 24 * 60 }, // 1週間前
            { method: "popup", minutes: 1 * 24 * 60 }  // 前日
          ]
        }
      );
      
      // フラグ更新を記録
      if (idxRegistered !== -1) {
        updatesToWrite.push({
          row: index + 2, // ヘッダー分+1
          col: idxRegistered + 1,
          value: "済"
        });
      }

      addedCount++;
      console.log(`✓ カレンダー登録: ${card} (${formatDateKey(cancel)})`);

    } catch (e) {
      console.error(`カレンダー登録エラー (${card}): ${e.message}`);
    }
  });

  // バッチでフラグを更新
  updatesToWrite.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });

  console.log(`✓ カレンダー登録完了: ${addedCount}件追加, ${skippedCount}件スキップ`);
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
    console.error("必須カラムが見つかりません");
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時刻をリセット
  
  const week = new Date(today);
  week.setDate(today.getDate() + 7); // 7日後

  const todayStr = formatDateKey(today);
  const weekStr = formatDateKey(week);

  const user = ADMIN_EMAIL;
  let sentCount = 0;
  const updatesToWrite = [];

  values.forEach((row, index) => {
    const card = row[idxCard];
    const cancel = row[idxCancel];
    const url = row[idxUrl] || "なし";
    const memo = row[idxMemo] || "なし";
    const notified = String(row[idxNotified] || "");

    if (!card || !(cancel instanceof Date)) return;

    // 日付を正規化
    const cancelDate = new Date(cancel);
    cancelDate.setHours(0, 0, 0, 0);
    const cancelStr = formatDateKey(cancelDate);

    // 今日または1週間後かつ未通知
    const shouldNotify = (cancelStr === todayStr || cancelStr === weekStr) 
                        && !notified.includes(cancelStr);

    if (shouldNotify) {
      try {
        const daysUntil = Math.ceil((cancelDate - today) / (1000 * 60 * 60 * 24));
        const urgency = daysUntil === 0 ? "【本日】" : `【あと${daysUntil}日】`;

        GmailApp.sendEmail(
          user,
          `${urgency}${card} の解約推奨日が迫っています`,
          `${card} の解約推奨日が近づいています。\n\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `カード名: ${card}\n` +
          `解約推奨日: ${Utilities.formatDate(cancelDate, "JST", "yyyy年MM月dd日")}\n` +
          `残り日数: ${daysUntil}日\n` +
          `━━━━━━━━━━━━━━━━━━\n\n` +
          `URL: ${url}\n` +
          `備考: ${memo}\n\n` +
          `※このメールは自動送信されています`
        );

        // 通知済みフラグを更新
        const rowNum = index + 2;
        const newNotified = notified ? `${notified},${cancelStr}` : cancelStr;
        updatesToWrite.push({
          row: rowNum,
          col: idxNotified + 1,
          value: newNotified
        });
        
        sentCount++;

      } catch (e) {
        console.error(`メール送信エラー (${card}): ${e.message}`);
      }
    }
  });

  // バッチでフラグを更新
  updatesToWrite.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });

  if (sentCount > 0) {
    console.log(`✓ メール送信完了: ${sentCount}件 (${new Date().toLocaleString('ja-JP')})`);
  } else {
    console.log(`メール送信対象なし (${new Date().toLocaleString('ja-JP')})`);
  }
}


/***********************************************
 * 6. 毎朝のトリガー自動設定
 ***********************************************/
function setupDailyTriggers() {
  // 既存トリガーをクリア
  ScriptApp.getProjectTriggers().forEach(t => {
    const funcName = t.getHandlerFunction();
    if (["sendCancelReminderEmails", "mainUpdate"].includes(funcName)) {
      ScriptApp.deleteTrigger(t);
      console.log(`既存トリガー削除: ${funcName}`);
    }
  });

  // メール通知: 毎朝7時（先に実行）
  ScriptApp.newTrigger("sendCancelReminderEmails")
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  console.log("✓ トリガー設定: sendCancelReminderEmails (毎朝7時)");

  // ランキング更新: 毎朝8時（後に実行、競合防止）
  ScriptApp.newTrigger("mainUpdate")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  console.log("✓ トリガー設定: mainUpdate (毎朝8時)");
}


/***********************************************
 * 7. 初期セットアップ（初回実行用）
 ***********************************************/
function initialSetup() {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("  初期セットアップ開始");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
  
  try {
    // 1. 契約管理シート作成
    console.log("\n[1/3] 契約管理シート作成中...");
    createContractManagementSheet();
    
    // 2. トリガー設定
    console.log("\n[2/3] トリガー設定中...");
    setupDailyTriggers();
    
    // 3. ランキング初回取得
    console.log("\n[3/3] ランキング初回取得中...");
    mainUpdate();
    
    console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━");
    console.log("  ✓ セットアップ完了！");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
    console.log("\n【次のステップ】");
    console.log("1. カード契約時: recordContract() を実行");
    console.log("2. カレンダー登録: registerCancelDateToCalendar() を実行");
    console.log("3. 契約管理シートで詳細を確認・編集");
    
  } catch (e) {
    const errorMsg = `初期セットアップ失敗: ${e.message}\nStack: ${e.stack}`;
    console.error(errorMsg);
    sendErrorNotification("初期セットアップ失敗", errorMsg);
    throw e;
  }
}


/***********************************************
 * ユーティリティ関数
 ***********************************************/

/**
 * 月末を考慮した安全な月加算
 * 
 * 例:
 *   2024/1/31 + 1ヶ月 → 2024/2/29 (うるう年)
 *   2024/8/31 + 6ヶ月 → 2025/2/28
 *   2024/5/15 + 3ヶ月 → 2024/8/15 (通常)
 */
function addMonthsSafely(date, months) {
  const result = new Date(date);
  const originalDay = result.getDate();
  
  // 月を加算
  result.setMonth(result.getMonth() + months);
  
  // 日付がずれた場合（例: 1/31 → 3/3）、前月の末日に修正
  if (result.getDate() !== originalDay) {
    // 当月の0日 = 前月の最終日
    result.setDate(0);
  }
  
  return result;
}

/**
 * 日付を YYYY-MM-DD 形式に変換（比較用）
 */
function formatDateKey(date) {
  return Utilities.formatDate(date, "JST", "yyyy-MM-dd");
}

/**
 * エラー通知メール送信
 */
function sendErrorNotification(subject, errorMessage) {
  try {
    const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    
    GmailApp.sendEmail(
      ADMIN_EMAIL,
      `[GASエラー] ${subject}`,
      `以下のエラーが発生しました:\n\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `エラー内容:\n${errorMessage}\n\n` +
      `発生時刻: ${timestamp}\n` +
      `━━━━━━━━━━━━━━━━━━\n\n` +
      `スプレッドシートを確認してください。`
    );
    
    console.log("✓ エラー通知メール送信完了");
    
  } catch (e) {
    console.error("エラー通知メール送信失敗:", e.message);
  }
}


/***********************************************
 * テスト用関数（開発時のみ使用）
 ***********************************************/

/**
 * addMonthsSafely のテスト
 */
function testAddMonthsSafely() {
  const testCases = [
    { date: new Date(2024, 0, 31), months: 1, expected: "2024-02-29" }, // うるう年
    { date: new Date(2023, 0, 31), months: 1, expected: "2023-02-28" }, // 平年
    { date: new Date(2024, 7, 31), months: 6, expected: "2025-02-28" },
    { date: new Date(2024, 4, 15), months: 3, expected: "2024-08-15" },
    { date: new Date(2024, 0, 30), months: 1, expected: "2024-02-29" },
  ];

  console.log("━━━ addMonthsSafely テスト ━━━");
  
  let passed = 0;
  testCases.forEach((tc, i) => {
    const result = addMonthsSafely(tc.date, tc.months);
    const resultStr = formatDateKey(result);
    const status = resultStr === tc.expected ? "✓ PASS" : "✗ FAIL";
    
    console.log(`Test ${i + 1}: ${status}`);
    console.log(`  入力: ${formatDateKey(tc.date)} + ${tc.months}ヶ月`);
    console.log(`  期待: ${tc.expected}`);
    console.log(`  結果: ${resultStr}\n`);
    
    if (resultStr === tc.expected) passed++;
  });
  
  console.log(`結果: ${passed}/${testCases.length} 件成功`);
}
