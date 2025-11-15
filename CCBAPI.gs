/***********************************************************
 * ポイントサイト横断｜クレカ案件トラッキング自動化 GAS
 * main.gs（Credit Card Bonuses API 対応版）
 *
 * 使用API: Credit Card Bonuses API (GitHub)
 * https://github.com/andenacitelli/credit-card-bonuses-api
 *
 * 含まれる機能：
 * - ランキング更新（GitHub API利用）
 * - 契約管理シートのテンプレ生成
 * - 契約記録登録（安全な日付計算）
 * - 解約推奨日のカレンダー登録（重複防止）
 * - メール通知（重複防止）
 * - トリガー自動設定
 ***********************************************************/

const API_URL = "https://raw.githubusercontent.com/andenacitelli/credit-card-bonuses-api/main/exports/data.json";
const ADMIN_EMAIL = Session.getActiveUser().getEmail();

/***********************************************
 * 1. ランキング更新：GitHub APIから取得
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
      timeoutSeconds: 30
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
      "順位",
      "カード名",
      "発行会社",
      "年会費",
      "ボーナス額",
      "必要利用額",
      "達成日数",
      "ネットワーク",
      "ビジネスカード",
      "URL",
      "備考"
    ];
    sheet.appendRow(headerRow);

    // データを処理してランキング作成（ボーナス額でソート）
    const processedData = data
      .filter(card => {
        // 有効なオファーがあるカードのみ
        return card.offers && card.offers.length > 0 && !card.discontinued;
      })
      .map(card => {
        const offer = card.offers[0]; // 最初のオファーを使用
        const bonusAmount = offer.amount && offer.amount.length > 0 ? offer.amount[0].amount : 0;
        
        return {
          card,
          offer,
          bonusAmount,
          bonusValue: calculateBonusValue(bonusAmount, card.currency)
        };
      })
      .sort((a, b) => b.bonusValue - a.bonusValue) // ボーナス価値でソート
      .slice(0, 50); // 上位50件

    // データ行をバッチで準備
    const dataRows = [];
    processedData.forEach((item, i) => {
      const { card, offer } = item;
      
      // ボーナス額の表示
      let bonusDisplay = "";
      if (offer.amount && offer.amount.length > 0) {
        const amt = offer.amount[0];
        bonusDisplay = formatBonus(amt.amount, amt.currency || card.currency);
      }

      // 年会費の表示
      const annualFee = card.isAnnualFeeWaived 
        ? `$${card.annualFee} (初年度無料)` 
        : `$${card.annualFee}`;

      // 備考欄
      let notes = [];
      if (card.isAnnualFeeWaived) notes.push("初年度年会費無料");
      if (offer.details) notes.push(offer.details);
      if (card.credits && card.credits.length > 0) {
        const creditValue = card.credits.reduce((sum, c) => sum + (c.value || 0), 0);
        notes.push(`特典総額: $${creditValue}`);
      }

      dataRows.push([
        i + 1,
        card.name,
        formatIssuer(card.issuer),
        annualFee,
        bonusDisplay,
        offer.spend ? `$${offer.spend.toLocaleString()}` : "不要",
        offer.days ? `${offer.days}日` : "無制限",
        formatNetwork(card.network),
        card.isBusiness ? "ビジネス" : "個人",
        card.url || "",
        notes.join(" | ") || "-"
      ]);
    });

    // バッチ書き込み
    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headerRow.length);
    
    // ヘッダーのスタイル設定
    sheet.getRange(1, 1, 1, headerRow.length)
      .setBackground("#4a86e8")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
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
    "獲得方法",
    "URL",
    "備考",
    "ボーナス獲得条件",
    "解約推奨日の算出根拠",
    "通知済み",
    "カレンダー登録済み"
  ];

  sheet.appendRow(header);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, header.length, 120);

  const comments = {
    "カード名": "ランキングシートからコピー",
    "契約日": "申し込み完了日 (YYYY/MM/DD)",
    "解約推奨日": "自動計算 or 手動入力",
    "獲得ポイント": "実際に獲得したポイント額",
    "獲得方法": "例: Chase UR、Amex MR など",
    "URL": "カード申込URL",
    "備考": "自由記入欄",
    "ボーナス獲得条件": "例: $5,000利用（90日以内）",
    "解約推奨日の算出根拠": "例: 契約から12ヶ月後",
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
 * 3. ランキングから契約管理へ自動転記
 ***********************************************/
function transferFromRanking(rowNumber) {
  const ss = SpreadsheetApp.getActive();
  const rankingSheet = ss.getSheetByName("ランキング");
  const contractSheet = ss.getSheetByName("契約管理");
  
  if (!rankingSheet) {
    throw new Error("ランキングシートが見つかりません。mainUpdate() を実行してください。");
  }
  
  if (!contractSheet) {
    createContractManagementSheet();
  }
  
  // ランキングシートから指定行のデータを取得
  const row = rankingSheet.getRange(rowNumber, 1, 1, 11).getValues()[0];
  
  const cardName = row[1]; // カード名
  const issuer = row[2]; // 発行会社
  const annualFee = row[3]; // 年会費
  const bonus = row[4]; // ボーナス額
  const spend = row[5]; // 必要利用額
  const days = row[6]; // 達成日数
  const url = row[9]; // URL
  const notes = row[10]; // 備考
  
  // 契約管理シートに追加
  recordContract({
    cardName: `${cardName} (${issuer})`,
    points: bonus,
    site: "申込サイト名を入力",
    url: url,
    memo: notes,
    conditions: `${spend}利用（${days}）`,
    basis: "ランキングから自動転記",
    monthsUntilCancel: 12
  });
  
  console.log(`✓ ${cardName} を契約管理シートに転記しました`);
}


/***********************************************
 * 4. 契約記録の登録
 ***********************************************/
function recordContract({ 
  cardName, 
  points, 
  site, 
  url, 
  memo, 
  conditions, 
  basis,
  monthsUntilCancel = 12 // デフォルト12ヶ月（米国は通常1年）
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
 * 5. 解約推奨日を Google カレンダーへ登録
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
    throw new Error("必須カラムが見つかりません");
  }

  const cal = CalendarApp.getDefaultCalendar();
  let addedCount = 0;
  let skippedCount = 0;

  const updatesToWrite = [];

  values.forEach((row, index) => {
    const card = row[idxCard];
    const cancel = row[idxCancel];
    const registered = row[idxRegistered];

    if (!card || !(cancel instanceof Date)) {
      skippedCount++;
      return;
    }

    if (registered === "済") {
      skippedCount++;
      return;
    }

    try {
      const eventTitle = `【解約推奨】${card}`;

      cal.createAllDayEvent(
        eventTitle,
        cancel,
        {
          description: `${card} の解約推奨日です。\n\nカードの利用状況を確認し、必要に応じて解約手続きを行ってください。\n\n※このリマインダーは自動生成されました`,
          reminders: [
            { method: "popup", minutes: 30 * 24 * 60 }, // 30日前
            { method: "popup", minutes: 7 * 24 * 60 },  // 1週間前
            { method: "popup", minutes: 1 * 24 * 60 }   // 前日
          ]
        }
      );
      
      if (idxRegistered !== -1) {
        updatesToWrite.push({
          row: index + 2,
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

  updatesToWrite.forEach(update => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });

  console.log(`✓ カレンダー登録完了: ${addedCount}件追加, ${skippedCount}件スキップ`);
}


/***********************************************
 * 6. メール通知（重複防止版）
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
  today.setHours(0, 0, 0, 0);
  
  const week = new Date(today);
  week.setDate(today.getDate() + 7);
  
  const month = new Date(today);
  month.setDate(today.getDate() + 30);

  const todayStr = formatDateKey(today);
  const weekStr = formatDateKey(week);
  const monthStr = formatDateKey(month);

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

    const cancelDate = new Date(cancel);
    cancelDate.setHours(0, 0, 0, 0);
    const cancelStr = formatDateKey(cancelDate);

    // 今日、1週間後、30日後のいずれかで未通知
    const shouldNotify = (
      (cancelStr === todayStr || cancelStr === weekStr || cancelStr === monthStr) 
      && !notified.includes(cancelStr)
    );

    if (shouldNotify) {
      try {
        const daysUntil = Math.ceil((cancelDate - today) / (1000 * 60 * 60 * 24));
        let urgency = "";
        
        if (daysUntil === 0) urgency = "【本日】";
        else if (daysUntil <= 7) urgency = `【あと${daysUntil}日】`;
        else urgency = `【あと${daysUntil}日】`;

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
 * 7. 毎朝のトリガー自動設定
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

  // メール通知: 毎朝7時
  ScriptApp.newTrigger("sendCancelReminderEmails")
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  console.log("✓ トリガー設定: sendCancelReminderEmails (毎朝7時)");

  // ランキング更新: 毎朝8時
  ScriptApp.newTrigger("mainUpdate")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  console.log("✓ トリガー設定: mainUpdate (毎朝8時)");
}


/***********************************************
 * 8. 初期セットアップ（初回実行用）
 ***********************************************/
function initialSetup() {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("  初期セットアップ開始");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
  
  try {
    console.log("\n[1/3] 契約管理シート作成中...");
    createContractManagementSheet();
    
    console.log("\n[2/3] トリガー設定中...");
    setupDailyTriggers();
    
    console.log("\n[3/3] ランキング初回取得中...");
    mainUpdate();
    
    console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━");
    console.log("  ✓ セットアップ完了！");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━");
    console.log("\n【次のステップ】");
    console.log("1. ランキングシートで気になるカードを確認");
    console.log("2. transferFromRanking(行番号) で契約管理へ転記");
    console.log("3. registerCancelDateToCalendar() でカレンダー登録");
    
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
 */
function addMonthsSafely(date, months) {
  const result = new Date(date);
  const originalDay = result.getDate();
  
  result.setMonth(result.getMonth() + months);
  
  if (result.getDate() !== originalDay) {
    result.setDate(0);
  }
  
  return result;
}

/**
 * 日付を YYYY-MM-DD 形式に変換
 */
function formatDateKey(date) {
  return Utilities.formatDate(date, "JST", "yyyy-MM-dd");
}

/**
 * ボーナス額の表示形式を整形
 */
function formatBonus(amount, currency) {
  if (!amount) return "情報なし";
  
  switch(currency) {
    case "USD":
      return `$${amount.toLocaleString()}`;
    case "CHASE":
    case "AMERICAN_EXPRESS":
    case "CAPITAL_ONE":
      return `${amount.toLocaleString()} pts`;
    default:
      return `${amount.toLocaleString()} ${currency}`;
  }
}

/**
 * 発行会社名を日本語化
 */
function formatIssuer(issuer) {
  const issuerMap = {
    "AMERICAN_EXPRESS": "アメックス",
    "CHASE": "チェース",
    "CITI": "シティ",
    "CAPITAL_ONE": "キャピタルワン",
    "BANK_OF_AMERICA": "バンカメ",
    "BARCLAYS": "バークレイズ",
    "BREX": "Brex"
  };
  return issuerMap[issuer] || issuer;
}

/**
 * ネットワーク名を日本語化
 */
function formatNetwork(network) {
  const networkMap = {
    "VISA": "Visa",
    "MASTERCARD": "Mastercard",
    "AMERICAN_EXPRESS": "Amex"
  };
  return networkMap[network] || network;
}

/**
 * ボーナス価値を計算（ソート用）
 */
function calculateBonusValue(amount, currency) {
  if (!amount) return 0;
  
  // 通貨ごとの価値係数（1ポイント = X ドル）
  const valueMap = {
    "USD": 1,
    "CHASE": 0.01,
    "AMERICAN_EXPRESS": 0.01,
    "CAPITAL_ONE": 0.01,
    "CITI": 0.01,
    "DELTA": 0.012,
    "UNITED": 0.012,
    "SOUTHWEST": 0.014,
    "MARRIOTT": 0.008,
    "HILTON": 0.005,
    "IHG": 0.005,
    "HYATT": 0.015
  };
  
  const coefficient = valueMap[currency] || 0.01;
  return amount * coefficient;
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
