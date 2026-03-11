/**
 * LPフォーム → Googleスプレッドシート連携 + 回答者へPDF自動送付
 * 
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成（または既存のシートを開く）
 * 2. ブラウザのURLから「スプレッドシートID」をコピー
 * 3. スプレッドシートのメニュー「拡張機能」→「Apps Script」を開く
 * 4. このファイルの内容をすべてコピーし、Apps Scriptの Code.gs に貼り付ける
 * 5. 下の SPREADSHEET_ID と PDF_FILE_ID を書き換える
 * 6. PDF_FILE_ID の取得方法:
 *    - 稼げるモデルハウス成功事例解説レポート.pdf をGoogleドライブにアップロード
 *    - ファイルの「共有」で「リンクを知っている全員が閲覧可」にしておく（リンク方式の場合必須）
 *    - URLの https://drive.google.com/file/d/【ここがファイルID】/view の「ここがファイルID」をコピー
 * 7. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」を選択（実行ユーザー: 自分、アクセス: 全員）
 * 8. 表示されたウェブアプリのURLを lp-modelhouse-report.html の FORM_SCRIPT_URL に設定
 *
 * 【重要】「次のユーザーとして実行」は必ず「自分」にすること。
 *
 * 【メール送信で「権限がありません」と出るとき】
 * → Apps Scriptエディタで「testSendPdf」を選択して「実行」する。
 *    初回に「〇〇がGoogleアカウントへのアクセスをリクエストしています」と出たら「許可」を押す。
 *    これでメール送信の権限が付与され、フォームからの送信でも送れるようになります。
 *
 * 【メールが届かないとき】
 * - デプロイ設定で「次のユーザーとして実行: 自分」になっているか確認
 * - PDF_FILE_ID が「YOUR_PDF_DRIVE_FILE_ID_HERE」のままになっていないか確認
 * - PDFは「ウェブアプリをデプロイした同じGoogleアカウント」のドライブにアップロードする
 * - 関数「testSendPdf」を選択して実行し、自分宛にテスト送信。エラー内容が「表示」→「実行ログ」に出ます
 * - 送信後、スプレッドシートの「PDF送付」列が「送付済」か「未設定または失敗」で原因を切り分け可能
 */

// ★★★ スプレッドシートのID ★★★
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// ★★★ レポートPDFのGoogleドライブファイルID ★★★
// ドライブにPDFをアップロード → 共有を「リンクを知っている全員が閲覧可」に → URLの/d/と/view/の間がID
const PDF_FILE_ID = 'YOUR_PDF_DRIVE_FILE_ID_HERE';

// シート名（必要に応じて変更可能）
const SHEET_NAME = 'フォーム回答';

// 自動送付メールの差出人名（任意）
const EMAIL_FROM_NAME = 'アルファーブレイン合同会社';

/**
 * 【動作確認用】Apps Scriptエディタで「testSendPdf」を選んで実行してください。
 * 自分のメールアドレスにテスト送信し、PDF取得・Gmail送信ができるか確認できます。
 * 初回は「Gmailで送信する権限」の許可が求められます。
 */
function testSendPdf() {
  const testEmail = Session.getActiveUser().getEmail();
  if (!testEmail) {
    Logger.log('エラー: 実行中のアカウントのメールアドレスを取得できません');
    return;
  }
  if (!PDF_FILE_ID || PDF_FILE_ID === 'YOUR_PDF_DRIVE_FILE_ID_HERE') {
    Logger.log('エラー: PDF_FILE_ID をドライブのファイルIDに設定してください');
    return;
  }
  try {
    sendPdfToRespondent(testEmail, 'テスト', 'テスト会社');
    Logger.log('OK: テストメールを送信しました。' + testEmail + ' の受信トレイ（迷惑メールも）を確認してください。');
  } catch (e) {
    Logger.log('エラー: メール送信に失敗しました。' + e.toString());
  }
}

/**
 * ブラウザでURLを直接開いた場合の表示（GETリクエスト用）
 */
function doGet() {
  return ContentService.createTextOutput('このURLはフォーム送信用です。LPページからお申し込みください。')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * CORS プリフライト用（ブラウザの fetch で必要）
 */
function doOptions() {
  return ContentService.createTextOutput('')
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setHeader('Access-Control-Max-Age', '86400');
}

/**
 * フォームから送信されたデータを受け取り、スプレッドシートに1行追加する
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const params = e.parameter || {};

    const company = params.company || '';
    const name = params.name || '';
    const role = params.role || '';
    const email = params.email || '';
    const phone = params.phone || '';

    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    const headers = getHeaderRow(sheet);
    const row = [
      timestamp,
      company,
      name,
      role,
      email,
      phone
    ];

    sheet.appendRow(row);
    ensurePdfColumnHeader(sheet);

    // 回答者にレポートPDFをメールで送付
    let mailSent = false;
    let statusText = '未設定または失敗';
    if (!email) {
      statusText = 'メールアドレスなし';
    } else if (!PDF_FILE_ID || PDF_FILE_ID === 'YOUR_PDF_DRIVE_FILE_ID_HERE') {
      statusText = '未設定(PDF_FILE_IDを設定してください)';
    } else {
      try {
        sendPdfToRespondent(email, name, company);
        mailSent = true;
        statusText = '送付済';
      } catch (mailErr) {
        var errMsg = (mailErr.message || mailErr.toString()).slice(0, 60);
        statusText = '失敗: ' + errMsg;
        Logger.log('PDF送付エラー: ' + mailErr.toString());
      }
    }
    sheet.getRange(sheet.getLastRow(), 7).setValue(statusText);

    return createJsonResponse(200, {
      result: 'success',
      message: mailSent ? '送信を受け付けました。ご登録のメールアドレスにレポートをお送りしました。' : '送信を受け付けました。'
    });
  } catch (err) {
    return createJsonResponse(500, { result: 'error', message: '送信に失敗しました。' });
  }
}

/**
 * 回答者にメールでPDFのダウンロードリンクを送付（DriveAppは使用しない）
 */
function sendPdfToRespondent(toEmail, name, company) {
  const subject = '【アルファーブレイン】稼げるモデルハウス成功事例解説レポート';
  const pdfUrl = 'https://drive.google.com/file/d/' + PDF_FILE_ID + '/view?usp=sharing';
  const body =
    (name ? name + ' 様\n\n' : '') +
    'このたびはレポートのダウンロードをお申し込みいただきありがとうございます。\n\n' +
    '下記リンクから「稼げるモデルハウス成功事例解説レポート」をご覧・ダウンロードいただけます。\n\n' +
    pdfUrl + '\n\n' +
    '※「プレビューできません」と出る場合は、画面の「ダウンロード」ボタンから保存してください。\n\n' +
    'ご確認のほどよろしくお願いいたします。\n\n' +
    '────────────────\n' +
    EMAIL_FROM_NAME + '\n';
  const options = { name: EMAIL_FROM_NAME };
  try {
    GmailApp.sendEmail(toEmail, subject, body, options);
  } catch (e) {
    MailApp.sendEmail(toEmail, subject, body, options);
  }
}

/**
 * スプレッドシートのシートを取得。なければ作成し、1行目にヘッダーを書く
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['送信日時', '会社名', 'ご担当者名', '役職', 'メールアドレス', '電話番号', 'PDF送付']);
    sheet.getRange('1:1').setFontWeight('bold');
  }
  return sheet;
}

/** 既存シートに「PDF送付」列があるか確認し、なければヘッダーを追加 */
function ensurePdfColumnHeader(sheet) {
  if (sheet.getLastRow() < 1) return;
  if (sheet.getRange(1, 7).getValue() !== 'PDF送付') {
    sheet.getRange(1, 7).setValue('PDF送付').setFontWeight('bold');
  }
}

/**
 * ヘッダー行があるか確認し、なければ1行目にヘッダーを書く
 */
function getHeaderRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(['送信日時', '会社名', 'ご担当者名', '役職', 'メールアドレス', '電話番号']);
    sheet.getRange('1:1').setFontWeight('bold');
  }
  return sheet.getRange(1, 1, 1, 6).getValues()[0];
}

/**
 * JSONレスポンス（CORS対応）
 */
function createJsonResponse(status, body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}
