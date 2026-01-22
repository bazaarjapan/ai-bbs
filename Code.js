/*
 * MIT License
 *
 * Copyright (c) 2024-2025 Noboru Ando @ Aoyama Gakuin University
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * ---------------------------------------------------------------
 * アプリ名         : Googleスプレッドシート掲示板システム
 * リポジトリ名     : AIBBS11
 * 作成・著作者     : 安藤昇＠青山学院大学
 * 作成日           : 2025/03/04
 * 更新日           : 2026/01/22
 * ---------------------------------------------------------------
 */

// グローバル変数
let SHEET;
let CONFIG_SHEET;
let PASSWORD;
let DRIVE_FOLDER_NAME;
let DRIVE_FOLDER;
let POSTS_PER_PAGE; // 1ページあたりの投稿表示件数

// 初期化済みフラグ
let initialized = false;

/**
 * 初期化関数 - 必要な変数を設定
 */
function initialize() {
  // すでに初期化済みの場合はスキップ
  if (initialized) {
    return true;
  }
  try {
    // アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // データ保存用のシート（シート1）を取得
    SHEET = ss.getSheetByName('シート1');
    if (!SHEET) {
      SHEET = ss.getSheets()[0]; // シート1がない場合は最初のシートを使用
    }

    // 設定用のシート（環境設定シート）を取得
    CONFIG_SHEET = ss.getSheetByName('環境設定');
    if (!CONFIG_SHEET) {
      // 環境設定シートがない場合は作成
      CONFIG_SHEET = ss.insertSheet('環境設定');
      CONFIG_SHEET.getRange('A1').setValue('パスワード');
      CONFIG_SHEET.getRange('A2').setValue('保存フォルダ名');
      CONFIG_SHEET.getRange('A3').setValue('表示件数');
      CONFIG_SHEET.getRange('B1').setValue('9999'); // デフォルトパスワード
      CONFIG_SHEET.getRange('B2').setValue('bbs_files'); // デフォルトフォルダ名
      CONFIG_SHEET.getRange('B3').setValue(5); // デフォルト表示件数
    }

    // 環境設定値を一度に取得（高速化）
    const configValues = CONFIG_SHEET.getRange('B1:B3').getValues();

    // パスワードを取得（数値の場合は文字列に変換）
    const passwordValue = configValues[0][0];
    PASSWORD = typeof passwordValue === 'number' ? String(passwordValue) : passwordValue;

    // 画像とファイル保存用フォルダ名を取得
    DRIVE_FOLDER_NAME = configValues[1][0];

    // 表示件数を取得
    const postsPerPageValue = configValues[2][0];
    POSTS_PER_PAGE = Number(postsPerPageValue) || 5; // 数値に変換、無効な場合はデフォルト5

    // Googleドライブのフォルダを取得または作成
    DRIVE_FOLDER = getDriveFolder(DRIVE_FOLDER_NAME);

    // シート1のヘッダーが存在しない場合は作成
    if (SHEET.getLastRow() === 0) {
      SHEET.appendRow(['ID', 'タイトル', '投稿文', '投稿日時', '更新日時', '表示']);
    }

    // 初期化完了フラグを設定
    initialized = true;

    return true;
  } catch (e) {
    Logger.log('初期化エラー: ' + e.toString());
    return false;
  }
}

/**
 * 指定された名前のGoogleドライブフォルダを取得または作成
 * 画像とファイルの両方の保存に使用されます
 */
function getDriveFolder(folderName) {
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    let mainFolder = null;
    if (folders.hasNext()) {
      mainFolder = folders.next();
      // 重複フォルダが存在する場合はゴミ箱へ移動
      while (folders.hasNext()) {
        const duplicateFolder = folders.next();
        duplicateFolder.setTrashed(true);
      }
      return mainFolder;
    }
    const folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  } catch (e) {
    Logger.log('フォルダ取得エラー: ' + e.toString());
    throw e;
  }
}

/**
 * doGet - Webアプリとして公開した際のエントリーポイント
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('掲示板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// キャッシュ用の変数
let postsCache = null;
let postsCacheTimestamp = 0;
const CACHE_TTL = 5 * 60 * 1000; // キャッシュの有効期間（5分）

/**
 * 投稿一覧を取得する関数
 */
function getPosts(page = 1, postsPerPage = null) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  // 表示件数が指定されていない場合は環境設定から取得
  if (postsPerPage === null) {
    postsPerPage = POSTS_PER_PAGE;
  }

  try {
    const now = new Date().getTime();

    // キャッシュが有効かチェック
    if (postsCache && (now - postsCacheTimestamp < CACHE_TTL)) {
      // キャッシュからデータを取得
      const posts = postsCache;

      // ページネーション
      const totalPosts = posts.length;
      const totalPages = Math.ceil(totalPosts / postsPerPage);
      const startIndex = (page - 1) * postsPerPage;
      const paginatedPosts = posts.slice(startIndex, startIndex + postsPerPage);

      return {
        posts: paginatedPosts,
        pagination: {
          currentPage: page,
          totalPages: totalPages,
          totalPosts: totalPosts
        }
      };
    }

    // キャッシュが無効な場合はデータを取得
    // データ行数をチェック
    const lastRow = SHEET.getLastRow();

    // データがない場合（ヘッダー行のみ）は空の結果を返す
    if (lastRow <= 1) {
      const emptyResult = {
        posts: [],
        pagination: {
          currentPage: page,
          totalPages: 0,
          totalPosts: 0
        }
      };

      // キャッシュを更新
      postsCache = [];
      postsCacheTimestamp = now;

      return emptyResult;
    }

    // データ範囲を取得（最適化：必要な列のみ取得）
    const dataRange = SHEET.getRange(2, 1, lastRow - 1, 6);
    const values = dataRange.getValues();

    // 表示フラグが1の投稿のみをフィルタリング
    const posts = values
      .filter(row => row[5] === 1) // 表示フラグが1の投稿のみを取得
      .map((row, index) => {
        return {
          id: row[0] || (index + 1),
          name: row[1] || '',
          text: row[2] || '',
          createdAt: row[3] ? new Date(row[3]).toLocaleString() : '',
          updatedAt: row[4] ? new Date(row[4]).toLocaleString() : ''
        };
      });

    // 投稿を新しい順に並べ替え
    posts.reverse();

    // キャッシュを更新
    postsCache = posts;
    postsCacheTimestamp = now;

    // ページネーション
    const totalPosts = posts.length;
    const totalPages = Math.ceil(totalPosts / postsPerPage);
    const startIndex = (page - 1) * postsPerPage;
    const paginatedPosts = posts.slice(startIndex, startIndex + postsPerPage);

    return {
      posts: paginatedPosts,
      pagination: {
        currentPage: page,
        totalPages: totalPages,
        totalPosts: totalPosts
      }
    };
  } catch (e) {
    Logger.log('投稿取得エラー: ' + e.toString());
    return { error: '投稿の取得に失敗しました: ' + e.toString() };
  }
}

/**
 * 新規投稿を保存する関数
 */
function savePost(postData, password) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  // パスワード認証
  if (password !== PASSWORD) {
    return { error: 'パスワードが正しくありません' };
  }

  // スプレッドシートのロックを取得
  const lock = LockService.getDocumentLock();

  try {
    // ロックを取得（最大10秒待機）
    if (!lock.tryLock(10000)) {
      return { error: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }

    // 現在時刻
    const now = new Date();

    // 最新のデータを取得して新しいIDを生成（同時書き込み対策）
    const lastRow = SHEET.getLastRow();
    let newId = 1;

    if (lastRow > 1) {
      // 最新のIDデータを取得
      const ids = SHEET.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      newId = Math.max(...ids.map(id => Number(id) || 0)) + 1;
    }

    // 新規投稿をシートに追加
    SHEET.appendRow([
      newId,
      postData.name,
      postData.text,
      now,
      now,
      1  // 表示フラグ（1=表示する）
    ]);

    // キャッシュをクリア
    postsCache = null;

    return { success: true, message: '投稿が保存されました' };
  } catch (e) {
    Logger.log('投稿保存エラー: ' + e.toString());
    return { error: '投稿の保存に失敗しました: ' + e.toString() };
  } finally {
    // 必ずロックを解放
    lock.releaseLock();
  }
}

/**
 * 投稿を更新する関数
 */
function updatePost(postData, password) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  // パスワード認証
  if (password !== PASSWORD) {
    return { error: 'パスワードが正しくありません' };
  }

  // スプレッドシートのロックを取得
  const lock = LockService.getDocumentLock();

  try {
    // ロックを取得（最大10秒待機）
    if (!lock.tryLock(10000)) {
      return { error: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }

    // 最新のデータを取得（同時編集対策）- ID列のみ取得して最適化
    const lastRow = SHEET.getLastRow();
    const idValues = SHEET.getRange(2, 1, lastRow - 1, 1).getValues();

    // 投稿IDに一致する行を検索
    let rowIndex = -1;
    for (let i = 0; i < idValues.length; i++) {
      if (idValues[i][0] == postData.id) {
        rowIndex = i + 2; // シートの行番号は1から始まり、ヘッダー行があるため+2
        break;
      }
    }

    if (rowIndex === -1) {
      return { error: '指定された投稿が見つかりません' };
    }

    // 現在時刻（更新日時）
    const now = new Date();

    // 作成日時を取得
    const createdAt = SHEET.getRange(rowIndex, 4, 1, 1).getValue();

    // 投稿を更新
    SHEET.getRange(rowIndex, 2, 1, 3).setValues([[
      postData.name,
      postData.text,
      createdAt, // 作成日時は変更しない
    ]]);

    // 更新日時を設定
    SHEET.getRange(rowIndex, 5).setValue(now);

    // キャッシュをクリア
    postsCache = null;

    return { success: true, message: '投稿が更新されました' };
  } catch (e) {
    Logger.log('投稿更新エラー: ' + e.toString());
    return { error: '投稿の更新に失敗しました: ' + e.toString() };
  } finally {
    // 必ずロックを解放
    lock.releaseLock();
  }
}

/**
 * 投稿を削除する関数
 */
function deletePost(postId, password) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  // パスワード認証
  if (password !== PASSWORD) {
    return { error: 'パスワードが正しくありません' };
  }

  // スプレッドシートのロックを取得
  const lock = LockService.getDocumentLock();

  try {
    // ロックを取得（最大10秒待機）
    if (!lock.tryLock(10000)) {
      return { error: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }

    // 最新のデータを取得（同時削除対策）- ID列のみ取得して最適化
    const lastRow = SHEET.getLastRow();
    const idValues = SHEET.getRange(2, 1, lastRow - 1, 1).getValues();

    // 投稿IDに一致する行を検索
    let rowIndex = -1;
    for (let i = 0; i < idValues.length; i++) {
      if (idValues[i][0] == postId) {
        rowIndex = i + 2; // シートの行番号は1から始まり、ヘッダー行があるため+2
        break;
      }
    }

    if (rowIndex === -1) {
      return { error: '指定された投稿が見つかりません' };
    }

    // 表示フラグを0に設定（非表示にする）
    SHEET.getRange(rowIndex, 6).setValue(0);

    // キャッシュをクリア
    postsCache = null;

    return { success: true, message: '投稿が非表示になりました' };
  } catch (e) {
    Logger.log('投稿削除エラー: ' + e.toString());
    return { error: '投稿の削除に失敗しました: ' + e.toString() };
  } finally {
    // 必ずロックを解放
    lock.releaseLock();
  }
}

/**
 * 画像をGoogleドライブにアップロードする関数
 */
function uploadImage(base64Data) {
  if (!initialize()) {
    throw new Error('初期化に失敗しました');
  }

  try {
    // Base64データからBlobを作成
    // データURLの形式（例：data:image/jpeg;base64,/9j/4AAQ...）からBase64部分を抽出
    const matches = base64Data.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);

    if (!matches || matches.length !== 3) {
      throw new Error('無効な画像データ形式です');
    }

    const contentType = matches[1];
    const base64EncodedData = matches[2];
    const decodedData = Utilities.base64Decode(base64EncodedData);
    const blob = Utilities.newBlob(decodedData, contentType);

    // ファイル名を生成（タイムスタンプを含む）
    const fileName = 'image_' + new Date().getTime();
    blob.setName(fileName);

    // Googleドライブにファイルをアップロード
    const file = DRIVE_FOLDER.createFile(blob);

    // ファイルを「リンクを知っている人全員」に共有設定
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ファイルのURLを返す
    return file.getUrl();
  } catch (e) {
    Logger.log('画像アップロードエラー: ' + e.toString());
    throw e;
  }
}

/**
 * Google Drive画像URLを処理する関数w640-h480
 */
function formatGoogleDriveUrl(url) {
  if (!url) return '';

  try {
    // Google DriveのURLかどうかを確認
    if (url.includes('drive.google.com')) {
      // ファイルIDを抽出
      const fileId = url.match(/[-\w]{25,}/);
      if (fileId) {
        // 新しい形式のURLを生成
        return `https://lh3.google.com/u/0/d/${fileId[0]}=w640-h480-iv1`;
      }
    }
    return url;
  } catch (e) {
    Logger.log('Error processing Google Drive URL:', e);
    return url;
  }
}

/**
 * エディタ用の画像をアップロードする関数
 */
function uploadImageToEditor(base64Image) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  try {
    // 画像のアップロード処理
    const imageUrl = uploadImage(base64Image);

    // 成功結果を返す
    return {
      success: true,
      imageUrl: formatGoogleDriveUrl(imageUrl) // 表示用にURLをフォーマット
    };
  } catch (e) {
    Logger.log('エディタ画像アップロードエラー: ' + e.toString());
    return { error: '画像のアップロードに失敗しました: ' + e.toString() };
  }
}

/**
 * パスワードを検証する関数
 */
function verifyPassword(password) {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }

  return { valid: password === PASSWORD };
}

/**
 * 表示件数設定を取得する関数
 */
function getDisplaySettings() {
  if (!initialize()) {
    return { error: '初期化に失敗しました' };
  }
  return { postsPerPage: POSTS_PER_PAGE };
}

/**
 * 表示件数設定を更新する関数
 */
function updatePostsPerPage(count) {
  if (!initialize()) {
    return { success: false, error: '初期化に失敗しました' };
  }

  const allowed = [5, 10, 20, 30];
  const numericCount = Number(count);
  if (!allowed.includes(numericCount)) {
    return { success: false, error: '無効な表示件数です' };
  }

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, error: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }

    CONFIG_SHEET.getRange('B3').setValue(numericCount);
    POSTS_PER_PAGE = numericCount;
    postsCache = null;
    postsCacheTimestamp = 0;

    return { success: true, postsPerPage: POSTS_PER_PAGE };
  } catch (e) {
    Logger.log('表示件数更新エラー: ' + e.toString());
    return { success: false, error: '表示件数の更新に失敗しました: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * HTMLに含めるJavaScriptを取得する関数
 */
function getScriptContent() {
  return HtmlService.createHtmlOutputFromFile('javascript').getContent();
}

/**
 * HTMLに含めるCSSを取得する関数
 */
function getStyleContent() {
  return HtmlService.createHtmlOutputFromFile('stylesheet').getContent();
}

/**
 * ファイルをGoogleドライブにアップロードする関数
 * @param {string} fileName - ファイル名
 * @param {string} fileData - Base64エンコードされたファイルデータ
 * @return {Object} アップロード結果
 */
function uploadFileToDrive(fileName, fileData) {
  if (!initialize()) {
    return {
      success: false,
      error: '初期化に失敗しました'
    };
  }

  try {
    // Base64データからバイナリデータを取得
    const contentType = fileData.match(/^data:(.+);base64,/)[1];
    const base64Data = fileData.replace(/^data:(.+);base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType);

    // ファイル名を設定（元のファイル名を保持）
    const safeFileName = fileName.replace(/[^\w\s.-]/g, '_'); // 安全なファイル名に変換
    const uniqueFileName = safeFileName + '_' + new Date().getTime(); // タイムスタンプを追加
    blob.setName(uniqueFileName);

    // 環境設定から取得したフォルダを使用
    const filesFolder = DRIVE_FOLDER;

    // ファイルをアップロード
    const file = filesFolder.createFile(blob);

    // ファイルを公開設定に変更
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ファイルURLを返す
    return {
      success: true,
      fileUrl: file.getUrl(),
      fileName: fileName
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
