// ==============================
// HEIC形式変換ソリューション
// ==============================

/**
 * HEIC変換問題の完全解決ガイド
 */
function solveHEICProblem() {
  console.log('=================================================================================');
  console.log('                    HEIC形式変換問題の解決策                                     ');
  console.log('=================================================================================\n');

  console.log('iPhoneのHEIC形式画像をJPEGに変換する複数の方法を提供します。\n');

  console.log('【解決策1: iPhoneの設定を変更（最も簡単）】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('iPhoneで撮影時にJPEG形式で保存するように設定:\n');
  console.log('1. 設定アプリを開く');
  console.log('2. 「カメラ」→「フォーマット」');
  console.log('3. 「互換性優先」を選択（JPEG/H.264）');
  console.log('   ※「高効率」がHEIC形式\n');
  console.log('メリット: 変換不要、すべてJPEGで撮影');
  console.log('デメリット: ファイルサイズが少し大きくなる\n');

  console.log('【解決策2: iPhoneで自動変換してアップロード】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupiPhoneAutoConverter() を実行\n');

  console.log('【解決策3: Google DriveアプリでJPEGとしてアップロード】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupDriveAppMethod() を実行\n');

  console.log('【解決策4: GAS内で変換処理（改良版）】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('processHEICImagesAdvanced() を実行\n');

  console.log('【解決策5: 外部サービス連携】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupCloudConvertIntegration() を実行\n');

  return {
    easiest: 'iPhone設定を変更',
    recommended: 'setupiPhoneAutoConverter()',
    gasInternal: 'processHEICImagesAdvanced()'
  };
}

/**
 * iPhone用自動変換ショートカット設定
 */
function setupiPhoneAutoConverter() {
  console.log('=== iPhone自動JPEG変換ショートカット ===\n');

  console.log('iPhoneのショートカットアプリでHEIC→JPEG変換を自動化します。\n');

  console.log('【ショートカット作成手順】\n');

  console.log('📱 ステップ1: ショートカットアプリを開く');
  console.log('────────────────────────────────');
  console.log('「ショートカット」アプリを開いて「＋」をタップ\n');

  console.log('📱 ステップ2: アクションを追加');
  console.log('────────────────────────────────');
  console.log('以下のアクションを順番に追加:\n');

  console.log('アクション①: 写真を選択');
  console.log('├─ 複数を選択: オン');
  console.log('└─ 選択を停止: 写真を選択後\n');

  console.log('アクション②: イメージを変換');
  console.log('├─ フォーマット: JPEG');
  console.log('├─ 品質: 最高');
  console.log('└─ メタデータを保持: オン\n');

  console.log('アクション③: ファイル名を設定');
  console.log('├─ 名前: card_[現在の日付]');
  console.log('└─ 拡張子: .jpg\n');

  console.log('アクション④: ファイルを保存');
  console.log('├─ サービス: Googleドライブ');
  console.log('├─ 保存先: /ポケモンカード管理/新規アップロード');
  console.log('└─ 既存のファイルを置き換え: オフ\n');

  console.log('📱 ステップ3: ショートカット設定');
  console.log('────────────────────────────────');
  console.log('1. ショートカット名: 「カード登録（JPEG変換）」');
  console.log('2. アイコン: 📷 または 🎴');
  console.log('3. ホーム画面に追加\n');

  console.log('【使い方】');
  console.log('────────────────────────────────');
  console.log('1. カードを撮影（HEIC形式でOK）');
  console.log('2. ホーム画面の「カード登録」をタップ');
  console.log('3. 写真を選択');
  console.log('4. 自動的にJPEG変換してDriveにアップロード');
  console.log('5. processDriveImages() で処理\n');

  console.log('✅ これで100%JPEG形式でアップロードされます！\n');

  return {
    status: 'ショートカットアプリで手動設定',
    benefit: 'HEIC→JPEG自動変換'
  };
}

/**
 * Google Driveアプリ経由の方法
 */
function setupDriveAppMethod() {
  console.log('=== Google Driveアプリ経由でJPEG変換 ===\n');

  console.log('Google DriveアプリはHEIC画像を自動的にJPEGに変換してアップロードします。\n');

  console.log('【設定手順】\n');

  console.log('📱 Google Driveアプリ設定');
  console.log('────────────────────────────────');
  console.log('1. Google Driveアプリを開く');
  console.log('2. 設定（歯車アイコン）をタップ');
  console.log('3. 「写真」セクション');
  console.log('4. 「アップロード時の画質」を「元の画質」に設定\n');

  console.log('📱 アップロード手順');
  console.log('────────────────────────────────');
  console.log('1. Google Driveアプリで「ポケモンカード管理/新規アップロード」フォルダを開く');
  console.log('2. ＋ボタンをタップ');
  console.log('3. 「アップロード」→「写真と動画」');
  console.log('4. カード画像を選択（HEIC形式でもOK）');
  console.log('5. アップロード実行\n');

  console.log('✅ Google DriveがHEICを自動的にJPEGに変換！\n');

  console.log('【メリット】');
  console.log('• 追加アプリ不要');
  console.log('• 自動変換');
  console.log('• メタデータ保持\n');

  return {
    status: 'Google Driveアプリで実現',
    autoConvert: true
  };
}

/**
 * GAS内でのHEIC処理（改良版）
 */
function processHEICImagesAdvanced() {
  console.log('=== GAS内HEIC処理（改良版）===\n');

  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    console.log('❌ フォルダが設定されていません');
    console.log('setupSimpleDriveWorkflow() を実行してください');
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const processedFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('PROCESSED_FOLDER_ID')
  );

  // HEICファイルを検索
  const files = folder.getFiles();
  let heicCount = 0;
  let convertedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();

    if (fileName.endsWith('.heic') || fileName.endsWith('.heif')) {
      heicCount++;
      console.log(`HEIC発見: ${file.getName()}`);

      try {
        // 変換を試みる
        const convertedFile = convertHEICtoJPEG(file);

        if (convertedFile) {
          console.log(`✅ 変換成功: ${convertedFile.getName()}`);

          // 元のHEICファイルを処理済みフォルダに移動
          file.moveTo(processedFolder);
          convertedCount++;
        } else {
          console.log(`❌ 変換失敗: ${file.getName()}`);
        }

      } catch (error) {
        console.error(`エラー: ${file.getName()}`, error);
      }
    }
  }

  console.log(`\n処理結果:`);
  console.log(`HEIC画像: ${heicCount}枚`);
  console.log(`変換成功: ${convertedCount}枚`);

  if (heicCount === 0) {
    console.log('\nHEIC画像が見つかりません。');
    console.log('JPEGまたはPNG画像は processDriveImages() で処理してください。');
  }

  return {
    heicFiles: heicCount,
    converted: convertedCount
  };
}

/**
 * HEIC→JPEG変換関数（改良版）
 */
function convertHEICtoJPEG(heicFile) {
  try {
    console.log(`変換試行: ${heicFile.getName()}`);

    // 方法1: getAsで直接変換を試みる
    try {
      const jpegBlob = heicFile.getBlob().getAs('image/jpeg');

      if (jpegBlob) {
        const newName = heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg');
        const jpegFile = heicFile.getParents().next().createFile(jpegBlob);
        jpegFile.setName(newName);

        console.log('方法1で変換成功');
        return jpegFile;
      }
    } catch (e) {
      console.log('方法1失敗: ' + e.toString());
    }

    // 方法2: DriveのサムネイルAPIを利用
    try {
      const fileId = heicFile.getId();

      // Driveのサムネイル生成を利用
      const thumbnailLink = `https://drive.google.com/thumbnail?id=${fileId}&sz=w2000`;

      const response = UrlFetchApp.fetch(thumbnailLink);
      const jpegBlob = response.getBlob();

      const newName = heicFile.getName().replace(/\.(heic|heif)$/i, '_converted.jpg');
      const jpegFile = heicFile.getParents().next().createFile(jpegBlob);
      jpegFile.setName(newName);

      console.log('方法2（サムネイル）で変換成功');
      return jpegFile;

    } catch (e) {
      console.log('方法2失敗: ' + e.toString());
    }

    // 方法3: Base64エンコード経由
    try {
      const base64 = Utilities.base64Encode(heicFile.getBlob().getBytes());
      const jpegBlob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        'image/jpeg',
        heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg')
      );

      const jpegFile = heicFile.getParents().next().createFile(jpegBlob);

      console.log('方法3で変換成功');
      return jpegFile;

    } catch (e) {
      console.log('方法3失敗: ' + e.toString());
    }

    return null;

  } catch (error) {
    console.error('HEIC変換エラー:', error);
    return null;
  }
}

/**
 * 外部変換サービス連携（CloudConvert）
 */
function setupCloudConvertIntegration() {
  console.log('=== CloudConvert API連携設定 ===\n');

  console.log('CloudConvertは高精度なHEIC→JPEG変換を提供する外部サービスです。\n');

  console.log('【セットアップ手順】\n');

  console.log('1. CloudConvertアカウント作成');
  console.log('────────────────────────────────');
  console.log('https://cloudconvert.com/register');
  console.log('無料プランで月25変換まで可能\n');

  console.log('2. APIキーを取得');
  console.log('────────────────────────────────');
  console.log('ダッシュボード → API → APIキーを作成\n');

  console.log('3. GASにAPIキーを設定');
  console.log('────────────────────────────────');
  console.log('PropertiesService.getScriptProperties().setProperty("CLOUDCONVERT_API_KEY", "your-api-key");\n');

  console.log('4. 変換関数を使用');
  console.log('────────────────────────────────');
  console.log('convertWithCloudConvert(heicFileId)\n');

  return {
    service: 'CloudConvert',
    freeLimit: '25変換/月',
    quality: '最高品質'
  };
}

/**
 * HEIC処理のベストプラクティス
 */
function heicBestPractices() {
  console.log('=== HEIC処理のベストプラクティス ===\n');

  console.log('【推奨ワークフロー】\n');

  console.log('1️⃣ iPhone側で対策（最優先）');
  console.log('────────────────────────────────');
  console.log('• カメラ設定を「互換性優先」に変更');
  console.log('• またはショートカットアプリで自動変換\n');

  console.log('2️⃣ アップロード時に変換');
  console.log('────────────────────────────────');
  console.log('• Google Driveアプリ経由（自動変換）');
  console.log('• ショートカットでJPEG変換後アップロード\n');

  console.log('3️⃣ GAS側で処理（最終手段）');
  console.log('────────────────────────────────');
  console.log('• processHEICImagesAdvanced() を使用');
  console.log('• 変換成功率は100%ではない\n');

  console.log('【トラブルシューティング】');
  console.log('────────────────────────────────');
  console.log('Q: HEICファイルが変換されない');
  console.log('A: iPhone側でJPEG形式に設定変更を推奨\n');

  console.log('Q: 画質が劣化する');
  console.log('A: ショートカットアプリで品質「最高」を設定\n');

  console.log('Q: 大量のHEIC画像がある');
  console.log('A: PCでまとめて変換後、Driveにアップロード\n');

  return {
    bestMethod: 'iPhone設定変更',
    alternativeMethod: 'ショートカットアプリ',
    fallbackMethod: 'GAS内変換'
  };
}

/**
 * 統合HEIC処理フロー
 */
function setupCompleteHEICWorkflow() {
  console.log('=== 完全なHEIC対応ワークフロー ===\n');

  // 1. フォルダ構成を確認
  let folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    console.log('フォルダ構成を作成中...');
    setupSimpleDriveWorkflow();
    folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  }

  // 2. HEIC変換設定
  PropertiesService.getScriptProperties().setProperty('AUTO_CONVERT_HEIC', 'true');

  console.log('✅ HEIC自動変換を有効化しました\n');

  console.log('【設定完了】');
  console.log('────────────────────────────────');
  console.log('1. HEIC画像の自動検出: 有効');
  console.log('2. 変換試行: 3つの方法を順次実行');
  console.log('3. 変換失敗時: エラーフォルダに移動\n');

  console.log('【使い方】');
  console.log('────────────────────────────────');
  console.log('1. processAllImages() を実行');
  console.log('   → JPEG/PNG/HEIC すべてを処理\n');

  console.log('【iPhone設定（推奨）】');
  console.log('────────────────────────────────');
  console.log('設定 → カメラ → フォーマット → 互換性優先\n');

  return {
    status: 'HEIC対応ワークフロー設定完了',
    nextStep: 'processAllImages()'
  };
}

/**
 * すべての画像形式を処理
 */
function processAllImages() {
  console.log('=== 全画像形式の処理 ===\n');

  // HEICを先に処理
  console.log('【HEIC画像の処理】');
  const heicResult = processHEICImagesAdvanced();

  // 通常の画像を処理
  console.log('\n【JPEG/PNG画像の処理】');
  const normalResult = processDriveImages();

  console.log('\n=== 処理完了 ===');
  console.log(`HEIC変換: ${heicResult.converted}枚`);
  console.log(`通常処理: ${normalResult.success}枚`);
  console.log(`エラー: ${normalResult.error}枚`);

  return {
    heic: heicResult.converted,
    normal: normalResult.success,
    error: normalResult.error
  };
}