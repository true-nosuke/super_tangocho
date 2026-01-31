// ========================================
// 多機能単語帳アプリ - Google Apps Script
// ========================================

// ========================================
// 定数定義
// ========================================
const SHEET_NAME = {
  WORD_DATA: '単語データ',
  CHALLENGE: '４択チャレンジ'
};

const COLUMN_INDEX = {
  IS_WEAK: 1,        // A列: 苦手かどうか
  ENGLISH: 2,        // B列: 英語
  JAPANESE: 3,       // C列: 日本語訳
  EXAMPLE: 4,        // D列: 例文
  EXAMPLE_JP: 5,     // E列: 例文（日本語訳）
  CORRECT: 6,        // F列: 正解回数
  INCORRECT: 7,      // G列: 不正解回数
  LAST_STUDY: 8      // H列: 最終学習日
};

// Groq API設定（スクリプトプロパティに保存してください）
const GROQ_API_KEY = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
const GROQ_API_URL = 'https://api.groq.com/openai/v1/chat/completions';

// ========================================
// メニュー作成
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('単語帳アプリ')
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('選択範囲を自動翻訳', 'translateSelectedCells')
    .addItem('選択範囲の例文を生成', 'generateExamplesForSelected')
    .addSeparator()
    .addItem('４択チャレンジを開始', 'startChallenge')
    .addSeparator()
    .addSubMenu(ui.createMenu('設定')
      .addItem('Groq APIキーを設定', 'setGroqApiKey'))
    .addToUi();
}

// ========================================
// サイドバー表示
// ========================================
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('単語帳コントロール')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ========================================
// 翻訳機能
// ========================================
/**
 * 選択中のセルの日本語を英語に翻訳
 */
function translateSelectedCells() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  
  const translatedValues = values.map(row => {
    return row.map(cell => {
      if (cell && typeof cell === 'string' && cell.trim() !== '') {
        try {
          return LanguageApp.translate(cell, 'ja', 'en');
        } catch (e) {
          Logger.log('翻訳エラー: ' + e);
          return cell;
        }
      }
      return cell;
    });
  });
  
  range.setValues(translatedValues);
  SpreadsheetApp.getActiveSpreadsheet().toast('翻訳が完了しました', '成功', 3);
}

/**
 * シート全体の空白の日本語訳セルに翻訳を挿入
 */
function translateEmptyJapaneseCells(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const englishValues = sheet.getRange(2, COLUMN_INDEX.ENGLISH, lastRow - 1, 1).getValues();
  const japaneseValues = sheet.getRange(2, COLUMN_INDEX.JAPANESE, lastRow - 1, 1).getValues();
  
  const translatedValues = [];
  
  for (let i = 0; i < englishValues.length; i++) {
    const english = englishValues[i][0];
    const japanese = japaneseValues[i][0];
    
    if (english && english.toString().trim() !== '' && (!japanese || japanese.toString().trim() === '')) {
      try {
        const translated = LanguageApp.translate(english, 'en', 'ja');
        translatedValues.push([translated]);
      } catch (e) {
        Logger.log(`翻訳エラー (行 ${i + 2}): ${e}`);
        translatedValues.push(['']);
      }
    } else {
      translatedValues.push([japanese || '']);
    }
  }
  
  if (translatedValues.length > 0) {
    sheet.getRange(2, COLUMN_INDEX.JAPANESE, translatedValues.length, 1).setValues(translatedValues);
  }
  
  return translatedValues.filter(v => v[0] !== '').length;
}

// ========================================
// Groq API による例文生成
// ========================================
/**
 * Groq APIキーを設定
 */
function setGroqApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Groq APIキー設定',
    'Groq APIキーを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const apiKey = result.getResponseText();
    PropertiesService.getScriptProperties().setProperty('GROQ_API_KEY', apiKey);
    ui.alert('APIキーが保存されました');
  }
}

/**
 * Groq APIを使用して例文を生成
 */
function generateExampleSentence(word, meaning) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  
  if (!apiKey) {
    throw new Error('Groq APIキーが設定されていません。メニューから設定してください。');
  }
  
  const prompt = `Create a simple and natural example sentence using the English word "${word}" (meaning: ${meaning}). 
Only output the example sentence in English, nothing else.`;
  
  const payload = {
    model: 'llama3-8b-8192',
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.7,
    max_tokens: 100
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(GROQ_API_URL, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.choices && json.choices.length > 0) {
      return json.choices[0].message.content.trim();
    }
    throw new Error('APIレスポンスが不正です');
  } catch (e) {
    Logger.log('Groq APIエラー: ' + e);
    throw e;
  }
}

/**
 * 選択範囲の単語に例文を生成
 */
function generateExamplesForSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('データがありません');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '例文生成',
    `選択中のシートの全ての単語に例文を生成しますか？\n(例文が空白のセルのみ)`,
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const data = sheet.getRange(2, COLUMN_INDEX.ENGLISH, lastRow - 1, 5).getValues();
  let generatedCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    const english = data[i][0];
    const japanese = data[i][1];
    const example = data[i][2];
    
    if (english && english.toString().trim() !== '' && (!example || example.toString().trim() === '')) {
      try {
        const exampleSentence = generateExampleSentence(english, japanese || '');
        const exampleJapanese = LanguageApp.translate(exampleSentence, 'en', 'ja');
        
        sheet.getRange(i + 2, COLUMN_INDEX.EXAMPLE).setValue(exampleSentence);
        sheet.getRange(i + 2, COLUMN_INDEX.EXAMPLE_JP).setValue(exampleJapanese);
        generatedCount++;
        
        Utilities.sleep(500); // API制限対策
      } catch (e) {
        Logger.log(`例文生成エラー (行 ${i + 2}): ${e}`);
      }
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`${generatedCount}件の例文を生成しました`, '完了', 3);
}

// ========================================
// 読み上げ機能（サイドバーから制御）
// ========================================
/**
 * 読み上げ用のデータを取得
 */
function getWordsForReading(scope) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  
  if (scope === 'selection') {
    sheet = ss.getActiveSheet();
    const range = sheet.getActiveRange();
    const row = range.getRow();
    const data = sheet.getRange(row, COLUMN_INDEX.ENGLISH, 1, 5).getValues()[0];
    
    return [{
      english: data[0] || '',
      japanese: data[1] || '',
      example: data[2] || '',
      exampleJp: data[3] || ''
    }];
  } else if (scope === 'sheet') {
    sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, COLUMN_INDEX.ENGLISH, lastRow - 1, 5).getValues();
    return data.map(row => ({
      english: row[0] || '',
      japanese: row[1] || '',
      example: row[2] || '',
      exampleJp: row[3] || ''
    })).filter(word => word.english);
  }
  
  return [];
}

// ========================================
// ４択チャレンジ機能
// ========================================
/**
 * チャレンジを開始
 */
function startChallenge() {
  const html = HtmlService.createHtmlOutputFromFile('Challenge')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '４択チャレンジ');
}

/**
 * チャレンジ用の問題データを取得
 */
function getChallengeData(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allWords = [];
  
  // データ収集
  if (config.scope === 'all') {
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      if (sheet.getName() !== SHEET_NAME.CHALLENGE) {
        allWords = allWords.concat(getWordsFromSheet(sheet));
      }
    });
  } else if (config.scope === 'current') {
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME.CHALLENGE) {
      allWords = getWordsFromSheet(sheet);
    }
  }
  
  // フィルタリング
  let filteredWords = allWords.filter(word => word.english && word.japanese);
  
  // 苦手な単語を優先
  if (config.prioritizeWeak) {
    filteredWords.sort((a, b) => {
      const aWeak = a.isWeak || (a.incorrect > a.correct);
      const bWeak = b.isWeak || (b.incorrect > b.correct);
      if (aWeak && !bWeak) return -1;
      if (!aWeak && bWeak) return 1;
      return 0;
    });
  }
  
  // 忘却曲線による優先
  if (config.prioritizeOld) {
    filteredWords.sort((a, b) => {
      const aTime = a.lastStudy ? new Date(a.lastStudy).getTime() : 0;
      const bTime = b.lastStudy ? new Date(b.lastStudy).getTime() : 0;
      return aTime - bTime;
    });
  }
  
  // 問題数に応じて切り出し
  const questions = filteredWords.slice(0, Math.min(config.questionCount, filteredWords.length));
  
  // 問題タイプに応じた選択肢生成
  return questions.map(word => {
    const question = generateQuestion(word, config.questionType, allWords);
    return question;
  });
}

/**
 * シートから単語データを取得
 */
function getWordsFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  return data.map((row, index) => ({
    sheetName: sheet.getName(),
    rowIndex: index + 2,
    isWeak: row[0] === true || row[0] === 'TRUE' || row[0] === '○',
    english: row[1] || '',
    japanese: row[2] || '',
    example: row[3] || '',
    exampleJp: row[4] || '',
    correct: parseInt(row[5]) || 0,
    incorrect: parseInt(row[6]) || 0,
    lastStudy: row[7] || null
  })).filter(word => word.english);
}

/**
 * 問題を生成
 */
function generateQuestion(correctWord, questionType, allWords) {
  let questionText = '';
  let correctAnswer = '';
  
  if (questionType === 'en-to-jp') {
    questionText = correctWord.english;
    correctAnswer = correctWord.japanese;
  } else if (questionType === 'jp-to-en') {
    questionText = correctWord.japanese;
    correctAnswer = correctWord.english;
  } else if (questionType === 'example') {
    if (correctWord.example) {
      // 単語を空欄にする
      questionText = correctWord.example.replace(new RegExp(correctWord.english, 'gi'), '______');
      correctAnswer = correctWord.english;
    } else {
      // 例文がない場合は英→日に切り替え
      questionText = correctWord.english;
      correctAnswer = correctWord.japanese;
    }
  }
  
  // ダミー選択肢を生成
  const choices = generateChoices(correctAnswer, allWords, questionType);
  
  return {
    questionText: questionText,
    choices: choices,
    correctAnswer: correctAnswer,
    wordData: correctWord
  };
}

/**
 * 選択肢を生成
 */
function generateChoices(correctAnswer, allWords, questionType) {
  const choices = [correctAnswer];
  const otherWords = allWords.filter(w => {
    if (questionType === 'en-to-jp' || questionType === 'example') {
      return w.japanese && w.japanese !== correctAnswer;
    } else {
      return w.english && w.english !== correctAnswer;
    }
  });
  
  // ランダムに3つ選択
  const shuffled = otherWords.sort(() => Math.random() - 0.5);
  for (let i = 0; i < Math.min(3, shuffled.length); i++) {
    if (questionType === 'en-to-jp' || questionType === 'example') {
      choices.push(shuffled[i].japanese);
    } else {
      choices.push(shuffled[i].english);
    }
  }
  
  // 選択肢が4つに満たない場合はダミーを追加
  while (choices.length < 4) {
    choices.push('(選択肢なし)');
  }
  
  // シャッフル
  return choices.sort(() => Math.random() - 0.5);
}

/**
 * 回答結果を記録
 */
function recordAnswer(wordData, isCorrect) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(wordData.sheetName);
  
  if (!sheet) return;
  
  const row = wordData.rowIndex;
  
  // 正解/不正解回数を更新
  if (isCorrect) {
    const currentCorrect = sheet.getRange(row, COLUMN_INDEX.CORRECT).getValue() || 0;
    sheet.getRange(row, COLUMN_INDEX.CORRECT).setValue(parseInt(currentCorrect) + 1);
  } else {
    const currentIncorrect = sheet.getRange(row, COLUMN_INDEX.INCORRECT).getValue() || 0;
    sheet.getRange(row, COLUMN_INDEX.INCORRECT).setValue(parseInt(currentIncorrect) + 1);
    
    // 不正解の場合は苦手マークをつける
    sheet.getRange(row, COLUMN_INDEX.IS_WEAK).setValue('○');
  }
  
  // 最終学習日時を更新
  sheet.getRange(row, COLUMN_INDEX.LAST_STUDY).setValue(new Date());
}

// ========================================
// 初期設定関数
// ========================================
/**
 * 単語データシートを初期化
 */
function initializeWordDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME.WORD_DATA);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME.WORD_DATA);
  }
  
  // ヘッダー設定
  const headers = [
    '苦手',
    '英語',
    '日本語訳',
    '例文',
    '例文（日本語訳）',
    '正解回数',
    '不正解回数',
    '最終学習日'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅を調整
  sheet.setColumnWidth(1, 60);   // 苦手
  sheet.setColumnWidth(2, 150);  // 英語
  sheet.setColumnWidth(3, 150);  // 日本語訳
  sheet.setColumnWidth(4, 300);  // 例文
  sheet.setColumnWidth(5, 300);  // 例文（日本語訳）
  sheet.setColumnWidth(6, 80);   // 正解回数
  sheet.setColumnWidth(7, 80);   // 不正解回数
  sheet.setColumnWidth(8, 150);  // 最終学習日
  
  SpreadsheetApp.getActiveSpreadsheet().toast('単語データシートを初期化しました', '完了', 3);
}

/**
 * チャレンジシートを初期化
 */
function initializeChallengeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME.CHALLENGE);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME.CHALLENGE);
  }
  
  sheet.clear();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('チャレンジシートを初期化しました', '完了', 3);
}
