const USER_SHEET_NAME = "user_backend";
const START_ROW = 4;
const DOMAIN = "@rehabstudio.online";
const RANDOM_STRING_LENGTH = 13;
const RANDOM_CHARACTERS = "!#$%&*+-./=?@_()0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";

// ====================================================================
// 1. メイン処理関数 (onChangeトリガーに設定)
// ====================================================================

/**
 * userシートのC/D列のデータ変更を検知し、ユーザー名とメールアドレス、ランダム文字列を計算し、
 * G/H/I列に「値」として書き込みます。処理後、J列の絵文字変換も実行します。
 */
function processUsernameAndEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    Logger.log("ターゲットシートが見つかりません: " + USER_SHEET_NAME);
    return;
  }

  const END_ROW = sheet.getLastRow();
  if (END_ROW < START_ROW) return;

  const dataRangeRowCount = END_ROW - START_ROW + 1;
  const D_COL = 3; // FirstName (旧D列 -> C列)
  const E_COL = 4; // LastName (旧E列 -> D列)
  const H_COL = 7; // Username (Output) (旧H列 -> G列)
  const I_COL = 8; // Email (Output) (旧I列 -> H列)
  const J_COL = 9; // Random String (Output) (旧J列 -> I列)

  // C列とD列のデータを一括で読み込み
  const dataRange = sheet.getRange(START_ROW, D_COL, dataRangeRowCount, 2);
  const data = dataRange.getValues();

  // G列からI列の既存の値を読み込み (G=インデックス0, H=1, I=2)
  const currentOutput = sheet.getRange(START_ROW, H_COL, dataRangeRowCount, 3).getValues();

  // 1. ユーザー名とメールアドレスを計算
  const calculatedResults = calculateUniqueUsername_Script(data);

  // 2. 書き込み用配列を準備 (G列, H列, I列)
  const outputArray = [];
  let changesMade = false; // G/H/I列への変更があったか

  for (let i = 0; i < data.length; i++) {
    const calculatedUsername = calculatedResults[i].username;
    const currentUsername = String(currentOutput[i][0] || "").trim(); // G列の既存値
    const currentEmail = String(currentOutput[i][1] || "").trim();    // H列の既存値
    const currentRandomString = String(currentOutput[i][2] || "").trim(); // I列の既存値
    
    let finalUsername = currentUsername;
    let finalEmail = currentEmail;
    let finalRandomString = currentRandomString;

    // A) Username (G列) の処理: 空またはエラーの場合にのみ計算結果で更新
    if (!currentUsername || currentUsername.startsWith("重複解消できません")) {
      finalUsername = calculatedUsername;
    } else {
      finalUsername = currentUsername; 
    }

    // B) Email (H列) の処理: ユーザー名が確定し、H列が空の場合のみ生成
    if (finalUsername && !currentEmail && !finalUsername.startsWith("重複解消できません")) {
      finalEmail = finalUsername + DOMAIN;
      changesMade = true;
    } else {
        finalEmail = currentEmail;
    }
    
    // C) Random String (I列) の処理: C/D列にデータがあり、I列が空の場合のみ生成
    const firstName = String(data[i][0] || "").trim();
    const lastName = String(data[i][1] || "").trim();
    
    if ((firstName || lastName) && !currentRandomString) {
        finalRandomString = generateRandomString(RANDOM_STRING_LENGTH, RANDOM_CHARACTERS);
        changesMade = true;
    } else {
        finalRandomString = currentRandomString;
    }

    // G列に変更があったかチェック
    if (finalUsername !== currentUsername) {
        changesMade = true;
    }
    
    outputArray.push([finalUsername, finalEmail, finalRandomString]);
  }
  
  // 3. 変更があった場合のみ、G列からI列に一括で値として書き込み
  if (changesMade) {
    sheet.getRange(START_ROW, H_COL, dataRangeRowCount, 3).setValues(outputArray);
  }

  // 4. J列の絵文字変換も実行
  processEmojiConversion(sheet, START_ROW);
}

// ====================================================================
// 2. スクリプト内でのユーザー名生成ロジック (変更なし)
// ====================================================================

/**
 * 配列全体を処理し、行の順序に基づいて重複しないユーザー名を計算します。
 */
function calculateUniqueUsername_Script(inputData) {
    const results = [];
    
    for (let i = 0; i < inputData.length; i++) {
        const currentFirstName = String(inputData[i][0] || "").trim().toLowerCase();
        const currentLastName = String(inputData[i][1] || "").trim().toLowerCase();
        
        if (!currentFirstName && !currentLastName) {
            results.push({username: '', email: ''});
            continue;
        }

        const generatedNames = new Set();
        let isFirstOccurrenceOfLastName = true;
        
        for (let j = 0; j < i; j++) {
            const ln = String(inputData[j][1] || "").trim().toLowerCase();
            
            if (ln !== currentLastName) continue;
            
            isFirstOccurrenceOfLastName = false;

            const fn_j = String(inputData[j][0] || "").trim().toLowerCase();
            let simulatedName = ln;
            let k_sim = 0;

            while (generatedNames.has(simulatedName) && simulatedName !== "") {
                k_sim++;
                if (k_sim > fn_j.length) {
                    simulatedName = "重複解消できません: " + ln;
                    break;
                }
                const prefix = fn_j.substring(0, k_sim);
                simulatedName = prefix + "." + ln;
            }
            generatedNames.add(simulatedName);
        }

        let username;

        if (isFirstOccurrenceOfLastName) {
            username = currentLastName;
        } else {
            let k = 0;
            let generatedNameAttempt = currentLastName;
            
            while (generatedNames.has(generatedNameAttempt) && generatedNameAttempt !== "") {
                k++;
                
                if (k > currentFirstName.length) {
                    generatedNameAttempt = "重複解消できません: " + currentLastName;
                    break;
                }
                
                const prefix = currentFirstName.substring(0, k);
                generatedNameAttempt = prefix + "." + currentLastName;
            }
            username = generatedNameAttempt;
        }
        
        const email = username && !username.startsWith("重複解消できません") ? username + DOMAIN : '';
        
        results.push({username: username, email: email});
    }
    
    return results;
}

// ====================================================================
// 3. I列ランダム文字列生成ロジック (変更なし)
// ====================================================================

/**
 * 指定された長さと文字セットでランダムな文字列を生成します。
 * @param {number} length 生成する文字列の長さ (13)。
 * @param {string} chars 使用する文字セット。
 * @return {string} ランダムな文字列。
 */
function generateRandomString(length, chars) {
  let result = '';
  for (let i = length; i > 0; --i) result += chars[Math.floor(Math.random() * chars.length)];
  return result;
}

// ====================================================================
// 4. J列 絵文字変換ロジック (列インデックス修正)
// ====================================================================

/**
 * J列のショートコードを絵文字に変換する処理を実行します。
 */
function processEmojiConversion(sheet, START_ROW) {
  const EMOJI_TARGET_COL = 10; // K列 -> J列 (11 -> 10)
  const END_ROW = sheet.getLastRow();
  
  if (END_ROW < START_ROW) return;

  const dataRangeRowCount = END_ROW - START_ROW + 1;

  const emojiMap = {
    ':white_check_mark:': '✅️', 
    ':+1:': '👍', 
  };
  
  // J列のデータをすべて読み込み
  const kColumnRange = sheet.getRange(START_ROW, EMOJI_TARGET_COL, dataRangeRowCount, 1);
  const kColumnValues = kColumnRange.getValues();
  
  const newValues = kColumnValues.map(row => [row[0]]); 
  let changesMade = false;
  
  for (let i = 0; i < kColumnValues.length; i++) {
    let currentValue = kColumnValues[i][0];
    
    if (typeof currentValue === 'string' && currentValue) {
      let tempValue = currentValue;
      let rowChanged = false;
      
      for (const shortcode in emojiMap) {
        if (tempValue.includes(shortcode)) {
          const regex = new RegExp(escapeRegExp(shortcode), 'g');
          tempValue = tempValue.replace(regex, emojiMap[shortcode]);
          rowChanged = true;
        }
      }
      
      if (rowChanged) {
        newValues[i][0] = tempValue;
        changesMade = true;
      }
    }
  }
  
  if (changesMade) {
    kColumnRange.setValues(newValues);
  }
}

/**
 * 正規表現で特殊文字として扱われる文字列をエスケープするヘルパー関数
 */
function escapeRegExp(string) {
  if (typeof string !== 'string') return string;
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}