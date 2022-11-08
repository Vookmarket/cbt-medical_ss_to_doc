let ss = SpreadsheetApp.getActiveSpreadsheet();
let ssQuestion = ss.getSheetByName('設問')
let ssCat = ss.getSheetByName('カテゴリー')
function loadSheet() {
  //最終行を取得する
  let ssQuestionLastRow = ssQuestion.getLastRow();

  //カテゴリーシートから基本情報を取得する
  let author = ssCat.getRange(1, 2).getValue();
  let date = new Date(ssCat.getRange(2, 2).getValue());
  let dateStr = Utilities.formatDate(date, 'JST', 'YYYYMMdd')
  let category = ssCat.getRange(3, 2).getValue();
  let majorItem = ssCat.getRange(4, 2).getValue();
  let mediumItem = ssCat.getRange(5, 2).getValue();
  let minorItem = ssCat.getRange(6, 2).getValue();

  //ドキュメントを作成する
  let name = author + '_' + dateStr;
  let doc = DocumentApp.create(name)
  let body = doc.getBody();

  for (i = 2; i <= ssQuestionLastRow; i++) {
    //設問シートから設問を取得する
    let question = ssQuestion.getRange(i, 2).getValue()
    Logger.log(question);
    //選択肢をリストで取得する
    selectLs = [
      ssQuestion.getRange(i, 3).getValue(),
      ssQuestion.getRange(i, 4).getValue(),
      ssQuestion.getRange(i, 5).getValue(),
      ssQuestion.getRange(i, 6).getValue(),
      ssQuestion.getRange(i, 7).getValue()
    ]
    Logger.log(selectLs);
    //正解選択肢番号を取得する
    let answer = ssQuestion.getRange(i, 8).getValue();

    //選択肢の出力文を作成する
    let selectParacraphs = []
    for (j = 0; j <= 4; j++) {
      if (j != answer - 1) {
        let selectText = '!#選択肢　' + selectLs[j]
        selectParacraphs.push(selectText);
      } else {
        let selectText = '!#正解　' + selectLs[j]
        selectParacraphs.push(selectText);
      }
    }
    arrayShuffle(selectParacraphs);

    //本文を出力する
    body.appendParagraph('!#開始');
    j = i - 1
    body.appendParagraph('!#問題コード ' + dateStr + '_' + j + '_' + author);
    body.appendParagraph('!#設問　' + question)
    //選択肢を出力する
    body.appendParagraph(selectParacraphs[0])
    body.appendParagraph(selectParacraphs[1])
    body.appendParagraph(selectParacraphs[2])
    body.appendParagraph(selectParacraphs[3])
    body.appendParagraph(selectParacraphs[4])
    body.appendParagraph('!#カテゴリ　' + category);
    body.appendParagraph('!#タグ　大項目①:' + majorItem);
    body.appendParagraph('!#タグ　中項目①:' + mediumItem);
    body.appendParagraph('!#タグ　小項目①:' + minorItem);
    body.appendParagraph('\n');
  }
}

function arrayShuffle(array) {
  for(var i = (array.length - 1); 0 < i; i--){

    // 0〜(i+1)の範囲で値を取得
    var r = Math.floor(Math.random() * (i + 1));

    // 要素の並び替えを実行
    var tmp = array[i];
    array[i] = array[r];
    array[r] = tmp;
  }
  return array;
}
