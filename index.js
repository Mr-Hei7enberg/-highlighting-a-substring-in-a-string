let dictionary = {
  "stop_words": [
    "Север",
    "Мороз",
    "Друг",
    "День"
  ]
}

function find(i, subStr, str) {
  // Алгоритм Кнута-Морриса-Пратта (КМП)
  // i-с какого места строки  ищем
  // j-с какого места образца ищем
  for (i; str[i]; ++i) {
    for (j = 0; ; ++j) {
      if (!subStr[j]) return i; // образец найден 
      if (str[i + j] != subStr[j]) break;
    }
    // пока не найден, продолжим внешний цикл
  }
  // образца нет
  return -1;
}

function searchIndex(word, text) {
  let res = [];                    
  let startIndex = 0;

  while (find(startIndex, word, text) !== -1) {
    let poisk = find(startIndex, word, text);
    let endIndex = poisk + word.length;

    if (res.length === 0) {
      res.push({
      startIndex : poisk,
      endIndex : endIndex
      });
    } 
    if (res[res.length - 1].startIndex !== poisk) {
      res.push({
      startIndex : poisk,
      endIndex : endIndex
      });
    }
    startIndex++
  }
 return res;
}

function substringSearch() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // range диапазон поиска
  let range = sheet.getRange("A1:G7").getValues();

  range.forEach((item, row) => {
    for (let column = 0; column < item.length; column++) {
      if (item[column] != "") {
          let color = SpreadsheetApp.newTextStyle().setForegroundColor("red").build();
          let richText = SpreadsheetApp.newRichTextValue().setText(item[column]);
        // item[column] текст ячейки
        dictionary.stop_words.forEach(word => {
          let arr = searchIndex(word.toLocaleLowerCase(), item[column].toLocaleLowerCase());
            arr.forEach(ind => {
                richText.setTextStyle(ind.startIndex, ind.endIndex, color);
            })
             sheet.getRange(row + 1, column + 1).setRichTextValue(richText.build()); 
        })
      }
    }
  })
}
