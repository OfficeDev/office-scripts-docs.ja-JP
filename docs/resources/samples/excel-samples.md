---
title: Office スクリプトの基本的なExcel on the web
description: スクリプト内のスクリプトと一緒にOfficeコード サンプルのコレクションExcel on the web。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 6d74d55556feb93e0f49da375b3c7896d439663f7f922e4ae135b6fdc6a40197
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847546"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Office スクリプトの基本的なExcel on the web

次のサンプルは、独自のブックで試す簡単なスクリプトです。 次の方法で使用Excel on the web。

1. **[自動化]** タブを開きます。
1. **[新しいスクリプト]** を選択します。
1. スクリプト全体を、選択したサンプルに置き換える。
1. コード **エディターの** 作業ウィンドウで [実行] を選択します。

## <a name="script-basics"></a>スクリプトの基本

これらのサンプルでは、スクリプトの基本的な構成要素Office示します。 これらのスクリプトを展開してソリューションを拡張し、一般的な問題を解決します。

### <a name="read-and-log-one-cell"></a>1 つのセルを読み取り、ログに記録する

このサンプルでは **、A1 の値を読み** 取り、コンソールに出力します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a>アクティブ セルの読み取り

このスクリプトは、現在のアクティブ セルの値をログに記録します。 複数のセルが選択されている場合は、左上のセルがログに記録されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>隣接するセルを変更する

このスクリプトは、相対参照を使用して隣接するセルを取得します。 アクティブ セルが一番上の行にある場合、スクリプトの一部は、現在選択されているセルの上にあるセルを参照しますので、失敗します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a>隣接するセルを変更する

このスクリプトは、アクティブ セルの書式設定を隣接セルにコピーします。 このスクリプトは、アクティブ セルがワークシートの端にない場合にのみ機能します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a>範囲内の各セルを変更する

このスクリプトは、現在選択されている範囲をループ処理します。 現在の書式設定をクリアし、各セルの塗りつぶしの色をランダムな色に設定します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a>特別な条件に基づいてセルのグループを取得する

このスクリプトは、現在のワークシートの使用範囲内のすべての空白セルを取得します。 次に、これらのすべてのセルを黄色の背景で強調表示します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a>コレクション

これらのサンプルは、ブック内のオブジェクトのコレクションで動作します。

### <a name="iterate-over-collections"></a>コレクションを反復処理する

このスクリプトは、ブック内のすべてのワークシートの名前を取得してログに記録します。 また、タブの色をランダムな色に設定します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a>コレクションのクエリと削除

このスクリプトは、新しいワークシートを作成します。 ワークシートの既存のコピーをチェックし、新しいシートを作成する前に削除します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a>日付

このセクションのサンプルでは、JavaScript Date オブジェクトの使い方 [を示](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) します。

次のサンプルでは、現在の日付と時刻を取得し、それらの値をアクティブワークシートの 2 つのセルに書き込みます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

次のサンプルでは、データに格納されている日付を読みExcel JavaScript Date オブジェクトに変換します。 日付の数値 [シリアル番号を](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) JavaScript Date の入力として使用します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>データの表示

これらのサンプルでは、ワークシート データを処理し、より良いビューまたは組織をユーザーに提供する方法を示します。

### <a name="apply-conditional-formatting"></a>条件付き書式の適用

このサンプルでは、ワークシートで現在使用されている範囲に条件付き書式を適用します。 条件付き書式は、値の上位 10% の緑の塗りつぶしです。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a>並べ替えテーブルを作成する

このサンプルでは、現在のワークシートの使用範囲からテーブルを作成し、最初の列に基づいて並べ替えを行います。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>ピボットテーブルから "Grand Total" 値をログに記録する

このサンプルでは、ブック内の最初のピボットテーブルを検索し、値を "Grand Total" セルに記録します (下の図では緑色で強調表示されています)。

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="[総合計] 行が緑色で強調表示された、果物の売上を示すピボットテーブル。":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a>データ検証を使用してドロップダウン リストを作成する

このスクリプトは、セルのドロップダウン選択リストを作成します。 選択した範囲の既存の値をリストの選択肢として使用します。

:::image type="content" source="../../images/sample-data-validation.png" alt-text="色の選択肢 '赤、青、緑' を含む 3 つのセルの範囲を示すワークシートで、ドロップダウン リストに表示されるのと同じ選択肢を示します。":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a>数式

これらのサンプルでは、Excelを使用し、スクリプトで使用する方法を示します。

### <a name="single-formula"></a>単一の数式

このスクリプトは、セルの数式を設定し、セルExcel値を個別に格納する方法を表示します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a>数式から `#SPILL!` 返されるエラーを処理する

このスクリプトは、TRANSPOSE 関数を使用して範囲 "A1:D2" を "A4:B7" にトランスポーズします。 トランスポーズでエラーが発生した場合は、ターゲット範囲をクリアし `#SPILL` 、数式を再度適用します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a>新しいサンプルの提案

新しいサンプルの提案を歓迎します。 他のスクリプト開発者に役立つ一般的なシナリオがある場合は、ページの下部にあるフィードバック セクションで教えて下さい。

## <a name="see-also"></a>関連項目

* [Sudhi Ramamurthy の YouTube の "Range basics"](https://youtu.be/4emjkOFdLBA)
* [Officeスクリプトのサンプルとシナリオ](samples-overview.md)
* [Excel on the web で Office スクリプトを記録、編集、作成する](../../tutorials/excel-tutorial.md)
