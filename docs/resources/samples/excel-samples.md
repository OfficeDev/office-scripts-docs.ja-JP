---
title: Excel on the webでのスクリプトのOfficeの基本スクリプト
description: Excel on the webのスクリプトで使用するコード サンプルのコレクションOffice。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545760"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Excel on the webでのスクリプトのOfficeの基本スクリプト

以下のサンプルは、独自のワークブックを試用するための簡単なスクリプトです。 Excel on the webでそれらを使用するには:

1. **[自動化]** タブを開きます。
2. **コード エディタ を** 押します。
3. [コード エディタ] 作業ウィンドウで **[新しいスクリプト** ] をクリックします。
4. スクリプト全体を、選択したサンプルに置き換えます。
5. [コード エディタ] 作業ウィンドウで [ **実行** ] をクリックします。

## <a name="script-basics"></a>スクリプトの基本

これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。 これらのスクリプトをスクリプトに追加して、ソリューションを拡張し、一般的な問題を解決します。

### <a name="read-and-log-one-cell"></a>1 つのセルの読み取りとログ記録

このサンプルでは **、A1** の値を読み取り、コンソールに出力します。

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

### <a name="read-the-active-cell"></a>アクティブ セルを読み取る

このスクリプトは、現在アクティブなセルの値をログに記録します。 複数のセルが選択されている場合は、一番上のセルがログに記録されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>隣接するセルを変更する

このスクリプトは、相対参照を使用して隣接するセルを取得します。 アクティブ セルが一番上の行にある場合、スクリプトの一部は、現在選択されているセルの上のセルを参照するため、失敗します。

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

### <a name="change-all-adjacent-cells"></a>隣接するすべてのセルを変更する

このスクリプトは、アクティブ セルの書式を隣接するセルにコピーします。 このスクリプトは、アクティブ セルがワークシートの端にない場合にのみ機能します。

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

このスクリプトは、現在選択されている範囲をループします。 現在の書式をクリアし、各セルの塗りつぶしの色をランダムな色に設定します。

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a>特殊な条件に基づいてセルのグループを取得する

このスクリプトは、現在のワークシートで使用されている範囲内のすべての空白セルを取得します。 次に、黄色の背景を持つすべてのセルを強調表示します。

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

これらのサンプルは、ブック内のオブジェクトのコレクションを扱います。

### <a name="iterate-over-collections"></a>コレクションを反復処理する

このスクリプトは、ブック内のすべてのワークシートの名前を取得し、ログに記録します。 タブの色もランダムな色に設定されます。

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

このセクションのサンプルでは、JavaScript [日付](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) オブジェクトの使用方法を示します。

次のサンプルでは、現在の日付と時刻を取得し、アクティブ ワークシート内の 2 つのセルにこれらの値を書き込みます。

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

次のサンプルでは、Excelに格納されている日付を読み取り、JavaScript Date オブジェクトに変換します。 JavaScript [日付の入力として日付の数値シリアル番号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) を使用します。

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

これらのサンプルでは、ワークシート データを操作する方法を示し、ユーザーにより良いビューや組織を提供します。

### <a name="apply-conditional-formatting"></a>条件付き書式の適用

このサンプルでは、ワークシートで現在使用されている範囲に条件付き書式を適用します。 条件付き書式は、値の上位 10% に対して緑の塗りつぶしです。

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

### <a name="create-a-sorted-table"></a>並べ替えられたテーブルを作成する

このサンプルでは、現在のワークシートで使用されている範囲からテーブルを作成し、最初の列に基づいて並べ替えます。

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>ピボットテーブルから "総計" 値をログに記録する

このサンプルでは、ブック内の最初のピボットテーブルを検索し、値を "総計" セルに記録します (下の図では緑色で強調表示されています)。

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="大合計行が緑色にハイライト表示された果物の売り上げを示すピボットテーブル":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a>データの入力規則を使用してドロップダウン リストを作成する

このスクリプトは、セルのドロップダウン選択リストを作成します。 選択した範囲の既存の値をリストの選択肢として使用します。

:::image type="content" source="../../images/sample-data-validation.png" alt-text="色の選択肢が 「赤、青、緑」を含む 3 つのセルの範囲を示すワークシートと、その横に表示されるドロップダウン リストと同じ選択肢":::

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

これらのサンプルでは、Excel式を使用し、スクリプトでの操作方法を示します。

### <a name="single-formula"></a>単一の式

このスクリプトは、セルの数式を設定し、次に、Excelがセルの数式と値を別々に格納する方法を表示します。

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>`#SPILL!`数式から返されたエラーを処理する

このスクリプトは、TRANSPOSE 関数を使用して、範囲 "A1:D2" を "A4:B7" に変換します。 転置がエラーの原因となった `#SPILL` 場合は、対象範囲をクリアし、数式を再度適用します。

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

新しいサンプルの提案を歓迎します。 他のスクリプト開発者に役立つ一般的なシナリオがある場合は、ページの下部にあるフィードバックセクションで教えてください。

## <a name="see-also"></a>関連項目

* [スーディ・ラマムルティのYouTubeでの「レンジの基本」](https://youtu.be/4emjkOFdLBA)
* [Officeスクリプトのサンプルとシナリオ](samples-overview.md)
* [Excel on the web で Office スクリプトを記録、編集、作成する](../../tutorials/excel-tutorial.md)
