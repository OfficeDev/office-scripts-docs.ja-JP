---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Excel on the webのOfficeスクリプトから組み込みの JavaScript API を呼び出す方法
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545048"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Officeスクリプトで組み込みの JavaScript オブジェクトを使用する

JavaScript には、スクリプトを JavaScript または[TypeScript](../overview/code-editor-environment.md) (JavaScript のスーパーセット) に含めるかどうかに関係なく、Officeスクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。 この記事では、Excel on the web用のスクリプトの組み込み JavaScript オブジェクトOffice使用する方法について説明します。

> [!NOTE]
> 組み込み JavaScript オブジェクトの完全なリストについては、Mozilla の [標準組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) の記事を参照してください。

## <a name="array"></a>配列

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を扱う標準化された方法を提供します。 配列は標準の JavaScript コンストラクトですが、範囲とコレクションの 2 つの主要な方法でスクリプトOfficeに関連しています。

### <a name="work-with-ranges"></a>範囲の操作

範囲には、その範囲内のセルに直接マップされる 2 次元配列が含まれます。 これらの配列には、その範囲内の各セルに関する特定の情報が含まれます。 たとえば、 `Range.getValues` これらのセルのすべての値を返します (2 次元配列の行と列は、そのワークシートのサブセクションの行と列にマッピングされます)。 `Range.getFormulas` と `Range.getNumberFormats` 同様に配列を返す他の頻繁に使用されるメソッドです `Range.getValues` 。

次のスクリプトは **、A1:D4** の範囲で"$" を含む任意の数値形式を検索します。 スクリプトは、それらのセルの塗りつぶしの色を "黄色" に設定します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="work-with-collections"></a>コレクションの操作

コレクションには、Excelオブジェクトの多くが含まれています。 コレクションは、Officeスクリプト API によって管理され、配列として公開されます。 たとえば、ワークシート内のすべての [図形](/javascript/api/office-scripts/excelscript/excelscript.shape) は、 `Shape[]` メソッドによって返されるに含 `Worksheet.getShapes` まれています。 この配列を使用してコレクションから値を読み取ったり、親オブジェクトのメソッドから特定のオブジェクトにアクセスしたりできます `get*` 。

> [!NOTE]
> これらのコレクション配列からオブジェクトを手動で追加したり削除したりしないでください。 `add`親オブジェクトのメソッドと `delete` 、コレクション型オブジェクトのメソッドを使用します。 たとえば、メソッドを使用して[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)に[テーブル](/javascript/api/office-scripts/excelscript/excelscript.table)を追加 `Worksheet.addTable` し `Table` 、using を削除 `Table.delete` します。

次のスクリプトは、現在のワークシート内のすべての図形の種類をログに記録します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

次のスクリプトは、現在のワークシート内の最も古い図形を削除します。

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>日付

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトは、スクリプト内の日付を処理するための標準化された方法を提供します。 `Date.now()` は、スクリプトのデータエントリにタイムスタンプを追加する場合に便利な、現在の日付と時刻を持つオブジェクトを生成します。

次のスクリプトは、ワークシートに現在の日付を追加します。 このメソッドを使用すると `toLocaleDateString` 、Excel値が日付として認識され、セルの数値書式が自動的に変更されます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

サンプルの [[日付の処理]](../resources/samples/excel-samples.md#dates) セクションには、日付に関連するスクリプトが追加されています。

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトは、一般的な数学演算のメソッドと定数を提供します。 これらは、ブックの計算エンジンを使用する必要なく、Excelでも多くの機能を提供します。 これにより、スクリプトがワークブックに対してクエリを実行する必要が生じなくなることが省かれ、パフォーマンスが向上します。

次のスクリプトは `Math.min` **、A1:D4** 範囲内の最小の数値を検索してログに記録するために使用します。 このサンプルでは、範囲全体に数値のみが含まれていることを前提とし、文字列は含まれません。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>外部 JavaScript ライブラリの使用はサポートされていません

Officeスクリプトは、外部のサードパーティ製ライブラリの使用をサポートしていません。 スクリプトは、組み込みの JavaScript オブジェクトとOfficeスクリプト API のみを使用できます。

## <a name="see-also"></a>関連項目

- [標準組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Officeスクリプト コード エディタ環境](../overview/code-editor-environment.md)
