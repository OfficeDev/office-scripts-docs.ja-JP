---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: 組み込みの JavaScript API を、Officeスクリプトから呼び出Excel on the web。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545048"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>スクリプトで組み込みの JavaScript オブジェクトをOfficeする

JavaScript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (JavaScript のスーパーセット) でスクリプトを実行するかどうかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが提供されています。 この記事では、スクリプトで組み込みの JavaScript オブジェクトの一部を使用Office説明Excel on the web。

> [!NOTE]
> すべての組み込み JavaScript オブジェクトの完全な一覧については、Mozilla の Standard 組み込 [みオブジェクトの記事を参照](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) してください。

## <a name="array"></a>配列

[Array オブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)は、スクリプト内の配列を扱う標準化された方法を提供します。 配列は標準的な JavaScript コンストラクトですが、Officeとコレクションの 2 つの主要な方法で関連付けされます。

### <a name="work-with-ranges"></a>範囲の使用

範囲には、その範囲内のセルに直接マップする複数の 2 次元配列が含まれます。 これらの配列には、その範囲内の各セルに関する特定の情報が含まれます。 たとえば、これらのセル内のすべての値を返します (2 次元配列の行と列は、そのワークシート のサブセクションの行と列 `Range.getValues` にマッピングされます)。 `Range.getFormulas` などの `Range.getNumberFormats` 配列を返す他の頻繁に使用されるメソッドです `Range.getValues` 。

次のスクリプトは **、A1:D4 範囲で** "$" を含む任意の数値形式を検索します。 スクリプトは、これらのセルの塗りつぶしの色を "黄色" に設定します。

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

### <a name="work-with-collections"></a>コレクションの使用

多Excelオブジェクトはコレクションに含まれています。 このコレクションは、スクリプト API Officeによって管理され、配列として公開されます。 たとえば、ワークシート [内のすべての Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) は、メソッドによって返されるオブジェクト `Shape[]` に含 `Worksheet.getShapes` まれているとします。 この配列を使用して、コレクションから値を読み取ることができます。また、親オブジェクトのメソッドから特定のオブジェクトに `get*` アクセスすることもできます。

> [!NOTE]
> これらのコレクション配列からオブジェクトを手動で追加または削除しない。 親オブジェクト `add` のメソッドとコレクション型オブジェクト `delete` のメソッドを使用します。 たとえば、メソッドを使用して[ワークシートに Table](/javascript/api/office-scripts/excelscript/excelscript.table)を追加し[](/javascript/api/office-scripts/excelscript/excelscript.worksheet) `Worksheet.addTable` 、using を削除 `Table` します `Table.delete` 。

次のスクリプトは、現在のワークシートのすべての図形の種類をログに記録します。

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

次のスクリプトは、現在のワークシートで最も古い図形を削除します。

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

[Date オブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)は、スクリプト内の日付を扱う標準化された方法を提供します。 `Date.now()` 現在の日付と時刻を持つオブジェクトを生成します。これは、スクリプトのデータ エントリにタイムスタンプを追加するときに便利です。

次のスクリプトは、現在の日付をワークシートに追加します。 メソッドを使用すると、Excel日付として認識され、セルの数値形式 `toLocaleDateString` が自動的に変更されます。

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

サンプル [の [日付の処理](../resources/samples/excel-samples.md#dates) ] セクションには、より多くの日付関連のスクリプトがあります。

## <a name="math"></a>数学

[Math オブジェクトは](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)、一般的な数学演算のメソッドと定数を提供します。 ブックの計算エンジンを使用せずに、Excelで使用できる多くの機能を提供します。 これにより、スクリプトがブックにクエリを実行する必要が省き、パフォーマンスが向上します。

次のスクリプトは `Math.min` **、A1:D4** 範囲の最小番号を検索してログに記録するために使用します。 このサンプルでは、文字列ではなく、範囲全体に数値だけが含まれていると想定しています。

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

Officeスクリプトは、外部のサード パーティ製ライブラリの使用をサポートしています。 スクリプトでは、組み込みの JavaScript オブジェクトとスクリプト API Office使用できます。

## <a name="see-also"></a>関連項目

- [標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Officeスクリプト コード エディター環境](../overview/code-editor-environment.md)
