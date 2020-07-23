---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Web 上の Excel で Office スクリプトから組み込みの JavaScript Api を呼び出す方法について説明します。
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229597"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Office スクリプトでの組み込みの JavaScript オブジェクトの使用

Javascript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (javascript のスーパーセット) のどちらでスクリプトを作成するかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。 この記事では、web 上の Excel の Office スクリプトに組み込まれている JavaScript オブジェクトのいくつかを使用する方法について説明します。

> [!NOTE]
> すべての組み込み JavaScript オブジェクトの完全なリストについては、「Mozilla の[標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)」記事を参照してください。

## <a name="array"></a>配列

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を操作するための標準化された方法を提供します。 配列は標準的な JavaScript コンストラクトですが、範囲とコレクションという2つの主な方法で Office スクリプトに関連しています。

### <a name="working-with-ranges"></a>範囲を使用して作業する

範囲には、その範囲内のセルに直接マップされるいくつかの2次元配列が含まれています。 これらの配列には、その範囲内の各セルに関する特定の情報が含まれています。 たとえば、は、 `Range.getValues` 2 次元配列の行と列がそのワークシートサブセクションの行と列にマッピングされているセルのすべての値を返します。 `Range.getFormulas`また、 `Range.getNumberFormats` のように配列を返すその他のメソッドもよく使用され `Range.getValues` ます。

次のスクリプトは、"$" が含まれている任意の番号書式の**A1: D4**範囲を検索します。 このスクリプトは、これらのセルの塗りつぶしの色を "黄" に設定します。

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

### <a name="working-with-collections"></a>コレクションを処理する

多くの Excel オブジェクトがコレクションに含まれています。 コレクションは Office スクリプト API によって管理され、配列として公開されます。 たとえば、ワークシート内のすべての[図形](/javascript/api/office-scripts/excelscript/excelscript.shape)は、 `Shape[]` メソッドによって返されるに含まれてい `Worksheet.getShapes` ます。 この配列を使用して、コレクションから値を取得したり、親オブジェクトのメソッドから特定のオブジェクトにアクセスしたりでき `get*` ます。

> [!NOTE]
> これらのコレクションの配列に対してオブジェクトを手動で追加または削除しないでください。 `add`親オブジェクトのメソッドと、コレクション型のオブジェクトのメソッドを使用し `delete` ます。 たとえば、メソッドを使用して[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)に[テーブル](/javascript/api/office-scripts/excelscript/excelscript.table)を追加 `Worksheet.addTable` し、 `Table` using を削除し `Table.delete` ます。

次のスクリプトは、現在のワークシート内のすべての図形の種類を記録します。

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

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトには、スクリプト内の日付を処理するための標準化された方法が用意されています。 `Date.now()`現在の日付と時刻を使用してオブジェクトを生成します。これは、スクリプトのデータ入力にタイムスタンプを追加するときに便利です。

次のスクリプトは、現在の日付をワークシートに追加します。 このメソッドを使用する `toLocaleDateString` と、Excel によって値が日付として認識され、セルの数値の書式が自動的に変更されることに注意してください。

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

サンプルの [[日付を使用して作業](../resources/excel-samples.md#dates)] セクションには、さらに多くの日付関連スクリプトがあります。

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトには、一般的な数値演算のためのメソッドと定数が用意されています。 これらは、ブックの計算エンジンを使用しなくても、Excel で使用できる多くの関数を提供します。 これにより、スクリプトはブックを照会する必要がなくなり、パフォーマンスが向上します。

次のスクリプトは、を使用して、 `Math.min` **A1: D4**範囲の最小数を検索して記録します。 この例では、範囲全体に文字列ではなく数値のみが含まれていることに注意してください。

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

Office スクリプトは、外部のサードパーティ製ライブラリの使用をサポートしていません。 スクリプトでは、組み込みの JavaScript オブジェクトと Office スクリプト Api のみを使用できます。

## <a name="see-also"></a>関連項目

- [標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office スクリプトのコードエディター環境](../overview/code-editor-environment.md)
