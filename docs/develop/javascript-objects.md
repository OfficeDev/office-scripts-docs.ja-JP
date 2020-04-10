---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Web 上の Excel で Office スクリプトから組み込みの JavaScript Api を呼び出す方法について説明します。
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: 54cadb6e9ce60e631488bbe7de00c29a6db35eb7
ms.sourcegitcommit: b13dedb5ee2048f0a244aa2294bf2c38697cb62c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215260"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Office スクリプトでの組み込みの JavaScript オブジェクトの使用

Javascript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (javascript のスーパーセット) のどちらでスクリプトを作成するかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。 この記事では、web 上の Excel の Office スクリプトに組み込まれている JavaScript オブジェクトのいくつかを使用する方法について説明します。

> [!NOTE]
> すべての組み込み JavaScript オブジェクトの完全なリストについては、「Mozilla の[標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)」記事を参照してください。

## <a name="array"></a>配列

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を操作するための標準化された方法を提供します。 配列は標準的な JavaScript コンストラクトですが、範囲とコレクションという2つの主な方法で Office スクリプトに関連しています。

### <a name="working-with-ranges"></a>範囲を使用して作業する

範囲には、その範囲内のセルに直接マップされるいくつかの2次元配列が含まれています。 これら`values`には、、 `formulas`、などの`numberFormat`プロパティが含まれます。 配列型のプロパティは、他のプロパティと同じように[読み込む](scripting-fundamentals.md#sync-and-load)必要があります。

次のスクリプトは、"$" が含まれている任意の番号書式の**A1: D4**範囲を検索します。 このスクリプトは、これらのセルの塗りつぶしの色を "黄" に設定します。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>コレクションを処理する

多くの Excel オブジェクトがコレクションに含まれています。 たとえば、ワークシート内のすべての[図形](/javascript/api/office-scripts/excel/excel.shape)は、 `Worksheet.shapes`プロパティとして、 [offecollection](/javascript/api/office-scripts/excel/excel.shapecollection)に含まれています。 各`*Collection`オブジェクトには`items` 、プロパティが含まれています。これは、そのコレクション内のオブジェクトを格納する配列です。 これは通常の JavaScript 配列と同様に処理できますが、コレクション内の項目を最初に読み込む必要があります。 コレクション内のすべてのオブジェクトのプロパティを操作する必要がある場合は、階層 load ステートメント (`items/propertyName`) を使用します。

次のスクリプトは、現在のワークシート内のすべての図形の種類を記録します。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

`getItem`または`getItemAt`メソッドを使用して、コレクションから個々のオブジェクトを読み込むことができます。 `getItem`名前のような一意の識別子を使用してオブジェクトを取得します (そのような名前は、多くの場合、スクリプトで指定されます)。 `getItemAt`コレクション内のインデックスを使用してオブジェクトを取得します。 オブジェクトを使用するには、 `await context.sync();`その前にコマンドを呼び出す必要があります。

次のスクリプトは、現在のワークシート内の最も古い図形を削除します。

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>日付

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトには、スクリプト内の日付を処理するための標準化された方法が用意されています。 `Date.now()`現在の日付と時刻を使用してオブジェクトを生成します。これは、スクリプトのデータ入力にタイムスタンプを追加するときに便利です。

次のスクリプトは、現在の日付をワークシートに追加します。 この`toLocaleDateString`メソッドを使用すると、Excel によって値が日付として認識され、セルの数値の書式が自動的に変更されることに注意してください。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

サンプルの [[日付を使用して作業](../resources/excel-samples.md#work-with-dates)] セクションには、さらに多くの日付関連スクリプトがあります。

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトには、一般的な数値演算のためのメソッドと定数が用意されています。 これらは、ブックの計算エンジンを使用しなくても、Excel で使用できる多くの関数を提供します。 これにより、スクリプトはブックを照会する必要がなくなり、パフォーマンスが向上します。

次のスクリプトは`Math.min` 、を使用して、 **A1: D4**範囲の最小数を検索して記録します。 この例では、範囲全体に文字列ではなく数値のみが含まれていることに注意してください。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
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
