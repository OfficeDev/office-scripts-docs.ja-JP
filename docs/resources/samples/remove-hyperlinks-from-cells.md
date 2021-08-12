---
title: ワークシート内の各セルからハイパーリンクをExcelする
description: '[スクリプト] を使用してOfficeワークシートの各セルからハイパーリンクを削除するExcelします。'
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 498d55ea1ee7926ab124d00795825660005c5e38e73ed5d90fe8f9208a583908
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847430"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>ワークシート内の各セルからハイパーリンクをExcelする

 このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。 ワークシートを走査し、セルに関連付けられているハイパーリンクがある場合は、ハイパーリンクをクリアしますが、セルの値は保持されます。 また、トラバーサルの完了に要する時間も記録します。

> [!NOTE]
> これは、セル数が 10k の場合<機能します。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに使用できる <a href="remove-hyperlinks.xlsx"> ブックremove-hyperlinks.xlsx</a> ファイル をダウンロードします。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-remove-hyperlinks"></a>サンプル コード: ハイパーリンクの削除

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>トレーニング ビデオ: ワークシート内の各セルからハイパーリンクをExcelする

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/v20fdinxpHU).
