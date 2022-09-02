---
title: Excel ワークシートの各セルからハイパーリンクを削除する
description: Office スクリプトを使用して Excel ワークシートの各セルからハイパーリンクを削除する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572627"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Excel ワークシートの各セルからハイパーリンクを削除する

 このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。 ワークシートを走査し、セルに関連付けられているハイパーリンクがある場合は、ハイパーリンクをクリアしますが、セルの値はそのまま保持されます。 また、トラバーサルの完了にかかる時間もログに記録されます。

> [!NOTE]
> これは、セル数が 10k <場合にのみ機能します。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックのファイル [remove-hyperlinks.xlsx](remove-hyperlinks.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-remove-hyperlinks"></a>サンプル コード: ハイパーリンクを削除する

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>トレーニング ビデオ: Excel ワークシートの各セルからハイパーリンクを削除する

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/v20fdinxpHU)。
