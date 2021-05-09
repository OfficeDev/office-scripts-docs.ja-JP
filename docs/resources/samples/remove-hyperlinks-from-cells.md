---
title: ワークシート内の各セルからハイパーリンクをExcelする
description: '[スクリプト] を使用してOfficeワークシートの各セルからハイパーリンクを削除するExcelします。'
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 048d01691377a7086bdba9ceb87ca98839cfa4d1
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285802"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="7b787-103">ワークシート内の各セルからハイパーリンクをExcelする</span><span class="sxs-lookup"><span data-stu-id="7b787-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="7b787-104">このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。</span><span class="sxs-lookup"><span data-stu-id="7b787-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="7b787-105">ワークシートを走査し、セルに関連付けられているハイパーリンクがある場合は、ハイパーリンクをクリアしますが、セルの値は保持されます。</span><span class="sxs-lookup"><span data-stu-id="7b787-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="7b787-106">また、トラバーサルの完了に要する時間も記録します。</span><span class="sxs-lookup"><span data-stu-id="7b787-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="7b787-107">これは、セル数が 10k の場合<機能します。</span><span class="sxs-lookup"><span data-stu-id="7b787-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="7b787-108">サンプル コード: ハイパーリンクの削除</span><span class="sxs-lookup"><span data-stu-id="7b787-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="7b787-109">このサンプルで <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> ファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="7b787-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="7b787-110">トレーニング ビデオ: ワークシート内の各セルからハイパーリンクをExcelする</span><span class="sxs-lookup"><span data-stu-id="7b787-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="7b787-111">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/v20fdinxpHU).</span><span class="sxs-lookup"><span data-stu-id="7b787-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
