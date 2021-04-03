---
title: Excel ワークシートの各セルからハイパーリンクを削除する
description: Excel ワークシートの各セルOfficeハイパーリンクを削除するには、スクリプトを使用する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 07b670aac3368e38b9b93283404befee608391a7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571294"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="e2b97-103">Excel ワークシートの各セルからハイパーリンクを削除する</span><span class="sxs-lookup"><span data-stu-id="e2b97-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="e2b97-104">このサンプルでは、現在のワークシートからすべてのハイパーリンクをクリアします。</span><span class="sxs-lookup"><span data-stu-id="e2b97-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="e2b97-105">ワークシートを走査し、セルに関連付けられているハイパーリンクがある場合は、ハイパーリンクをクリアしますが、セルの値は保持されます。</span><span class="sxs-lookup"><span data-stu-id="e2b97-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="e2b97-106">また、トラバーサルの完了に要する時間も記録します。</span><span class="sxs-lookup"><span data-stu-id="e2b97-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="e2b97-107">これは、セル数が 10k の場合<機能します。</span><span class="sxs-lookup"><span data-stu-id="e2b97-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="e2b97-108">サンプル コード: ハイパーリンクの削除</span><span class="sxs-lookup"><span data-stu-id="e2b97-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="e2b97-109">このサンプルで <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> ファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="e2b97-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {

  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);
  const targetRange = sheet.getUsedRange(true);
  if (!targetRange) {
    console.log(`There is no data in the worksheet. `)
    return;
  }
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  const totalCells = rowCount * colCount;
  if (totalCells > 10000) {
    console.log("Too many cells to operate with. Consider editing script to use selected range and then remove hyperlinks in batches. " + targetRange.getAddress());
    return;
  }
  // Call the helper function to remove the hyperlinks. 
  removeHyperLink(targetRange);
  return;
}

/**
 * Removes hyperlink for each cell in the target range. Logs the time it takes to complete traversal.
 * @param targetRange Target range to clear the hyperlinks from.
 */
function removeHyperLink(targetRange: ExcelScript.Range): void {
  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);
  let clearedCount = 0;
  let cellsVisited = 0;

  let groupStart = new Date().getTime();
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      cellsVisited++;
      if (cellsVisited % 50 === 0) {
        let groupEnd = new Date().getTime();
        console.log(`Completed ${cellsVisited} cells out of ${rowCount * colCount}. This group took: ${(groupEnd - groupStart) / 1000} seconds to complete.`);
        groupStart = new Date().getTime();
      }
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
  console.log(`Done. Inspected ${cellsVisited} cells. Cleared hyperlinks in: ${clearedCount} cells`);
  return;
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="e2b97-110">トレーニング ビデオ: Excel ワークシートの各セルからハイパーリンクを削除する</span><span class="sxs-lookup"><span data-stu-id="e2b97-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="e2b97-111">[![Excel ワークシートの各セルからハイパーリンクを削除する方法について、ステップバイステップのビデオを見る](../../images/hyperlinks-vid.jpg)](https://youtu.be/v20fdinxpHU "Excel ワークシートの各セルからハイパーリンクを削除する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="e2b97-111">[![Watch step-by-step video on how to remove hyperlinks from each cell in an Excel worksheet](../../images/hyperlinks-vid.jpg)](https://youtu.be/v20fdinxpHU "Step-by-step video on how to remove hyperlinks from each cell in an Excel worksheet")</span></span>
