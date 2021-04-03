---
title: 複数の Excel テーブルのデータを 1 つのテーブルに結合する
description: 複数の Excel テーブルのデータOffice 1 つのテーブルに結合するために、スクリプトを使用する方法について学習します。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 2f3f7232216f686946861d8c2cdec44013333ec7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571417"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="65447-103">複数の Excel テーブルのデータを 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="65447-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="65447-104">このサンプルでは、複数の Excel テーブルのデータを、すべての行を含む 1 つのテーブルに結合します。</span><span class="sxs-lookup"><span data-stu-id="65447-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="65447-105">使用されているテーブルはすべて同じ構造を持つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="65447-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="65447-106">このスクリプトには、次の 2 つのバリエーションがあります。</span><span class="sxs-lookup"><span data-stu-id="65447-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="65447-107">最初 [のスクリプトは](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) 、Excel ファイル内のすべてのテーブルを結合します。</span><span class="sxs-lookup"><span data-stu-id="65447-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="65447-108">2 [番目のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) は、ワークシートのセット内のテーブルを選択的に取得します。</span><span class="sxs-lookup"><span data-stu-id="65447-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="65447-109">サンプル コード: 複数の Excel テーブルのデータを 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="65447-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="65447-110">サンプル ファイルをダウンロード <a href="tables-copy.xlsx"> してtables-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="65447-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="65447-111">サンプル コード: 選択したワークシートの複数の Excel テーブルのデータを 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="65447-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="65447-112">サンプル ファイルをダウンロード <a href="tables-select-copy.xlsx"> してtables-select-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="65447-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="65447-113">トレーニング ビデオ: 複数の Excel テーブルのデータを 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="65447-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="65447-114">[![複数の Excel テーブルのデータを 1 つのテーブルに結合する方法について、ステップバイステップのビデオを見る](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "複数の Excel テーブルのデータを 1 つのテーブルに結合する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="65447-114">[![Watch step-by-step video on how to combine data from multiple Excel tables into a single table](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Step-by-step video on how to combine data from multiple Excel tables into a single table")</span></span>
