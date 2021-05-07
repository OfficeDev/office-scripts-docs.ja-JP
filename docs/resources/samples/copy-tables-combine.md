---
title: 複数のテーブルのデータExcel 1 つのテーブルに結合する
description: 複数のテーブルから 1 Officeテーブルのデータを結合するために、Excelスクリプトを使用する方法について学習します。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232446"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="9b9b3-103">複数のテーブルのデータExcel 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="9b9b3-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="9b9b3-104">このサンプルでは、複数のテーブルExcelデータを、すべての行を含む 1 つのテーブルに結合します。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="9b9b3-105">使用されているテーブルはすべて同じ構造を持つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="9b9b3-106">このスクリプトには、次の 2 つのバリエーションがあります。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="9b9b3-107">最初[のスクリプトは](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)、ファイル内のすべてのテーブルを結合Excelします。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="9b9b3-108">2 [番目のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) は、ワークシートのセット内のテーブルを選択的に取得します。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="9b9b3-109">サンプル コード: 複数のテーブルのデータExcel 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="9b9b3-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="9b9b3-110">サンプル ファイルをダウンロード <a href="tables-copy.xlsx"> してtables-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="9b9b3-111">サンプル コード: 選択したワークシート内の複数Excelテーブルのデータを 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="9b9b3-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="9b9b3-112">サンプル ファイルをダウンロード <a href="tables-select-copy.xlsx"> してtables-select-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="9b9b3-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="9b9b3-113">トレーニング ビデオ: 複数のテーブルのデータExcel 1 つのテーブルに結合する</span><span class="sxs-lookup"><span data-stu-id="9b9b3-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="9b9b3-114">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/di-8JukK3Lc).</span><span class="sxs-lookup"><span data-stu-id="9b9b3-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
