---
title: 複数のテーブルのデータExcel 1 つのテーブルに結合する
description: スクリプトを使用して複数のOfficeテーブルのデータを 1 つのExcelに結合する方法について学習します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: bfe2f4af579e5ee59d20e12fe4850dd9c41bcc56
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585892"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>複数のテーブルのデータExcel 1 つのテーブルに結合する

このサンプルでは、複数のテーブルExcelデータを、すべての行を含む 1 つのテーブルに結合します。 使用されているテーブルはすべて同じ構造を持つ必要があります。

このスクリプトには、次の 2 つのバリエーションがあります。

1. 最初[のスクリプトは、](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)ファイル内のすべてのテーブルを結合Excelします。
1. [2 番目のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table)は、ワークシートのセット内のテーブルを選択的に取得します。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="tables-copy.xlsx">tables-copy.xlsx</a> ブックのダウンロード を行います。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>サンプル コード: 複数のテーブルのデータExcel 1 つのテーブルに結合する

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>サンプル コード: 選択したワークシート内の複数Excelテーブルのデータを 1 つのテーブルに結合する

サンプル ファイルをダウンロード <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> 、次のスクリプトと一緒に使用して、自分で試してみてください。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>トレーニング ビデオ: 複数のテーブルのデータExcel 1 つのテーブルに結合する

[Sudhi Ramamurthy が YouTube でこのサンプルを見るのを見る](https://youtu.be/di-8JukK3Lc)。
