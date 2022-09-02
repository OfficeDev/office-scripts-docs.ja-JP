---
title: 複数の Excel テーブルのデータを 1 つのテーブルに結合する
description: Office スクリプトを使用して、複数の Excel テーブルのデータを 1 つのテーブルに結合する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3db510514c676b9012fd47abc2a7e92492a9cf87
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572452"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>複数の Excel テーブルのデータを 1 つのテーブルに結合する

このサンプルでは、複数の Excel テーブルのデータを、すべての行を含む 1 つのテーブルに結合します。 使用されているすべてのテーブルの構造が同じであることを前提としています。

このスクリプトには 2 つのバリエーションがあります。

1. [最初のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)は、Excel ファイル内のすべてのテーブルを結合します。
1. [2 番目のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table)は、ワークシートのセット内のテーブルを選択的に取得します。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックの [tables-copy.xlsx](tables-copy.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>サンプル コード: 複数の Excel テーブルのデータを 1 つのテーブルに結合する

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>サンプル コード: 選択したワークシートの複数の Excel テーブルのデータを 1 つのテーブルに結合する

[tables-select-copy.xlsx](tables-select-copy.xlsx)サンプル ファイルをダウンロードし、次のスクリプトで使用して自分で試してください。

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>トレーニング ビデオ: 複数の Excel テーブルのデータを 1 つのテーブルに結合する

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/di-8JukK3Lc)。
