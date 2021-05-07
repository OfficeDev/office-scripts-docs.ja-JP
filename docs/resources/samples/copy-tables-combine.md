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
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>複数のテーブルのデータExcel 1 つのテーブルに結合する

このサンプルでは、複数のテーブルExcelデータを、すべての行を含む 1 つのテーブルに結合します。 使用されているテーブルはすべて同じ構造を持つ必要があります。

このスクリプトには、次の 2 つのバリエーションがあります。

1. 最初[のスクリプトは](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)、ファイル内のすべてのテーブルを結合Excelします。
1. 2 [番目のスクリプト](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) は、ワークシートのセット内のテーブルを選択的に取得します。

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>サンプル コード: 複数のテーブルのデータExcel 1 つのテーブルに結合する

サンプル ファイルをダウンロード <a href="tables-copy.xlsx"> してtables-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>サンプル コード: 選択したワークシート内の複数Excelテーブルのデータを 1 つのテーブルに結合する

サンプル ファイルをダウンロード <a href="tables-select-copy.xlsx"> してtables-select-copy.xlsx</a> スクリプトと一緒に使用して、自分で試してみてください。

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>トレーニング ビデオ: 複数のテーブルのデータExcel 1 つのテーブルに結合する

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/di-8JukK3Lc).
