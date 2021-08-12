---
title: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする
description: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする方法について学習します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 5815ae9f40ec1c529bbdc19575239e94712479d3db8a8c602cc33a270538811c
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847567"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>アクティブ セルの場所に基づいてテーブル列フィルターをクリアする

このサンプルでは、アクティブセルの場所に基づいてテーブル列フィルターをクリアします。 このスクリプトは、セルがテーブルの一部かどうかを検出し、テーブル列を決定し、そのセルに適用されているフィルターをクリアします。

フィルターをクリアする前に (および後で再適用する) 前にフィルターを保存する方法の[](move-rows-across-tables.md)詳細については、「フィルターを保存してテーブル間で行を移動する」、より高度なサンプルを参照してください。

_列フィルターをクリアする前に (アクティブ セルに注意してください)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="列フィルターをクリアする前のアクティブ セル。":::

_列フィルターをクリアした後_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="列フィルターをクリアした後のアクティブ なセル。":::

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="table-with-filter.xlsx"> 使用table-with-filter.xlsx</a> ブックのブックをダウンロードします。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>サンプル コード: アクティブ セルに基づいてテーブル列フィルターをクリアする

次のスクリプトは、アクティブなセルの場所に基づいてテーブル列フィルターをクリアし、テーブルを持つ任意のExcelに適用できます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
