---
title: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする
description: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする方法について学習します。
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232502"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>アクティブ セルの場所に基づいてテーブル列フィルターをクリアする

このサンプルでは、アクティブセルの場所に基づいてテーブル列フィルターをクリアします。 このスクリプトは、セルがテーブルの一部かどうかを検出し、テーブル列を決定し、そのセルに適用されているフィルターをクリアします。

フィルターをクリアする前に (および後で再適用する) 前にフィルターを保存する方法の[](move-rows-across-tables.md)詳細については、「フィルターを保存してテーブル間で行を移動する」、より高度なサンプルを参照してください。

_列フィルターをクリアする前に (アクティブ セルに注意してください)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="列フィルターをクリアする前のアクティブ セル":::

_列フィルターをクリアした後_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="列フィルターをクリアした後のアクティブ セル":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>サンプル コード: アクティブ セルに基づいてテーブル列フィルターをクリアする

次のスクリプトは、アクティブなセルの場所に基づいてテーブル列フィルターをクリアし、テーブルを持つ任意のExcelに適用できます。 便宜上、このファイルをダウンロード<a href="table-with-filter.xlsx">して使用table-with-filter.xlsx。 </a>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
