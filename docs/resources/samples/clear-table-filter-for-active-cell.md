---
title: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする
description: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする方法について学習します。
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 4f8353fb5480812b7b63e7a9b3ffb11ece2a8c6c
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755085"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>アクティブ セルの場所に基づいてテーブル列フィルターをクリアする

このサンプルでは、アクティブセルの場所に基づいてテーブル列フィルターをクリアします。 このスクリプトは、セルがテーブルの一部かどうかを検出し、テーブル列を決定し、そのセルに適用されているフィルターをクリアします。

フィルターをクリアする前に (および後で再適用する) 前にフィルターを保存する方法の[](move-rows-across-tables.md)詳細については、「フィルターを保存してテーブル間で行を移動する」、より高度なサンプルを参照してください。

_列フィルターをクリアする前に (アクティブ セルに注意してください)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="列フィルターをクリアする前のアクティブ セル。":::

_列フィルターをクリアした後_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="列フィルターをクリアした後のアクティブ なセル。":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>サンプル コード: アクティブ セルに基づいてテーブル列フィルターをクリアする

次のスクリプトは、アクティブセルの場所に基づいてテーブル列フィルターをクリアし、テーブルを持つ任意の Excel ファイルに適用できます。 便宜上、このファイルをダウンロード<a href="table-with-filter.xlsx">して使用table-with-filter.xlsx。 </a>

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

## <a name="training-video-clear-table-column-filter-based-on-active-cell-location"></a>トレーニング ビデオ: アクティブ セルの場所に基づいてテーブル列フィルターをクリアする

範囲を操作する方法の例については [、「Range basics training videos」を参照してください](range-basics.md#training-videos-range-basics)。
