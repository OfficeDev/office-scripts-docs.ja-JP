---
title: テーブル列フィルターを削除する
description: アクティブなセルの場所に基づいてテーブル列フィルターをクリアする方法について説明します。
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e016f7f2af9e7553229f3b3b19007e011879de8e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572522"
---
# <a name="remove-table-column-filters"></a>テーブル列フィルターを削除する

このサンプルでは、アクティブなセルの場所に基づいて、テーブル列からフィルターを削除します。 スクリプトは、セルがテーブルの一部であるかどうかを検出し、テーブル列を決定し、それに適用されるすべてのフィルターをクリアします。

フィルターをクリアする前にフィルターを保存する (後で再適用する) 方法の詳細については、より高度なサンプルである [フィルターを保存してテーブル間で行を移動](move-rows-across-tables.md)する方法に関するページを参照してください。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックの [table-with-filter.xlsx](table-with-filter.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>サンプル コード: 作業中のセルに基づいてテーブル列フィルターをクリアする

次のスクリプトは、アクティブなセルの場所に基づいてテーブル列フィルターをクリアし、テーブルを含む任意の Excel ファイルに適用できます。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>列フィルターをクリアする前に (アクティブなセルに注意してください)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="列フィルターをクリアする前のアクティブセル。":::

## <a name="after-clearing-column-filter"></a>列フィルターをクリアした後

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="列フィルターをクリアした後のアクティブセル。":::
