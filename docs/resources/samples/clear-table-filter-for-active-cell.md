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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="4ea7d-103">アクティブ セルの場所に基づいてテーブル列フィルターをクリアする</span><span class="sxs-lookup"><span data-stu-id="4ea7d-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="4ea7d-104">このサンプルでは、アクティブセルの場所に基づいてテーブル列フィルターをクリアします。</span><span class="sxs-lookup"><span data-stu-id="4ea7d-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="4ea7d-105">このスクリプトは、セルがテーブルの一部かどうかを検出し、テーブル列を決定し、そのセルに適用されているフィルターをクリアします。</span><span class="sxs-lookup"><span data-stu-id="4ea7d-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="4ea7d-106">フィルターをクリアする前に (および後で再適用する) 前にフィルターを保存する方法の[](move-rows-across-tables.md)詳細については、「フィルターを保存してテーブル間で行を移動する」、より高度なサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="4ea7d-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="4ea7d-107">_列フィルターをクリアする前に (アクティブ セルに注意してください)_</span><span class="sxs-lookup"><span data-stu-id="4ea7d-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="列フィルターをクリアする前のアクティブ セル":::

<span data-ttu-id="4ea7d-109">_列フィルターをクリアした後_</span><span class="sxs-lookup"><span data-stu-id="4ea7d-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="列フィルターをクリアした後のアクティブ セル":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="4ea7d-111">サンプル コード: アクティブ セルに基づいてテーブル列フィルターをクリアする</span><span class="sxs-lookup"><span data-stu-id="4ea7d-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="4ea7d-112">次のスクリプトは、アクティブなセルの場所に基づいてテーブル列フィルターをクリアし、テーブルを持つ任意のExcelに適用できます。</span><span class="sxs-lookup"><span data-stu-id="4ea7d-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="4ea7d-113">便宜上、このファイルをダウンロード<a href="table-with-filter.xlsx">して使用table-with-filter.xlsx。 </a></span><span class="sxs-lookup"><span data-stu-id="4ea7d-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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
