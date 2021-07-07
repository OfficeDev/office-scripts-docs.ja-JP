---
title: スクリプトを使用してテーブル間で行Officeする
description: フィルターを保存し、フィルターを処理して再適用することで、テーブル間で行を移動する方法について学習します。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 860521de166108d5a8355ea246c1bfe77e0e064b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313758"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="80703-103">フィルターを保存し、フィルターを処理して再適用することで、テーブル間で行を移動する</span><span class="sxs-lookup"><span data-stu-id="80703-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="80703-104">このスクリプトでは、次のことが行われます。</span><span class="sxs-lookup"><span data-stu-id="80703-104">This script does the following:</span></span>

* <span data-ttu-id="80703-105">列の値が一部の値と等しいソース テーブルから行を _選択します_。</span><span class="sxs-lookup"><span data-stu-id="80703-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="80703-106">選択した行を別のワークシートの別の (ターゲット) テーブルに移動します。</span><span class="sxs-lookup"><span data-stu-id="80703-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="80703-107">ソース テーブルに関連するフィルターを再適用します。</span><span class="sxs-lookup"><span data-stu-id="80703-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="ブックの前と後のスクリーンショット。":::

## <a name="sample-excel-file"></a><span data-ttu-id="80703-109">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="80703-109">Sample Excel file</span></span>

<span data-ttu-id="80703-110">すぐに使用できる <a href="input-table-filters.xlsx"> ブックinput-table-filters.xlsx</a> ファイル をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="80703-110">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="80703-111">次のスクリプトを追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="80703-111">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="80703-112">サンプル コード: 範囲の値を使用して行を移動する</span><span class="sxs-lookup"><span data-stu-id="80703-112">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="80703-113">トレーニング ビデオ: テーブル間で行を移動する</span><span class="sxs-lookup"><span data-stu-id="80703-113">Training video: Move rows across tables</span></span>

<span data-ttu-id="80703-114">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/_3t3Pk4i2L0).</span><span class="sxs-lookup"><span data-stu-id="80703-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span> <span data-ttu-id="80703-115">ビデオのソリューションには、2 つのスクリプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="80703-115">There are two scripts shown in the video's solution.</span></span> <span data-ttu-id="80703-116">主な違いは、行の選択方法です。</span><span class="sxs-lookup"><span data-stu-id="80703-116">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="80703-117">1 つ目のバリアントでは、テーブル フィルターを適用し、表示範囲を読み取って行を選択します。</span><span class="sxs-lookup"><span data-stu-id="80703-117">In the first variant, the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="80703-118">2 番目の行は、値を読み取り、行の値 (このページのサンプルで使用される値) を抽出して選択します。</span><span class="sxs-lookup"><span data-stu-id="80703-118">In the second, the rows are selected by reading the values and extracting the row values (which is what the sample on this page uses).</span></span>
