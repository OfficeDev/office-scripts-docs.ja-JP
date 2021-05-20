---
title: Excel on the webでのスクリプトのOfficeの基本スクリプト
description: Excel on the webのスクリプトで使用するコード サンプルのコレクションOffice。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545760"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="33bf4-103">Excel on the webでのスクリプトのOfficeの基本スクリプト</span><span class="sxs-lookup"><span data-stu-id="33bf4-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="33bf4-104">以下のサンプルは、独自のワークブックを試用するための簡単なスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="33bf4-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="33bf4-105">Excel on the webでそれらを使用するには:</span><span class="sxs-lookup"><span data-stu-id="33bf4-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="33bf4-106">**[自動化]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="33bf4-107">**コード エディタ を** 押します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="33bf4-108">[コード エディタ] 作業ウィンドウで **[新しいスクリプト** ] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="33bf4-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="33bf4-109">スクリプト全体を、選択したサンプルに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="33bf4-110">[コード エディタ] 作業ウィンドウで [ **実行** ] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="33bf4-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="33bf4-111">スクリプトの基本</span><span class="sxs-lookup"><span data-stu-id="33bf4-111">Script basics</span></span>

<span data-ttu-id="33bf4-112">これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="33bf4-113">これらのスクリプトをスクリプトに追加して、ソリューションを拡張し、一般的な問題を解決します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="33bf4-114">1 つのセルの読み取りとログ記録</span><span class="sxs-lookup"><span data-stu-id="33bf4-114">Read and log one cell</span></span>

<span data-ttu-id="33bf4-115">このサンプルでは **、A1** の値を読み取り、コンソールに出力します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="33bf4-116">アクティブ セルを読み取る</span><span class="sxs-lookup"><span data-stu-id="33bf4-116">Read the active cell</span></span>

<span data-ttu-id="33bf4-117">このスクリプトは、現在アクティブなセルの値をログに記録します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="33bf4-118">複数のセルが選択されている場合は、一番上のセルがログに記録されます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="33bf4-119">隣接するセルを変更する</span><span class="sxs-lookup"><span data-stu-id="33bf4-119">Change an adjacent cell</span></span>

<span data-ttu-id="33bf4-120">このスクリプトは、相対参照を使用して隣接するセルを取得します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="33bf4-121">アクティブ セルが一番上の行にある場合、スクリプトの一部は、現在選択されているセルの上のセルを参照するため、失敗します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="33bf4-122">隣接するすべてのセルを変更する</span><span class="sxs-lookup"><span data-stu-id="33bf4-122">Change all adjacent cells</span></span>

<span data-ttu-id="33bf4-123">このスクリプトは、アクティブ セルの書式を隣接するセルにコピーします。</span><span class="sxs-lookup"><span data-stu-id="33bf4-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="33bf4-124">このスクリプトは、アクティブ セルがワークシートの端にない場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="33bf4-125">範囲内の各セルを変更する</span><span class="sxs-lookup"><span data-stu-id="33bf4-125">Change each individual cell in a range</span></span>

<span data-ttu-id="33bf4-126">このスクリプトは、現在選択されている範囲をループします。</span><span class="sxs-lookup"><span data-stu-id="33bf4-126">This script loops over the currently select range.</span></span> <span data-ttu-id="33bf4-127">現在の書式をクリアし、各セルの塗りつぶしの色をランダムな色に設定します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="33bf4-128">特殊な条件に基づいてセルのグループを取得する</span><span class="sxs-lookup"><span data-stu-id="33bf4-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="33bf4-129">このスクリプトは、現在のワークシートで使用されている範囲内のすべての空白セルを取得します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="33bf4-130">次に、黄色の背景を持つすべてのセルを強調表示します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-130">It then highlights all those cells with a yellow background.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="33bf4-131">コレクション</span><span class="sxs-lookup"><span data-stu-id="33bf4-131">Collections</span></span>

<span data-ttu-id="33bf4-132">これらのサンプルは、ブック内のオブジェクトのコレクションを扱います。</span><span class="sxs-lookup"><span data-stu-id="33bf4-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="33bf4-133">コレクションを反復処理する</span><span class="sxs-lookup"><span data-stu-id="33bf4-133">Iterate over collections</span></span>

<span data-ttu-id="33bf4-134">このスクリプトは、ブック内のすべてのワークシートの名前を取得し、ログに記録します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="33bf4-135">タブの色もランダムな色に設定されます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-135">It also sets the their tab colors to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="33bf4-136">コレクションのクエリと削除</span><span class="sxs-lookup"><span data-stu-id="33bf4-136">Query and delete from a collection</span></span>

<span data-ttu-id="33bf4-137">このスクリプトは、新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-137">This script creates a new worksheet.</span></span> <span data-ttu-id="33bf4-138">ワークシートの既存のコピーをチェックし、新しいシートを作成する前に削除します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a><span data-ttu-id="33bf4-139">日付</span><span class="sxs-lookup"><span data-stu-id="33bf4-139">Dates</span></span>

<span data-ttu-id="33bf4-140">このセクションのサンプルでは、JavaScript [日付](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) オブジェクトの使用方法を示します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="33bf4-141">次のサンプルでは、現在の日付と時刻を取得し、アクティブ ワークシート内の 2 つのセルにこれらの値を書き込みます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="33bf4-142">次のサンプルでは、Excelに格納されている日付を読み取り、JavaScript Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="33bf4-143">JavaScript [日付の入力として日付の数値シリアル番号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) を使用します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="33bf4-144">データの表示</span><span class="sxs-lookup"><span data-stu-id="33bf4-144">Display data</span></span>

<span data-ttu-id="33bf4-145">これらのサンプルでは、ワークシート データを操作する方法を示し、ユーザーにより良いビューや組織を提供します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="33bf4-146">条件付き書式の適用</span><span class="sxs-lookup"><span data-stu-id="33bf4-146">Apply conditional formatting</span></span>

<span data-ttu-id="33bf4-147">このサンプルでは、ワークシートで現在使用されている範囲に条件付き書式を適用します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="33bf4-148">条件付き書式は、値の上位 10% に対して緑の塗りつぶしです。</span><span class="sxs-lookup"><span data-stu-id="33bf4-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="33bf4-149">並べ替えられたテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="33bf4-149">Create a sorted table</span></span>

<span data-ttu-id="33bf4-150">このサンプルでは、現在のワークシートで使用されている範囲からテーブルを作成し、最初の列に基づいて並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="33bf4-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="33bf4-151">ピボットテーブルから "総計" 値をログに記録する</span><span class="sxs-lookup"><span data-stu-id="33bf4-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="33bf4-152">このサンプルでは、ブック内の最初のピボットテーブルを検索し、値を "総計" セルに記録します (下の図では緑色で強調表示されています)。</span><span class="sxs-lookup"><span data-stu-id="33bf4-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="大合計行が緑色にハイライト表示された果物の売り上げを示すピボットテーブル":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="33bf4-154">データの入力規則を使用してドロップダウン リストを作成する</span><span class="sxs-lookup"><span data-stu-id="33bf4-154">Create a drop-down list using data validation</span></span>

<span data-ttu-id="33bf4-155">このスクリプトは、セルのドロップダウン選択リストを作成します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="33bf4-156">選択した範囲の既存の値をリストの選択肢として使用します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-156">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="色の選択肢が 「赤、青、緑」を含む 3 つのセルの範囲を示すワークシートと、その横に表示されるドロップダウン リストと同じ選択肢":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a><span data-ttu-id="33bf4-158">数式</span><span class="sxs-lookup"><span data-stu-id="33bf4-158">Formulas</span></span>

<span data-ttu-id="33bf4-159">これらのサンプルでは、Excel式を使用し、スクリプトでの操作方法を示します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="33bf4-160">単一の式</span><span class="sxs-lookup"><span data-stu-id="33bf4-160">Single formula</span></span>

<span data-ttu-id="33bf4-161">このスクリプトは、セルの数式を設定し、次に、Excelがセルの数式と値を別々に格納する方法を表示します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="33bf4-162">`#SPILL!`数式から返されたエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="33bf4-162">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="33bf4-163">このスクリプトは、TRANSPOSE 関数を使用して、範囲 "A1:D2" を "A4:B7" に変換します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="33bf4-164">転置がエラーの原因となった `#SPILL` 場合は、対象範囲をクリアし、数式を再度適用します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-164">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a><span data-ttu-id="33bf4-165">新しいサンプルの提案</span><span class="sxs-lookup"><span data-stu-id="33bf4-165">Suggest new samples</span></span>

<span data-ttu-id="33bf4-166">新しいサンプルの提案を歓迎します。</span><span class="sxs-lookup"><span data-stu-id="33bf4-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="33bf4-167">他のスクリプト開発者に役立つ一般的なシナリオがある場合は、ページの下部にあるフィードバックセクションで教えてください。</span><span class="sxs-lookup"><span data-stu-id="33bf4-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="33bf4-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="33bf4-168">See also</span></span>

* [<span data-ttu-id="33bf4-169">スーディ・ラマムルティのYouTubeでの「レンジの基本」</span><span class="sxs-lookup"><span data-stu-id="33bf4-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="33bf4-170">Officeスクリプトのサンプルとシナリオ</span><span class="sxs-lookup"><span data-stu-id="33bf4-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="33bf4-171">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="33bf4-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
