---
title: Office スクリプトの基本的なExcel on the web
description: スクリプト内のスクリプトと一緒にOfficeコード サンプルのコレクションExcel on the web。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545760"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="82e9b-103">Office スクリプトの基本的なExcel on the web</span><span class="sxs-lookup"><span data-stu-id="82e9b-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="82e9b-104">次のサンプルは、独自のブックで試す簡単なスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="82e9b-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="82e9b-105">次の方法で使用Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="82e9b-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="82e9b-106">**[自動化]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="82e9b-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="82e9b-107">[コード **エディター] を押します**。</span><span class="sxs-lookup"><span data-stu-id="82e9b-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="82e9b-108">コード **エディターの作業** ウィンドウで [新しいスクリプト] を押します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="82e9b-109">スクリプト全体を、選択したサンプルに置き換える。</span><span class="sxs-lookup"><span data-stu-id="82e9b-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="82e9b-110">コード **エディターの** 作業ウィンドウで [実行] を押します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="82e9b-111">スクリプトの基本</span><span class="sxs-lookup"><span data-stu-id="82e9b-111">Script basics</span></span>

<span data-ttu-id="82e9b-112">これらのサンプルでは、スクリプトの基本的な構成要素Office示します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="82e9b-113">これらをスクリプトに追加して、ソリューションを拡張し、一般的な問題を解決します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="82e9b-114">1 つのセルを読み取り、ログに記録する</span><span class="sxs-lookup"><span data-stu-id="82e9b-114">Read and log one cell</span></span>

<span data-ttu-id="82e9b-115">このサンプルでは **、A1 の値を読み** 取り、コンソールに出力します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="82e9b-116">アクティブ セルの読み取り</span><span class="sxs-lookup"><span data-stu-id="82e9b-116">Read the active cell</span></span>

<span data-ttu-id="82e9b-117">このスクリプトは、現在のアクティブ セルの値をログに記録します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="82e9b-118">複数のセルが選択されている場合は、左上のセルがログに記録されます。</span><span class="sxs-lookup"><span data-stu-id="82e9b-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="82e9b-119">隣接するセルを変更する</span><span class="sxs-lookup"><span data-stu-id="82e9b-119">Change an adjacent cell</span></span>

<span data-ttu-id="82e9b-120">このスクリプトは、相対参照を使用して隣接するセルを取得します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="82e9b-121">アクティブ セルが一番上の行にある場合、スクリプトの一部は、現在選択されているセルの上にあるセルを参照しますので、失敗します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="82e9b-122">隣接するセルを変更する</span><span class="sxs-lookup"><span data-stu-id="82e9b-122">Change all adjacent cells</span></span>

<span data-ttu-id="82e9b-123">このスクリプトは、アクティブ セルの書式設定を隣接セルにコピーします。</span><span class="sxs-lookup"><span data-stu-id="82e9b-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="82e9b-124">このスクリプトは、アクティブ セルがワークシートの端にない場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="82e9b-125">範囲内の各セルを変更する</span><span class="sxs-lookup"><span data-stu-id="82e9b-125">Change each individual cell in a range</span></span>

<span data-ttu-id="82e9b-126">このスクリプトは、現在選択されている範囲をループ処理します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-126">This script loops over the currently select range.</span></span> <span data-ttu-id="82e9b-127">現在の書式設定をクリアし、各セルの塗りつぶしの色をランダムな色に設定します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="82e9b-128">特別な条件に基づいてセルのグループを取得する</span><span class="sxs-lookup"><span data-stu-id="82e9b-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="82e9b-129">このスクリプトは、現在のワークシートの使用範囲内のすべての空白セルを取得します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="82e9b-130">次に、これらのすべてのセルを黄色の背景で強調表示します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-130">It then highlights all those cells with a yellow background.</span></span>

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

## <a name="collections"></a><span data-ttu-id="82e9b-131">コレクション</span><span class="sxs-lookup"><span data-stu-id="82e9b-131">Collections</span></span>

<span data-ttu-id="82e9b-132">これらのサンプルは、ブック内のオブジェクトのコレクションで動作します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="82e9b-133">コレクションを反復処理する</span><span class="sxs-lookup"><span data-stu-id="82e9b-133">Iterate over collections</span></span>

<span data-ttu-id="82e9b-134">このスクリプトは、ブック内のすべてのワークシートの名前を取得してログに記録します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="82e9b-135">また、タブの色をランダムな色に設定します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-135">It also sets the their tab colors to a random color.</span></span>

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

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="82e9b-136">コレクションのクエリと削除</span><span class="sxs-lookup"><span data-stu-id="82e9b-136">Query and delete from a collection</span></span>

<span data-ttu-id="82e9b-137">このスクリプトは、新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-137">This script creates a new worksheet.</span></span> <span data-ttu-id="82e9b-138">ワークシートの既存のコピーをチェックし、新しいシートを作成する前に削除します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

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

## <a name="dates"></a><span data-ttu-id="82e9b-139">日付</span><span class="sxs-lookup"><span data-stu-id="82e9b-139">Dates</span></span>

<span data-ttu-id="82e9b-140">このセクションのサンプルでは、JavaScript Date オブジェクトの使い方 [を示](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="82e9b-141">次のサンプルでは、現在の日付と時刻を取得し、それらの値をアクティブワークシートの 2 つのセルに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="82e9b-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="82e9b-142">次のサンプルでは、データに格納されている日付を読みExcel JavaScript Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="82e9b-143">日付の数値 [シリアル番号を](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) JavaScript Date の入力として使用します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="82e9b-144">データの表示</span><span class="sxs-lookup"><span data-stu-id="82e9b-144">Display data</span></span>

<span data-ttu-id="82e9b-145">これらのサンプルでは、ワークシート データを処理し、より良いビューまたは組織をユーザーに提供する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="82e9b-146">条件付き書式の適用</span><span class="sxs-lookup"><span data-stu-id="82e9b-146">Apply conditional formatting</span></span>

<span data-ttu-id="82e9b-147">このサンプルでは、ワークシートで現在使用されている範囲に条件付き書式を適用します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="82e9b-148">条件付き書式は、値の上位 10% の緑の塗りつぶしです。</span><span class="sxs-lookup"><span data-stu-id="82e9b-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="82e9b-149">並べ替えテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="82e9b-149">Create a sorted table</span></span>

<span data-ttu-id="82e9b-150">このサンプルでは、現在のワークシートの使用範囲からテーブルを作成し、最初の列に基づいて並べ替えを行います。</span><span class="sxs-lookup"><span data-stu-id="82e9b-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="82e9b-151">ピボットテーブルから "Grand Total" 値をログに記録する</span><span class="sxs-lookup"><span data-stu-id="82e9b-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="82e9b-152">このサンプルでは、ブック内の最初のピボットテーブルを検索し、値を "Grand Total" セルに記録します (下の図では緑色で強調表示されています)。</span><span class="sxs-lookup"><span data-stu-id="82e9b-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="[総計] 行が緑色で強調表示された、果物の売上を示すピボットテーブル":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="82e9b-154">データ検証を使用してドロップダウン リストを作成する</span><span class="sxs-lookup"><span data-stu-id="82e9b-154">Create a drop-down list using data validation</span></span>

<span data-ttu-id="82e9b-155">このスクリプトは、セルのドロップダウン選択リストを作成します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="82e9b-156">選択した範囲の既存の値をリストの選択肢として使用します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-156">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="色の選択肢 '赤、青、緑' を含む 3 つのセルの範囲を示すワークシートと、その横にドロップダウン リストに表示されるのと同じ選択肢":::

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

## <a name="formulas"></a><span data-ttu-id="82e9b-158">数式</span><span class="sxs-lookup"><span data-stu-id="82e9b-158">Formulas</span></span>

<span data-ttu-id="82e9b-159">これらのサンプルでは、Excelを使用し、スクリプトで使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="82e9b-160">単一の数式</span><span class="sxs-lookup"><span data-stu-id="82e9b-160">Single formula</span></span>

<span data-ttu-id="82e9b-161">このスクリプトは、セルの数式を設定し、セルExcel値を個別に格納する方法を表示します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="82e9b-162">数式から `#SPILL!` 返されるエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="82e9b-162">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="82e9b-163">このスクリプトは、TRANSPOSE 関数を使用して範囲 "A1:D2" を "A4:B7" にトランスポーズします。</span><span class="sxs-lookup"><span data-stu-id="82e9b-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="82e9b-164">トランスポーズでエラーが発生した場合は、ターゲット範囲をクリアし `#SPILL` 、数式を再度適用します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-164">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

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

## <a name="suggest-new-samples"></a><span data-ttu-id="82e9b-165">新しいサンプルの提案</span><span class="sxs-lookup"><span data-stu-id="82e9b-165">Suggest new samples</span></span>

<span data-ttu-id="82e9b-166">新しいサンプルの提案を歓迎します。</span><span class="sxs-lookup"><span data-stu-id="82e9b-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="82e9b-167">他のスクリプト開発者に役立つ一般的なシナリオがある場合は、ページの下部にあるフィードバック セクションで教えて下さい。</span><span class="sxs-lookup"><span data-stu-id="82e9b-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="82e9b-168">関連項目</span><span class="sxs-lookup"><span data-stu-id="82e9b-168">See also</span></span>

* [<span data-ttu-id="82e9b-169">Sudhi Ramamurthy の YouTube の "Range basics"</span><span class="sxs-lookup"><span data-stu-id="82e9b-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="82e9b-170">Officeスクリプトのサンプルとシナリオ</span><span class="sxs-lookup"><span data-stu-id="82e9b-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="82e9b-171">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="82e9b-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
