---
title: Excel on the web の Office スクリプトのサンプルスクリプト
description: Web 上の Excel の Office スクリプトで使用するコードサンプルのコレクションです。
ms.date: 08/04/2020
localization_priority: Normal
ms.openlocfilehash: 4f8d6f2395a841a8dcba2ea0e712e645a84a6d91
ms.sourcegitcommit: 1c88abcf5df16a05913f12df89490ce843cfebe2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/13/2020
ms.locfileid: "46665230"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="26301-103">Web 上の Excel での Office スクリプトのサンプルスクリプト (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="26301-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="26301-104">次のサンプルは、独自のブックで試すことができる簡単なスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="26301-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="26301-105">Web 上の Excel で使用するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="26301-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="26301-106">**[自動化]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="26301-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="26301-107">**コードエディター**を押します。</span><span class="sxs-lookup"><span data-stu-id="26301-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="26301-108">コードエディターの作業ウィンドウで、[ **新しいスクリプト** ] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="26301-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="26301-109">スクリプト全体を、選択したサンプルに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="26301-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="26301-110">コードエディターの作業ウィンドウで、[ **実行** ] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="26301-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="26301-111">スクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="26301-111">Scripting basics</span></span>

<span data-ttu-id="26301-112">これらのサンプルでは、Office スクリプトの基本的な構成要素を示します。</span><span class="sxs-lookup"><span data-stu-id="26301-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="26301-113">これらをスクリプトに追加して、ソリューションを拡張し、一般的な問題を解決します。</span><span class="sxs-lookup"><span data-stu-id="26301-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="26301-114">1つのセルを読み取り、ログに記録する</span><span class="sxs-lookup"><span data-stu-id="26301-114">Read and log one cell</span></span>

<span data-ttu-id="26301-115">この例では、 **A1** の値を読み取り、コンソールに出力します。</span><span class="sxs-lookup"><span data-stu-id="26301-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="26301-116">アクティブセルを読み取る</span><span class="sxs-lookup"><span data-stu-id="26301-116">Read the active cell</span></span>

<span data-ttu-id="26301-117">このスクリプトは、現在アクティブなセルの値を記録します。</span><span class="sxs-lookup"><span data-stu-id="26301-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="26301-118">複数のセルが選択されている場合は、一番左側のセルがログに記録されます。</span><span class="sxs-lookup"><span data-stu-id="26301-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="26301-119">隣接するセルを変更する</span><span class="sxs-lookup"><span data-stu-id="26301-119">Change an adjacent cell</span></span>

<span data-ttu-id="26301-120">このスクリプトは、相対参照を使用して隣接するセルを取得します。</span><span class="sxs-lookup"><span data-stu-id="26301-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="26301-121">アクティブセルが一番上の行にある場合は、現在選択されているセルを参照しているため、スクリプトの一部が失敗することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="26301-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="26301-122">隣接するすべてのセルを変更する</span><span class="sxs-lookup"><span data-stu-id="26301-122">Change all adjacent cells</span></span>

<span data-ttu-id="26301-123">このスクリプトは、アクティブセルの書式を隣接するセルにコピーします。</span><span class="sxs-lookup"><span data-stu-id="26301-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="26301-124">このスクリプトは、アクティブセルがワークシートの端にない場合にのみ機能することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="26301-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="26301-125">範囲内の各セルを変更する</span><span class="sxs-lookup"><span data-stu-id="26301-125">Change each individual cell in a range</span></span>

<span data-ttu-id="26301-126">このスクリプトは、現在の選択範囲をループします。</span><span class="sxs-lookup"><span data-stu-id="26301-126">This script loops over the currently select range.</span></span> <span data-ttu-id="26301-127">現在の書式をクリアし、各セルの塗りつぶしの色をランダムな色に設定します。</span><span class="sxs-lookup"><span data-stu-id="26301-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
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

## <a name="collections"></a><span data-ttu-id="26301-128">コレクション</span><span class="sxs-lookup"><span data-stu-id="26301-128">Collections</span></span>

<span data-ttu-id="26301-129">これらのサンプルは、ブック内のオブジェクトのコレクションに対して機能します。</span><span class="sxs-lookup"><span data-stu-id="26301-129">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="26301-130">コレクションの反復処理</span><span class="sxs-lookup"><span data-stu-id="26301-130">Iterating over collections</span></span>

<span data-ttu-id="26301-131">このスクリプトは、ブック内のすべてのワークシートの名前を取得してログ記録します。</span><span class="sxs-lookup"><span data-stu-id="26301-131">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="26301-132">また、タブの色をランダムな色に設定します。</span><span class="sxs-lookup"><span data-stu-id="26301-132">It also sets the their tab colors to a random color.</span></span>

```typescript
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

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="26301-133">コレクションを照会および削除する</span><span class="sxs-lookup"><span data-stu-id="26301-133">Querying and deleting from a collection</span></span>

<span data-ttu-id="26301-134">このスクリプトは、新しいワークシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="26301-134">This script creates a new worksheet.</span></span> <span data-ttu-id="26301-135">ワークシートの既存のコピーがあるかどうかを確認し、新しいシートを作成する前に削除します。</span><span class="sxs-lookup"><span data-stu-id="26301-135">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
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

## <a name="dates"></a><span data-ttu-id="26301-136">日付</span><span class="sxs-lookup"><span data-stu-id="26301-136">Dates</span></span>

<span data-ttu-id="26301-137">このセクションのサンプルは、JavaScript の [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) オブジェクトを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="26301-137">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="26301-138">次の例では、現在の日付と時刻を取得し、アクティブなワークシート内の2つのセルにこれらの値を書き込みます。</span><span class="sxs-lookup"><span data-stu-id="26301-138">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="26301-139">次の例では、Excel に保存されている日付を読み取って、JavaScript の Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="26301-139">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="26301-140">[日付のシリアル番号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)は、JavaScript 日付の入力として使用されます。</span><span class="sxs-lookup"><span data-stu-id="26301-140">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="26301-141">データの表示</span><span class="sxs-lookup"><span data-stu-id="26301-141">Display data</span></span>

<span data-ttu-id="26301-142">これらのサンプルは、ワークシートデータを操作し、ユーザーにより良い表示や組織を提供する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="26301-142">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="26301-143">条件付き書式の適用</span><span class="sxs-lookup"><span data-stu-id="26301-143">Apply conditional formatting</span></span>

<span data-ttu-id="26301-144">この例では、ワークシートで現在使用されている範囲に条件付き書式を適用します。</span><span class="sxs-lookup"><span data-stu-id="26301-144">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="26301-145">条件付き書式は、値の上位10% の緑の塗りつぶしです。</span><span class="sxs-lookup"><span data-stu-id="26301-145">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="26301-146">並べ替えられたテーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="26301-146">Create a sorted table</span></span>

<span data-ttu-id="26301-147">次の使用例は、現在のワークシートの使用範囲から表を作成し、最初の列に基づいて並べ替えます。</span><span class="sxs-lookup"><span data-stu-id="26301-147">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="26301-148">ピボットテーブルから "総計" 値を記録する</span><span class="sxs-lookup"><span data-stu-id="26301-148">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="26301-149">次の例では、ブックの最初のピボットテーブルを検索し、次の図のように、[総計] セル (緑で強調表示されている) に値を記録します。</span><span class="sxs-lookup"><span data-stu-id="26301-149">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![総計行が緑色で強調表示された果物 sales ピボットテーブル。](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="formulas"></a><span data-ttu-id="26301-151">式</span><span class="sxs-lookup"><span data-stu-id="26301-151">Formulas</span></span>

<span data-ttu-id="26301-152">これらのサンプルでは、Excel の数式を使用して、スクリプト内でそれらを操作する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="26301-152">These samples use Excel formulas and show how to work with them in scripts.</span></span>

## <a name="single-formula"></a><span data-ttu-id="26301-153">単一の数式</span><span class="sxs-lookup"><span data-stu-id="26301-153">Single formula</span></span>

<span data-ttu-id="26301-154">次のスクリプトは、セルの数式を設定し、Excel がセルの数式と値を個別に格納する方法を表示します。</span><span class="sxs-lookup"><span data-stu-id="26301-154">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
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

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="26301-155">数式からの結果を Spilling する</span><span class="sxs-lookup"><span data-stu-id="26301-155">Spilling results from a formula</span></span>

<span data-ttu-id="26301-156">このスクリプトは、転置関数を使用して、範囲 "A1: D2" を "A4: B7" に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="26301-156">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="26301-157">転置した結果、#SPILL エラーが発生した場合は、対象範囲をクリアし、数式を再度適用します。</span><span class="sxs-lookup"><span data-stu-id="26301-157">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
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

## <a name="scenario-samples"></a><span data-ttu-id="26301-158">シナリオサンプル</span><span class="sxs-lookup"><span data-stu-id="26301-158">Scenario samples</span></span>

<span data-ttu-id="26301-159">大規模な現実世界のソリューションを紹介するサンプルについては、「 [Office スクリプトのサンプルシナリオ](scenarios/sample-scenario-overview.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="26301-159">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="26301-160">新しいサンプルを提案する</span><span class="sxs-lookup"><span data-stu-id="26301-160">Suggest new samples</span></span>

<span data-ttu-id="26301-161">新しいサンプルの提案を歓迎します。</span><span class="sxs-lookup"><span data-stu-id="26301-161">We welcome suggestions for new samples.</span></span> <span data-ttu-id="26301-162">他のスクリプト開発者を支援する一般的なシナリオがある場合は、以下のフィードバックセクションでご連絡ください。</span><span class="sxs-lookup"><span data-stu-id="26301-162">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
