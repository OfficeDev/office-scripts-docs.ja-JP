---
title: Office スクリプトの範囲の基本
description: '[スクリプト] で Range オブジェクトを使用するOffice説明します。'
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571303"
---
# <a name="range-basics"></a><span data-ttu-id="1daca-103">範囲の基本</span><span class="sxs-lookup"><span data-stu-id="1daca-103">Range basics</span></span>

<span data-ttu-id="1daca-104">`Range` は、スクリプト Excel オブジェクト Office内の基礎オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="1daca-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="1daca-105">[範囲 API を使用](/javascript/api/office-scripts/excelscript/excelscript.range) すると、グリッドで使用できるデータと形式の両方にアクセスし、ワークシート、テーブル、グラフなどの Excel 内の他の主要オブジェクトをリンクできます。</span><span class="sxs-lookup"><span data-stu-id="1daca-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="1daca-106">範囲は、"A1:B4" などのアドレスを使用するか、指定されたセル セットの名前付きキーである名前付きアイテムを使用して識別されます。</span><span class="sxs-lookup"><span data-stu-id="1daca-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="1daca-107">Excel オブジェクト モデルでは、セルとセルのグループの両方を範囲 と呼 _ばれます_。</span><span class="sxs-lookup"><span data-stu-id="1daca-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="1daca-108">`Range` セル内のデータなどのセル レベルの属性や、セルレベルの属性 (書式、罫線など) を含めることもできます。</span><span class="sxs-lookup"><span data-stu-id="1daca-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="1daca-109">`Range` また、少なくとも 1 つのセルで構成されるユーザーの選択を介して取得できます。</span><span class="sxs-lookup"><span data-stu-id="1daca-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="1daca-110">範囲を操作する際には、セルと範囲の関係を明確に保つ必要があります。</span><span class="sxs-lookup"><span data-stu-id="1daca-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="1daca-111">スクリプトで最も頻繁に使用されるゲッター、セッター、その他の便利なメソッドのコア セットを次に示します。</span><span class="sxs-lookup"><span data-stu-id="1daca-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="1daca-112">これは、API ジャーニーの開始点として最適です。</span><span class="sxs-lookup"><span data-stu-id="1daca-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="1daca-113">以降のセクションでは、メソッドをグループ化し、オブジェクトの API のロックを解除し始めるに当たって、メンタル モデルの構築 `Range` に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="1daca-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="1daca-114">スクリプトの例</span><span class="sxs-lookup"><span data-stu-id="1daca-114">Example scripts</span></span>

* [<span data-ttu-id="1daca-115">基本的な読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="1daca-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="1daca-116">ワークシートの最後に行を追加する</span><span class="sxs-lookup"><span data-stu-id="1daca-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="1daca-117">列フィルターのクリア</span><span class="sxs-lookup"><span data-stu-id="1daca-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="1daca-118">一意の色で各セルに色を付け</span><span class="sxs-lookup"><span data-stu-id="1daca-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="1daca-119">2 次元 (2D) 配列を使用して値を使用して範囲を更新する</span><span class="sxs-lookup"><span data-stu-id="1daca-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="1daca-120">基本的な読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="1daca-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="1daca-121">ワークシートの最後に行を追加する</span><span class="sxs-lookup"><span data-stu-id="1daca-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="1daca-122">一意の色で各セルに色を付け</span><span class="sxs-lookup"><span data-stu-id="1daca-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="1daca-123">2D 配列を使用して値を使用して範囲を更新する</span><span class="sxs-lookup"><span data-stu-id="1daca-123">Update range with values using 2D array</span></span>

<span data-ttu-id="1daca-124">2D 配列の値に基づいて更新する範囲ディメンションを動的に計算します。</span><span class="sxs-lookup"><span data-stu-id="1daca-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="1daca-125">トレーニング ビデオ: 範囲の基本</span><span class="sxs-lookup"><span data-stu-id="1daca-125">Training videos: Range basics</span></span>

<span data-ttu-id="1daca-126">_範囲の基本_</span><span class="sxs-lookup"><span data-stu-id="1daca-126">_Range basics_</span></span>

<span data-ttu-id="1daca-127">[![Range の基本に関するステップバイステップのビデオを見る](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Range の基本に関するステップ バイ ステップ のビデオ")</span><span class="sxs-lookup"><span data-stu-id="1daca-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="1daca-128">_ワークシートの最後に行を追加する_</span><span class="sxs-lookup"><span data-stu-id="1daca-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="1daca-129">[![ワークシートの最後に行を追加する方法について、ステップバイステップのビデオを見る](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "ワークシートの最後に行を追加する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="1daca-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="1daca-130">範囲メタデータを返すメソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="1daca-131">getAddress(), getAddressLocal()</span><span class="sxs-lookup"><span data-stu-id="1daca-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="1daca-132">getCellCount()</span><span class="sxs-lookup"><span data-stu-id="1daca-132">getCellCount()</span></span>
* <span data-ttu-id="1daca-133">getRowCount(), getColumnCount()</span><span class="sxs-lookup"><span data-stu-id="1daca-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="1daca-134">指定した範囲に関連付けられたデータ/定数を返すメソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="1daca-135">単一セル値として返される</span><span class="sxs-lookup"><span data-stu-id="1daca-135">Returned as single cell value</span></span>

* <span data-ttu-id="1daca-136">getFormula(), getFormulaLocal()</span><span class="sxs-lookup"><span data-stu-id="1daca-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="1daca-137">getFormulaR1C1()</span><span class="sxs-lookup"><span data-stu-id="1daca-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="1daca-138">getNumberFormat(), getNumberFormatLocal()</span><span class="sxs-lookup"><span data-stu-id="1daca-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="1daca-139">getText()</span><span class="sxs-lookup"><span data-stu-id="1daca-139">getText()</span></span>
* <span data-ttu-id="1daca-140">getValue()</span><span class="sxs-lookup"><span data-stu-id="1daca-140">getValue()</span></span>
* <span data-ttu-id="1daca-141">getValueType()</span><span class="sxs-lookup"><span data-stu-id="1daca-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="1daca-142">2D 配列として返される (範囲全体)</span><span class="sxs-lookup"><span data-stu-id="1daca-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="1daca-143">getFormulas(), getFormulasLocal()</span><span class="sxs-lookup"><span data-stu-id="1daca-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="1daca-144">getFormulasR1C1()</span><span class="sxs-lookup"><span data-stu-id="1daca-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="1daca-145">getNumberFormatCategories()</span><span class="sxs-lookup"><span data-stu-id="1daca-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="1daca-146">getNumberFormats(), getNumberFormatsLocal()</span><span class="sxs-lookup"><span data-stu-id="1daca-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="1daca-147">getTexts()</span><span class="sxs-lookup"><span data-stu-id="1daca-147">getTexts()</span></span>
* <span data-ttu-id="1daca-148">getValues()</span><span class="sxs-lookup"><span data-stu-id="1daca-148">getValues()</span></span>
* <span data-ttu-id="1daca-149">getValueTypes()</span><span class="sxs-lookup"><span data-stu-id="1daca-149">getValueTypes()</span></span>
* <span data-ttu-id="1daca-150">getHidden()</span><span class="sxs-lookup"><span data-stu-id="1daca-150">getHidden()</span></span>
* <span data-ttu-id="1daca-151">getIsEntireRow()</span><span class="sxs-lookup"><span data-stu-id="1daca-151">getIsEntireRow()</span></span>
* <span data-ttu-id="1daca-152">getIsEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="1daca-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="1daca-153">他の範囲オブジェクトを返すメソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-153">Methods that return other range object</span></span>

* <span data-ttu-id="1daca-154">getSurroundingRegion() -- VBA の CurrentRegion に似ている</span><span class="sxs-lookup"><span data-stu-id="1daca-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="1daca-155">getCell(row, column)</span><span class="sxs-lookup"><span data-stu-id="1daca-155">getCell(row, column)</span></span>
* <span data-ttu-id="1daca-156">getColumn(column)</span><span class="sxs-lookup"><span data-stu-id="1daca-156">getColumn(column)</span></span>
* <span data-ttu-id="1daca-157">getColumnHidden()</span><span class="sxs-lookup"><span data-stu-id="1daca-157">getColumnHidden()</span></span>
* <span data-ttu-id="1daca-158">getColumnsAfter(count)</span><span class="sxs-lookup"><span data-stu-id="1daca-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="1daca-159">getColumnsBefore(count)</span><span class="sxs-lookup"><span data-stu-id="1daca-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="1daca-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="1daca-160">getEntireColumn()</span></span>
* <span data-ttu-id="1daca-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="1daca-161">getEntireRow()</span></span>
* <span data-ttu-id="1daca-162">getLastCell()</span><span class="sxs-lookup"><span data-stu-id="1daca-162">getLastCell()</span></span>
* <span data-ttu-id="1daca-163">getLastColumn()</span><span class="sxs-lookup"><span data-stu-id="1daca-163">getLastColumn()</span></span>
* <span data-ttu-id="1daca-164">getLastRow()</span><span class="sxs-lookup"><span data-stu-id="1daca-164">getLastRow()</span></span>
* <span data-ttu-id="1daca-165">getRow(row)</span><span class="sxs-lookup"><span data-stu-id="1daca-165">getRow(row)</span></span>
* <span data-ttu-id="1daca-166">getRowHidden()</span><span class="sxs-lookup"><span data-stu-id="1daca-166">getRowHidden()</span></span>
* <span data-ttu-id="1daca-167">getRowsAbove(count)</span><span class="sxs-lookup"><span data-stu-id="1daca-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="1daca-168">getRowsBelow(count)</span><span class="sxs-lookup"><span data-stu-id="1daca-168">getRowsBelow(count)</span></span>

<span data-ttu-id="1daca-169">**重要/興味深い**</span><span class="sxs-lookup"><span data-stu-id="1daca-169">**Important/Interesting**</span></span>

* <span data-ttu-id="1daca-170">_workbook_.getSelectedRange()</span><span class="sxs-lookup"><span data-stu-id="1daca-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="1daca-171">_workbook_.getActiveCell()</span><span class="sxs-lookup"><span data-stu-id="1daca-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="1daca-172">getUsedRange(valuesOnly)</span><span class="sxs-lookup"><span data-stu-id="1daca-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="1daca-173">getAbsoluteResizedRange(numRows, numColumns)</span><span class="sxs-lookup"><span data-stu-id="1daca-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="1daca-174">getOffsetRange(rowOffset, columnOffset)</span><span class="sxs-lookup"><span data-stu-id="1daca-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="1daca-175">getResizedRange(deltaRows, deltaColumns)</span><span class="sxs-lookup"><span data-stu-id="1daca-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="1daca-176">別の範囲オブジェクトとの関連で範囲オブジェクトを返すメソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="1daca-177">getBoundingRect(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="1daca-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="1daca-178">getIntersection(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="1daca-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="1daca-179">他のオブジェクト (範囲以外のオブジェクト) を返すメソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="1daca-180">getDirectPrecedents()</span><span class="sxs-lookup"><span data-stu-id="1daca-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="1daca-181">getWorksheet()</span><span class="sxs-lookup"><span data-stu-id="1daca-181">getWorksheet()</span></span>
* <span data-ttu-id="1daca-182">getTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="1daca-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="1daca-183">getPivotTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="1daca-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="1daca-184">getDataValidation()</span><span class="sxs-lookup"><span data-stu-id="1daca-184">getDataValidation()</span></span>
* <span data-ttu-id="1daca-185">getPredefinedCellStyle()</span><span class="sxs-lookup"><span data-stu-id="1daca-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="1daca-186">Set メソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="1daca-187">単数形のセル セット メソッド</span><span class="sxs-lookup"><span data-stu-id="1daca-187">Singular cell set methods</span></span>

* <span data-ttu-id="1daca-188">setFormula(formula)</span><span class="sxs-lookup"><span data-stu-id="1daca-188">setFormula(formula)</span></span>
* <span data-ttu-id="1daca-189">setFormulaLocal(formulaLocal)</span><span class="sxs-lookup"><span data-stu-id="1daca-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="1daca-190">setFormulaR1C1(formulaR1C1)</span><span class="sxs-lookup"><span data-stu-id="1daca-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="1daca-191">setNumberFormatLocal(numberFormatLocal)</span><span class="sxs-lookup"><span data-stu-id="1daca-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="1daca-192">setValue(value)</span><span class="sxs-lookup"><span data-stu-id="1daca-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="1daca-193">2D /範囲全体の設定方法</span><span class="sxs-lookup"><span data-stu-id="1daca-193">2D / entire range set methods</span></span>

* <span data-ttu-id="1daca-194">setFormulas(formulas)</span><span class="sxs-lookup"><span data-stu-id="1daca-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="1daca-195">setFormulasLocal(formulasLocal)</span><span class="sxs-lookup"><span data-stu-id="1daca-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="1daca-196">setFormulasR1C1(formulasR1C1)</span><span class="sxs-lookup"><span data-stu-id="1daca-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="1daca-197">setNumberFormat(numberFormat)</span><span class="sxs-lookup"><span data-stu-id="1daca-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="1daca-198">setNumberFormats(numberFormats)</span><span class="sxs-lookup"><span data-stu-id="1daca-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="1daca-199">setNumberFormatsLocal(numberFormatsLocal)</span><span class="sxs-lookup"><span data-stu-id="1daca-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="1daca-200">setValues(values)</span><span class="sxs-lookup"><span data-stu-id="1daca-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="1daca-201">その他の方法</span><span class="sxs-lookup"><span data-stu-id="1daca-201">Other methods</span></span>

* <span data-ttu-id="1daca-202">merge(across)</span><span class="sxs-lookup"><span data-stu-id="1daca-202">merge(across)</span></span>
* <span data-ttu-id="1daca-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="1daca-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="1daca-204">近日対応予定</span><span class="sxs-lookup"><span data-stu-id="1daca-204">Coming soon</span></span>

* <span data-ttu-id="1daca-205">範囲エッジ API</span><span class="sxs-lookup"><span data-stu-id="1daca-205">Range edge APIs</span></span>
