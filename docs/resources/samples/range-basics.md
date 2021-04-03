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
# <a name="range-basics"></a>範囲の基本

`Range` は、スクリプト Excel オブジェクト Office内の基礎オブジェクトです。 [範囲 API を使用](/javascript/api/office-scripts/excelscript/excelscript.range) すると、グリッドで使用できるデータと形式の両方にアクセスし、ワークシート、テーブル、グラフなどの Excel 内の他の主要オブジェクトをリンクできます。

範囲は、"A1:B4" などのアドレスを使用するか、指定されたセル セットの名前付きキーである名前付きアイテムを使用して識別されます。 Excel オブジェクト モデルでは、セルとセルのグループの両方を範囲 と呼 _ばれます_。 `Range` セル内のデータなどのセル レベルの属性や、セルレベルの属性 (書式、罫線など) を含めることもできます。

`Range` また、少なくとも 1 つのセルで構成されるユーザーの選択を介して取得できます。 範囲を操作する際には、セルと範囲の関係を明確に保つ必要があります。

スクリプトで最も頻繁に使用されるゲッター、セッター、その他の便利なメソッドのコア セットを次に示します。 これは、API ジャーニーの開始点として最適です。 以降のセクションでは、メソッドをグループ化し、オブジェクトの API のロックを解除し始めるに当たって、メンタル モデルの構築 `Range` に役立ちます。

## <a name="example-scripts"></a>スクリプトの例

* [基本的な読み取りおよび書き込み](#basic-read-and-write)
* [ワークシートの最後に行を追加する](#add-row-at-the-end-of-worksheet)
* [列フィルターのクリア](clear-table-filter-for-active-cell.md)
* [一意の色で各セルに色を付け](#color-each-cell-with-unique-color)
* [2 次元 (2D) 配列を使用して値を使用して範囲を更新する](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>基本的な読み取りおよび書き込み

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

### <a name="add-row-at-the-end-of-worksheet"></a>ワークシートの最後に行を追加する

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

### <a name="color-each-cell-with-unique-color"></a>一意の色で各セルに色を付け

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

### <a name="update-range-with-values-using-2d-array"></a>2D 配列を使用して値を使用して範囲を更新する

2D 配列の値に基づいて更新する範囲ディメンションを動的に計算します。

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

## <a name="training-videos-range-basics"></a>トレーニング ビデオ: 範囲の基本

_範囲の基本_

[![Range の基本に関するステップバイステップのビデオを見る](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Range の基本に関するステップ バイ ステップ のビデオ")

_ワークシートの最後に行を追加する_

[![ワークシートの最後に行を追加する方法について、ステップバイステップのビデオを見る](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "ワークシートの最後に行を追加する方法に関するステップバイステップのビデオ")

## <a name="methods-that-return-some-range-metadata"></a>範囲メタデータを返すメソッド

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>指定した範囲に関連付けられたデータ/定数を返すメソッド

### <a name="returned-as-single-cell-value"></a>単一セル値として返される

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### <a name="returned-as-2d-arrays-whole-range"></a>2D 配列として返される (範囲全体)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## <a name="methods-that-return-other-range-object"></a>他の範囲オブジェクトを返すメソッド

* getSurroundingRegion() -- VBA の CurrentRegion に似ている
* getCell(row, column)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

**重要/興味深い**

* _workbook_.getSelectedRange()
* _workbook_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>別の範囲オブジェクトとの関連で範囲オブジェクトを返すメソッド

* getBoundingRect(anotherRange)
* getIntersection(anotherRange)

## <a name="methods-that-return-other-objects-non-range-objects"></a>他のオブジェクト (範囲以外のオブジェクト) を返すメソッド

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## <a name="set-methods"></a>Set メソッド

### <a name="singular-cell-set-methods"></a>単数形のセル セット メソッド

* setFormula(formula)
* setFormulaLocal(formulaLocal)
* setFormulaR1C1(formulaR1C1)
* setNumberFormatLocal(numberFormatLocal)
* setValue(value)

### <a name="2d--entire-range-set-methods"></a>2D /範囲全体の設定方法

* setFormulas(formulas)
* setFormulasLocal(formulasLocal)
* setFormulasR1C1(formulasR1C1)
* setNumberFormat(numberFormat)
* setNumberFormats(numberFormats)
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## <a name="other-methods"></a>その他の方法

* merge(across)
* unmerge()

## <a name="coming-soon"></a>近日対応予定

* 範囲エッジ API
