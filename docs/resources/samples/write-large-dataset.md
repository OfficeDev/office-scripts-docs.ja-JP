---
title: 大規模なデータセットを記述する場合のパフォーマンスの最適化
description: 大規模なデータセットをスクリプトに記述するときにパフォーマンスを最適化するOfficeします。
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 190072e58238be95a2939f73dcda077ed91db848
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571495"
---
# <a name="performance-optimization-when-writing-a-large-dataset"></a><span data-ttu-id="e051e-103">大規模なデータセットを記述する場合のパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="e051e-103">Performance optimization when writing a large dataset</span></span>

## <a name="basic-performance-optimization"></a><span data-ttu-id="e051e-104">基本的なパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="e051e-104">Basic performance optimization</span></span>

<span data-ttu-id="e051e-105">スクリプトのパフォーマンスの基本Office、「Getting Started」[](getting-started.md#basic-performance-considerations)の記事の「パフォーマンス」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e051e-105">For performance basics in Office Scripts, see the [performance section](getting-started.md#basic-performance-considerations) of the Getting Started article.</span></span>

## <a name="sample-code-optimize-performance-of-a-large-dataset"></a><span data-ttu-id="e051e-106">サンプル コード: 大規模なデータセットのパフォーマンスを最適化する</span><span class="sxs-lookup"><span data-stu-id="e051e-106">Sample code: Optimize performance of a large dataset</span></span>

<span data-ttu-id="e051e-107">Range `setValues()` API では、範囲の値を設定できます。</span><span class="sxs-lookup"><span data-stu-id="e051e-107">The `setValues()` Range API allows setting the values of a range.</span></span> <span data-ttu-id="e051e-108">この API には、データ サイズ、ネットワーク設定など、さまざまな要因に応じてデータの制限があります。大量のデータを確実に更新するには、より小さなチャンクでデータ更新を行う方法を考える必要があります。</span><span class="sxs-lookup"><span data-stu-id="e051e-108">This API has data limitations depending on various factors such as data size, network settings, etc. In order to reliably update a large range of data, you'll need to think about doing data updates in smaller chunks.</span></span> <span data-ttu-id="e051e-109">このスクリプトはこれを実行し、範囲の行をチャンク単位で書き込み、大きな範囲を更新する必要がある場合は、小さな部分で行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="e051e-109">This script attempts to do this and writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller parts.</span></span> <span data-ttu-id="e051e-110">**警告**: さまざまなサイズでテストされていないので、スクリプトでこれを使用する場合は注意してください。</span><span class="sxs-lookup"><span data-stu-id="e051e-110">**Warning**: It has not been tested across various sizes so be aware of that if you want to use this in your script.</span></span> <span data-ttu-id="e051e-111">テストの機会が得たので、さまざまなデータ サイズに対するパフォーマンスに関する結果を更新します。</span><span class="sxs-lookup"><span data-stu-id="e051e-111">As we have opportunity to test, we'll update with findings around how it performs for various data sizes.</span></span>

<span data-ttu-id="e051e-112">このスクリプトはチャンクごとに 1K セルを選択しますが、上書きして、その動作をテストできます。</span><span class="sxs-lookup"><span data-stu-id="e051e-112">This script selects 1K cells per chunk but you can override to test out how it works for you.</span></span> <span data-ttu-id="e051e-113">6 列のデータを含む 100k 行を更新します。</span><span class="sxs-lookup"><span data-stu-id="e051e-113">It updates 100k rows with 6 columns of data.</span></span> <span data-ttu-id="e051e-114">空白のシートでこれを実行して調べてください。</span><span class="sxs-lookup"><span data-stu-id="e051e-114">Run this on a blank sheet to examine.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  let data: (string | number | boolean)[][] = [];
  // Number of rows in the random data (x 6 columns).
  const sampleRows = 100000;

  console.log(`Generating data...`)
  // Dynamically generate some random data for testing purpose. 
  for (let i = 0; i < sampleRows; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInChunks(sheet.getRange("B2"), data);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`)
  }

  return;
}

function updateRangeInChunks(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInChunk: number = 10000
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per chunk setting: ${cellsInChunk}`);
  if (!values) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0 || values[0].length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log(`No need to chunk -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
  console.log("Rows per chunk: " + rowsPerChunk);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let chunkCount = 0;

  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerChunk) {
      chunkCount++;
      console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);

    }
  }
  console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInChunk).toFixed(8)} seconds per ${cellsInChunk} cells-chunk.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */

function updateNextChunk(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerChunk: number,
  totalRowsUpdated: number
) {

  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

## <a name="training-video-optimize-performance-when-writing-a-large-dataset"></a><span data-ttu-id="e051e-115">トレーニング ビデオ: 大規模なデータセットを記述するときにパフォーマンスを最適化する</span><span class="sxs-lookup"><span data-stu-id="e051e-115">Training video: Optimize performance when writing a large dataset</span></span>

<span data-ttu-id="e051e-116">[![大規模なデータセットを記述するときにパフォーマンスを最適化する方法に関するビデオを見る](../../images/largedata-vid.png)](https://youtu.be/BP9Kp0Ltj7U "大規模なデータセットを記述するときにパフォーマンスを最適化する方法に関するビデオ")</span><span class="sxs-lookup"><span data-stu-id="e051e-116">[![Watch video on how to optimize performance when writing a large dataset](../../images/largedata-vid.png)](https://youtu.be/BP9Kp0Ltj7U "Video on how to optimize performance when writing a large dataset")</span></span>
