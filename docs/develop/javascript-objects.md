---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Web 上の Excel で Office スクリプトから組み込みの JavaScript Api を呼び出す方法について説明します。
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 1c8ac757574e8c4be64b373f8d4bf421ddfa0c79
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999261"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="f04c4-103">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="f04c4-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="f04c4-104">Javascript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (javascript のスーパーセット) のどちらでスクリプトを作成するかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="f04c4-105">この記事では、web 上の Excel の Office スクリプトに組み込まれている JavaScript オブジェクトのいくつかを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="f04c4-106">すべての組み込み JavaScript オブジェクトの完全なリストについては、「Mozilla の[標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)」記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f04c4-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="f04c4-107">配列</span><span class="sxs-lookup"><span data-stu-id="f04c4-107">Array</span></span>

<span data-ttu-id="f04c4-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を操作するための標準化された方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="f04c4-109">配列は標準的な JavaScript コンストラクトですが、範囲とコレクションという2つの主な方法で Office スクリプトに関連しています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="f04c4-110">範囲を使用して作業する</span><span class="sxs-lookup"><span data-stu-id="f04c4-110">Working with ranges</span></span>

<span data-ttu-id="f04c4-111">範囲には、その範囲内のセルに直接マップされるいくつかの2次元配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="f04c4-112">これらの配列には、その範囲内の各セルに関する特定の情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="f04c4-113">たとえば、は、 `Range.getValues` 2 次元配列の行と列がそのワークシートサブセクションの行と列にマッピングされているセルのすべての値を返します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="f04c4-114">`Range.getFormulas`また、 `Range.getNumberFormats` のように配列を返すその他のメソッドもよく使用され `Range.getValues` ます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="f04c4-115">次のスクリプトは、"$" が含まれている任意の番号書式の**A1: D4**範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="f04c4-116">このスクリプトは、これらのセルの塗りつぶしの色を "黄" に設定します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-116">The script sets the fill color in those cells to "yellow".</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a><span data-ttu-id="f04c4-117">コレクションを処理する</span><span class="sxs-lookup"><span data-stu-id="f04c4-117">Working with collections</span></span>

<span data-ttu-id="f04c4-118">多くの Excel オブジェクトがコレクションに含まれています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="f04c4-119">コレクションは Office スクリプト API によって管理され、配列として公開されます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="f04c4-120">たとえば、ワークシート内のすべての[図形](/javascript/api/office-scripts/excelscript/excelscript.shape)は、 `Shape[]` メソッドによって返されるに含まれてい `Worksheet.getShapes` ます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="f04c4-121">この配列を使用して、コレクションから値を取得したり、親オブジェクトのメソッドから特定のオブジェクトにアクセスしたりでき `get*` ます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="f04c4-122">これらのコレクションの配列に対してオブジェクトを手動で追加または削除しないでください。</span><span class="sxs-lookup"><span data-stu-id="f04c4-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="f04c4-123">`add`親オブジェクトのメソッドと、コレクション型のオブジェクトのメソッドを使用し `delete` ます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="f04c4-124">たとえば、メソッドを使用して[ワークシート](/javascript/api/office-scripts/excelscript/excelscript.worksheet)に[テーブル](/javascript/api/office-scripts/excelscript/excelscript.table)を追加 `Worksheet.addTable` し、 `Table` using を削除し `Table.delete` ます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="f04c4-125">次のスクリプトは、現在のワークシート内のすべての図形の種類を記録します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-125">The following script logs the type of every shape in the current worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

<span data-ttu-id="f04c4-126">次のスクリプトは、現在のワークシート内の最も古い図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-126">The following script deletes the oldest shape in the current worksheet.</span></span>

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a><span data-ttu-id="f04c4-127">日付</span><span class="sxs-lookup"><span data-stu-id="f04c4-127">Date</span></span>

<span data-ttu-id="f04c4-128">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトには、スクリプト内の日付を処理するための標準化された方法が用意されています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="f04c4-129">`Date.now()`現在の日付と時刻を使用してオブジェクトを生成します。これは、スクリプトのデータ入力にタイムスタンプを追加するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="f04c4-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="f04c4-130">次のスクリプトは、現在の日付をワークシートに追加します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="f04c4-131">このメソッドを使用する `toLocaleDateString` と、Excel によって値が日付として認識され、セルの数値の書式が自動的に変更されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f04c4-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

<span data-ttu-id="f04c4-132">サンプルの [[日付を使用して作業](../resources/excel-samples.md#work-with-dates)] セクションには、さらに多くの日付関連スクリプトがあります。</span><span class="sxs-lookup"><span data-stu-id="f04c4-132">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="f04c4-133">数学</span><span class="sxs-lookup"><span data-stu-id="f04c4-133">Math</span></span>

<span data-ttu-id="f04c4-134">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトには、一般的な数値演算のためのメソッドと定数が用意されています。</span><span class="sxs-lookup"><span data-stu-id="f04c4-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="f04c4-135">これらは、ブックの計算エンジンを使用しなくても、Excel で使用できる多くの関数を提供します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="f04c4-136">これにより、スクリプトはブックを照会する必要がなくなり、パフォーマンスが向上します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="f04c4-137">次のスクリプトは、を使用して、 `Math.min` **A1: D4**範囲の最小数を検索して記録します。</span><span class="sxs-lookup"><span data-stu-id="f04c4-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="f04c4-138">この例では、範囲全体に文字列ではなく数値のみが含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f04c4-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="f04c4-139">外部 JavaScript ライブラリの使用はサポートされていません</span><span class="sxs-lookup"><span data-stu-id="f04c4-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="f04c4-140">Office スクリプトは、外部のサードパーティ製ライブラリの使用をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="f04c4-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="f04c4-141">スクリプトでは、組み込みの JavaScript オブジェクトと Office スクリプト Api のみを使用できます。</span><span class="sxs-lookup"><span data-stu-id="f04c4-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="f04c4-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="f04c4-142">See also</span></span>

- [<span data-ttu-id="f04c4-143">標準の組み込みオブジェクト</span><span class="sxs-lookup"><span data-stu-id="f04c4-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="f04c4-144">Office スクリプトのコードエディター環境</span><span class="sxs-lookup"><span data-stu-id="f04c4-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
