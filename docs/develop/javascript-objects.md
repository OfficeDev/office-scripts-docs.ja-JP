---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: 組み込みの JavaScript API を、Officeスクリプトから呼び出Excel on the web。
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: e3b36265f235678eee18fbf369058b165da46210
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232404"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="4f2dd-103">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="4f2dd-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="4f2dd-104">JavaScript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (JavaScript のスーパーセット) でスクリプトを実行するかどうかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが提供されています。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="4f2dd-105">この記事では、スクリプトで組み込みの JavaScript オブジェクトの一部を使用Office説明Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="4f2dd-106">すべての組み込み JavaScript オブジェクトの完全な一覧については、Mozilla の Standard 組み込 [みオブジェクトの記事を参照](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) してください。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="4f2dd-107">配列</span><span class="sxs-lookup"><span data-stu-id="4f2dd-107">Array</span></span>

<span data-ttu-id="4f2dd-108">[Array オブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)は、スクリプト内の配列を扱う標準化された方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="4f2dd-109">配列は標準的な JavaScript コンストラクトですが、Officeとコレクションの 2 つの主要な方法で関連付けされます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="4f2dd-110">範囲の操作</span><span class="sxs-lookup"><span data-stu-id="4f2dd-110">Working with ranges</span></span>

<span data-ttu-id="4f2dd-111">範囲には、その範囲内のセルに直接マップする複数の 2 次元配列が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="4f2dd-112">これらの配列には、その範囲内の各セルに関する特定の情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="4f2dd-113">たとえば、これらのセル内のすべての値を返します (2 次元配列の行と列は、そのワークシート のサブセクションの行と列 `Range.getValues` にマッピングされます)。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="4f2dd-114">`Range.getFormulas` などの `Range.getNumberFormats` 配列を返す他の頻繁に使用されるメソッドです `Range.getValues` 。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="4f2dd-115">次のスクリプトは **、A1:D4 範囲で** "$" を含む任意の数値形式を検索します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="4f2dd-116">スクリプトは、これらのセルの塗りつぶしの色を "黄色" に設定します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="4f2dd-117">コレクションの操作</span><span class="sxs-lookup"><span data-stu-id="4f2dd-117">Working with collections</span></span>

<span data-ttu-id="4f2dd-118">多Excelオブジェクトはコレクションに含まれています。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="4f2dd-119">このコレクションは、スクリプト API Officeによって管理され、配列として公開されます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="4f2dd-120">たとえば、ワークシート [内のすべての Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) は、メソッドによって返されるオブジェクト `Shape[]` に含 `Worksheet.getShapes` まれているとします。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="4f2dd-121">この配列を使用して、コレクションから値を読み取ることができます。また、親オブジェクトのメソッドから特定のオブジェクトに `get*` アクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="4f2dd-122">これらのコレクション配列からオブジェクトを手動で追加または削除しない。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="4f2dd-123">親オブジェクト `add` のメソッドとコレクション型オブジェクト `delete` のメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="4f2dd-124">たとえば、メソッドを使用して[ワークシートに Table](/javascript/api/office-scripts/excelscript/excelscript.table)を追加し[](/javascript/api/office-scripts/excelscript/excelscript.worksheet) `Worksheet.addTable` 、using を削除 `Table` します `Table.delete` 。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="4f2dd-125">次のスクリプトは、現在のワークシートのすべての図形の種類をログに記録します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="4f2dd-126">次のスクリプトは、現在のワークシートで最も古い図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="4f2dd-127">日付</span><span class="sxs-lookup"><span data-stu-id="4f2dd-127">Date</span></span>

<span data-ttu-id="4f2dd-128">[Date オブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)は、スクリプト内の日付を扱う標準化された方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="4f2dd-129">`Date.now()` 現在の日付と時刻を持つオブジェクトを生成します。これは、スクリプトのデータ エントリにタイムスタンプを追加するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="4f2dd-130">次のスクリプトは、現在の日付をワークシートに追加します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="4f2dd-131">メソッドを使用すると、Excel日付として認識され、セルの数値形式 `toLocaleDateString` が自動的に変更されます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="4f2dd-132">サンプル [の [日付の処理](../resources/samples/excel-samples.md#dates) ] セクションには、より多くの日付関連のスクリプトがあります。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="4f2dd-133">数学</span><span class="sxs-lookup"><span data-stu-id="4f2dd-133">Math</span></span>

<span data-ttu-id="4f2dd-134">[Math オブジェクトは](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)、一般的な数学演算のメソッドと定数を提供します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="4f2dd-135">ブックの計算エンジンを使用せずに、Excelで使用できる多くの機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="4f2dd-136">これにより、スクリプトがブックにクエリを実行する必要が省き、パフォーマンスが向上します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="4f2dd-137">次のスクリプトは `Math.min` **、A1:D4** 範囲の最小番号を検索してログに記録するために使用します。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="4f2dd-138">このサンプルでは、文字列ではなく、範囲全体に数値だけが含まれていると想定しています。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="4f2dd-139">外部 JavaScript ライブラリの使用はサポートされていません</span><span class="sxs-lookup"><span data-stu-id="4f2dd-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="4f2dd-140">Officeスクリプトは、外部のサード パーティ製ライブラリの使用をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="4f2dd-141">スクリプトでは、組み込みの JavaScript オブジェクトとスクリプト API Office使用できます。</span><span class="sxs-lookup"><span data-stu-id="4f2dd-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="4f2dd-142">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="4f2dd-142">See also</span></span>

- [<span data-ttu-id="4f2dd-143">標準の組み込みオブジェクト</span><span class="sxs-lookup"><span data-stu-id="4f2dd-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="4f2dd-144">Officeスクリプト コード エディター環境</span><span class="sxs-lookup"><span data-stu-id="4f2dd-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
