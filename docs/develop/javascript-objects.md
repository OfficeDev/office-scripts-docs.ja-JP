---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Web 上の Excel で Office スクリプトから組み込みの JavaScript Api を呼び出す方法について説明します。
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: 54cadb6e9ce60e631488bbe7de00c29a6db35eb7
ms.sourcegitcommit: b13dedb5ee2048f0a244aa2294bf2c38697cb62c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215260"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="47e17-103">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="47e17-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="47e17-104">Javascript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (javascript のスーパーセット) のどちらでスクリプトを作成するかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="47e17-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="47e17-105">この記事では、web 上の Excel の Office スクリプトに組み込まれている JavaScript オブジェクトのいくつかを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="47e17-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="47e17-106">すべての組み込み JavaScript オブジェクトの完全なリストについては、「Mozilla の[標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)」記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="47e17-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="47e17-107">配列</span><span class="sxs-lookup"><span data-stu-id="47e17-107">Array</span></span>

<span data-ttu-id="47e17-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を操作するための標準化された方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="47e17-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="47e17-109">配列は標準的な JavaScript コンストラクトですが、範囲とコレクションという2つの主な方法で Office スクリプトに関連しています。</span><span class="sxs-lookup"><span data-stu-id="47e17-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="47e17-110">範囲を使用して作業する</span><span class="sxs-lookup"><span data-stu-id="47e17-110">Working with ranges</span></span>

<span data-ttu-id="47e17-111">範囲には、その範囲内のセルに直接マップされるいくつかの2次元配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="47e17-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="47e17-112">これら`values`には、、 `formulas`、などの`numberFormat`プロパティが含まれます。</span><span class="sxs-lookup"><span data-stu-id="47e17-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="47e17-113">配列型のプロパティは、他のプロパティと同じように[読み込む](scripting-fundamentals.md#sync-and-load)必要があります。</span><span class="sxs-lookup"><span data-stu-id="47e17-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="47e17-114">次のスクリプトは、"$" が含まれている任意の番号書式の**A1: D4**範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="47e17-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="47e17-115">このスクリプトは、これらのセルの塗りつぶしの色を "黄" に設定します。</span><span class="sxs-lookup"><span data-stu-id="47e17-115">The script sets the fill color in those cells to "yellow".</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a><span data-ttu-id="47e17-116">コレクションを処理する</span><span class="sxs-lookup"><span data-stu-id="47e17-116">Working with collections</span></span>

<span data-ttu-id="47e17-117">多くの Excel オブジェクトがコレクションに含まれています。</span><span class="sxs-lookup"><span data-stu-id="47e17-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="47e17-118">たとえば、ワークシート内のすべての[図形](/javascript/api/office-scripts/excel/excel.shape)は、 `Worksheet.shapes`プロパティとして、 [offecollection](/javascript/api/office-scripts/excel/excel.shapecollection)に含まれています。</span><span class="sxs-lookup"><span data-stu-id="47e17-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="47e17-119">各`*Collection`オブジェクトには`items` 、プロパティが含まれています。これは、そのコレクション内のオブジェクトを格納する配列です。</span><span class="sxs-lookup"><span data-stu-id="47e17-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="47e17-120">これは通常の JavaScript 配列と同様に処理できますが、コレクション内の項目を最初に読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="47e17-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="47e17-121">コレクション内のすべてのオブジェクトのプロパティを操作する必要がある場合は、階層 load ステートメント (`items/propertyName`) を使用します。</span><span class="sxs-lookup"><span data-stu-id="47e17-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="47e17-122">次のスクリプトは、現在のワークシート内のすべての図形の種類を記録します。</span><span class="sxs-lookup"><span data-stu-id="47e17-122">The following script logs the type of every shape in the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

<span data-ttu-id="47e17-123">`getItem`または`getItemAt`メソッドを使用して、コレクションから個々のオブジェクトを読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="47e17-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="47e17-124">`getItem`名前のような一意の識別子を使用してオブジェクトを取得します (そのような名前は、多くの場合、スクリプトで指定されます)。</span><span class="sxs-lookup"><span data-stu-id="47e17-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="47e17-125">`getItemAt`コレクション内のインデックスを使用してオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="47e17-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="47e17-126">オブジェクトを使用するには、 `await context.sync();`その前にコマンドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="47e17-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="47e17-127">次のスクリプトは、現在のワークシート内の最も古い図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="47e17-127">The following script deletes the oldest shape in the current worksheet.</span></span>

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a><span data-ttu-id="47e17-128">日付</span><span class="sxs-lookup"><span data-stu-id="47e17-128">Date</span></span>

<span data-ttu-id="47e17-129">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトには、スクリプト内の日付を処理するための標準化された方法が用意されています。</span><span class="sxs-lookup"><span data-stu-id="47e17-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="47e17-130">`Date.now()`現在の日付と時刻を使用してオブジェクトを生成します。これは、スクリプトのデータ入力にタイムスタンプを追加するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="47e17-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="47e17-131">次のスクリプトは、現在の日付をワークシートに追加します。</span><span class="sxs-lookup"><span data-stu-id="47e17-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="47e17-132">この`toLocaleDateString`メソッドを使用すると、Excel によって値が日付として認識され、セルの数値の書式が自動的に変更されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="47e17-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

<span data-ttu-id="47e17-133">サンプルの [[日付を使用して作業](../resources/excel-samples.md#work-with-dates)] セクションには、さらに多くの日付関連スクリプトがあります。</span><span class="sxs-lookup"><span data-stu-id="47e17-133">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="47e17-134">数学</span><span class="sxs-lookup"><span data-stu-id="47e17-134">Math</span></span>

<span data-ttu-id="47e17-135">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトには、一般的な数値演算のためのメソッドと定数が用意されています。</span><span class="sxs-lookup"><span data-stu-id="47e17-135">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="47e17-136">これらは、ブックの計算エンジンを使用しなくても、Excel で使用できる多くの関数を提供します。</span><span class="sxs-lookup"><span data-stu-id="47e17-136">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="47e17-137">これにより、スクリプトはブックを照会する必要がなくなり、パフォーマンスが向上します。</span><span class="sxs-lookup"><span data-stu-id="47e17-137">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="47e17-138">次のスクリプトは`Math.min` 、を使用して、 **A1: D4**範囲の最小数を検索して記録します。</span><span class="sxs-lookup"><span data-stu-id="47e17-138">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="47e17-139">この例では、範囲全体に文字列ではなく数値のみが含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="47e17-139">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="47e17-140">外部 JavaScript ライブラリの使用はサポートされていません</span><span class="sxs-lookup"><span data-stu-id="47e17-140">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="47e17-141">Office スクリプトは、外部のサードパーティ製ライブラリの使用をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="47e17-141">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="47e17-142">スクリプトでは、組み込みの JavaScript オブジェクトと Office スクリプト Api のみを使用できます。</span><span class="sxs-lookup"><span data-stu-id="47e17-142">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="47e17-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="47e17-143">See also</span></span>

- [<span data-ttu-id="47e17-144">標準の組み込みオブジェクト</span><span class="sxs-lookup"><span data-stu-id="47e17-144">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="47e17-145">Office スクリプトのコードエディター環境</span><span class="sxs-lookup"><span data-stu-id="47e17-145">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
