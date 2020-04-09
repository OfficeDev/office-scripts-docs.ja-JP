---
title: Office スクリプトでの組み込みの JavaScript オブジェクトの使用
description: Web 上の Excel で Office スクリプトから組み込みの JavaScript Api を呼び出す方法について説明します。
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191014"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="20bb1-103">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="20bb1-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="20bb1-104">Javascript には、JavaScript または[TypeScript](../overview/code-editor-environment.md) (javascript のスーパーセット) のどちらでスクリプトを作成するかに関係なく、Office スクリプトで使用できるいくつかの組み込みオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="20bb1-105">この記事では、web 上の Excel の Office スクリプトに組み込まれている JavaScript オブジェクトのいくつかを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="20bb1-106">すべての組み込み JavaScript オブジェクトの完全なリストについては、「Mozilla の[標準の組み込みオブジェクト](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)」記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="20bb1-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="20bb1-107">配列</span><span class="sxs-lookup"><span data-stu-id="20bb1-107">Array</span></span>

<span data-ttu-id="20bb1-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)オブジェクトは、スクリプト内の配列を操作するための標準化された方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="20bb1-109">配列は標準的な JavaScript コンストラクトですが、範囲とコレクションという2つの主な方法で Office スクリプトに関連しています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="20bb1-110">範囲を使用して作業する</span><span class="sxs-lookup"><span data-stu-id="20bb1-110">Working with ranges</span></span>

<span data-ttu-id="20bb1-111">範囲には、その範囲内のセルに直接マップされるいくつかの2次元配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="20bb1-112">これら`values`には、、 `formulas`、などの`numberFormat`プロパティが含まれます。</span><span class="sxs-lookup"><span data-stu-id="20bb1-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="20bb1-113">配列型のプロパティは、他のプロパティと同じように[読み込む](scripting-fundamentals.md#sync-and-load)必要があります。</span><span class="sxs-lookup"><span data-stu-id="20bb1-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="20bb1-114">次のスクリプトは、"$" が含まれている任意の番号書式の**A1: D4**範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="20bb1-115">このスクリプトは、これらのセルの塗りつぶしの色を "黄" に設定します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="20bb1-116">コレクションを処理する</span><span class="sxs-lookup"><span data-stu-id="20bb1-116">Working with collections</span></span>

<span data-ttu-id="20bb1-117">多くの Excel オブジェクトがコレクションに含まれています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="20bb1-118">たとえば、ワークシート内のすべての[図形](/javascript/api/office-scripts/excel/excel.shape)は、 `Worksheet.shapes`プロパティとして、 [offecollection](/javascript/api/office-scripts/excel/excel.shapecollection)に含まれています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="20bb1-119">各`*Collection`オブジェクトには`items` 、プロパティが含まれています。これは、そのコレクション内のオブジェクトを格納する配列です。</span><span class="sxs-lookup"><span data-stu-id="20bb1-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="20bb1-120">これは通常の JavaScript 配列と同様に処理できますが、コレクション内の項目を最初に読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="20bb1-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="20bb1-121">コレクション内のすべてのオブジェクトのプロパティを操作する必要がある場合は、階層 load ステートメント (`items/propertyName`) を使用します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="20bb1-122">次のスクリプトは、現在のワークシート内のすべての図形の種類を記録します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="20bb1-123">`getItem`または`getItemAt`メソッドを使用して、コレクションから個々のオブジェクトを読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="20bb1-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="20bb1-124">`getItem`名前のような一意の識別子を使用してオブジェクトを取得します (そのような名前は、多くの場合、スクリプトで指定されます)。</span><span class="sxs-lookup"><span data-stu-id="20bb1-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="20bb1-125">`getItemAt`コレクション内のインデックスを使用してオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="20bb1-126">オブジェクトを使用するには、 `await context.sync();`その前にコマンドを呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="20bb1-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="20bb1-127">次のスクリプトは、現在のワークシート内の最も古い図形を削除します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="20bb1-128">日付</span><span class="sxs-lookup"><span data-stu-id="20bb1-128">Date</span></span>

<span data-ttu-id="20bb1-129">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)オブジェクトには、スクリプト内の日付を処理するための標準化された方法が用意されています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="20bb1-130">`Date.now()`現在の日付と時刻を使用してオブジェクトを生成します。これは、スクリプトのデータ入力にタイムスタンプを追加するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="20bb1-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="20bb1-131">次のスクリプトは、現在の日付をワークシートに追加します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="20bb1-132">この`toLocaleDateString`メソッドを使用すると、Excel によって値が日付として認識され、セルの数値の書式が自動的に変更されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="20bb1-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="20bb1-133">サンプルの [[日付を使用して作業](../resources/excel-samples.md#work-with-dates)] セクションには、さらに多くの日付関連スクリプトがあります。</span><span class="sxs-lookup"><span data-stu-id="20bb1-133">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="20bb1-134">数学</span><span class="sxs-lookup"><span data-stu-id="20bb1-134">Math</span></span>

<span data-ttu-id="20bb1-135">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)オブジェクトには、一般的な数値演算のためのメソッドと定数が用意されています。</span><span class="sxs-lookup"><span data-stu-id="20bb1-135">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="20bb1-136">これらは、ブックの計算エンジンを使用しなくても、Excel で使用できる多くの関数を提供します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-136">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="20bb1-137">これにより、スクリプトはブックを照会する必要がなくなり、パフォーマンスが向上します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-137">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="20bb1-138">次のスクリプトは`Math.min` 、を使用して、 **A1: D4**範囲の最小数を検索して記録します。</span><span class="sxs-lookup"><span data-stu-id="20bb1-138">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="20bb1-139">この例では、範囲全体に文字列ではなく数値のみが含まれていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="20bb1-139">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="20bb1-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="20bb1-140">See also</span></span>

- [<span data-ttu-id="20bb1-141">標準の組み込みオブジェクト</span><span class="sxs-lookup"><span data-stu-id="20bb1-141">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="20bb1-142">Office スクリプトのコードエディター環境</span><span class="sxs-lookup"><span data-stu-id="20bb1-142">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
