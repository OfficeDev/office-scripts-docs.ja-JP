---
title: Excel on the web での Office スクリプトのスクリプトの基本事項
description: Office スクリプトを作成する前に理解しておくべきオブジェクト モデルの情報と他の基本事項について説明します。
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 9c3c10e283e40f1e719e73106bcdacfcff44dbc9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074509"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="7b4af-103">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="7b4af-103">Scripting fundamentals for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="7b4af-104">この記事では、Office スクリプトの技術的な側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="7b4af-105">Excel オブジェクトどうしが連携する仕組みや、コード エディターがブックと同期する仕組みについて説明します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="7b4af-106">TypeScript: オフィス スクリプトの言語</span><span class="sxs-lookup"><span data-stu-id="7b4af-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="7b4af-107">オフィス スクリプトは [TypeScript](https://www.typescriptlang.org/docs/home.html) で書かれており、[JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) のスーパーセットです。</span><span class="sxs-lookup"><span data-stu-id="7b4af-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="7b4af-108">JavaScript に慣れている場合は、コードの多くが両言語で共通しているため、知識を引き継ぐことができます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="7b4af-109">Office スクリプトのコーディング作業を始める前に、初心者レベルのプログラミング知識を身に付けておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="7b4af-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="7b4af-110">以下のリソースは、Office スクリプトのコーディング面を理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="7b4af-111">`main` 機能: スクリプトの開始点</span><span class="sxs-lookup"><span data-stu-id="7b4af-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="7b4af-112">各スクリプトには、最初のパラメーターとして `ExcelScript.Workbook` 型の `main` 関数を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="7b4af-113">関数を実行すると、Excel アプリケーションはブックを最初のパラメーターとして指定して、この `main` 関数を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="7b4af-114">`ExcelScript.Workbook` は常に最初のパラメータである必要があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="7b4af-115">スクリプトを実行すると、`main` 関数の内部のコードが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="7b4af-116">`main` は、スクリプト内の他の関数を呼び出すことができますが、関数に含まれていないコードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="7b4af-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="7b4af-117">スクリプトは、他の Office スクリプトを呼び出すことはできません。</span><span class="sxs-lookup"><span data-stu-id="7b4af-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="7b4af-118">[Power Automate](https://flow.microsoft.com) では、スクリプトをフローで接続することができます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="7b4af-119">スクリプトとフローの間のデータの受け渡しは、`main` メソッドのパラメーターと戻り値を介して行われます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="7b4af-120">Office スクリプトと Power Automate を統合する方法については、 [「Power Automate で Office スクリプトを実行する」](power-automate-integration.md)で詳しく説明しています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="7b4af-121">オブジェクト モデルの概要</span><span class="sxs-lookup"><span data-stu-id="7b4af-121">Object model overview</span></span>

<span data-ttu-id="7b4af-122">スクリプトを作成するには、Office スクリプト API がどのように連携しているかを理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="7b4af-123">ブックのコンポーネントには、相互に特定の関係があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="7b4af-124">多くの点で、これらの関係は Excel UI の関係と一致しています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="7b4af-125">**ブック** には、1 つ以上の **ワークシート** が含まれます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="7b4af-126">**ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="7b4af-127">**Range** は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="7b4af-128">**Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="7b4af-129">**ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="7b4af-130">**ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="7b4af-131">ブック</span><span class="sxs-lookup"><span data-stu-id="7b4af-131">Workbook</span></span>

<span data-ttu-id="7b4af-132">すべてのスクリプトには、`main` 関数によって `Workbook` 型の `workbook` オブジェクトが提供されています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="7b4af-133">これは、スクリプトが Excel ブックを操作するための最上位レベルのオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="7b4af-134">次のスクリプトは、アクティブなワークシートをブックから取得し、その名前を記録します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="7b4af-135">範囲</span><span class="sxs-lookup"><span data-stu-id="7b4af-135">Ranges</span></span>

<span data-ttu-id="7b4af-136">範囲とは、ブック内の連続したセルのグループのことです。</span><span class="sxs-lookup"><span data-stu-id="7b4af-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="7b4af-137">スクリプトでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。</span><span class="sxs-lookup"><span data-stu-id="7b4af-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="7b4af-138">範囲には、値、数式、書式の 3 つの主要プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="7b4af-139">これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="7b4af-140">`getValues`、`getFormulas`、`getFormat` を介してアクセスします。</span><span class="sxs-lookup"><span data-stu-id="7b4af-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="7b4af-141">値と数式は、`setValues` と `setFormulas` で変更できますが、書式は、個別に設定されている複数の小さなオブジェクトから構成されている `RangeFormat` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="7b4af-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="7b4af-142">範囲は、2 次元配列を使用して情報を管理します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="7b4af-143">Office スクリプト フレームワークでの配列の扱いについては、「[範囲での作業](javascript-objects.md#work-with-ranges)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7b4af-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="7b4af-144">サンプル範囲</span><span class="sxs-lookup"><span data-stu-id="7b4af-144">Range sample</span></span>

<span data-ttu-id="7b4af-145">次のサンプルで、売上記録の作成方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="7b4af-146">このスクリプトは、`Range` オブジェクトを使用して、値、数式、書式の一部を設定しています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.54],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

<span data-ttu-id="7b4af-147">このスクリプトを実行すると、現在のワークシートに次のデータが作成されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="値の行、数式の列、フォーマットされたヘッダーを含む売上記録を含むワークシート。":::

### <a name="the-types-of-range-values"></a><span data-ttu-id="7b4af-149">レンジ値の種類</span><span class="sxs-lookup"><span data-stu-id="7b4af-149">The types of Range values</span></span>

<span data-ttu-id="7b4af-150">各セルには値があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-150">Each cell has value.</span></span> <span data-ttu-id="7b4af-151">この値は、セルに入力された基本的な値であり、Excel で表示されるテキストとは異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-151">This value is the underlying value entered into the cell, which may be different from the text displayed in Excel.</span></span> <span data-ttu-id="7b4af-152">たとえば、セルに日付として "2021 年 5 月 2 日" が表示されていても、実際の値は「44318」であることがあります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-152">For example, you might see "5/2/2021" displayed in the cell as a date, but the actual value is 44318.</span></span> <span data-ttu-id="7b4af-153">この表示は、数値表示形式で変更できますが、セル内の実際の値やタイプは、新しい値が設定されたときにのみ変更されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-153">This display can be changed with the number format, but the actual value and type in the cell only changes when a new value is set.</span></span>

<span data-ttu-id="7b4af-154">セルの値を使用する場合には、セルや範囲からどのような値を得ることを期待しているのかを TypeScript に伝達することが重要です。</span><span class="sxs-lookup"><span data-stu-id="7b4af-154">When you are using the cell value, it's important to tell TypeScript what value you are expecting to get from a cell or range.</span></span> <span data-ttu-id="7b4af-155">セルには、次のいずれかのタブの種類を選択します: `string`、`number`、または `boolean`。</span><span class="sxs-lookup"><span data-stu-id="7b4af-155">A cell contains one of the following types: `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="7b4af-156">スクリプトが返された値をこれらの型の 1 つとして処理するためには、その型を宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-156">In order for your script to treat the returned values as one of those types, you must declare the type.</span></span>

<span data-ttu-id="7b4af-157">次のスクリプトは、前のサンプルのテーブルから平均価格を取得します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-157">The following script gets the average price from the table in the previous sample.</span></span> <span data-ttu-id="7b4af-158">コード `priceRange.getValues() as number[][]` を確認します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-158">Note the code `priceRange.getValues() as number[][]`.</span></span> <span data-ttu-id="7b4af-159">この[アサート](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions)は、範囲の値の型が `number[][]` であることを主張します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-159">This [asserts](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) the type of the range values to be a `number[][]`.</span></span> <span data-ttu-id="7b4af-160">この配列のすべての値は、スクリプトで数字として処理されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-160">All the values in that array can then be treated as numbers in the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="7b4af-161">グラフ、表、およびその他のデータ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7b4af-161">Charts, tables, and other data objects</span></span>

<span data-ttu-id="7b4af-162">スクリプトを使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-162">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="7b4af-163">表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-163">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="7b4af-164">これらはコレクションに格納され、この記事の後半で説明します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-164">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="7b4af-165">テーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="7b4af-165">Create a table</span></span>

<span data-ttu-id="7b4af-p116">データ入力範囲を使ってテーブルを作成します。書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-p116">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="7b4af-168">次のスクリプトでは、前のサンプルの範囲を使用して表を作成します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-168">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="7b4af-169">前のデータを含むワークシート上でこのスクリプトを実行すると、次のテーブルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-169">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="前の売上記録から作成された表を含むワークシート。":::

### <a name="create-a-chart"></a><span data-ttu-id="7b4af-171">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="7b4af-171">Create a chart</span></span>

<span data-ttu-id="7b4af-172">グラフを作成すると、範囲内のデータを視覚化できます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-172">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="7b4af-173">スクリプトでさまざまな種類のグラフを作成できます。いずれのグラフも、必要に応じてカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-173">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="7b4af-174">次のスクリプトで、3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-174">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

<span data-ttu-id="7b4af-175">前の表を含むワークシート上でこのスクリプトを実行すると、次のグラフが作成されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-175">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。":::

## <a name="collections"></a><span data-ttu-id="7b4af-177">コレクション</span><span class="sxs-lookup"><span data-stu-id="7b4af-177">Collections</span></span>

<span data-ttu-id="7b4af-178">Excel オブジェクトは、1 つ以上の同じ種類のオブジェクトのコレクションがある場合、それらを配列に格納します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-178">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="7b4af-179">たとえば、`Workbook` オブジェクトには `Worksheet[]` が含まれます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-179">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="7b4af-180">この配列は `Workbook.getWorksheets()` メソッドでアクセスします。</span><span class="sxs-lookup"><span data-stu-id="7b4af-180">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="7b4af-181">複数の `get` メソッド (`Worksheet.getCharts()` など) は、オブジェクト コレクション全体を配列として返します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-181">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="7b4af-182">このパターンは、Office スクリプトの API 全体で見ることができます。たとえば、`Worksheet` オブジェクトには `getTables()` メソッドがあり、`Table[]` を返し、`Table` オブジェクトには `getColumns()` メソッドがあり、`TableColumn[]` を返すといったことです。</span><span class="sxs-lookup"><span data-stu-id="7b4af-182">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="7b4af-183">返された配列は通常の配列なので、スクリプトでは通常の配列操作がすべて可能です。</span><span class="sxs-lookup"><span data-stu-id="7b4af-183">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="7b4af-184">配列のインデックス値を使用して、コレクション内の個々のオブジェクトにアクセスすることもできます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-184">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="7b4af-185">たとえば、`workbook.getTables()[0]` はコレクション内の最初のテーブルを返します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-185">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="7b4af-186">Office スクリプト フレームワークで組み込みの配列機能を使用する方法については、「[コレクションでの作業](javascript-objects.md#work-with-collections)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7b4af-186">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="7b4af-187">個々のオブジェクトには、`get` メソッドを通してコレクションからアクセスします。</span><span class="sxs-lookup"><span data-stu-id="7b4af-187">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="7b4af-188">単一の `get` メソッド (`Worksheet.getTable(name)` など) は、単一のオブジェクトを返し、特定のオブジェクトの ID または名前を要求します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-188">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="7b4af-189">この ID や名前は通常、スクリプトや Excel の UI で設定します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-189">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="7b4af-p121">次のスクリプトはブック内のすべてのテーブルを取得します。これにより、ヘッダーが表示され、フィルター ボタンが表示され、テーブル スタイルが "TableStyleLight1" に設定されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-p121">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="7b4af-192">スクリプトを使用して Excel オブジェクトを追加する</span><span class="sxs-lookup"><span data-stu-id="7b4af-192">Add Excel objects with a script</span></span>

<span data-ttu-id="7b4af-193">親オブジェクトで使用可能な対応する `add` メソッドを呼び出すことにより、プログラムでテーブルやグラフなどのドキュメント オブジェクトを追加できます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-193">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7b4af-194">コレクション配列にオブジェクトを手動で追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="7b4af-194">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="7b4af-195">親オブジェクトに `add` メソッドを使用します。たとえば、`Worksheet.addTable` メソッドを使用して、`Worksheet` に `Table` を追加します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-195">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="7b4af-196">次のスクリプトは、ブック内の最初のワークシートに Excel のテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-196">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="7b4af-197">作成されたテーブルは、`addTable` メソッドによって返されます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-197">Note that the created table is returned by the `addTable` method.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> <span data-ttu-id="7b4af-198">ほとんどの Excel オブジェクトには `setName` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-198">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="7b4af-199">これにより、スクリプトの後半や、同じワークブックを扱う他のスクリプトで、Excel オブジェクトに簡単にアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-199">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="7b4af-200">コレクションにオブジェクトが存在することを確認する</span><span class="sxs-lookup"><span data-stu-id="7b4af-200">Verify an object exists in the collection</span></span>

<span data-ttu-id="7b4af-201">スクリプトでは、続行する前にテーブルなどのオブジェクトが存在するかどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7b4af-201">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="7b4af-202">スクリプトや Excel の UI で与えられた名前を使って、必要なオブジェクトを特定し、それに応じて行動します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-202">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="7b4af-203">`get` メソッドは、要求されたオブジェクトがコレクションに存在しない場合、`undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-203">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="7b4af-204">次のスクリプトは、"MyTable" という名前のテーブルを要求し、`if...else` ステートメントを使用してテーブルが見つかったかどうか確認します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-204">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

<span data-ttu-id="7b4af-205">Office スクリプトで一般的なパターンは、スクリプトを実行するたびに表やグラフなどのオブジェクトを再作成することです。</span><span class="sxs-lookup"><span data-stu-id="7b4af-205">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="7b4af-206">以前のデータが不要な場合は、新しいオブジェクトを作成する前に以前のオブジェクトを削除するのがよいでしょう。</span><span class="sxs-lookup"><span data-stu-id="7b4af-206">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="7b4af-207">これにより、他のユーザーによってもたらされた名前の競合やその他の相違を避けることができます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-207">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="7b4af-208">次のスクリプトは、"MyTable" という名前のテーブルがあればそれを削除し、同じ名前の新しいテーブルを追加します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-208">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="7b4af-209">スクリプトを使用して Excel オブジェクトを削除する</span><span class="sxs-lookup"><span data-stu-id="7b4af-209">Remove Excel objects with a script</span></span>

<span data-ttu-id="7b4af-210">オブジェクトを削除するには、オブジェクトの `delete` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-210">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="7b4af-211">オブジェクトを追加する場合と同様に、コレクション配列からオブジェクトを手動で削除しないでください。</span><span class="sxs-lookup"><span data-stu-id="7b4af-211">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="7b4af-212">コレクション型のオブジェクトの `delete` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-212">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="7b4af-213">たとえば、`Table.delete` を使用して `Worksheet` から `Table` を削除します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-213">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="7b4af-214">次のスクリプトは、ブック内の最初のワークシートを削除します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-214">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="7b4af-215">オブジェクト モデルに関する参考資料</span><span class="sxs-lookup"><span data-stu-id="7b4af-215">Further reading on the object model</span></span>

<span data-ttu-id="7b4af-216">「[Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)」に、Office スクリプトで使用されるオブジェクトが包括的にまとめられています。</span><span class="sxs-lookup"><span data-stu-id="7b4af-216">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="7b4af-217">目次を使用して、詳細を確認したいクラスに移動できます。</span><span class="sxs-lookup"><span data-stu-id="7b4af-217">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="7b4af-218">よく参照されているページのいくつかを次に示します。</span><span class="sxs-lookup"><span data-stu-id="7b4af-218">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="7b4af-219">グラフ</span><span class="sxs-lookup"><span data-stu-id="7b4af-219">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="7b4af-220">コメント</span><span class="sxs-lookup"><span data-stu-id="7b4af-220">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="7b4af-221">PivotTable</span><span class="sxs-lookup"><span data-stu-id="7b4af-221">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="7b4af-222">Range</span><span class="sxs-lookup"><span data-stu-id="7b4af-222">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="7b4af-223">範囲の形式</span><span class="sxs-lookup"><span data-stu-id="7b4af-223">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="7b4af-224">図形</span><span class="sxs-lookup"><span data-stu-id="7b4af-224">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="7b4af-225">表</span><span class="sxs-lookup"><span data-stu-id="7b4af-225">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="7b4af-226">ブック</span><span class="sxs-lookup"><span data-stu-id="7b4af-226">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="7b4af-227">ワークシート</span><span class="sxs-lookup"><span data-stu-id="7b4af-227">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="7b4af-228">関連項目</span><span class="sxs-lookup"><span data-stu-id="7b4af-228">See also</span></span>

- [<span data-ttu-id="7b4af-229">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="7b4af-229">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="7b4af-230">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="7b4af-230">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="7b4af-231">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7b4af-231">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="7b4af-232">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="7b4af-232">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="7b4af-233">Office スクリプトでのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="7b4af-233">Best practices in Office Scripts</span></span>](best-practices.md)
