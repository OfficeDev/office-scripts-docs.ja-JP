---
title: Web 上の Excel での Office スクリプトのスクリプトの基礎
description: Office スクリプトを記述する前に知っておくべきオブジェクトモデル情報およびその他の基本事項。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700352"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="b26ab-103">Web 上の Excel での Office スクリプトのスクリプトの基本事項 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b26ab-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="b26ab-104">この記事では、Office スクリプトの技術的な側面について紹介します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="b26ab-105">Excel オブジェクトがどのように連携するか、およびコードエディターがブックとどのように同期されるかについて説明します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="b26ab-106">オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="b26ab-106">Object model</span></span>

<span data-ttu-id="b26ab-107">Excel Api について理解するには、ブックの各コンポーネントが互いにどのように関係しているかを理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="b26ab-108">**ブック**に1つ以上の**ワークシート**が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b26ab-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="b26ab-109">**ワークシート**は、 **Range**オブジェクトを通じてセルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="b26ab-110">**範囲**は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="b26ab-111">**範囲**は、**表**、**グラフ**、**図形**、およびその他のデータビジュアライゼーションまたは組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="b26ab-112">**ワークシート**には、個々のシートに存在するデータオブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="b26ab-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="b26ab-113">**ブック全体**のデータオブジェクト (**テーブル**など) のコレクションが含まれ**ています**。</span><span class="sxs-lookup"><span data-stu-id="b26ab-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="b26ab-114">Ranges</span><span class="sxs-lookup"><span data-stu-id="b26ab-114">Ranges</span></span>

<span data-ttu-id="b26ab-115">範囲は、ブック内の連続したセルのグループです。</span><span class="sxs-lookup"><span data-stu-id="b26ab-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="b26ab-116">通常、スクリプトでは、A1 形式の表記を使用します (たとえば **、行** **B**の1つのセルの場合は**B3** **、列**2 から**F**は**2** ~ **4**の範囲を定義**するため)** 。</span><span class="sxs-lookup"><span data-stu-id="b26ab-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="b26ab-117">範囲には、、、 `values`および`formulas` `format`の3つの主要なプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="b26ab-118">これらのプロパティは、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="b26ab-119">範囲のサンプル</span><span class="sxs-lookup"><span data-stu-id="b26ab-119">Range sample</span></span>

<span data-ttu-id="b26ab-120">次の例は、sales レコードを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b26ab-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="b26ab-121">このスクリプトは`Range` 、オブジェクトを使用して、値、式、および書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

<span data-ttu-id="b26ab-122">このスクリプトを実行すると、現在のワークシートに次のデータが作成されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-122">Running this script creates the following data in the current worksheet:</span></span>

![値の行、数式の列、および書式設定されたヘッダーを示す sales レコード。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="b26ab-124">グラフ、表、およびその他のデータオブジェクト</span><span class="sxs-lookup"><span data-stu-id="b26ab-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="b26ab-125">スクリプトは、Excel 内でデータ構造と視覚エフェクトを作成して操作できます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="b26ab-126">表とグラフは、よく使用されるオブジェクトの2つですが、Api はピボットテーブル、図形、画像などをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="b26ab-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="b26ab-127">テーブルを作成する</span><span class="sxs-lookup"><span data-stu-id="b26ab-127">Creating a table</span></span>

<span data-ttu-id="b26ab-128">データの埋め込まれた範囲を使用してテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="b26ab-129">書式設定とテーブルコントロール (フィルターなど) が範囲に自動的に適用されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="b26ab-130">次のスクリプトは、前の例の範囲を使用してテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="b26ab-131">以前のデータを含むワークシートでこのスクリプトを実行すると、次のテーブルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![前の販売レコードから作成されたテーブル。](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="b26ab-133">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="b26ab-133">Creating a chart</span></span>

<span data-ttu-id="b26ab-134">範囲内のデータを視覚化するためのグラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="b26ab-135">スクリプトを使用すると、さまざまなグラフをさまざまな方法で使用できます。これらはそれぞれのニーズに合わせてカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="b26ab-136">次のスクリプトは、3つのアイテムの簡単な縦棒グラフを作成し、それをワークシートの一番上の100ピクセルの下に配置します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="b26ab-137">前の表を使用して、このスクリプトをワークシートで実行すると、次のグラフが作成されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![前の売上レコードからの3つのアイテムの数量を示す縦棒グラフ。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="b26ab-139">オブジェクトモデルについてのさらなる閲覧</span><span class="sxs-lookup"><span data-stu-id="b26ab-139">Further reading on the object model</span></span>

<span data-ttu-id="b26ab-140">[Office SCRIPTS API リファレンスドキュメント](/javascript/api/office-scripts/overview)は、office スクリプトで使用されるオブジェクトの包括的な一覧です。</span><span class="sxs-lookup"><span data-stu-id="b26ab-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="b26ab-141">そこで、目次を使用して、詳細について知りたいクラスに移動できます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="b26ab-142">一般的に表示されるページをいくつか次に示します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="b26ab-143">Chart</span><span class="sxs-lookup"><span data-stu-id="b26ab-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="b26ab-144">Comment</span><span class="sxs-lookup"><span data-stu-id="b26ab-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="b26ab-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="b26ab-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="b26ab-146">Range</span><span class="sxs-lookup"><span data-stu-id="b26ab-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="b26ab-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="b26ab-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="b26ab-148">Shape</span><span class="sxs-lookup"><span data-stu-id="b26ab-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="b26ab-149">Table</span><span class="sxs-lookup"><span data-stu-id="b26ab-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="b26ab-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="b26ab-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="b26ab-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="b26ab-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="b26ab-152">`main`関数</span><span class="sxs-lookup"><span data-stu-id="b26ab-152">`main` function</span></span>

<span data-ttu-id="b26ab-153">すべての Office スクリプトには`main` 、 `Excel.RequestContext`型定義を含む次のシグネチャを持つ関数が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="b26ab-154">スクリプトの実行時`main`に関数内のコードが実行されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="b26ab-155">`main`スクリプトで他の関数を呼び出すことはできますが、関数に含まれていないコードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="b26ab-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="b26ab-156">[Context/文脈に従う]</span><span class="sxs-lookup"><span data-stu-id="b26ab-156">Context</span></span>

<span data-ttu-id="b26ab-157">関数`main`は、と`Excel.RequestContext`いう名前`context`のパラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="b26ab-158">スクリプトと`context`ブックの間のブリッジと考えることができます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="b26ab-159">スクリプトは、 `context`オブジェクトを使用してブックにアクセス`context`し、それを使用してデータをやり取りします。</span><span class="sxs-lookup"><span data-stu-id="b26ab-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="b26ab-160">スクリプト`context`と Excel が異なるプロセスと場所で実行されているため、オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="b26ab-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="b26ab-161">このスクリプトでは、クラウド内のブックに対して変更を加えるか、データを照会する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="b26ab-162">オブジェクト`context`はこれらのトランザクションを管理します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="b26ab-163">同期と読み込み</span><span class="sxs-lookup"><span data-stu-id="b26ab-163">Sync and Load</span></span>

<span data-ttu-id="b26ab-164">スクリプトとブックは異なる場所で実行されるので、2つの間のデータ転送には時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="b26ab-165">スクリプトのパフォーマンスを向上させるために、スクリプトとブックを同期`sync`する操作をスクリプトが明示的に呼び出すまで、コマンドはキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="b26ab-166">スクリプトは、次のいずれかが必要になるまで、独立して動作することができます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="b26ab-167">ブックからデータを読み取ります (操作`load`の後)。</span><span class="sxs-lookup"><span data-stu-id="b26ab-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="b26ab-168">ブックにデータを書き込みます (通常はスクリプトが完了したため)。</span><span class="sxs-lookup"><span data-stu-id="b26ab-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="b26ab-169">次の図は、スクリプトとブックの間の制御フローの例を示しています。</span><span class="sxs-lookup"><span data-stu-id="b26ab-169">The following image shows an example control flow between the script and workbook:</span></span>

![スクリプトからブックに対する読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="b26ab-171">同期</span><span class="sxs-lookup"><span data-stu-id="b26ab-171">Sync</span></span>

<span data-ttu-id="b26ab-172">スクリプトでブックのデータの読み取りまたは書き込みを行う必要がある場合は`RequestContext.sync` 、次のようにメソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="b26ab-173">`context.sync()`は、スクリプトの終了時に暗黙的に呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="b26ab-174">操作が`sync`完了すると、ブックは、スクリプトが指定した書き込み操作を反映するように更新されます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="b26ab-175">書き込み操作は、Excel オブジェクト (たとえば`range.format.fill.color = "red"`、 `range.format.autoFitColumns()`) にプロパティを設定するか、またはプロパティを変更するメソッドを呼び出しています (例:)。</span><span class="sxs-lookup"><span data-stu-id="b26ab-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="b26ab-176">また`sync` 、操作を使用`load`して、スクリプトが要求したブックの値を読み取ります (次のセクションで説明します)。</span><span class="sxs-lookup"><span data-stu-id="b26ab-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="b26ab-177">ネットワークによっては、スクリプトとブックを同期するときに時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="b26ab-178">スクリプトの実行速度を速く`sync`するには、呼び出しの数を最小限に抑える必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="b26ab-179">読み込め</span><span class="sxs-lookup"><span data-stu-id="b26ab-179">Load</span></span>

<span data-ttu-id="b26ab-180">スクリプトを読み取る前に、ブックからデータを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="b26ab-181">ただし、多くの場合、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="b26ab-182">代わりに、この`load`メソッドを使用すると、スクリプトの状態を特定し、ブックから取得するデータを指定できます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="b26ab-183">この`load`メソッドは、すべての Excel オブジェクトで使用できます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="b26ab-184">スクリプトは、オブジェクトのプロパティを読み取れるように読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="b26ab-185">そうしないと、エラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="b26ab-186">次の例では`Range` 、オブジェクトを使用して、 `load`メソッドがデータを読み込むために使用できる3つの方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="b26ab-187">Intent</span><span class="sxs-lookup"><span data-stu-id="b26ab-187">Intent</span></span> |<span data-ttu-id="b26ab-188">コマンドの例</span><span class="sxs-lookup"><span data-stu-id="b26ab-188">Example Command</span></span> | <span data-ttu-id="b26ab-189">効果</span><span class="sxs-lookup"><span data-stu-id="b26ab-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="b26ab-190">1つのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="b26ab-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="b26ab-191">この例では、1つのプロパティ (この場合は、この範囲内の値の2次元配列) を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="b26ab-192">複数のプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="b26ab-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="b26ab-193">コンマで区切られたリストからすべてのプロパティを読み込みます。この例では、値、行の数、および列の数を指定します。</span><span class="sxs-lookup"><span data-stu-id="b26ab-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="b26ab-194">すべてを読み込む</span><span class="sxs-lookup"><span data-stu-id="b26ab-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="b26ab-195">範囲のすべてのプロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-195">Loads all the properties on the range.</span></span> <span data-ttu-id="b26ab-196">これは、不要なデータを取得してスクリプトを低速にするために推奨されるソリューションではありません。</span><span class="sxs-lookup"><span data-stu-id="b26ab-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="b26ab-197">この値は、スクリプトのテスト時、またはオブジェクトのすべてのプロパティが必要な場合にのみ使用してください。</span><span class="sxs-lookup"><span data-stu-id="b26ab-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="b26ab-198">スクリプトは、読み込ま`context.sync()`れた値を読み取る前に呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="b26ab-199">また、コレクション全体に対してプロパティを読み込むこともできます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="b26ab-200">すべての collection オブジェクトに`items`は、そのコレクション内のオブジェクトを含む配列であるプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b26ab-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="b26ab-201">を`items`階層呼び出し`items\myProperty`の開始として使用し、 `load`各アイテムの指定されたプロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="b26ab-202">次の使用例は`resolved` 、ワークシート`Comment`の`CommentCollection`オブジェクト内のすべてのオブジェクトのプロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b26ab-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="b26ab-203">Office スクリプトでのコレクションの使用の詳細については、「 [Office スクリプトの組み込み JavaScript オブジェクトの使用」の記事の「Array」セクション](javascript-objects.md#array)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b26ab-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="b26ab-204">関連項目</span><span class="sxs-lookup"><span data-stu-id="b26ab-204">See also</span></span>

- [<span data-ttu-id="b26ab-205">Web 上の Excel で Office スクリプトを記録、編集、および作成する</span><span class="sxs-lookup"><span data-stu-id="b26ab-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="b26ab-206">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="b26ab-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="b26ab-207">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b26ab-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="b26ab-208">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="b26ab-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
