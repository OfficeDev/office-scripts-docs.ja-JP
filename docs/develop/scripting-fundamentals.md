---
title: Excel on the web での Office スクリプトのスクリプトの基本事項
description: Office スクリプトを作成する前に理解しておくべきオブジェクト モデルの情報と他の基本事項について説明します。
ms.date: 04/24/2020
localization_priority: Priority
ms.openlocfilehash: 8449654e359f665677f3d416a8e28fa4d6930f26
ms.sourcegitcommit: 350bd2447f616fa87bb23ac826c7731fb813986b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/28/2020
ms.locfileid: "43919799"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9e5df-103">Excel on the web での Office スクリプトのスクリプトの基本事項 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="9e5df-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9e5df-104">この記事では、Office スクリプトの技術的な側面について説明します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="9e5df-105">Excel オブジェクトどうしが連携する仕組みや、コード エディターがブックと同期する仕組みについて説明します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="9e5df-106">オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="9e5df-106">Object model</span></span>

<span data-ttu-id="9e5df-107">Excel API について理解するには、ブックの構成要素が互いにどのように関連しているかを理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="9e5df-108">**ブック** には、1 つ以上の **ワークシート** が含まれます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="9e5df-109">**ワークシート** では、**Range** オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="9e5df-110">**Range** は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="9e5df-111">**Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="9e5df-112">**ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="9e5df-113">**ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="9e5df-114">範囲</span><span class="sxs-lookup"><span data-stu-id="9e5df-114">Ranges</span></span>

<span data-ttu-id="9e5df-115">範囲とは、ブック内の連続したセルのグループのことです。</span><span class="sxs-lookup"><span data-stu-id="9e5df-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="9e5df-116">スクリプトでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。</span><span class="sxs-lookup"><span data-stu-id="9e5df-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="9e5df-117">範囲には `values`、`formulas`、`format` の 3 つの主要なプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="9e5df-118">これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="9e5df-119">サンプル範囲</span><span class="sxs-lookup"><span data-stu-id="9e5df-119">Range sample</span></span>

<span data-ttu-id="9e5df-120">次のサンプルで、売上記録の作成方法を示します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="9e5df-121">このスクリプトは、`Range` オブジェクトを使用して、値、数式、書式を設定しています。</span><span class="sxs-lookup"><span data-stu-id="9e5df-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="9e5df-122">このスクリプトを実行すると、現在のワークシートに次のデータが作成されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-122">Running this script creates the following data in the current worksheet:</span></span>

![値の行、数式の列、書式設定されたヘッダーを示す売上記録。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="9e5df-124">グラフ、表、およびその他のデータ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e5df-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="9e5df-125">スクリプトを使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="9e5df-126">表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="9e5df-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="9e5df-127">表の作成</span><span class="sxs-lookup"><span data-stu-id="9e5df-127">Creating a table</span></span>

<span data-ttu-id="9e5df-128">データが入力された範囲を使用することにより、表を作成します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="9e5df-129">書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="9e5df-130">次のスクリプトでは、前のサンプルの範囲を使用して表を作成します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="9e5df-131">前のデータを含むワークシート上でこのスクリプトを実行すると、次のテーブルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![前の売上記録から作成された表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="9e5df-133">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="9e5df-133">Creating a chart</span></span>

<span data-ttu-id="9e5df-134">グラフを作成すると、範囲内のデータを視覚化できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="9e5df-135">スクリプトでさまざまな種類のグラフを作成できます。いずれのグラフも、必要に応じてカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="9e5df-136">次のスクリプトで、3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="9e5df-137">前の表を含むワークシート上でこのスクリプトを実行すると、次のグラフが作成されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="9e5df-139">オブジェクト モデルに関する参考資料</span><span class="sxs-lookup"><span data-stu-id="9e5df-139">Further reading on the object model</span></span>

<span data-ttu-id="9e5df-140">「[Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)」に、Office スクリプトで使用されるオブジェクトが包括的にまとめられています。</span><span class="sxs-lookup"><span data-stu-id="9e5df-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="9e5df-141">目次を使用して、詳細を確認したいクラスに移動できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="9e5df-142">よく参照されているページのいくつかを次に示します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="9e5df-143">グラフ</span><span class="sxs-lookup"><span data-stu-id="9e5df-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="9e5df-144">コメント</span><span class="sxs-lookup"><span data-stu-id="9e5df-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="9e5df-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="9e5df-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="9e5df-146">Range</span><span class="sxs-lookup"><span data-stu-id="9e5df-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="9e5df-147">範囲の形式</span><span class="sxs-lookup"><span data-stu-id="9e5df-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="9e5df-148">図形</span><span class="sxs-lookup"><span data-stu-id="9e5df-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="9e5df-149">表</span><span class="sxs-lookup"><span data-stu-id="9e5df-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="9e5df-150">ブック</span><span class="sxs-lookup"><span data-stu-id="9e5df-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="9e5df-151">ワークシート</span><span class="sxs-lookup"><span data-stu-id="9e5df-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="9e5df-152">`main` 関数</span><span class="sxs-lookup"><span data-stu-id="9e5df-152">`main` function</span></span>

<span data-ttu-id="9e5df-153">どの Office スクリプトにも、次のシグネチャで、`Excel.RequestContext` 型の定義を含む `main` 関数を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="9e5df-154">スクリプトを実行すると、`main` 関数の内部のコードが実行されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="9e5df-155">`main` は、スクリプト内の他の関数を呼び出すことができますが、関数に含まれていないコードは実行されません。</span><span class="sxs-lookup"><span data-stu-id="9e5df-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="9e5df-156">コンテキスト</span><span class="sxs-lookup"><span data-stu-id="9e5df-156">Context</span></span>

<span data-ttu-id="9e5df-157">`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="9e5df-158">`context` は、スクリプトとブックの間のブリッジと見なすことができます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="9e5df-159">スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。</span><span class="sxs-lookup"><span data-stu-id="9e5df-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="9e5df-160">スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="9e5df-161">スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="9e5df-162">`context` オブジェクトは、それらのトランザクションを管理します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="9e5df-163">同期と読み込み</span><span class="sxs-lookup"><span data-stu-id="9e5df-163">Sync and Load</span></span>

<span data-ttu-id="9e5df-164">スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="9e5df-165">スクリプトのパフォーマンスを向上させるため、スクリプトが明示的に `sync` 操作を呼び出してスクリプトとブックを同期するまで、コマンドはキューに登録されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="9e5df-166">スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="9e5df-167">ブックからデータを読み取る (`load` 操作または [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult) を返すメソッドの後)。</span><span class="sxs-lookup"><span data-stu-id="9e5df-167">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).</span></span>
- <span data-ttu-id="9e5df-168">ブックにデータを書き込む (通常はスクリプトが完了した結果)。</span><span class="sxs-lookup"><span data-stu-id="9e5df-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="9e5df-169">次の図に、スクリプトとブックの間の制御フローの例を示します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-169">The following image shows an example control flow between the script and workbook:</span></span>

![スクリプトからブックに対して実行される読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="9e5df-171">同期</span><span class="sxs-lookup"><span data-stu-id="9e5df-171">Sync</span></span>

<span data-ttu-id="9e5df-172">スクリプトでブックに対するデータの読み取りや書き込みが必要になる場合、次のように `RequestContext.sync` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="9e5df-173">スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="9e5df-174">`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="9e5df-175">書き込み操作とは、Excel オブジェクトに任意のプロパティを設定すること (`range.format.fill.color = "red"` など)、またはプロパティを変更するメソッドを呼び出すこと (`range.format.autoFitColumns()` など) を意味します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="9e5df-176">また、`sync` 操作では、スクリプトが `load` 操作または `ClientResult` を返すメソッドを使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。</span><span class="sxs-lookup"><span data-stu-id="9e5df-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="9e5df-177">ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="9e5df-178">スクリプトの実行速度を高めるため、`sync` 呼び出しは最小限に抑えることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9e5df-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="9e5df-179">読み込み</span><span class="sxs-lookup"><span data-stu-id="9e5df-179">Load</span></span>

<span data-ttu-id="9e5df-180">スクリプトでは、ブックからデータを読み込んでから、そのデータを読み取る必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="9e5df-181">しかし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="9e5df-182">代わりに、`load` メソッドを使用すると、どのデータをブックから取得する必要があるかをスクリプトで具体的に指定できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="9e5df-183">`load` メソッドは、すべての Excel オブジェクトで使用できます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="9e5df-184">スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。</span><span class="sxs-lookup"><span data-stu-id="9e5df-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="9e5df-185">これに従わないと、エラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="9e5df-186">次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="9e5df-187">目的</span><span class="sxs-lookup"><span data-stu-id="9e5df-187">Intent</span></span> |<span data-ttu-id="9e5df-188">コマンドの例</span><span class="sxs-lookup"><span data-stu-id="9e5df-188">Example Command</span></span> | <span data-ttu-id="9e5df-189">効果</span><span class="sxs-lookup"><span data-stu-id="9e5df-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="9e5df-190">1 つのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="9e5df-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="9e5df-191">単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="9e5df-192">複数のプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="9e5df-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="9e5df-193">コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="9e5df-194">すべてを読み込む</span><span class="sxs-lookup"><span data-stu-id="9e5df-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="9e5df-195">範囲のすべてのプロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-195">Loads all the properties on the range.</span></span> <span data-ttu-id="9e5df-196">このソリューションは、不要なデータを取得することによりスクリプトの速度が低下するため、推奨されません。</span><span class="sxs-lookup"><span data-stu-id="9e5df-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="9e5df-197">スクリプトをテストする場合、またはオブジェクトのすべてのプロパティが必要な場合にのみ使用してください。</span><span class="sxs-lookup"><span data-stu-id="9e5df-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="9e5df-198">スクリプトでは、読み込まれた値を読み取る前に、`context.sync()` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="9e5df-199">また、コレクション全体のプロパティを読み込むこともできます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="9e5df-200">どのコレクション オブジェクトにも、`items` プロパティがあります。これは、そのコレクションのオブジェクトを格納する配列です。</span><span class="sxs-lookup"><span data-stu-id="9e5df-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="9e5df-201">`items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="9e5df-202">次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="9e5df-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="9e5df-203">Office スクリプトでのコレクションの使用方法の詳細については、[「Office スクリプトでの組み込みの JavaScript オブジェクトの使用」の「配列」セクション](javascript-objects.md#array)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9e5df-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

### <a name="clientresult"></a><span data-ttu-id="9e5df-204">ClientResult</span><span class="sxs-lookup"><span data-stu-id="9e5df-204">ClientResult</span></span>

<span data-ttu-id="9e5df-205">ブックから情報を返すメソッドには、`load`/`sync` パラダイムと似たパターンがあります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-205">Methods that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="9e5df-206">たとえば、`TableCollection.getCount` はコレクション内のテーブルの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-206">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="9e5df-207">`getCount` は `ClientResult<number>` を返します。つまり、返される `ClientResult` の `value` プロパティは数値になります。</span><span class="sxs-lookup"><span data-stu-id="9e5df-207">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the return `ClientResult` is a number.</span></span> <span data-ttu-id="9e5df-208">`context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="9e5df-208">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="9e5df-209">プロパティの読み込みと同様、`value` は、`sync` が呼び出されるまでは、ローカルの "空の" 値です。</span><span class="sxs-lookup"><span data-stu-id="9e5df-209">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="9e5df-210">次のスクリプトは、ブック内のテーブルの総数を取得し、その数をコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="9e5df-210">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let tableCount = context.workbook.tables.getCount();

  // This sync call implicitly loads tableCount.value.
  // Any other ClientResult values are loaded too.
  await context.sync();

  // Trying to log the value before calling sync would throw an error.
  console.log(tableCount.value);
}
```

## <a name="see-also"></a><span data-ttu-id="9e5df-211">関連項目</span><span class="sxs-lookup"><span data-stu-id="9e5df-211">See also</span></span>

- [<span data-ttu-id="9e5df-212">Excel on the web で Office スクリプトを記録、編集、作成する</span><span class="sxs-lookup"><span data-stu-id="9e5df-212">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="9e5df-213">Excel on the web で Office スクリプトを使用してブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="9e5df-213">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="9e5df-214">Office スクリプト API リファレンス</span><span class="sxs-lookup"><span data-stu-id="9e5df-214">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="9e5df-215">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="9e5df-215">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
