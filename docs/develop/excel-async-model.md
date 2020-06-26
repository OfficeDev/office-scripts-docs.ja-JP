---
title: 従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用
description: Office スクリプト非同期 Api の入門と、従来のスクリプトでロード/同期パターンを使用する方法について説明します。
ms.date: 06/22/2020
localization_priority: Normal
ms.openlocfilehash: c7b3c1401ecc2b4d0371590e71f61ae6e9ad8a9d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878854"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a><span data-ttu-id="ef190-103">従来のスクリプトをサポートするための Office スクリプト非同期 Api の使用</span><span class="sxs-lookup"><span data-stu-id="ef190-103">Using the Office Scripts Async APIs to support legacy scripts</span></span>

<span data-ttu-id="ef190-104">この記事では、従来の非同期の Api を使用してスクリプトを記述する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ef190-104">This article will teach you how to write scripts using the legacy, async, APIs.</span></span> <span data-ttu-id="ef190-105">これらの Api は、標準の同期された Office スクリプト Api と同じコア機能を備えていますが、スクリプトとブックとの間のデータ同期を制御する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-105">These APIs have the same core functionality as the standard, synchronous Office Scripts APIs, but they require that your script control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef190-106">Async モデルは、現在の[API モデル](scripting-fundamentals.md?view=office-scripts)を実装する前に作成されたスクリプトでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="ef190-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts).</span></span> <span data-ttu-id="ef190-107">スクリプトは、作成時に作成した API モデルに完全にロックされます。</span><span class="sxs-lookup"><span data-stu-id="ef190-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="ef190-108">これは、レガシスクリプトを新しいモデルに変換する場合は、新しいスクリプトを使用する必要があることも意味します。</span><span class="sxs-lookup"><span data-stu-id="ef190-108">This also means that if you want to convert a legacy script to the new model, you must use a brand new script.</span></span> <span data-ttu-id="ef190-109">現在のモデルは使いやすいため、変更時に古いスクリプトを新しいモデルに更新することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ef190-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="ef190-110">この移行を実行する方法については、「[従来の非同期スクリプトを現在のモデルに変換](#converting-legacy-async-scripts-to-the-current-model)する」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef190-110">The [Converting legacy async scripts to the current model](#converting-legacy-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="ef190-111">`main` 関数</span><span class="sxs-lookup"><span data-stu-id="ef190-111">`main` function</span></span>

<span data-ttu-id="ef190-112">非同期 Api を使用するスクリプトは、別の関数を備えてい `main` ます。</span><span class="sxs-lookup"><span data-stu-id="ef190-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="ef190-113">これは `async` 、を `Excel.RequestContext` 最初のパラメーターとして持つ関数です。</span><span class="sxs-lookup"><span data-stu-id="ef190-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="ef190-114">コンテキスト</span><span class="sxs-lookup"><span data-stu-id="ef190-114">Context</span></span>

<span data-ttu-id="ef190-115">`main` 関数は、`context` という名前の `Excel.RequestContext` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="ef190-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="ef190-116">`context` は、スクリプトとブックの間のブリッジと見なすことができます。</span><span class="sxs-lookup"><span data-stu-id="ef190-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="ef190-117">スクリプトは、`context` オブジェクトを使用してブックにアクセスし、その `context` を使用してデータをやり取りします。</span><span class="sxs-lookup"><span data-stu-id="ef190-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="ef190-118">スクリプトと Excel は異なるプロセスや場所で実行されているため、`context` オブジェクトが必要になります。</span><span class="sxs-lookup"><span data-stu-id="ef190-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="ef190-119">スクリプトで、クラウドのブックに変更を加えたり、そのブックからデータをクエリしたりする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="ef190-120">`context` オブジェクトは、それらのトランザクションを管理します。</span><span class="sxs-lookup"><span data-stu-id="ef190-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="ef190-121">同期と読み込み</span><span class="sxs-lookup"><span data-stu-id="ef190-121">Sync and Load</span></span>

<span data-ttu-id="ef190-122">スクリプトとブックは別の場所で実行されるため、両者の間でデータを転送するには時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="ef190-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="ef190-123">非同期 API では、スクリプトとブックを同期する操作をスクリプトが明示的に呼び出すまで、コマンドがキューに登録され `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="ef190-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="ef190-124">スクリプトは、次のどちらかを実行することが必要になるまで、独立して動作できます。</span><span class="sxs-lookup"><span data-stu-id="ef190-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="ef190-125">ブックからデータを読み取る (`load` 操作または [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult?view=office-scripts-async) を返すメソッドの後)。</span><span class="sxs-lookup"><span data-stu-id="ef190-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult?view=office-scripts-async)).</span></span>
- <span data-ttu-id="ef190-126">ブックにデータを書き込む (通常はスクリプトが完了した結果)。</span><span class="sxs-lookup"><span data-stu-id="ef190-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="ef190-127">次の図に、スクリプトとブックの間の制御フローの例を示します。</span><span class="sxs-lookup"><span data-stu-id="ef190-127">The following image shows an example control flow between the script and workbook:</span></span>

![スクリプトからブックに対して実行される読み取りおよび書き込み操作を示す図。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="ef190-129">同期</span><span class="sxs-lookup"><span data-stu-id="ef190-129">Sync</span></span>

<span data-ttu-id="ef190-130">非同期スクリプトでブックのデータを読み取る必要がある場合、またはブックにデータを書き込む必要がある場合は、次のようにメソッドを呼び出し `RequestContext.sync` ます。</span><span class="sxs-lookup"><span data-stu-id="ef190-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="ef190-131">スクリプトが終了すると、`context.sync()` が暗黙的に呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ef190-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="ef190-132">`sync` 操作が完了すると、ブックが更新され、スクリプトが指定した書き込み操作が反映されます。</span><span class="sxs-lookup"><span data-stu-id="ef190-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="ef190-133">書き込み操作とは、Excel オブジェクトに任意のプロパティを設定すること (`range.format.fill.color = "red"` など)、またはプロパティを変更するメソッドを呼び出すこと (`range.format.autoFitColumns()` など) を意味します。</span><span class="sxs-lookup"><span data-stu-id="ef190-133">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="ef190-134">また、`sync` 操作では、スクリプトが `load` 操作または `ClientResult` を返すメソッドを使用して要求したブックから任意の値が読み取られます (次のセクションを参照)。</span><span class="sxs-lookup"><span data-stu-id="ef190-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="ef190-135">ネットワークによっては、スクリプトとブックを同期するのに時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="ef190-136">`sync`スクリプトの実行速度を速くするために、呼び出しの数を最小限に抑えます。</span><span class="sxs-lookup"><span data-stu-id="ef190-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="ef190-137">それ以外の場合、非同期 Api は標準の同期 Api よりも高速ではありません。</span><span class="sxs-lookup"><span data-stu-id="ef190-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="ef190-138">読み込み</span><span class="sxs-lookup"><span data-stu-id="ef190-138">Load</span></span>

<span data-ttu-id="ef190-139">非同期スクリプトを読み取る前に、ブックからデータを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="ef190-140">ただし、ブック全体からデータを読み込むと、スクリプトの速度が大幅に低下します。</span><span class="sxs-lookup"><span data-stu-id="ef190-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="ef190-141">このメソッドを使用すると、 `load` ブックからどのようなデータを取得するかをスクリプトで明示的に指定できます。</span><span class="sxs-lookup"><span data-stu-id="ef190-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="ef190-142">`load` メソッドは、すべての Excel オブジェクトで使用できます。</span><span class="sxs-lookup"><span data-stu-id="ef190-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="ef190-143">スクリプトでは、オブジェクトのプロパティを読み込んでからでなければ、それらを読み取ることができません。</span><span class="sxs-lookup"><span data-stu-id="ef190-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="ef190-144">そうしないと、エラーになります。</span><span class="sxs-lookup"><span data-stu-id="ef190-144">Not doing so results in an error.</span></span>

<span data-ttu-id="ef190-145">次の例では、`Range` オブジェクトを使用して、`load` メソッドでデータを読み込む方法を示します。</span><span class="sxs-lookup"><span data-stu-id="ef190-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="ef190-146">目的</span><span class="sxs-lookup"><span data-stu-id="ef190-146">Intent</span></span> |<span data-ttu-id="ef190-147">コマンドの例</span><span class="sxs-lookup"><span data-stu-id="ef190-147">Example Command</span></span> | <span data-ttu-id="ef190-148">効果</span><span class="sxs-lookup"><span data-stu-id="ef190-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="ef190-149">1 つのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="ef190-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="ef190-150">単一のプロパティ (この例では、範囲内の値の 2 次元配列) を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ef190-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="ef190-151">複数のプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="ef190-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="ef190-152">コンマで区切られたリストからすべてのプロパティ (この例では、値、行数、列数) を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ef190-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="ef190-153">すべてを読み込む</span><span class="sxs-lookup"><span data-stu-id="ef190-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="ef190-154">範囲のすべてのプロパティを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="ef190-154">Loads all the properties on the range.</span></span> <span data-ttu-id="ef190-155">これは、不要なデータを取得してスクリプトを低速にするために推奨される解決策ではありません。</span><span class="sxs-lookup"><span data-stu-id="ef190-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="ef190-156">これは、スクリプトをテストする場合、またはオブジェクトのすべてのプロパティを必要とする場合にのみ使用してください。</span><span class="sxs-lookup"><span data-stu-id="ef190-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="ef190-157">スクリプトでは、読み込まれた値を読み取る前に、`context.sync()` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

<span data-ttu-id="ef190-158">また、コレクション全体のプロパティを読み込むこともできます。</span><span class="sxs-lookup"><span data-stu-id="ef190-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="ef190-159">Async API のすべてのコレクションオブジェクトには、 `items` そのコレクション内のオブジェクトを含む配列であるプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="ef190-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="ef190-160">`items` を `load` に対する階層呼び出し (`items\myProperty`) の最初に使用すると、それらの項目それぞれの指定されたプロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="ef190-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="ef190-161">次の例では、ワークシートの `CommentCollection` オブジェクトに含まれる各 `Comment` オブジェクトの `resolved` プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="ef190-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a><span data-ttu-id="ef190-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="ef190-162">ClientResult</span></span>

<span data-ttu-id="ef190-163">ブックから情報を返す非同期 API のメソッドには、パラダイムに似たパターンがあり `load` / `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="ef190-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="ef190-164">たとえば、`TableCollection.getCount` はコレクション内のテーブルの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="ef190-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="ef190-165">`getCount` は `ClientResult<number>` を返します。つまり、返される `ClientResult` の `value` プロパティは数値になります。</span><span class="sxs-lookup"><span data-stu-id="ef190-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the return `ClientResult` is a number.</span></span> <span data-ttu-id="ef190-166">`context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="ef190-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="ef190-167">プロパティの読み込みと同様、`value` は、`sync` が呼び出されるまでは、ローカルの "空の" 値です。</span><span class="sxs-lookup"><span data-stu-id="ef190-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="ef190-168">次のスクリプトは、ブック内のテーブルの総数を取得し、その数をコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="ef190-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-legacy-async-scripts-to-the-current-model"></a><span data-ttu-id="ef190-169">従来の非同期スクリプトを現在のモデルに変換する</span><span class="sxs-lookup"><span data-stu-id="ef190-169">Converting legacy async scripts to the current model</span></span>

<span data-ttu-id="ef190-170">現在の API モデルでは、、、またはを使用しません `load` `sync` `RequestContext` 。</span><span class="sxs-lookup"><span data-stu-id="ef190-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="ef190-171">これにより、スクリプトがより簡単に作成および管理できるようになります。</span><span class="sxs-lookup"><span data-stu-id="ef190-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="ef190-172">古いスクリプトを変換するための最善のリソースは、[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)です。</span><span class="sxs-lookup"><span data-stu-id="ef190-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="ef190-173">ここでは、特定のシナリオについてコミュニティにサポートを求めることができます。</span><span class="sxs-lookup"><span data-stu-id="ef190-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="ef190-174">次のガイダンスは、実行する必要のある一般的な手順の概要を示すために役立ちます。</span><span class="sxs-lookup"><span data-stu-id="ef190-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="ef190-175">新しいスクリプトを作成し、それに古い非同期コードをコピーします。</span><span class="sxs-lookup"><span data-stu-id="ef190-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="ef190-176">代わりに、現在の方法を使用して、古いメソッド署名を含めないようにしてください `main` `function main(workbook: ExcelScript.Workbook)` 。</span><span class="sxs-lookup"><span data-stu-id="ef190-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="ef190-177">との呼び出しをすべて削除し `load` `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="ef190-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="ef190-178">これらは不要になりました。</span><span class="sxs-lookup"><span data-stu-id="ef190-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="ef190-179">すべてのプロパティが削除されました。</span><span class="sxs-lookup"><span data-stu-id="ef190-179">All properties have been removed.</span></span> <span data-ttu-id="ef190-180">これらのオブジェクトに and メソッドを使用してアクセスできるようになった `get` `set` ので、これらのプロパティ参照をメソッド呼び出しに切り替える必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="ef190-181">たとえば、次のようなプロパティアクセスを使用してセルの塗りつぶしの色を設定するのではなく、次の `mySheet.getRange("A2:C2").format.fill.color = "blue";` ようなメソッドを使用します。`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="ef190-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="ef190-182">コレクションクラスは、配列に置き換えられました。</span><span class="sxs-lookup"><span data-stu-id="ef190-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="ef190-183">`add` `get` これらのコレクションクラスのメソッドとメソッドは、コレクションを所有していたオブジェクトに移動されたので、それに応じて参照を更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef190-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="ef190-184">たとえば、ブックの最初のワークシートから "MyChart" という名前のグラフを取得するには、次のコードを使用 `workbook.getWorksheets()[0].getChart("MyChart");` します。</span><span class="sxs-lookup"><span data-stu-id="ef190-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="ef190-185">`[0]`で返されるの最初の値にアクセスするには、に注意し `Worksheet[]` て `getWorksheets()` ください。</span><span class="sxs-lookup"><span data-stu-id="ef190-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="ef190-186">わかりやすくするために名前が変更されたメソッドもあります。</span><span class="sxs-lookup"><span data-stu-id="ef190-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="ef190-187">詳細については、「 [Office SCRIPTS API リファレンス](/javascript/api/office-scripts/overview?view=office-scripts)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef190-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="ef190-188">Office スクリプトの非同期 API リファレンスドキュメント</span><span class="sxs-lookup"><span data-stu-id="ef190-188">Office Scripts Async API reference documentation</span></span>

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
