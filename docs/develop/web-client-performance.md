---
title: Office スクリプトのパフォーマンスを向上させる
description: Excel ブックとスクリプトの間の通信を理解することで、より高速なスクリプトを作成できます。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878899"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="3b92f-103">Office スクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="3b92f-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="3b92f-104">Office スクリプトの目的は、頻繁に実行される一連のタスクを自動化して時間を節約することです。</span><span class="sxs-lookup"><span data-stu-id="3b92f-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="3b92f-105">低速なスクリプトは、ワークフローを高速化しないように感じられます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="3b92f-106">ほとんどの場合、スクリプトは完全に機能し、期待どおりに実行されます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="3b92f-107">ただし、パフォーマンスに影響する可能性のあるいくつかの avoidable のシナリオがあります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="3b92f-108">時間のかかるスクリプトの最も一般的な原因は、ブックとの通信が多すぎることです。</span><span class="sxs-lookup"><span data-stu-id="3b92f-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="3b92f-109">スクリプトは、ローカルコンピューター上で実行されます。ブックはクラウド内に存在します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="3b92f-110">場合によっては、スクリプトによってローカルデータがブックの内容と同期されます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="3b92f-111">これは、このようなバックグラウンドでの同期が発生したときに、(などの) 書き込み操作 `workbook.addWorksheet()` がブックにのみ適用されることを意味します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="3b92f-112">同様に、どのような読み取り操作 (など) でも、その `myRange.getValues()` 時点でスクリプトのブックからデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="3b92f-113">どちらの場合も、スクリプトはデータを処理する前に情報をフェッチします。</span><span class="sxs-lookup"><span data-stu-id="3b92f-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="3b92f-114">たとえば、次のコードでは、使用されている範囲内の行数を正確に記録します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="3b92f-115">Office スクリプト Api は、ブックまたはスクリプト内のすべてのデータが正確で、必要に応じて最新の状態になっていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="3b92f-116">スクリプトを正しく実行するために、これらの同期について心配する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="3b92f-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="3b92f-117">ただし、このスクリプトからクラウドへの通信を認識することは、不要なネットワーク呼び出しを回避するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="3b92f-118">パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="3b92f-118">Performance optimizations</span></span>

<span data-ttu-id="3b92f-119">クラウドへの通信を減らすための簡単な手法を適用することができます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="3b92f-120">次のパターンは、スクリプトを高速化するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="3b92f-121">繰り返しループ内ではなく、ブックのデータを1回読み取ります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="3b92f-122">不要な `console.log` ステートメントを削除します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="3b92f-123">Try/catch ブロックを使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="3b92f-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="3b92f-124">ループの外部でブックのデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="3b92f-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="3b92f-125">ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーすることができます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="3b92f-126">同じ通話を繰り返し作成するのではなく、可能な限りデータをローカルに保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="3b92f-127">これは、ループを処理するときに特に当てはまります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="3b92f-128">ワークシートの使用されている範囲の負の数のカウントを取得するスクリプトについて検討します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="3b92f-129">スクリプトは、使用されている範囲内のすべてのセルを反復処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="3b92f-130">そのためには、範囲、行の数、および列の数が必要です。</span><span class="sxs-lookup"><span data-stu-id="3b92f-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="3b92f-131">ループを開始する前に、これらをローカル変数として格納する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="3b92f-132">それ以外の場合、ループが反復処理されるたびに、ブックが強制的に返されます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="3b92f-133">実験として、でループを置き換えてみてください `usedRangeValues` `usedRange.getValues()` 。</span><span class="sxs-lookup"><span data-stu-id="3b92f-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="3b92f-134">大きな範囲を扱うときに、スクリプトの実行にかかる時間が非常に長くなることがあります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="3b92f-135">不要なステートメントを削除する `console.log`</span><span class="sxs-lookup"><span data-stu-id="3b92f-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="3b92f-136">コンソールログは、[スクリプトをデバッグ](../testing/troubleshooting.md)するための非常に重要なツールです。</span><span class="sxs-lookup"><span data-stu-id="3b92f-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="3b92f-137">ただし、ログに記録された情報が最新であることを確認するために、スクリプトは強制的にブックと同期されます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="3b92f-138">スクリプトを共有する前に、不要なログ記録ステートメント (テストに使用されるものなど) を削除することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="3b92f-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="3b92f-139">これにより、 `console.log()` ステートメントがループ内にない限り、通常、パフォーマンスの問題が発生することはありません。</span><span class="sxs-lookup"><span data-stu-id="3b92f-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="3b92f-140">Try/catch ブロックの使用を避ける</span><span class="sxs-lookup"><span data-stu-id="3b92f-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="3b92f-141">スクリプトの予想される制御フローの一部として[ `try` / `catch` ブロック](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)を使用することはお勧めしません。</span><span class="sxs-lookup"><span data-stu-id="3b92f-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="3b92f-142">ほとんどのエラーは、ブックから返されたオブジェクトをチェックすることで回避できます。</span><span class="sxs-lookup"><span data-stu-id="3b92f-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="3b92f-143">たとえば、次のスクリプトは、ブックから返されたテーブルが存在することを確認してから、行を追加します。</span><span class="sxs-lookup"><span data-stu-id="3b92f-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="3b92f-144">大文字と小文字を区別するヘルプ</span><span class="sxs-lookup"><span data-stu-id="3b92f-144">Case-by-case help</span></span>

<span data-ttu-id="3b92f-145">Office スクリプトプラットフォームが[パワー自動化](https://flow.microsoft.com/)、[アダプティブカード](https://docs.microsoft.com/adaptive-cards)、その他の製品間の機能に拡張されると、スクリプトブックの通信の詳細がより複雑になります。</span><span class="sxs-lookup"><span data-stu-id="3b92f-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="3b92f-146">スクリプトの実行速度を速くする必要がある場合は、[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3b92f-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="3b92f-147">専門家が検索してヘルプを見つけられるように、質問に "office スクリプト" というタグを付けてください。</span><span class="sxs-lookup"><span data-stu-id="3b92f-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="3b92f-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="3b92f-148">See also</span></span>

- [<span data-ttu-id="3b92f-149">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="3b92f-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="3b92f-150">MDN web ドキュメント: ループと反復</span><span class="sxs-lookup"><span data-stu-id="3b92f-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
