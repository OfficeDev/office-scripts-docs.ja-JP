---
title: Officeスクリプトのパフォーマンスを向上させる
description: Excelブックとスクリプトの間の通信を理解して、より高速なスクリプトを作成します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544992"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="e11a3-103">Officeスクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="e11a3-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="e11a3-104">Officeスクリプトの目的は、一般的に実行される一連のタスクを自動化して時間を節約することです。</span><span class="sxs-lookup"><span data-stu-id="e11a3-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="e11a3-105">低速なスクリプトは、ワークフローを高速化できないような感じがします。</span><span class="sxs-lookup"><span data-stu-id="e11a3-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="e11a3-106">ほとんどの場合、スクリプトは完全に問題なく、期待どおりに実行されます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="e11a3-107">ただし、パフォーマンスに影響を与える可能性のある回避可能なシナリオがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="e11a3-108">スクリプトが遅い最も一般的な理由は、ブックとの通信が過剰である場合です。</span><span class="sxs-lookup"><span data-stu-id="e11a3-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="e11a3-109">スクリプトはローカル コンピューターで実行されますが、ブックはクラウドに存在します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="e11a3-110">スクリプトは、特定の時間に、ローカル データとブックのデータを同期します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="e11a3-111">つまり、書き込み操作 ( など `workbook.addWorksheet()` ) は、このバックグラウンド同期が行われる場合にのみブックに適用されます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="e11a3-112">同様に、読み取り操作 ( など `myRange.getValues()` ) は、その時点でスクリプトのワークブックからデータを取得するだけです。</span><span class="sxs-lookup"><span data-stu-id="e11a3-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="e11a3-113">いずれの場合も、スクリプトはデータに対して動作する前に情報をフェッチします。</span><span class="sxs-lookup"><span data-stu-id="e11a3-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="e11a3-114">たとえば、次のコードは、使用範囲内の行数を正確に記録します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="e11a3-115">Officeスクリプト API は、ワークブックまたはスクリプト内のデータが正確で、必要に応じて最新であることを保証します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="e11a3-116">スクリプトを正しく実行するために、これらの同期について心配する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e11a3-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="e11a3-117">ただし、このスクリプトからクラウドへの通信を認識すると、不要なネットワーク呼び出しを回避できます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="e11a3-118">パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="e11a3-118">Performance optimizations</span></span>

<span data-ttu-id="e11a3-119">クラウドへの通信を減らすのに役立つ簡単な手法を適用できます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="e11a3-120">次のパターンは、スクリプトの高速化に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="e11a3-121">ループ内で繰り返しブック データを読み取るのではなく、ブックデータを 1 回読み取ります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="e11a3-122">不要な `console.log` ステートメントを削除します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="e11a3-123">try/catch ブロックの使用は避けてください。</span><span class="sxs-lookup"><span data-stu-id="e11a3-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="e11a3-124">ループの外側でブック データを読み取る</span><span class="sxs-lookup"><span data-stu-id="e11a3-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="e11a3-125">ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーできます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="e11a3-126">同じ呼び出しを繰り返し行うのではなく、可能な限りデータをローカルに保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="e11a3-127">これは、ループを扱う場合に特に当てはまります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="e11a3-128">ワークシートの使用範囲内の負の数を取得するスクリプトを検討してください。</span><span class="sxs-lookup"><span data-stu-id="e11a3-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="e11a3-129">スクリプトは、使用範囲内のすべてのセルを反復処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="e11a3-130">そのためには、範囲、行数、列数が必要です。</span><span class="sxs-lookup"><span data-stu-id="e11a3-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="e11a3-131">ループを開始する前に、ローカル変数として格納する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="e11a3-132">そうしないと、ループの各反復処理によってブックに強制的に戻ります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="e11a3-133">実験として、 `usedRangeValues` ループ内でを に置き換えて `usedRange.getValues()` みます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="e11a3-134">大きな範囲を扱う場合、スクリプトの実行にかなり時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="e11a3-135">`try...catch`ループ内または周囲のループでブロックを使用しないようにする</span><span class="sxs-lookup"><span data-stu-id="e11a3-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="e11a3-136">[`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)ループまたは周囲のループでステートメントを使用することはお勧めしません。</span><span class="sxs-lookup"><span data-stu-id="e11a3-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="e11a3-137">これは、ループ内のデータの読み取りを避ける必要があるのと同じ理由です: 各反復処理では、スクリプトがブックと同期してエラーがスローされないように強制します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="e11a3-138">ほとんどのエラーは、ブックから返されたオブジェクトをチェックすることで回避できます。</span><span class="sxs-lookup"><span data-stu-id="e11a3-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="e11a3-139">たとえば、次のスクリプトは、行を追加する前に、ブックによって返されたテーブルが存在することを確認します。</span><span class="sxs-lookup"><span data-stu-id="e11a3-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="e11a3-140">不要な `console.log` ステートメントを削除する</span><span class="sxs-lookup"><span data-stu-id="e11a3-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="e11a3-141">コンソールログはスクリプトをデバッグするための重要なツール [です](../testing/troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="e11a3-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="e11a3-142">ただし、スクリプトは、記録された情報が最新であることを確認するために、ブックと同期する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="e11a3-143">スクリプトを共有する前に、不要なログ記録ステートメント (テストに使用するものなど) を削除することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="e11a3-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="e11a3-144">通常、ステートメントがループ内にある場合を除き、パフォーマンスに問題 `console.log()` はありません。</span><span class="sxs-lookup"><span data-stu-id="e11a3-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="e11a3-145">ケースバイケースのヘルプ</span><span class="sxs-lookup"><span data-stu-id="e11a3-145">Case-by-case help</span></span>

<span data-ttu-id="e11a3-146">Officeスクリプト プラットフォームが[Power Automate、](https://flow.microsoft.com/)[アダプティブ カード](/adaptive-cards)、その他の製品間機能を使用するように拡張するにつれて、スクリプトとワークブックの通信の詳細がより複雑になります。</span><span class="sxs-lookup"><span data-stu-id="e11a3-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="e11a3-147">スクリプトの実行速度を上げるためのヘルプが必要な場合は、 [Microsoft Q&A](/answers/topics/office-scripts-dev.html)を通じて連絡を取ってください。</span><span class="sxs-lookup"><span data-stu-id="e11a3-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="e11a3-148">専門家がそれを見つけて助けることができるように、あなたの質問に「オフィススクリプト-dev」とタグを付けてください。</span><span class="sxs-lookup"><span data-stu-id="e11a3-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="e11a3-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="e11a3-149">See also</span></span>

- [<span data-ttu-id="e11a3-150">Excel on the web での Office スクリプトのスクリプトの基本事項</span><span class="sxs-lookup"><span data-stu-id="e11a3-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="e11a3-151">MDN Web ドキュメント: ループと反復</span><span class="sxs-lookup"><span data-stu-id="e11a3-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
