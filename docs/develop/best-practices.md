---
title: Office スクリプトでのベスト プラクティス
description: 一般的な問題を防止し、予期しない入力またはデータOfficeできる堅牢なスクリプトを記述する方法。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546032"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="8131d-103">Office スクリプトでのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8131d-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="8131d-104">これらのパターンとプラクティスは、スクリプトが毎回正常に実行されるのを助けるために設計されています。</span><span class="sxs-lookup"><span data-stu-id="8131d-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="8131d-105">ワークフローの自動化を開始する場合、一般的な落とし穴をExcelしてください。</span><span class="sxs-lookup"><span data-stu-id="8131d-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="8131d-106">オブジェクトが存在する確認</span><span class="sxs-lookup"><span data-stu-id="8131d-106">Verify an object is present</span></span>

<span data-ttu-id="8131d-107">スクリプトは、多くの場合、ブックに存在する特定のワークシートまたはテーブルに依存します。</span><span class="sxs-lookup"><span data-stu-id="8131d-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="8131d-108">ただし、スクリプトの実行の間に名前が変更または削除される場合があります。</span><span class="sxs-lookup"><span data-stu-id="8131d-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="8131d-109">これらのテーブルまたはワークシートがメソッドを呼び出す前に存在する場合は、スクリプトが突然終了しないか確認できます。</span><span class="sxs-lookup"><span data-stu-id="8131d-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="8131d-110">次のサンプル コードは、ブックに "Index" ワークシートが存在する場合にチェックします。</span><span class="sxs-lookup"><span data-stu-id="8131d-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="8131d-111">ワークシートが存在する場合、スクリプトは範囲を取得して続行します。</span><span class="sxs-lookup"><span data-stu-id="8131d-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="8131d-112">存在しない場合、スクリプトはカスタム エラー メッセージをログに記録します。</span><span class="sxs-lookup"><span data-stu-id="8131d-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="8131d-113">TypeScript 演算子 `?` は、メソッドを呼び出す前にオブジェクトが存在するかどうかをチェックします。</span><span class="sxs-lookup"><span data-stu-id="8131d-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="8131d-114">これにより、オブジェクトが存在しない場合に特別な操作を行う必要が生じなかった場合に、コードの効率化が図れる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="8131d-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="8131d-115">データとブックの状態を最初に検証する</span><span class="sxs-lookup"><span data-stu-id="8131d-115">Validate data and workbook state first</span></span>

<span data-ttu-id="8131d-116">データを操作する前に、すべてのワークシート、テーブル、図形、その他のオブジェクトが存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8131d-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="8131d-117">前のパターンを使用して、すべてがブック内にあるか確認し、期待に一致します。</span><span class="sxs-lookup"><span data-stu-id="8131d-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="8131d-118">データが書き込まれる前にこれを行って、スクリプトがブックを部分的な状態に残すのを確認します。</span><span class="sxs-lookup"><span data-stu-id="8131d-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="8131d-119">次のスクリプトでは、"Table1" と "Table2" という名前の 2 つのテーブルが存在する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8131d-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="8131d-120">スクリプトは、最初にテーブルが存在する場合はチェックし、ステートメントで終わり、存在しない場合は適切なメッセージ `return` で終わります。</span><span class="sxs-lookup"><span data-stu-id="8131d-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="8131d-121">検証が別の関数で行っている場合でも、関数からステートメントを発行してスクリプト `return` を終了する必要 `main` があります。</span><span class="sxs-lookup"><span data-stu-id="8131d-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="8131d-122">サブ関数から戻しても、スクリプトは終了しない。</span><span class="sxs-lookup"><span data-stu-id="8131d-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="8131d-123">次のスクリプトは、前のスクリプトと同じ動作をします。</span><span class="sxs-lookup"><span data-stu-id="8131d-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="8131d-124">違いは、関数が `main` 関数を呼び出 `inputPresent` してすべてを検証する点です。</span><span class="sxs-lookup"><span data-stu-id="8131d-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="8131d-125">`inputPresent` 必要なすべての入力が存在するかどうかを示すブール値 ( `true` `false` または ) を返します。</span><span class="sxs-lookup"><span data-stu-id="8131d-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="8131d-126">関数 `main` は、そのブール値を使用して、スクリプトの継続または終了を決定します。</span><span class="sxs-lookup"><span data-stu-id="8131d-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="8131d-127">ステートメントを使用する `throw` 場合</span><span class="sxs-lookup"><span data-stu-id="8131d-127">When to use a `throw` statement</span></span>

<span data-ttu-id="8131d-128">ステートメント [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) は、予期しないエラーが発生したかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="8131d-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="8131d-129">コードは直ちに終了します。</span><span class="sxs-lookup"><span data-stu-id="8131d-129">It ends the code immediately.</span></span> <span data-ttu-id="8131d-130">ほとんどの場合、スクリプトから実行 `throw` する必要はなんらない。</span><span class="sxs-lookup"><span data-stu-id="8131d-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="8131d-131">通常、スクリプトは、問題が原因でスクリプトの実行に失敗したとユーザーに自動的に通知します。</span><span class="sxs-lookup"><span data-stu-id="8131d-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="8131d-132">ほとんどの場合、エラー メッセージと関数のステートメントでスクリプトを終了しても `return` 十分 `main` です。</span><span class="sxs-lookup"><span data-stu-id="8131d-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="8131d-133">ただし、スクリプトがプロセス フローの一部として実行されているPower Automate、フローの続行を停止できます。</span><span class="sxs-lookup"><span data-stu-id="8131d-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="8131d-134">ステートメント `throw` はスクリプトを停止し、フローにも停止を指示します。</span><span class="sxs-lookup"><span data-stu-id="8131d-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="8131d-135">次のスクリプトは、テーブルチェックの例 `throw` でステートメントを使用する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8131d-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="8131d-136">ステートメントを使用する `try...catch` 場合</span><span class="sxs-lookup"><span data-stu-id="8131d-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="8131d-137">この [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) ステートメントは、API 呼び出しが失敗した場合に検出し、スクリプトの実行を続行する方法です。</span><span class="sxs-lookup"><span data-stu-id="8131d-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="8131d-138">範囲に対して大規模なデータ更新を実行する次のスニペットを検討してください。</span><span class="sxs-lookup"><span data-stu-id="8131d-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="8131d-139">Web `someLargeValues` が処理できるExcelより大きい場合、呼び出し `setValues()` は失敗します。</span><span class="sxs-lookup"><span data-stu-id="8131d-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="8131d-140">その後、ランタイム エラーが発生してスクリプト [も失敗します](../testing/troubleshooting.md#runtime-errors)。</span><span class="sxs-lookup"><span data-stu-id="8131d-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="8131d-141">このステートメントを使用すると、スクリプトをすぐに終了して既定のエラーを表示することなく、スクリプトで `try...catch` この条件を認識できます。</span><span class="sxs-lookup"><span data-stu-id="8131d-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="8131d-142">スクリプト ユーザーに優れたエクスペリエンスを提供する方法の 1 つは、カスタム エラー メッセージを表示する方法です。</span><span class="sxs-lookup"><span data-stu-id="8131d-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="8131d-143">次のスニペットは、読者に `try...catch` 役立つエラー情報をログに記録するステートメントを示しています。</span><span class="sxs-lookup"><span data-stu-id="8131d-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="8131d-144">エラーを処理するもう 1 つの方法は、エラー ケースを処理するフォールバック動作を持つ方法です。</span><span class="sxs-lookup"><span data-stu-id="8131d-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="8131d-145">次のスニペットでは、ブロックを使用して別のメソッドを試して、更新プログラムを小さな部分に分割し、エラー `catch` を回避します。</span><span class="sxs-lookup"><span data-stu-id="8131d-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="8131d-146">大きな範囲を更新する方法の完全な例については、「大きなデータセットを記述 [する」を参照してください](../resources/samples/write-large-dataset.md)。</span><span class="sxs-lookup"><span data-stu-id="8131d-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="8131d-147">ループ `try...catch` の内側または周囲を使用すると、スクリプトの速度が低下します。</span><span class="sxs-lookup"><span data-stu-id="8131d-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="8131d-148">パフォーマンスの詳細については、「ブロックの使用 [を避ける」を `try...catch` 参照してください](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)。</span><span class="sxs-lookup"><span data-stu-id="8131d-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="8131d-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="8131d-149">See also</span></span>

- [<span data-ttu-id="8131d-150">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="8131d-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="8131d-151">Office スクリプトを使用した Power Automate のトラブルシューティング情報</span><span class="sxs-lookup"><span data-stu-id="8131d-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="8131d-152">スクリプトを使用したプラットフォームOffice制限</span><span class="sxs-lookup"><span data-stu-id="8131d-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="8131d-153">スクリプトのパフォーマンスをOfficeする</span><span class="sxs-lookup"><span data-stu-id="8131d-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
