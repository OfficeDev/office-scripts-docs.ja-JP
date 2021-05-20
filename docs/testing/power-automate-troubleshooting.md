---
title: Power Automateで実行されているOfficeスクリプトのトラブルシューティング
description: ヒント、プラットフォーム情報、およびOfficeスクリプトとPower Automateの統合に関する既知の問題。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545572"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a><span data-ttu-id="6048b-103">Power Automateで実行されているOfficeスクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="6048b-103">Troubleshoot Office Scripts running in Power Automate</span></span>

<span data-ttu-id="6048b-104">Power Automateを使用すると、Officeスクリプトの自動化を次のレベルに引き上げます。</span><span class="sxs-lookup"><span data-stu-id="6048b-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="6048b-105">ただし、Power Automateは独立したExcel セッションでスクリプトを実行するため、注意すべき重要な点がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="6048b-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="6048b-106">Power Automateでスクリプト Officeを使用し始めたばかりの場合は、Power Automateを使用して[Officeスクリプトを実行](../develop/power-automate-integration.md)してプラットフォームについて学んでください。</span><span class="sxs-lookup"><span data-stu-id="6048b-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-relative-references"></a><span data-ttu-id="6048b-107">相対参照を避ける</span><span class="sxs-lookup"><span data-stu-id="6048b-107">Avoid relative references</span></span>

<span data-ttu-id="6048b-108">Power Automateは、選択したExcelブックでスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="6048b-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="6048b-109">この場合、ブックは閉じられる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="6048b-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="6048b-110">など、ユーザーの現在の状態に依存する API `Workbook.getActiveWorksheet` は、Power Automateで異なる動作をする可能性があります。</span><span class="sxs-lookup"><span data-stu-id="6048b-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="6048b-111">これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照がPower Automateフローに存在しないためです。</span><span class="sxs-lookup"><span data-stu-id="6048b-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="6048b-112">一部の相対参照 API は、Power Automateでエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="6048b-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="6048b-113">その他のユーザーの状態を意味する既定の動作があります。</span><span class="sxs-lookup"><span data-stu-id="6048b-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="6048b-114">スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="6048b-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="6048b-115">これにより、ワークシートが並べ替えられた場合でも、Power Automateフローの一貫性が保たれます。</span><span class="sxs-lookup"><span data-stu-id="6048b-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a><span data-ttu-id="6048b-116">Power Automateフローの実行時に失敗するスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="6048b-116">Script methods that fail when run Power Automate flows</span></span>

<span data-ttu-id="6048b-117">次のメソッドは、Power Automate フローのスクリプトから呼び出されるとエラーをスローし、失敗します。</span><span class="sxs-lookup"><span data-stu-id="6048b-117">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="6048b-118">クラス</span><span class="sxs-lookup"><span data-stu-id="6048b-118">Class</span></span> | <span data-ttu-id="6048b-119">メソッド</span><span class="sxs-lookup"><span data-stu-id="6048b-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="6048b-120">グラフ</span><span class="sxs-lookup"><span data-stu-id="6048b-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="6048b-121">Range</span><span class="sxs-lookup"><span data-stu-id="6048b-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="6048b-122">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="6048b-123">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="6048b-124">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="6048b-125">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="6048b-126">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="6048b-127">Power Automateフローでの既定の動作を持つスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="6048b-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="6048b-128">次のメソッドは、既定の動作を使用します。</span><span class="sxs-lookup"><span data-stu-id="6048b-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="6048b-129">クラス</span><span class="sxs-lookup"><span data-stu-id="6048b-129">Class</span></span> | <span data-ttu-id="6048b-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="6048b-130">Method</span></span> | <span data-ttu-id="6048b-131">Power Automate動作</span><span class="sxs-lookup"><span data-stu-id="6048b-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="6048b-132">ブック</span><span class="sxs-lookup"><span data-stu-id="6048b-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="6048b-133">ブックの最初のワークシート、またはメソッドによって現在アクティブになっているワークシートを返 `Worksheet.activate` します。</span><span class="sxs-lookup"><span data-stu-id="6048b-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="6048b-134">ワークシート</span><span class="sxs-lookup"><span data-stu-id="6048b-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="6048b-135">ワークシートを の目的でアクティブなワークシートとしてマーク `Workbook.getActiveWorksheet` します。</span><span class="sxs-lookup"><span data-stu-id="6048b-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="6048b-136">ファイル ブラウザー コントロールを使用してブックを選択する</span><span class="sxs-lookup"><span data-stu-id="6048b-136">Select workbooks with the file browser control</span></span>

<span data-ttu-id="6048b-137">Power Automate フローの **[スクリプトの実行**] ステップを構築する場合は、フローの一部であるブックを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6048b-137">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="6048b-138">ブックの名前を手動で入力する代わりに、ファイル ブラウザを使用してブックを選択します。</span><span class="sxs-lookup"><span data-stu-id="6048b-138">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[ピッカー ファイル ブラウザーの表示] オプションを表示する [スクリプトの実行] アクションをPower Automate":::

<span data-ttu-id="6048b-140">Power Automateの制限に関する詳細なコンテキストと、ブックの動的選択に対する潜在的な回避策については[、Microsoft Power Automate Communityのこのスレッドを](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)参照してください。</span><span class="sxs-lookup"><span data-stu-id="6048b-140">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="6048b-141">タイム ゾーンの違い</span><span class="sxs-lookup"><span data-stu-id="6048b-141">Time zone differences</span></span>

<span data-ttu-id="6048b-142">Excelファイルには、固有の場所やタイムゾーンがありません。</span><span class="sxs-lookup"><span data-stu-id="6048b-142">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="6048b-143">ユーザーがブックを開くたびに、セッションは日付の計算にユーザーのローカル タイム ゾーンを使用します。</span><span class="sxs-lookup"><span data-stu-id="6048b-143">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="6048b-144">Power Automateは常に UTC を使用します。</span><span class="sxs-lookup"><span data-stu-id="6048b-144">Power Automate always uses UTC.</span></span>

<span data-ttu-id="6048b-145">スクリプトで日付または時刻を使用する場合、スクリプトをローカルでテストするときと、Power Automateを実行する場合と動作に違いが生じる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="6048b-145">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="6048b-146">Power Automateを使用すると、時間の変換、書式設定、および調整ができます。</span><span class="sxs-lookup"><span data-stu-id="6048b-146">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="6048b-147">Power Automateのこれらの関数を使用する方法については、[フロー内の日付と時刻の操作](https://flow.microsoft.com/blog/working-with-dates-and-times/)を参照してください[ `main` 。](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)</span><span class="sxs-lookup"><span data-stu-id="6048b-147">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="6048b-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="6048b-148">See also</span></span>

- [<span data-ttu-id="6048b-149">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="6048b-149">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="6048b-150">Power Automate を使用した Office スクリプトの実行</span><span class="sxs-lookup"><span data-stu-id="6048b-150">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="6048b-151">Excel Online (Business) コネクタ リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="6048b-151">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
