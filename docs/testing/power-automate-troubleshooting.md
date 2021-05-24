---
title: Officeで実行されているスクリプトのトラブルシューティングPower Automate
description: ヒント、プラットフォーム情報、および既知の問題と、スクリプトとスクリプトのOffice統合Power Automate。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545572"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a><span data-ttu-id="2bd70-103">Officeで実行されているスクリプトのトラブルシューティングPower Automate</span><span class="sxs-lookup"><span data-stu-id="2bd70-103">Troubleshoot Office Scripts running in Power Automate</span></span>

<span data-ttu-id="2bd70-104">Power Automateスクリプトオートメーションを次Officeレベルに移動できます。</span><span class="sxs-lookup"><span data-stu-id="2bd70-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="2bd70-105">ただし、Power Automateに独立したセッションでスクリプトを実行Excel、いくつかの重要な点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2bd70-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="2bd70-106">Power Automate で Office スクリプトを使用する場合は、Office スクリプトを Power Automate で実行[](../develop/power-automate-integration.md)して、プラットフォームについて説明します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-relative-references"></a><span data-ttu-id="2bd70-107">相対参照を避ける</span><span class="sxs-lookup"><span data-stu-id="2bd70-107">Avoid relative references</span></span>

<span data-ttu-id="2bd70-108">Power Automate、選択したブックでスクリプトをExcel代わりに実行します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="2bd70-109">この場合、ブックが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="2bd70-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="2bd70-110">ユーザーの現在の状態 (など) に依存する API は、ユーザーの動作 `Workbook.getActiveWorksheet` が異Power Automate。</span><span class="sxs-lookup"><span data-stu-id="2bd70-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="2bd70-111">これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照がビュー フロー内に存在Power Automateです。</span><span class="sxs-lookup"><span data-stu-id="2bd70-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="2bd70-112">一部の相対参照 API は、エラーをスロー Power Automate。</span><span class="sxs-lookup"><span data-stu-id="2bd70-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="2bd70-113">他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。</span><span class="sxs-lookup"><span data-stu-id="2bd70-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="2bd70-114">スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="2bd70-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="2bd70-115">これにより、ワークシートPower Automate場合でも、一貫性のあるフローを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2bd70-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a><span data-ttu-id="2bd70-116">フローの実行時に失敗するスクリプト メソッドPower Automateします。</span><span class="sxs-lookup"><span data-stu-id="2bd70-116">Script methods that fail when run Power Automate flows</span></span>

<span data-ttu-id="2bd70-117">次のメソッドは、エラーをスローし、エラー フロー内のスクリプトから呼び出Power Automateします。</span><span class="sxs-lookup"><span data-stu-id="2bd70-117">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="2bd70-118">クラス</span><span class="sxs-lookup"><span data-stu-id="2bd70-118">Class</span></span> | <span data-ttu-id="2bd70-119">メソッド</span><span class="sxs-lookup"><span data-stu-id="2bd70-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="2bd70-120">グラフ</span><span class="sxs-lookup"><span data-stu-id="2bd70-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="2bd70-121">Range</span><span class="sxs-lookup"><span data-stu-id="2bd70-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="2bd70-122">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="2bd70-123">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="2bd70-124">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="2bd70-125">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="2bd70-126">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="2bd70-127">スクリプト フローの既定の動作を持つスクリプト メソッドPower Automateします。</span><span class="sxs-lookup"><span data-stu-id="2bd70-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="2bd70-128">次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="2bd70-129">クラス</span><span class="sxs-lookup"><span data-stu-id="2bd70-129">Class</span></span> | <span data-ttu-id="2bd70-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="2bd70-130">Method</span></span> | <span data-ttu-id="2bd70-131">Power Automate動作</span><span class="sxs-lookup"><span data-stu-id="2bd70-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="2bd70-132">ブック</span><span class="sxs-lookup"><span data-stu-id="2bd70-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="2bd70-133">ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="2bd70-134">ワークシート</span><span class="sxs-lookup"><span data-stu-id="2bd70-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="2bd70-135">の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="2bd70-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="2bd70-136">ファイル ブラウザー コントロールを使用してブックを選択する</span><span class="sxs-lookup"><span data-stu-id="2bd70-136">Select workbooks with the file browser control</span></span>

<span data-ttu-id="2bd70-137">アプリケーション フローの **スクリプトの実行** ステップをPower Automate、フローの一部であるブックを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2bd70-137">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="2bd70-138">ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-138">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[Power Automateファイル ブラウザーの表示] オプションを示すスクリプトの実行アクション":::

<span data-ttu-id="2bd70-140">ブックの動的選択に関するPower Automateの制限と潜在的な回避策に関する詳細なコンテキストについては、Microsoft Power Automate Community のこのスレッドを[参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="2bd70-140">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="2bd70-141">タイム ゾーンの違い</span><span class="sxs-lookup"><span data-stu-id="2bd70-141">Time zone differences</span></span>

<span data-ttu-id="2bd70-142">Excelファイルに固有の場所やタイム ゾーンが存在しない。</span><span class="sxs-lookup"><span data-stu-id="2bd70-142">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="2bd70-143">ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-143">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="2bd70-144">Power Automateは常に UTC を使用します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-144">Power Automate always uses UTC.</span></span>

<span data-ttu-id="2bd70-145">スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と、スクリプトを実行するときに動作の違いPower Automate。</span><span class="sxs-lookup"><span data-stu-id="2bd70-145">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="2bd70-146">Power Automateを使用すると、変換、書式設定、調整を行います。</span><span class="sxs-lookup"><span data-stu-id="2bd70-146">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="2bd70-147">Power Automate[](https://flow.microsoft.com/blog/working-with-dates-and-times/)および[ `main` Parameters: Pass](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) data to a script でこれらの関数を使用する方法については、「フロー内の日付と時刻の操作」を参照して、スクリプトの時間情報を提供する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="2bd70-147">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="2bd70-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="2bd70-148">See also</span></span>

- [<span data-ttu-id="2bd70-149">スクリプトOfficeトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="2bd70-149">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="2bd70-150">Power Automate を使用した Office スクリプトの実行</span><span class="sxs-lookup"><span data-stu-id="2bd70-150">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="2bd70-151">Excel Online (Business) コネクタ リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="2bd70-151">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
