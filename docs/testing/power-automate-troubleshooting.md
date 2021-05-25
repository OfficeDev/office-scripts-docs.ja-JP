---
title: Officeで実行されているスクリプトのトラブルシューティングPower Automate
description: ヒント、プラットフォーム情報、および既知の問題と、スクリプトとスクリプトのOffice統合Power Automate。
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 3d114b8b9aceb95285ecfc78ddbd868541b9f04c
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631665"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a><span data-ttu-id="4cffc-103">Officeで実行されているスクリプトのトラブルシューティングPower Automate</span><span class="sxs-lookup"><span data-stu-id="4cffc-103">Troubleshoot Office Scripts running in Power Automate</span></span>

<span data-ttu-id="4cffc-104">Power Automateスクリプトオートメーションを次Officeレベルに移動できます。</span><span class="sxs-lookup"><span data-stu-id="4cffc-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="4cffc-105">ただし、Power Automateに独立したセッションでスクリプトを実行Excel、いくつかの重要な点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="4cffc-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="4cffc-106">Power Automate で Office スクリプトを使用する場合は、Office スクリプトを Power Automate で実行[](../develop/power-automate-integration.md)して、プラットフォームについて説明します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-relative-references"></a><span data-ttu-id="4cffc-107">相対参照を避ける</span><span class="sxs-lookup"><span data-stu-id="4cffc-107">Avoid relative references</span></span>

<span data-ttu-id="4cffc-108">Power Automate、選択したブックでスクリプトをExcel代わりに実行します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="4cffc-109">この場合、ブックが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="4cffc-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="4cffc-110">ユーザーの現在の状態 (など) に依存する API は、ユーザーの動作 `Workbook.getActiveWorksheet` が異Power Automate。</span><span class="sxs-lookup"><span data-stu-id="4cffc-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="4cffc-111">これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照がビュー フロー内に存在Power Automateです。</span><span class="sxs-lookup"><span data-stu-id="4cffc-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="4cffc-112">一部の相対参照 API は、エラーをスロー Power Automate。</span><span class="sxs-lookup"><span data-stu-id="4cffc-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="4cffc-113">他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。</span><span class="sxs-lookup"><span data-stu-id="4cffc-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="4cffc-114">スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="4cffc-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="4cffc-115">これにより、ワークシートPower Automate場合でも、一貫性のあるフローを作成できます。</span><span class="sxs-lookup"><span data-stu-id="4cffc-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a><span data-ttu-id="4cffc-116">スクリプト フローで実行すると失敗するスクリプト メソッドPower Automateします。</span><span class="sxs-lookup"><span data-stu-id="4cffc-116">Script methods that fail when run in Power Automate flows</span></span>

<span data-ttu-id="4cffc-117">次のメソッドは、エラーをスローし、エラー フロー内のスクリプトから呼び出Power Automateします。</span><span class="sxs-lookup"><span data-stu-id="4cffc-117">The following methods throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="4cffc-118">クラス</span><span class="sxs-lookup"><span data-stu-id="4cffc-118">Class</span></span> | <span data-ttu-id="4cffc-119">Method</span><span class="sxs-lookup"><span data-stu-id="4cffc-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="4cffc-120">グラフ</span><span class="sxs-lookup"><span data-stu-id="4cffc-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="4cffc-121">Range</span><span class="sxs-lookup"><span data-stu-id="4cffc-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="4cffc-122">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="4cffc-123">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="4cffc-124">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="4cffc-125">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="4cffc-126">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="4cffc-127">スクリプト フローの既定の動作を持つスクリプト メソッドPower Automateします。</span><span class="sxs-lookup"><span data-stu-id="4cffc-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="4cffc-128">次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="4cffc-129">クラス</span><span class="sxs-lookup"><span data-stu-id="4cffc-129">Class</span></span> | <span data-ttu-id="4cffc-130">Method</span><span class="sxs-lookup"><span data-stu-id="4cffc-130">Method</span></span> | <span data-ttu-id="4cffc-131">Power Automate動作</span><span class="sxs-lookup"><span data-stu-id="4cffc-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="4cffc-132">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="4cffc-133">ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="4cffc-134">ワークシート</span><span class="sxs-lookup"><span data-stu-id="4cffc-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="4cffc-135">の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="4cffc-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="data-refresh-not-supported-in-power-automate"></a><span data-ttu-id="4cffc-136">データ更新は、データ更新プログラムではPower Automate</span><span class="sxs-lookup"><span data-stu-id="4cffc-136">Data refresh not supported in Power Automate</span></span>

<span data-ttu-id="4cffc-137">Officeスクリプトは、スクリプトで実行するとデータを更新Power Automate。</span><span class="sxs-lookup"><span data-stu-id="4cffc-137">Office Scripts can't refresh data when run in Power Automate.</span></span> <span data-ttu-id="4cffc-138">フローで呼び `PivotTable.refresh` 出された場合は何もしないなどのメソッド。</span><span class="sxs-lookup"><span data-stu-id="4cffc-138">Methods such as `PivotTable.refresh` do nothing when called in a flow.</span></span> <span data-ttu-id="4cffc-139">さらに、Power Automateリンクを使用する数式のデータ更新はトリガーされません。</span><span class="sxs-lookup"><span data-stu-id="4cffc-139">Additionally, Power Automate doesn't trigger a data refresh for formulas that use workbook links.</span></span>

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a><span data-ttu-id="4cffc-140">スクリプト フローで実行するときに何もしないスクリプト メソッドPower Automateします。</span><span class="sxs-lookup"><span data-stu-id="4cffc-140">Script methods that do nothing when run in Power Automate flows</span></span>

<span data-ttu-id="4cffc-141">次のメソッドは、スクリプトを使用して呼び出した場合、スクリプトPower Automate。</span><span class="sxs-lookup"><span data-stu-id="4cffc-141">The following methods do nothing in a script when called through Power Automate.</span></span> <span data-ttu-id="4cffc-142">それでも正常に返され、エラーはスローしません。</span><span class="sxs-lookup"><span data-stu-id="4cffc-142">They still return successfully and don't throw any errors.</span></span>

| <span data-ttu-id="4cffc-143">クラス</span><span class="sxs-lookup"><span data-stu-id="4cffc-143">Class</span></span> | <span data-ttu-id="4cffc-144">Method</span><span class="sxs-lookup"><span data-stu-id="4cffc-144">Method</span></span> |
|--|--|
| [<span data-ttu-id="4cffc-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="4cffc-145">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [<span data-ttu-id="4cffc-146">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-146">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [<span data-ttu-id="4cffc-147">ブック</span><span class="sxs-lookup"><span data-stu-id="4cffc-147">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [<span data-ttu-id="4cffc-148">ワークシート</span><span class="sxs-lookup"><span data-stu-id="4cffc-148">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="4cffc-149">ファイル ブラウザー コントロールを使用してブックを選択する</span><span class="sxs-lookup"><span data-stu-id="4cffc-149">Select workbooks with the file browser control</span></span>

<span data-ttu-id="4cffc-150">アプリケーション フローの **スクリプトの実行** ステップをPower Automate、フローの一部であるブックを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4cffc-150">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="4cffc-151">ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-151">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[Power Automateファイル ブラウザーの表示] オプションを示すスクリプトの実行アクション":::

<span data-ttu-id="4cffc-153">ブックの動的選択に関するPower Automateの制限と潜在的な回避策に関する詳細なコンテキストについては、Microsoft Power Automate Community のこのスレッドを[参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="4cffc-153">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="4cffc-154">タイム ゾーンの違い</span><span class="sxs-lookup"><span data-stu-id="4cffc-154">Time zone differences</span></span>

<span data-ttu-id="4cffc-155">Excelファイルに固有の場所やタイム ゾーンが存在しない。</span><span class="sxs-lookup"><span data-stu-id="4cffc-155">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="4cffc-156">ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-156">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="4cffc-157">Power Automateは常に UTC を使用します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-157">Power Automate always uses UTC.</span></span>

<span data-ttu-id="4cffc-158">スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と、スクリプトを実行するときに動作の違いPower Automate。</span><span class="sxs-lookup"><span data-stu-id="4cffc-158">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="4cffc-159">Power Automateを使用すると、変換、書式設定、調整を行います。</span><span class="sxs-lookup"><span data-stu-id="4cffc-159">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="4cffc-160">Power Automate[](https://flow.microsoft.com/blog/working-with-dates-and-times/)および[ `main` Parameters: Pass](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) data to a script でこれらの関数を使用する方法については、「フロー内の日付と時刻の操作」を参照して、スクリプトの時間情報を提供する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4cffc-160">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="4cffc-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="4cffc-161">See also</span></span>

- [<span data-ttu-id="4cffc-162">スクリプトOfficeトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="4cffc-162">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="4cffc-163">Power Automate を使用した Office スクリプトの実行</span><span class="sxs-lookup"><span data-stu-id="4cffc-163">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="4cffc-164">Excel Online (Business) コネクタ リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="4cffc-164">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
