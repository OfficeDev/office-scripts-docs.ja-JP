---
title: Office スクリプトを使用した Power Automate のトラブルシューティング情報
description: Office Scripts と Power Automate の統合に関するヒント、プラットフォーム情報、既知の問題。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933267"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a><span data-ttu-id="dff46-103">Office スクリプトを使用した Power Automate のトラブルシューティング情報</span><span class="sxs-lookup"><span data-stu-id="dff46-103">Troubleshooting information for Power Automate with Office Scripts</span></span>

<span data-ttu-id="dff46-104">Power Automate を使用すると、Officeスクリプトの自動化を次のレベルに進めできます。</span><span class="sxs-lookup"><span data-stu-id="dff46-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="dff46-105">ただし、Power Automate は独立した Excel セッションでユーザーに代わってスクリプトを実行しますが、注意が必要ないくつかの重要な点があります。</span><span class="sxs-lookup"><span data-stu-id="dff46-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="dff46-106">Power Automate で Office スクリプトを使い始め始めたばかりの場合は、Power Automate を使った [Office スクリプト](../develop/power-automate-integration.md) の実行から始め、プラットフォームについて学習してください。</span><span class="sxs-lookup"><span data-stu-id="dff46-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="dff46-107">相対参照の使用を避ける</span><span class="sxs-lookup"><span data-stu-id="dff46-107">Avoid using relative references</span></span>

<span data-ttu-id="dff46-108">Power Automate は、選択した Excel ブックでスクリプトをユーザーに代わって実行します。</span><span class="sxs-lookup"><span data-stu-id="dff46-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="dff46-109">この場合、ブックが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="dff46-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="dff46-110">Power Automate では、ユーザーの現在の状態に依存する API (など) の動作 `Workbook.getActiveWorksheet` が異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="dff46-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="dff46-111">これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照が Power Automate フローに存在しないためです。</span><span class="sxs-lookup"><span data-stu-id="dff46-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="dff46-112">一部の相対参照 API は Power Automate でエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="dff46-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="dff46-113">他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。</span><span class="sxs-lookup"><span data-stu-id="dff46-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="dff46-114">スクリプトを設計する場合は、ワークシートと範囲の絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="dff46-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="dff46-115">これにより、ワークシートが再配置された場合でも、Power Automate フローの一貫性が維持されます。</span><span class="sxs-lookup"><span data-stu-id="dff46-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a><span data-ttu-id="dff46-116">Power Automate フローの実行時に失敗するスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="dff46-116">Script methods that fail when run Power Automate flows</span></span>

<span data-ttu-id="dff46-117">次のメソッドは、Power Automate フローのスクリプトから呼び出された場合にエラーをスローし、失敗します。</span><span class="sxs-lookup"><span data-stu-id="dff46-117">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="dff46-118">クラス</span><span class="sxs-lookup"><span data-stu-id="dff46-118">Class</span></span> | <span data-ttu-id="dff46-119">Method</span><span class="sxs-lookup"><span data-stu-id="dff46-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="dff46-120">Chart</span><span class="sxs-lookup"><span data-stu-id="dff46-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="dff46-121">Range</span><span class="sxs-lookup"><span data-stu-id="dff46-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="dff46-122">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="dff46-123">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="dff46-124">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="dff46-125">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="dff46-126">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="dff46-127">Power Automate フローでの既定の動作を持つスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="dff46-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="dff46-128">次のメソッドは、ユーザーの現在の状態の代用として、既定の動作を使用します。</span><span class="sxs-lookup"><span data-stu-id="dff46-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="dff46-129">クラス</span><span class="sxs-lookup"><span data-stu-id="dff46-129">Class</span></span> | <span data-ttu-id="dff46-130">Method</span><span class="sxs-lookup"><span data-stu-id="dff46-130">Method</span></span> | <span data-ttu-id="dff46-131">Power Automate の動作</span><span class="sxs-lookup"><span data-stu-id="dff46-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="dff46-132">Workbook</span><span class="sxs-lookup"><span data-stu-id="dff46-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="dff46-133">ブック内の最初のワークシート、またはメソッドによって現在アクティブになっているワークシートを返 `Worksheet.activate` します。</span><span class="sxs-lookup"><span data-stu-id="dff46-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="dff46-134">Worksheet</span><span class="sxs-lookup"><span data-stu-id="dff46-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="dff46-135">次の目的のために、ワークシートをアクティブ ワークシートとしてマークします `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="dff46-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="dff46-136">ファイル ブラウザー コントロールを使用してブックを選択する</span><span class="sxs-lookup"><span data-stu-id="dff46-136">Select workbooks with the file browser control</span></span>

<span data-ttu-id="dff46-137">Power Automate フロー **のスクリプト** 実行ステップを作成する場合は、フローの一部であるブックを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dff46-137">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="dff46-138">ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。</span><span class="sxs-lookup"><span data-stu-id="dff46-138">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

![Power Automate で "スクリプトの実行" アクションを作成する場合のファイル ブラウザー オプション](../images/power-automate-file-browser.png)

<span data-ttu-id="dff46-140">Power Automate の制限の詳細と、ブックの動的な選択に対する潜在的な回避策については、Microsoft Power Automate コミュニティのこのスレッド [を参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="dff46-140">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="dff46-141">タイム ゾーンの違い</span><span class="sxs-lookup"><span data-stu-id="dff46-141">Time zone differences</span></span>

<span data-ttu-id="dff46-142">Excel ファイルには、固有の場所やタイム ゾーンが存在しない。</span><span class="sxs-lookup"><span data-stu-id="dff46-142">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="dff46-143">ユーザーがブックを開くたび、セッションは日付の計算にユーザーのローカルのタイムゾーンを使用します。</span><span class="sxs-lookup"><span data-stu-id="dff46-143">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="dff46-144">Power Automate は常に UTC を使用します。</span><span class="sxs-lookup"><span data-stu-id="dff46-144">Power Automate always uses UTC.</span></span>

<span data-ttu-id="dff46-145">スクリプトで日付や時刻を使用する場合、スクリプトをローカルでテストする場合と Power Automate を使用してスクリプトを実行する場合とで動作が異なる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="dff46-145">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="dff46-146">Power Automate を使用すると、変換、書式設定、および時間の調整を行います。</span><span class="sxs-lookup"><span data-stu-id="dff46-146">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="dff46-147">Power [](https://flow.microsoft.com/blog/working-with-dates-and-times/) Automate と Parameters でこれらの関数を使用する方法の手順については、「フロー内の日付と時刻を操作する[ `main` :](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script)スクリプトにデータを渡す」を参照して、その時間情報をスクリプトに提供する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="dff46-147">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="dff46-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="dff46-148">See also</span></span>

- [<span data-ttu-id="dff46-149">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="dff46-149">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="dff46-150">Power Automate Officeスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="dff46-150">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="dff46-151">Excel Online (Business) コネクタのリファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="dff46-151">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
