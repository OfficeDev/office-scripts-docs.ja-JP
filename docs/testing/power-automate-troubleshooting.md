---
title: Power Automate with Office スクリプト
description: スクリプトと Power Automate の間の統合に関するヒント、プラットフォーム情報、既知Office問題。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755008"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a><span data-ttu-id="e81bc-103">Power Automate with Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="e81bc-103">Troubleshooting information for Power Automate with Office Scripts</span></span>

<span data-ttu-id="e81bc-104">Power Automation を使用すると、スクリプトOfficeを次のレベルに進めできます。</span><span class="sxs-lookup"><span data-stu-id="e81bc-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="e81bc-105">ただし、Power Automate は独立した Excel セッションでスクリプトを代理で実行しますので、いくつかの重要な点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e81bc-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="e81bc-106">Power Automate を使用して Office スクリプトを使い始める場合は、「Power [Automate](../develop/power-automate-integration.md) を使用して Office スクリプトを実行する」から始め、プラットフォームについて学習してください。</span><span class="sxs-lookup"><span data-stu-id="e81bc-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="e81bc-107">相対参照の使用を避ける</span><span class="sxs-lookup"><span data-stu-id="e81bc-107">Avoid using relative references</span></span>

<span data-ttu-id="e81bc-108">Power Automate は、選択した Excel ブックでスクリプトを代理で実行します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="e81bc-109">この場合、ブックが閉じられます。</span><span class="sxs-lookup"><span data-stu-id="e81bc-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="e81bc-110">Power Automate では、ユーザーの現在の状態 (など) に依存する API の動作 `Workbook.getActiveWorksheet` が異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="e81bc-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="e81bc-111">これは、API がユーザーのビューまたはカーソルの相対位置に基づいており、その参照が Power Automate フローに存在しないのでです。</span><span class="sxs-lookup"><span data-stu-id="e81bc-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="e81bc-112">一部の相対参照 API は、Power Automate でエラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="e81bc-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="e81bc-113">他のユーザーは、ユーザーの状態を意味する既定の動作を持っています。</span><span class="sxs-lookup"><span data-stu-id="e81bc-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="e81bc-114">スクリプトを設計する場合は、ワークシートと範囲に絶対参照を使用してください。</span><span class="sxs-lookup"><span data-stu-id="e81bc-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="e81bc-115">これにより、ワークシートが再配置された場合でも、Power Automate フローの整合性が保たれる。</span><span class="sxs-lookup"><span data-stu-id="e81bc-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a><span data-ttu-id="e81bc-116">Power Automate フローの実行時に失敗するスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="e81bc-116">Script methods that fail when run Power Automate flows</span></span>

<span data-ttu-id="e81bc-117">次のメソッドは、Power Automate フローのスクリプトから呼び出された場合にエラーをスローし、失敗します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-117">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="e81bc-118">クラス</span><span class="sxs-lookup"><span data-stu-id="e81bc-118">Class</span></span> | <span data-ttu-id="e81bc-119">Method</span><span class="sxs-lookup"><span data-stu-id="e81bc-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="e81bc-120">Chart</span><span class="sxs-lookup"><span data-stu-id="e81bc-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="e81bc-121">Range</span><span class="sxs-lookup"><span data-stu-id="e81bc-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="e81bc-122">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="e81bc-123">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="e81bc-124">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="e81bc-125">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="e81bc-126">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="e81bc-127">Power Automate フローの既定の動作を持つスクリプト メソッド</span><span class="sxs-lookup"><span data-stu-id="e81bc-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="e81bc-128">次のメソッドは、ユーザーの現在の状態の代りとして、既定の動作を使用します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="e81bc-129">クラス</span><span class="sxs-lookup"><span data-stu-id="e81bc-129">Class</span></span> | <span data-ttu-id="e81bc-130">Method</span><span class="sxs-lookup"><span data-stu-id="e81bc-130">Method</span></span> | <span data-ttu-id="e81bc-131">Power Automate の動作</span><span class="sxs-lookup"><span data-stu-id="e81bc-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="e81bc-132">Workbook</span><span class="sxs-lookup"><span data-stu-id="e81bc-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="e81bc-133">ブックの最初のワークシート、またはメソッドによって現在アクティブ化されているワークシートのいずれかを返 `Worksheet.activate` します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="e81bc-134">Worksheet</span><span class="sxs-lookup"><span data-stu-id="e81bc-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="e81bc-135">の目的でワークシートをアクティブなワークシートとしてマークします `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="e81bc-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="e81bc-136">ファイル ブラウザー コントロールを使用してブックを選択する</span><span class="sxs-lookup"><span data-stu-id="e81bc-136">Select workbooks with the file browser control</span></span>

<span data-ttu-id="e81bc-137">Power Automate フローの **スクリプト** の実行ステップを作成する場合は、フローの一部であるブックを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e81bc-137">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="e81bc-138">ブックの名前を手動で入力する代わりに、ファイル ブラウザーを使用してブックを選択します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-138">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="[選択ウィンドウのファイル ブラウザーを表示する] オプションを示す Power Automate Run スクリプト アクション。":::

<span data-ttu-id="e81bc-140">Power Automate の制限に関する詳細なコンテキストと、ブックの動的選択に関する潜在的な回避策の説明については、Microsoft Power Automate コミュニティのこのスレッド [を参照してください](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="e81bc-140">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="e81bc-141">タイム ゾーンの違い</span><span class="sxs-lookup"><span data-stu-id="e81bc-141">Time zone differences</span></span>

<span data-ttu-id="e81bc-142">Excel ファイルには、固有の場所やタイム ゾーンが存在しない。</span><span class="sxs-lookup"><span data-stu-id="e81bc-142">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="e81bc-143">ユーザーがブックを開くたび、そのユーザーのローカル タイム ゾーンを日付の計算に使用します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-143">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="e81bc-144">Power Automate は常に UTC を使用します。</span><span class="sxs-lookup"><span data-stu-id="e81bc-144">Power Automate always uses UTC.</span></span>

<span data-ttu-id="e81bc-145">スクリプトで日付または時刻を使用する場合、スクリプトがローカルでテストされる場合と Power Automate を使用して実行する場合の動作の違いがあります。</span><span class="sxs-lookup"><span data-stu-id="e81bc-145">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="e81bc-146">Power Automate を使用すると、変換、書式設定、および調整を行います。</span><span class="sxs-lookup"><span data-stu-id="e81bc-146">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="e81bc-147">「Power [Automate」](https://flow.microsoft.com/blog/working-with-dates-and-times/)および[ `main` 「Parameters: Passing](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) data to a script」のこれらの関数の使い方については、「フロー内の日付と時刻を操作する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e81bc-147">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="e81bc-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="e81bc-148">See also</span></span>

- [<span data-ttu-id="e81bc-149">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e81bc-149">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="e81bc-150">Power Automate Officeスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="e81bc-150">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="e81bc-151">Excel Online (Business) コネクタリファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="e81bc-151">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
