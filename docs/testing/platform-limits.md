---
title: プラットフォームの制限と要件 (スクリプトOffice)
description: スクリプトと一緒に使用する場合Officeスクリプトのリソース制限とブラウザー Excel on the web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545582"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="88f7f-103">プラットフォームの制限と要件 (スクリプトOffice)</span><span class="sxs-lookup"><span data-stu-id="88f7f-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="88f7f-104">スクリプトの開発時に注意する必要があるプラットフォームのOfficeがあります。</span><span class="sxs-lookup"><span data-stu-id="88f7f-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="88f7f-105">この記事では、ブラウザーのサポートとデータ制限について詳OfficeスクリプトのExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="88f7f-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="88f7f-106">ブラウザのサポート</span><span class="sxs-lookup"><span data-stu-id="88f7f-106">Browser support</span></span>

<span data-ttu-id="88f7f-107">Officeスクリプトは、Web のサポート[をサポートOfficeブラウザーで動作します](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)。</span><span class="sxs-lookup"><span data-stu-id="88f7f-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="88f7f-108">ただし、JavaScript の一部の機能は、11 Internet Explorer (IE 11) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88f7f-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="88f7f-109">[ES6](https://www.w3schools.com/Js/js_es6.asp)以降で導入された機能は、IE 11 では動作しません。</span><span class="sxs-lookup"><span data-stu-id="88f7f-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="88f7f-110">組織内のユーザーが引き続きそのブラウザーを使用している場合は、共有するときに、その環境でスクリプトをテストしてください。</span><span class="sxs-lookup"><span data-stu-id="88f7f-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="88f7f-111">サードパーティの Cookie</span><span class="sxs-lookup"><span data-stu-id="88f7f-111">Third-party cookies</span></span>

<span data-ttu-id="88f7f-112">ブラウザーで [自動化] タブを表示するには、サードパーティの Cookie が **有効になっている必要** Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="88f7f-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="88f7f-113">タブが表示されていない場合は、ブラウザーの設定を確認します。</span><span class="sxs-lookup"><span data-stu-id="88f7f-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="88f7f-114">プライベート ブラウザー セッションを使用している場合は、その度にこの設定を再び有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="88f7f-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="88f7f-115">一部のブラウザーでは、この設定を "サードパーティ Cookie" ではなく"すべての Cookie" と呼ぶ場合があります。</span><span class="sxs-lookup"><span data-stu-id="88f7f-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="88f7f-116">一般的なブラウザーで Cookie 設定を調整する手順</span><span class="sxs-lookup"><span data-stu-id="88f7f-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="88f7f-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="88f7f-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="88f7f-118">Edge</span><span class="sxs-lookup"><span data-stu-id="88f7f-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="88f7f-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="88f7f-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="88f7f-120">Safari</span><span class="sxs-lookup"><span data-stu-id="88f7f-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="88f7f-121">データの上限</span><span class="sxs-lookup"><span data-stu-id="88f7f-121">Data limits</span></span>

<span data-ttu-id="88f7f-122">データを一度にExcelできるデータの量と、トランザクションを実行できるPower Automate制限があります。</span><span class="sxs-lookup"><span data-stu-id="88f7f-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="88f7f-123">Excel</span><span class="sxs-lookup"><span data-stu-id="88f7f-123">Excel</span></span>

<span data-ttu-id="88f7f-124">Excelを使用してブックを呼び出す場合、Web の場合、次の制限があります。</span><span class="sxs-lookup"><span data-stu-id="88f7f-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="88f7f-125">要求と応答は **5 MB に制限されています**。</span><span class="sxs-lookup"><span data-stu-id="88f7f-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="88f7f-126">範囲は 500 万 **セルに制限されます**。</span><span class="sxs-lookup"><span data-stu-id="88f7f-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="88f7f-127">大規模なデータセットを扱う際にエラーが発生する場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。</span><span class="sxs-lookup"><span data-stu-id="88f7f-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="88f7f-128">例については、「大規模なデータセットを書 [き込む」サンプルを参照](../resources/samples/write-large-dataset.md) してください。</span><span class="sxs-lookup"><span data-stu-id="88f7f-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="88f7f-129">[Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-)のような API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="88f7f-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="88f7f-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="88f7f-130">Power Automate</span></span>

<span data-ttu-id="88f7f-131">ユーザーが OfficeスクリプトPower Automate使用する場合、各ユーザーは 1 日にスクリプトの実行アクションに対して **400 回の呼び出しに制限されます**。</span><span class="sxs-lookup"><span data-stu-id="88f7f-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="88f7f-132">この制限は、UTC の午前 12:00 にリセットされます。</span><span class="sxs-lookup"><span data-stu-id="88f7f-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="88f7f-133">このPower Automateには使用上の制限があります。これは次の記事で確認できます。</span><span class="sxs-lookup"><span data-stu-id="88f7f-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="88f7f-134">サーバーの制限と構成Power Automate</span><span class="sxs-lookup"><span data-stu-id="88f7f-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="88f7f-135">オンライン (Business) コネクタExcel既知の問題と制限事項</span><span class="sxs-lookup"><span data-stu-id="88f7f-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="88f7f-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="88f7f-136">See also</span></span>

- [<span data-ttu-id="88f7f-137">スクリプトOfficeトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="88f7f-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="88f7f-138">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="88f7f-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="88f7f-139">スクリプトのパフォーマンスをOfficeする</span><span class="sxs-lookup"><span data-stu-id="88f7f-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="88f7f-140">スクリプトの基本OfficeスクリプトのExcel on the web</span><span class="sxs-lookup"><span data-stu-id="88f7f-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
