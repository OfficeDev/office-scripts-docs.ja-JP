---
title: Office スクリプトを使用したプラットフォームの制限と要件
description: Excel on the webで使用する場合のリソース制限とOfficeスクリプトのブラウザサポート
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545582"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="4512a-103">Office スクリプトを使用したプラットフォームの制限と要件</span><span class="sxs-lookup"><span data-stu-id="4512a-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="4512a-104">Officeスクリプトを開発する際に注意する必要があるプラットフォームの制限事項がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="4512a-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="4512a-105">この記事では、ブラウザーのサポートとExcel on the web用のOffice スクリプトのデータ制限について詳しく説明します。</span><span class="sxs-lookup"><span data-stu-id="4512a-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="4512a-106">ブラウザのサポート</span><span class="sxs-lookup"><span data-stu-id="4512a-106">Browser support</span></span>

<span data-ttu-id="4512a-107">Officeスクリプトは、web の[Officeをサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。</span><span class="sxs-lookup"><span data-stu-id="4512a-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="4512a-108">ただし、一部の JavaScript 機能は、インターネット エクスプ ローラー 11 (IE 11) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4512a-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="4512a-109">[ES6 以降](https://www.w3schools.com/Js/js_es6.asp)で導入された機能は、IE 11 では動作しません。</span><span class="sxs-lookup"><span data-stu-id="4512a-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="4512a-110">組織のユーザーがそのブラウザーを引き続き使用している場合は、スクリプトを共有するときに必ずその環境でスクリプトをテストしてください。</span><span class="sxs-lookup"><span data-stu-id="4512a-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="4512a-111">サードパーティのクッキー</span><span class="sxs-lookup"><span data-stu-id="4512a-111">Third-party cookies</span></span>

<span data-ttu-id="4512a-112">お使いのブラウザでは、Excel on the webの[**自動化**]タブを表示するためにサードパーティのクッキーを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4512a-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="4512a-113">タブが表示されていない場合は、ブラウザの設定を確認してください。</span><span class="sxs-lookup"><span data-stu-id="4512a-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="4512a-114">プライベートブラウザセッションを使用している場合は、毎回この設定を再度有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4512a-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="4512a-115">ブラウザによっては、この設定を「サードパーティのクッキー」ではなく「すべてのクッキー」と呼んでいます。</span><span class="sxs-lookup"><span data-stu-id="4512a-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="4512a-116">一般的なブラウザでクッキーの設定を調整する手順</span><span class="sxs-lookup"><span data-stu-id="4512a-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="4512a-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="4512a-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="4512a-118">Edge</span><span class="sxs-lookup"><span data-stu-id="4512a-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="4512a-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="4512a-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="4512a-120">Safari</span><span class="sxs-lookup"><span data-stu-id="4512a-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="4512a-121">データの上限</span><span class="sxs-lookup"><span data-stu-id="4512a-121">Data limits</span></span>

<span data-ttu-id="4512a-122">一度に転送できるデータExcel量と、個々のPower Automateトランザクションの数には制限があります。</span><span class="sxs-lookup"><span data-stu-id="4512a-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="4512a-123">Excel</span><span class="sxs-lookup"><span data-stu-id="4512a-123">Excel</span></span>

<span data-ttu-id="4512a-124">スクリプトを使用してブックを呼び出す場合、web のExcelには次の制限があります。</span><span class="sxs-lookup"><span data-stu-id="4512a-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="4512a-125">要求と応答は 5 **MB** に制限されています。</span><span class="sxs-lookup"><span data-stu-id="4512a-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="4512a-126">範囲は **500 万個のセル** に制限されます。</span><span class="sxs-lookup"><span data-stu-id="4512a-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="4512a-127">大きなデータセットを扱うときにエラーが発生した場合は、より大きな範囲ではなく、複数の小さい範囲を使用してみてください。</span><span class="sxs-lookup"><span data-stu-id="4512a-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="4512a-128">例については、「 [大規模なデータセットの記述サンプル」](../resources/samples/write-large-dataset.md) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4512a-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="4512a-129">[Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-)などの API を使用して、大きな範囲ではなく特定のセルをターゲットにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="4512a-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="4512a-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="4512a-130">Power Automate</span></span>

<span data-ttu-id="4512a-131">Power AutomateでOfficeスクリプトを使用する場合、各ユーザーは **1 日あたりスクリプトの実行アクションに対して 400 回の呼び出しを行う** 必要があります。</span><span class="sxs-lookup"><span data-stu-id="4512a-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="4512a-132">この制限は、UTC の午前 12 時にリセットされます。</span><span class="sxs-lookup"><span data-stu-id="4512a-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="4512a-133">Power Automate プラットフォームには、次の記事で説明する使用制限もあります。</span><span class="sxs-lookup"><span data-stu-id="4512a-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="4512a-134">Power Automateの制限と構成</span><span class="sxs-lookup"><span data-stu-id="4512a-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="4512a-135">Excel オンライン (ビジネス) コネクタの既知の問題と制限事項</span><span class="sxs-lookup"><span data-stu-id="4512a-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="4512a-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="4512a-136">See also</span></span>

- [<span data-ttu-id="4512a-137">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="4512a-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="4512a-138">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="4512a-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="4512a-139">Officeスクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="4512a-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="4512a-140">Excel on the webでのスクリプトのスクリプトOfficeの基礎</span><span class="sxs-lookup"><span data-stu-id="4512a-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
