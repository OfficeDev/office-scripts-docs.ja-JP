---
title: Office スクリプトを使用したプラットフォームの制限と要件
description: Web 上の Excel で使用する場合の Office スクリプトのリソース制限とブラウザーサポート
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930079"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="e1132-103">Office スクリプトを使用したプラットフォームの制限と要件</span><span class="sxs-lookup"><span data-stu-id="e1132-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="e1132-104">Office スクリプトを開発する際には、いくつかのプラットフォームの制限事項に注意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1132-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="e1132-105">この記事では、web 上の Excel 用 Office スクリプトのブラウザーのサポートとデータの制限について説明します。</span><span class="sxs-lookup"><span data-stu-id="e1132-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="e1132-106">ブラウザのサポート</span><span class="sxs-lookup"><span data-stu-id="e1132-106">Browser support</span></span>

<span data-ttu-id="e1132-107">Office スクリプト [は、web 用の office をサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。</span><span class="sxs-lookup"><span data-stu-id="e1132-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="e1132-108">ただし、一部の JavaScript 機能は Internet Explorer 11 (IE 11) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e1132-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="e1132-109">ES6 以降で導入された機能は、IE 11 で [は](https://www.w3schools.com/Js/js_es6.asp) 動作しません。</span><span class="sxs-lookup"><span data-stu-id="e1132-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="e1132-110">組織内のユーザーが依然としてそのブラウザーを使用している場合は、その環境でスクリプトを共有するときに必ずテストしてください。</span><span class="sxs-lookup"><span data-stu-id="e1132-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="e1132-111">サードパーティの cookie</span><span class="sxs-lookup"><span data-stu-id="e1132-111">Third-party cookies</span></span>

<span data-ttu-id="e1132-112">ブラウザーでは、web 上の Excel で [ **自動化** ] タブが表示されるように、サードパーティの cookie が有効になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e1132-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="e1132-113">タブが表示されていない場合は、ブラウザーの設定を確認します。</span><span class="sxs-lookup"><span data-stu-id="e1132-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="e1132-114">プライベートブラウザーセッションを使用している場合は、この設定を毎回有効にしなければならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="e1132-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="e1132-115">一部のブラウザーは、"サードパーティの cookie" ではなく "すべての cookie" としてこの設定を参照します。</span><span class="sxs-lookup"><span data-stu-id="e1132-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="e1132-116">一般的なブラウザーで cookie 設定を調整するための手順</span><span class="sxs-lookup"><span data-stu-id="e1132-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="e1132-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="e1132-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="e1132-118">Edge</span><span class="sxs-lookup"><span data-stu-id="e1132-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="e1132-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="e1132-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="e1132-120">Safari</span><span class="sxs-lookup"><span data-stu-id="e1132-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="e1132-121">データの上限</span><span class="sxs-lookup"><span data-stu-id="e1132-121">Data limits</span></span>

<span data-ttu-id="e1132-122">一度に転送できる Excel データの量と、実行できる個々の電力を自動化するトランザクションの数には制限があります。</span><span class="sxs-lookup"><span data-stu-id="e1132-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="e1132-123">Excel</span><span class="sxs-lookup"><span data-stu-id="e1132-123">Excel</span></span>

<span data-ttu-id="e1132-124">スクリプトを使用してブックを呼び出すときに、web 用の Excel には次の制限があります。</span><span class="sxs-lookup"><span data-stu-id="e1132-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="e1132-125">要求と応答は **5 mb** に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e1132-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="e1132-126">範囲は **500万のセル** に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e1132-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="e1132-127">大規模なデータセットを処理するときにエラーが発生した場合は、大きな範囲ではなく、複数の狭い範囲を使用してください。</span><span class="sxs-lookup"><span data-stu-id="e1132-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="e1132-128">範囲外の [セル](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) のような api を使用して、大きな範囲ではなく特定のセルを対象にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="e1132-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="e1132-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="e1132-129">Power Automate</span></span>

<span data-ttu-id="e1132-130">Office スクリプトを電源自動化と共に使用する場合、1 **日あたりの通話** は最大200に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e1132-130">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="e1132-131">この制限は、12:00 AM UTC でリセットされます。</span><span class="sxs-lookup"><span data-stu-id="e1132-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="e1132-132">Power 自動プラットフォームにも使用上の制限があります。これは、「 [Power 自動検出の制限と構成](/power-automate/limits-and-config)」に記載されています。</span><span class="sxs-lookup"><span data-stu-id="e1132-132">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="e1132-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="e1132-133">See also</span></span>

- [<span data-ttu-id="e1132-134">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e1132-134">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="e1132-135">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="e1132-135">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="e1132-136">Office スクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="e1132-136">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="e1132-137">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="e1132-137">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
