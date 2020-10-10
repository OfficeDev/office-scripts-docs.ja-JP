---
title: Office スクリプトを使用したプラットフォームの制限と要件
description: Web 上の Excel で使用する場合の Office スクリプトのリソース制限とブラウザーサポート
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: df468192f443b912e26411e46c9f953e046e55ec
ms.sourcegitcommit: 42fa3b629c93930b4e73e9c4c01d0c8bdf6d7487
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/09/2020
ms.locfileid: "48411558"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="cbeef-103">Office スクリプトを使用したプラットフォームの制限と要件</span><span class="sxs-lookup"><span data-stu-id="cbeef-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="cbeef-104">Office スクリプトを開発する際には、いくつかのプラットフォームの制限事項に注意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbeef-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="cbeef-105">この記事では、web 上の Excel 用 Office スクリプトのブラウザーのサポートとデータの制限について説明します。</span><span class="sxs-lookup"><span data-stu-id="cbeef-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="cbeef-106">ブラウザのサポート</span><span class="sxs-lookup"><span data-stu-id="cbeef-106">Browser support</span></span>

<span data-ttu-id="cbeef-107">Office スクリプト [は、web 用の office をサポート](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)する任意のブラウザーで動作します。</span><span class="sxs-lookup"><span data-stu-id="cbeef-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="cbeef-108">ただし、一部の JavaScript 機能は Internet Explorer 11 (IE 11) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cbeef-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="cbeef-109">ES6 以降で導入された機能は、IE 11 で [は](https://www.w3schools.com/Js/js_es6.asp) 動作しません。</span><span class="sxs-lookup"><span data-stu-id="cbeef-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="cbeef-110">組織内のユーザーが依然としてそのブラウザーを使用している場合は、その環境でスクリプトを共有するときに必ずテストしてください。</span><span class="sxs-lookup"><span data-stu-id="cbeef-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="cbeef-111">サードパーティの cookie</span><span class="sxs-lookup"><span data-stu-id="cbeef-111">Third-party cookies</span></span>

<span data-ttu-id="cbeef-112">ブラウザーでは、web 上の Excel で [ **自動化** ] タブが表示されるように、サードパーティの cookie が有効になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbeef-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="cbeef-113">タブが表示されていない場合は、ブラウザーの設定を確認します。</span><span class="sxs-lookup"><span data-stu-id="cbeef-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="cbeef-114">プライベートブラウザーセッションを使用している場合は、この設定を毎回有効にしなければならない場合があります。</span><span class="sxs-lookup"><span data-stu-id="cbeef-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="cbeef-115">一部のブラウザーは、"サードパーティの cookie" ではなく "すべての cookie" としてこの設定を参照します。</span><span class="sxs-lookup"><span data-stu-id="cbeef-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="cbeef-116">データの上限</span><span class="sxs-lookup"><span data-stu-id="cbeef-116">Data limits</span></span>

<span data-ttu-id="cbeef-117">一度に転送できる Excel データの量と、実行できる個々の電力を自動化するトランザクションの数には制限があります。</span><span class="sxs-lookup"><span data-stu-id="cbeef-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="cbeef-118">Excel</span><span class="sxs-lookup"><span data-stu-id="cbeef-118">Excel</span></span>

<span data-ttu-id="cbeef-119">スクリプトを使用してブックを呼び出すときに、web 用の Excel には次の制限があります。</span><span class="sxs-lookup"><span data-stu-id="cbeef-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="cbeef-120">要求と応答は **5 mb**に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cbeef-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="cbeef-121">範囲は **500万のセル**に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cbeef-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="cbeef-122">大規模なデータセットを処理するときにエラーが発生した場合は、大きな範囲ではなく、複数の狭い範囲を使用してください。</span><span class="sxs-lookup"><span data-stu-id="cbeef-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="cbeef-123">範囲外の [セル](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) のような api を使用して、大きな範囲ではなく特定のセルを対象にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="cbeef-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="cbeef-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="cbeef-124">Power Automate</span></span>

<span data-ttu-id="cbeef-125">Office スクリプトを電源自動化と共に使用する場合、1 **日あたりの通話**は最大200に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cbeef-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="cbeef-126">この制限は、12:00 AM UTC でリセットされます。</span><span class="sxs-lookup"><span data-stu-id="cbeef-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="cbeef-127">Power 自動プラットフォームにも使用上の制限があります。これは、「 [Power 自動検出の制限と構成](/power-automate/limits-and-config)」に記載されています。</span><span class="sxs-lookup"><span data-stu-id="cbeef-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="cbeef-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="cbeef-128">See also</span></span>

- [<span data-ttu-id="cbeef-129">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="cbeef-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="cbeef-130">Office スクリプトの効果を元に戻す</span><span class="sxs-lookup"><span data-stu-id="cbeef-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="cbeef-131">Office スクリプトのパフォーマンスを向上させる</span><span class="sxs-lookup"><span data-stu-id="cbeef-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="cbeef-132">Web 上の Excel での Office スクリプトのスクリプトの基礎</span><span class="sxs-lookup"><span data-stu-id="cbeef-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
