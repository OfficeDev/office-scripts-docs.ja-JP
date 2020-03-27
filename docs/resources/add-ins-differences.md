---
title: Office スクリプトと Office アドインの相違点
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978729"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="43faf-103">Office スクリプトと Office アドインの相違点</span><span class="sxs-lookup"><span data-stu-id="43faf-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="43faf-104">Office アドインと Office スクリプトには、多くの共通点があります。</span><span class="sxs-lookup"><span data-stu-id="43faf-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="43faf-105">どちらも、Office JavaScript API の名前空間を`Excel`使用して、Excel ブックの自動制御を提供します。</span><span class="sxs-lookup"><span data-stu-id="43faf-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="43faf-106">ただし、Office スクリプトの範囲は、より制限されています。</span><span class="sxs-lookup"><span data-stu-id="43faf-106">However, Office Scripts are more limited in their scope.</span></span>

![さまざまな Office 機能拡張ソリューションのフォーカス領域を示す4つの領域の図。](../images/office-programmability-diagram.png)

<span data-ttu-id="43faf-109">Office スクリプトは、作業ウィンドウが開いている間は Office アドインが保持されるのに対して、手動ボタンを押すか、[電源自動化](https://flow.microsoft.com/)で手順として、完了するために実行します。</span><span class="sxs-lookup"><span data-stu-id="43faf-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="43faf-110">これは、アドインがセッション中に状態を維持できることを意味しますが、Office スクリプトでは実行の間に内部状態は保持されません。</span><span class="sxs-lookup"><span data-stu-id="43faf-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="43faf-111">Excel 拡張機能がスクリプトプラットフォームの機能を超える必要がある場合は、office アドインの[ドキュメント](/office/dev/add-ins)にアクセスして、office アドインの詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="43faf-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="43faf-112">この記事の残りの部分では、Office アドインと Office スクリプトの主な違いについて説明します。</span><span class="sxs-lookup"><span data-stu-id="43faf-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="43faf-113">プラットフォームのサポート</span><span class="sxs-lookup"><span data-stu-id="43faf-113">Platform Support</span></span>

<span data-ttu-id="43faf-114">Office アドインはプラットフォーム間で機能します。</span><span class="sxs-lookup"><span data-stu-id="43faf-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="43faf-115">これらは、Windows デスクトップ、Mac、iOS、および web プラットフォーム間で機能し、それぞれに同じ操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="43faf-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="43faf-116">この点については、個々の API のドキュメントに記載されている例外を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43faf-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="43faf-117">Office スクリプトは、現在 web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="43faf-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="43faf-118">すべての記録、編集、実行は、web プラットフォーム上で実行されます。</span><span class="sxs-lookup"><span data-stu-id="43faf-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="43faf-119">API</span><span class="sxs-lookup"><span data-stu-id="43faf-119">APIs</span></span>

<span data-ttu-id="43faf-120">Office スクリプトは、ほとんどの Excel JavaScript Api をサポートしています。これは、2つのプラットフォーム間で多くの機能が重なっていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="43faf-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="43faf-121">2つの例外として、イベントと共通 Api があります。</span><span class="sxs-lookup"><span data-stu-id="43faf-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="43faf-122">イベント</span><span class="sxs-lookup"><span data-stu-id="43faf-122">Events</span></span>

<span data-ttu-id="43faf-123">Office スクリプトは[イベント](/office/dev/add-ins/excel/excel-add-ins-events)をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="43faf-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="43faf-124">すべてのスクリプトは、コードを 1 `main`つのメソッドで実行し、終了します。</span><span class="sxs-lookup"><span data-stu-id="43faf-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="43faf-125">イベントがトリガーされると再アクティブ化されないため、イベントを登録できません。</span><span class="sxs-lookup"><span data-stu-id="43faf-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="43faf-126">共通 API</span><span class="sxs-lookup"><span data-stu-id="43faf-126">Common APIs</span></span>

<span data-ttu-id="43faf-127">Office スクリプトでは、[共通 api](/javascript/api/office)を使用できません。</span><span class="sxs-lookup"><span data-stu-id="43faf-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="43faf-128">一般的な Api でのみサポートされている認証、ダイアログウィンドウ、またはその他の機能が必要な場合は、Office のスクリプトではなく、Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43faf-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="43faf-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="43faf-129">See also</span></span>

- [<span data-ttu-id="43faf-130">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="43faf-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="43faf-131">Office スクリプトと VBA マクロの相違点</span><span class="sxs-lookup"><span data-stu-id="43faf-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="43faf-132">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="43faf-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="43faf-133">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="43faf-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
