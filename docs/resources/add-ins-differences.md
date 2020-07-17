---
title: Office スクリプトと Office アドインの違い
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: fc2029780190672c633e00e26f44273e4311c754
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878662"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="f8b13-103">Office スクリプトと Office アドインの違い</span><span class="sxs-lookup"><span data-stu-id="f8b13-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="f8b13-104">Office アドインと Office スクリプトには、多くの共通点があります。</span><span class="sxs-lookup"><span data-stu-id="f8b13-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="f8b13-105">どちらも、Excel ブックの JavaScript API の自動制御を提供します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="f8b13-106">ただし、Office スクリプト Api は、Office JavaScript API の特殊な同期バージョンです。</span><span class="sxs-lookup"><span data-stu-id="f8b13-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

![さまざまな Office 機能拡張ソリューションのフォーカス領域を示す4つの領域の図。](../images/office-programmability-diagram.png)

<span data-ttu-id="f8b13-109">Office スクリプトは、作業ウィンドウが開いている間は Office アドインが保持されるのに対して、手動ボタンを押すか、[電源自動化](https://flow.microsoft.com/)で手順として、完了するために実行します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="f8b13-110">これは、アドインがセッション中に状態を維持できることを意味しますが、Office スクリプトでは実行の間に内部状態は保持されません。</span><span class="sxs-lookup"><span data-stu-id="f8b13-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="f8b13-111">Excel 拡張機能がスクリプトプラットフォームの機能を超える必要がある場合は、office アドインの[ドキュメント](/office/dev/add-ins)にアクセスして、office アドインの詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="f8b13-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="f8b13-112">この記事の残りの部分では、Office アドインと Office スクリプトの主な違いについて説明します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="f8b13-113">プラットフォームのサポート</span><span class="sxs-lookup"><span data-stu-id="f8b13-113">Platform Support</span></span>

<span data-ttu-id="f8b13-114">Office アドインはプラットフォーム間で機能します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="f8b13-115">これらは、Windows デスクトップ、Mac、iOS、および web プラットフォーム間で機能し、それぞれに同じ操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="f8b13-116">この点については、個々の API のドキュメントに記載されている例外を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f8b13-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="f8b13-117">Office スクリプトは、現在 web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f8b13-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="f8b13-118">すべての記録、編集、実行は、web プラットフォーム上で実行されます。</span><span class="sxs-lookup"><span data-stu-id="f8b13-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="f8b13-119">API</span><span class="sxs-lookup"><span data-stu-id="f8b13-119">APIs</span></span>

<span data-ttu-id="f8b13-120">Office アドイン用の Office JavaScript Api の同期バージョンはありません。標準の Office スクリプト api はプラットフォームに固有のものであり、パラダイムの使用を避けるために多くの最適化と変更が行われてい `load` / `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="f8b13-120">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="f8b13-121">[Excel JavaScript api](/javascript/api/excel?view=excel-js-preview)の一部は、 [Office スクリプト非同期 api](../develop/excel-async-model.md)と互換性があります。</span><span class="sxs-lookup"><span data-stu-id="f8b13-121">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="f8b13-122">一部のサンプルおよびアドインコードブロックは、 `Excel.run` 最小限の翻訳でブロックに移植できます。</span><span class="sxs-lookup"><span data-stu-id="f8b13-122">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="f8b13-123">2つのプラットフォームは機能を共有していますが、ギャップがあります。</span><span class="sxs-lookup"><span data-stu-id="f8b13-123">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="f8b13-124">Office アドインには、office アドインには含まれませんが、イベントと共通 Api はない2つの主要な API セットがあります。</span><span class="sxs-lookup"><span data-stu-id="f8b13-124">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="f8b13-125">イベント</span><span class="sxs-lookup"><span data-stu-id="f8b13-125">Events</span></span>

<span data-ttu-id="f8b13-126">Office スクリプトは[イベント](/office/dev/add-ins/excel/excel-add-ins-events)をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="f8b13-126">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="f8b13-127">すべてのスクリプトは、コードを1つのメソッドで実行し `main` 、終了します。</span><span class="sxs-lookup"><span data-stu-id="f8b13-127">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="f8b13-128">イベントがトリガーされると再アクティブ化されないため、イベントを登録できません。</span><span class="sxs-lookup"><span data-stu-id="f8b13-128">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="f8b13-129">共通 API</span><span class="sxs-lookup"><span data-stu-id="f8b13-129">Common APIs</span></span>

<span data-ttu-id="f8b13-130">Office スクリプトでは、[共通 api](/javascript/api/office)を使用できません。</span><span class="sxs-lookup"><span data-stu-id="f8b13-130">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="f8b13-131">一般的な Api でのみサポートされている認証、ダイアログウィンドウ、またはその他の機能が必要な場合は、Office のスクリプトではなく、Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8b13-131">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="f8b13-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="f8b13-132">See also</span></span>

- [<span data-ttu-id="f8b13-133">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="f8b13-133">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="f8b13-134">Office スクリプトと VBA マクロの相違点</span><span class="sxs-lookup"><span data-stu-id="f8b13-134">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="f8b13-135">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f8b13-135">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="f8b13-136">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f8b13-136">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
