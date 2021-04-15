---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API のOffice違い。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755099"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="b80bb-103">Office スクリプトと Office アドインの違い</span><span class="sxs-lookup"><span data-stu-id="b80bb-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="b80bb-104">OfficeアドインとスクリプトOffice共通点が多い。</span><span class="sxs-lookup"><span data-stu-id="b80bb-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="b80bb-105">どちらも JavaScript API の Excel ブックの自動制御を提供します。</span><span class="sxs-lookup"><span data-stu-id="b80bb-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="b80bb-106">ただし、Officeスクリプト API は、JavaScript API の特殊な同期Officeです。</span><span class="sxs-lookup"><span data-stu-id="b80bb-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当てしていますが、Office スクリプトはエンド ユーザーに対応します (Office Web アドインはプロフェッショナル開発者を対象としています)。":::

<span data-ttu-id="b80bb-108">Officeスクリプトは、手動ボタンを押して完了するか [、Power Automate](https://flow.microsoft.com/)のステップとして実行しますが、Office アドインは作業ウィンドウが開いている間も保持されます。</span><span class="sxs-lookup"><span data-stu-id="b80bb-108">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="b80bb-109">つまり、アドインはセッション中に状態を維持できるのに対し、Officeスクリプトは実行の間に内部状態を維持できません。</span><span class="sxs-lookup"><span data-stu-id="b80bb-109">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="b80bb-110">Excel 拡張機能がスクリプト プラットフォームの機能を超える必要がある場合は [、Office](/office/dev/add-ins) アドインのドキュメントを参照して、Office アドインの詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="b80bb-110">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="b80bb-111">この記事の残りの部分では、アドインとスクリプトの主な違OfficeについてOfficeします。</span><span class="sxs-lookup"><span data-stu-id="b80bb-111">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="b80bb-112">プラットフォームサポート</span><span class="sxs-lookup"><span data-stu-id="b80bb-112">Platform Support</span></span>

<span data-ttu-id="b80bb-113">Officeはクロスプラットフォームです。</span><span class="sxs-lookup"><span data-stu-id="b80bb-113">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="b80bb-114">Windows デスクトップ、Mac、iOS、および Web プラットフォーム間で動作し、それぞれに同じエクスペリエンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b80bb-114">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="b80bb-115">この例外は、個々の API のドキュメントに示されています。</span><span class="sxs-lookup"><span data-stu-id="b80bb-115">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="b80bb-116">Officeスクリプトは現在、Web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b80bb-116">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="b80bb-117">すべての記録、編集、および実行は、Web プラットフォーム上で行われます。</span><span class="sxs-lookup"><span data-stu-id="b80bb-117">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="b80bb-118">API</span><span class="sxs-lookup"><span data-stu-id="b80bb-118">APIs</span></span>

<span data-ttu-id="b80bb-119">アドイン用の JavaScript API Office同期バージョンOfficeはありません。標準のOfficeスクリプト API はプラットフォームに固有であり、パラダイムの使用を避けるための多数の最適化と変更 `load` / `sync` があります。</span><span class="sxs-lookup"><span data-stu-id="b80bb-119">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="b80bb-120">[Excel JavaScript API の](/javascript/api/excel?view=excel-js-preview&preserve-view=true)一部は、スクリプト非同期 API Office[互換性があります](../develop/excel-async-model.md)。</span><span class="sxs-lookup"><span data-stu-id="b80bb-120">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="b80bb-121">一部のサンプルとアドイン コード ブロックは、最小限の変換でブロック `Excel.run` に移植できます。</span><span class="sxs-lookup"><span data-stu-id="b80bb-121">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="b80bb-122">2 つのプラットフォームは機能を共有しますが、ギャップがあります。</span><span class="sxs-lookup"><span data-stu-id="b80bb-122">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="b80bb-123">2 つの主要な API セットは、Officeが含まれていますが、スクリプトOfficeイベントと共通 API ではありません。</span><span class="sxs-lookup"><span data-stu-id="b80bb-123">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="b80bb-124">イベント</span><span class="sxs-lookup"><span data-stu-id="b80bb-124">Events</span></span>

<span data-ttu-id="b80bb-125">Officeスクリプトはイベントをサポート [していない](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="b80bb-125">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="b80bb-126">すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。</span><span class="sxs-lookup"><span data-stu-id="b80bb-126">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="b80bb-127">イベントがトリガーされると再アクティブ化されないので、イベントを登録できません。</span><span class="sxs-lookup"><span data-stu-id="b80bb-127">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="b80bb-128">共通 API</span><span class="sxs-lookup"><span data-stu-id="b80bb-128">Common APIs</span></span>

<span data-ttu-id="b80bb-129">Officeは共通 [API を使用できません](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="b80bb-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="b80bb-130">一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b80bb-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="b80bb-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="b80bb-131">See also</span></span>

- [<span data-ttu-id="b80bb-132">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="b80bb-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="b80bb-133">スクリプトと VBA Officeの違い</span><span class="sxs-lookup"><span data-stu-id="b80bb-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="b80bb-134">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="b80bb-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="b80bb-135">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b80bb-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
