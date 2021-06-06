---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API Office違い。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 46f5f2ea6fea15e9506f5c7d30941311fc2e669e
ms.sourcegitcommit: 0bfc9472d107e32c804029659317f8e81fec5d19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779364"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="54195-103">Office スクリプトと Office アドインの違い</span><span class="sxs-lookup"><span data-stu-id="54195-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="54195-104">各スクリプトと OfficeアドインOfficeの違いを理解し、各アドインをいつ使用する必要が生じ得るのかについて理解します。</span><span class="sxs-lookup"><span data-stu-id="54195-104">Understand the differences between Office Scripts and Office Add-ins to know when to use each one.</span></span> <span data-ttu-id="54195-105">Officeスクリプトは、ワークフローの改善を探しているすべてのユーザーが迅速に作成するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="54195-105">Office Scripts are designed to be quickly made by anyone looking to improve their workflow.</span></span> <span data-ttu-id="54195-106">Officeアドインは、リボン ボタンと作業ウィンドウOffice対話型の UI と統合します。</span><span class="sxs-lookup"><span data-stu-id="54195-106">Office Add-ins integrate with the Office UI for a more interactive experience through ribbon buttons and task panes.</span></span> <span data-ttu-id="54195-107">Officeアドインは、カスタム関数を提供することで、組み込Excel機能を拡張できます。</span><span class="sxs-lookup"><span data-stu-id="54195-107">Office Add-ins can also expand built-in Excel functions by providing custom functions.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当て、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象とします)":::

<span data-ttu-id="54195-109">Officeスクリプトは手動でボタンを押して実行するか[、Power Automate](https://flow.microsoft.com/)でステップとして実行しますが、Office アドインは構成方法に応じて実行を続行します。</span><span class="sxs-lookup"><span data-stu-id="54195-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins continue running depending on how they are configured.</span></span> <span data-ttu-id="54195-110">たとえば、作業ウィンドウが閉Office実行を続行するアドインを構成できます。</span><span class="sxs-lookup"><span data-stu-id="54195-110">For example, you can configure an Office Add-in to continue running even when its task pane is closed.</span></span> <span data-ttu-id="54195-111">つまり、Officeアドインはセッション中に状態を維持しますが、Officeスクリプトは実行の間に内部状態を維持します。</span><span class="sxs-lookup"><span data-stu-id="54195-111">This means that Office Add-ins maintain state during a session, whereas Office Scripts don't maintain an internal state between runs.</span></span> <span data-ttu-id="54195-112">構築するソリューションに保守状態が必要な場合は、Office アドインの[](/office/dev/add-ins)ドキュメントを参照して、Office アドインの詳細を確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="54195-112">If the solution you are building requires a maintained state, you should visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="54195-113">この記事の残りの部分では、アドインとスクリプトの主なOfficeについてOfficeします。</span><span class="sxs-lookup"><span data-stu-id="54195-113">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="54195-114">プラットフォームサポート</span><span class="sxs-lookup"><span data-stu-id="54195-114">Platform Support</span></span>

<span data-ttu-id="54195-115">Officeアドインはクロスプラットフォームです。</span><span class="sxs-lookup"><span data-stu-id="54195-115">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="54195-116">デスクトップ、Mac、Windows Web プラットフォーム間で動作し、それぞれで同じエクスペリエンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="54195-116">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="54195-117">この例外は、個々の API のドキュメントに示されています。</span><span class="sxs-lookup"><span data-stu-id="54195-117">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="54195-118">Officeスクリプトは現在、ユーザーがサポートしているExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="54195-118">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="54195-119">すべての記録、編集、および実行は、Web プラットフォーム上で行われます。</span><span class="sxs-lookup"><span data-stu-id="54195-119">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="54195-120">API</span><span class="sxs-lookup"><span data-stu-id="54195-120">APIs</span></span>

<span data-ttu-id="54195-121">OfficeアドインOffice Office スクリプト API の JavaScript API はいくつかの機能を共有しますが、プラットフォームは異なります。</span><span class="sxs-lookup"><span data-stu-id="54195-121">While the Office JavaScript APIs for Office Add-ins and the Office Scripts APIs share some functionality, they are different platforms.</span></span> <span data-ttu-id="54195-122">スクリプト Office API は、JavaScript API モデルの最適化された同期Excelサブセットです。</span><span class="sxs-lookup"><span data-stu-id="54195-122">The Office Scripts APIs are an optimized, synchronous subset of the Excel JavaScript API model.</span></span> <span data-ttu-id="54195-123">大きな違いは、アドイン `load` / `sync` でのパラダイムの使用です。さらに、アドインはイベント用の API と、共通 API と呼ばれる Excel以外の広範な機能セットを提供します。</span><span class="sxs-lookup"><span data-stu-id="54195-123">The major difference is usage of the `load`/`sync` paradigm with add-ins. Additionally, add-ins offer APIs for events and a broader set of functionality outside of Excel, known as the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="54195-124">イベント</span><span class="sxs-lookup"><span data-stu-id="54195-124">Events</span></span>

<span data-ttu-id="54195-125">Officeスクリプトは、ブック レベルのイベントを[サポートしていない](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="54195-125">Office Scripts do not support workbook-level [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="54195-126">スクリプトは、ユーザーがスクリプトの [**実行**] ボタンを押すか、スクリプトを使用してトリガー Power Automate。</span><span class="sxs-lookup"><span data-stu-id="54195-126">Scripts are either triggered by users pressing the **Run** button for a script or through Power Automate.</span></span> <span data-ttu-id="54195-127">すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。</span><span class="sxs-lookup"><span data-stu-id="54195-127">Every script runs the code in a single `main` method, then ends.</span></span>

### <a name="common-apis"></a><span data-ttu-id="54195-128">共通 API</span><span class="sxs-lookup"><span data-stu-id="54195-128">Common APIs</span></span>

<span data-ttu-id="54195-129">Officeスクリプトで共通[API を使用することはできません](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="54195-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="54195-130">一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="54195-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="54195-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="54195-131">See also</span></span>

- [<span data-ttu-id="54195-132">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="54195-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="54195-133">スクリプトと VBA Officeの違い</span><span class="sxs-lookup"><span data-stu-id="54195-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="54195-134">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="54195-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="54195-135">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="54195-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
