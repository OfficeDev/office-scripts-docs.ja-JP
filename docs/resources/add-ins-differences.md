---
title: Office スクリプトと Office アドインの違い
description: スクリプトとアドインの動作Office API Office違い。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232635"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="c4468-103">Office スクリプトと Office アドインの違い</span><span class="sxs-lookup"><span data-stu-id="c4468-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="c4468-104">Officeアドインとカスタム スクリプトOffice共通点が多い。</span><span class="sxs-lookup"><span data-stu-id="c4468-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="c4468-105">どちらも JavaScript API を使用してブックExcel制御を提供します。</span><span class="sxs-lookup"><span data-stu-id="c4468-105">They both offer automated control of an Excel workbook a JavaScript API.</span></span> <span data-ttu-id="c4468-106">ただし、Officeスクリプト API は、JavaScript API の特殊な同期Officeです。</span><span class="sxs-lookup"><span data-stu-id="c4468-106">However, the Office Scripts APIs are a specialized, synchronous version of the Office JavaScript API.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="さまざまな拡張ソリューションのフォーカス領域を示す 4 象限Office図。Office スクリプトと Office Web アドインの両方が Web とコラボレーションに焦点を当て、Office スクリプトはエンド ユーザーに対応します (一方、Office Web アドインはプロの開発者を対象とします)":::

<span data-ttu-id="c4468-108">Officeスクリプトは、手動ボタンを押して完了するか[、Power Automate](https://flow.microsoft.com/)でステップとして実行しますが、Office アドインは作業ウィンドウを開いている間も保持されます。</span><span class="sxs-lookup"><span data-stu-id="c4468-108">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="c4468-109">つまり、アドインはセッション中に状態を維持できるのに対し、Officeスクリプトは実行の間に内部状態を維持できません。</span><span class="sxs-lookup"><span data-stu-id="c4468-109">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="c4468-110">Excel 拡張機能がスクリプト プラットフォームの機能を超える必要がある場合は、Office アドインのドキュメントを参照して[、Office](/office/dev/add-ins)アドインの詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="c4468-110">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="c4468-111">この記事の残りの部分では、アドインとスクリプトの主なOfficeについてOfficeします。</span><span class="sxs-lookup"><span data-stu-id="c4468-111">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="c4468-112">プラットフォームサポート</span><span class="sxs-lookup"><span data-stu-id="c4468-112">Platform Support</span></span>

<span data-ttu-id="c4468-113">Officeアドインはクロスプラットフォームです。</span><span class="sxs-lookup"><span data-stu-id="c4468-113">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="c4468-114">デスクトップ、Mac、Windows Web プラットフォーム間で動作し、それぞれで同じエクスペリエンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c4468-114">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="c4468-115">この例外は、個々の API のドキュメントに示されています。</span><span class="sxs-lookup"><span data-stu-id="c4468-115">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="c4468-116">Officeスクリプトは現在、ユーザーがサポートしているExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="c4468-116">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="c4468-117">すべての記録、編集、および実行は、Web プラットフォーム上で行われます。</span><span class="sxs-lookup"><span data-stu-id="c4468-117">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="c4468-118">API</span><span class="sxs-lookup"><span data-stu-id="c4468-118">APIs</span></span>

<span data-ttu-id="c4468-119">アドイン用の JavaScript API Office同期バージョンOfficeはありません。標準のOfficeスクリプト API はプラットフォームに固有であり、パラダイムの使用を避けるための多数の最適化と変更 `load` / `sync` があります。</span><span class="sxs-lookup"><span data-stu-id="c4468-119">There is no synchronous version of the Office JavaScript APIs for Office Add-ins. The standard Office Scripts APIs are unique to the platform and have numerous optimizations and alterations to avoid the usage of the `load`/`sync` paradigm.</span></span>

<span data-ttu-id="c4468-120">JavaScript API の[一部Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)スクリプト非同期 API[とOffice互換性があります](../develop/excel-async-model.md)。</span><span class="sxs-lookup"><span data-stu-id="c4468-120">Some of the [Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) are compatible with the [Office Scripts Async APIs](../develop/excel-async-model.md).</span></span> <span data-ttu-id="c4468-121">一部のサンプルとアドイン コード ブロックは、最小限の変換でブロック `Excel.run` に移植できます。</span><span class="sxs-lookup"><span data-stu-id="c4468-121">Some samples and add-in code blocks could be ported to `Excel.run` blocks with minimal translation.</span></span> <span data-ttu-id="c4468-122">2 つのプラットフォームは機能を共有しますが、ギャップがあります。</span><span class="sxs-lookup"><span data-stu-id="c4468-122">While the two platforms share functionality, there are gaps.</span></span> <span data-ttu-id="c4468-123">2 つの主要な API セットは、Officeに含まれますが、スクリプトOfficeイベントと共通 API ではありません。</span><span class="sxs-lookup"><span data-stu-id="c4468-123">The two major API sets that Office Add-ins have but Office Scripts do not are events and the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="c4468-124">イベント</span><span class="sxs-lookup"><span data-stu-id="c4468-124">Events</span></span>

<span data-ttu-id="c4468-125">Officeスクリプトはイベントをサポート[していない](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="c4468-125">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="c4468-126">すべてのスクリプトでコードが 1 つのメソッドで `main` 実行され、終了します。</span><span class="sxs-lookup"><span data-stu-id="c4468-126">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="c4468-127">イベントがトリガーされると再アクティブ化されないので、イベントを登録できません。</span><span class="sxs-lookup"><span data-stu-id="c4468-127">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="c4468-128">共通 API</span><span class="sxs-lookup"><span data-stu-id="c4468-128">Common APIs</span></span>

<span data-ttu-id="c4468-129">Officeスクリプトで共通[API を使用することはできません](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="c4468-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="c4468-130">一般的な API でのみサポートされている認証、ダイアログ ウィンドウ、その他の機能が必要な場合は、Office スクリプトではなく Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c4468-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="c4468-131">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="c4468-131">See also</span></span>

- [<span data-ttu-id="c4468-132">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="c4468-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="c4468-133">スクリプトと VBA Officeの違い</span><span class="sxs-lookup"><span data-stu-id="c4468-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="c4468-134">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="c4468-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="c4468-135">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="c4468-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
