---
title: Office スクリプトと Office アドインの相違点
description: Office スクリプトと Office アドインの動作と API の違い。
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700395"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="d653c-103">Office スクリプトと Office アドインの相違点</span><span class="sxs-lookup"><span data-stu-id="d653c-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="d653c-104">Office アドインと Office スクリプトには、多くの共通点があります。</span><span class="sxs-lookup"><span data-stu-id="d653c-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="d653c-105">どちらも、Office JavaScript API の名前空間を`Excel`使用して、Excel ブックの自動制御を提供します。</span><span class="sxs-lookup"><span data-stu-id="d653c-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="d653c-106">ただし、Office スクリプトの範囲は、より制限されています。</span><span class="sxs-lookup"><span data-stu-id="d653c-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="d653c-107">Office スクリプトは、手動のボタンを押すことで完了まで実行されます。 Office アドインは、ユーザーの操作に依存し、ブックの使用中は保持されます。</span><span class="sxs-lookup"><span data-stu-id="d653c-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="d653c-108">Excel 拡張機能がスクリプトプラットフォームの機能を超える必要がある場合は、office アドインの[ドキュメント](/office/dev/add-ins)にアクセスして、office アドインの詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="d653c-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="d653c-109">この記事の残りの部分では、Office アドインと Office スクリプトの主な違いについて説明します。</span><span class="sxs-lookup"><span data-stu-id="d653c-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="d653c-110">プラットフォームのサポート</span><span class="sxs-lookup"><span data-stu-id="d653c-110">Platform Support</span></span>

<span data-ttu-id="d653c-111">Office アドインはプラットフォーム間で機能します。</span><span class="sxs-lookup"><span data-stu-id="d653c-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="d653c-112">これらは、Windows デスクトップ、Mac、iOS、および web プラットフォーム間で機能し、それぞれに同じ操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="d653c-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="d653c-113">この点については、個々の API のドキュメントに記載されている例外を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d653c-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="d653c-114">Office スクリプトは、現在 web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d653c-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="d653c-115">すべての記録、編集、実行は、web プラットフォーム上で実行されます。</span><span class="sxs-lookup"><span data-stu-id="d653c-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="d653c-116">API</span><span class="sxs-lookup"><span data-stu-id="d653c-116">APIs</span></span>

<span data-ttu-id="d653c-117">Office スクリプトは、ほとんどの Excel JavaScript Api をサポートしています。これは、2つのプラットフォーム間で多くの機能が重なっていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="d653c-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="d653c-118">2つの例外として、イベントと共通 Api があります。</span><span class="sxs-lookup"><span data-stu-id="d653c-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="d653c-119">イベント</span><span class="sxs-lookup"><span data-stu-id="d653c-119">Events</span></span>

<span data-ttu-id="d653c-120">Office スクリプトは[イベント](/office/dev/add-ins/excel/excel-add-ins-events)をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="d653c-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="d653c-121">すべてのスクリプトは、コードを 1 `main`つのメソッドで実行し、終了します。</span><span class="sxs-lookup"><span data-stu-id="d653c-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="d653c-122">イベントがトリガーされると再アクティブ化されないため、イベントを登録できません。</span><span class="sxs-lookup"><span data-stu-id="d653c-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="d653c-123">共通 API</span><span class="sxs-lookup"><span data-stu-id="d653c-123">Common APIs</span></span>

<span data-ttu-id="d653c-124">Office スクリプトでは、[共通 api](/javascript/api/office)を使用できません。</span><span class="sxs-lookup"><span data-stu-id="d653c-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="d653c-125">一般的な Api でのみサポートされている認証、ダイアログウィンドウ、またはその他の機能が必要な場合は、Office のスクリプトではなく、Office アドインを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d653c-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="d653c-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="d653c-126">See also</span></span>

- [<span data-ttu-id="d653c-127">Web 上の Excel での Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="d653c-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="d653c-128">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="d653c-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="d653c-129">Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="d653c-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)