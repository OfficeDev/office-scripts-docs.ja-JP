---
title: Office スクリプトと VBA マクロの相違点
description: Office スクリプトと Excel VBA マクロの動作と API の違い。
ms.date: 06/30/2020
localization_priority: Normal
ms.openlocfilehash: 8c246545943341607a7aced4da792b8e49880cb0
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616690"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a><span data-ttu-id="687b4-103">Office スクリプトと VBA マクロの相違点</span><span class="sxs-lookup"><span data-stu-id="687b4-103">Differences between Office Scripts and VBA macros</span></span>

<span data-ttu-id="687b4-104">Office スクリプトと VBA マクロには、多くの共通点があります。</span><span class="sxs-lookup"><span data-stu-id="687b4-104">Office Scripts and VBA macros have a lot in common.</span></span> <span data-ttu-id="687b4-105">両方のユーザーが、使いやすいアクションレコーダーを使用してソリューションを自動化し、それらのレコーディングを編集できるようにします。</span><span class="sxs-lookup"><span data-stu-id="687b4-105">They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings.</span></span> <span data-ttu-id="687b4-106">両方のフレームワークは、プログラマが開発者にとって Excel で小さなプログラムを作成することを考慮していない可能性があるユーザーを支援するためのものです。</span><span class="sxs-lookup"><span data-stu-id="687b4-106">Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.</span></span>
<span data-ttu-id="687b4-107">基本的な違いは、デスクトップソリューション用に VBA マクロが開発されており、Office スクリプトが、ガイド原則としてクロスプラットフォームのサポートとセキュリティを使用して設計されているということです。</span><span class="sxs-lookup"><span data-stu-id="687b4-107">The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles.</span></span> <span data-ttu-id="687b4-108">現時点では、Office スクリプトは web 上の Excel でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="687b4-108">Currently, Office Scripts are only supported in Excel on the web.</span></span>

![さまざまな Office 機能拡張ソリューションに対するフォーカスの領域を示す4つの領域の図。](../images/office-programmability-diagram.png)

<span data-ttu-id="687b4-111">この記事では、VBA マクロ (および VBA) と Office スクリプトの主な違いについて説明します。</span><span class="sxs-lookup"><span data-stu-id="687b4-111">This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts.</span></span> <span data-ttu-id="687b4-112">Office スクリプトは Excel でのみ使用可能であるため、ここで説明する唯一のホストです。</span><span class="sxs-lookup"><span data-stu-id="687b4-112">Since Office Scripts are only available for Excel, that is the only host being discussed here.</span></span>

## <a name="platform-and-ecosystem"></a><span data-ttu-id="687b4-113">プラットフォームとエコシステム</span><span class="sxs-lookup"><span data-stu-id="687b4-113">Platform and ecosystem</span></span>

<span data-ttu-id="687b4-114">VBA はデスクトップ向けに設計されており、Office スクリプトは web 用に設計されています。</span><span class="sxs-lookup"><span data-stu-id="687b4-114">VBA is designed for the desktop and Office Scripts are designed for the web.</span></span> <span data-ttu-id="687b4-115">VBA は、ユーザーのデスクトップと対話して、COM や OLE などの同様のテクノロジに接続できます。</span><span class="sxs-lookup"><span data-stu-id="687b4-115">VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE.</span></span> <span data-ttu-id="687b4-116">ただし、VBA には、インターネットを呼び出すための便利な方法はありません。</span><span class="sxs-lookup"><span data-stu-id="687b4-116">However, VBA has no convenient way to call out to the internet.</span></span>

<span data-ttu-id="687b4-117">Office スクリプトでは、JavaScript の汎用ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="687b4-117">Office Scripts use a universal runtime for JavaScript.</span></span> <span data-ttu-id="687b4-118">これにより、スクリプトの実行に使用されているコンピューターに関係なく、一貫性のある動作とアクセスが可能になります。</span><span class="sxs-lookup"><span data-stu-id="687b4-118">This gives consistent behavior and accessibility, regardless of the machine being used to run the script.</span></span> <span data-ttu-id="687b4-119">また、他の web サービスへの呼び出しを行うこともできます。</span><span class="sxs-lookup"><span data-stu-id="687b4-119">They can also make calls to other web services.</span></span>

## <a name="security"></a><span data-ttu-id="687b4-120">セキュリティ</span><span class="sxs-lookup"><span data-stu-id="687b4-120">Security</span></span>

<span data-ttu-id="687b4-121">VBA マクロには、Excel と同じセキュリティクリアランスがあります。</span><span class="sxs-lookup"><span data-stu-id="687b4-121">VBA macros have the same security clearance as Excel.</span></span> <span data-ttu-id="687b4-122">これにより、デスクトップへのフルアクセスが可能になります。</span><span class="sxs-lookup"><span data-stu-id="687b4-122">This gives them full access to your desktop.</span></span> <span data-ttu-id="687b4-123">Office スクリプトは、ブックをホストするマシンではなく、ブックへのアクセスのみが可能です。</span><span class="sxs-lookup"><span data-stu-id="687b4-123">Office Scripts only have access to the workbook, not the machine hosting the workbook.</span></span> <span data-ttu-id="687b4-124">さらに、スクリプトでは、JavaScript 認証トークンを共有できないため、スクリプトは外部サービスで認証されません。</span><span class="sxs-lookup"><span data-stu-id="687b4-124">Additionally, no JavaScript authentication tokens can be shared with scripts, so scripts can never authenticate with an external service.</span></span>

<span data-ttu-id="687b4-125">管理者には、VBA マクロに関する3つのオプションがあります。テナントのすべてのマクロを許可するか、テナントにマクロを許可しないか、署名された証明書によるマクロのみを許可します。</span><span class="sxs-lookup"><span data-stu-id="687b4-125">Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates.</span></span> <span data-ttu-id="687b4-126">このように細分化されていないと、1つの不良アクターを分離するのが困難になります。</span><span class="sxs-lookup"><span data-stu-id="687b4-126">This lack of granularity makes it hard to isolate a single bad actor.</span></span> <span data-ttu-id="687b4-127">現時点では、Office スクリプトはテナントに対してオンまたはオフになっています。</span><span class="sxs-lookup"><span data-stu-id="687b4-127">Currently, Office Scripts are either on or off for a tenant.</span></span> <span data-ttu-id="687b4-128">しかし、管理者が個々のスクリプトやスクリプト作成者をより詳細に制御できるようにしています。</span><span class="sxs-lookup"><span data-stu-id="687b4-128">However, we are working to give admins more control over individual scripts and script creators.</span></span>

## <a name="coverage"></a><span data-ttu-id="687b4-129">割合</span><span class="sxs-lookup"><span data-stu-id="687b4-129">Coverage</span></span>

<span data-ttu-id="687b4-130">現時点では、VBA は、デスクトップクライアントで使用可能な Excel 機能の詳細な範囲を提供しています。</span><span class="sxs-lookup"><span data-stu-id="687b4-130">Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client.</span></span> <span data-ttu-id="687b4-131">Office スクリプトでは、web 上の Excel のほぼすべてのシナリオについて説明します。</span><span class="sxs-lookup"><span data-stu-id="687b4-131">Office Scripts cover nearly all of the scenarios for Excel on the web.</span></span> <span data-ttu-id="687b4-132">また、web 上の新機能の debut により、Office スクリプトはアクションレコーダーと JavaScript Api の両方に対してそれらをサポートします。</span><span class="sxs-lookup"><span data-stu-id="687b4-132">Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.</span></span>

## <a name="power-automate"></a><span data-ttu-id="687b4-133">Power Automate</span><span class="sxs-lookup"><span data-stu-id="687b4-133">Power Automate</span></span>

<span data-ttu-id="687b4-134">Office スクリプトは、電源自動化を使用して実行できます。</span><span class="sxs-lookup"><span data-stu-id="687b4-134">Office Scripts can be run through Power Automate.</span></span> <span data-ttu-id="687b4-135">ブックは、スケジュールまたはイベントドリブンフローによって更新できます。これにより、Excel を開かなくてもワークフローを自動化できます。</span><span class="sxs-lookup"><span data-stu-id="687b4-135">Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel.</span></span> <span data-ttu-id="687b4-136">これは、ブックが OneDrive に保存されていて、そのユーザーが Excel のデスクトップ、Mac、または web クライアントを使用しているかどうかに関係なく、フローがスクリプトを実行できることを意味します。</span><span class="sxs-lookup"><span data-stu-id="687b4-136">This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.</span></span>

<span data-ttu-id="687b4-137">VBA には Power オートメーションコネクタがありません。</span><span class="sxs-lookup"><span data-stu-id="687b4-137">VBA doesn't have a Power Automate connector.</span></span> <span data-ttu-id="687b4-138">サポートされているすべての VBA シナリオでは、マクロの実行にユーザーが参加していました。</span><span class="sxs-lookup"><span data-stu-id="687b4-138">All supported VBA scenarios involved a user attending to the macro's execution.</span></span>

## <a name="see-also"></a><span data-ttu-id="687b4-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="687b4-139">See also</span></span>

- [<span data-ttu-id="687b4-140">Excel on the web の Office スクリプト</span><span class="sxs-lookup"><span data-stu-id="687b4-140">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="687b4-141">Office スクリプトと Office アドインの違い</span><span class="sxs-lookup"><span data-stu-id="687b4-141">Differences between Office Scripts and Office Add-ins</span></span>](add-ins-differences.md)
- [<span data-ttu-id="687b4-142">Office スクリプトのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="687b4-142">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="687b4-143">Excel VBA リファレンス</span><span class="sxs-lookup"><span data-stu-id="687b4-143">Excel VBA reference</span></span>](/office/vba/api/overview/excel)
