---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: スクリプトで外部 API 呼び出しを行うOffice。
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570312"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="15e60-103">Office スクリプトでの外部 API 呼び出しのサポート</span><span class="sxs-lookup"><span data-stu-id="15e60-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="15e60-104">スクリプト作成者は、プラットフォームのプレビュー 段階で外部 [API](https://developer.mozilla.org/docs/Web/API) を使用する場合、一貫した動作を期待してはならない。</span><span class="sxs-lookup"><span data-stu-id="15e60-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="15e60-105">そのため、重要なスクリプト シナリオでは外部 API に依存しません。</span><span class="sxs-lookup"><span data-stu-id="15e60-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="15e60-106">外部 API への呼び出しは、通常の状況では Power Automate 経由ではなく、Excel アプリケーション経由 [でのみ実行できます](#external-calls-from-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="15e60-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="15e60-107">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="15e60-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="15e60-108">管理者は、このような呼び出しに対するファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="15e60-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="15e60-109">操作 `fetch`</span><span class="sxs-lookup"><span data-stu-id="15e60-109">Working with `fetch`</span></span>

<span data-ttu-id="15e60-110">フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="15e60-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="15e60-111">これは `async` API なので、スクリプトの署名を `main` 調整する必要があります。</span><span class="sxs-lookup"><span data-stu-id="15e60-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="15e60-112">関数を `main` 作成 `async` し、 を返します `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="15e60-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="15e60-113">また、呼び出しと取得 `await` `fetch` も確認する必要 `json` があります。</span><span class="sxs-lookup"><span data-stu-id="15e60-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="15e60-114">これにより、スクリプトが終了する前にこれらの操作が確実に完了します。</span><span class="sxs-lookup"><span data-stu-id="15e60-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="15e60-115">次のスクリプトは、指定された URL のテスト サーバーから `fetch` JSON データを取得するために使用します。</span><span class="sxs-lookup"><span data-stu-id="15e60-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

<span data-ttu-id="15e60-116">Office スクリプトのサンプル シナリオ [: NOAA](../resources/scenarios/noaa-data-fetch.md) の水位データをグラフ化すると、国立海洋大気局の潮流データベースからレコードを取得するために使用されるフェッチ コマンドが示されています。</span><span class="sxs-lookup"><span data-stu-id="15e60-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="15e60-117">Power Automate からの外部通話</span><span class="sxs-lookup"><span data-stu-id="15e60-117">External calls from Power Automate</span></span>

<span data-ttu-id="15e60-118">Power Automate を使用してスクリプトを実行すると、外部 API 呼び出しは失敗します。</span><span class="sxs-lookup"><span data-stu-id="15e60-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="15e60-119">これは、Excel クライアントを使用してスクリプトを実行する場合と Power Automate を使用する場合の動作の違いです。</span><span class="sxs-lookup"><span data-stu-id="15e60-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="15e60-120">フローに組み込む前に、スクリプトでそのような参照を確認してください。</span><span class="sxs-lookup"><span data-stu-id="15e60-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="15e60-121">Power [Automate Excel Online](/connectors/excelonlinebusiness) コネクタを介して行われた外部呼び出しは、既存のデータ損失防止ポリシーを支持するために失敗します。</span><span class="sxs-lookup"><span data-stu-id="15e60-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="15e60-122">ただし、Power Automate を介して実行されるスクリプトは、組織外および組織のファイアウォールの外部で実行されます。</span><span class="sxs-lookup"><span data-stu-id="15e60-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="15e60-123">この外部環境で悪意のあるユーザーから保護するために、管理者はスクリプトの使用Officeできます。</span><span class="sxs-lookup"><span data-stu-id="15e60-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="15e60-124">管理者は、Power Automate で Excel Online コネクタを無効にするか、Office スクリプト管理者コントロールを使用して Web 上の Excel Office [を無効にできます](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="15e60-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="15e60-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="15e60-125">See also</span></span>

- [<span data-ttu-id="15e60-126">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="15e60-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="15e60-127">Office スクリプトのサンプル シナリオ: NOAA からの水位データのグラフ</span><span class="sxs-lookup"><span data-stu-id="15e60-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
