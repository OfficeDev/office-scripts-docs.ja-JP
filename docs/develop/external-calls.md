---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: Office スクリプトで外部 API 呼び出しを行うためのサポートとガイダンス。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878819"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="16f2b-103">Office スクリプトでの外部 API 呼び出しのサポート</span><span class="sxs-lookup"><span data-stu-id="16f2b-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="16f2b-104">Office スクリプトプラットフォームは、[外部 api](https://developer.mozilla.org/docs/Web/API)への呼び出しをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="16f2b-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="16f2b-105">ただし、これらの呼び出しは適切な状況で実行することができます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="16f2b-106">外部呼び出しは、Excel クライアントを使用してのみ行うことができます。[通常の状況で](#external-calls-from-power-automate)は、電力の自動処理は行われません。</span><span class="sxs-lookup"><span data-stu-id="16f2b-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="16f2b-107">スクリプト作成者は、プラットフォームのプレビューフェーズ中に外部 Api を使用するときに、一貫した動作を期待する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="16f2b-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="16f2b-108">これは、JavaScript ランタイムがブックとの対話を管理する方法に起因します。</span><span class="sxs-lookup"><span data-stu-id="16f2b-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="16f2b-109">このスクリプトは、API 呼び出しが完了する前に終了することができます (または `Promise` 完全に解決される)。</span><span class="sxs-lookup"><span data-stu-id="16f2b-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="16f2b-110">そのため、重要なスクリプトシナリオでは外部 Api に依存しません。</span><span class="sxs-lookup"><span data-stu-id="16f2b-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="16f2b-111">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="16f2b-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="16f2b-112">管理者は、このような呼び出しに対してファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="16f2b-113">外部 Api の定義ファイル</span><span class="sxs-lookup"><span data-stu-id="16f2b-113">Definition files for external APIs</span></span>

<span data-ttu-id="16f2b-114">Office スクリプトには、外部 Api の定義ファイルは含まれていません。</span><span class="sxs-lookup"><span data-stu-id="16f2b-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="16f2b-115">このような Api を使用すると、定義が欠落しているとコンパイル時エラーが生成されます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="16f2b-116">次のスクリプトに示すように、Api は引き続き実行されます (ただし、Excel クライアントで実行する場合のみ)。</span><span class="sxs-lookup"><span data-stu-id="16f2b-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="16f2b-117">電源自動化からの外部通話</span><span class="sxs-lookup"><span data-stu-id="16f2b-117">External calls from Power Automate</span></span>

<span data-ttu-id="16f2b-118">電源自動化を使用してスクリプトを実行すると、外部 API 呼び出しは失敗します。</span><span class="sxs-lookup"><span data-stu-id="16f2b-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="16f2b-119">これは、Excel クライアントを使用してスクリプトを実行する場合と Power オートメーションを使用する場合の動作の違いです。</span><span class="sxs-lookup"><span data-stu-id="16f2b-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="16f2b-120">スクリプトをフローに組み込む前に、そのような参照について必ずチェックしてください。</span><span class="sxs-lookup"><span data-stu-id="16f2b-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="16f2b-121">Power [Online](/connectors/excelonlinebusiness)の外部通話の失敗は、既存のデータ損失防止ポリシーを守るために役立ちます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="16f2b-122">ただし、電源自動化によって実行されるスクリプトは、組織外、組織のファイアウォールの外側にあります。</span><span class="sxs-lookup"><span data-stu-id="16f2b-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="16f2b-123">この外部環境で悪意のあるユーザーからの保護を強化するために、管理者は Office スクリプトの使用を制御することができます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="16f2b-124">管理者は、Power turn で Excel Online コネクタを無効にするか、 [Office scripts administrator コントロール](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)を使用して web 上の Excel の office スクリプトをオフにすることができます。</span><span class="sxs-lookup"><span data-stu-id="16f2b-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="see-also"></a><span data-ttu-id="16f2b-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="16f2b-125">See also</span></span>

- [<span data-ttu-id="16f2b-126">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="16f2b-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)