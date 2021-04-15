---
title: スクリプトで外部 API 呼び出Officeする
description: スクリプトで外部 API 呼び出しを行うOfficeします。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754804"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="48124-103">スクリプトからの外部 API Office呼び出し</span><span class="sxs-lookup"><span data-stu-id="48124-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="48124-104">Officeスクリプトを使用すると、 [外部 API 呼び出しのサポートが制限されます](../../develop/external-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="48124-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="48124-105">OAuth2 タイプの認証フローをサインインまたは使用する方法はありません。</span><span class="sxs-lookup"><span data-stu-id="48124-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="48124-106">すべてのキーと資格情報をハードコード (または別のソースから読み取る) 必要があります。</span><span class="sxs-lookup"><span data-stu-id="48124-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="48124-107">API の資格情報とキーを格納するインフラストラクチャはありません。</span><span class="sxs-lookup"><span data-stu-id="48124-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="48124-108">これは、ユーザーが管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48124-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="48124-109">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される場合や、内部ブックに外部データが取り込まれたりする場合があります。</span><span class="sxs-lookup"><span data-stu-id="48124-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="48124-110">管理者は、このような呼び出しに対するファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="48124-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="48124-111">外部通話に依存する前に、必ずローカル ポリシーに確認してください。</span><span class="sxs-lookup"><span data-stu-id="48124-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="48124-112">スクリプトが API 呼び出しを使用する場合、Power Automate シナリオでは機能しません。</span><span class="sxs-lookup"><span data-stu-id="48124-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="48124-113">Power Automate の HTTP アクションまたは同等のアクションを使用して、データを外部サービスから取得または外部サービスにプッシュする必要があります。</span><span class="sxs-lookup"><span data-stu-id="48124-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="48124-114">外部 API 呼び出しには非同期 API 構文が含まれるので、非同期通信の仕組みについて少し高度な知識が必要です。</span><span class="sxs-lookup"><span data-stu-id="48124-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="48124-115">依存関係を取得する前に、データ スループットの量を確認してください。</span><span class="sxs-lookup"><span data-stu-id="48124-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="48124-116">たとえば、外部データセット全体を引き下げないのが最適な選択肢ではなく、代わりにページネーションを使用してデータをチャンク単位で取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="48124-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="48124-117">有用な知識とリソース</span><span class="sxs-lookup"><span data-stu-id="48124-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="48124-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): ほとんどの場合、API 呼び出しを使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="48124-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="48124-119">[ `async` : この動作を理解します `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)。</span><span class="sxs-lookup"><span data-stu-id="48124-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="48124-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): この動作を理解します。</span><span class="sxs-lookup"><span data-stu-id="48124-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="48124-121">手順</span><span class="sxs-lookup"><span data-stu-id="48124-121">Steps</span></span>

1. <span data-ttu-id="48124-122">プレフィックスを `main` 追加して、関数を非同期関数としてマーク `async` します。</span><span class="sxs-lookup"><span data-stu-id="48124-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="48124-123">たとえば、`async function main(workbook: ExcelScript.Workbook)` などです。</span><span class="sxs-lookup"><span data-stu-id="48124-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="48124-124">どの種類の API 呼び出しを行っていますか?</span><span class="sxs-lookup"><span data-stu-id="48124-124">Which type of API call are you making?</span></span> <span data-ttu-id="48124-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="48124-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="48124-126">詳細については、REST API の資料を参照してください。</span><span class="sxs-lookup"><span data-stu-id="48124-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="48124-127">サービス API エンドポイント、認証要件、ヘッダーなどを取得します。</span><span class="sxs-lookup"><span data-stu-id="48124-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="48124-128">コードの完了と開発時間の検証に役立つ入力または `interface` 出力を定義します。</span><span class="sxs-lookup"><span data-stu-id="48124-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="48124-129">詳細については [、ビデオ](#training-video-how-to-make-external-api-calls) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="48124-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="48124-130">コード、テスト、最適化。</span><span class="sxs-lookup"><span data-stu-id="48124-130">Code, test, optimize.</span></span> <span data-ttu-id="48124-131">API 呼び出しルーチンの関数を作成して、スクリプトの他の部分から再利用可能にしたり、別のスクリプトで再利用したりすることができます (コピー貼り付けは、この方法ではるかに簡単になります)。</span><span class="sxs-lookup"><span data-stu-id="48124-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="48124-132">シナリオ</span><span class="sxs-lookup"><span data-stu-id="48124-132">Scenario</span></span>

<span data-ttu-id="48124-133">このスクリプトは、ユーザーの GitHub リポジトリに関する基本情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="48124-133">This script gets basic information about the user's GitHub repositories.</span></span>

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="48124-134">サンプルで使用されるリソース</span><span class="sxs-lookup"><span data-stu-id="48124-134">Resources used in the sample</span></span>

1. [<span data-ttu-id="48124-135">リポジトリ Github API リファレンスを取得します。</span><span class="sxs-lookup"><span data-stu-id="48124-135">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="48124-136">API 呼び出しの出力: Web ブラウザーまたは任意の HTTP インターフェイスに移動して入力し `https://api.github.com/users/{USERNAME}/repos` 、{USERNAME} プレースホルダーを Github ID に置き換える。</span><span class="sxs-lookup"><span data-stu-id="48124-136">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="48124-137">取得される情報: repo.name、repo.size、repo.owner.id、repo.license?。name</span><span class="sxs-lookup"><span data-stu-id="48124-137">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="48124-138">サンプル コード: ユーザーの GitHub リポジトリに関する基本情報を取得する</span><span class="sxs-lookup"><span data-stu-id="48124-138">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="48124-139">トレーニング ビデオ: 外部 API 呼び出しを行う方法</span><span class="sxs-lookup"><span data-stu-id="48124-139">Training video: How to make external API calls</span></span>

<span data-ttu-id="48124-140">[![外部 API 呼び出しの実行方法に関するビデオを見る](../../images/api-vid.png)](https://youtu.be/fulP29J418E "外部 API 呼び出しを行う方法に関するビデオ")</span><span class="sxs-lookup"><span data-stu-id="48124-140">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
