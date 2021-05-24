---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: スクリプトで外部 API 呼び出しを行うOffice。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545083"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="cf11b-103">Office スクリプトでの外部 API 呼び出しのサポート</span><span class="sxs-lookup"><span data-stu-id="cf11b-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="cf11b-104">スクリプト作成者は、プラットフォームのプレビュー 段階で外部 [API](https://developer.mozilla.org/docs/Web/API) を使用する場合、一貫した動作を期待してはならない。</span><span class="sxs-lookup"><span data-stu-id="cf11b-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="cf11b-105">そのため、重要なスクリプト シナリオでは外部 API に依存しません。</span><span class="sxs-lookup"><span data-stu-id="cf11b-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="cf11b-106">外部 API への呼び出しは、通常の状況Excelアプリケーションを介Power Automate[実行できます](#external-calls-from-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="cf11b-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="cf11b-107">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="cf11b-108">管理者は、このような呼び出しに対するファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="cf11b-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="cf11b-109">外部呼び出し用にスクリプトを構成する</span><span class="sxs-lookup"><span data-stu-id="cf11b-109">Configure your script for external calls</span></span>

<span data-ttu-id="cf11b-110">外部呼び出 [しは非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトがとしてマークされている必要があります `async` 。</span><span class="sxs-lookup"><span data-stu-id="cf11b-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="cf11b-111">次に示すように、プレフィックスを関数に追加 `async` `main` し、それを `Promise` 返すようにします。</span><span class="sxs-lookup"><span data-stu-id="cf11b-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="cf11b-112">他の情報を返すスクリプトは、その種類の `Promise` 1 つを返す可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="cf11b-113">たとえば、スクリプトでオブジェクトを返す必要がある場合、 `Employee` 戻り値の署名は次のようになります。 `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="cf11b-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="cf11b-114">そのサービスを呼び出すには、外部サービスのインターフェイスを学習する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="cf11b-115">REST API を使用 `fetch` [している場合](https://wikipedia.org/wiki/Representational_state_transfer)は、返されるデータの JSON 構造を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="cf11b-116">スクリプトの入力と出力の両方について、必要な JSON 構造に一致 `interface` するを検討してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="cf11b-117">これにより、スクリプトの型の安全性が向上します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-117">This gives the script more type safety.</span></span> <span data-ttu-id="cf11b-118">この例については、「スクリプトからフェッチを使用する[」でOfficeできます](../resources/samples/external-fetch-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="cf11b-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="cf11b-119">スクリプトからの外部呼び出しOffice制限</span><span class="sxs-lookup"><span data-stu-id="cf11b-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="cf11b-120">OAuth2 タイプの認証フローをサインインまたは使用する方法はありません。</span><span class="sxs-lookup"><span data-stu-id="cf11b-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="cf11b-121">すべてのキーと資格情報をハードコード (または別のソースから読み取る) 必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="cf11b-122">API の資格情報とキーを格納するインフラストラクチャはありません。</span><span class="sxs-lookup"><span data-stu-id="cf11b-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="cf11b-123">これは、ユーザーが管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="cf11b-124">ドキュメント Cookie、 `localStorage` および `sessionStorage` オブジェクトはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf11b-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="cf11b-125">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される場合や、内部ブックに外部データが取り込まれたりする場合があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="cf11b-126">管理者は、このような呼び出しに対するファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="cf11b-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="cf11b-127">外部通話に依存する前に、必ずローカル ポリシーに確認してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="cf11b-128">依存関係を取得する前に、データ スループットの量を確認してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="cf11b-129">たとえば、外部データセット全体を引き下げないのが最適な選択肢ではなく、代わりにページネーションを使用してデータをチャンク単位で取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="cf11b-130">を使用して情報を取得する `fetch`</span><span class="sxs-lookup"><span data-stu-id="cf11b-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="cf11b-131">フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="cf11b-132">これは `async` API なので、スクリプトの署名を `main` 調整する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="cf11b-133">関数を `main` 作成 `async` し、 を返します `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="cf11b-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="cf11b-134">また、呼び出しと取得 `await` `fetch` も確認する必要 `json` があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="cf11b-135">これにより、スクリプトが終了する前にこれらの操作が確実に完了します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="cf11b-136">取得した JSON データは、 `fetch` スクリプトで定義されているインターフェイスと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="cf11b-137">スクリプトは型をサポートしていないのでOffice値を特定の型[に割り当てる必要 `any` があります](typescript-restrictions.md#no-any-type-in-office-scripts)。</span><span class="sxs-lookup"><span data-stu-id="cf11b-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="cf11b-138">返されるプロパティの名前と種類については、サービスのドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="cf11b-139">次に、一致するインターフェイスまたはインターフェイスをスクリプトに追加します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="cf11b-140">次のスクリプトは、指定された URL のテスト サーバーから `fetch` JSON データを取得するために使用します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="cf11b-141">データを `JSONData` 一致する型として格納するインターフェイスに注意してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a><span data-ttu-id="cf11b-142">その他 `fetch` のサンプル</span><span class="sxs-lookup"><span data-stu-id="cf11b-142">Other `fetch` samples</span></span>

* <span data-ttu-id="cf11b-143">[[スクリプトの外部フェッチ](../resources/samples/external-fetch-calls.md)呼び出しOffice使用する] サンプルは、ユーザーのリポジトリに関する基本情報を取得するGitHub示しています。</span><span class="sxs-lookup"><span data-stu-id="cf11b-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="cf11b-144">Office スクリプトのサンプル シナリオ[: NOAA](../resources/scenarios/noaa-data-fetch.md)の Graph 水レベルデータは、国立海洋大気局のタイドと Currents データベースからレコードを取得するために使用されるフェッチ コマンドを示しています。</span><span class="sxs-lookup"><span data-stu-id="cf11b-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="cf11b-145">外部からの外部通話Power Automate</span><span class="sxs-lookup"><span data-stu-id="cf11b-145">External calls from Power Automate</span></span>

<span data-ttu-id="cf11b-146">スクリプトを使用してスクリプトを実行すると、外部 API 呼び出しPower Automate。</span><span class="sxs-lookup"><span data-stu-id="cf11b-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="cf11b-147">これは、スクリプトをアプリケーションを介して実行する場合と、Excelスクリプトを実行Power Automate。</span><span class="sxs-lookup"><span data-stu-id="cf11b-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="cf11b-148">フローに組み込む前に、スクリプトでそのような参照を確認してください。</span><span class="sxs-lookup"><span data-stu-id="cf11b-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="cf11b-149">データを外部サービスから取得または外部サービスにプッシュするには [、Azure AD](/connectors/webcontents/) または他の同等のアクションで HTTP を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf11b-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="cf11b-150">既存のデータ損失防止ポリシーを[Power Automate Excel、オンライン](/connectors/excelonlinebusiness)コネクタを介して行われた外部通話は失敗します。</span><span class="sxs-lookup"><span data-stu-id="cf11b-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="cf11b-151">ただし、組織の外部Power Automate、組織のファイアウォールの外部で実行されるスクリプトは実行されます。</span><span class="sxs-lookup"><span data-stu-id="cf11b-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="cf11b-152">この外部環境で悪意のあるユーザーからの保護を強化するために、管理者はスクリプトの使用Officeできます。</span><span class="sxs-lookup"><span data-stu-id="cf11b-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="cf11b-153">管理者は、Excel で Excel Power Automate Online コネクタを無効にするか、Office スクリプト管理者Excel on the webを使用して Office スクリプトを[無効にできます](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="cf11b-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="cf11b-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="cf11b-154">See also</span></span>

* [<span data-ttu-id="cf11b-155">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="cf11b-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="cf11b-156">Office スクリプトで外部取得呼び出しを使用する</span><span class="sxs-lookup"><span data-stu-id="cf11b-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="cf11b-157">Officeスクリプトのサンプル シナリオ: noAA Graphデータを使用する</span><span class="sxs-lookup"><span data-stu-id="cf11b-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
