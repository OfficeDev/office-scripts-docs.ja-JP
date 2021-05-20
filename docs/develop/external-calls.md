---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: Office スクリプトで外部 API 呼び出しを行うためのサポートとガイダンス。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545083"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="afb01-103">Office スクリプトでの外部 API 呼び出しのサポート</span><span class="sxs-lookup"><span data-stu-id="afb01-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="afb01-104">スクリプト作成者は、プラットフォームのプレビュー段階で [外部 API を](https://developer.mozilla.org/docs/Web/API) 使用する場合に一貫した動作を期待すべきではありません。</span><span class="sxs-lookup"><span data-stu-id="afb01-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="afb01-105">そのため、重要なスクリプト シナリオでは外部 API に依存しないでください。</span><span class="sxs-lookup"><span data-stu-id="afb01-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="afb01-106">外部 API への呼び出しは、[通常の状況では](#external-calls-from-power-automate)Power Automateではなく、Excel アプリケーションを通じてのみ行うことができます。</span><span class="sxs-lookup"><span data-stu-id="afb01-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="afb01-107">外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="afb01-108">管理者は、このような呼び出しに対してファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="afb01-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="afb01-109">外部呼び出し用のスクリプトの構成</span><span class="sxs-lookup"><span data-stu-id="afb01-109">Configure your script for external calls</span></span>

<span data-ttu-id="afb01-110">外部呼び出しは [非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトが `async` .</span><span class="sxs-lookup"><span data-stu-id="afb01-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="afb01-111">次に `async` `main` 示すように、関数にプレフィックスを追加し、 `Promise` を返すようにします。</span><span class="sxs-lookup"><span data-stu-id="afb01-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="afb01-112">他の情報を返すスクリプト `Promise` は、その型の を返すことができます。</span><span class="sxs-lookup"><span data-stu-id="afb01-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="afb01-113">たとえば、スクリプトがオブジェクトを返す必要がある場合 `Employee` 、返されるシグネチャは次のようになります。 `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="afb01-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="afb01-114">そのサービスを呼び出すためには、外部サービスのインターフェイスを学習する必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="afb01-115">または REST API を使用している場合は `fetch` 、返されるデータの JSON 構造を決定する必要があります。 [](https://wikipedia.org/wiki/Representational_state_transfer)</span><span class="sxs-lookup"><span data-stu-id="afb01-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="afb01-116">スクリプトへの入力とスクリプトからの出力の両方について、必要な `interface` JSON 構造に一致するように を作成することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="afb01-117">これにより、スクリプトのタイプ セーフが強化されます。</span><span class="sxs-lookup"><span data-stu-id="afb01-117">This gives the script more type safety.</span></span> <span data-ttu-id="afb01-118">この例については、「 [Office スクリプトからのフェッチを使用する 」を参照してください](../resources/samples/external-fetch-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="afb01-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="afb01-119">Office スクリプトからの外部呼び出しに関する制限事項</span><span class="sxs-lookup"><span data-stu-id="afb01-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="afb01-120">サインインしたり、OAuth2 タイプの認証フローを使用する方法はありません。</span><span class="sxs-lookup"><span data-stu-id="afb01-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="afb01-121">すべてのキーと資格情報は、ハードコード (または別のソースから読み取る) する必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="afb01-122">API の資格情報とキーを格納するインフラストラクチャはありません。</span><span class="sxs-lookup"><span data-stu-id="afb01-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="afb01-123">これはユーザーが管理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="afb01-124">ドキュメントの Cookie、 `localStorage` および `sessionStorage` オブジェクトはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="afb01-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="afb01-125">外部呼び出しによって、機密データが望ましくないエンドポイントに公開されたり、外部データが内部ワークブックに取り込まれる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="afb01-126">管理者は、このような呼び出しに対してファイアウォール保護を確立できます。</span><span class="sxs-lookup"><span data-stu-id="afb01-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="afb01-127">外部呼び出しに依存する前に、必ずローカル ポリシーを確認してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="afb01-128">依存関係を取得する前に、データ スループットの量を確認してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="afb01-129">たとえば、外部データセット全体をプルダウンするのが最適な方法ではない場合があり、代わりにページ分割を使用してデータをチャンク単位で取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="afb01-130">で情報を取得 `fetch`</span><span class="sxs-lookup"><span data-stu-id="afb01-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="afb01-131">[フェッチ API は](https://developer.mozilla.org/docs/Web/API/Fetch_API)、外部サービスから情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="afb01-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="afb01-132">これは `async` API なので、スクリプトの署名を調整する必要 `main` があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="afb01-133">関数を `main` 作成 `async` し、 を返します `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="afb01-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="afb01-134">また、 `await` `fetch` 呼び出しと取得を確認する必要があります `json` 。</span><span class="sxs-lookup"><span data-stu-id="afb01-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="afb01-135">これにより、スクリプトが終了する前にこれらの操作が完了します。</span><span class="sxs-lookup"><span data-stu-id="afb01-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="afb01-136">によって取得される JSON データは、 `fetch` スクリプトで定義されているインターフェイスと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="afb01-137">返される値は[、Office スクリプトは `any` 型をサポートしていないため、特定の型に](typescript-restrictions.md#no-any-type-in-office-scripts)割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="afb01-138">返されるプロパティの名前と型を確認するには、サービスのドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="afb01-139">次に、一致するインターフェイスをスクリプトに追加します。</span><span class="sxs-lookup"><span data-stu-id="afb01-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="afb01-140">次のスクリプトは `fetch` 、指定された URL のテスト サーバーから JSON データを取得するために使用します。</span><span class="sxs-lookup"><span data-stu-id="afb01-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="afb01-141">`JSONData`データを一致する型として格納するインターフェイスに注意してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

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

### <a name="other-fetch-samples"></a><span data-ttu-id="afb01-142">その他の `fetch` サンプル</span><span class="sxs-lookup"><span data-stu-id="afb01-142">Other `fetch` samples</span></span>

* <span data-ttu-id="afb01-143">[Officeスクリプトで外部フェッチ呼び出しを使用するサンプルでは](../resources/samples/external-fetch-calls.md)、ユーザーのGitHubリポジトリに関する基本情報を取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="afb01-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="afb01-144">[Officeスクリプトのサンプルシナリオ:NOAAからの水位データGraph](../resources/scenarios/noaa-data-fetch.md)は、米国海洋大気局の潮汐と電流データベースからレコードを取得するために使用されているフェッチコマンドを示しています。</span><span class="sxs-lookup"><span data-stu-id="afb01-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="afb01-145">Power Automateからの外部通話</span><span class="sxs-lookup"><span data-stu-id="afb01-145">External calls from Power Automate</span></span>

<span data-ttu-id="afb01-146">Power Automateを指定してスクリプトを実行すると、外部 API 呼び出しが失敗します。</span><span class="sxs-lookup"><span data-stu-id="afb01-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="afb01-147">これは、Excelアプリケーションを通じてスクリプトを実行する場合とPower Automateを使用する場合の動作の違いです。</span><span class="sxs-lookup"><span data-stu-id="afb01-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="afb01-148">フローに組み込む前に、スクリプトでそのような参照を確認してください。</span><span class="sxs-lookup"><span data-stu-id="afb01-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="afb01-149">[Azure AD](/connectors/webcontents/)と共に HTTP を使用するか、他の同等のアクションを使用して、データを外部サービスから取得またはプッシュする必要があります。</span><span class="sxs-lookup"><span data-stu-id="afb01-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="afb01-150">[Power Automate Excelオンライン コネクタ](/connectors/excelonlinebusiness)を介して行われた外部呼び出しは、既存のデータ損失防止ポリシーを守るために失敗します。</span><span class="sxs-lookup"><span data-stu-id="afb01-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="afb01-151">ただし、Power Automateを介して実行されるスクリプトは、組織の外部および組織のファイアウォールの外部で実行されます。</span><span class="sxs-lookup"><span data-stu-id="afb01-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="afb01-152">この外部環境で悪意のあるユーザーから保護を強化するために、管理者はOfficeスクリプトの使用を制御できます。</span><span class="sxs-lookup"><span data-stu-id="afb01-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="afb01-153">管理者は、Power AutomateでExcelオンライン コネクタを無効にするか、Office スクリプト[管理者コントロール](/microsoft-365/admin/manage/manage-office-scripts-settings)を使用してExcel on the web用Officeスクリプトを無効にできます。</span><span class="sxs-lookup"><span data-stu-id="afb01-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="afb01-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="afb01-154">See also</span></span>

* [<span data-ttu-id="afb01-155">Office スクリプトでの組み込みの JavaScript オブジェクトの使用</span><span class="sxs-lookup"><span data-stu-id="afb01-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="afb01-156">Office スクリプトで外部取得呼び出しを使用する</span><span class="sxs-lookup"><span data-stu-id="afb01-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="afb01-157">Officeスクリプトのサンプル シナリオ: NOAA からの水位データのGraph</span><span class="sxs-lookup"><span data-stu-id="afb01-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
