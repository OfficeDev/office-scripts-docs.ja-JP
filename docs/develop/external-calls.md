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
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプト作成者は、プラットフォームのプレビュー段階で [外部 API を](https://developer.mozilla.org/docs/Web/API) 使用する場合に一貫した動作を期待すべきではありません。 そのため、重要なスクリプト シナリオでは外部 API に依存しないでください。

外部 API への呼び出しは、[通常の状況では](#external-calls-from-power-automate)Power Automateではなく、Excel アプリケーションを通じてのみ行うことができます。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対してファイアウォール保護を確立できます。

## <a name="configure-your-script-for-external-calls"></a>外部呼び出し用のスクリプトの構成

外部呼び出しは [非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトが `async` . 次に `async` `main` 示すように、関数にプレフィックスを追加し、 `Promise` を返すようにします。

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 他の情報を返すスクリプト `Promise` は、その型の を返すことができます。 たとえば、スクリプトがオブジェクトを返す必要がある場合 `Employee` 、返されるシグネチャは次のようになります。 `: Promise <Employee>`

そのサービスを呼び出すためには、外部サービスのインターフェイスを学習する必要があります。 または REST API を使用している場合は `fetch` 、返されるデータの JSON 構造を決定する必要があります。 [](https://wikipedia.org/wiki/Representational_state_transfer) スクリプトへの入力とスクリプトからの出力の両方について、必要な `interface` JSON 構造に一致するように を作成することを検討してください。 これにより、スクリプトのタイプ セーフが強化されます。 この例については、「 [Office スクリプトからのフェッチを使用する 」を参照してください](../resources/samples/external-fetch-calls.md)。

### <a name="limitations-with-external-calls-from-office-scripts"></a>Office スクリプトからの外部呼び出しに関する制限事項

* サインインしたり、OAuth2 タイプの認証フローを使用する方法はありません。 すべてのキーと資格情報は、ハードコード (または別のソースから読み取る) する必要があります。
* API の資格情報とキーを格納するインフラストラクチャはありません。 これはユーザーが管理する必要があります。
* ドキュメントの Cookie、 `localStorage` および `sessionStorage` オブジェクトはサポートされていません。 
* 外部呼び出しによって、機密データが望ましくないエンドポイントに公開されたり、外部データが内部ワークブックに取り込まれる可能性があります。 管理者は、このような呼び出しに対してファイアウォール保護を確立できます。 外部呼び出しに依存する前に、必ずローカル ポリシーを確認してください。
* 依存関係を取得する前に、データ スループットの量を確認してください。 たとえば、外部データセット全体をプルダウンするのが最適な方法ではない場合があり、代わりにページ分割を使用してデータをチャンク単位で取得する必要があります。

## <a name="retrieve-information-with-fetch"></a>で情報を取得 `fetch`

[フェッチ API は](https://developer.mozilla.org/docs/Web/API/Fetch_API)、外部サービスから情報を取得します。 これは `async` API なので、スクリプトの署名を調整する必要 `main` があります。 関数を `main` 作成 `async` し、 を返します `Promise<void>` 。 また、 `await` `fetch` 呼び出しと取得を確認する必要があります `json` 。 これにより、スクリプトが終了する前にこれらの操作が完了します。

によって取得される JSON データは、 `fetch` スクリプトで定義されているインターフェイスと一致する必要があります。 返される値は[、Office スクリプトは `any` 型をサポートしていないため、特定の型に](typescript-restrictions.md#no-any-type-in-office-scripts)割り当てる必要があります。 返されるプロパティの名前と型を確認するには、サービスのドキュメントを参照してください。 次に、一致するインターフェイスをスクリプトに追加します。

次のスクリプトは `fetch` 、指定された URL のテスト サーバーから JSON データを取得するために使用します。 `JSONData`データを一致する型として格納するインターフェイスに注意してください。

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

### <a name="other-fetch-samples"></a>その他の `fetch` サンプル

* [Officeスクリプトで外部フェッチ呼び出しを使用するサンプルでは](../resources/samples/external-fetch-calls.md)、ユーザーのGitHubリポジトリに関する基本情報を取得する方法を示します。
* [Officeスクリプトのサンプルシナリオ:NOAAからの水位データGraph](../resources/scenarios/noaa-data-fetch.md)は、米国海洋大気局の潮汐と電流データベースからレコードを取得するために使用されているフェッチコマンドを示しています。

## <a name="external-calls-from-power-automate"></a>Power Automateからの外部通話

Power Automateを指定してスクリプトを実行すると、外部 API 呼び出しが失敗します。 これは、Excelアプリケーションを通じてスクリプトを実行する場合とPower Automateを使用する場合の動作の違いです。 フローに組み込む前に、スクリプトでそのような参照を確認してください。

[Azure AD](/connectors/webcontents/)と共に HTTP を使用するか、他の同等のアクションを使用して、データを外部サービスから取得またはプッシュする必要があります。

> [!WARNING]
> [Power Automate Excelオンライン コネクタ](/connectors/excelonlinebusiness)を介して行われた外部呼び出しは、既存のデータ損失防止ポリシーを守るために失敗します。 ただし、Power Automateを介して実行されるスクリプトは、組織の外部および組織のファイアウォールの外部で実行されます。 この外部環境で悪意のあるユーザーから保護を強化するために、管理者はOfficeスクリプトの使用を制御できます。 管理者は、Power AutomateでExcelオンライン コネクタを無効にするか、Office スクリプト[管理者コントロール](/microsoft-365/admin/manage/manage-office-scripts-settings)を使用してExcel on the web用Officeスクリプトを無効にできます。

## <a name="see-also"></a>関連項目

* [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
* [Office スクリプトで外部取得呼び出しを使用する](../resources/samples/external-fetch-calls.md)
* [Officeスクリプトのサンプル シナリオ: NOAA からの水位データのGraph](../resources/scenarios/noaa-data-fetch.md)
