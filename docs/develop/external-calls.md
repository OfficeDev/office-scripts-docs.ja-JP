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
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプト作成者は、プラットフォームのプレビュー 段階で外部 [API](https://developer.mozilla.org/docs/Web/API) を使用する場合、一貫した動作を期待してはならない。 そのため、重要なスクリプト シナリオでは外部 API に依存しません。

外部 API への呼び出しは、通常の状況Excelアプリケーションを介Power Automate[実行できます](#external-calls-from-power-automate)。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。

## <a name="configure-your-script-for-external-calls"></a>外部呼び出し用にスクリプトを構成する

外部呼び出 [しは非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトがとしてマークされている必要があります `async` 。 次に示すように、プレフィックスを関数に追加 `async` `main` し、それを `Promise` 返すようにします。

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 他の情報を返すスクリプトは、その種類の `Promise` 1 つを返す可能性があります。 たとえば、スクリプトでオブジェクトを返す必要がある場合、 `Employee` 戻り値の署名は次のようになります。 `: Promise <Employee>`

そのサービスを呼び出すには、外部サービスのインターフェイスを学習する必要があります。 REST API を使用 `fetch` [している場合](https://wikipedia.org/wiki/Representational_state_transfer)は、返されるデータの JSON 構造を決定する必要があります。 スクリプトの入力と出力の両方について、必要な JSON 構造に一致 `interface` するを検討してください。 これにより、スクリプトの型の安全性が向上します。 この例については、「スクリプトからフェッチを使用する[」でOfficeできます](../resources/samples/external-fetch-calls.md)。

### <a name="limitations-with-external-calls-from-office-scripts"></a>スクリプトからの外部呼び出しOffice制限

* OAuth2 タイプの認証フローをサインインまたは使用する方法はありません。 すべてのキーと資格情報をハードコード (または別のソースから読み取る) 必要があります。
* API の資格情報とキーを格納するインフラストラクチャはありません。 これは、ユーザーが管理する必要があります。
* ドキュメント Cookie、 `localStorage` および `sessionStorage` オブジェクトはサポートされていません。 
* 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される場合や、内部ブックに外部データが取り込まれたりする場合があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。 外部通話に依存する前に、必ずローカル ポリシーに確認してください。
* 依存関係を取得する前に、データ スループットの量を確認してください。 たとえば、外部データセット全体を引き下げないのが最適な選択肢ではなく、代わりにページネーションを使用してデータをチャンク単位で取得する必要があります。

## <a name="retrieve-information-with-fetch"></a>を使用して情報を取得する `fetch`

フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。 これは `async` API なので、スクリプトの署名を `main` 調整する必要があります。 関数を `main` 作成 `async` し、 を返します `Promise<void>` 。 また、呼び出しと取得 `await` `fetch` も確認する必要 `json` があります。 これにより、スクリプトが終了する前にこれらの操作が確実に完了します。

取得した JSON データは、 `fetch` スクリプトで定義されているインターフェイスと一致している必要があります。 スクリプトは型をサポートしていないのでOffice値を特定の型[に割り当てる必要 `any` があります](typescript-restrictions.md#no-any-type-in-office-scripts)。 返されるプロパティの名前と種類については、サービスのドキュメントを参照してください。 次に、一致するインターフェイスまたはインターフェイスをスクリプトに追加します。

次のスクリプトは、指定された URL のテスト サーバーから `fetch` JSON データを取得するために使用します。 データを `JSONData` 一致する型として格納するインターフェイスに注意してください。

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

### <a name="other-fetch-samples"></a>その他 `fetch` のサンプル

* [[スクリプトの外部フェッチ](../resources/samples/external-fetch-calls.md)呼び出しOffice使用する] サンプルは、ユーザーのリポジトリに関する基本情報を取得するGitHub示しています。
* Office スクリプトのサンプル シナリオ[: NOAA](../resources/scenarios/noaa-data-fetch.md)の Graph 水レベルデータは、国立海洋大気局のタイドと Currents データベースからレコードを取得するために使用されるフェッチ コマンドを示しています。

## <a name="external-calls-from-power-automate"></a>外部からの外部通話Power Automate

スクリプトを使用してスクリプトを実行すると、外部 API 呼び出しPower Automate。 これは、スクリプトをアプリケーションを介して実行する場合と、Excelスクリプトを実行Power Automate。 フローに組み込む前に、スクリプトでそのような参照を確認してください。

データを外部サービスから取得または外部サービスにプッシュするには [、Azure AD](/connectors/webcontents/) または他の同等のアクションで HTTP を使用する必要があります。

> [!WARNING]
> 既存のデータ損失防止ポリシーを[Power Automate Excel、オンライン](/connectors/excelonlinebusiness)コネクタを介して行われた外部通話は失敗します。 ただし、組織の外部Power Automate、組織のファイアウォールの外部で実行されるスクリプトは実行されます。 この外部環境で悪意のあるユーザーからの保護を強化するために、管理者はスクリプトの使用Officeできます。 管理者は、Excel で Excel Power Automate Online コネクタを無効にするか、Office スクリプト管理者Excel on the webを使用して Office スクリプトを[無効にできます](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>関連項目

* [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
* [Office スクリプトで外部取得呼び出しを使用する](../resources/samples/external-fetch-calls.md)
* [Officeスクリプトのサンプル シナリオ: noAA Graphデータを使用する](../resources/scenarios/noaa-data-fetch.md)
