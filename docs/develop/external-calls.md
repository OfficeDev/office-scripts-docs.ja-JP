---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: スクリプトで外部 API 呼び出しを行うOffice。
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: e7be505f13529e1d3bcff22ce9fa18cc36148f7b
ms.sourcegitcommit: 79ce4fad6d284b1aa71f5ad6d2938d9ad6a09fee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2022
ms.locfileid: "63459607"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプトは、外部サービスへの呼び出しをサポートします。 これらのサービスを使用して、ブックにデータなどの情報を提供します。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。

> [!IMPORTANT]
> 外部 API への呼び出しは、通常の状況Excel、Power Automateを介して[行う必要があります](#external-calls-from-power-automate)。

## <a name="configure-your-script-for-external-calls"></a>外部呼び出し用にスクリプトを構成する

外部呼び出 [しは非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトがとしてマークされている必要があります `async`。 次に示 `async` すように、 `main` プレフィックスを関数に追加し、それを `Promise`返すようにします。

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 他の情報を返すスクリプトは、その種類の 1 `Promise` つを返す可能性があります。 たとえば、スクリプトでオブジェクトを返す必要 `Employee` がある場合、戻り値の署名は次のようになります。 `: Promise <Employee>`

そのサービスを呼び出すには、外部サービスのインターフェイスを学習する必要があります。 REST API を使用`fetch`[している場合](https://wikipedia.org/wiki/Representational_state_transfer)は、返されるデータの JSON 構造を決定する必要があります。 スクリプトの入力と出力の両方について `interface` 、必要な JSON 構造に一致するを検討してください。 これにより、スクリプトの型の安全性が向上します。 この例については、「スクリプトからフェッチを使用する[」でOfficeできます](../resources/samples/external-fetch-calls.md)。

### <a name="limitations-with-external-calls-from-office-scripts"></a>スクリプトからの外部呼び出しOffice制限

* OAuth2 タイプの認証フローをサインインまたは使用する方法はありません。 すべてのキーと資格情報をハードコード (または別のソースから読み取る) 必要があります。
* API の資格情報とキーを格納するインフラストラクチャはありません。 これは、ユーザーが管理する必要があります。
* ドキュメント Cookie、 `localStorage`およびオブジェクト `sessionStorage` はサポートされていません。
* 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される場合や、内部ブックに外部データが取り込まれたりする場合があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。 外部通話に依存する前に、必ずローカル ポリシーに確認してください。
* 依存関係を取得する前に、データ スループットの量を確認してください。 たとえば、外部データセット全体を引き下げないのが最適な選択肢ではなく、代わりにページネーションを使用してデータをチャンク単位で取得する必要があります。

## <a name="retrieve-information-with-fetch"></a>を使用して情報を取得する `fetch`

フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。 これは API なので `async` 、スクリプトの署名を `main` 調整する必要があります。 関数を作成 `main` します `async`。 また、呼び出しと取得 `await` も確認 `fetch` する必要 `json` があります。 これにより、スクリプトが終了する前にこれらの操作が確実に完了します。

取得した JSON データは、スクリプト `fetch` で定義されているインターフェイスと一致している必要があります。 返される値は、スクリプトが型をサポートしていないOffice型に割[り当てる必要`any`があります](typescript-restrictions.md#no-any-type-in-office-scripts)。 返されるプロパティの名前と種類については、サービスのドキュメントを参照してください。 次に、一致するインターフェイスまたはインターフェイスをスクリプトに追加します。

次のスクリプトは、指定 `fetch` された URL のテスト サーバーから JSON データを取得するために使用します。 データを `JSONData` 一致する型として格納するインターフェイスに注意してください。

```TypeScript
async function main(workbook: ExcelScript.Workbook){
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

* [[スクリプトの外部フェッチ](../resources/samples/external-fetch-calls.md)呼びOffice使用する] サンプルは、ユーザーのリポジトリに関する基本情報を取得するGitHub示しています。
* Office スクリプトのサンプル シナリオ[: NOAA の Graph](../resources/scenarios/noaa-data-fetch.md) 水レベルデータは、全米海洋大気局のタイドと Currents データベースからレコードを取得するために使用されるフェッチ コマンドを示しています。

## <a name="external-calls-from-power-automate"></a>ユーザーからの外部通話Power Automate

スクリプトを実行すると、外部 API 呼び出しは失敗し、Power Automate。 これは、スクリプトをアプリケーション経由で実行する場合と、Excelスクリプトを実行する場合Power Automate。 フローに組み込む前に、スクリプトでそのような参照を確認してください。

外部サービスからデータを取得または外部サービスにプッシュするには、Azure ADまたは他の同等のアクションで [HTTP](/connectors/webcontents/) を使用する必要があります。

> [!WARNING]
> 既存のデータ損失防止ポリシーをPower Automate Excel[、](/connectors/excelonlinebusiness)オンライン コネクタを介して行われた外部呼び出しは失敗します。 ただし、組織のファイアウォールPower Automate実行されるスクリプトは、組織の外部および組織のファイアウォールの外部で実行されます。 この外部環境で悪意のあるユーザーからの保護を強化するために、管理者はスクリプトの使用Officeできます。 管理者は、Power Automate で Excel Online コネクタを無効にするか、Office スクリプト管理者Excel on the webをOffice[できます](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>関連項目

* [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
* [Office スクリプトで外部取得呼び出しを使用する](../resources/samples/external-fetch-calls.md)
* [Officeスクリプトのサンプル シナリオ: NOAA Graphの水レベルデータの作成](../resources/scenarios/noaa-data-fetch.md)
