---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: Office スクリプトで外部 API 呼び出しを行うためのサポートとガイダンス。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088044"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプトは、外部サービスへの呼び出しをサポートします。 これらのサービスを使用して、ブックにデータやその他の情報を提供します。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。

> [!IMPORTANT]
> 外部 API の呼び出しは、[通常の状況](#external-calls-from-power-automate)ではPower Automateを通じてではなく、Excel アプリケーションを介してのみ行うことができます。

## <a name="configure-your-script-for-external-calls"></a>外部呼び出し用にスクリプトを構成する

外部呼び出しは [非同期](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) であり、スクリプトに `async`. 次に `async` 示すように、関数にプレフィックスを `main` 追加して返 `Promise`すようにします。

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 他の情報を返すスクリプトは、その型を `Promise` 返すことができます。 たとえば、スクリプトでオブジェクトを返す必要がある `Employee` 場合、戻りシグネチャは次のようになります。 `: Promise <Employee>`

そのサービスを呼び出すには、外部サービスのインターフェイスについて学習する必要があります。 [API または REST API を](https://wikipedia.org/wiki/Representational_state_transfer)使用`fetch`している場合は、返されるデータの JSON 構造を決定する必要があります。 スクリプトとの間の入力と出力の両方について、必要な JSON 構造に一致するようにすることを `interface` 検討してください。 これにより、スクリプトの安全性が向上します。 この例については、「[Office スクリプトからのフェッチの使用](../resources/samples/external-fetch-calls.md)」を参照してください。

### <a name="limitations-with-external-calls-from-office-scripts"></a>Office スクリプトからの外部呼び出しに関する制限事項

* サインインしたり、OAuth2 の種類の認証フローを使用したりする方法はありません。 すべてのキーと資格情報はハードコーディングする (または別のソースから読み取る) 必要があります。
* API 資格情報とキーを格納するインフラストラクチャはありません。 これはユーザーが管理する必要があります。
* ドキュメント Cookie、 `localStorage`および `sessionStorage` オブジェクトはサポートされていません。
* 外部呼び出しにより、機密データが望ましくないエンドポイントに公開されたり、外部データが内部ブックに取り込まれたりすることがあります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。 外部呼び出しに依存する前に、必ずローカル ポリシーを確認してください。
* 依存関係を取得する前に、データ スループットの量を確認してください。 たとえば、外部データセット全体をプルダウンすることは最適なオプションではない可能性があり、代わりに改ページ分割を使用してデータをチャンク単位で取得する必要があります。

## <a name="retrieve-information-with-fetch"></a>を使用して情報を取得する `fetch`

[フェッチ API](https://developer.mozilla.org/docs/Web/API/Fetch_API) は、外部サービスから情報を取得します。 これは `async` API であるため、スクリプトのシグネチャを調整する `main` 必要があります。 関数`async`を作成します`main`。 また、呼び出しと`json`取得も`await``fetch`必ず行う必要があります。 これにより、スクリプトが終了する前にこれらの操作が確実に完了します。

取得される `fetch` JSON データは、スクリプトで定義されているインターフェイスと一致する必要があります。 Office スクリプトでは型がサポートされていないため、返される値を特定[の型に`any`](typescript-restrictions.md#no-any-type-in-office-scripts)割り当てる必要があります。 返されるプロパティの名前と型を確認するには、サービスのドキュメントを参照する必要があります。 次に、一致するインターフェイスまたはインターフェイスをスクリプトに追加します。

次のスクリプトは、指定された URL のテスト サーバーから JSON データを取得するために使用 `fetch` します。 データを `JSONData` 一致する型として格納するインターフェイスに注意してください。

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
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

* [Office スクリプトサンプルで外部フェッチ呼び出しを使用](../resources/samples/external-fetch-calls.md)するサンプルでは、ユーザーのGitHubリポジトリに関する基本情報を取得する方法を示します。
* [Office スクリプトのサンプル シナリオ: NOAA から水レベルのデータをGraph](../resources/scenarios/noaa-data-fetch.md)すると、海洋大気局の Tides および Currents データベースからレコードを取得するために使用される fetch コマンドが示されています。

## <a name="external-calls-from-power-automate"></a>Power Automateからの外部呼び出し

スクリプトがPower Automateで実行されると、外部 API 呼び出しは失敗します。 これは、Excel アプリケーションを使用してスクリプトを実行することと、Power Automateを介してスクリプトを実行する場合の動作上の違いです。 フローにビルドする前に、スクリプトでこのような参照がないか確認してください。

Azure AD またはその他の同等のアクション [で HTTP](/connectors/webcontents/) を使用して、外部サービスからデータをプルするか、外部サービスにプッシュする必要があります。

> [!WARNING]
> Power Automate [Excel Online コネクタ](/connectors/excelonlinebusiness)を介して行われた外部呼び出しは、既存のデータ損失防止ポリシーを維持するために失敗します。 ただし、Power Automateを介して実行されるスクリプトは、組織の外部および組織のファイアウォールの外部で実行されます。 この外部環境で悪意のあるユーザーからの保護を強化するために、管理者は Office スクリプトの使用を制御できます。 管理者は、Power Automateで Excel Online コネクタを無効にするか、Office [スクリプト管理者コントロール](/microsoft-365/admin/manage/manage-office-scripts-settings)を使用してExcel on the webのスクリプトOffice無効にすることができます。

## <a name="see-also"></a>関連項目

* [JSON を使用して、Office スクリプトとの間でデータを渡す](use-json.md)
* [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
* [Office スクリプトで外部取得呼び出しを使用する](../resources/samples/external-fetch-calls.md)
* [Office スクリプトのサンプル シナリオ: NOAA から水レベルのデータをGraphする](../resources/scenarios/noaa-data-fetch.md)
