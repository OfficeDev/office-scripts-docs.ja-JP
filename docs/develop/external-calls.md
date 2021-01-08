---
title: Office スクリプトでの外部 API 呼び出しのサポート
description: 外部スクリプトで外部 API 呼び出しを行う場合のサポートOfficeガイダンス。
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784145"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプト作成者は、プラットフォームのプレビュー フェーズ中に外部 [API](https://developer.mozilla.org/docs/Web/API) を使用するときに一貫した動作を期待してはならない。 そのため、重要なスクリプト シナリオでは外部 API に依存しません。

外部 API への呼び出しは、通常の状況では Power Automate 経由ではなく、Excel アプリケーション [を介して行う必要があります](#external-calls-from-power-automate)。

> [!CAUTION]
> 外部呼び出しでは、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対してファイアウォール保護を確立できます。

## <a name="working-with-fetch"></a>操作 `fetch`

フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。 これは `async` API なので、スクリプトの署名を `main` 調整する必要があります。 関数を `main` 作成 `async` し、それを返す `Promise<void>` 必要があります。 通話と取得も `await` `fetch` 必ず行う必要 `json` があります。 これにより、スクリプトが終了する前にこれらの操作が確実に完了します。

次のスクリプトは、 `fetch` 指定された URL のテスト サーバーから JSON データを取得するために使用します。

```typescript
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

Office スクリプトのサンプル シナリオ [: NOAA](../resources/scenarios/noaa-data-fetch.md) からのグラフの水レベル データは、National National Wateric and の Administration の大島と現在のデータベースからレコードを取得するために使用されるフェッチ コマンドを示しています。

## <a name="external-calls-from-power-automate"></a>Power Automate からの外部通話

Power Automate を使用してスクリプトを実行すると、外部 API 呼び出しは失敗します。 これは、Excel クライアントを使用してスクリプトを実行する場合と Power Automate を使用する場合の動作の違いです。 フローに組み込む前に、スクリプトでそのような参照を確認してください。

> [!WARNING]
> Power Automate [Excel Online](/connectors/excelonlinebusiness) コネクタを介して行われる外部呼び出しは、既存のデータ損失防止ポリシーをサポートするために失敗します。 ただし、Power Automate を介して実行されるスクリプトは、組織の外部および組織のファイアウォールの外側で実行されます。 この外部環境で悪意のあるユーザーからの保護を強化するために、管理者はスクリプトの使用Officeできます。 管理者は、Power Automate で Excel Online コネクタを無効にするか、Office Scripts 管理者による Web 上の Excel 用スクリプトの Office スクリプトを [無効にできます](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>関連項目

- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
- [Office スクリプトのサンプル シナリオ: NOAA からのグラフの水レベル データ](../resources/scenarios/noaa-data-fetch.md)
