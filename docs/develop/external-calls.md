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
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

スクリプト作成者は、プラットフォームのプレビュー 段階で外部 [API](https://developer.mozilla.org/docs/Web/API) を使用する場合、一貫した動作を期待してはならない。 そのため、重要なスクリプト シナリオでは外部 API に依存しません。

外部 API への呼び出しは、通常の状況では Power Automate 経由ではなく、Excel アプリケーション経由 [でのみ実行できます](#external-calls-from-power-automate)。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対するファイアウォール保護を確立できます。

## <a name="working-with-fetch"></a>操作 `fetch`

フェッチ [API は、](https://developer.mozilla.org/docs/Web/API/Fetch_API) 外部サービスから情報を取得します。 これは `async` API なので、スクリプトの署名を `main` 調整する必要があります。 関数を `main` 作成 `async` し、 を返します `Promise<void>` 。 また、呼び出しと取得 `await` `fetch` も確認する必要 `json` があります。 これにより、スクリプトが終了する前にこれらの操作が確実に完了します。

次のスクリプトは、指定された URL のテスト サーバーから `fetch` JSON データを取得するために使用します。

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

Office スクリプトのサンプル シナリオ [: NOAA](../resources/scenarios/noaa-data-fetch.md) の水位データをグラフ化すると、国立海洋大気局の潮流データベースからレコードを取得するために使用されるフェッチ コマンドが示されています。

## <a name="external-calls-from-power-automate"></a>Power Automate からの外部通話

Power Automate を使用してスクリプトを実行すると、外部 API 呼び出しは失敗します。 これは、Excel クライアントを使用してスクリプトを実行する場合と Power Automate を使用する場合の動作の違いです。 フローに組み込む前に、スクリプトでそのような参照を確認してください。

> [!WARNING]
> Power [Automate Excel Online](/connectors/excelonlinebusiness) コネクタを介して行われた外部呼び出しは、既存のデータ損失防止ポリシーを支持するために失敗します。 ただし、Power Automate を介して実行されるスクリプトは、組織外および組織のファイアウォールの外部で実行されます。 この外部環境で悪意のあるユーザーから保護するために、管理者はスクリプトの使用Officeできます。 管理者は、Power Automate で Excel Online コネクタを無効にするか、Office スクリプト管理者コントロールを使用して Web 上の Excel Office [を無効にできます](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>関連項目

- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)
- [Office スクリプトのサンプル シナリオ: NOAA からの水位データのグラフ](../resources/scenarios/noaa-data-fetch.md)
