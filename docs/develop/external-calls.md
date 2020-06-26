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
# <a name="external-api-call-support-in-office-scripts"></a>Office スクリプトでの外部 API 呼び出しのサポート

Office スクリプトプラットフォームは、[外部 api](https://developer.mozilla.org/docs/Web/API)への呼び出しをサポートしていません。 ただし、これらの呼び出しは適切な状況で実行することができます。 外部呼び出しは、Excel クライアントを使用してのみ行うことができます。[通常の状況で](#external-calls-from-power-automate)は、電力の自動処理は行われません。

スクリプト作成者は、プラットフォームのプレビューフェーズ中に外部 Api を使用するときに、一貫した動作を期待する必要はありません。 これは、JavaScript ランタイムがブックとの対話を管理する方法に起因します。 このスクリプトは、API 呼び出しが完了する前に終了することができます (または `Promise` 完全に解決される)。 そのため、重要なスクリプトシナリオでは外部 Api に依存しません。

> [!CAUTION]
> 外部呼び出しにより、機密データが望ましくないエンドポイントに公開される可能性があります。 管理者は、このような呼び出しに対してファイアウォール保護を確立できます。

## <a name="definition-files-for-external-apis"></a>外部 Api の定義ファイル

Office スクリプトには、外部 Api の定義ファイルは含まれていません。 このような Api を使用すると、定義が欠落しているとコンパイル時エラーが生成されます。 次のスクリプトに示すように、Api は引き続き実行されます (ただし、Excel クライアントで実行する場合のみ)。

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

## <a name="external-calls-from-power-automate"></a>電源自動化からの外部通話

電源自動化を使用してスクリプトを実行すると、外部 API 呼び出しは失敗します。 これは、Excel クライアントを使用してスクリプトを実行する場合と Power オートメーションを使用する場合の動作の違いです。 スクリプトをフローに組み込む前に、そのような参照について必ずチェックしてください。

> [!WARNING]
> Power [Online](/connectors/excelonlinebusiness)の外部通話の失敗は、既存のデータ損失防止ポリシーを守るために役立ちます。 ただし、電源自動化によって実行されるスクリプトは、組織外、組織のファイアウォールの外側にあります。 この外部環境で悪意のあるユーザーからの保護を強化するために、管理者は Office スクリプトの使用を制御することができます。 管理者は、Power turn で Excel Online コネクタを無効にするか、 [Office scripts administrator コントロール](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)を使用して web 上の Excel の office スクリプトをオフにすることができます。

## <a name="see-also"></a>関連項目

- [Office スクリプトでの組み込みの JavaScript オブジェクトの使用](javascript-objects.md)