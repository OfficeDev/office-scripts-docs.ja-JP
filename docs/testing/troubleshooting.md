---
title: スクリプトOfficeトラブルシューティング
description: スクリプトのデバッグのヒントとOfficeヘルプ リソース。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545558"
---
# <a name="troubleshoot-office-scripts"></a>スクリプトOfficeトラブルシューティング

スクリプトを開発Office、間違いを犯す可能性があります。 大丈夫です。 問題を見つけてスクリプトを完全に機能するためのツールがあります。

## <a name="types-of-errors"></a>エラーの種類

Officeスクリプトエラーは、次の 2 つのカテゴリに分類されます。

* コンパイル時のエラーまたは警告
* ランタイム エラー

### <a name="compile-time-errors"></a>コンパイル時エラー

コンパイル時のエラーと警告は、最初はコード エディターに表示されます。 これらは、エディターの波状の赤い下線で表示されます。 また、[コード エディター] 作業ウィンドウ **の** 下部にある [問題] タブにも表示されます。 エラーを選択すると、問題の詳細と解決策の提案が表示されます。 コンパイル時のエラーは、スクリプトを実行する前に対処する必要があります。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに表示されるコンパイラ エラー":::

オレンジ色の警告の下線と灰色の情報メッセージが表示される場合があります。 これらは、スクリプトが意図しない効果を持つ可能性があるパフォーマンスの提案や他の可能性を示します。 このような警告は、却下する前に注意して調べる必要があります。

### <a name="runtime-errors"></a>ランタイム エラー

ランタイム エラーは、スクリプトのロジックの問題が原因で発生します。 これは、スクリプトで使用されるオブジェクトがブック内に含めなかったり、テーブルの形式が予想と異なっている、またはスクリプトの要件と現在のブックの間に若干の不一致が生じていった場合に発生する可能性があります。 次のスクリプトは、"TestSheet" という名前のワークシートが存在しない場合にエラーを生成します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>コンソール メッセージ

コンパイル時と実行時の両方のエラーは、スクリプトの実行時にコンソールにエラー メッセージを表示します。 問題が発生した行番号を指定します。 問題の根本原因は、コンソールで示されているコードとは異なるコード行である可能性があります。

次の図は、明示的なコンパイラ エラーのコンソール[出力を `any` ](../develop/typescript-restrictions.md)示しています。 エラー文字列の `[5, 16]` 先頭にあるテキストに注意してください。 これは、エラーが 5 行目で、文字 16 から始まるかどうかを示します。
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="明示的な 'any' エラー メッセージを表示するコード エディター コンソール":::

次の図は、実行時エラーのコンソール出力を示しています。 ここでは、既存のワークシートの名前を持つワークシートを追加します。 ここでも、エラーの前の "2 行目" に注意して、調査する行を表示します。
:::image type="content" source="../images/runtime-error-console.png" alt-text="'addWorksheet' 呼び出しからのエラーを表示するコード エディター コンソール":::

## <a name="console-logs"></a>コンソール ログ

ステートメントを使用してメッセージを画面に印刷 `console.log` します。 これらのログには、変数の現在の値、またはトリガーされるコード パスが表示されます。 これを行うには、任意 `console.log` のオブジェクトをパラメーターとして呼び出します。 通常、コンソール `string` で読み取りが最も簡単な型は a です。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

渡された文字列は、作業ウィンドウの下部にあるコード エディターのログ コンソール `console.log` に表示されます。 ログは [出力] タブ **にあります** が、ログの書き込み時にタブが自動的にフォーカスを取得します。

ログはブックには影響を与えかねない。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>[自動化] タブが表示されないか、Officeスクリプトが使用できない

次の手順は、[自動化] タブに関連する問題のトラブルシューティングに役立つExcel on the web。

1. [ライセンスにスクリプトMicrosoft 365含Officeしてください](../overview/excel.md#requirements)。
1. [ブラウザーがサポートされていないことを確認します](platform-limits.md#browser-support)。
1. [サードパーティの Cookie が有効になっているか確認します](platform-limits.md#third-party-cookies)。
1. [管理者が管理センターのスクリプトOffice無効にMicrosoft 365します](/microsoft-365/admin/manage/manage-office-scripts-settings)。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>スクリプトのトラブルシューティングを行Power Automate

スクリプトの実行に関する詳細については、「Power Automateで実行されているスクリプトOffice[トラブルシューティング」を参照Power Automate。](power-automate-troubleshooting.md)

## <a name="help-resources"></a>ヘルプ リソース

[スタック オーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングの問題を支援する開発者のコミュニティです。 多くの場合、スタック オーバーフローのクイック検索を使用して、問題の解決策を見つける可能性があります。 そうでない場合は、質問をして"office-scripts" タグでタグ付けします。 アドインではなく、Office *スクリプト* を作成Office *してください*。

JavaScript API で問題がOffice場合は[、OfficeDev/office-js](https://github.com/OfficeDev/office-js)リポジトリにGitHubしてください。 製品チームのメンバーは、問題に対応し、さらに支援を提供します。 **OfficeDev/office-js** リポジトリに問題を作成すると、製品チームが対処する必要がある javaScript API ライブラリOfficeに欠陥が見つかりました。

アクション レコーダーまたはエディターで問題が発生した場合は、ヘルプ ウィンドウの [フィードバック] >をExcel。

## <a name="see-also"></a>関連項目

- [Office スクリプトでのベスト プラクティス](../develop/best-practices.md)
- [スクリプトを使用したプラットフォームOffice制限](platform-limits.md)
- [スクリプトのパフォーマンスをOfficeする](../develop/web-client-performance.md)
- [PowerAutomate Office実行されているスクリプトのトラブルシューティング](power-automate-troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
