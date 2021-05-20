---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのデバッグに関するヒントとテクニック、およびヘルプ リソース。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545558"
---
# <a name="troubleshoot-office-scripts"></a>Office スクリプトのトラブルシューティング

Officeスクリプトを開発する際に、間違いを犯す可能性があります。 大丈夫です。 問題を見つけ、スクリプトを完璧に動作させるために役立つツールがあります。

## <a name="types-of-errors"></a>エラーの種類

Officeスクリプト エラーは、次の 2 つのカテゴリのいずれかに分類されます。

* コンパイル時のエラーまたは警告
* ランタイム エラー

### <a name="compile-time-errors"></a>コンパイル時エラー

コンパイル時のエラーと警告は、最初はコード エディターに表示されます。 これらは、エディタで赤い波線で示されます。 また、[コード エディタ] 作業ウィンドウの下部にある [ **問題** ] タブにも表示されます。 エラーを選択すると、問題の詳細が表示され、解決策が提案されます。 コンパイル時エラーは、スクリプトを実行する前に対処する必要があります。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに表示されるコンパイラ エラー":::

オレンジ色の警告の下線や灰色の情報メッセージが表示されることもあります。 これらは、スクリプトが意図しない影響を及ぼす可能性があるパフォーマンスの提案やその他の可能性を示します。 このような警告は、それらを却下する前に綿密に検討する必要があります。

### <a name="runtime-errors"></a>ランタイム エラー

ランタイム エラーは、スクリプトのロジックの問題が原因で発生します。 スクリプトで使用されているオブジェクトがブックに含まれなかったり、テーブルの書式が予想と異なる場合、スクリプトの要件と現在のブックとの間に若干の不一致が生じ、その他の若干の相違が生じます。 次のスクリプトは、"TestSheet" という名前のワークシートが存在しない場合にエラーを生成します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>コンソール メッセージ

コンパイル時エラーと実行時エラーの両方が、スクリプトの実行時にコンソールにエラーメッセージを表示します。 問題が発生した行番号を指定します。 問題の根本原因は、コンソールに示されているものとは異なるコード行である可能性があります。

次の図は、[明示的 `any` ](../develop/typescript-restrictions.md)なコンパイラ エラーのコンソール出力を示しています。 `[5, 16]`エラー文字列の先頭にあるテキストを書き留めます。 これは、エラーが 5 行目で、文字 16 から始まる場合を示します。
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="コード エディターコンソールで明示的な 'any' エラー メッセージが表示される":::

次の図は、ランタイム エラーのコンソール出力を示しています。 ここでは、既存のワークシートの名前を持つワークシートを追加しようとします。 ここでも、エラーの前にある 「行 2」に注意して、調査する行を示します。
:::image type="content" source="../images/runtime-error-console.png" alt-text="'addWorksheet' 呼び出しからのエラーを表示するコード エディター コンソール":::

## <a name="console-logs"></a>コンソールログ

ステートメントを使用してメッセージを画面に出力 `console.log` します。 これらのログには、変数の現在の値や、トリガーされるコードパスを示すことができます。 これを行うには、 `console.log` 任意のオブジェクトをパラメーターとして呼び出します。 通常、 `string` コンソールで読み取るのが最も簡単なタイプは a です。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

渡された文字列 `console.log` は、作業ウィンドウの下部にあるコード エディタのログ コンソールに表示されます。 ログは **[出力** ]タブにありますが、ログが書き込まれると自動的にフォーカスが移動します。

ログはブックには影響しません。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>[自動化] タブが表示されないか、スクリプトが使用できないOffice

次の手順は、[**自動化**] タブがExcel on the webに表示されない問題のトラブルシューティングに役立ちます。

1. [Microsoft 365 ライセンスに [Office スクリプト] が含まれていることを確認](../overview/excel.md#requirements)します。
1. [お使いのブラウザがサポートされていることを確認](platform-limits.md#browser-support)します。
1. [サードパーティの Cookie が有効になっていることを確認](platform-limits.md#third-party-cookies)します。
1. [管理者がMicrosoft 365管理センターのOfficeスクリプトを無効にしていないことを確認](/microsoft-365/admin/manage/manage-office-scripts-settings)します。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Power Automateのスクリプトのトラブルシューティング

Power Automateを通じてスクリプトを実行する方法については、「 [Power Automate で実行されているOfficeスクリプトのトラブルシューティング](power-automate-troubleshooting.md)」を参照してください。

## <a name="help-resources"></a>ヘルプ リソース

[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングの問題を支援する開発者のコミュニティです。 スタックオーバーフロー検索を迅速に行うと、問題の解決策を見つけることができます。 それ以下の場合は、質問をして「オフィススクリプト」タグでタグ付けします。 Office アドインではなく、Office *スクリプト* を作成していることを必ずお *知りください*。

Office JavaScript API で問題が発生した場合は[、OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHubリポジトリで問題を作成します。 製品チームのメンバーは、問題に対応し、さらなる支援を提供します。 **OfficeDev/office-js** リポジトリで問題を作成すると、製品チームが対処する必要がある Office JavaScript API ライブラリに問題が見つかったことが示されます。

アクション レコーダまたはエディタに問題がある場合は、Excelの **ヘルプ>フィードバック** ボタンからフィードバックを送信します。

## <a name="see-also"></a>関連項目

- [Office スクリプトのベスト プラクティス](../develop/best-practices.md)
- [Officeスクリプトを使用したプラットフォームの制限](platform-limits.md)
- [Officeスクリプトのパフォーマンスを向上させる](../develop/web-client-performance.md)
- [PowerAutomate で実行されているOffice スクリプトのトラブルシューティング](power-automate-troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
