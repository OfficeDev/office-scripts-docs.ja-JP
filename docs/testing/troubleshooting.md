---
title: Office Scripts のトラブルシューティング
description: Office スクリプトのデバッグのヒントと手法、およびヘルプ リソース。
ms.date: 11/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e673d39b6249ccc7598b832d6478cc8dc0751f6
ms.sourcegitcommit: f5fc9146d5c096e3a580a3fa8f9714147c548df4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2022
ms.locfileid: "66038680"
---
# <a name="troubleshoot-office-scripts"></a>Office Scripts のトラブルシューティング

Office スクリプトを開発すると、間違いを犯す可能性があります。 大丈夫です。 問題を見つけてスクリプトを完全に動作させるためのツールがあります。

> [!NOTE]
> Power Automateを使用したOffice スクリプトに固有のトラブルシューティングのアドバイスについては、「[Power Automateで実行されているOffice スクリプトのトラブルシューティング](power-automate-troubleshooting.md)」を参照してください。

## <a name="types-of-errors"></a>エラーの種類

Office スクリプトエラーは、次の 2 つのカテゴリのいずれかに分類されます。

* コンパイル時のエラーまたは警告
* ランタイム エラー

### <a name="compile-time-errors"></a>コンパイル時エラー

コンパイル時のエラーと警告は、最初はコード エディターに表示されます。 これらは、エディターの波状の赤い下線で表示されます。 また、コード エディター作業ウィンドウの下部にある [ **問題** ] タブにも表示されます。 エラーを選択すると、問題の詳細が表示され、解決策が提案されます。 コンパイル時エラーは、スクリプトを実行する前に対処する必要があります。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="コード エディターのホバー テキストに表示されるコンパイラ エラー。":::

オレンジ色の警告の下線と灰色の情報メッセージが表示される場合もあります。 これらは、パフォーマンスの提案や、スクリプトが意図しない影響を及ぼす可能性があるその他の可能性を示しています。 このような警告は、無視する前に注意深く調べる必要があります。

### <a name="runtime-errors"></a>ランタイム エラー

ランタイム エラーは、スクリプト内のロジックの問題が原因で発生します。 これは、スクリプトで使用されるオブジェクトがブックに含まれていないか、テーブルの形式が予想とは異なる場合や、スクリプトの要件と現在のブックとの間の若干の不一致が原因である可能性があります。 次のスクリプトは、"TestSheet" という名前のワークシートが存在しない場合にエラーを生成します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>コンソール メッセージ

コンパイル時エラーとランタイム エラーの両方で、スクリプトの実行時にコンソールにエラー メッセージが表示されます。 問題が発生した行番号を指定します。 問題の根本原因は、コンソールに示されているものとは異なるコード行である可能性があることに注意してください。

次の図は、 [明示的な `any`](../develop/typescript-restrictions.md) コンパイラ エラーのコンソール出力を示しています。 エラー文字列の先頭にあるテキスト `[5, 16]` に注意してください。 これは、エラーが 5 行目にあり、文字 16 以降であることを示します。
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="明示的な 'any' エラー メッセージを表示するコード エディター コンソール。":::

次の図は、ランタイム エラーのコンソール出力を示しています。 ここでは、スクリプトは、既存のワークシートの名前を持つワークシートを追加しようとします。 ここでも、エラーの前にある "2 行目" に注意して、調査する行を示します。
:::image type="content" source="../images/runtime-error-console.png" alt-text="'addWorksheet' 呼び出しからのエラーを表示するコード エディター コンソール。":::

## <a name="console-logs"></a>コンソール ログ

ステートメントを使用してメッセージを画面に `console.log` 出力します。 これらのログには、変数の現在の値、またはトリガーされるコード パスが表示されます。 これを行うには、パラメーターとして任意のオブジェクトを呼び出 `console.log` します。 通常、a `string` はコンソールで読み取る最も簡単な型です。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

渡された `console.log` 文字列は、作業ウィンドウの下部にあるコード エディターのログ コンソールに表示されます。 ログは [ **出力** ] タブにありますが、タブはログの書き込み時に自動的にフォーカスを取得します。

ログはブックには影響しません。

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>[自動化] タブが表示されないか、スクリプトを使用できないOfficeする

次の手順は、[**自動化**] タブがExcel on the webに表示されない問題のトラブルシューティングに役立ちます。

1. [Microsoft 365 ライセンスにOfficeスクリプトが含まれていることを確認します](../overview/excel.md#requirements)。
1. [ブラウザーがサポートされていることを確認します](platform-limits.md#browser-support)。
1. [サード パーティの Cookie が有効になっていることを確認します](platform-limits.md#third-party-cookies)。
1. [管理者がMicrosoft 365 管理センターのスクリプトOffice無効になっていないことを確認](/microsoft-365/admin/manage/manage-office-scripts-settings)します。
1. テナントに外部ユーザーまたはゲスト ユーザーとしてログインしていないことを確認します。

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>ヘルプ リソース

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) は、コーディングの問題を支援する開発者のコミュニティです。 多くの場合、スタック オーバーフローのクイック検索を使用して問題の解決策を見つけることができます。 そうでない場合は、質問をして、"office-scripts" タグでタグ付けします。 Office *アドイン* ではなく、Office *スクリプト* を作成していることを確認してください。

## <a name="see-also"></a>関連項目

- [Office スクリプトでのベスト プラクティス](../develop/best-practices.md)
- [Office スクリプトを使用したプラットフォームの制限](platform-limits.md)
- [Office Scripts のパフォーマンスの改善](../develop/web-client-performance.md)
- [PowerAutomate で実行されているOffice スクリプトのトラブルシューティング](power-automate-troubleshooting.md)
- [Office スクリプトの効果を元に戻す](undo.md)
