---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのヒントとテクニック、およびヘルプリソースをデバッグします。
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700358"
---
# <a name="troubleshooting-office-scripts"></a>Office スクリプトのトラブルシューティング

Office スクリプトを開発する際には、誤りが発生することがあります。 大丈夫です。 問題を見つけてスクリプトを完全に動作させるためのツールが用意されています。

## <a name="console-logs"></a>コンソールログ

トラブルシューティング中に、画面にメッセージを出力することもできます。 これにより、変数の現在の値や、どのコードパスがトリガーされているかを確認できます。 これを行うには、テキストをコンソールに記録します。

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> オブジェクトプロパティを`load`ログに記録`sync`する前に、ワークシートデータとブックを忘れずに使用してください。

に`console.log`渡される文字列は、コードエディターのログコンソールに表示されます。 コンソールをオンにするには、**省略記号**ボタンを押して [**ログ...** ] を選択します。

ログがブックに影響を与えることはありません。

## <a name="error-messages"></a>エラー メッセージ

Excel スクリプトで問題が発生すると、エラーが生成されます。 **ログを表示**するかどうかの確認を求めるポップアップが表示されます。 そのボタンを押してコンソールを開き、エラーを表示します。

## <a name="help-resources"></a>ヘルプリソース

[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)は、コーディングに関する問題の解決に役立つ開発者のコミュニティです。 多くの場合、クイックスタックオーバーフロー検索を使用して、問題の解決策を見つけることができます。 そうでない場合は、質問して、「office スクリプト」タグでタグを付けてください。 Office*アドイン*ではなく、office*スクリプト*を作成していることを必ずお伝えください。

Office JavaScript API で問題が発生した場合は、 [Officedev/Office/js](https://github.com/OfficeDev/office-js) GitHub リポジトリに問題を作成します。 製品チームのメンバーは問題に対応し、さらに支援を提供します。 **Officedev/office-js**リポジトリで問題を発生させることは、製品チームが対処する必要のある OFFICE JavaScript API ライブラリに問題が見つかったことを示しています。

操作レコーダーまたは Editor に問題がある場合は、Excel の**ヘルプ > フィードバック**ボタンを使用してフィードバックを送信してください。

## <a name="see-also"></a>関連項目

- [Web 上の Excel での Office スクリプト](../overview/excel.md)
- [Web 上の Excel での Office スクリプトのスクリプトの基礎](../develop/scripting-fundamentals.md)
- [Office スクリプトの効果を元に戻す](undo.md)
