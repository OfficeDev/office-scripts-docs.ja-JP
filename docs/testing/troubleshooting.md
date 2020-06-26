---
title: Office スクリプトのトラブルシューティング
description: Office スクリプトのヒントとテクニック、およびヘルプリソースをデバッグします。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 6448980eec45214a589444229db0fd781b9fea13
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878620"
---
# <a name="troubleshooting-office-scripts"></a>Office スクリプトのトラブルシューティング

Office スクリプトを開発する際には、誤りが発生することがあります。 大丈夫です。 問題を見つけてスクリプトを完全に動作させるためのツールが用意されています。

## <a name="console-logs"></a>コンソールログ

トラブルシューティング中に、画面にメッセージを出力することもできます。 これにより、変数の現在の値や、どのコードパスがトリガーされているかを確認できます。 これを行うには、テキストをコンソールに記録します。

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

に渡される文字列 `console.log` は、コードエディターのログコンソールに表示されます。 コンソールをオンにするには、**省略記号**ボタンを押して [**ログ...** ] を選択します。

ログがブックに影響を与えることはありません。

## <a name="error-messages"></a>エラー メッセージ

Excel スクリプトで問題が発生すると、エラーが生成されます。 **ログを表示**するかどうかの確認を求めるポップアップが表示されます。 そのボタンを押してコンソールを開き、エラーを表示します。

## <a name="help-resources"></a>ヘルプリソース

[スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)は、コーディングに関する問題の解決に役立つ開発者のコミュニティです。 多くの場合、クイックスタックオーバーフロー検索を使用して、問題の解決策を見つけることができます。 そうでない場合は、質問して、「office スクリプト」タグでタグを付けてください。 Office*アドイン*ではなく、office*スクリプト*を作成していることを必ずお伝えください。

Office JavaScript API で問題が発生した場合は、 [Officedev/Office/js](https://github.com/OfficeDev/office-js) GitHub リポジトリに問題を作成します。 製品チームのメンバーは問題に対応し、さらに支援を提供します。 **Officedev/office-js**リポジトリで問題を発生させることは、製品チームが対処する必要のある OFFICE JavaScript API ライブラリに問題が見つかったことを示しています。

操作レコーダーまたは Editor に問題がある場合は、Excel の**ヘルプ > フィードバック**ボタンを使用してフィードバックを送信してください。

## <a name="see-also"></a>関連項目

- [Excel on the web の Office スクリプト](../overview/excel.md)
- [Web 上の Excel での Office スクリプトのスクリプトの基礎](../develop/scripting-fundamentals.md)
- [Office スクリプトの効果を元に戻す](undo.md)
- [Office スクリプトのパフォーマンスを向上させる](../develop/web-client-performance.md)
