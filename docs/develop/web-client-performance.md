---
title: Office スクリプトのパフォーマンスを向上させる
description: Excel ブックとスクリプトの間の通信を理解することで、より高速なスクリプトを作成できます。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878899"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Office スクリプトのパフォーマンスを向上させる

Office スクリプトの目的は、頻繁に実行される一連のタスクを自動化して時間を節約することです。 低速なスクリプトは、ワークフローを高速化しないように感じられます。 ほとんどの場合、スクリプトは完全に機能し、期待どおりに実行されます。 ただし、パフォーマンスに影響する可能性のあるいくつかの avoidable のシナリオがあります。

時間のかかるスクリプトの最も一般的な原因は、ブックとの通信が多すぎることです。 スクリプトは、ローカルコンピューター上で実行されます。ブックはクラウド内に存在します。 場合によっては、スクリプトによってローカルデータがブックの内容と同期されます。 これは、このようなバックグラウンドでの同期が発生したときに、(などの) 書き込み操作 `workbook.addWorksheet()` がブックにのみ適用されることを意味します。 同様に、どのような読み取り操作 (など) でも、その `myRange.getValues()` 時点でスクリプトのブックからデータを取得します。 どちらの場合も、スクリプトはデータを処理する前に情報をフェッチします。 たとえば、次のコードでは、使用されている範囲内の行数を正確に記録します。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office スクリプト Api は、ブックまたはスクリプト内のすべてのデータが正確で、必要に応じて最新の状態になっていることを確認します。 スクリプトを正しく実行するために、これらの同期について心配する必要はありません。 ただし、このスクリプトからクラウドへの通信を認識することは、不要なネットワーク呼び出しを回避するのに役立ちます。

## <a name="performance-optimizations"></a>パフォーマンスの最適化

クラウドへの通信を減らすための簡単な手法を適用することができます。 次のパターンは、スクリプトを高速化するのに役立ちます。

- 繰り返しループ内ではなく、ブックのデータを1回読み取ります。
- 不要な `console.log` ステートメントを削除します。
- Try/catch ブロックを使用しないでください。

### <a name="read-workbook-data-outside-of-a-loop"></a>ループの外部でブックのデータを読み取る

ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーすることができます。 同じ通話を繰り返し作成するのではなく、可能な限りデータをローカルに保存する必要があります。 これは、ループを処理するときに特に当てはまります。

ワークシートの使用されている範囲の負の数のカウントを取得するスクリプトについて検討します。 スクリプトは、使用されている範囲内のすべてのセルを反復処理する必要があります。 そのためには、範囲、行の数、および列の数が必要です。 ループを開始する前に、これらをローカル変数として格納する必要があります。 それ以外の場合、ループが反復処理されるたびに、ブックが強制的に返されます。

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> 実験として、でループを置き換えてみてください `usedRangeValues` `usedRange.getValues()` 。 大きな範囲を扱うときに、スクリプトの実行にかかる時間が非常に長くなることがあります。

### <a name="remove-unnecessary-consolelog-statements"></a>不要なステートメントを削除する `console.log`

コンソールログは、 [スクリプトをデバッグ](../testing/troubleshooting.md)するための非常に重要なツールです。 ただし、ログに記録された情報が最新であることを確認するために、スクリプトは強制的にブックと同期されます。 スクリプトを共有する前に、不要なログ記録ステートメント (テストに使用されるものなど) を削除することを検討してください。 これにより、 `console.log()` ステートメントがループ内にない限り、通常、パフォーマンスの問題が発生することはありません。

### <a name="avoid-using-trycatch-blocks"></a>Try/catch ブロックの使用を避ける

スクリプトの予想される制御フローの一部として[ `try` / `catch` ブロック](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)を使用することはお勧めしません。 ほとんどのエラーは、ブックから返されたオブジェクトをチェックすることで回避できます。 たとえば、次のスクリプトは、ブックから返されたテーブルが存在することを確認してから、行を追加します。

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>大文字と小文字を区別するヘルプ

Office スクリプトプラットフォームが [パワー自動化](https://flow.microsoft.com/)、 [アダプティブカード](https://docs.microsoft.com/adaptive-cards)、その他の製品間の機能に拡張されると、スクリプトブックの通信の詳細がより複雑になります。 スクリプトの実行速度を速くする必要がある場合は、 [スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-scripts)を参照してください。 専門家が検索してヘルプを見つけられるように、質問に "office スクリプト" というタグを付けてください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [MDN web ドキュメント: ループと反復](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
