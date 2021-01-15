---
title: スクリプトのパフォーマンスをOfficeする
description: Excel ブックとスクリプトの間の通信を理解して、より高速なスクリプトを作成します。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867871"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>スクリプトのパフォーマンスをOfficeする

スクリプトを使用Office、一般的に実行される一連のタスクを自動化して時間を節約します。 遅いスクリプトは、ワークフローの速度を上げないような気がします。 ほとんどの場合、スクリプトは完全に正常に実行され、期待通り実行されます。 ただし、パフォーマンスに影響を与える可能性がある、いくつかの回避可能なシナリオがあります。

スクリプトの速度が遅い最も一般的な理由は、ブックとの通信が過剰である点です。 スクリプトはローカル コンピューターで実行され、ブックはクラウド内に存在します。 スクリプトによって、ローカル データがブックのローカル データと同期される場合があります。 つまり、書き込み操作 (など) がブックに適用されるのは、この背後での同期が発生した `workbook.addWorksheet()` 場合のみです。 同様に、読み取り操作 (たとえば) は、その時点でスクリプトのブックからのデータ `myRange.getValues()` のみを取得します。 どちらの場合も、スクリプトはデータに対して動作する前に情報をフェッチします。 たとえば、次のコードは、使用範囲内の行数を正確に記録します。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office スクリプト API は、ブックまたはスクリプト内のデータが正確で、必要に応じて最新の情報を提供します。 スクリプトを正しく実行するために、これらの同期について心配する必要はありません。 ただし、このスクリプトからクラウドへの通信を認識すると、不要なネットワーク呼び出しを回避するのに役立ちます。

## <a name="performance-optimizations"></a>パフォーマンスの最適化

クラウドへの通信を減らすのに役立つ簡単な手法を適用できます。 次のパターンは、スクリプトの高速化に役立ちます。

- ループ内で繰り返しではなく、ブックのデータを 1 回読み取ります。
- 不要なステートメント `console.log` を削除します。
- try/catch ブロックは使用しないようにします。

### <a name="read-workbook-data-outside-of-a-loop"></a>ループ外のブック データの読み取り

ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーできます。 同じ呼び出しを繰り返し行うのではなく、できる限りローカルにデータを保存する必要があります。 これは、ループを処理する場合に特に当てはまる場合です。

ワークシートの使用範囲の負の数を取得するスクリプトを検討します。 スクリプトは、使用範囲内のすべてのセルに対して反復処理を行う必要があります。 これを行うには、範囲、行数、列数が必要です。 ループを開始する前に、ローカル変数として格納する必要があります。 それ以外の場合、ループを繰り返すごとに強制的にブックに戻ります。

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
> 実験として、ループ内で置き換 `usedRangeValues` えしてみてください `usedRange.getValues()` 。 大きな範囲を処理する場合、スクリプトの実行にかなり時間がかかる場合があります。

### <a name="remove-unnecessary-consolelog-statements"></a>不要なステートメント `console.log` を削除する

コンソール ログは、スクリプトをデバッグ [するための重要なツールです](../testing/troubleshooting.md)。 ただし、ログに記録された情報を最新の状態にするためのスクリプトを強制的にブックと同期します。 スクリプトを共有する前に、不要なログ 記録ステートメント (テスト用など) を削除してください。 これは通常、ステートメントがループ内にある場合を限り、パフォーマンスの問題 `console.log()` を顕著に引き起こします。

### <a name="avoid-using-trycatch-blocks"></a>try/catch ブロックの使用を避ける

スクリプトの予想される制御フロー[ `try` / `catch` の](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)一部としてブロックを使用はお勧めしません。 ブックから返されたオブジェクトをチェックすることで、ほとんどのエラーを回避できます。 たとえば、次のスクリプトは、行を追加する前に、ブックから返されたテーブルが存在するかどうかを確認します。

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

## <a name="case-by-case-help"></a>ケース バイ ケースのヘルプ

Office スクリプト プラットフォームが[拡張され、Power Automate、](https://flow.microsoft.com/)[アダプティブ](/adaptive-cards)カード、その他の製品間の機能で動作するほど、スクリプト ブックの通信の詳細が複雑になります。 スクリプトの実行速度を速くするためにヘルプが必要な場合は [、Stack Overflow からご確認ください](https://stackoverflow.com/questions/tagged/office-scripts)。 質問に必ず 「office-scripts」というタグを付け、専門家が質問を見つけて支援します。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [MDN Web ドキュメント: ループと反復](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)