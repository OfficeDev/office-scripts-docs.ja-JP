---
title: Officeスクリプトのパフォーマンスを向上させる
description: Excelブックとスクリプトの間の通信を理解して、より高速なスクリプトを作成します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544992"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Officeスクリプトのパフォーマンスを向上させる

Officeスクリプトの目的は、一般的に実行される一連のタスクを自動化して時間を節約することです。 低速なスクリプトは、ワークフローを高速化できないような感じがします。 ほとんどの場合、スクリプトは完全に問題なく、期待どおりに実行されます。 ただし、パフォーマンスに影響を与える可能性のある回避可能なシナリオがいくつかあります。

スクリプトが遅い最も一般的な理由は、ブックとの通信が過剰である場合です。 スクリプトはローカル コンピューターで実行されますが、ブックはクラウドに存在します。 スクリプトは、特定の時間に、ローカル データとブックのデータを同期します。 つまり、書き込み操作 ( など `workbook.addWorksheet()` ) は、このバックグラウンド同期が行われる場合にのみブックに適用されます。 同様に、読み取り操作 ( など `myRange.getValues()` ) は、その時点でスクリプトのワークブックからデータを取得するだけです。 いずれの場合も、スクリプトはデータに対して動作する前に情報をフェッチします。 たとえば、次のコードは、使用範囲内の行数を正確に記録します。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Officeスクリプト API は、ワークブックまたはスクリプト内のデータが正確で、必要に応じて最新であることを保証します。 スクリプトを正しく実行するために、これらの同期について心配する必要はありません。 ただし、このスクリプトからクラウドへの通信を認識すると、不要なネットワーク呼び出しを回避できます。

## <a name="performance-optimizations"></a>パフォーマンスの最適化

クラウドへの通信を減らすのに役立つ簡単な手法を適用できます。 次のパターンは、スクリプトの高速化に役立ちます。

- ループ内で繰り返しブック データを読み取るのではなく、ブックデータを 1 回読み取ります。
- 不要な `console.log` ステートメントを削除します。
- try/catch ブロックの使用は避けてください。

### <a name="read-workbook-data-outside-of-a-loop"></a>ループの外側でブック データを読み取る

ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーできます。 同じ呼び出しを繰り返し行うのではなく、可能な限りデータをローカルに保存する必要があります。 これは、ループを扱う場合に特に当てはまります。

ワークシートの使用範囲内の負の数を取得するスクリプトを検討してください。 スクリプトは、使用範囲内のすべてのセルを反復処理する必要があります。 そのためには、範囲、行数、列数が必要です。 ループを開始する前に、ローカル変数として格納する必要があります。 そうしないと、ループの各反復処理によってブックに強制的に戻ります。

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
> 実験として、 `usedRangeValues` ループ内でを に置き換えて `usedRange.getValues()` みます。 大きな範囲を扱う場合、スクリプトの実行にかなり時間がかかる場合があります。

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>`try...catch`ループ内または周囲のループでブロックを使用しないようにする

[`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)ループまたは周囲のループでステートメントを使用することはお勧めしません。 これは、ループ内のデータの読み取りを避ける必要があるのと同じ理由です: 各反復処理では、スクリプトがブックと同期してエラーがスローされないように強制します。 ほとんどのエラーは、ブックから返されたオブジェクトをチェックすることで回避できます。 たとえば、次のスクリプトは、行を追加する前に、ブックによって返されたテーブルが存在することを確認します。

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

### <a name="remove-unnecessary-consolelog-statements"></a>不要な `console.log` ステートメントを削除する

コンソールログはスクリプトをデバッグするための重要なツール [です](../testing/troubleshooting.md)。 ただし、スクリプトは、記録された情報が最新であることを確認するために、ブックと同期する必要があります。 スクリプトを共有する前に、不要なログ記録ステートメント (テストに使用するものなど) を削除することを検討してください。 通常、ステートメントがループ内にある場合を除き、パフォーマンスに問題 `console.log()` はありません。

## <a name="case-by-case-help"></a>ケースバイケースのヘルプ

Officeスクリプト プラットフォームが[Power Automate、](https://flow.microsoft.com/)[アダプティブ カード](/adaptive-cards)、その他の製品間機能を使用するように拡張するにつれて、スクリプトとワークブックの通信の詳細がより複雑になります。 スクリプトの実行速度を上げるためのヘルプが必要な場合は、 [Microsoft Q&A](/answers/topics/office-scripts-dev.html)を通じて連絡を取ってください。 専門家がそれを見つけて助けることができるように、あなたの質問に「オフィススクリプト-dev」とタグを付けてください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [MDN Web ドキュメント: ループと反復](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
