---
title: スクリプトのパフォーマンスをOfficeする
description: ブックとスクリプトの間の通信を理解Excelスクリプトを作成します。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2deb417d41c4be663efaf83735459eab26146410
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585633"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>スクリプトのパフォーマンスをOfficeする

スクリプトの目的Office、一般的に実行される一連のタスクを自動化して時間を節約します。 遅いスクリプトは、ワークフローの速度を上げないような気がします。 ほとんどの場合、スクリプトは完全に正常に実行され、期待通り実行されます。 ただし、パフォーマンスに影響を与える可能性がある回避可能なシナリオがいくつかある。

遅いスクリプトの最も一般的な理由は、ブックとの過剰な通信です。 スクリプトはローカル コンピューター上で実行され、ブックはクラウドに存在します。 特定の時間に、スクリプトはローカル データをブックのローカル データと同期します。 つまり、書き込み操作 ( `workbook.addWorksheet()`など) は、この舞台裏の同期が発生した場合にのみブックに適用されます。 同様に、読み取り操作 ( `myRange.getValues()`など) は、その時点でスクリプトのブックからのデータのみを取得します。 どちらの場合も、スクリプトはデータに対して動作する前に情報をフェッチします。 たとえば、次のコードでは、使用範囲内の行数を正確に記録します。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Officeスクリプト API を使用すると、ブックまたはスクリプト内のデータが正確で、必要に応じて最新の情報になります。 スクリプトが正しく実行されるように、これらの同期について心配する必要はありません。 ただし、このスクリプト間通信を認識すると、不要なネットワーク呼び出しを回避できます。

## <a name="performance-optimizations"></a>パフォーマンスの最適化

クラウドへの通信を減らすのに役立つ簡単な手法を適用できます。 次のパターンは、スクリプトの高速化に役立ちます。

- ループ内で繰り返しではなく、ブック データを 1 回読み取ります。
- 不要なステートメントを `console.log` 削除します。
- try/catch ブロックを使用しないようにします。

### <a name="read-workbook-data-outside-of-a-loop"></a>ループ外のブック データの読み取り

ブックからデータを取得するメソッドは、ネットワーク呼び出しをトリガーできます。 同じ呼び出しを繰り返し行うのではなく、可能な限りローカルにデータを保存する必要があります。 これは、ループを扱う場合に特に当てはまる。

ワークシートの使用範囲の負の数値の数を取得するスクリプトを検討します。 スクリプトでは、使用範囲内のすべてのセルを反復処理する必要があります。 これを行うには、範囲、行数、列数が必要です。 ループを開始する前に、ローカル変数として格納する必要があります。 それ以外の場合、ループの繰り返しごとに強制的にブックに戻ります。

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
> 実験として、ループ内の置き換 `usedRangeValues` えをお試しください `usedRange.getValues()`。 大きな範囲を扱う場合、スクリプトの実行にかなり時間がかかる場合があります。

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>ループ内または周囲 `try...catch` のループでブロックを使用しないようにする

ループまたは周囲のループでステートメント [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) を使用することをお勧めしません。 これは、ループ内のデータの読み取りを避けるのと同じ理由です。各反復では、スクリプトは強制的にブックと同期し、エラーがスローされていないか確認します。 ほとんどのエラーは、ブックから返されるオブジェクトをチェックすることで回避できます。 たとえば、次のスクリプトは、行を追加する前に、ブックによって返されるテーブルが存在することを確認します。

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

### <a name="remove-unnecessary-consolelog-statements"></a>不要なステートメントを `console.log` 削除する

コンソール ログは、スクリプトをデバッグ [するための重要なツールです](../testing/troubleshooting.md)。 ただし、スクリプトはブックと強制的に同期し、ログに記録された情報が最新の状態に更新されます。 スクリプトを共有する前に、不要なログ ステートメント (テストに使用されるログ ステートメントなど) を削除してください。 これは通常、ステートメントがループに `console.log()` 含まれる場合を限り、パフォーマンスの問題を引き起こします。

## <a name="case-by-case-help"></a>ケース バイ ケース のヘルプ

Officeスクリプト プラットフォームが[拡張され、Power Automate](https://flow.microsoft.com/)、[アダプティブ](/adaptive-cards) カード、その他のクロスプロダクト機能を使用して動作する場合、スクリプト ブックの通信の詳細が複雑になります。 スクリプトの実行速度を上げ支援が必要な場合は、 [Microsoft Q&してください](/answers/topics/office-scripts-excel-dev.html)。 専門家が質問を見つけて支援できるよう、質問に "office-scripts-dev" をタグ付けしてください。

## <a name="see-also"></a>関連項目

- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [MDN Web ドキュメント: ループと反復](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
