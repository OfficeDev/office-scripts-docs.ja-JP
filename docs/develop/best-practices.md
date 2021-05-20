---
title: Office スクリプトのベスト プラクティス
description: 一般的な問題を回避し、予期しない入力やデータを処理できる堅牢なOfficeスクリプトを記述する方法。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546032"
---
# <a name="best-practices-in-office-scripts"></a>Office スクリプトのベスト プラクティス

これらのパターンとプラクティスは、スクリプトが毎回正常に実行されるように設計されています。 Excelワークフローの自動化を開始する際に、一般的な落とし穴を避けるために使用します。

## <a name="verify-an-object-is-present"></a>オブジェクトが存在することを確認する

スクリプトは、ブック内に存在する特定のワークシートまたはテーブルに依存することがよくあります。 ただし、スクリプトの実行の間に名前が変更されたり削除されたりする場合があります。 メソッドを呼び出す前に、これらのテーブルまたはワークシートが存在するかどうかを確認することで、スクリプトが突然終了しないようにすることができます。

次のサンプル コードは、"インデックス" ワークシートがブックに存在しているかどうかを確認します。 ワークシートが存在する場合、スクリプトは範囲を取得して処理を続行します。 存在しない場合、スクリプトはカスタム エラー メッセージをログに記録します。

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

TypeScript `?` 演算子は、メソッドを呼び出す前にオブジェクトが存在するかどうかをチェックします。 これにより、オブジェクトが存在しないときに特別な操作を行う必要がない場合に、コードの合理化が可能になります。

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>データとブックの状態を最初に検証する

データを操作する前に、ワークシート、表、図形、およびその他のオブジェクトがすべて存在することを確認します。 前のパターンを使用して、すべてがワークブック内にあり、期待に合っているかどうかを確認します。 データが書き込まれる前にこれを行うと、スクリプトがワークブックを部分的な状態に置き去りにすることがなくなります。

次のスクリプトでは、"Table1" と "Table2" という名前のテーブルが 2 つ存在する必要があります。 スクリプトは、まずテーブルが存在するかどうか確認し、ステートメント `return` と適切でない場合は適切なメッセージで終わります。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

検証が別の関数で行われている場合でも、 `return` その関数からステートメントを発行してスクリプトを終了する必要があります `main` 。 サブ関数から戻っても、スクリプトは終了しません。

次のスクリプトは、前のスクリプトと同じ動作をします。 違いは、 `main` 関数が関数を呼び出 `inputPresent` してすべてを検証することです。 `inputPresent``true` `false` は、必要な入力がすべて存在するかどうかを示すブール値 ( または ) を返します。 `main`この関数は、そのブール値を使用して、スクリプトの続行または終了を決定します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a>ステートメントを使用する場合 `throw`

[`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw)ステートメントは、予期しないエラーが発生したことを示します。 コードはすぐに終了します。 ほとんどの場合、スクリプトから実行する必要はありません `throw` 。 通常、スクリプトは、問題が原因でスクリプトが実行できなかったことを自動的にユーザーに通知します。 ほとんどの場合、エラー メッセージと関数のステートメントでスクリプトを終了するだけで十分です `return` `main` 。

ただし、スクリプトがPower Automateフローの一部として実行されている場合は、フローの続行を停止する必要があります。 `throw`ステートメントはスクリプトを停止し、フローにも停止するように指示します。

次のスクリプトは、 `throw` テーブル チェックの例でステートメントを使用する方法を示しています。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a>ステートメントを使用する場合 `try...catch`

[`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)このステートメントは、API 呼び出しが失敗したかどうかを検出し、スクリプトの実行を続行する方法です。

範囲で大規模なデータ更新を実行する次のスニペットを検討してください。

```TypeScript
range.setValues(someLargeValues);
```

`someLargeValues`Web で処理できるExcelより大きい場合、 `setValues()` 呼び出しは失敗します。 スクリプトは [、ランタイム エラー](../testing/troubleshooting.md#runtime-errors)で失敗します。 この `try...catch` ステートメントを使用すると、スクリプトを直ちに終了して既定のエラーを表示することなく、スクリプトがこの状態を認識できます。

スクリプト ユーザーに、より優れたエクスペリエンスを提供する方法の 1 つは、カスタム エラー メッセージを提示することです。 次のスニペットは、 `try...catch` より多くのエラー情報を記録して、読者を助けるステートメントを示しています。

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

エラーを処理するもう 1 つの方法は、エラーのケースを処理するフォールバック動作を持つということです。 次のスニペットでは `catch` 、ブロックを使用して、更新を小さく分割してエラーを回避する別の方法を試します。

> [!TIP]
> 大きな範囲を更新する方法の完全な例については、「 [大きなデータセットを記述する](../resources/samples/write-large-dataset.md)」を参照してください。

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> `try...catch`ループの内側または周囲を使用すると、スクリプトの速度が低下します。 パフォーマンスの詳細については、 [ `try...catch` ブロックの使用を避ける](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)を参照してください。

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトを使用した Power Automate のトラブルシューティング情報](../testing/power-automate-troubleshooting.md)
- [Officeスクリプトを使用したプラットフォームの制限](../testing/platform-limits.md)
- [Officeスクリプトのパフォーマンスを向上させる](web-client-performance.md)
