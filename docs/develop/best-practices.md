---
title: Office スクリプトでのベスト プラクティス
description: 一般的な問題を防止し、予期しない入力またはデータOfficeできる堅牢なスクリプトを記述する方法。
ms.date: 12/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 689196e1a0ca70c999ec8048de64190cbfe75581
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585766"
---
# <a name="best-practices-in-office-scripts"></a>Office スクリプトでのベスト プラクティス

これらのパターンとプラクティスは、スクリプトが毎回正常に実行されるのを助けるために設計されています。 ワークフローの自動化を開始する場合、一般的な落とし穴をExcelします。

## <a name="use-the-action-recorder-to-learn-new-features"></a>アクション レコーダーを使用して新機能を学習する

Excel多くのことを行います。 ほとんどのスクリプトはスクリプト化できます。 アクション レコーダーは、ユーザーのExcelを記録し、コードに変換します。 これは、さまざまな機能がスクリプトを使用してどのように動作Office方法です。 特定のアクションのコードが必要な場合は、アクション レコーダーに切り替えてアクションを実行し、[コードとしてコピー] を選択して、結果のコードをスクリプトに貼り付けます。

:::image type="content" source="../images/action-recorder-copy-code.png" alt-text="[コードとしてコピー] ボタンが強調表示された、操作レコーダー作業ウィンドウ。":::

## <a name="verify-an-object-is-present"></a>オブジェクトが存在する確認

スクリプトは、多くの場合、ブックに存在する特定のワークシートまたはテーブルに依存します。 ただし、スクリプトの実行の間に名前が変更または削除される場合があります。 これらのテーブルまたはワークシートがメソッドを呼び出す前に存在する場合は、スクリプトが突然終了しないか確認できます。

次のサンプル コードは、ブックに "Index" ワークシートが存在する場合にチェックします。 ワークシートが存在する場合、スクリプトは範囲を取得して続行します。 存在しない場合、スクリプトはカスタム エラー メッセージをログに記録します。

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

TypeScript 演算子は `?` 、メソッドを呼び出す前にオブジェクトが存在するかどうかをチェックします。 これにより、オブジェクトが存在しない場合に特別な操作を行う必要が生じなかった場合に、コードの効率化が図れる可能性があります。

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>データとブックの状態を最初に検証する

データを操作する前に、すべてのワークシート、テーブル、図形、その他のオブジェクトが存在する必要があります。 前のパターンを使用して、すべてがブック内にあるか確認し、期待に一致します。 データが書き込まれる前にこれを行って、スクリプトがブックを部分的な状態に残すのを確認します。

次のスクリプトでは、"Table1" と "Table2" という名前の 2 つのテーブルが存在する必要があります。 スクリプトは、最初にテーブルが `return` 存在する場合はチェックし、ステートメントで終わり、存在しない場合は適切なメッセージで終わります。

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

  // Continue...
}
```

検証が別の関数で `return` 行っている場合でも、関数からステートメントを発行してスクリプトを終了する必要 `main` があります。 サブ関数から戻しても、スクリプトは終了しない。

次のスクリプトは、前のスクリプトと同じ動作をします。 違いは、関数が関数 `main` を呼び出してすべてを `inputPresent` 検証する点です。 `inputPresent`必要なすべての入力が存在するかどうかを示`false`すブール値 (`true` または ) を返します。 関数 `main` は、そのブール値を使用して、スクリプトの継続または終了を決定します。

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>ステートメントを使用する `throw` 場合

ステートメント [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) は、予期しないエラーが発生したかどうかを示します。 コードは直ちに終了します。 ほとんどの場合、スクリプトから実行する `throw` 必要はなんらない。 通常、スクリプトは、問題が原因でスクリプトの実行に失敗したとユーザーに自動的に通知します。 ほとんどの場合、エラー メッセージ `return` と関数のステートメントでスクリプトを終了しても十分 `main` です。

ただし、スクリプトがプロセス フローの一部として実行Power Automate、フローの続行を停止できます。 ステートメント `throw` はスクリプトを停止し、フローにも停止を指示します。

次のスクリプトは、テーブルチェックの例で `throw` ステートメントを使用する方法を示しています。

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

## <a name="when-to-use-a-trycatch-statement"></a>ステートメントを使用する `try...catch` 場合

この [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) ステートメントは、API 呼び出しが失敗した場合に検出し、スクリプトの実行を続行する方法です。

範囲に対して大規模なデータ更新を実行する次のスニペットを検討してください。

```TypeScript
range.setValues(someLargeValues);
```

処理`someLargeValues`できる値よりExcel for the web場合、呼び出しは`setValues()`失敗します。 スクリプトは、実行時エラーでも [失敗します](../testing/troubleshooting.md#runtime-errors)。 この `try...catch` ステートメントを使用すると、スクリプトをすぐに終了して既定のエラーを表示することなく、スクリプトでこの条件を認識できます。

スクリプト ユーザーに優れたエクスペリエンスを提供する方法の 1 つは、カスタム エラー メッセージを表示する方法です。 次のスニペットは、読者に役立 `try...catch` つエラー情報をログに記録するステートメントを示しています。

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

エラーを処理するもう 1 つの方法は、エラー ケースを処理するフォールバック動作を持つ方法です。 次のスニペットでは、ブロックを `catch` 使用して別のメソッドを試して、更新プログラムを小さな部分に分割し、エラーを回避します。

> [!TIP]
> 大きな範囲を更新する方法の完全な例については、「大きなデータセットを記述 [する」を参照してください](../resources/samples/write-large-dataset.md)。

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
> ループの `try...catch` 内側または周囲を使用すると、スクリプトの速度が低下します。 パフォーマンスの詳細については、「ブロックの使用 [を避ける」を参照 `try...catch` してください](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)。

## <a name="see-also"></a>関連項目

- [Office スクリプトのトラブルシューティング](../testing/troubleshooting.md)
- [Office スクリプトを使用した Power Automate のトラブルシューティング情報](../testing/power-automate-troubleshooting.md)
- [スクリプトを使用したプラットフォームOffice制限](../testing/platform-limits.md)
- [スクリプトのパフォーマンスをOfficeする](web-client-performance.md)
