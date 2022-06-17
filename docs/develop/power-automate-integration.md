---
title: Power Automate を使用した Office スクリプトの実行
description: Power Automate ワークフローを使用して Excel on the web の Office スクリプトを取得する方法。
ms.date: 05/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85c335eeb736ec544eccb2fbdbe819bdbef6848c
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128231"
---
# <a name="run-office-scripts-with-power-automate"></a>Power Automate を使用した Office スクリプトの実行

[Power Automate](https://flow.microsoft.com) では、より大規模で自動化されたワークフローに Office スクリプトを追加できます。Power Automate を使用すると、ワークシートのテーブルにメールの内容を追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりできます。

## <a name="get-started"></a>概要

Power Automate を初めて使用する場合は、「[Power Automate に関する入門情報](/power-automate/getting-started)」にアクセスすることをお勧めします。 そちらで、利用可能なすべてのオートメーションの可能性について詳しく学ぶことができます。 このドキュメントでは、Power Automate での Office スクリプトの動作と、それが Excel エクスペリエンスの改善にどのように役立つかに重点が置かれています。

Power Automate と Office スクリプトの統合を開始するには、チュートリアル「[Power Automate でスクリプトの使用を開始する](../tutorials/excel-power-automate-manual.md)」に従ってください。 単純なスクリプトを呼び出すフローの作成方法について学ぶことができます。 このチュートリアルと「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)」のチュートリアルが完了したら、こちらに戻り、Office スクリプトを Power Automate フローに接続する方法の詳細をご確認ください。

## <a name="excel-online-business-connector"></a>Excel Online (Business) コネクタ

[コネクタ](/connectors/connectors)は、Power Automate とアプリケーション間のブリッジです。 [Excel Online (Business) コネクタ](/connectors/excelonlinebusiness)を使用すると、フローに Excel ブックへのアクセスが提供されます。 "スクリプトの実行" アクションにより、選択したブックからアクセスできるすべての Office スクリプトを呼び出すことができます。 また、フローによってデータを提供したり、フローの後の手順用にスクリプトで情報を返したりできるよう、スクリプトに入力パラメーターを指定することもできます。

> [!IMPORTANT]
> "スクリプトの実行" アクションにより、Excel コネクタを使用するユーザーにブックとそのデータへの重要なアクセス権が付与されます。 さらに、「[Power Automate からの外部呼び出し](external-calls.md)」で説明されているとおり、外部 API の呼び出しを行うスクリプトにセキュリティ上のリスクがあります。 管理者が機密性の高いデータの流出を懸念している場合は、Excel Online コネクタをオフにするか、[Office スクリプト管理者制御](/microsoft-365/admin/manage/manage-office-scripts-settings)で Office スクリプトへのアクセスを制限することができます。

> [!IMPORTANT]
> Power Automateでは、現時点ではSharePointに格納されているスクリプトはサポート **されていません**。

## <a name="data-transfer-in-flows-for-scripts"></a>スクリプトのフローでのデータ転送

Power Automate を使用すると、フローのステップ間でデータの一部を渡すことができます。 スクリプトを構成して、必要な種類の情報を受け入れたり、フローに必要なものをブックから返したりすることができます。 スクリプトの入力は、(`workbook: ExcelScript.Workbook` に加えて) `main` 関数にパラメーターを追加することによって指定されます。 スクリプトからの出力は、`main` に戻り値の型を追加することによって宣言されます。

> [!NOTE]
> フローで "スクリプトの実行" ブロックを作成すると、承認されたパラメーターと返された型が入力されます。 スクリプトのパラメーターまたは戻り値の型を変更する場合は、フローの "スクリプトの実行" ブロックを再実行する必要があります。 これにより、データが正しく解析されていることが確認されます。

次のセクションでは、Power Automate で使用されるスクリプトの入力と出力の詳細について説明します。 このトピックについて学ぶための実践的なアプローチが必要な場合は、「[自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)」チュートリアルを試すか、[タスクの自動アラーム](../resources/scenarios/task-reminders.md)のサンプル シナリオを確認してください。

### <a name="main-parameters-pass-data-to-a-script"></a>`main` パラメーター: スクリプトにデータを渡す

すべてのスクリプト入力は、`main` 関数の追加パラメーターとして指定されます。 たとえば、入力として名前を表す `string` をスクリプトで受け入れるようにする場合は、`main` 署名を `function main(workbook: ExcelScript.Workbook, name: string)` に変更します。

Power Automate でフローを構成する場合、スクリプト入力を静的な値、[式](/power-automate/use-expressions-in-conditions)、または動的なコンテンツとして指定できます。 個々のサービスのコネクタの詳細については、[Power Automate コネクタに関するドキュメント](/connectors/)を参照してください。

#### <a name="type-restrictions"></a>入力の制限

スクリプトの `main` 関数に入力パラメーターを追加する場合は、次の上限や制限を検討してください。 これらは、スクリプトの戻り値の型にも適用されます。

1. 最後のパラメーターは `ExcelScript.Workbook` の型にする必要があります。 パラメーター名は関係ありません。

1. 型 `string`、`number`、`boolean`、`unknown`、`object`、および `undefined` がサポートされています。

1. 前述の型の配列 (`[]` と `Array<T>` スタイルの両方) がサポートされています。 入れ子になった配列もサポートされています。

1. 共用体型は、単一の型に属するリテラルの共用体 (`"Left", 5` ではなく、`"Left" | "Right"` など) の場合に許可されます。 undefined を含むサポートされる型の共用体 (`string | undefined` など) もサポートされます。

1. オブジェクト型は、型 `string`、`number`、`boolean`、サポートされている配列、または他のサポートされているオブジェクトのプロパティが含まれる場合に許可されます。 次の例は、パラメーターの型としてサポートされる入れ子にされたオブジェクトを示しています。

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. オブジェクトのインターフェイスまたはクラス定義はスクリプトで定義されている必要があります。 次の例のように、オブジェクトをインラインで匿名で定義することができます。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>省略可能なパラメーターと既定のパラメーター

1. 省略可能なパラメーターは許可され、省略可能な修飾子 `?` で示されます (例: `function main(workbook: ExcelScript.Workbook, Name?: string)`)。

1. 既定のパラメーター値は許可されています (例: `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`)。

### <a name="return-data-from-a-script"></a>スクリプトからデータを返す

スクリプトではブックからデータを返すことができ、Power Automate フローの動的なコンテンツとして使用することができます。 [前述の型制限](#type-restrictions) が戻り値の型に適用されます。 オブジェクトを返すには、戻り値の型構文を `main` 関数に追加します。 たとえば、スクリプトから `string` 値を返す場合、`main` シグネチャは `function main(workbook: ExcelScript.Workbook): string` になります。

## <a name="example"></a>例

次のスクリーンショットは、[GitHub](https://github.com/) の問題がお客様に割り当てられるたびにトリガーされる Power Automate フローを示しています。 このフローでは、Excel ブックのテーブルに問題を追加するスクリプトが実行されます。 そのテーブルに 5 つ以上の問題がある場合、フローでメール アラームが送信されます。

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="サンプルのフローを示す Power Automate フロー エディター。":::

スクリプトの `main` 関数では、問題の ID と問題のタイトルが入力パラメーターとして指定され、スクリプトによって問題テーブルの行数が返されます。

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>関連項目

- [手動 Power Automation フローからスクリプトを呼び出す](../tutorials/excel-power-automate-manual.md)
- [自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)
- [自動で実行される Power Automate フローにスクリプトからデータを返す](../tutorials/excel-power-automate-returns.md)
- [Office スクリプトを使用した Power Automate のトラブルシューティング情報](../testing/power-automate-troubleshooting.md)
- [Power Automate の使用を開始する](/power-automate/getting-started)
- [Excel Online (Business) コネクタ リファレンス ドキュメント](/connectors/excelonlinebusiness/)
