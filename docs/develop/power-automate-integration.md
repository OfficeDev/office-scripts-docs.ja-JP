---
title: Power Automate Officeスクリプトを実行する
description: Power Automate ワークフロー Office操作する Web 上の Excel スクリプトを取得する方法。
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755071"
---
# <a name="run-office-scripts-with-power-automate"></a>Power Automate Officeスクリプトを実行する

[Power Automate を](https://flow.microsoft.com) 使用すると、Officeスクリプトを大規模で自動化されたワークフローに追加できます。 Power Automate を使用すると、ワークシートのテーブルに電子メールの内容を追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりすることができます。

## <a name="getting-started"></a>はじめに

Power Automate を使用する場合は、「Power Automate の使用を開始 [する」にアクセスすることをお勧めします](/power-automate/getting-started)。 そこで、利用可能なすべてのオートメーションの可能性について詳しくは、ご覧ください。 このドキュメントでは、Power Automate Officeスクリプトがどのように動作し、Excel エクスペリエンスを向上させるのに役立つのかについて重点的に取り上っています。

Power Automate スクリプトと Officeスクリプトの組み合Office、Power Automate を使用したスクリプトの [使用を開始するチュートリアルに従います](../tutorials/excel-power-automate-manual.md)。 これにより、単純なスクリプトを呼び出すフローを作成する方法を説明します。 このチュートリアルと、自動的に実行される [Power Automate](../tutorials/excel-power-automate-trigger.md) フロー チュートリアルのスクリプトへのデータの渡しを完了したら、Office スクリプトを Power Automate フローに接続する方法の詳細については、ここに戻します。

## <a name="excel-online-business-connector"></a>Excel Online (Business) コネクタ

[コネクタは](/connectors/connectors) 、Power Automate とアプリケーションの間のブリッジです。 [Excel Online (Business) コネクタを使用すると](/connectors/excelonlinebusiness)、フローから Excel ブックにアクセスできます。 "スクリプトの実行" アクションでは、選択したブックからアクセスOfficeスクリプトを呼び出します。 また、フローによってデータを提供できるようスクリプトにパラメーターを入力したり、フロー内の後の手順に関する情報をスクリプトから返したりすることもできます。

> [!IMPORTANT]
> "スクリプトの実行" アクションにより、Excel コネクタを使用するユーザーはブックとそのデータに重要なアクセス権を与えます。 さらに [、「Power Automate](external-calls.md)からの外部呼び出し」で説明したように、外部 API 呼び出しを行うスクリプトにはセキュリティリスクがあります。 管理者が機密性の高いデータの露出に関心がある場合は、Excel Online コネクタをオフにするか、Office スクリプト管理者コントロールを使用して Office スクリプトへのアクセス [を制限できます](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="data-transfer-in-flows-for-scripts"></a>スクリプトのフローでのデータ転送

Power Automate を使用すると、フローのステップ間でデータを渡します。 必要な情報の種類を受け入れ、フローで必要な情報をブックから返すスクリプトを構成できます。 スクリプトの入力は、(に加えて) 関数にパラメーター `main` を追加することで指定されます `workbook: ExcelScript.Workbook` 。 スクリプトからの出力は、 に戻り値の型を追加することで宣言されます `main` 。

> [!NOTE]
> フローで "スクリプトの実行" ブロックを作成すると、受け入れられるパラメーターと返される型が設定されます。 スクリプトのパラメーターまたは戻り値の種類を変更する場合は、フローの "スクリプトの実行" ブロックをやり直す必要があります。 これにより、データが正しく解析されます。

次のセクションでは、Power Automate で使用されるスクリプトの入力と出力の詳細について説明します。 このトピックの学習に関する実践的なアプローチが必要な場合は、自動実行の[Power Automate](../tutorials/excel-power-automate-trigger.md)フローチュートリアルでスクリプトにデータを渡すチュートリアル[](../resources/scenarios/task-reminders.md)を試してみるか、自動タスク リマインダーのサンプル シナリオを参照してください。

### <a name="main-parameters-passing-data-to-a-script"></a>`main` パラメーター: スクリプトにデータを渡す

すべてのスクリプト入力は、関数の追加パラメーターとして指定 `main` されます。 たとえば、名前を入力として表すスクリプトを受け入れる場合は、署名 `string` を `main` に変更します `function main(workbook: ExcelScript.Workbook, name: string)` 。

Power Automate でフローを構成する場合は、スクリプト入力を静的な値、式、または動的 [コンテンツとして指定](/power-automate/use-expressions-in-conditions)できます。 個々のサービスのコネクタの詳細については [、「Power Automate Connector」のドキュメントを参照してください](/connectors/)。

スクリプトの関数に入力パラメーターを追加する場合は、次の許容値と `main` 制限を考慮してください。

1. 最初のパラメーターは型である必要があります `ExcelScript.Workbook` 。 パラメーター名は関係ありません。

2. すべてのパラメーターには、型 (または `string` など) が必要 `number` です。

3. 基本的な型 `string` `number` `boolean` `any` 、、、、 `unknown` `object` `undefined` がサポートされています。

4. 前に示した基本型の配列がサポートされています。

5. 入れ子になった配列はパラメーターとしてサポートされます (ただし、戻り値の型としてサポートされません)。

6. 共用体の型は、単一の型 (など) に属するリテラルの共用体である場合に使用できます `"Left" | "Right"` 。 未定義のサポートされている型の共用体もサポートされています (など `string | undefined` )。

7. オブジェクト型は、型、、サポートされている配列、または他のサポートされているオブジェクトのプロパティが含まれている `string` `number` `boolean` 場合に使用できます。 次の例は、パラメーターの種類としてサポートされている入れ子になったオブジェクトを示しています。

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. オブジェクトには、スクリプトで定義されているインターフェイスまたはクラス定義が必要です。 次の例のように、オブジェクトをインラインで匿名で定義できます。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. 省略可能なパラメーターは許可され、省略可能な修飾子 (たとえば) を使用して `?` 指定できます `function main(workbook: ExcelScript.Workbook, Name?: string)` 。

10. 既定のパラメーター値を使用できます (たとえば `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

### <a name="returning-data-from-a-script"></a>スクリプトからデータを返す

スクリプトは、Power Automate フローで動的コンテンツとして使用するブックからデータを返します。 入力パラメーターと同様に、Power Automate は戻り値の種類にいくつかの制限を設定します。

1. 基本の型 `string` `number` `boolean` `void` 、、、 `undefined` がサポートされています。

2. 戻り値の型として使用される Union 型は、スクリプト パラメーターとして使用する場合と同じ制限に従います。

3. 配列型は、型 、、または `string` の場合 `number` に使用できます `boolean` 。 また、この型がサポートされている共用体またはサポートされているリテラル型である場合にも使用できます。

4. 戻り値の型として使用されるオブジェクトの種類は、スクリプト パラメーターとして使用する場合と同じ制限に従います。

5. 暗黙的な型指定はサポートされています。定義された型と同じルールに従う必要があります。

## <a name="example"></a>例

次のスクリーンショットは、GitHub の問題が割り当てられるたびにトリガーされる [Power Automate](https://github.com/) フローを示しています。 フローは、Excel ブック内のテーブルに問題を追加するスクリプトを実行します。 そのテーブルに 5 つ以上の問題がある場合、フローはメールリマインダーを送信します。

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="例のフローを示す Power Automate フロー エディター。":::

スクリプトの関数は、問題 ID と発行タイトルを入力パラメーターとして指定し、スクリプトは問題テーブル内の行数 `main` を返します。

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

- [Power Automate Office Excel で Web 上のスクリプトを実行する](../tutorials/excel-power-automate-manual.md)
- [自動で実行される Power Automate フロー内で、データをスクリプトに渡す](../tutorials/excel-power-automate-trigger.md)
- [自動で実行される Power Automate フローにスクリプトからデータを返す](../tutorials/excel-power-automate-returns.md)
- [Power Automate with Office スクリプト](../testing/power-automate-troubleshooting.md)
- [Power Automate の使用を開始する](/power-automate/getting-started)
- [Excel Online (Business) コネクタリファレンス ドキュメント](/connectors/excelonlinebusiness/)
