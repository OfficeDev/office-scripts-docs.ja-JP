---
title: Office スクリプトとパワー自動化を統合する
description: Power 自動ワークフローを使用して、web 上の Excel で Office スクリプトを取得する方法について説明します。
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878849"
---
# <a name="integrate-office-scripts-with-power-automate"></a>Office スクリプトとパワー自動化を統合する

[パワー自動](https://flow.microsoft.com)では、スクリプトをより大きなワークフローに統合します。 Power オートメーションでは、ワークシートのテーブルに電子メールの内容を追加したり、ブックのコメントに基づいてプロジェクト管理ツールでアクションを作成したりするなどの操作を実行できます。 電力を自動自動化することが初めての場合は、「 [Power オートメーションの使用を開始](/power-automate/getting-started)する」を参照することをお勧めします。 ここでは、複数のサービス間でワークフローを自動化する方法について詳しく知ることができます。

> [!IMPORTANT]
> 現時点では、[共有フロー](/power-automate/share-buttons)から Office スクリプトを実行することはできません。 パワー自動化を使用しても、スクリプトを作成したユーザーのみが実行できます。

## <a name="getting-started"></a>はじめに

Power オートメーションと Office のスクリプトの組み合わせを開始するには、チュートリアルの次の手順を実行し[て、Power 自動化を使用したスクリプトの使用を開始](../tutorials/excel-power-automate-manual.md)します。 これにより、簡単なスクリプトを呼び出すフローを作成する方法を学習できます。 そのチュートリアルを完了し、Power オートメーションチュートリアルを使用して[自動的にスクリプトを実行](../tutorials/excel-power-automate-trigger.md)すると、プラットフォームの統合の詳細については、こちらを参照してください。

## <a name="excel-online-business-connector"></a>Excel Online (Business) コネクタ

[コネクタ](/connectors/connectors)は、電力の自動化とアプリケーションの間のブリッジです。 [Excel Online (Business) コネクタ](/connectors/excelonlinebusiness)を使用すると、excel ブックへのアクセスがフローに付与されます。 "スクリプトを実行する" アクションを使用すると、選択したブックからアクセス可能な Office スクリプトを呼び出すことができます。 フローを使用してスクリプトを実行できるだけでなく、スクリプト内のフローを使用して、ブックとの間でデータをやり取りすることもできます。

> [!IMPORTANT]
> "スクリプトを実行する" アクションを実行すると、Excel コネクタを使用するユーザーに、ブックとそのデータに対して重要なアクセス権が与えられます。 また、外部の[呼び出しからの外部呼び出し](external-calls.md)について説明するように、外部 API を呼び出すスクリプトにはセキュリティリスクがあります。 管理者が非常に機密性の高いデータの公開を懸念している場合は、Excel Online コネクタをオフにするか、 [Office スクリプト管理者コントロール](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)を使用して office スクリプトへのアクセスを制限することができます。

## <a name="passing-data-from-power-automate-into-a-script"></a>Power オートメーションからスクリプトにデータを渡す

すべてのスクリプトの入力は、関数の追加パラメーターとして指定され `main` ます。 たとえば、入力として名前を表すを受け入れるスクリプトが必要な場合は、 `string` `main` 署名をに変更し `function main(workbook: ExcelScript.Workbook, name: string)` ます。

Power 自動化でフローを構成するときは、スクリプトの入力を静的な値、[式](/power-automate/use-expressions-in-conditions)、または動的コンテンツとして指定できます。 個々のサービスのコネクタの詳細については、「[電源自動化コネクタ](/connectors/)」のドキュメントを参照してください。

入力パラメーターをスクリプトの関数に追加するときは `main` 、次の制限と制限事項を考慮してください。

1. 最初のパラメーターの型はでなければなりません `ExcelScript.Workbook` 。 そのパラメーター名は重要ではありません。

2. すべてのパラメーターには、型を指定する必要があります。

3. 基本的な型、、、、、、 `string` `number` `boolean` `any` `unknown` `object` 、 `undefined` がサポートされています。

4. 前にリストされていた基本的な種類の配列がサポートされています。

5. 入れ子になった配列は、パラメーターとしてサポートされます (戻り値の型としてではありません)。

6. 共用体型は、1つの型 ( `string` 、 `number` 、または) に属するリテラルの和集合である場合に使用でき `boolean` ます。 サポートされている型と未定義の共用体もサポートされています。

7. オブジェクトの種類は、型 `string` 、 `number` 、、 `boolean` サポートされている配列、またはその他のサポートされているオブジェクトのプロパティが含まれている場合に許可されます。 次の例は、パラメータタイプとしてサポートされているネストされたオブジェクトを示しています。

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

8. オブジェクトのインターフェイスまたはクラス定義は、スクリプトで定義されている必要があります。 また、次の例に示すように、オブジェクトを匿名でインラインで定義することもできます。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. 省略可能なパラメーターを指定できます。オプションの修飾子 (たとえば、) を使用することもでき `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` ます。

10. 既定のパラメーター値を使用できます (例 `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` :

## <a name="returning-data-from-a-script-back-to-power-automate"></a>スクリプトからのデータを Power 自動化に戻す

スクリプトは、Power オートメーションフローで動的コンテンツとして使用するブックからのデータを返すことができます。 入力パラメーターと同様に、Power オートメーションでは戻り値の型にいくつかの制限が課されます。

1. 基本的な型、、、、、 `string` `number` がサポートされてい `boolean` `void` `undefined` ます。

2. 戻り値の型として使用される共用体型は、スクリプトパラメーターとして使用する場合と同じ制限に従います。

3. 配列型は `string` 、型、、またはのいずれかである場合に使用でき `number` `boolean` ます。 また、型がサポートされている共用体型またはサポートされているリテラル型の場合にも使用できます。

4. 戻り値の型として使用されるオブジェクトの種類は、スクリプトパラメーターとして使用する場合と同じ制限に従います。

5. 暗黙的な入力はサポートされていますが、定義された型と同じルールに従う必要があります。

## <a name="avoid-using-relative-references"></a>相対参照の使用を避ける

Power オートメーションは、ユーザーの代わりに、選択した Excel ブックでスクリプトを実行します。 これが発生すると、ブックが閉じられる場合があります。 など、ユーザーの現在の状態に依存する API は、 `Workbook.getActiveWorksheet` 電力の自動処理によって実行されると失敗します。 スクリプトを設計するときは、必ずワークシートおよび範囲の絶対参照を使用してください。

次の関数は、Power オートメーションフローでスクリプトから呼び出されたときにエラーをスローして失敗します。

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a>例

次のスクリーンショットは、 [GitHub](https://github.com/)の問題がユーザーに割り当てられたときにトリガーされる電源自動化フローを示しています。 このフローは、Excel ブックのテーブルに問題を追加するスクリプトを実行します。 そのテーブルに5つ以上の問題がある場合、フローはメール事前通知を送信します。

![電源自動化フローエディターに示されている例のフロー。](../images/power-automate-parameter-return-sample.png)

`main`スクリプトの関数は、[案件 ID] と [issue title] を入力パラメーターとして指定し、スクリプトは issue テーブル内の行数を返します。

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

- [Power オートメーションを使用して web 上の Excel で Office スクリプトを実行する](../tutorials/excel-power-automate-manual.md)
- [パワー自動化を使用してスクリプトを自動的に実行する](../tutorials/excel-power-automate-trigger.md)
- [Excel on the web での Office スクリプトのスクリプトの基本事項](scripting-fundamentals.md)
- [Power Automate の使用を開始する](/power-automate/getting-started)
- [Excel Online (ビジネス向け) コネクタのリファレンスドキュメント](/connectors/excelonlinebusiness/)
