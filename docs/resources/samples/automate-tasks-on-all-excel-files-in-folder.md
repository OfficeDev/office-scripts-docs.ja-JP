---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: フォルダー内のすべてのファイルに対してスクリプトExcel実行する方法について説明OneDrive for Business。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: fad1483fbcddaf541874630e8a4e5a06faa784627d44d17ea2ab7ca0af1550a4
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847413"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>フォルダー内のすべての Excel ファイルでスクリプトを実行する

このプロジェクトは、フォルダー内のすべてのファイルに対して一連の自動化タスクを実行OneDrive for Business。 また、フォルダー内のフォルダー SharePointすることもできます。
このプロパティは、Excelファイルに対して計算を実行し、書式設定を追加し、同僚にコメント[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)挿入します。

## <a name="sample-excel-files"></a>サンプル Excel ファイル

この <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> 必要なすべてのブックの詳細をダウンロードします。 これらのファイルを Sales というタイトルのフォルダーに **展開します**。 次のスクリプトをスクリプト コレクションに追加して、サンプルを自分で試してみてください。

## <a name="sample-code-add-formatting-and-insert-comment"></a>サンプル コード: 書式の追加とコメントの挿入

これは、個々のブックで実行されるスクリプトです。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automateフロー: フォルダー内のすべてのブックでスクリプトを実行する

このフローは、"Sales" フォルダー内のすべてのブックでスクリプトを実行します。

1. 新しいインスタント クラウド **フローを作成します**。
1. [フロー **を手動でトリガーする] を選択し** 、[作成] を **選択します**。
1. [フォルダー内 **のファイルの一** 覧] **OneDrive for Businessを使用** する新 **しい手順を追加** します。

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. 抽出されたブックを含む "Sales" フォルダーを選択します。
1. ブックのみを選択するには、[新しい手順]を選択し、[条件] を **選択します**。 条件には、次の値を使用します。
    1. **名前**(ファイルOneDrive値)
    1. "ends with"
    1. "xlsx"

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="後続Power Automateを各ファイルに適用する条件ブロックを指定します。":::
1. [**はい] ブランチで**、[スクリプトの実行] アクションExcel **オンライン (Business)** コネクタ **を追加** します。 アクションには、次の値を使用します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: **Id** (OneDrive ID 値)
    1. **スクリプト**: スクリプト名

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. フローを保存し、試してみてください。[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>トレーニング ビデオ: フォルダー内のすべてのファイルExcelスクリプトを実行する

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/xMg711o7k6w).
