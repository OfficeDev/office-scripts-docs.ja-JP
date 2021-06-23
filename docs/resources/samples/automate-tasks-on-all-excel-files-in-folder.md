---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: フォルダー内のすべてのファイルに対してスクリプトExcel実行する方法について説明OneDrive for Business。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 83e091a8b009bac577da9ed53dcf4139c1b845c9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074586"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>フォルダー内のすべての Excel ファイルでスクリプトを実行する

このプロジェクトは、フォルダー内のすべてのファイルに対して一連の自動化タスクを実行OneDrive for Business。 また、フォルダー内のフォルダー SharePointすることもできます。
このプロパティは、Excelファイルに対して計算を実行し、書式設定を追加し、同僚にコメント[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)挿入します。

ファイルをダウンロード<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>サンプルで使用されている Sales というタイトルのフォルダーにファイルを抽出し、自分で試してみてください。

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
1. [フロー **を手動でトリガーする] を選択し** 、[作成] を **押します**。
1. [フォルダー内 **のファイルの一** 覧] **OneDrive for Businessを使用** する新 **しい手順を追加** します。

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. 抽出されたブックを含む "Sales" フォルダーを選択します。
1. ブックのみを選択するには、[新しい手順] を選択し、[条件]**を選択****し**、次の値を設定します。
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
1. フローを保存し、試してみてください。

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>トレーニング ビデオ: フォルダー内のすべてのファイルExcelスクリプトを実行する

[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/xMg711o7k6w).
