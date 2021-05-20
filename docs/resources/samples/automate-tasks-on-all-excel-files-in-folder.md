---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: OneDrive for Businessのフォルダ内のすべてのExcel ファイルに対してスクリプトを実行する方法について説明します。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545793"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>フォルダー内のすべての Excel ファイルでスクリプトを実行する

このプロジェクトは、OneDrive for Businessのフォルダにあるすべてのファイルに対して、一連のオートメーション タスクを実行します。 また、SharePointフォルダでも使用できます。
Excel ファイルに対して計算を実行し、書式を追加し、同僚[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)コメントを挿入します。

<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">ファイルhighlight-alert-excel-files.zip</a>をダウンロードし、このサンプルで使用されている **Sales** というタイトルのフォルダにファイルを抽出し、自分で試してみてください!

## <a name="sample-code-add-formatting-and-insert-comment"></a>サンプル コード: 書式を追加してコメントを挿入する

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automateフロー: フォルダー内のすべてのブックに対してスクリプトを実行します。

このフローは、"Sales" フォルダー内のすべてのブックでスクリプトを実行します。

1. 新しい **インスタント クラウド フロー** を作成する :
1. [ **フローを手動でトリガーする] を** 選択し、[ **作成]** を押します。
1. **[OneDrive for Business** コネクタ] アクションと [**フォルダー内のファイルを一覧表示する]** アクションを使用する **新しい手順** を追加します。

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Power Automate の完成したOneDrive for Business コネクタ":::
1. 抽出したワークブックを含む "Sales" フォルダーを選択します。
1. ブックのみが選択されていることを確認するには、[ **新しいステップ**] を選択し、[ **条件** ] を選択して次の値を設定します。
    1. **名前**(OneDrive ファイル名の値)
    1. 「で終わる」
    1. "xlsx"。

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="後続のアクションを各ファイルに適用するPower Automate条件ブロック":::
1. [**はいの場合**] の下で、[**スクリプトの実行**] アクションを使用して **Excelオンライン (ビジネス)** コネクタを追加します。 アクションには次の値を使用します。
    1. **場所**: OneDrive for Business
    1. **ドキュメント ライブラリ**: OneDrive
    1. **ファイル**: **ID** (OneDriveファイル ID 値)
    1. **スクリプト**: スクリプト名

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Power Automateの完了Excelオンライン (ビジネス) コネクタ":::
1. フローを保存し、それを試してみてください。

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>トレーニング ビデオ: フォルダー内のすべてのExcel ファイルに対してスクリプトを実行する

[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/xMg711o7k6w)。
