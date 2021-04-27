---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: OneDrive for Business のフォルダー内のすべての Excel ファイルでスクリプトを実行する方法について説明します。
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571489"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>フォルダー内のすべての Excel ファイルでスクリプトを実行する

このプロジェクトは、OneDrive for Business のフォルダー内のすべてのファイルに対して一連の自動化タスクを実行します。 SharePoint フォルダーでも使用できます。
Excel ファイルの計算を実行し、書式設定を追加し、同僚にコメント [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 挿入します。

## <a name="sample-code-add-formatting-and-insert-comment"></a>サンプル コード: 書式の追加とコメントの挿入

ファイルをダウンロード<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>サンプルで使用されている Sales というタイトルのフォルダーにファイルを抽出し、自分で試してみてください。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>トレーニング ビデオ: フォルダー内のすべての Excel ファイルでスクリプトを実行する

[](https://youtu.be/xMg711o7k6w) OneDrive for Business または SharePoint フォルダー内のすべての Excel ファイルでスクリプトを実行する方法の詳細なビデオをご覧ください。