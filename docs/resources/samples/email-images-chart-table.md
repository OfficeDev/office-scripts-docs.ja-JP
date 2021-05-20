---
title: Excelのグラフと表の画像を電子メールで送信する
description: OfficeスクリプトとPower Automateを使用して、Excelのグラフと表の画像を抽出して電子メールで送信する方法について説明します。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545778"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>OfficeスクリプトとPower Automateを使用して、チャートとテーブルの画像を電子メールで送信する

このサンプルでは、Office スクリプトとPower Automateを使用してグラフを作成します。 次に、チャートとそのベーステーブルの画像を電子メールで送信します。

## <a name="example-scenario"></a>シナリオ例

* 計算して最新の結果を得る。
* グラフを作成します。
* チャートとテーブルの画像を取得します。
* Power Automateで画像を電子メールで送信します。

_入力データ_

:::image type="content" source="../../images/input-data.png" alt-text="入力データのテーブルを示すワークシート":::

_出力チャート_

:::image type="content" source="../../images/chart-created.png" alt-text="顧客別の支払額を示して作成された縦棒グラフ":::

_Power Automateフローを通じて受信された電子メール_

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれたExcelチャートを示すフローによって送信された電子メール":::

## <a name="solution"></a>ソリューション

このソリューションには、次の 2 つの部分があります。

1. [グラフと表を計算して抽出Excel Officeスクリプト](#sample-code-calculate-and-extract-excel-chart-and-table)
1. スクリプトを呼び出して結果を電子メールで送信するPower Automateフロー。 この方法の例については、「 [Power Automate を使用した自動化ワークフローの作成](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)」を参照してください。

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>サンプル コード: グラフとテーブルExcel計算および抽出

次のスクリプトは、Excelのグラフと表を計算して抽出します。

サンプルファイル <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> をダウンロードし、このスクリプトで使用して自分で試してみてください!

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automateフロー: チャートとテーブルの画像を電子メールで送信する

このフローはスクリプトを実行し、返された画像を電子メールで送信します。

1. 新しい **インスタント クラウド フロー** を作成する :
1. [ **フローを手動でトリガーする] を** 選択し、[ **作成]** を押します。
1. **[スクリプトの実行**] アクションで **、オンライン (ビジネス)** コネクタExcelを使用する **新しいステップ** を追加します。 アクションには次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: ワークブック ([ファイル選択で選択 )](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: スクリプト名

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Power Automateの完了Excelオンライン (ビジネス) コネクタ":::
1. このサンプルでは、電子メール クライアントとしてOutlookを使用します。 サポートする任意の電子メール コネクタPower Automateを使用できますが、残りの手順では、Outlook選択したと想定しています。 **Office 365 Outlook** コネクタと **送信と電子メール (V2)** アクションを使用する **新しい手順** を追加します。 アクションには次の値を使用します。
    * **To**: テスト用の電子メール アカウント (または個人用メール)
    * **件名**: レポートデータを確認してください
    * [ **本文** ] フィールドで 、[コード ビュー] ( `</>` ) を選択し、次のように入力します。

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Power Automateの完成したOffice 365 Outlook コネクタ":::
1. フローを保存し、それを試してみてください。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>トレーニングビデオ:チャートとテーブルの画像を抽出して電子メールで送信する

[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/152GJyqc-Kw)。
