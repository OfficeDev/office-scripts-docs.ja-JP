---
title: グラフと表の画像Excelメールで送信する
description: '[スクリプトとスクリプト] OfficeをPower Automateして、グラフと表の画像を抽出Excelメールで送信する方法について学習します。'
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2930a70a5bed4eb49f33f315460ae32f40b5a2f2
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585507"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>グラフOffice表の画像Power Automateメールを送信するには、スクリプトとスクリプトを使用します。

このサンプルでは、OfficeスクリプトとPower Automateを使用してグラフを作成します。 次に、グラフとその基本テーブルの画像を電子メールで送信します。

## <a name="example-scenario"></a>シナリオ例

* 最新の結果を取得するために計算します。
* グラフを作成します。
* グラフと表の画像を取得します。
* 画像にメールを送信Power Automate。

_入力データ_

:::image type="content" source="../../images/input-data.png" alt-text="入力データの表を示すワークシート。":::

_出力グラフ_

:::image type="content" source="../../images/chart-created.png" alt-text="顧客による金額を示す列グラフが作成されました。":::

_メール フローを通じて受信Power Automateメール_

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれたグラフのExcelによって送信される電子メール。":::

## <a name="solution"></a>ソリューション

このソリューションには、次の 2 つの部分があります。

1. [グラフOfficeテーブルを計算して抽出Excelスクリプト](#sample-code-calculate-and-extract-excel-chart-and-table)
1. スクリプトPower Automateし、結果を電子メールで送信するフローを示します。 これを行う方法の例については、「自動ワークフロー[を作成する」](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)を参照Power Automate。

## <a name="sample-excel-file"></a>サンプル Excel ファイル

すぐに <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> ブックのダウンロード を行います。 次のスクリプトを追加して、サンプルを自分で試してみてください。

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>サンプル コード: グラフと表のExcel抽出する

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automateフロー: グラフと表の画像をメールで送信する

このフローはスクリプトを実行し、返された画像を電子メールで送信します。

1. 新しいインスタント クラウド **フローを作成します**。
1. [フロー **を手動でトリガーする] を選択し、[** 作成] を **選択します**。
1. [スクリプト **の実行]** アクションを使用して、Excel **(Business)** コネクタを使用する新しい **手順を追加** します。 アクションには、次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: ブック ([ファイル選択ウィンドウで選択](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **スクリプト**: スクリプト名

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. このサンプルでは、Outlookクライアントとして使用します。 サポートされている任意の電子メール コネクタPower Automate使用できますが、残りの手順では、ユーザーがサポートを選択Outlook。 新しい **手順を追加** して、**Office 365 Outlookおよび電子** メール **(V2) アクションを使用** します。 アクションには、次の値を使用します。
    * **To**: テストメール アカウント (または個人用メール)
    * **件名**: レポート データを確認してください
    * [本文 **] フィールドで** [コード ビュー] () を選択`</>`し、次の値を入力します。

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="完了したOffice 365 OutlookコネクタをPower Automate。":::
1. フローを保存し、試してみてください。[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>トレーニング ビデオ: グラフとテーブルの画像を抽出して電子メールで送信する

[Sudhi Ramamurthy が YouTube でこのサンプルを見るのを見る](https://youtu.be/152GJyqc-Kw)。
