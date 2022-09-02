---
title: Excel グラフとテーブルの画像をEmailする
description: Office スクリプトと Power Automate を使用して Excel グラフとテーブルの画像を抽出して電子メールで送信する方法について説明します。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572466"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Office スクリプトと Power Automate を使用して、グラフとテーブルの画像を電子メールで送信する

このサンプルでは、Office スクリプトと Power Automate を使用してグラフを作成します。 次に、グラフとその基本テーブルの画像を電子メールで送信します。

## <a name="example-scenario"></a>シナリオ例

* 最新の結果を取得するために計算します。
* グラフを作成します。
* グラフとテーブルの画像を取得します。
* Power Automate を使用してイメージをEmailします。

_入力データ_

:::image type="content" source="../../images/input-data.png" alt-text="入力データのテーブルを示すワークシート。":::

_出力グラフ_

:::image type="content" source="../../images/chart-created.png" alt-text="顧客の支払額を示す縦棒グラフ。":::

_Power Automate フローを介して受信したEmail_

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれた Excel グラフを示すフローによって送信された電子メール。":::

## <a name="solution"></a>ソリューション

このソリューションには、次の 2 つの部分があります。

1. [Excel のグラフとテーブルを計算して抽出する Office スクリプト](#sample-code-calculate-and-extract-excel-chart-and-table)
1. スクリプトを呼び出し、結果を電子メールで送信する Power Automate フロー。 これを行う方法の例については、「 [Power Automate を使用して自動化されたワークフローを作成する](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)」を参照してください。

## <a name="sample-excel-file"></a>Excel ファイルのサンプル

すぐに使用できるブックの [email-chart-table.xlsx](email-chart-table.xlsx) をダウンロードします。 サンプルを自分で試すには、次のスクリプトを追加します。

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>サンプル コード: Excel グラフとテーブルを計算して抽出する

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate フロー: グラフとテーブルイメージをEmailする

このフローはスクリプトを実行し、返された画像に電子メールを送信します。

1. 新しい **インスタント クラウド フロー** を作成します。
1. [ **手動でフローをトリガーする** ] を選択し、[ **作成**] を選択します。
1. **スクリプトの実行** アクションで **Excel Online (Business)** コネクタを使用する **新しいステップ** を追加します。 アクションには次の値を使用します。
    * **場所**: OneDrive for Business
    * **ドキュメント ライブラリ**: OneDrive
    * **ファイル**: ブック ([ファイル選択子で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **スクリプト**: スクリプト名

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Power Automate で完成した Excel Online (Business) コネクタ。":::
1. このサンプルでは、Outlook を電子メール クライアントとして使用します。 Power Automate でサポートされている電子メール コネクタを使用できますが、残りの手順では Outlook を選択したことを前提としています。 **Office 365 Outlook** コネクタと **送信と電子メール (V2)** アクションを使用する **新しい手順** を追加します。 アクションには次の値を使用します。
    * **To**: テスト用メール アカウント (または個人用メール)
    * **件名**: レポート データを確認してください
    * **[本文**] フィールドで [コード ビュー] (`</>`) を選択し、次のように入力します。

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Power Automate の完成したOffice 365 Outlook コネクタ。":::
1. フローを保存して試してください。フロー エディター ページの **[テスト** ] ボタンを使用するか、[ **マイ フロー** ] タブでフローを実行します。メッセージが表示されたら、必ずアクセスを許可してください。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>トレーニング ビデオ: グラフとテーブルの画像を抽出して電子メールで送信する

[YouTube でこのサンプルを見る、スディ Ramamurthy のチュートリアルをご覧ください](https://youtu.be/152GJyqc-Kw)。
