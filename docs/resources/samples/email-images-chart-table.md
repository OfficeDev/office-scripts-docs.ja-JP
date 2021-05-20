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
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="5a6c2-103">OfficeスクリプトとPower Automateを使用して、チャートとテーブルの画像を電子メールで送信する</span><span class="sxs-lookup"><span data-stu-id="5a6c2-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="5a6c2-104">このサンプルでは、Office スクリプトとPower Automateを使用してグラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="5a6c2-105">次に、チャートとそのベーステーブルの画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="5a6c2-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="5a6c2-106">Example scenario</span></span>

* <span data-ttu-id="5a6c2-107">計算して最新の結果を得る。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="5a6c2-108">グラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-108">Create chart.</span></span>
* <span data-ttu-id="5a6c2-109">チャートとテーブルの画像を取得します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-109">Get chart and table images.</span></span>
* <span data-ttu-id="5a6c2-110">Power Automateで画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="5a6c2-111">_入力データ_</span><span class="sxs-lookup"><span data-stu-id="5a6c2-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="入力データのテーブルを示すワークシート":::

<span data-ttu-id="5a6c2-113">_出力チャート_</span><span class="sxs-lookup"><span data-stu-id="5a6c2-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="顧客別の支払額を示して作成された縦棒グラフ":::

<span data-ttu-id="5a6c2-115">_Power Automateフローを通じて受信された電子メール_</span><span class="sxs-lookup"><span data-stu-id="5a6c2-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれたExcelチャートを示すフローによって送信された電子メール":::

## <a name="solution"></a><span data-ttu-id="5a6c2-117">ソリューション</span><span class="sxs-lookup"><span data-stu-id="5a6c2-117">Solution</span></span>

<span data-ttu-id="5a6c2-118">このソリューションには、次の 2 つの部分があります。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="5a6c2-119">グラフと表を計算して抽出Excel Officeスクリプト</span><span class="sxs-lookup"><span data-stu-id="5a6c2-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="5a6c2-120">スクリプトを呼び出して結果を電子メールで送信するPower Automateフロー。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="5a6c2-121">この方法の例については、「 [Power Automate を使用した自動化ワークフローの作成](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="5a6c2-122">サンプル コード: グラフとテーブルExcel計算および抽出</span><span class="sxs-lookup"><span data-stu-id="5a6c2-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="5a6c2-123">次のスクリプトは、Excelのグラフと表を計算して抽出します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="5a6c2-124">サンプルファイル <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> をダウンロードし、このスクリプトで使用して自分で試してみてください!</span><span class="sxs-lookup"><span data-stu-id="5a6c2-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="5a6c2-125">Power Automateフロー: チャートとテーブルの画像を電子メールで送信する</span><span class="sxs-lookup"><span data-stu-id="5a6c2-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="5a6c2-126">このフローはスクリプトを実行し、返された画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="5a6c2-127">新しい **インスタント クラウド フロー** を作成する :</span><span class="sxs-lookup"><span data-stu-id="5a6c2-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="5a6c2-128">[ **フローを手動でトリガーする] を** 選択し、[ **作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="5a6c2-129">**[スクリプトの実行**] アクションで **、オンライン (ビジネス)** コネクタExcelを使用する **新しいステップ** を追加します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5a6c2-130">アクションには次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="5a6c2-131">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5a6c2-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="5a6c2-132">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5a6c2-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="5a6c2-133">**ファイル**: ワークブック ([ファイル選択で選択 )](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="5a6c2-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="5a6c2-134">**スクリプト**: スクリプト名</span><span class="sxs-lookup"><span data-stu-id="5a6c2-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Power Automateの完了Excelオンライン (ビジネス) コネクタ":::
1. <span data-ttu-id="5a6c2-136">このサンプルでは、電子メール クライアントとしてOutlookを使用します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="5a6c2-137">サポートする任意の電子メール コネクタPower Automateを使用できますが、残りの手順では、Outlook選択したと想定しています。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="5a6c2-138">**Office 365 Outlook** コネクタと **送信と電子メール (V2)** アクションを使用する **新しい手順** を追加します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="5a6c2-139">アクションには次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="5a6c2-140">**To**: テスト用の電子メール アカウント (または個人用メール)</span><span class="sxs-lookup"><span data-stu-id="5a6c2-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="5a6c2-141">**件名**: レポートデータを確認してください</span><span class="sxs-lookup"><span data-stu-id="5a6c2-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="5a6c2-142">[ **本文** ] フィールドで 、[コード ビュー] ( `</>` ) を選択し、次のように入力します。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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
1. <span data-ttu-id="5a6c2-144">フローを保存し、それを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="5a6c2-145">トレーニングビデオ:チャートとテーブルの画像を抽出して電子メールで送信する</span><span class="sxs-lookup"><span data-stu-id="5a6c2-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="5a6c2-146">[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/152GJyqc-Kw)。</span><span class="sxs-lookup"><span data-stu-id="5a6c2-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
