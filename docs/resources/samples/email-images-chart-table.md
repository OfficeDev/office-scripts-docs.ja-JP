---
title: グラフと表の画像Excelメールで送信する
description: '[スクリプト] と [Office] Power Automateを使用して、グラフと表の画像Excelメールを送信する方法について学習します。'
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 5eb20025462614d62774ae6c088bdf397dcfb39d
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074593"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="23c5f-103">グラフOfficeの画像をPower Automateする場合は、スクリプトとスクリプトを使用してメールを送信します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="23c5f-104">このサンプルでは、OfficeスクリプトとPower Automateを使用してグラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="23c5f-105">次に、グラフとその基本テーブルの画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="23c5f-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="23c5f-106">Example scenario</span></span>

* <span data-ttu-id="23c5f-107">最新の結果を取得するために計算します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="23c5f-108">グラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-108">Create chart.</span></span>
* <span data-ttu-id="23c5f-109">グラフと表の画像を取得します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-109">Get chart and table images.</span></span>
* <span data-ttu-id="23c5f-110">画像にメールを送信Power Automate。</span><span class="sxs-lookup"><span data-stu-id="23c5f-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="23c5f-111">_入力データ_</span><span class="sxs-lookup"><span data-stu-id="23c5f-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="入力データの表を示すワークシート。":::

<span data-ttu-id="23c5f-113">_出力グラフ_</span><span class="sxs-lookup"><span data-stu-id="23c5f-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="顧客による金額を示す列グラフが作成されました。":::

<span data-ttu-id="23c5f-115">_メール フローを通じて受信Power Automateメール_</span><span class="sxs-lookup"><span data-stu-id="23c5f-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれたグラフのExcelによって送信される電子メール。":::

## <a name="solution"></a><span data-ttu-id="23c5f-117">ソリューション</span><span class="sxs-lookup"><span data-stu-id="23c5f-117">Solution</span></span>

<span data-ttu-id="23c5f-118">このソリューションには、次の 2 つの部分があります。</span><span class="sxs-lookup"><span data-stu-id="23c5f-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="23c5f-119">グラフOfficeテーブルを計算して抽出するExcelスクリプト</span><span class="sxs-lookup"><span data-stu-id="23c5f-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="23c5f-120">スクリプトPower Automate結果を電子メールで送信するフローを示します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="23c5f-121">これを行う方法の例については、「自動ワークフローを作成[する」を参照Power Automate。](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="23c5f-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="23c5f-122">サンプル コード: グラフと表のExcel抽出する</span><span class="sxs-lookup"><span data-stu-id="23c5f-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="23c5f-123">次のスクリプトは、グラフと表のExcel抽出します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="23c5f-124">サンプル ファイルをダウンロード <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> このスクリプトで使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="23c5f-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="23c5f-125">Power Automateフロー: グラフと表の画像をメールで送信する</span><span class="sxs-lookup"><span data-stu-id="23c5f-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="23c5f-126">このフローはスクリプトを実行し、返された画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="23c5f-127">新しいインスタント クラウド **フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="23c5f-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="23c5f-128">[フロー **を手動でトリガーする] を選択し** 、[作成] を **押します**。</span><span class="sxs-lookup"><span data-stu-id="23c5f-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="23c5f-129">[スクリプト **の実行]** アクションを使用して、Excel **(Business)** コネクタを使用する新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="23c5f-130">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="23c5f-131">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="23c5f-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="23c5f-132">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="23c5f-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="23c5f-133">**ファイル**: ブック ([ファイル選択ウィンドウで選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="23c5f-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="23c5f-134">**スクリプト**: スクリプト名</span><span class="sxs-lookup"><span data-stu-id="23c5f-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. <span data-ttu-id="23c5f-136">このサンプルでは、Outlookクライアントとして使用します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="23c5f-137">サポートされている任意の電子メール コネクタPower Automate使用できますが、残りの手順では、メール コネクタを選択Outlook。</span><span class="sxs-lookup"><span data-stu-id="23c5f-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="23c5f-138">新しい **手順を追加** して、Office 365 Outlook **および電子** メール **(V2) アクションを使用** します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="23c5f-139">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="23c5f-140">**To**: テスト用メール アカウント (または個人用メール)</span><span class="sxs-lookup"><span data-stu-id="23c5f-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="23c5f-141">**件名**: レポート データを確認してください</span><span class="sxs-lookup"><span data-stu-id="23c5f-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="23c5f-142">[本文 **] フィールドで** 、[コード ビュー] ( ) を選択 `</>` し、次の値を入力します。</span><span class="sxs-lookup"><span data-stu-id="23c5f-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Office 365 OutlookでPower Automate。":::
1. <span data-ttu-id="23c5f-144">フローを保存し、試してみてください。</span><span class="sxs-lookup"><span data-stu-id="23c5f-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="23c5f-145">トレーニング ビデオ: グラフとテーブルの画像を抽出して電子メールで送信する</span><span class="sxs-lookup"><span data-stu-id="23c5f-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="23c5f-146">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="23c5f-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
