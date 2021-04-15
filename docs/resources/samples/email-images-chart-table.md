---
title: Excel グラフと表の画像をメールで送信する
description: Excel のグラフとテーブルのOfficeを抽出して電子メールで送信するには、スクリプトと Power Automate を使用する方法について学習します。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754811"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="39396-103">グラフOffice表の画像を電子メールで送信するには、スクリプトと Power Automate を使用します。</span><span class="sxs-lookup"><span data-stu-id="39396-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="39396-104">このサンプルでは、Officeスクリプトと Power Automate を使用してグラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="39396-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="39396-105">次に、グラフとその基本テーブルの画像を電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="39396-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="39396-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="39396-106">Example scenario</span></span>

* <span data-ttu-id="39396-107">最新の結果を取得するために計算します。</span><span class="sxs-lookup"><span data-stu-id="39396-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="39396-108">グラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="39396-108">Create chart.</span></span>
* <span data-ttu-id="39396-109">グラフと表の画像を取得します。</span><span class="sxs-lookup"><span data-stu-id="39396-109">Get chart and table images.</span></span>
* <span data-ttu-id="39396-110">Power Automate を使用して画像にメールを送信します。</span><span class="sxs-lookup"><span data-stu-id="39396-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="39396-111">_入力データ_</span><span class="sxs-lookup"><span data-stu-id="39396-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="入力データの表を示すワークシート。":::

<span data-ttu-id="39396-113">_出力グラフ_</span><span class="sxs-lookup"><span data-stu-id="39396-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="顧客による金額を示す列グラフが作成されました。":::

<span data-ttu-id="39396-115">_Power Automate フローを通じて受信された電子メール_</span><span class="sxs-lookup"><span data-stu-id="39396-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="本文に埋め込まれた Excel グラフを示すフローによって送信される電子メール。":::

## <a name="solution"></a><span data-ttu-id="39396-117">ソリューション</span><span class="sxs-lookup"><span data-stu-id="39396-117">Solution</span></span>

<span data-ttu-id="39396-118">このソリューションには、次の 2 つの部分があります。</span><span class="sxs-lookup"><span data-stu-id="39396-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="39396-119">Excel Officeテーブルを計算して抽出するためのスクリプト</span><span class="sxs-lookup"><span data-stu-id="39396-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="39396-120">スクリプトを呼び出して結果を電子メールで送信する Power Automate フロー。</span><span class="sxs-lookup"><span data-stu-id="39396-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="39396-121">これを行う方法の例については、「Power Automat を使用して自動化された [ワークフローを作成する」を参照してください](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="39396-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="39396-122">サンプル コード: Excel グラフとテーブルを計算して抽出する</span><span class="sxs-lookup"><span data-stu-id="39396-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="39396-123">次のスクリプトは、Excel のグラフとテーブルを計算して抽出します。</span><span class="sxs-lookup"><span data-stu-id="39396-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="39396-124">サンプル ファイルをダウンロード <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> このスクリプトで使用して、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="39396-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="39396-125">トレーニング ビデオ: グラフとテーブルの画像を抽出して電子メールで送信する</span><span class="sxs-lookup"><span data-stu-id="39396-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="39396-126">[![グラフとテーブルの画像を抽出して電子メールで送信する方法について、ステップバイステップのビデオを見る](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "グラフとテーブルの画像を抽出して電子メールで送信する方法に関するステップバイステップのビデオ")</span><span class="sxs-lookup"><span data-stu-id="39396-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
