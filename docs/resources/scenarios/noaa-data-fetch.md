---
title: 'Office スクリプトのサンプル シナリオ: NOAA からのグラフの水レベル データ'
description: NOAA データベースから JSON データをフェッチし、それを使用してグラフを作成するサンプル。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 5b0b4e3675cbe053368f63123d819f0dab626e60
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867878"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="983c7-103">Office スクリプトのサンプル シナリオ: NOAA から水レベルのデータを取得およびグラフ化する</span><span class="sxs-lookup"><span data-stu-id="983c7-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="983c7-104">このシナリオでは、National National Wateric and の の管理のシアトルステーションで、水レベル [をプロットする必要があります](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)。</span><span class="sxs-lookup"><span data-stu-id="983c7-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="983c7-105">外部データを使用してスプレッドシートにデータを入力し、グラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="983c7-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="983c7-106">このコマンドを使用して NOAA の多用データベースと Currents データベースにクエリを実行するスクリプト `fetch` [を開発します](https://tidesandcurrents.noaa.gov/)。</span><span class="sxs-lookup"><span data-stu-id="983c7-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="983c7-107">これは、指定した期間にわたって記録された水レベルを取得します。</span><span class="sxs-lookup"><span data-stu-id="983c7-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="983c7-108">情報は JSON として返されます。そのため、スクリプトの一部は、その情報を範囲の値に変換します。</span><span class="sxs-lookup"><span data-stu-id="983c7-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="983c7-109">データがスプレッドシートに入った後、グラフの作成に使用されます。</span><span class="sxs-lookup"><span data-stu-id="983c7-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="983c7-110">スクリプトのスキルの説明</span><span class="sxs-lookup"><span data-stu-id="983c7-110">Scripting skills covered</span></span>

- <span data-ttu-id="983c7-111">外部 API 呼び出し ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="983c7-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="983c7-112">JSON 解析</span><span class="sxs-lookup"><span data-stu-id="983c7-112">JSON parsing</span></span>
- <span data-ttu-id="983c7-113">グラフ</span><span class="sxs-lookup"><span data-stu-id="983c7-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="983c7-114">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="983c7-114">Setup instructions</span></span>

1. <span data-ttu-id="983c7-115">Excel on the web でブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="983c7-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="983c7-116">[自動化] **タブで** 、[すべてのスクリプト] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="983c7-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="983c7-117">[コード **エディター] 作業ウィンドウで** 、[新しいスクリプト] **を選択し** 、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="983c7-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);

      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1", 
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    }
    ```

1. <span data-ttu-id="983c7-118">スクリプトの名前を **NOAA 水レベル グラフに変更し** 、保存します。</span><span class="sxs-lookup"><span data-stu-id="983c7-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="983c7-119">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="983c7-119">Running the script</span></span>

<span data-ttu-id="983c7-120">任意のワークシートで **、NOAA 水位グラフ スクリプトを実行** します。</span><span class="sxs-lookup"><span data-stu-id="983c7-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="983c7-121">このスクリプトは、2020 年 12 月 25 日から 2020 年 12 月 27 日まで、水レベルのデータをフェッチします。</span><span class="sxs-lookup"><span data-stu-id="983c7-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="983c7-122">スクリプトの先頭にある変数は、異なる日付を使用するか、別のステーション情報を取得 `const` するために変更できます。</span><span class="sxs-lookup"><span data-stu-id="983c7-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="983c7-123">データ [取得用の CO-OPS API](https://api.tidesandcurrents.noaa.gov/api/prod/) では、このすべてのデータを取得する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="983c7-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="983c7-124">スクリプトの実行後</span><span class="sxs-lookup"><span data-stu-id="983c7-124">After running the script</span></span>

![スクリプトを実行した後のワークシートには、いくつかの水位データとグラフが表示されます。](../../images/scenario-noaa-water-level-after.png)
