---
title: 'Office スクリプトのサンプル シナリオ: NOAA からの水位データのグラフ'
description: NOAA データベースから JSON データをフェッチし、それを使用してグラフを作成するサンプル。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 1b9ad9115efe5b1924499f160f39d4b8fe654763
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570165"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="435da-103">Officeスクリプトのサンプル シナリオ: NOAA からの水レベルデータのフェッチとグラフ化</span><span class="sxs-lookup"><span data-stu-id="435da-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="435da-104">このシナリオでは、国立海洋大気局のシアトルステーションで水位を [プロットする必要があります](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)。</span><span class="sxs-lookup"><span data-stu-id="435da-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="435da-105">外部データを使用してスプレッドシートにデータを入力し、グラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="435da-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="435da-106">このコマンドを使用して NOAA Tides データベースと Currents データベースにクエリを実行するスクリプト `fetch` [を開発します](https://tidesandcurrents.noaa.gov/)。</span><span class="sxs-lookup"><span data-stu-id="435da-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="435da-107">この場合、指定した期間にわたって水位が記録されます。</span><span class="sxs-lookup"><span data-stu-id="435da-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="435da-108">情報は JSON として返されます。そのため、スクリプトの一部は、それを範囲の値に変換します。</span><span class="sxs-lookup"><span data-stu-id="435da-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="435da-109">データがスプレッドシートに入った後、グラフの作成に使用されます。</span><span class="sxs-lookup"><span data-stu-id="435da-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="435da-110">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="435da-110">Scripting skills covered</span></span>

- <span data-ttu-id="435da-111">外部 API 呼び出し ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="435da-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="435da-112">JSON の解析</span><span class="sxs-lookup"><span data-stu-id="435da-112">JSON parsing</span></span>
- <span data-ttu-id="435da-113">グラフ</span><span class="sxs-lookup"><span data-stu-id="435da-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="435da-114">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="435da-114">Setup instructions</span></span>

1. <span data-ttu-id="435da-115">Web 上の Excel を使用してブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="435da-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="435da-116">[自動化] **タブで** 、[すべてのスクリプト] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="435da-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="435da-117">[コード **エディター] 作業ウィンドウ** で、[ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="435da-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
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

1. <span data-ttu-id="435da-118">スクリプトの名前を **NOAA 水位グラフに変更し** 、保存します。</span><span class="sxs-lookup"><span data-stu-id="435da-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="435da-119">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="435da-119">Running the script</span></span>

<span data-ttu-id="435da-120">任意のワークシートで **、NOAA 水位グラフ スクリプトを実行** します。</span><span class="sxs-lookup"><span data-stu-id="435da-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="435da-121">このスクリプトは、2020 年 12 月 25 日から 2020 年 12 月 27 日まで水位データをフェッチします。</span><span class="sxs-lookup"><span data-stu-id="435da-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="435da-122">スクリプトの先頭にある変数は、異なる日付を使用したり、異なるステーション情報を取得したりするために `const` 変更できます。</span><span class="sxs-lookup"><span data-stu-id="435da-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="435da-123">[CO-OPS API For Data Retrieval では](https://api.tidesandcurrents.noaa.gov/api/prod/)、このすべてのデータを取得する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="435da-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="435da-124">スクリプトの実行後</span><span class="sxs-lookup"><span data-stu-id="435da-124">After running the script</span></span>

![スクリプトを実行した後のワークシートには、水位データとグラフが表示されます。](../../images/scenario-noaa-water-level-after.png)
