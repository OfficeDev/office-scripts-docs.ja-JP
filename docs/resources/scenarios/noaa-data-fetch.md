---
title: 'Office スクリプトのサンプル シナリオ: NOAA からの水位データのグラフ'
description: NOAA データベースから JSON データをフェッチし、それを使用してグラフを作成するサンプル。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: ba4836cd0782ab7f2158aeaaa562c851927b90f7
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755120"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Officeスクリプトのサンプル シナリオ: NOAA からの水レベルデータのフェッチとグラフ化

このシナリオでは、国立海洋大気局のシアトルステーションで水位を [プロットする必要があります](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)。 外部データを使用してスプレッドシートにデータを入力し、グラフを作成します。

このコマンドを使用して NOAA Tides データベースと Currents データベースにクエリを実行するスクリプト `fetch` [を開発します](https://tidesandcurrents.noaa.gov/)。 この場合、指定した期間にわたって水位が記録されます。 情報は JSON として返されます。そのため、スクリプトの一部は、それを範囲の値に変換します。 データがスプレッドシートに入った後、グラフの作成に使用されます。

## <a name="scripting-skills-covered"></a>スクリプティングのスキルをカバー

- 外部 API 呼び出し ( `fetch` )
- JSON の解析
- グラフ

## <a name="setup-instructions"></a>セットアップ手順

1. Web 上の Excel を使用してブックを開きます。

1. [自動化] **タブで** 、[すべてのスクリプト] **を選択します**。

1. [コード **エディター] 作業ウィンドウ** で、[ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

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

1. スクリプトの名前を **NOAA 水位グラフに変更し** 、保存します。

## <a name="running-the-script"></a>スクリプトを実行する

任意のワークシートで **、NOAA 水位グラフ スクリプトを実行** します。 このスクリプトは、2020 年 12 月 25 日から 2020 年 12 月 27 日まで水位データをフェッチします。 スクリプトの先頭にある変数は、異なる日付を使用したり、異なるステーション情報を取得したりするために `const` 変更できます。 [CO-OPS API For Data Retrieval では](https://api.tidesandcurrents.noaa.gov/api/prod/)、このすべてのデータを取得する方法について説明します。

### <a name="after-running-the-script"></a>スクリプトの実行後

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="スクリプトを実行した後のワークシートには、水位データとグラフが表示されます。":::
