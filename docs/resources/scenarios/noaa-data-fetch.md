---
title: 'Office スクリプトのサンプル シナリオ: NOAA から水レベルのデータをGraphする'
description: NOAA データベースから JSON データをフェッチし、それを使用してグラフを作成するサンプル。
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4181edae7d8a46ae381ddfb1a2893b03faffd9b
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088100"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office スクリプトのサンプル シナリオ: NOAA から水レベルのデータをフェッチしてグラフ化する

このシナリオでは、 [米国海洋大気局のシアトル局](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)で水位をプロットする必要があります。 外部データを使用してスプレッドシートを設定し、グラフを作成します。

このコマンドを使用 `fetch` して [NOAA Tides および Currents データベース](https://tidesandcurrents.noaa.gov/)に対するクエリを実行するスクリプトを開発します。 これは、特定の期間にわたって記録される水位を取得します。 情報は [JSON](https://www.w3schools.com/whatis/whatis_json.asp) として返されるため、スクリプトの一部では、その情報が範囲の値に変換されます。 データがスプレッドシートに含まれると、グラフを作成するために使用されます。

JSON の操作の詳細については、「[JSON を使用して、Office スクリプトとの間でデータを渡す](../../develop/use-json.md)」を参照してください。

## <a name="scripting-skills-covered"></a>スクリプティング スキルの説明

- 外部 API 呼び出し (`fetch`)
- JSON 解析
- グラフ

## <a name="setup-instructions"></a>セットアップ手順

1. Excel on the webを使用してブックを開きます。

1. [ **自動化** ] タブで [ **新しいスクリプト** ] を選択し、次のスクリプトをエディターに貼り付けます。

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook) {
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
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
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
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. スクリプトの名前を **NOAA 水位グラフ** に変更して保存します。

## <a name="running-the-script"></a>スクリプトを実行する

任意のワークシートで、 **NOAA 水位チャート** スクリプトを実行します。 このスクリプトは、2020 年 12 月 25 日から 2020 年 12 月 27 日までの水位データをフェッチします。 スクリプトの先頭の変数を `const` 変更して、異なる日付を使用したり、異なるステーション情報を取得したりできます。 [データ取得用 CO-OPS API](https://api.tidesandcurrents.noaa.gov/api/prod/) では、このすべてのデータを取得する方法について説明します。

### <a name="after-running-the-script"></a>スクリプトを実行した後

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="スクリプトを実行した後のワークシートには、水位データとグラフが表示されます。":::
