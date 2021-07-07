---
title: 'Officeスクリプトのサンプル シナリオ: Web ダウンロードの分析'
description: ブック内の生のインターネット トラフィック データをExcel、その情報をテーブルに整理する前に、元の場所を決定するサンプル。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: ec7ccd2a4f534be825ee4aa358f5d81becc53937
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313933"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Officeスクリプトのサンプル シナリオ: Web ダウンロードの分析

このシナリオでは、会社の Web サイトからダウンロード レポートを分析する必要があります。 この分析の目的は、Web トラフィックが米国または世界の他の場所から送信されるかどうかを判断します。

同僚が生データをブックにアップロードします。 毎週のデータ セットには、独自のワークシートがあります。 また、週 **別の傾向を** 示す表とグラフを含むサマリー ワークシートがあります。

アクティブなワークシートの毎週のダウンロード データを分析するスクリプトを開発します。 各ダウンロードに関連付けられた IP アドレスを解析し、それが米国から送信されたかどうかを判断します。 答えは、ブール値 ("TRUE" または "FALSE") としてワークシートに挿入され、条件付き書式がそれらのセルに適用されます。 IP アドレスの場所の結果はワークシートで合計され、サマリー テーブルにコピーされます。

## <a name="scripting-skills-covered"></a>スクリプティングのスキルをカバー

- テキスト解析
- スクリプトのサブ機能
- 条件付き書式
- テーブル

## <a name="setup-instructions"></a>セットアップ手順

1. ユーザー <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a>にダウンロードOneDrive。

1. ブックを開き、Excel for the web。

1. [自動化] **タブで** 、[新しい **スクリプト] を選択** し、次のスクリプトをエディターに貼り付けます。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
        * End the script early if either object is not in the workbook.
        */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }
  
      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();
  
      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
            "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1);
  
      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);
  
      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];
  
      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }
  
      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];
  
      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);
  
      // Call the local function to add summary data to the worksheet.
      addSummaryData();
  
      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);
  
      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();
  
      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();
  
      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;
  
      /**
       * A function to add summary data on the worksheet.
        */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
          formula1: "=TRUE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
          formula1: "=FALSE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
      // Split the IP address into octets.
      let octets = ipAddress.split(".");
  
      // Create a number for each octet and do the math to create the integer value of the IP address.
      let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
      return fullNum;
    }
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

1. スクリプトの名前を **[Web ダウンロードの分析] に変更し** 、保存します。

## <a name="running-the-script"></a>スクリプトを実行する

[週] ワークシート **に \* \* 移動** し、[Web ダウンロードの分析]**スクリプトを実行** します。 スクリプトは、現在のシートに条件付き書式と場所のラベル付けを適用します。 また、[概要] ワークシート **も更新** されます。

### <a name="before-running-the-script"></a>スクリプトを実行する前に

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="生の Web トラフィック データを表示するワークシート。":::

### <a name="after-running-the-script"></a>スクリプトの実行後

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="前の Web トラフィック行で書式設定された IP 場所情報を表示するワークシート。":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="スクリプトが実行されたワークシートをまとめたサマリー テーブルとグラフ。":::
