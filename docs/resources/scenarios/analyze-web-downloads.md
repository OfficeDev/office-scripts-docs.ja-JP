---
title: 'Office スクリプトサンプルシナリオ: web ダウンロードを分析する'
description: Excel ブックでインターネットトラフィックを生に取得し、その情報をテーブルに整理する前に元の場所を特定するサンプル。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 2a74fada55115faf79f0b625b8a7cd6352deb651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878648"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Office スクリプトサンプルシナリオ: web ダウンロードを分析する

このシナリオでは、企業の web サイトからダウンロードレポートを分析する作業を行っています。 この分析の目的は、web トラフィックが米国または世界の他の場所から送られてくるかどうかを判断することです。

仕事仲間がブックに生データをアップロードします。 各週のデータセットには、独自のワークシートがあります。 また、週単位の傾向を示す表とグラフを含む**サマリー**ワークシートもあります。

作業中のワークシートのデータを週単位でダウンロードするスクリプトを開発します。 各ダウンロードに関連付けられている IP アドレスを解析し、それが US からのものであるかどうかを判断します。 答えは、ワークシートのブール値 ("TRUE" または "FALSE") として挿入され、それらのセルに条件付き書式が適用されます。 IP アドレスの場所の結果がワークシートに合計され、サマリーテーブルにコピーされます。

## <a name="scripting-skills-covered"></a>スクリプト作成スキルの説明

- テキストの解析
- スクリプト内の subfunctions
- 条件付き書式
- テーブル

## <a name="demo-video"></a>デモビデオ

このサンプルは、2020年2月に、Office アドインの開発者コミュニティコールの一部として使用されています。

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a>セットアップの手順

1. OneDrive に<a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a>をダウンロードします。

2. Web 用の Excel でブックを開きます。

3. [**自動化**] タブで、**コードエディター**を開きます。

4. [**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the Summary worksheet and table.
      let summaryWorksheet = workbook.getWorksheet("Summary");
      let summaryTable = summaryWorksheet?.getTable("Table1");
      if (!summaryWorksheet || !summaryTable) {
          console.log("The script expects the Summary worksheet with a summary table named Table1. Please download the correct template and try again.");
          return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (!currentWorksheet.getName().toLocaleLowerCase().startsWith("week")) {
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
      let ipRangeValues = ipRange.getValues();
      let logRangeValues = logRange.getValues();
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
        let summaryHeaderRow = currentWorksheet
            .getRange("J1:M1");
        let summaryContentRow = currentWorksheet
            .getRange("J2:M2");
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
        conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#F8CCAD");
        conditionalFormatTrue.getCellValue().setRule({
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

5. スクリプトの名前を変更して**Web ダウンロードを分析**し、保存します。

## <a name="running-the-script"></a>スクリプトを実行する

任意の** \* \* 週**のワークシートに移動し、[Web 用の**ダウンロードの分析**] スクリプトを実行します。 このスクリプトは、現在のシートに条件付き書式と場所のラベルを適用します。 **サマリー**ワークシートも更新されます。

### <a name="before-running-the-script"></a>スクリプトを実行する前に

![生の web トラフィックデータを示すワークシート。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>スクリプトを実行した後

![以前の web トラフィック行で書式設定された IP 位置情報を示すワークシート。](../../images/scenario-analyze-web-downloads-after.png)

![スクリプトが実行されたワークシートの要約を示す表とグラフ。](../../images/scenario-analyze-web-downloads-table.png)
