---
title: 'Office スクリプトサンプルシナリオ: web ダウンロードを分析する'
description: Excel ブックでインターネットトラフィックを生に取得し、その情報をテーブルに整理する前に元の場所を特定するサンプル。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700412"
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

1. <a href="analyze-web-downloads.xlsx">Analyze-web-downloads</a>を OneDrive にダウンロードします。

2. Web 用の Excel でブックを開きます。

3. [**自動化**] タブで、**コードエディター**を開きます。

4. [**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
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

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
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

任意の**\*週**のワークシートに移動し、[Web 用の**ダウンロードの分析**] スクリプトを実行します。 このスクリプトは、現在のシートに条件付き書式と場所のラベルを適用します。 **サマリー**ワークシートも更新されます。

### <a name="before-running-the-script"></a>スクリプトを実行する前に

![生の web トラフィックデータを示すワークシート。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>スクリプトを実行した後

![以前の web トラフィック行で書式設定された IP 位置情報を示すワークシート。](../../images/scenario-analyze-web-downloads-after.png)

![スクリプトが実行されたワークシートの要約を示す表とグラフ。](../../images/scenario-analyze-web-downloads-table.png)
