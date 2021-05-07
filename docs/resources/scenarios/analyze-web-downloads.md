---
title: 'Officeスクリプトのサンプル シナリオ: Web ダウンロードの分析'
description: ブック内の生のインターネット トラフィック データをExcel、その情報をテーブルに整理する前に、元の場所を決定するサンプル。
ms.date: 04/27/2021
localization_priority: Normal
ms.openlocfilehash: 6c5958e9957ca49c370ae34456236bdd15f41c44
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232712"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="8b549-103">Officeスクリプトのサンプル シナリオ: Web ダウンロードの分析</span><span class="sxs-lookup"><span data-stu-id="8b549-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="8b549-104">このシナリオでは、会社の Web サイトからダウンロード レポートを分析する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8b549-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="8b549-105">この分析の目的は、Web トラフィックが米国または世界の他の場所から送信されるかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="8b549-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="8b549-106">同僚が生データをブックにアップロードします。</span><span class="sxs-lookup"><span data-stu-id="8b549-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="8b549-107">毎週のデータ セットには、独自のワークシートがあります。</span><span class="sxs-lookup"><span data-stu-id="8b549-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="8b549-108">また、週 **別の傾向を** 示す表とグラフを含むサマリー ワークシートがあります。</span><span class="sxs-lookup"><span data-stu-id="8b549-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="8b549-109">アクティブなワークシートの毎週のダウンロード データを分析するスクリプトを開発します。</span><span class="sxs-lookup"><span data-stu-id="8b549-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="8b549-110">各ダウンロードに関連付けられた IP アドレスを解析し、それが米国から送信されたかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="8b549-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="8b549-111">答えは、ブール値 ("TRUE" または "FALSE") としてワークシートに挿入され、条件付き書式がそれらのセルに適用されます。</span><span class="sxs-lookup"><span data-stu-id="8b549-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="8b549-112">IP アドレスの場所の結果はワークシートで合計され、サマリー テーブルにコピーされます。</span><span class="sxs-lookup"><span data-stu-id="8b549-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="8b549-113">スクリプティングのスキルをカバー</span><span class="sxs-lookup"><span data-stu-id="8b549-113">Scripting skills covered</span></span>

- <span data-ttu-id="8b549-114">テキスト解析</span><span class="sxs-lookup"><span data-stu-id="8b549-114">Text parsing</span></span>
- <span data-ttu-id="8b549-115">スクリプトのサブ機能</span><span class="sxs-lookup"><span data-stu-id="8b549-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="8b549-116">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="8b549-116">Conditional formatting</span></span>
- <span data-ttu-id="8b549-117">テーブル</span><span class="sxs-lookup"><span data-stu-id="8b549-117">Tables</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="8b549-118">セットアップ手順</span><span class="sxs-lookup"><span data-stu-id="8b549-118">Setup instructions</span></span>

1. <span data-ttu-id="8b549-119">ユーザー <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a>にダウンロードOneDrive。</span><span class="sxs-lookup"><span data-stu-id="8b549-119">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="8b549-120">Web 用のブックExcel開きます。</span><span class="sxs-lookup"><span data-stu-id="8b549-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="8b549-121">[自動化] **タブで** 、[すべてのスクリプト] **を開きます**。</span><span class="sxs-lookup"><span data-stu-id="8b549-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="8b549-122">[コード **エディター] 作業ウィンドウ** で、[新しいスクリプト] **を押** して、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="8b549-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="8b549-123">スクリプトの名前を **[Web ダウンロードの分析] に変更し** 、保存します。</span><span class="sxs-lookup"><span data-stu-id="8b549-123">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="8b549-124">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="8b549-124">Running the script</span></span>

<span data-ttu-id="8b549-125">[週] ワークシート **に \* \* 移動** し、[Web ダウンロードの分析]**スクリプトを実行** します。</span><span class="sxs-lookup"><span data-stu-id="8b549-125">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="8b549-126">スクリプトは、現在のシートに条件付き書式と場所のラベル付けを適用します。</span><span class="sxs-lookup"><span data-stu-id="8b549-126">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="8b549-127">また、[概要] ワークシート **も更新** されます。</span><span class="sxs-lookup"><span data-stu-id="8b549-127">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="8b549-128">スクリプトを実行する前に</span><span class="sxs-lookup"><span data-stu-id="8b549-128">Before running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="生の Web トラフィック データを表示するワークシート":::

### <a name="after-running-the-script"></a><span data-ttu-id="8b549-130">スクリプトの実行後</span><span class="sxs-lookup"><span data-stu-id="8b549-130">After running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="前の Web トラフィック行で書式設定された IP 場所情報を表示するワークシート":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="スクリプトが実行されたワークシートをまとめたサマリー テーブルとグラフ":::
