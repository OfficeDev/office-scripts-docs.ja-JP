---
title: 'Office スクリプトサンプルシナリオ: web ダウンロードを分析する'
description: Excel ブックでインターネットトラフィックを生に取得し、その情報をテーブルに整理する前に元の場所を特定するサンプル。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 425d2af432d6b3c4b7604daf7935d2cc1ec059a8
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999268"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="12348-103">Office スクリプトサンプルシナリオ: web ダウンロードを分析する</span><span class="sxs-lookup"><span data-stu-id="12348-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="12348-104">このシナリオでは、企業の web サイトからダウンロードレポートを分析する作業を行っています。</span><span class="sxs-lookup"><span data-stu-id="12348-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="12348-105">この分析の目的は、web トラフィックが米国または世界の他の場所から送られてくるかどうかを判断することです。</span><span class="sxs-lookup"><span data-stu-id="12348-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="12348-106">仕事仲間がブックに生データをアップロードします。</span><span class="sxs-lookup"><span data-stu-id="12348-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="12348-107">各週のデータセットには、独自のワークシートがあります。</span><span class="sxs-lookup"><span data-stu-id="12348-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="12348-108">また、週単位の傾向を示す表とグラフを含む**サマリー**ワークシートもあります。</span><span class="sxs-lookup"><span data-stu-id="12348-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="12348-109">作業中のワークシートのデータを週単位でダウンロードするスクリプトを開発します。</span><span class="sxs-lookup"><span data-stu-id="12348-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="12348-110">各ダウンロードに関連付けられている IP アドレスを解析し、それが US からのものであるかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="12348-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="12348-111">答えは、ワークシートのブール値 ("TRUE" または "FALSE") として挿入され、それらのセルに条件付き書式が適用されます。</span><span class="sxs-lookup"><span data-stu-id="12348-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="12348-112">IP アドレスの場所の結果がワークシートに合計され、サマリーテーブルにコピーされます。</span><span class="sxs-lookup"><span data-stu-id="12348-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="12348-113">スクリプト作成スキルの説明</span><span class="sxs-lookup"><span data-stu-id="12348-113">Scripting skills covered</span></span>

- <span data-ttu-id="12348-114">テキストの解析</span><span class="sxs-lookup"><span data-stu-id="12348-114">Text parsing</span></span>
- <span data-ttu-id="12348-115">スクリプト内の subfunctions</span><span class="sxs-lookup"><span data-stu-id="12348-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="12348-116">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="12348-116">Conditional formatting</span></span>
- <span data-ttu-id="12348-117">テーブル</span><span class="sxs-lookup"><span data-stu-id="12348-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="12348-118">デモビデオ</span><span class="sxs-lookup"><span data-stu-id="12348-118">Demo video</span></span>

<span data-ttu-id="12348-119">このサンプルは、2020年2月に、Office アドインの開発者コミュニティコールの一部として使用されています。</span><span class="sxs-lookup"><span data-stu-id="12348-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> <span data-ttu-id="12348-120">このビデオに示されているコードは、古い API モデル ( [Office スクリプト非同期 api](../../develop/excel-async-model.md)) を使用しています。</span><span class="sxs-lookup"><span data-stu-id="12348-120">The code shown in this video uses an older API model (the [Office Scripts Async APIs](../../develop/excel-async-model.md)).</span></span> <span data-ttu-id="12348-121">このページに示されているサンプルは更新されていますが、コードは録音とは少し異なります。</span><span class="sxs-lookup"><span data-stu-id="12348-121">The sample presented on this page has been updated, but the code looks a little different from the recording.</span></span> <span data-ttu-id="12348-122">この変更は、プレゼンターのデモのスクリプトまたはその他のコンテンツの動作には影響しません。</span><span class="sxs-lookup"><span data-stu-id="12348-122">The changes don't affect the behavior of the script or the other content in the presenter's demo.</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="12348-123">セットアップの手順</span><span class="sxs-lookup"><span data-stu-id="12348-123">Setup instructions</span></span>

1. <span data-ttu-id="12348-124">OneDrive に<a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a>をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="12348-124">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="12348-125">Web 用の Excel でブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="12348-125">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="12348-126">[**自動化**] タブで、**コードエディター**を開きます。</span><span class="sxs-lookup"><span data-stu-id="12348-126">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="12348-127">[**コードエディター** ] 作業ウィンドウで、[**新しいスクリプト**] をクリックし、次のスクリプトをエディターに貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="12348-127">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="12348-128">スクリプトの名前を変更して**Web ダウンロードを分析**し、保存します。</span><span class="sxs-lookup"><span data-stu-id="12348-128">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="12348-129">スクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="12348-129">Running the script</span></span>

<span data-ttu-id="12348-130">任意の\*\* \* \* 週**のワークシートに移動し、[Web 用の**ダウンロードの分析\*\*] スクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="12348-130">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="12348-131">このスクリプトは、現在のシートに条件付き書式と場所のラベルを適用します。</span><span class="sxs-lookup"><span data-stu-id="12348-131">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="12348-132">**サマリー**ワークシートも更新されます。</span><span class="sxs-lookup"><span data-stu-id="12348-132">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="12348-133">スクリプトを実行する前に</span><span class="sxs-lookup"><span data-stu-id="12348-133">Before running the script</span></span>

![生の web トラフィックデータを示すワークシート。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="12348-135">スクリプトを実行した後</span><span class="sxs-lookup"><span data-stu-id="12348-135">After running the script</span></span>

![以前の web トラフィック行で書式設定された IP 位置情報を示すワークシート。](../../images/scenario-analyze-web-downloads-after.png)

![スクリプトが実行されたワークシートの要約を示す表とグラフ。](../../images/scenario-analyze-web-downloads-table.png)
