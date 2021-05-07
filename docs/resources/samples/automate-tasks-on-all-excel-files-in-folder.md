---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: フォルダー内のすべてのファイルに対してスクリプトExcel実行する方法について説明OneDrive for Business。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a6b869e2b346635e2b28fa7c6273c1a86a5bc5c5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232628"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="dd681-103">フォルダー内のすべての Excel ファイルでスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="dd681-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="dd681-104">このプロジェクトは、フォルダー内のすべてのファイルに対して一連の自動化タスクを実行OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="dd681-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="dd681-105">また、フォルダー内のフォルダー SharePointすることもできます。</span><span class="sxs-lookup"><span data-stu-id="dd681-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="dd681-106">このプロパティは、Excelファイルに対して計算を実行し、書式設定を追加し、同僚にコメント[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)挿入します。</span><span class="sxs-lookup"><span data-stu-id="dd681-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="dd681-107">ファイルをダウンロード<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>サンプルで使用されている Sales というタイトルのフォルダーにファイルを抽出し、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="dd681-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="dd681-108">サンプル コード: 書式の追加とコメントの挿入</span><span class="sxs-lookup"><span data-stu-id="dd681-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="dd681-109">これは、個々のブックで実行されるスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="dd681-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="dd681-110">Power Automateフロー: フォルダー内のすべてのブックでスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="dd681-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="dd681-111">このフローは、"Sales" フォルダー内のすべてのブックでスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="dd681-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="dd681-112">新しいインスタント クラウド **フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="dd681-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="dd681-113">[フロー **を手動でトリガーする] を選択し** 、[作成] を **押します**。</span><span class="sxs-lookup"><span data-stu-id="dd681-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="dd681-114">[フォルダー内 **のファイルの一** 覧] **OneDrive for Businessを使用** する新 **しい手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="dd681-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="OneDrive for Businessの完成したコネクタPower Automate":::
1. <span data-ttu-id="dd681-116">抽出されたブックを含む "Sales" フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="dd681-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="dd681-117">ブックのみを選択するには、[新しい手順] を選択し、[条件]**を選択\*\*\*\*し**、次の値を設定します。</span><span class="sxs-lookup"><span data-stu-id="dd681-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="dd681-118">**名前**(ファイルOneDrive値)</span><span class="sxs-lookup"><span data-stu-id="dd681-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="dd681-119">"ends with"</span><span class="sxs-lookup"><span data-stu-id="dd681-119">"ends with"</span></span>
    1. <span data-ttu-id="dd681-120">"xlsx"</span><span class="sxs-lookup"><span data-stu-id="dd681-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="各Power Automateに後続のアクションを適用する条件ブロック":::
1. <span data-ttu-id="dd681-122">[**はい] ブランチの** 下に、[スクリプトの実行 (プレビュー) アクションExcel **オンライン (Business)** コネクタ **を追加** します。</span><span class="sxs-lookup"><span data-stu-id="dd681-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="dd681-123">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="dd681-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="dd681-124">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="dd681-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="dd681-125">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="dd681-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="dd681-126">**ファイル**: **Id** (OneDrive ID 値)</span><span class="sxs-lookup"><span data-stu-id="dd681-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="dd681-127">**スクリプト**: スクリプト名</span><span class="sxs-lookup"><span data-stu-id="dd681-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="オンライン (Excel) コネクタの完成Power Automate":::
1. <span data-ttu-id="dd681-129">フローを保存し、試してみてください。</span><span class="sxs-lookup"><span data-stu-id="dd681-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="dd681-130">トレーニング ビデオ: フォルダー内のすべてのファイルExcelスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="dd681-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="dd681-131">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="dd681-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
