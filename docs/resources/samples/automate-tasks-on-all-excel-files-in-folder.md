---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: OneDrive for Businessのフォルダ内のすべてのExcel ファイルに対してスクリプトを実行する方法について説明します。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545793"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="4904a-103">フォルダー内のすべての Excel ファイルでスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="4904a-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="4904a-104">このプロジェクトは、OneDrive for Businessのフォルダにあるすべてのファイルに対して、一連のオートメーション タスクを実行します。</span><span class="sxs-lookup"><span data-stu-id="4904a-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="4904a-105">また、SharePointフォルダでも使用できます。</span><span class="sxs-lookup"><span data-stu-id="4904a-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="4904a-106">Excel ファイルに対して計算を実行し、書式を追加し、同僚[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)コメントを挿入します。</span><span class="sxs-lookup"><span data-stu-id="4904a-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="4904a-107"><a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">ファイルhighlight-alert-excel-files.zip</a>をダウンロードし、このサンプルで使用されている **Sales** というタイトルのフォルダにファイルを抽出し、自分で試してみてください!</span><span class="sxs-lookup"><span data-stu-id="4904a-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="4904a-108">サンプル コード: 書式を追加してコメントを挿入する</span><span class="sxs-lookup"><span data-stu-id="4904a-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="4904a-109">これは、個々のブックで実行されるスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="4904a-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="4904a-110">Power Automateフロー: フォルダー内のすべてのブックに対してスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="4904a-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="4904a-111">このフローは、"Sales" フォルダー内のすべてのブックでスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="4904a-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="4904a-112">新しい **インスタント クラウド フロー** を作成する :</span><span class="sxs-lookup"><span data-stu-id="4904a-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="4904a-113">[ **フローを手動でトリガーする] を** 選択し、[ **作成]** を押します。</span><span class="sxs-lookup"><span data-stu-id="4904a-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="4904a-114">**[OneDrive for Business** コネクタ] アクションと [**フォルダー内のファイルを一覧表示する]** アクションを使用する **新しい手順** を追加します。</span><span class="sxs-lookup"><span data-stu-id="4904a-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Power Automate の完成したOneDrive for Business コネクタ":::
1. <span data-ttu-id="4904a-116">抽出したワークブックを含む "Sales" フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="4904a-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="4904a-117">ブックのみが選択されていることを確認するには、[ **新しいステップ**] を選択し、[ **条件** ] を選択して次の値を設定します。</span><span class="sxs-lookup"><span data-stu-id="4904a-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="4904a-118">**名前**(OneDrive ファイル名の値)</span><span class="sxs-lookup"><span data-stu-id="4904a-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="4904a-119">「で終わる」</span><span class="sxs-lookup"><span data-stu-id="4904a-119">"ends with"</span></span>
    1. <span data-ttu-id="4904a-120">"xlsx"。</span><span class="sxs-lookup"><span data-stu-id="4904a-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="後続のアクションを各ファイルに適用するPower Automate条件ブロック":::
1. <span data-ttu-id="4904a-122">[**はいの場合**] の下で、[**スクリプトの実行**] アクションを使用して **Excelオンライン (ビジネス)** コネクタを追加します。</span><span class="sxs-lookup"><span data-stu-id="4904a-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="4904a-123">アクションには次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="4904a-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="4904a-124">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="4904a-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="4904a-125">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="4904a-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="4904a-126">**ファイル**: **ID** (OneDriveファイル ID 値)</span><span class="sxs-lookup"><span data-stu-id="4904a-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="4904a-127">**スクリプト**: スクリプト名</span><span class="sxs-lookup"><span data-stu-id="4904a-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Power Automateの完了Excelオンライン (ビジネス) コネクタ":::
1. <span data-ttu-id="4904a-129">フローを保存し、それを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="4904a-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="4904a-130">トレーニング ビデオ: フォルダー内のすべてのExcel ファイルに対してスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="4904a-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="4904a-131">[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/xMg711o7k6w)。</span><span class="sxs-lookup"><span data-stu-id="4904a-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
