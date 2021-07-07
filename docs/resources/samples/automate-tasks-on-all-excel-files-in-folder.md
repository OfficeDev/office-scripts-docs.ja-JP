---
title: フォルダー内のすべての Excel ファイルでスクリプトを実行する
description: フォルダー内のすべてのファイルに対してスクリプトExcel実行する方法について説明OneDrive for Business。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313898"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="c47a2-103">フォルダー内のすべての Excel ファイルでスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="c47a2-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="c47a2-104">このプロジェクトは、フォルダー内のすべてのファイルに対して一連の自動化タスクを実行OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="c47a2-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="c47a2-105">また、フォルダー内のフォルダー SharePointすることもできます。</span><span class="sxs-lookup"><span data-stu-id="c47a2-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="c47a2-106">このプロパティは、Excelファイルに対して計算を実行し、書式設定を追加し、同僚にコメント[@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)挿入します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="c47a2-107">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="c47a2-107">Sample Excel files</span></span>

<span data-ttu-id="c47a2-108">この <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> 必要なすべてのブックの詳細をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="c47a2-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="c47a2-109">これらのファイルを Sales というタイトルのフォルダーに **展開します**。</span><span class="sxs-lookup"><span data-stu-id="c47a2-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="c47a2-110">次のスクリプトをスクリプト コレクションに追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="c47a2-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="c47a2-111">サンプル コード: 書式の追加とコメントの挿入</span><span class="sxs-lookup"><span data-stu-id="c47a2-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="c47a2-112">これは、個々のブックで実行されるスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="c47a2-112">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="c47a2-113">Power Automateフロー: フォルダー内のすべてのブックでスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="c47a2-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="c47a2-114">このフローは、"Sales" フォルダー内のすべてのブックでスクリプトを実行します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="c47a2-115">新しいインスタント クラウド **フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="c47a2-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="c47a2-116">[フロー **を手動でトリガーする] を選択し** 、[作成] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="c47a2-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="c47a2-117">[フォルダー内 **のファイルの一** 覧] **OneDrive for Businessを使用** する新 **しい手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="完了したOneDrive for BusinessコネクタをPower Automate。":::
1. <span data-ttu-id="c47a2-119">抽出されたブックを含む "Sales" フォルダーを選択します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="c47a2-120">ブックのみを選択するには、[新しい手順] を選択し、[条件]**を選択\*\*\*\*し**、次の値を設定します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="c47a2-121">**名前**(ファイルOneDrive値)</span><span class="sxs-lookup"><span data-stu-id="c47a2-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="c47a2-122">"ends with"</span><span class="sxs-lookup"><span data-stu-id="c47a2-122">"ends with"</span></span>
    1. <span data-ttu-id="c47a2-123">"xlsx"</span><span class="sxs-lookup"><span data-stu-id="c47a2-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="後続Power Automateを各ファイルに適用する条件ブロックを指定します。":::
1. <span data-ttu-id="c47a2-125">[**はい] ブランチで**、[スクリプトの実行] アクションExcel **オンライン (Business)** コネクタ **を追加** します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="c47a2-126">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="c47a2-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="c47a2-127">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="c47a2-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="c47a2-128">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c47a2-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="c47a2-129">**ファイル**: **Id** (OneDrive ID 値)</span><span class="sxs-lookup"><span data-stu-id="c47a2-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="c47a2-130">**スクリプト**: スクリプト名</span><span class="sxs-lookup"><span data-stu-id="c47a2-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="オンライン (Excel) コネクタの完成Power Automate。":::
1. <span data-ttu-id="c47a2-132">フローを保存し、試してみてください。[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。</span><span class="sxs-lookup"><span data-stu-id="c47a2-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="c47a2-133">トレーニング ビデオ: フォルダー内のすべてのファイルExcelスクリプトを実行する</span><span class="sxs-lookup"><span data-stu-id="c47a2-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="c47a2-134">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="c47a2-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
