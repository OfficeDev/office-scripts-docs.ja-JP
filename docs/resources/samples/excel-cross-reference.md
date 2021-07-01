---
title: ファイルとファイルExcel相互参照Power Automate
description: スクリプトとスクリプトを使用Office、Power Automateファイルを相互参照して書式設定するExcelします。
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202876"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="7e6ca-103">ファイルとファイルExcel相互参照Power Automate</span><span class="sxs-lookup"><span data-stu-id="7e6ca-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="7e6ca-104">このソリューションでは、2 つのファイル間でデータを比較Excel不一致を見つける方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="7e6ca-105">このスクリプトはOfficeを使用してデータを分析し、Power Automate間の通信を行います。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="7e6ca-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="7e6ca-106">Example scenario</span></span>

<span data-ttu-id="7e6ca-107">今後の会議にスピーカーをスケジュールしているイベント コーディネーターです。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="7e6ca-108">イベント データは 1 つのスプレッドシートに、スピーカーの登録は別のスプレッドシートに保持します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="7e6ca-109">2 つのブックの同期を確実に行う場合は、Officeスクリプトを使用して、潜在的な問題を強調表示します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="7e6ca-110">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="7e6ca-110">Sample Excel files</span></span>

<span data-ttu-id="7e6ca-111">このソリューションで使用されている次のファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="7e6ca-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7e6ca-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="7e6ca-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7e6ca-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="7e6ca-114">サンプル コード: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="7e6ca-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="7e6ca-115">サンプル コード: スピーカー登録の検証</span><span class="sxs-lookup"><span data-stu-id="7e6ca-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="7e6ca-116">Power Automateフロー: ブック全体の不整合を確認する</span><span class="sxs-lookup"><span data-stu-id="7e6ca-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="7e6ca-117">このフローは、最初のブックからイベント情報を抽出し、そのデータを使用して 2 番目のブックを検証します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="7e6ca-118">新しいインスタント [Power Automate](https://flow.microsoft.com)にサインインし、**新しいインスタント クラウド フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="7e6ca-119">[フロー **を手動でトリガーする] を選択し** 、[作成] を **押します**。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="7e6ca-120">[スクリプト **の実行]** アクションを使用して、Excel **(Business)** コネクタを使用する新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7e6ca-121">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="7e6ca-122">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="7e6ca-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="7e6ca-123">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="7e6ca-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="7e6ca-124">**ファイル**: event-data.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="7e6ca-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="7e6ca-125">**スクリプト**: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="7e6ca-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="最初のスクリプトExcelオンライン (Business) コネクタの完成Power Automate。":::

1. <span data-ttu-id="7e6ca-127">[スクリプトの実行 **] アクション** を使用して、Excel **(Business)** コネクタを使用する 2 番目の新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7e6ca-128">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="7e6ca-129">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="7e6ca-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="7e6ca-130">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="7e6ca-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="7e6ca-131">**ファイル**: speaker-registration.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="7e6ca-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="7e6ca-132">**スクリプト**: スピーカー登録の検証</span><span class="sxs-lookup"><span data-stu-id="7e6ca-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="2 番目Excelのオンライン (Business) コネクタの完成Power Automate。":::
1. <span data-ttu-id="7e6ca-134">このサンプルでは、Outlookクライアントとして使用します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="7e6ca-135">サポートされている任意の電子メール コネクタPower Automate使用できます。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="7e6ca-136">新しい **手順を追加** して、Office 365 Outlook **および電子** メール **(V2) アクションを使用** します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="7e6ca-137">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="7e6ca-138">**To**: テスト用メール アカウント (または個人用メール)</span><span class="sxs-lookup"><span data-stu-id="7e6ca-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="7e6ca-139">**件名**: イベントの検証結果</span><span class="sxs-lookup"><span data-stu-id="7e6ca-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="7e6ca-140">**本文**: result (_Run スクリプト 2 からの **動的コンテンツ**_)</span><span class="sxs-lookup"><span data-stu-id="7e6ca-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Office 365 OutlookでPower Automate。":::
1. <span data-ttu-id="7e6ca-142">フローを保存し、[テスト] **を選択** して試します。"不一致が見つかりました" というメールを受信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="7e6ca-143">データにはレビューが必要です。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-143">Data requires your review."</span></span> <span data-ttu-id="7e6ca-144">これは、グループ内の行と **speaker-registrations.xlsx行の** 間に違 **いevent-data.xlsx。**</span><span class="sxs-lookup"><span data-stu-id="7e6ca-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="7e6ca-145">[speaker-registrations.xlsxを **開** き、スピーカー登録リストに潜在的な問題があるいくつかの強調表示されたセルを表示します。</span><span class="sxs-lookup"><span data-stu-id="7e6ca-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
