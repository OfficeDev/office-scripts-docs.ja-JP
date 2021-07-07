---
title: ファイルとファイルExcel相互参照Power Automate
description: スクリプトとスクリプトを使用Office、Power Automateファイルを相互参照して書式設定するExcelします。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313961"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="5d28a-103">ファイルとファイルExcel相互参照Power Automate</span><span class="sxs-lookup"><span data-stu-id="5d28a-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="5d28a-104">このソリューションでは、2 つのファイル間でデータを比較Excel不一致を見つける方法を示します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="5d28a-105">このスクリプトはOfficeを使用してデータを分析し、Power Automate間の通信を行います。</span><span class="sxs-lookup"><span data-stu-id="5d28a-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="5d28a-106">シナリオ例</span><span class="sxs-lookup"><span data-stu-id="5d28a-106">Example scenario</span></span>

<span data-ttu-id="5d28a-107">今後の会議にスピーカーをスケジュールしているイベント コーディネーターです。</span><span class="sxs-lookup"><span data-stu-id="5d28a-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="5d28a-108">イベント データは 1 つのスプレッドシートに、スピーカーの登録は別のスプレッドシートに保持します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="5d28a-109">2 つのブックの同期を確実に行う場合は、Officeスクリプトを使用して、潜在的な問題を強調表示します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="5d28a-110">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="5d28a-110">Sample Excel files</span></span>

<span data-ttu-id="5d28a-111">次のファイルをダウンロードして、サンプルのすぐに使用できるブックを取得します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="5d28a-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="5d28a-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="5d28a-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="5d28a-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="5d28a-114">次のスクリプトを追加して、サンプルを自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="5d28a-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="5d28a-115">サンプル コード: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="5d28a-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="5d28a-116">サンプル コード: スピーカー登録の検証</span><span class="sxs-lookup"><span data-stu-id="5d28a-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="5d28a-117">Power Automateフロー: ブック全体の不整合を確認する</span><span class="sxs-lookup"><span data-stu-id="5d28a-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="5d28a-118">このフローは、最初のブックからイベント情報を抽出し、そのデータを使用して 2 番目のブックを検証します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="5d28a-119">新しいインスタント [Power Automate](https://flow.microsoft.com)にサインインし、**新しいインスタント クラウド フローを作成します**。</span><span class="sxs-lookup"><span data-stu-id="5d28a-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="5d28a-120">[フロー **を手動でトリガーする] を選択し** 、[作成] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="5d28a-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="5d28a-121">[スクリプト **の実行]** アクションを使用して、Excel **(Business)** コネクタを使用する新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5d28a-122">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="5d28a-123">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5d28a-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="5d28a-124">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5d28a-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="5d28a-125">**ファイル**: event-data.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="5d28a-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="5d28a-126">**スクリプト**: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="5d28a-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="最初のスクリプトExcelオンライン (Business) コネクタの完成Power Automate。":::

1. <span data-ttu-id="5d28a-128">[スクリプトの実行 **] アクション** を使用して、Excel **(Business)** コネクタを使用する 2 番目の新しい **手順を追加** します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5d28a-129">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="5d28a-130">**場所**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5d28a-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="5d28a-131">**ドキュメント ライブラリ**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="5d28a-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="5d28a-132">**ファイル**: speaker-registration.xlsx ([ファイル選択で選択)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="5d28a-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="5d28a-133">**スクリプト**: スピーカー登録の検証</span><span class="sxs-lookup"><span data-stu-id="5d28a-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="2 番目Excelのオンライン (Business) コネクタの完成Power Automate。":::
1. <span data-ttu-id="5d28a-135">このサンプルでは、Outlookクライアントとして使用します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="5d28a-136">サポートされている任意の電子メール コネクタPower Automate使用できます。</span><span class="sxs-lookup"><span data-stu-id="5d28a-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="5d28a-137">新しい **手順を追加** して、Office 365 Outlook **および電子** メール **(V2) アクションを使用** します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="5d28a-138">アクションには、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="5d28a-139">**To**: テスト用メール アカウント (または個人用メール)</span><span class="sxs-lookup"><span data-stu-id="5d28a-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="5d28a-140">**件名**: イベントの検証結果</span><span class="sxs-lookup"><span data-stu-id="5d28a-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="5d28a-141">**本文**: result (_Run スクリプト 2 からの **動的コンテンツ**_)</span><span class="sxs-lookup"><span data-stu-id="5d28a-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Office 365 OutlookでPower Automate。":::
1. <span data-ttu-id="5d28a-143">フローを保存します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-143">Save the flow.</span></span> <span data-ttu-id="5d28a-144">[フロー エディター **] ページ** の [テスト] ボタンを使用するか、[マイ フロー] タブでフロー **を実行** します。メッセージが表示されたら、必ずアクセスを許可してください。</span><span class="sxs-lookup"><span data-stu-id="5d28a-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="5d28a-145">"不一致が見つかりました" というメールを受信する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5d28a-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="5d28a-146">データにはレビューが必要です。</span><span class="sxs-lookup"><span data-stu-id="5d28a-146">Data requires your review."</span></span> <span data-ttu-id="5d28a-147">これは、グループ内の行と **speaker-registrations.xlsx行の** 間に違 **いevent-data.xlsx。**</span><span class="sxs-lookup"><span data-stu-id="5d28a-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="5d28a-148">[speaker-registrations.xlsxを **開** き、スピーカー登録リストに潜在的な問題があるいくつかの強調表示されたセルを表示します。</span><span class="sxs-lookup"><span data-stu-id="5d28a-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
