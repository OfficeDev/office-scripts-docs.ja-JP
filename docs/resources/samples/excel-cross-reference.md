---
title: ファイルを相互参照してExcelする
description: スクリプトとスクリプトを使用Office、Power Automateファイルを相互参照して書式設定するExcelします。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 7cc10787190e7ba8f5984ddda8b3c770eb0f7d8a
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285907"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="1c24a-103">ファイルを相互参照してExcelする</span><span class="sxs-lookup"><span data-stu-id="1c24a-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="1c24a-104">このソリューションは、2 つの Excel ファイルを相互参照および書式設定する方法を、スクリプトとスクリプトを使用Office示Power Automate。</span><span class="sxs-lookup"><span data-stu-id="1c24a-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="1c24a-105">プロジェクトでは、次の結果が得されます。</span><span class="sxs-lookup"><span data-stu-id="1c24a-105">The project achieves the following:</span></span>

1. <span data-ttu-id="1c24a-106">1 つのスクリプトの <a href="events.xlsx"> 実行アクションevents.xlsx</a> を使用して、イベント データを抽出します。</span><span class="sxs-lookup"><span data-stu-id="1c24a-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="1c24a-107">そのデータをイベント トランザクション データを含む 2 番目の Excel ファイルに渡し、そのデータを使用して、Office Scripts を使用して、データの基本的な検証と、不足しているデータまたは不正確なデータの書式設定を行います。</span><span class="sxs-lookup"><span data-stu-id="1c24a-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="1c24a-108">結果をレビュー者に電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="1c24a-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="1c24a-109">詳細については、「クロス リファレンス」[を参照し、](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)スクリプトを使用して 2 Excel ファイルOfficeしてください。</span><span class="sxs-lookup"><span data-stu-id="1c24a-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="1c24a-110">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="1c24a-110">Sample Excel files</span></span>

<span data-ttu-id="1c24a-111">このソリューションで使用されている次のファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="1c24a-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="1c24a-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="1c24a-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="1c24a-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="1c24a-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="1c24a-114">サンプル コード: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="1c24a-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];
  
  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
      let [event, date, location, capacity] = row;
      records.push({
          event: event as string,
          date: date as number, 
          location: location as string,
          capacity: capacity as number
      })
  }

  // Log the event data to the console and return it for a flow.
  console.log(JSON.stringify(records));
  return records;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="1c24a-115">サンプル コード: イベント トランザクションの検証</span><span class="sxs-lookup"><span data-stu-id="1c24a-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);
    
 // Apply some basic formatting for readability.
  table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
  table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let overallMatch = true;

  // Iterate over every row.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [event, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyObject of keysObject) {
      if (keyObject.event === event) {
        match = true;

        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (keyObject.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.capacity !== capacity) {
          overallMatch = false;
          range.getCell(i, 3).getFormat()
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
  }
  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="1c24a-116">トレーニング ビデオ: クロスリファレンスと書式設定を行Excelファイル</span><span class="sxs-lookup"><span data-stu-id="1c24a-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="1c24a-117">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="1c24a-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
