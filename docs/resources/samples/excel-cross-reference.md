---
title: Excel ファイルの相互参照とフォーマット
description: OfficeスクリプトとPower Automateを使用して、Excel ファイルを相互参照およびフォーマットする方法について説明します。
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: f07395eb4e6c77b7aee3776e3252d135bc690a6f
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545767"
---
# <a name="cross-reference-and-format-an-excel-file"></a>Excel ファイルの相互参照とフォーマット

このソリューションでは、OfficeスクリプトとPower Automateを使用して、2 つのExcel ファイルを相互参照およびフォーマットする方法を示します。

プロジェクトは、次のことが実現します。

1. 1 つのスクリプト実行アクションを使用して <a href="events.xlsx">events.xlsx</a> からイベント データを抽出します。
1. そのデータをイベント トランザクション データを含む 2 番目のExcel ファイルに渡し、そのデータを使用して、Office スクリプトを使用してデータの基本的な検証と、欠落または不正なデータの書式設定を行います。
1. 結果をレビュー担当者に電子メールで送信します。

詳細については、「[クロス リファレンス」および「 Office スクリプトを使用した 2 つのExcel ファイルのフォーマット](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)」を参照してください。

## <a name="sample-excel-files"></a>サンプル Excel ファイル

このソリューションで使用されている以下のファイルをダウンロードして、自分で試してみてください!

1. <a href="events.xlsx">events.xlsx</a>
1. <a href="event-transactions.xlsx">event-transactions.xlsx</a>

## <a name="sample-code-get-event-data"></a>サンプル コード: イベント データを取得する

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

## <a name="sample-code-validate-event-transactions"></a>サンプル コード: イベント トランザクションの検証

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a>トレーニング ビデオ: Excel ファイルの相互参照とフォーマット

[スーディ・ラマムルティがこのサンプルをYouTubeで歩くのを見てください](https://youtu.be/dVwqBf483qo")。
