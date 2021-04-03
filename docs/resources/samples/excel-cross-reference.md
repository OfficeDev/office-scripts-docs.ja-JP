---
title: Excel ファイルの相互参照と書式設定
description: Excel ファイルを相互参照Office書式設定するには、スクリプトと Power Automate を使用する方法について説明します。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 287de604733b7e6a126d0c81cb4e23351e558c61
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571513"
---
# <a name="cross-reference-and-format-an-excel-file"></a>Excel ファイルの相互参照と書式設定

このソリューションは、スクリプトと Power Automate を使用して 2 つの Excel ファイルを相互参照および書式設定Office示します。

プロジェクトでは、次の結果が得されます。

1. 1 つのスクリプトの <a href="events.xlsx"> 実行アクションevents.xlsx</a> を使用して、イベント データを抽出します。
1. イベント トランザクション データを含む 2 番目の Excel ファイルにデータを渡し、そのデータを使用してデータの基本的な検証を行い、Office スクリプトを使用してデータの不足または誤ったデータの書式設定を行います。
1. 結果をレビュー者に電子メールで送信します。

詳細については、「クロス リファレンス」を参照し、スクリプトを使用して 2 つの [Excel ファイルOfficeしてください](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)。

## <a name="sample-excel-files"></a>サンプル Excel ファイル

このソリューションで使用されている次のファイルをダウンロードして、自分で試してみてください。

1. <a href="events.xlsx">events.xlsx</a>
1. <a href="event-transactions.xlsx">event-transactions.xlsx</a>

## <a name="sample-code-get-event-data"></a>サンプル コード: イベント データの取得

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
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
    console.log(JSON.stringify(records))
    return records;
}

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
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      for (let keyObject of keysObject){
        if (keyObject.event === event) {
          match = true;
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
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
}

interface EventData {
event: string
date: number
location: string
capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a>トレーニング ビデオ: Excel ファイルの相互参照と書式設定

[![Excel ファイルを相互参照して書式設定する方法の詳細なビデオを見る](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Excel ファイルを相互参照して書式設定する方法の詳細なビデオ")
