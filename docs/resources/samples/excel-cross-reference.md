---
title: ファイルを相互参照してExcelする
description: スクリプトとスクリプトを使用Office、Power Automateファイルを相互参照して書式設定するExcelします。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 858fe561c1a82f471bc3c0f43d81e457fb02b627
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232383"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="0ea8d-103">ファイルを相互参照してExcelする</span><span class="sxs-lookup"><span data-stu-id="0ea8d-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="0ea8d-104">このソリューションは、2 つの Excel ファイルを相互参照および書式設定する方法を、スクリプトとスクリプトを使用Office示Power Automate。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="0ea8d-105">プロジェクトでは、次の結果が得されます。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-105">The project achieves the following:</span></span>

1. <span data-ttu-id="0ea8d-106">1 つのスクリプトの <a href="events.xlsx"> 実行アクションevents.xlsx</a> を使用して、イベント データを抽出します。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="0ea8d-107">そのデータをイベント トランザクション データを含む 2 番目の Excel ファイルに渡し、そのデータを使用して、Office Scripts を使用して、データの基本的な検証と、不足しているデータまたは不正確なデータの書式設定を行います。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="0ea8d-108">結果をレビュー者に電子メールで送信します。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="0ea8d-109">詳細については、「クロス リファレンス」[を参照し、](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)スクリプトを使用して 2 Excel ファイルOfficeしてください。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="0ea8d-110">サンプル Excel ファイル</span><span class="sxs-lookup"><span data-stu-id="0ea8d-110">Sample Excel files</span></span>

<span data-ttu-id="0ea8d-111">このソリューションで使用されている次のファイルをダウンロードして、自分で試してみてください。</span><span class="sxs-lookup"><span data-stu-id="0ea8d-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="0ea8d-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="0ea8d-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="0ea8d-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="0ea8d-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="0ea8d-114">サンプル コード: イベント データの取得</span><span class="sxs-lookup"><span data-stu-id="0ea8d-114">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="0ea8d-115">サンプル コード: イベント トランザクションの検証</span><span class="sxs-lookup"><span data-stu-id="0ea8d-115">Sample code: Validate event transactions</span></span>

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="0ea8d-116">トレーニング ビデオ: クロスリファレンスと書式設定を行Excelファイル</span><span class="sxs-lookup"><span data-stu-id="0ea8d-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="0ea8d-117">[Sudhi Ramamurthy が YouTube でこのサンプルを歩くのを見る](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="0ea8d-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
